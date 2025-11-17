# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, shutil, time, io, zipfile, re, base64, pytz
from pathlib import Path
from datetime import datetime
from typing import Tuple
from openpyxl import load_workbook
from xylella_processor import process_pdf

# ───────────────────────────────────────────────
# Limpa ficheiros temporários
# ───────────────────────────────────────────────
def clean_old_tmp_artifacts():
    """
    Limpa ficheiros antigos em /tmp gerados por versões anteriores da app
    (xylella_session_*, *_ocr_debug.txt, process_log.csv, process_summary_*.txt).
    Corre no arranque da app.
    """
    base_tmp = Path(tempfile.gettempdir())

    # Limpa pastas de sessão antigas
    for d in base_tmp.glob("xylella_session_*"):
        try:
            shutil.rmtree(d, ignore_errors=True)
            print(f"🧹 Apagada pasta de sessão antiga: {d}")
        except Exception as e:
            print(f"⚠️ Não foi possível apagar {d}: {e}")

    # Limpa ficheiros de debug soltos no /tmp
    for pattern in ["*_ocr_debug.txt", "process_log.csv", "process_summary_*.txt"]:
        for f in base_tmp.glob(pattern):
            try:
                f.unlink()
                print(f"🧹 Apagado artefacto antigo: {f}")
            except Exception as e:
                print(f"⚠️ Não foi possível apagar {f}: {e}")

    # Limpa PDFs temporários soltos
    for f in base_tmp.glob("*.pdf"):
        try:
            f.unlink()
            print(f"🧹 PDF temporário antigo removido: {f}")
        except Exception as e:
            print(f"⚠️ Não foi possível remover {f}: {e}")

    # Limpa diretórios vazios restantes
    for d in base_tmp.iterdir():
        try:
            if d.is_dir() and not any(d.iterdir()):
                d.rmdir()
                print(f"🧹 Diretório vazio removido: {d}")
        except Exception:
            pass


def clean_temp_folder(path: str | Path):
    """Apaga a pasta temporária indicada, com debug opcional."""
    path = Path(path)
    if not path.exists():
        print(f"ℹ️ Pasta {path} já não existe.")
        return

    remaining = []
    for root, dirs, files in os.walk(path):
        for file in files:
            remaining.append(os.path.join(root, file))

    if remaining:
        print("⚠️ Ficheiros temporários ainda presentes:")
        for f in remaining:
            print("   └──", f)
    else:
        print("✅ Pasta temporária vazia.")

    try:
        shutil.rmtree(path, ignore_errors=True)
        print("🧹 Pasta temporária apagada com sucesso.")
    except Exception as e:
        print(f"❌ Erro ao apagar a pasta temporária: {e}")


# ───────────────────────────────────────────────
# Configuração base
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")

# Limpeza global de /tmp (artefactos antigos de execuções anteriores)
clean_old_tmp_artifacts()

st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 ficheiro Excel por requisição.")

if "processed_files" not in st.session_state:
    st.session_state.processed_files = set()

# ───────────────────────────────────────────────
# CSS — estilo azul + animações suaves
# ───────────────────────────────────────────────
st.markdown("""
<style>
.stButton > button[kind="primary"]{
  background:#CA4300!important;border:1px solid #CA4300!important;color:#fff!important;
  font-weight:600!important;border-radius:6px!important;transition:background-color .2s ease-in-out!important;
}
.stButton > button[kind="primary"]:hover{background:#A13700!important;border-color:#A13700!important;}
[data-testid="stFileUploader"]>div:first-child{border:2px dashed #CA4300!important;border-radius:10px!important;padding:1rem!important}

/* Caixas */
.file-box{border-radius:8px;padding:.6rem 1rem;margin-bottom:.5rem;opacity:0;animation:fadeIn .4s ease forwards}
@keyframes fadeIn{from{opacity:0;transform:translateY(-4px)}to{opacity:1;transform:translateY(0)}}
.fadeOut{animation:fadeOut .5s ease forwards}
@keyframes fadeOut{from{opacity:1;transform:translateY(0)}to{opacity:0;transform:translateY(-3px)}}

/* Estados */
.file-box.active{background:#E8F1FB;border-left:4px solid #2B6CB0}
.file-box.success{background:#e6f9ee;border-left:4px solid #1a7f37}
.file-box.warning{background:#fff8e5;border-left:4px solid #e6a100}
.file-box.error{background:#fdeaea;border-left:4px solid #cc0000}

/* Texto */
.file-title{font-size:.9rem;font-weight:600;color:#1A365D}
.file-sub{font-size:.8rem;color:#2A4365}

/* Pontinhos animados */
.dots::after{content:'...';display:inline-block;animation:dots 1.5s steps(4,end) infinite}
@keyframes dots{
  0%,20%{color:rgba(42,67,101,0);text-shadow:.25em 0 0 rgba(42,67,101,0),.5em 0 0 rgba(42,67,101,0)}
  40%{color:#2A4365;text-shadow:.25em 0 0 rgba(42,67,101,0),.5em 0 0 rgba(42,67,101,0)}
  60%{text-shadow:.25em 0 0 #2A4365,.5em 0 0 rgba(42,67,101,0)}
  80%,100%{text-shadow:.25em 0 0 #2A4365,.5em 0 0 #2A4365}
}

/* Botão clean (branco) */
.clean-btn{background:#fff!important;border:1px solid #ccc!important;color:#333!important;font-weight:600!important;
border-radius:8px!important;padding:.5rem 1.2rem!important;transition:all .2s ease}
.clean-btn:hover{border-color:#999!important;color:#000!important}
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Estado e reset
# ───────────────────────────────────────────────
if "stage" not in st.session_state:
    st.session_state.stage = "idle"
if "uploads" not in st.session_state:
    st.session_state.uploads = None

# ✅ Anti-duplicação
if "processed_files" not in st.session_state:
    st.session_state.processed_files = set()


def reset_app():
    st.session_state.stage = "idle"
    st.session_state.uploads = None
    st.session_state.processed_files = set()


# ───────────────────────────────────────────────
# Auxiliares
# ───────────────────────────────────────────────
def read_e1_counts(xlsx_path: str) -> Tuple[int | None, int | None]:
    try:
        wb = load_workbook(xlsx_path, data_only=True)
        ws = wb.worksheets[0]
        val = str(ws["E1"].value or "")
        m = re.search(r"(\d+)\s*/\s*(\d+)", val)
        if m:
            return int(m.group(1)), int(m.group(2))
    except Exception:
        pass
    return None, None


def build_zip_with_summary(excel_files: list[str], summary_text: str) -> bytes:
    """
    Cria um ZIP apenas com:
      • ficheiros Excel gerados
      • summary.txt com o resumo do processamento
    (sem incluir ficheiros de debug OCR).
    """
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as z:
        for p in excel_files:
            if os.path.exists(p):
                z.write(p, arcname=os.path.basename(p))
        z.writestr("summary.txt", summary_text)
    mem.seek(0)
    return mem.read()


# ───────────────────────────────────────────────
# Interface principal
# ───────────────────────────────────────────────
if st.session_state.stage == "idle":
    uploads = st.file_uploader(
        "📂 Carrega um ou vários PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        key="file_uploader",
    )
    if uploads:
        if st.button("📄 Processar ficheiros de Input", type="primary"):
            st.session_state.uploads = uploads
            st.session_state.stage = "processing"
            st.rerun()
    else:
        st.info("💡 Carrega um ficheiro PDF para ativar o botão de processamento.")

elif st.session_state.stage == "processing":
    st.info("⏳ A processar ficheiros... aguarde até o processo terminar.")
    uploads = st.session_state.uploads
    session_dir = tempfile.mkdtemp(prefix="xylella_session_")
    final_dir = Path.cwd() / "output_final"
    final_dir.mkdir(exist_ok=True)
    start_ts = time.time()

    all_excel = []
    summary_lines = []
    error_count = 0
    warning_count = 0
    total = len(uploads)
    progress = st.progress(0.0)

    for i, up in enumerate(uploads, start=1):
        if up.name in st.session_state.processed_files:
            progress.progress(i / total)
            continue

        placeholder = st.empty()
        placeholder.markdown(
            f"""
            <div class='file-box active'>
              <div class='file-title'>📄 {up.name}</div>
              <div class='file-sub'>Ficheiro {i} de {total} — a processar<span class="dots"></span></div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        tmpdir = Path(tempfile.mkdtemp(dir=session_dir))
        tmp_pdf = tmpdir / up.name
        with open(tmp_pdf, "wb") as f:
            f.write(up.getbuffer())

        # OUTPUT_DIR por sessão / por PDF
        os.environ["OUTPUT_DIR"] = str(tmpdir)

        created = process_pdf(str(tmp_pdf))

        # ───────────────────────────────────────────────
        # DEBUG NO ECRÃ (DESATIVADO, MAS PRONTO A USAR)
        # ───────────────────────────────────────────────
         debug_files = list(tmpdir.glob("*_ocr_debug.txt"))
         if debug_files:
             st.subheader(f"Ficheiros OCR Debug ({up.name})")
             for fpath in debug_files:
                 st.write(f"📄 {fpath.name}")
                 with open(fpath, "r", encoding="utf-8") as f:
                    st.text(f.read())

        st.session_state.processed_files.add(up.name)

        if not created:
            error_count += 1
            html = (
                "<div class='file-box error'>"
                f"<div class='file-title'>📄 {up.name}</div>"
                "<div class='file-sub'>❌ Erro: nenhum ficheiro gerado.</div>"
                "</div>"
            )
            placeholder.markdown(html, unsafe_allow_html=True)
            summary_lines.append(f"{up.name}: erro - nenhum ficheiro gerado.")
        else:
            req_count = len(created)
            sample_count_total = 0
            discrepancies = []

            for fp in created:
                dest = final_dir / Path(fp).name
                shutil.copy(fp, dest)
                all_excel.append(str(dest))
                exp, proc = read_e1_counts(str(dest))

                # Substitui None ou string vazia por 0
                exp = int(exp) if exp not in (None, "", " ") else 0
                proc = int(proc) if proc not in (None, "", " ") else 0

                if proc:
                    sample_count_total += proc
                    if exp != proc:
                        discrepancies.append(
                            f"⚠️ {Path(fp).name} (processadas: {proc} / declaradas: {exp})"
                        )

            box_class = "warning" if discrepancies else "success"
            if discrepancies:
                warning_count += 1
                discrep_html = (
                    "<div class='file-sub'>⚠️ <b>"
                    + str(len(discrepancies))
                    + "</b> discrepância(s):<br>"
                    + "<br>".join(discrepancies)
                    + "</div>"
                )
            else:
                discrep_html = ""

            html = (
                f"<div class='file-box {box_class}'>"
                f"<div class='file-title'>📄 {up.name}</div>"
                f"<div class='file-sub'><b>{req_count}</b> requisição(ões), "
                f"<b>{sample_count_total}</b> amostras.</div>"
                f"{discrep_html}</div>"
            )
            placeholder.markdown(html, unsafe_allow_html=True)

            # 📋 Resumo multilinha
            summary_lines.append(
                f"{up.name}: {req_count} requisições, {sample_count_total} amostras"
                + (f" ⚠️ {len(discrepancies)} discrepância(s)." if discrepancies else "")
            )

            for fp in created:
                name = Path(fp).name
                exp, proc = read_e1_counts(str(fp))

                try:
                    exp = int(exp) if exp not in (None, "", " ") else 0
                except Exception:
                    exp = 0
                try:
                    proc = int(proc) if proc not in (None, "", " ") else 0
                except Exception:
                    proc = 0

                if proc > 0 and exp != proc:
                    if exp == 0:
                        summary_lines.append(
                            f"   ↳ ⚠️ {name} (processadas: {proc} / declaradas: ausente ou 0)"
                        )
                    else:
                        summary_lines.append(
                            f"   ↳ ⚠️ {name} (processadas: {proc} / declaradas: {exp})"
                        )
                else:
                    summary_lines.append(f"   ↳ {name}")

        progress.progress(i / total)
        time.sleep(0.5)

    total_time = time.time() - start_ts
    lisbon_tz = pytz.timezone("Europe/Lisbon")
    now_local = datetime.now(lisbon_tz)
    total_reqs = len(all_excel)

    # 🧪 cálculo rigoroso — soma só 1 valor por PDF
    total_amostras = 0
    pdf_seen = set()

    for l in summary_lines:
        if l.strip().startswith("↳"):
            continue

        pdf_name = l.split(":")[0].strip()
        if pdf_name in pdf_seen:
            continue
        pdf_seen.add(pdf_name)

        m_proc = re.search(r"processadas:\s*(\d+)", l)
        m_amos = re.search(r"(\d+)\s+amostra", l)

        if m_proc:
            total_amostras += int(m_proc.group(1))
        elif m_amos:
            total_amostras += int(m_amos.group(1))

    summary_text = "\n".join(summary_lines)
    summary_text += f"\n\n📊 Total: {len(all_excel)} ficheiro(s) Excel"
    summary_text += f"\n🧪 Total de amostras: {total_amostras}"
    summary_text += f"\n⏱️ Tempo total: {total_time:.1f} segundos"
    summary_text += f"\n📅 Executado em: {now_local:%d/%m/%Y às %H:%M:%S}"

    # 🧹 Limpeza da pasta temporária de sessão (inclui OCR debug)
    try:
        clean_temp_folder(session_dir)
        summary_text += "\n🧹 Pasta temporária apagada com sucesso."
    except Exception as e:
        summary_text += f"\n⚠️ Falha ao apagar pasta temporária: {e}"

    if warning_count:
        summary_text += f"\n⚠️ {warning_count} ficheiro(s) com discrepâncias"
    if error_count:
        summary_text += f"\n❌ {error_count} ficheiro(s) com erro (sem ficheiros Excel gerados)"

    # ZIP só com Excel + summary.txt
    zip_bytes = build_zip_with_summary(all_excel, summary_text)
    zip_name = f"xylella_output_{now_local:%Y%m%d_%H%M%S}.zip"

    st.markdown(
        f"""
    <div style='text-align:center;margin-top:1.5rem;'>
      <h3>🏁 Processamento concluído!</h3>
      <p>Foram gerados <b>{total_reqs}</b> ficheiro(s) Excel,
      com um total de <b>{total_amostras}</b> amostras processadas.<br>
      Tempo total de execução: <b>{total_time:.1f} segundos</b>.<br>
      Executado em: <b>{now_local:%d/%m/%Y às %H:%M:%S}</b>.</p>
    </div>""",
        unsafe_allow_html=True,
    )

    zip_b64 = base64.b64encode(zip_bytes).decode()

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(
            f"<a href='data:application/zip;base64,{zip_b64}' download='{zip_name}'>"
            "<button class='clean-btn' style='width:100%;'>⬇️ Descarregar resultados (ZIP)</button>"
            "</a>",
            unsafe_allow_html=True,
        )
    with col2:
        st.button(
            "🔁 Novo processamento",
            type="secondary",
            use_container_width=True,
            on_click=reset_app,
        )
