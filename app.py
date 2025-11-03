# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, shutil, time, io, zipfile, re, base64, pytz
from pathlib import Path
from datetime import datetime
from typing import Tuple, List, Dict, Any
from openpyxl import load_workbook
from concurrent.futures import ThreadPoolExecutor, as_completed
from xylella_processor import process_pdf  # devolve paths absolutos .xlsx

# ───────────────────────────────────────────────
# Configuração base
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 ficheiro Excel por requisição.")

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

/* Botão branco */
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

def reset_app():
    st.session_state.stage = "idle"
    st.session_state.uploads = None
    st.experimental_rerun()

# ───────────────────────────────────────────────
# Auxiliares
# ───────────────────────────────────────────────
def read_e1_counts(xlsx_path: str) -> Tuple[int | None, int | None]:
    """
    Lê E1: 'Nº Amostras: {esperado} / {processado}'.
    Devolve (esperado, processado), mesmo que a célula esteja fundida.
    """
    import time
    for attempt in range(3):  # tenta 3 vezes (em caso de ficheiro em escrita)
        try:
            wb = load_workbook(xlsx_path, data_only=True)
            ws = wb.active
            candidates = [ws["E1"].value, ws.cell(1, 5).value, ws.cell(1, 6).value]
            val = next((v for v in candidates if isinstance(v, str) and "/" in v), "")
            m = re.search(r"(\d+)\s*/\s*(\d+)", val)
            if m:
                exp, proc = int(m.group(1)), int(m.group(2))
                return exp, proc
            break
        except PermissionError:
            time.sleep(0.5)  # aguarda ficheiro terminar de gravar
        except Exception as e:
            print(f"[WARN] Falha ao ler E1 em {xlsx_path}: {e}")
            break
    return None, None


def copy_if_new(src: str, dest_dir: Path) -> str | None:
    """Evita duplicações: só copia se não existir com o mesmo tamanho."""
    try:
        src_p = Path(src)
        dest_p = dest_dir / src_p.name
        if dest_p.exists():
            if dest_p.stat().st_size == src_p.stat().st_size:
                return str(dest_p)  # já existe igual
        shutil.copy(src_p, dest_p)
        return str(dest_p)
    except Exception:
        return None

# ───────────────────────────────────────────────
# Interface principal
# ───────────────────────────────────────────────
if st.session_state.stage == "idle":
    uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True, key="file_uploader")
    if uploads:
        if st.button("📄 Processar ficheiros de Input", type="primary"):
            st.session_state.uploads = uploads
            st.session_state.stage = "processing"
            st.rerun()
    else:
        st.info("💡 Carrega um ficheiro PDF para ativar o botão de processamento.")

elif st.session_state.stage == "processing":
    uploads = st.session_state.get("uploads") or []
    if not uploads:
        st.warning("⚠️ Nenhum ficheiro encontrado. Volte atrás e carregue novamente os PDFs.")
        st.session_state.stage = "idle"
        st.rerun()

    st.info("⏳ A processar ficheiros... aguarde até o processo terminar.")

    session_dir = tempfile.mkdtemp(prefix="xylella_session_")
    final_dir = Path.cwd() / "output_final"
    final_dir.mkdir(exist_ok=True)
    start_ts = time.time()

    # placeholders por ficheiro (UI criada no thread principal)
    placeholders: Dict[str, Any] = {}
    total = len(uploads)
    for idx, up in enumerate(uploads, start=1):
        ph = st.empty()
        ph.markdown(
            f"<div class='file-box active'><div class='file-title'>📄 {up.name}</div>"
            f"<div class='file-sub'>Ficheiro {idx} de {total} — a processar<span class='dots'></span></div></div>",
            unsafe_allow_html=True,
        )
        placeholders[up.name] = ph

    progress = st.progress(0.0)
    completed = 0

    all_excel: List[str] = []
    summary_lines: List[str] = []
    warning_count = 0
    error_count = 0
    overall_samples = 0

    outdirs: List[Path] = []

    def worker(idx: int, up) -> Dict[str, Any]:
        """
        Executa o processamento de 1 PDF (thread worker).
        Devolve dict com: base, created(list), req_count, samples, discrepancies(list[str]), tmpdir
        """
        base = up.name
        tmpdir = Path(tempfile.mkdtemp(dir=session_dir))
        outdirs.append(tmpdir)
        tmp_pdf = tmpdir / base
        with open(tmp_pdf, "wb") as f:
            f.write(up.getbuffer())
        os.environ["OUTPUT_DIR"] = str(tmpdir)

        t0 = time.time()
        created_paths = process_pdf(str(tmp_pdf))  # já pode ter paralelismo interno
        elapsed = time.time() - t0

        # Ler expected/proc de cada xlsx criado e recolher discrepâncias
        req_count = len(created_paths or [])
        samples_total = 0
        discrepancies: List[str] = []
        per_req_excel: List[str] = []

        if created_paths:
            for fp in created_paths:
                per_req_excel.append(Path(fp).name)
                exp, proc = read_e1_counts(fp)
                if proc:
                    samples_total += proc
                if exp is not None and proc is not None and exp != proc:
                    discrepancies.append(f"⚠️ {Path(fp).name} (processadas: {proc} / declaradas: {exp})")

        return {
            "base": base,
            "created": created_paths or [],
            "req_count": req_count,
            "samples": samples_total,
            "discrepancies": discrepancies,
            "elapsed": elapsed,
            "tmpdir": tmpdir,
            "per_req_excel": per_req_excel,
        }

    # Lança threads (nível PDF) e trata resultados à medida que terminam
    with ThreadPoolExecutor(max_workers=min(3, total)) as ex:
        futures = {ex.submit(worker, i, up): up.name for i, up in enumerate(uploads, 1)}
        for fut in as_completed(futures):
            res = fut.result()
            base = res["base"]
            ph = placeholders.get(base)

            if not res["created"]:
                error_count += 1
                html = (
                    f"<div class='file-box error'><div class='file-title'>📄 {base}</div>"
                    f"<div class='file-sub'>❌ Erro: nenhum ficheiro gerado.</div></div>"
                )
                if ph:
                    ph.markdown(html, unsafe_allow_html=True)
                summary_lines.append(f"{base}: erro - nenhum ficheiro gerado.")
            else:
                # Copiar resultados (sem duplicar)
                copied_for_zip: List[str] = []
                for fp in res["created"]:
                    copied = copy_if_new(fp, final_dir)
                    if copied:
                        all_excel.append(copied)
                        copied_for_zip.append(copied)

                # UI final por ficheiro
                box_class = "warning" if res["discrepancies"] else "success"
                if res["discrepancies"]:
                    warning_count += 1
                    discrep_html = (
                        "<div class='file-sub'>⚠️ <b>" + str(len(res["discrepancies"])) +
                        "</b> discrepância(s):<br>" + "<br>".join(res["discrepancies"]) + "</div>"
                    )
                else:
                    discrep_html = ""

                html = (
                    f"<div class='file-box {box_class}'>"
                    f"<div class='file-title'>📄 {base}</div>"
                    f"<div class='file-sub'><b>{res['req_count']}</b> requisição(ões), "
                    f"<b>{res['samples']}</b> amostras.</div>"
                    f"{discrep_html}"
                    f"</div>"
                )
                if ph:
                    ph.markdown(html, unsafe_allow_html=True)

                # Resumo (com sublinhas por xlsx, marcando os com discrepância)
                summary_lines.append(
                    f"{base}: {res['req_count']} requisições, {res['samples']} amostras" +
                    (f" ⚠️ {len(res['discrepancies'])} discrepância(s)." if res["discrepancies"] else "")
                )
                # Primeiro, listar as discrepâncias como sublinhas
                discrep_names = set()
                for d in res["discrepancies"]:
                    summary_lines.append(f"   ↳ {d}")
                    # d começa com "⚠️ <excel_name> ..." → extrair nome
                    m = re.match(r"^⚠️\s+(.+?\.xlsx)\b", d)
                    if m:
                        discrep_names.add(m.group(1))

                # Depois, listar os restantes excels (sem ⚠️)
                for name in res["per_req_excel"]:
                    if name not in discrep_names:
                        summary_lines.append(f"   ↳ {name}")

                overall_samples += res["samples"]

            completed += 1
            progress.progress(completed / total)

    total_time = time.time() - start_ts
    debug_files = collect_debug_files([Path(d) for d in set(outdirs)])
    lisbon_tz = pytz.timezone("Europe/Lisbon")
    now_local = datetime.now(lisbon_tz)

    # ──────────────── SECÇÃO FINAL ────────────────
    if all_excel:
        summary_text = "\n".join(summary_lines)
        summary_text += f"\n\n📊 Total: {len(all_excel)} ficheiro(s) Excel"
        summary_text += f"\n🧪 Total de amostras: {overall_samples}"
        summary_text += f"\n⏱️ Tempo total: {total_time:.1f} segundos"
        summary_text += f"\n📅 Executado em: {now_local:%d/%m/%Y às %H:%M:%S}"
        if warning_count:
            summary_text += f"\n⚠️ {warning_count} ficheiro(s) com discrepâncias"
        if error_count:
            summary_text += f"\n❌ {error_count} ficheiro(s) com erro (sem ficheiros Excel gerados)"

        zip_bytes = build_zip_with_summary(all_excel, debug_files, summary_text)
        zip_name = f"xylella_output_{now_local:%Y%m%d_%H%M%S}.zip"

        # Bloco final
        st.markdown(f"""
        <div style='text-align:center;margin-top:1.5rem;'>
          <h3>🏁 Processamento concluído!</h3>
          <p>Foram gerados <b>{len(all_excel)}</b> ficheiro(s) Excel,
          com um total de <b>{overall_samples}</b> amostras processadas.<br>
          Tempo total de execução: <b>{total_time:.1f} segundos</b>.<br>
          Executado em: <b>{now_local:%d/%m/%Y às %H:%M:%S}</b>.</p>
        </div>
        """, unsafe_allow_html=True)

        zip_b64 = base64.b64encode(zip_bytes).decode()
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(
                f"<a href='data:application/zip;base64,{zip_b64}' download='{zip_name}'>"
                f"<button class='clean-btn' style='width:100%;'>⬇️ Descarregar resultados (ZIP)</button></a>",
                unsafe_allow_html=True
            )
        with col2:
            st.button("🔁 Novo processamento", type="secondary", use_container_width=True, on_click=reset_app)
    else:
        st.error("⚠️ Nenhum ficheiro Excel foi detetado para incluir no ZIP.")
