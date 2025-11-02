# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, shutil, time, io, zipfile, re, base64, itertools, pytz
from pathlib import Path
from datetime import datetime
from typing import List, Tuple
from openpyxl import load_workbook
from xylella_processor import process_pdf

# ───────────────────────────────────────────────
# Configuração base
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 ficheiro Excel por requisição.")

# ───────────────────────────────────────────────
# CSS — estilo limpo e azul
# ───────────────────────────────────────────────
st.markdown("""
<style>
.stButton > button[kind="primary"] {
  background-color: #CA4300 !important;
  border: 1px solid #CA4300 !important;
  color: #fff !important;
  font-weight: 600 !important;
  border-radius: 6px !important;
  transition: background-color 0.2s ease-in-out !important;
}
.stButton > button[kind="primary"]:hover {
  background-color: #A13700 !important;
  border-color: #A13700 !important;
}
[data-testid="stFileUploader"] > div:first-child {
  border: 2px dashed #CA4300 !important;
  border-radius: 10px !important;
  padding: 1rem !important;
}
.file-box {
  background-color: #E8F1FB;
  border-left: 4px solid #2B6CB0;
  padding: 0.6rem 1rem;
  border-radius: 8px;
  margin-bottom: 0.5rem;
}
.file-title { font-size: 0.9rem; font-weight: 600; color: #1A365D; }
.file-sub { font-size: 0.8rem; color: #2A4365; }
.clean-btn {
  background-color: #fff !important;
  border: 1px solid #ccc !important;
  color: #333 !important;
  font-weight: 600 !important;
  border-radius: 8px !important;
  padding: 0.5rem 1.2rem !important;
  transition: all 0.2s ease-in-out !important;
}
.clean-btn:hover { border-color: #999 !important; color: #000 !important; }
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Estado
# ───────────────────────────────────────────────
if "stage" not in st.session_state:
    st.session_state.stage = "idle"
if "uploads" not in st.session_state:
    st.session_state.uploads = None

# ───────────────────────────────────────────────
# Funções auxiliares
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


def collect_debug_files(output_dirs: list[Path]) -> list[str]:
    debug_files = []
    for pattern in ["*_ocr_debug.txt", "process_log.csv", "process_summary_*.txt"]:
        for d in output_dirs:
            for f in d.glob(pattern):
                debug_files.append(str(f))
    return debug_files


def build_zip_with_summary(excel_files: list[str], debug_files: list[str], summary_text: str) -> bytes:
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as z:
        for p in excel_files:
            if os.path.exists(p):
                z.write(p, arcname=os.path.basename(p))
        for d in debug_files:
            if os.path.exists(d):
                z.write(d, arcname=f"debug/{os.path.basename(d)}")
        z.writestr("summary.txt", summary_text)
    mem.seek(0)
    return mem.read()

# ───────────────────────────────────────────────
# Função: renderiza ecrã inicial (upload)
# ───────────────────────────────────────────────
def render_home():
    st.session_state.stage = "idle"
    st.session_state.uploads = None
    st.markdown("<h3>🧪 Xylella Processor</h3>", unsafe_allow_html=True)
    st.caption("Carrega um ou vários PDFs para processar novamente.")
    uploads = st.file_uploader(
        "📂 Carrega um ou vários PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"file_uploader_{time.time()}"  # força widget novo
    )
    if uploads:
        if st.button("📄 Processar ficheiros de Input", type="primary"):
            st.session_state.uploads = uploads
            st.session_state.stage = "processing"
            st.experimental_rerun()
    else:
        st.info("💡 Carrega um ficheiro PDF para ativar o botão de processamento.")
    return

# ───────────────────────────────────────────────
# Interface principal
# ───────────────────────────────────────────────
if st.session_state.stage == "idle":
    uploads = st.file_uploader(
        "📂 Carrega um ou vários PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        key="file_uploader"
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
    start_time = time.time()

    all_excel, outdirs, summary_lines = [], [], []
    total = len(uploads)
    progress = st.progress(0)

    for i, up in enumerate(uploads, start=1):
        placeholder = st.empty()

        # Animação breve
        for frame in itertools.cycle(["..."]):
            placeholder.markdown(
                f"""
                <div class='file-box'>
                    <div class='file-title'>📄 {up.name}</div>
                    <div class='file-sub'>Ficheiro {i} de {total} — a processar{frame}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            time.sleep(0.3)
            break

        tmpdir = Path(tempfile.mkdtemp(dir=session_dir))
        tmp_pdf = tmpdir / up.name
        with open(tmp_pdf, "wb") as f:
            f.write(up.getbuffer())

        os.environ["OUTPUT_DIR"] = str(tmpdir)
        outdirs.append(tmpdir)
        created = process_pdf(str(tmp_pdf))

        if not created:
            placeholder.warning(f"⚠️ Nenhum ficheiro gerado para {up.name}")
            summary_lines.append(f"{up.name}: sem ficheiros gerados.")
        else:
            req_count = len(created)
            total_samples, discrepancies = 0, []
            for fp in created:
                dest = final_dir / Path(fp).name
                shutil.copy(fp, dest)
                all_excel.append(str(dest))
                exp, proc = read_e1_counts(str(dest))
                if exp and proc:
                    total_samples += proc
                    if exp != proc:
                        discrepancies.append(f"{Path(fp).name} (processadas: {proc} / declaradas: {exp})")
            discrep_str = " </br> ⚠️ Discrepâncias em " + "; ".join(discrepancies) if discrepancies else ""
            placeholder.success(f"✅ {up.name}:</br><b>{req_count}</b> requisição(ões), <b>{total_samples}</b> amostras{discrep_str}.")
            summary_lines.append(f"{up.name}: {req_count} requisições, {total_samples} amostras{discrep_str}.")

        progress.progress(i / total)
        time.sleep(0.2)

    total_time = time.time() - start_time

    # ──────────────── SECÇÃO FINAL ────────────────
    if all_excel:
        debug_files = collect_debug_files(outdirs)
        lisbon_tz = pytz.timezone("Europe/Lisbon")
        now_local = datetime.now(lisbon_tz)

        summary_text = "\n".join(summary_lines)
        summary_text += f"\n\n📊 Total: {len(all_excel)} ficheiro(s) Excel"
        summary_text += f"\n🧪 Total de amostras: {sum(int(m.group(1)) for l in summary_lines if (m := re.search(r'(\\d+)\\s+amostra', l)))}"
        summary_text += f"\n⏱️ Tempo total: {total_time:.1f} segundos"
        summary_text += f"\n📅 Executado em: {now_local:%d/%m/%Y às %H:%M:%S}"
        zip_bytes = build_zip_with_summary(all_excel, debug_files, summary_text)
        zip_name = f"xylella_output_{now_local:%Y%m%d_%H%M%S}.zip"

        total_reqs = len(all_excel)
        total_amostras = sum(
            int(m.group(1)) for l in summary_lines if (m := re.search(r"(\d+)\s+amostra", l))
        )

        st.markdown(f"""
        <div style='text-align:center;margin-top:1.5rem;'>
            <h3>🏁 Processamento concluído!</h3>
            <p>Foram gerados <b>{total_reqs}</b> ficheiro(s) Excel,
            com um total de <b>{total_amostras}</b> amostras processadas.<br>
            Tempo total de execução: <b>{total_time:.1f} segundos</b>.<br>
            Executado em: <b>{now_local:%d/%m/%Y às %H:%M:%S}</b>.</p>
        </div>
        """, unsafe_allow_html=True)

        zip_b64 = base64.b64encode(zip_bytes).decode()

        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"""
            <a href="data:application/zip;base64,{zip_b64}" download="{zip_name}">
                <button class="clean-btn" style="width:100%;">⬇️ Descarregar resultados (ZIP)</button>
            </a>
            """, unsafe_allow_html=True)
        def reset_app():
            st.session_state.stage = "idle"
            st.session_state.uploads = None
        
        # no final do processamento:
        with col2:
            st.button(
                "🔁 Novo processamento",
                type="secondary",
                use_container_width=True,
                on_click=reset_app
            )
