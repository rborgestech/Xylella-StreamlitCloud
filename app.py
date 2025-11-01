# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, shutil, time, io, zipfile, re
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
# CSS — estilo limpo e azul para blocos
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
.file-title {
  font-size: 0.9rem;
  font-weight: 600;
  color: #1A365D;
}
.file-sub {
  font-size: 0.8rem;
  color: #2A4365;
}
.small-text { font-size: 0.85rem; color: #333; }
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Estado
# ───────────────────────────────────────────────
if "stage" not in st.session_state:
    st.session_state.stage = "idle"
if "uploads" not in st.session_state:
    st.session_state.uploads = None
if "reset_flag" not in st.session_state:
    st.session_state.reset_flag = False

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

def collect_debug_files(output_dirs: List[Path]) -> List[str]:
    debug_files = []
    for pattern in ["*_ocr_debug.txt", "process_log.csv", "process_summary_*.txt"]:
        for d in output_dirs:
            for f in d.glob(pattern):
                debug_files.append(str(f))
    return debug_files

def build_zip_with_summary(excel_files: List[str], debug_files: List[str], summary_text: str) -> bytes:
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
# INTERFACE PRINCIPAL
# ───────────────────────────────────────────────
if st.session_state.stage == "idle":
    uploads = st.file_uploader(
        "📂 Carrega um ou vários PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        key="file_uploader"
    )

    if uploads:
        if st.button("📄 Processar ficheiros de Input", type="primary", key="start_processing"):
            st.session_state.uploads = uploads
            st.session_state.stage = "processing"
            st.rerun()
    else:
        st.info("💡 Carrega um ficheiro PDF para ativar o botão de processamento.")

elif st.session_state.stage == "processing":
    st.info("⏳ A processar ficheiros... aguarde até o processo terminar.")
    st.divider()

    uploads = st.session_state.uploads
    session_dir = tempfile.mkdtemp(prefix="xylella_session_")
    final_dir = Path.cwd() / "output_final"
    final_dir.mkdir(exist_ok=True)
    start_time = time.time()

    all_excel, outdirs, summary_lines = [], [], []
    total = len(uploads)
    progress = st.progress(0)

    for i, up in enumerate(uploads, start=1):
        st.markdown(
            f"""
            <div class='file-box'>
                <div class='file-title'>📄 {up.name}</div>
                <div class='file-sub'>Ficheiro {i} de {total} — a processar...</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        tmpdir = Path(tempfile.mkdtemp(dir=session_dir))
        tmp_pdf = tmpdir / up.name
        with open(tmp_pdf, "wb") as f:
            f.write(up.getbuffer())

        os.environ["OUTPUT_DIR"] = str(tmpdir)
        outdirs.append(tmpdir)
        created = process_pdf(str(tmp_pdf))

        if not created:
            st.warning(f"⚠️ Nenhum ficheiro gerado para {up.name}")
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
            discrep_str = " ⚠️ Discrepâncias em " + "; ".join(discrepancies) if discrepancies else ""
            st.success(f"✅ {up.name}: {req_count} requisição(ões), {total_samples} amostras{discrep_str}.")
            summary_lines.append(f"{up.name}: {req_count} requisições, {total_samples} amostras{discrep_str}.")
        progress.progress(i / total)
        time.sleep(0.3)

    total_time = time.time() - start_time

    if all_excel:
        debug_files = collect_debug_files(outdirs)
        summary_text = "\n".join(summary_lines)
        summary_text += f"\n\n📊 Total: {len(all_excel)} ficheiro(s) Excel\n⏱️ Tempo total: {total_time:.1f} segundos"
        zip_bytes = build_zip_with_summary(all_excel, debug_files, summary_text)
        zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"

        st.success(f"🏁 Processamento concluído ({len(all_excel)} ficheiros Excel gerados).")

        # Marca que o utilizador fez download
        def mark_for_reset():
            st.session_state.reset_flag = True

        st.download_button(
            "⬇️ Descarregar resultados (ZIP)",
            data=zip_bytes,
            file_name=zip_name,
            mime="application/zip",
            on_click=mark_for_reset
        )
    else:
        st.error("⚠️ Nenhum ficheiro Excel foi detetado para incluir no ZIP.")
        shutil.rmtree(session_dir, ignore_errors=True)
        st.session_state.stage = "idle"
        st.session_state.uploads = None
        st.rerun()

# ───────────────────────────────────────────────
# RESET SEGURO APÓS DOWNLOAD
# ───────────────────────────────────────────────
# RESET SEGURO APÓS DOWNLOAD (adiado ligeiramente)
if st.session_state.reset_flag:
    with st.empty():
        st.info("🔄 A reiniciar interface...")
        time.sleep(1.2)
    st.session_state.reset_flag = False
    st.session_state.stage = "idle"
    st.session_state.uploads = None
    try:
        st.rerun()
    except Exception:
        pass

