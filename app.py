# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, shutil, time, io, zipfile, re
from pathlib import Path
from datetime import datetime
from typing import List, Tuple
from openpyxl import load_workbook
from xylella_processor import process_pdf

# ───────────────────────────────────────────────
# Configuração base e estilo
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 ficheiro Excel por requisição.")

# CSS — tons laranja (#CA4300)
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
.stButton > button[kind="primary"]:hover,
.stButton > button[kind="primary"]:focus {
  background-color: #A13700 !important;
  border: 1px solid #A13700 !important;
}
.stButton > button[kind="primary"][disabled],
.stButton > button[kind="primary"][disabled]:hover {
  background-color: #b3b3b3 !important;
  border: 1px solid #b3b3b3 !important;
  color: #f2f2f2 !important;
  cursor: not-allowed !important;
}
[data-testid="stFileUploader"] > div:first-child {
  border: 2px dashed #CA4300 !important;
  border-radius: 10px !important;
  padding: 1rem !important;
}
[data-testid="stFileUploader"] > div:first-child:hover {
  border-color: #A13700 !important;
}
[data-testid="stFileUploader"] > div:focus-within {
  border-color: #CA4300 !important;
  box-shadow: none !important;
}
:root {
  --primary-color: #CA4300 !important;
  --secondary-color: #CA4300 !important;
  --accent-color: #CA4300 !important;
}
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Estado
# ───────────────────────────────────────────────
if "processing" not in st.session_state:
    st.session_state.processing = False
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = "uploader_1"

# ───────────────────────────────────────────────
# Helpers
# ───────────────────────────────────────────────
def read_e1_counts(xlsx_path: str) -> Tuple[int | None, int | None]:
    """Lê a célula E1 do Excel e extrai 'Nº Amostras: X / Y'."""
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
    """Recolhe ficheiros de debug (_ocr_debug.txt, process_log.csv, summaries)."""
    debug_files = []
    patterns = ["*_ocr_debug.txt", "process_log.csv", "process_summary_*.txt"]
    for outdir in output_dirs:
        for pat in patterns:
            debug_files.extend(str(f) for f in outdir.glob(pat) if f.exists())
    return debug_files

def build_zip_with_summary(excel_files: List[str], debug_files: List[str], summary_text: str) -> bytes:
    """Constrói um ZIP com Excel, debug/ e summary.txt."""
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as z:
        for p in excel_files:
            if p and os.path.exists(p):
                z.write(p, arcname=os.path.basename(p))
        for d in debug_files:
            if d and os.path.exists(d):
                z.write(d, arcname=f"debug/{os.path.basename(d)}")
        z.writestr("summary.txt", summary_text or "")
    mem.seek(0)
    return mem.read()

# ───────────────────────────────────────────────
# Interface (oculta durante processamento)
# ───────────────────────────────────────────────
if not st.session_state.processing:
    uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True,
                               key=st.session_state.uploader_key)
    start = st.button("📄 Processar ficheiros de Input", type="primary",
                      disabled=not uploads or st.session_state.processing)
else:
    uploads, start = None, False
    st.info("⚙️ A processar... aguarda alguns segundos.")

# ───────────────────────────────────────────────
# Execução principal
# ───────────────────────────────────────────────
if start and uploads:
    st.session_state.processing = True
    session_dir = tempfile.mkdtemp(prefix="xylella_session_")
    final_dir = Path.cwd() / "output_final"
    final_dir.mkdir(exist_ok=True)

    try:
        all_excel, outdirs, summary_lines = [], [], []
        progress = st.progress(0)
        total = len(uploads)

        for i, up in enumerate(uploads, start=1):
            st.markdown(f"### 📄 {up.name}")
            st.write("⏳ Início de processamento...")

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
                total_samples, discrepancy_msgs = 0, []
                for fp in created:
                    dest = final_dir / Path(fp).name
                    shutil.copy(fp, dest)
                    all_excel.append(str(dest))

                    exp, proc = read_e1_counts(str(dest))
                    if exp is not None and proc is not None:
                        total_samples += proc
                        diff = proc - exp
                        if diff != 0:
                            discrepancy_msgs.append(
                                f"⚠️ discrepância ({proc} processadas / {exp} declaradas)"
                            )

                discrep_str = f" ({'; '.join(discrepancy_msgs)})" if discrepancy_msgs else ""
                st.success(f"✅ {up.name}: {req_count} requisição(ões), {total_samples} amostras{discrep_str}.")
                summary_lines.append(f"{up.name}: {req_count} requisições, {total_samples} amostras{discrep_str}.")

            progress.progress(i / total)
            time.sleep(0.2)

        if all_excel:
            debug_files = collect_debug_files(outdirs)
            summary_text = "\n".join(summary_lines) + f"\n\n📊 Total: {len(all_excel)} ficheiro(s) Excel"
            zip_bytes = build_zip_with_summary(all_excel, debug_files, summary_text)
            zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"

            st.success(f"🏁 Processamento concluído ({len(all_excel)} ficheiros Excel gerados).")
            st.download_button("⬇️ Descarregar resultados (ZIP)", data=zip_bytes,
                               file_name=zip_name, mime="application/zip")

            # Limpa automaticamente a lista de ficheiros carregados
            st.session_state.uploader_key = f"uploader_{datetime.now().timestamp()}"
        else:
            st.error("⚠️ Nenhum ficheiro Excel foi detetado para incluir no ZIP.")
    finally:
        shutil.rmtree(session_dir, ignore_errors=True)
        st.session_state.processing = False

# Estado inativo
if not st.session_state.processing:
    st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")
