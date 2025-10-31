# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, shutil, time, re, io, zipfile
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from xylella_processor import process_pdf

# ───────────────────────────────────────────────
# Configuração base
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 Excel por requisição.")

# ───────────────────────────────────────────────
# CSS — laranja #CA4300
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
.stButton > button[kind="primary"]:hover,
.stButton > button[kind="primary"]:focus,
.stButton > button[kind="primary"]:active {
  background-color: #A13700 !important;
  border: 1px solid #A13700 !important;
  color: #fff !important;
  box-shadow: none !important;
  outline: none !important;
}
[data-testid="stFileUploader"] > div:first-child {
  border: 2px dashed #CA4300 !important;
  border-radius: 10px !important;
  padding: 1rem !important;
  transition: border-color 0.3s ease-in-out;
}
[data-testid="stFileUploader"] > div:first-child:hover {
  border-color: #A13700 !important;
}
[data-testid="stFileUploader"] > div:focus-within {
  border-color: #CA4300 !important;
  box-shadow: none !important;
}
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Estado
# ───────────────────────────────────────────────
if "processing" not in st.session_state:
    st.session_state.processing = False
if "uploads" not in st.session_state:
    st.session_state.uploads = []

# ───────────────────────────────────────────────
# Interface
# ───────────────────────────────────────────────
if not st.session_state.processing:
    uploaded = st.file_uploader(
        "📂 Carrega um ou vários PDFs",
        type=["pdf"],
        accept_multiple_files=True,
    )
    if uploaded:
        st.session_state.uploads = uploaded

if not st.session_state.processing and st.session_state.uploads:
    start = st.button("📄 Processar ficheiros de Input", type="primary")
else:
    start = False

# ───────────────────────────────────────────────
# Função auxiliar
# ───────────────────────────────────────────────
def read_e1_counts(xlsx_path: str):
    """Lê 'Nº Amostras: X / Y' da célula E1 (Esperado / Processado)."""
    declared, processed = None, None
    try:
        wb = load_workbook(xlsx_path, data_only=False)
        ws = wb.worksheets[0]
        val = str(ws["E1"].value or "")
        m = re.search(r"(\d+)\s*/\s*(\d+)", val)
        if m:
            declared = int(m.group(1))
            processed = int(m.group(2))
    except Exception:
        pass
    return declared, processed

# ───────────────────────────────────────────────
# Execução principal
# ───────────────────────────────────────────────
if start and st.session_state.uploads:
    st.session_state.processing = True
    try:
        st.info("⚙️ A processar ficheiros... aguarda alguns segundos.")
        all_excel, debug_files, summary_lines = [], [], []
        progress = st.progress(0.0)
        total = len(st.session_state.uploads)

        for i, up in enumerate(st.session_state.uploads, start=1):
            st.markdown(f"### 📄 {up.name}")
            st.write(f"⏳ A processar ficheiro {i}/{total}...")

            tmpdir = Path(tempfile.mkdtemp(prefix="xylella_"))
            tmp_path = tmpdir / up.name
            with open(tmp_path, "wb") as f:
                f.write(up.getbuffer())

            os.environ["OUTPUT_DIR"] = str(tmpdir)
            created = process_pdf(str(tmp_path))

            if not created:
                st.warning(f"⚠️ Nenhum ficheiro gerado para {up.name}")
                continue

            req_count = len(created)
            total_samples, discrepancy_msgs = 0, []

            for fp in created:
                declared, processed = read_e1_counts(fp)
                if processed:
                    total_samples += processed
                if declared is not None and processed is not None and declared != processed:
                    diff = processed - declared
                    discrepancy_msgs.append(f"{Path(fp).name}: Esperado {declared}, Processado {processed} (Δ {diff:+d})")
                all_excel.append(fp)
                st.success(f"✅ {Path(fp).name} gravado")

            if discrepancy_msgs:
                st.warning(f"✅ {up.name}: {req_count} requisições, {total_samples} amostras (⚠️ discrepâncias: {', '.join(discrepancy_msgs)})")
            else:
                st.success(f"✅ {up.name}: {req_count} requisições, {total_samples} amostras (sem discrepâncias)")

            summary_lines.append(f"{up.name}: {req_count} requisições, {total_samples} amostras.")

            # recolhe debug files
            for pat in ["*_ocr_debug.txt", "process_log.csv", "process_summary_*.txt"]:
                for f in tmpdir.glob(pat):
                    debug_files.append(str(f))

            progress.progress(i / total)
            time.sleep(0.2)

        # ZIP final com debug + summary
        if all_excel:
            summary_lines.append(f"\n📊 Total: {len(all_excel)} ficheiro(s) Excel gerado(s)")
            summary_text = "\n".join(summary_lines)

            zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
            mem = io.BytesIO()
            with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as z:
                for f in all_excel:
                    z.write(f, arcname=os.path.basename(f))
                for dbg in debug_files:
                    z.write(dbg, arcname=f"debug/{os.path.basename(dbg)}")
                z.writestr("summary.txt", summary_text)
            mem.seek(0)

            st.success(f"🏁 Processamento concluído ({len(all_excel)} ficheiros Excel gerados).")
            st.download_button(
                "⬇️ Descarregar resultados (ZIP)",
                data=mem.read(),
                file_name=zip_name,
                mime="application/zip",
                type="primary"
            )
            st.balloons()

        else:
            st.error("⚠️ Nenhum ficheiro Excel foi detetado para incluir no ZIP.")

    except Exception as e:
        st.error(f"❌ Erro inesperado: {e}")

    finally:
        st.session_state.processing = False

# ───────────────────────────────────────────────
# Botão para limpar lista de ficheiros
# ───────────────────────────────────────────────
if not st.session_state.processing and st.session_state.uploads:
    if st.button("🗑️ Limpar lista de ficheiros carregados"):
        st.session_state.uploads = []
        st.experimental_rerun()
