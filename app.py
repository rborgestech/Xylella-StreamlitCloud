# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, shutil, time, io, zipfile, re
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from xylella_processor import process_pdf, build_zip

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
}
.stButton > button[kind="primary"]:hover,
.stButton > button[kind="primary"]:focus,
.stButton > button[kind="primary"]:active {
  background-color: #A13700 !important;
  border: 1px solid #A13700 !important;
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
}
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Função auxiliar: ler E1 (esperado/processado)
# ───────────────────────────────────────────────
def read_e1_counts(xlsx_path: str):
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

# ───────────────────────────────────────────────
# Interface de Upload
# ───────────────────────────────────────────────
uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True)

if "processing" not in st.session_state:
    st.session_state.processing = False

start = st.button("📄 Processar ficheiros de Input", type="primary", disabled=st.session_state.processing or not uploads)

# ───────────────────────────────────────────────
# Execução principal
# ───────────────────────────────────────────────
if start and uploads:
    st.session_state.processing = True
    try:
        st.info("⚙️ A processar ficheiros... aguarda alguns segundos.")
        all_excel, all_debug, summary_lines = [], [], []

        final_dir = Path.cwd() / "output_final"
        final_dir.mkdir(exist_ok=True)

        progress = st.progress(0)
        total = len(uploads)

        for i, up in enumerate(uploads, start=1):
            st.markdown(f"### 📄 {up.name}")
            st.write(f"⏳ A processar ficheiro {i}/{total}...")

            tmpdir = Path(tempfile.mkdtemp())
            tmp_path = tmpdir / up.name
            with open(tmp_path, "wb") as f:
                f.write(up.getbuffer())

            os.environ["OUTPUT_DIR"] = str(tmpdir)
            created = process_pdf(str(tmp_path))

            if not created:
                st.warning(f"⚠️ Nenhum ficheiro gerado para {up.name}")
            else:
                total_samples, discrepancies = 0, []
                for fp in created:
                    declared, processed = read_e1_counts(fp)
                    if processed:
                        total_samples += processed
                    if declared and processed and declared != processed:
                        diff = processed - declared
                        discrepancies.append(f"{Path(fp).name}: Esperado {declared}, Processado {processed} (Δ {diff:+d})")

                    dest = final_dir / Path(fp).name
                    shutil.copy(fp, dest)
                    all_excel.append(str(dest))
                    st.success(f"✅ {Path(fp).name} gravado")

                msg = f"✅ {up.name}: {len(created)} requisições, {total_samples} amostras"
                if discrepancies:
                    msg += f" ⚠️ Discrepâncias detectadas: {', '.join(discrepancies)}"
                st.info(msg)
                summary_lines.append(msg)

            # adicionar ficheiros de debug (txt, csv)
            for pattern in ["*_ocr_debug.txt", "process_log.csv", "process_summary_*.txt"]:
                for dbg in tmpdir.glob(pattern):
                    all_debug.append(str(dbg))

            progress.progress(i / total)
            time.sleep(0.3)

        # ───────────────────────────────────────────────
        # Criação do ZIP final com debug/ e summary.txt
        # ───────────────────────────────────────────────
        if all_excel:
            summary_lines.append(f"\n📊 Total: {len(all_excel)} ficheiro(s) Excel gerado(s)")
            summary_text = "\n".join(summary_lines)

            mem = io.BytesIO()
            with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as z:
                for f in all_excel:
                    z.write(f, arcname=os.path.basename(f))
                for dbg in all_debug:
                    z.write(dbg, arcname=f"debug/{os.path.basename(dbg)}")
                z.writestr("summary.txt", summary_text)
            mem.seek(0)

            zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
            st.success(f"🏁 Processamento concluído ({len(all_excel)} ficheiros Excel gerados).")
            st.download_button("⬇️ Descarregar resultados (ZIP)", data=mem.read(),
                               file_name=zip_name, mime="application/zip", type="primary")

            # Botão para limpar lista
            if st.button("🗑️ Limpar lista de ficheiros carregados"):
                st.experimental_rerun()
        else:
            st.error("⚠️ Nenhum ficheiro Excel foi detetado para incluir no ZIP.")

    finally:
        st.session_state.processing = False

else:
    st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")
