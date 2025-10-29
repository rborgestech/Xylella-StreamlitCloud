import streamlit as st
import tempfile, os, zipfile
from datetime import datetime
from xylella_processor import process_pdf, write_to_template

# ⬇️ IMPORTA o teu processador

st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor (Cloud)")

st.markdown("Faz upload de **um ou vários PDFs**. Vou gerar **um Excel por requisição**.")

uploaded_files = st.file_uploader("Carrega os PDFs", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("📄 Processar ficheiros de Input"):
        with st.spinner("A processar…"):
            tmpdir = tempfile.mkdtemp()
            outdir = os.path.join(tmpdir, "output")
            os.makedirs(outdir, exist_ok=True)

            out_paths = []
            for up in uploaded_files:
                # guardar PDF temporário
                in_path = os.path.join(tmpdir, up.name)
                with open(in_path, "wb") as f:
                    f.write(up.read())

                # correr o teu pipeline
                rows = process_pdf(in_path)
                # o nome base do excel: <nome>_req1.xlsx, etc., fica a cargo do write_to_template
                write_to_template(rows, os.path.join(outdir, os.path.splitext(up.name)[0]))

            # zip para download
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            zip_path = os.path.join(tmpdir, f"xylella_output_{stamp}.zip")
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
                for root, _, files in os.walk(outdir):
                    for fn in files:
                        full = os.path.join(root, fn)
                        arc = os.path.relpath(full, outdir)
                        z.write(full, arc)

        st.success("✅ Concluído!")
        with open(zip_path, "rb") as f:
            st.download_button("⬇️ Descarregar resultados (ZIP)", f, file_name=os.path.basename(zip_path))
