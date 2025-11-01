# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, shutil, time
from pathlib import Path
from datetime import datetime
from xylella_processor import process_pdf, build_zip

# ───────────────────────────────────────────────
# Configuração base
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 Excel por requisição.")

# ───────────────────────────────────────────────
# CSS — laranja #CA4300 e sem vermelhos
# ───────────────────────────────────────────────
st.markdown("""
<style>
/* Botão principal */
.stButton > button[kind="primary"] {
  background-color: #CA4300 !important;
  border: 1px solid #CA4300 !important;
  color: #fff !important;
  font-weight: 600 !important;
  border-radius: 6px !important;
  transition: background-color 0.2s ease-in-out !important;
}

/* Hover, Focus, Active */
.stButton > button[kind="primary"]:hover,
.stButton > button[kind="primary"]:focus,
.stButton > button[kind="primary"]:active {
  background-color: #A13700 !important;
  border: 1px solid #A13700 !important;
  color: #fff !important;
  box-shadow: none !important;
  outline: none !important;
}

/* Disabled */
.stButton > button[kind="primary"][disabled],
.stButton > button[kind="primary"][disabled]:hover {
  background-color: #b3b3b3 !important;
  border: 1px solid #b3b3b3 !important;
  color: #f2f2f2 !important;
  cursor: not-allowed !important;
  box-shadow: none !important;
}

/* File uploader */
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

/* Cores globais */
:root {
  --primary-color: #CA4300 !important;
  --secondary-color: #CA4300 !important;
  --accent-color: #CA4300 !important;
  --text-selection-color: #CA4300 !important;
}
</style>
""", unsafe_allow_html=True)

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
        st.info("⚙️ A processar... isto pode demorar alguns segundos.")
        all_excel = []

        final_dir = Path.cwd() / "output_final"
        final_dir.mkdir(exist_ok=True)

        progress = st.progress(0)
        total = len(uploads)

        for i, up in enumerate(uploads, start=1):
            st.markdown(f"### 📄 {up.name}")
            st.write("⏳ Início de processamento...")

            tmpdir = tempfile.mkdtemp()
            tmp_path = os.path.join(tmpdir, up.name)
            with open(tmp_path, "wb") as f:
                f.write(up.getbuffer())

            os.environ["OUTPUT_DIR"] = tmpdir
            created = process_pdf(tmp_path)

            if not created:
                st.warning(f"⚠️ Nenhum ficheiro gerado para {up.name}")
            else:
                for fp in created:
                    dest = final_dir / Path(fp).name
                    shutil.copy(fp, dest)
                    all_excel.append(str(dest))
                    st.success(f"✅ {Path(fp).name} gravado")

            progress.progress(i / total)
            time.sleep(0.3)

        if all_excel:
            zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
            zip_bytes = build_zip(all_excel)
            st.success(f"🏁 Processamento concluído ({len(all_excel)} ficheiros Excel gerados).")
            st.download_button("⬇️ Descarregar resultados (ZIP)", data=zip_bytes,
                               file_name=zip_name, mime="application/zip")
        else:
            st.error("⚠️ Nenhum ficheiro Excel foi detetado para incluir no ZIP.")

    finally:
        st.session_state.processing = False

else:
    st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")
