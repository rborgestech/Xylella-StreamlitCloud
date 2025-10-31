# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, shutil, time
from pathlib import Path
from datetime import datetime
from xylella_processor import process_pdf, build_zip

st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 Excel por requisição.")

# ───────────────────────────────────────────────
# CSS Laranja SGS
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
.stButton > button[kind="primary"]:hover {
  background-color: #A13700 !important;
  border: 1px solid #A13700 !important;
}
[data-testid="stFileUploader"] > div:first-child {
  border: 2px dashed #CA4300 !important;
  border-radius: 10px !important;
  padding: 1rem !important;
}
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Estado
# ───────────────────────────────────────────────
if "processing" not in st.session_state:
    st.session_state.processing = False

# ───────────────────────────────────────────────
# Interface
# ───────────────────────────────────────────────
if not st.session_state.processing:
    uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True)
    start = st.button("📄 Processar ficheiros de Input", type="primary", disabled=not uploads)

    if start and uploads:
        st.session_state.uploads = uploads
        st.session_state.processing = True
        st.experimental_rerun()
else:
    st.info("⏳ A processar ficheiros... aguarde até o processo terminar.")
    uploads = st.session_state.uploads

    # Execução principal
    session_dir = tempfile.mkdtemp(prefix="xylella_session_")
    all_excel = []
    total = len(uploads)
    progress = st.progress(0)
    status_text = st.empty()

    try:
        for i, up in enumerate(uploads, start=1):
            status_text.markdown(f"### 📄 A processar ficheiro **{i}/{total}**: `{up.name}`")

            tmpdir = tempfile.mkdtemp(dir=session_dir)
            tmp_path = os.path.join(tmpdir, up.name)
            with open(tmp_path, "wb") as f:
                f.write(up.getbuffer())

            os.environ["OUTPUT_DIR"] = tmpdir
            result = process_pdf(tmp_path)

            if isinstance(result, tuple) and len(result) == 3:
                created, n_amostras, discrepancias = result
            else:
                created, n_amostras, discrepancias = result, None, None

            if not created:
                st.warning(f"⚠️ Nenhum ficheiro gerado para {up.name}")
            else:
                for fp in created:
                    all_excel.append(fp)
                    msg = f"✅ {Path(fp).name} gravado"
                    if n_amostras is not None:
                        msg += f" — {n_amostras} amostras"
                        if discrepancias:
                            msg += f", {discrepancias} discrepâncias"
                    st.success(msg)

            progress.progress(i / total)
            time.sleep(0.2)

        # ZIP final
        if all_excel:
            zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
            zip_bytes = build_zip(all_excel)
            st.success(f"🏁 Processamento concluído ({len(all_excel)} ficheiros Excel gerados).")
            st.download_button("⬇️ Descarregar resultados (ZIP)", data=zip_bytes,
                               file_name=zip_name, mime="application/zip")
            st.balloons()
        else:
            st.error("⚠️ Nenhum ficheiro Excel foi detetado para incluir no ZIP.")

    except Exception as e:
        st.error(f"❌ Erro inesperado: {e}")

    finally:
        shutil.rmtree(session_dir, ignore_errors=True)
        st.session_state.processing = False
        st.session_state.uploads = None
        st.experimental_rerun()
