import streamlit as st
import tempfile, os
from pathlib import Path
from xylella_processor import process_pdf_with_stats, build_zip

# ───────────────────────────────────────────────
# Configuração base do Streamlit
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 Excel por requisição.")

# ───────────────────────────────────────────────
# CSS — botão azul escuro e sem bordas vermelhas
# ───────────────────────────────────────────────
st.markdown("""
<style>
.stButton > button[kind="primary"] {
  background: #0b3d91 !important;
  border-color: #0b3d91 !important;
  color: #ffffff !important;
  box-shadow: none !important;
}
.stButton > button[kind="primary"]:hover,
.stButton > button[kind="primary"]:focus,
.stButton > button[kind="primary"]:active {
  background: #0a357f !important;
  border-color: #0a357f !important;
  outline: none !important;
  box-shadow: none !important;
}
:root {
  --focus-ring: 0 0 0 0 rgba(0,0,0,0) !important;
}
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Interface de Upload
# ───────────────────────────────────────────────
uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True)

if "processing" not in st.session_state:
    st.session_state.processing = False

btn = st.button("📄 Processar ficheiros de Input", type="primary", disabled=st.session_state.processing)

# ───────────────────────────────────────────────
# Processamento
# ───────────────────────────────────────────────
if btn and uploads:
    st.session_state.processing = True
    try:
        st.info("⚙️ A processar... aguarda alguns segundos.")
        log_lines = []
        all_excel = []

        for up in uploads:
            st.markdown(f"### 📄 {up.name}")
            st.write("⏳ Início de processamento...")

            tmpdir = tempfile.mkdtemp()
            tmp_path = os.path.join(tmpdir, up.name)
            with open(tmp_path, "wb") as f:
                f.write(up.getbuffer())

            os.environ["OUTPUT_DIR"] = tmpdir
            files, stats = process_pdf_with_stats(tmp_path)
            all_excel.extend(files)

            if stats["req_count"] == 0:
                line = f"{up.name}: 0 requisições, 0 amostras."
            else:
                line = f"{up.name}: {stats['req_count']} requisições, {stats['samples_total']} amostras."
            log_lines.append(line)
            st.write(line)

            for item in stats["per_req"]:
                st.write(f" ✅ Requisição {item['req']}: {item['samples']} amostras → {Path(item['file']).name}")

        if all_excel:
            zip_bytes = build_zip(all_excel, log_lines=log_lines)
            st.success("🏁 Processamento concluído.")
            st.download_button("⬇️ Descarregar resultados (ZIP)", data=zip_bytes,
                               file_name="xylella_resultados.zip", mime="application/zip")

    finally:
        st.session_state.processing = False

else:
    st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")
