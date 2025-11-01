import streamlit as st
from pathlib import Path
from datetime import datetime
import shutil
import os

from core_xylella import process_pdf_sync

# ─────────────────────────────────────────────────────────────────────
# Diretório base de output
# ─────────────────────────────────────────────────────────────────────
BASE_OUTPUT_DIR = Path("output_final")
BASE_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Estado inicial
if "uploaded_files" not in st.session_state:
    st.session_state.uploaded_files = []
if "generated_files" not in st.session_state:
    st.session_state.generated_files = []
if "processing" not in st.session_state:
    st.session_state.processing = False
if "output_run_dir" not in st.session_state:
    st.session_state.output_run_dir = ""

# ─────────────────────────────────────────────────────────────────────
# Cabeçalho
# ─────────────────────────────────────────────────────────────────────
st.title("📄 Processador de Requisições Xylella")
st.caption("Versão OCR Azure + Parser Colab")

# Botão para limpar tudo
if st.button("🧹 Limpar dados"):
    st.session_state.uploaded_files = []
    st.session_state.generated_files = []
    st.session_state.output_run_dir = ""
    st.experimental_rerun()

# ─────────────────────────────────────────────────────────────────────
# Upload
# ─────────────────────────────────────────────────────────────────────
uploaded = st.file_uploader("Seleciona um ou mais PDFs:", type=["pdf"], accept_multiple_files=True, disabled=st.session_state.processing)

if uploaded:
    for f in uploaded:
        if f.name not in [u.name for u in st.session_state.uploaded_files]:
            st.session_state.uploaded_files.append(f)

# ─────────────────────────────────────────────────────────────────────
# Processar ficheiros
# ─────────────────────────────────────────────────────────────────────
if st.session_state.uploaded_files and not st.session_state.processing:
    if st.button("🚀 Processar ficheiros", disabled=st.session_state.processing):
        st.session_state.processing = True
        st.experimental_rerun()

# ─────────────────────────────────────────────────────────────────────
# Processamento real
# ─────────────────────────────────────────────────────────────────────
if st.session_state.processing and st.session_state.uploaded_files:
    # Criar subpasta com timestamp
    ts = datetime.now().strftime("run_%Y%m%d_%H%M%S")
    output_dir = BASE_OUTPUT_DIR / ts
    output_dir.mkdir(parents=True, exist_ok=True)
    st.session_state.output_run_dir = str(output_dir)

    st.write("⏳ A processar ficheiros... aguarde até o processo terminar.")
    with st.spinner("A processar..."):
        for uploaded_file in st.session_state.uploaded_files:
            filename = uploaded_file.name
            temp_path = output_dir / filename
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.read())
            st.write(f"📄 A processar: {filename}")
            try:
                os.environ["OUTPUT_DIR"] = str(output_dir.resolve())
                generated_paths = process_pdf_sync(str(temp_path))
                if generated_paths:
                    st.session_state.generated_files.extend(generated_paths)
                    for path in generated_paths:
                        st.success(f"✅ Ficheiro gerado: {os.path.basename(path)}")
                else:
                    st.warning(f"⚠️ Nenhum ficheiro gerado para {filename}.")
            except Exception as e:
                st.error(f"❌ Erro ao processar {filename}: {e}")

    st.session_state.processing = False

# ─────────────────────────────────────────────────────────────────────
# Ficheiros gerados
# ─────────────────────────────────────────────────────────────────────
if st.session_state.generated_files:
    st.subheader("📂 Ficheiros gerados")
    for path in st.session_state.generated_files:
        path_obj = Path(path)
        if path_obj.exists():
            with open(path_obj, "rb") as f:
                st.download_button(
                    label=f"⬇️ Download {path_obj.name}",
                    data=f,
                    file_name=path_obj.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    # ZIP de todos
    if len(st.session_state.generated_files) > 1:
        from zipfile import ZipFile
        zip_path = Path(st.session_state.output_run_dir) / "ficheiros_xylella.zip"
        with ZipFile(zip_path, "w") as zipf:
            for file_path in st.session_state.generated_files:
                zipf.write(file_path, arcname=Path(file_path).name)
        with open(zip_path, "rb") as zf:
            st.download_button("📦 Download ZIP", data=zf, file_name="ficheiros_xylella.zip", mime="application/zip")

    # Limpa lista depois de exportar
    st.session_state.uploaded_files = []
    st.session_state.generated_files = []

# ─────────────────────────────────────────────────────────────────────
# Fim
# ─────────────────────────────────────────────────────────────────────
st.write("---")
st.caption("Desenvolvido para SGS por [NeoPackS] · Versão cloud")
