# app.py (versão final corrigida)

import os
import streamlit as st
from pathlib import Path
from xylella_processor import process_pdf_with_stats, build_zip

# ✨ Configuração base
st.set_page_config(page_title="Xylella Processor", layout="wide")
st.title("Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 Excel por requisição.")

# Diretório base de output temporário
temp_dir = Path("/tmp")
temp_dir.mkdir(parents=True, exist_ok=True)
os.environ["OUTPUT_DIR"] = str(temp_dir)

# Estado inicial
if "uploaded_files" not in st.session_state:
    st.session_state.uploaded_files = []
if "processing" not in st.session_state:
    st.session_state.processing = False
if "generated" not in st.session_state:
    st.session_state.generated = []

# Upload de ficheiros
uploaded = st.file_uploader("Carrega um ou mais PDFs", type="pdf", accept_multiple_files=True)
if uploaded:
    for f in uploaded:
        dest = temp_dir / f.name
        with open(dest, "wb") as out:
            out.write(f.read())
        st.session_state.uploaded_files.append(dest)
    st.success(f"{len(uploaded)} ficheiro(s) carregado(s). Pronto para processar.")

# Mostrar lista de ficheiros carregados
if st.session_state.uploaded_files:
    st.markdown("## 📄 Ficheiros em processamento")
    for f in st.session_state.uploaded_files:
        st.write(f.name)

# Botão para iniciar processamento
if st.session_state.uploaded_files and not st.session_state.processing:
    if st.button("Processar ficheiros"):
        st.session_state.processing = True
        st.rerun()

# Processamento efetivo
if st.session_state.processing:
    with st.spinner("\u23f3 A processar ficheiros... aguarde até o processo terminar."):
        all_excels = []
        for file_path in st.session_state.uploaded_files:
            st.markdown(f"### 📄 {file_path.name}")
            try:
                excels, stats, debug = process_pdf_with_stats(str(file_path))
                if not excels:
                    st.warning(f"⚠️ Nenhum ficheiro gerado para {file_path.name}.")
                else:
                    st.success(f"✅ {len(excels)} ficheiro(s) gerado(s).")
                    st.session_state.generated.extend(excels)
                    all_excels.extend(excels)
            except Exception as e:
                st.error(f"❌ Erro ao processar {file_path.name}: {e}")

        # Gera ZIP com todos os ficheiros
        if all_excels:
            zip_bytes, zip_name = build_zip(all_excels)
            st.download_button("📁 Download ZIP com resultados", zip_bytes, file_name=zip_name)

    # Limpar estado após processamento
    if st.button("Limpar dados"):
        st.session_state.uploaded_files = []
        st.session_state.generated = []
        st.session_state.processing = False
        st.rerun()

# Footer
st.markdown("---")
st.caption("Desenvolvido para Projeto Xylella | Streamlit App")
