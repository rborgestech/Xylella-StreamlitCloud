import streamlit as st
import os
from xylella_processor import process_pdf, build_zip

st.set_page_config(page_title="Xylella Processor", layout="centered")

# ─────────────────────────────────────────────
# Estado global
# ─────────────────────────────────────────────
if "results" not in st.session_state:
    st.session_state["results"] = []
if "is_processing" not in st.session_state:
    st.session_state["is_processing"] = False

st.title("🧪 Processamento de Requisições Xylella")

# Upload de ficheiros PDF
uploaded_files = st.file_uploader(
    "Selecionar ficheiros PDF para processar",
    type=["pdf"],
    accept_multiple_files=True,
    disabled=st.session_state["is_processing"]
)

# Botão para iniciar processamento
process_btn = st.button("🚀 Processar ficheiros", disabled=st.session_state["is_processing"])

# ─────────────────────────────────────────────
# Processar PDFs
# ─────────────────────────────────────────────
if process_btn and uploaded_files:
    st.session_state["is_processing"] = True
    st.session_state["results"].clear()

    for file in uploaded_files:
        fname = file.name
        placeholder = st.empty()
        placeholder.info(f"⏳ Início de processamento: {fname}")

        try:
            # Guardar temporariamente o ficheiro PDF
            temp_path = os.path.join("/mount/src/xylella-streamlitcloud/input_tmp", fname)
            os.makedirs(os.path.dirname(temp_path), exist_ok=True)
            with open(temp_path, "wb") as f:
                f.write(file.getbuffer())

            # Executar o processamento (retorna 1 valor, como antes)
            result = process_pdf(temp_path)

            # Atualizar interface
            if result and os.path.exists(result):
                placeholder.success(f"✅ {os.path.basename(result)} gravado")
                st.session_state["results"].append(result)
            else:
                placeholder.warning(f"⚠️ {fname}: ficheiro não gerado ou vazio")

        except Exception as e:
            placeholder.error(f"❌ Erro ao processar {fname}: {e}")

    st.session_state["is_processing"] = False

# ─────────────────────────────────────────────
# Exportar ZIP (apenas após processamento)
# ─────────────────────────────────────────────
if not st.session_state["is_processing"] and st.session_state["results"]:
    st.divider()
    st.success(f"📄 {len(st.session_state['results'])} ficheiro(s) processado(s) com sucesso.")

    if st.button("📦 Exportar resultados (ZIP)"):
        zip_path = build_zip(st.session_state["results"])
        st.download_button(
            label="⬇️ Descarregar ZIP",
            data=open(zip_path, "rb").read(),
            file_name=os.path.basename(zip_path),
            mime="application/zip",
        )

        # Limpar resultados e upload (prepara próximo processamento)
        st.session_state["results"].clear()
        st.experimental_rerun()
