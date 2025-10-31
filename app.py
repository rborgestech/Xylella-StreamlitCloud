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

# Upload
uploaded_files = st.file_uploader(
    "Selecionar ficheiros PDF para processar",
    type=["pdf"],
    accept_multiple_files=True,
    disabled=st.session_state["is_processing"]
)

# Botões de ação
col1, col2 = st.columns(2)
with col1:
    process_btn = st.button("🚀 Processar ficheiros", disabled=st.session_state["is_processing"])
with col2:
    export_btn = st.button("📦 Exportar resultados (ZIP)", disabled=st.session_state["is_processing"])

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
            # Guarda PDF temporário
            temp_path = os.path.join("/mount/src/xylella-streamlitcloud/input_tmp", fname)
            os.makedirs(os.path.dirname(temp_path), exist_ok=True)
            with open(temp_path, "wb") as f:
                f.write(file.getbuffer())

            # Executa processamento
            out_path, summary = process_pdf(temp_path)
            n_amostras = summary.get("samples", 0)
            esperado = summary.get("expected", None)

            if esperado is not None and esperado != n_amostras:
                placeholder.warning(f"⚠️ {fname}: {n_amostras} amostras — discrepância (esperado {esperado})")
            else:
                placeholder.success(f"✅ {fname}: {n_amostras} amostras gravadas")

            st.session_state["results"].append(out_path)

        except Exception as e:
            placeholder.error(f"❌ Erro ao processar {fname}: {e}")

    st.session_state["is_processing"] = False

# ─────────────────────────────────────────────
# Exportar ZIP
# ─────────────────────────────────────────────
if export_btn and st.session_state["results"]:
    zip_path = build_zip(st.session_state["results"])
    st.success(f"📦 Exportação concluída: {os.path.basename(zip_path)}")
    st.download_button(
        label="⬇️ Descarregar ZIP",
        data=open(zip_path, "rb").read(),
        file_name=os.path.basename(zip_path),
        mime="application/zip",
    )
    # Limpa resultados e ficheiros carregados
    st.session_state["results"].clear()
    st.experimental_rerun()

