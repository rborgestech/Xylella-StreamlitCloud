# xylella_processor.py
# -*- coding: utf-8 -*-
"""
Adaptador Streamlit para o Projeto Xylella.
Responsável por:
- Ligar a interface Streamlit ao motor core (core_xylella.py)
- Gerir uploads e processamento de PDFs
- Manter compatibilidade com versões anteriores (process_pdf / write_to_template)
"""

from pathlib import Path
import streamlit as st
from core_xylella import process_pdf_sync, write_to_template

# ───────────────────────────────────────────────────────────────
# Caminhos globais (robustos para Streamlit Cloud)
# ───────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "Output"
OUTPUT_DIR.mkdir(exist_ok=True)

TEMPLATE_FILENAME = "TEMPLATE_PXF_SGS.xlsx"
TEMPLATE_PATH = BASE_DIR / TEMPLATE_FILENAME

# ───────────────────────────────────────────────────────────────
# Interface Streamlit (UI)
# ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧬", layout="centered")
st.title("🧬 Xylella Processor")
st.markdown(
    "Plataforma automática de **processamento de requisições Xylella fastidiosa** com geração de relatórios Excel."
)

uploaded_files = st.file_uploader(
    "📤 Carrega um ou mais ficheiros PDF:",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.info("⚙️ A processar os ficheiros... Isto pode demorar alguns segundos por PDF.")
    for uploaded_file in uploaded_files:
        pdf_path = OUTPUT_DIR / uploaded_file.name
        with open(pdf_path, "wb") as f:
            f.write(uploaded_file.read())

        st.write(f"📄 **{uploaded_file.name}**")
        with st.spinner("A extrair dados e gerar Excel..."):
            try:
                rows = process_pdf_sync(str(pdf_path))
                st.success(f"✅ {len(rows)} amostras extraídas com sucesso!")

                excel_path = OUTPUT_DIR / (uploaded_file.name.replace(".pdf", ".xlsx"))
                if excel_path.exists():
                    with open(excel_path, "rb") as f:
                        st.download_button(
                            label="📥 Descarregar Excel",
                            data=f,
                            file_name=excel_path.name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.warning("⚠️ Ficheiro Excel não encontrado após o processamento.")
            except Exception as e:
                st.error(f"❌ Erro ao processar {uploaded_file.name}: {e}")

    st.success("🏁 Todos os ficheiros foram processados.")
else:
    st.info("💡 Carrega um ficheiro PDF para começar o processamento.")

st.markdown("---")
st.caption("Desenvolvido para o Projeto Xylella 🧪 — versão Streamlit Cloud.")

# ───────────────────────────────────────────────────────────────
# Retrocompatibilidade com scripts antigos
# ───────────────────────────────────────────────────────────────
# Estes aliases permitem que ficheiros legados (ex: app.py) continuem a importar:
# from xylella_processor import process_pdf, write_to_template
process_pdf = process_pdf_sync
