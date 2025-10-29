# xylella_processor.py
# -*- coding: utf-8 -*-
"""
Interface Streamlit para o processamento de PDFs do Projeto Xylella.
Executa o OCR e gera automaticamente ficheiros Excel no formato do TEMPLATE_PXF_SGS.xlsx.

Fluxo:
1. O utilizador faz upload de um ou mais PDFs.
2. Cada PDF é processado com o motor definido em core_xylella.py.
3. São gerados ficheiros Excel no diretório 'Output', prontos para download.

Compatível com Streamlit Cloud.
"""

from pathlib import Path
import streamlit as st
from core_xylella import process_pdf_sync

# ─────────────────────────────────────────────────────────────────────
# Configuração geral
# ─────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧬", layout="centered")
st.title("🧬 Xylella Processor")
st.markdown("Processa automaticamente **requisições Xylella** e gera relatórios em Excel.")

BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "Output"
OUTPUT_DIR.mkdir(exist_ok=True)

# ─────────────────────────────────────────────────────────────────────
# Upload de ficheiros PDF
# ─────────────────────────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "📤 Carrega um ou mais ficheiros PDF de requisições Xylella:",
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

                # Caminho do ficheiro Excel de saída
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

# Rodapé
st.markdown("---")
st.caption("Desenvolvido para o Projeto Xylella 🧪 — versão Streamlit Cloud.")
