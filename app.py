# app.py
import streamlit as st
import os
from pathlib import Path
from xylella_processor import process_pdf_with_stats, build_zip_with_summary

# Diretório temporário (usado pelo core)
OUTPUT_DIR = Path("/tmp")

st.set_page_config(page_title="Xylella Processor", page_icon="🦟")
st.title("\U0001F99F Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 Excel por requisição.")

if "uploaded_files" not in st.session_state:
    st.session_state.uploaded_files = []
    st.session_state.results = []

# Upload
uploaded = st.file_uploader("Carrega um ou mais PDFs", type="pdf", accept_multiple_files=True)
if uploaded:
    st.session_state.uploaded_files = uploaded
    st.session_state.results = []

# Botão para processar
if st.session_state.uploaded_files:
    if st.button("\U0001F680 Processar ficheiros"):
        with st.spinner("\u23F3 A processar ficheiros... aguarde até o processo terminar."):
            results_all = []
            for up in st.session_state.uploaded_files:
                pdf_path = os.path.join("/tmp", up.name)
                with open(pdf_path, "wb") as f:
                    f.write(up.getvalue())
                result, stats, debug_files = process_pdf_with_stats(pdf_path)
                results_all.append((up.name, result, stats))
            st.session_state.results = results_all

# Apresentar resultados
if st.session_state.results:
    summary_lines = []
    all_excels = []
    all_debug = []

    for name, files, stats in st.session_state.results:
        st.subheader(f"\U0001F4C4 {name}")
        if files:
            st.success(f"{len(files)} ficheiro(s) gerado(s).")
        else:
            st.warning(f"⚠️ Nenhum ficheiro gerado para {name}.")

        summary_lines.append(f"📄 {name}: {stats['req_count']} requisições, {stats['samples_total']} amostras")
        for req in stats["per_req"]:
            if req["diff"]:
                summary_lines.append(f"    ⚠️ Diferença na requisição {req['req']}: esperado {req['expected']}, processado {req['samples']}")
        all_excels.extend(files)

    # ZIP final
    summary_text = "Resumo do processamento:\n" + "\n".join(summary_lines)
    zip_bytes, zip_name = build_zip_with_summary(all_excels, [], summary_text)
    st.download_button("⬇️ Descarregar ZIP", data=zip_bytes, file_name=zip_name, mime="application/zip")
