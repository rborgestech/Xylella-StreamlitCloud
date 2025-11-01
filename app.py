# app.py (versão com summary detalhado)

import os
import streamlit as st
from pathlib import Path
from xylella_processor import process_pdf_with_stats, build_zip_with_summary

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
if "processed_files" not in st.session_state:
    st.session_state.processed_files = set()
if "summary_lines" not in st.session_state:
    st.session_state.summary_lines = []

# Upload de ficheiros
uploaded = st.file_uploader("Carrega um ou mais PDFs", type="pdf", accept_multiple_files=True)
if uploaded:
    for f in uploaded:
        dest = temp_dir / f.name
        if dest not in st.session_state.uploaded_files:
            with open(dest, "wb") as out:
                out.write(f.read())
            st.session_state.uploaded_files.append(dest)
    st.success(f"{len(uploaded)} ficheiro(s) carregado(s). Pronto para processar.")

# Mostrar lista de ficheiros carregados
if st.session_state.uploaded_files:
    st.markdown("## 📄 Ficheiros em processamento")
    shown = set()
    for f in st.session_state.uploaded_files:
        if f.name not in shown:
            st.write(f.name)
            shown.add(f.name)

# Botão para iniciar processamento
if st.session_state.uploaded_files and not st.session_state.processing:
    if st.button("Processar ficheiros"):
        st.session_state.processing = True
        st.rerun()

# Processamento efetivo
if st.session_state.processing:
    with st.spinner("⏳ A processar ficheiros... aguarde até o processo terminar."):
        all_excels = []
        summary_lines = []

        for file_path in st.session_state.uploaded_files:
            if file_path.name in st.session_state.processed_files:
                continue
            st.markdown(f"### 📄 {file_path.name}")
            try:
                excels, stats, debug = process_pdf_with_stats(str(file_path))
                if not excels:
                    st.warning(f"⚠️ Nenhum ficheiro gerado para {file_path.name}.")
                else:
                    st.success(f"✅ {len(excels)} ficheiro(s) gerado(s).")
                    st.session_state.generated.extend(excels)
                    all_excels.extend(excels)
                    st.session_state.processed_files.add(file_path.name)

                    # Construir linha de resumo
                    line = f"📄 {stats['pdf_name']}: {stats['req_count']} requisições, {stats['samples_total']} amostras"
                    for req in stats['per_req']:
                        if req["expected"] is not None:
                            diff = req["samples"] - req["expected"]
                            if diff != 0:
                                line += f" | Req {req['req']}: {req['samples']} vs {req['expected']} declaradas (diferença {diff:+d})"
                    summary_lines.append(line)
            except Exception as e:
                st.error(f"❌ Erro ao processar {file_path.name}: {e}")

        st.session_state.summary_lines = summary_lines

        # Gera ZIP com todos os ficheiros + summary
        if all_excels:
            summary_text = "Resumo do processamento:\n" + "\n".join(summary_lines)
            zip_bytes, zip_name = build_zip_with_summary(all_excels, [], summary_text)
            st.download_button("📁 Download ZIP com resultados", zip_bytes, file_name=zip_name)

    # Limpar estado após processamento
    if st.button("Limpar dados"):
        st.session_state.uploaded_files = []
        st.session_state.generated = []
        st.session_state.processing = False
        st.session_state.processed_files = set()
        st.session_state.summary_lines = []
        st.rerun()

# Footer
st.markdown("---")
st.caption("Desenvolvido para Projeto Xylella | Streamlit App")
