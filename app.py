# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, shutil, time
from pathlib import Path
from datetime import datetime
from xylella_processor import process_pdf, build_zip

# ───────────────────────────────────────────────
# Configuração base
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 Excel por requisição.")

# ───────────────────────────────────────────────
# CSS — tema SGS + estilos visuais
# ───────────────────────────────────────────────
st.markdown("""
<style>
/* Botões principais */
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

/* File uploader */
[data-testid="stFileUploader"] > div:first-child {
  border: 2px dashed #CA4300 !important;
  border-radius: 10px !important;
  padding: 1rem !important;
}

/* Caixas */
.success-box {
  background-color: #E8F5E9;
  border-left: 5px solid #2E7D32;
  padding: 0.7rem 1rem;
  border-radius: 6px;
  margin-bottom: 0.4rem;
}
.warning-box {
  background-color: #FFF8E1;
  border-left: 5px solid #FBC02D;
  padding: 0.7rem 1rem;
  border-radius: 6px;
  margin-bottom: 0.4rem;
}
.info-box {
  background-color: #E3F2FD;
  border-left: 5px solid #1E88E5;
  padding: 0.7rem 1rem;
  border-radius: 6px;
  margin-bottom: 0.4rem;
}

/* Botões finais */
.button-row {
  display: flex;
  justify-content: center;
  align-items: center;
  gap: 1rem;
  margin-top: 1.5rem;
}
.stDownloadButton button, .stButton button {
  background-color: #ffffff !important;
  border: 1.5px solid #CA4300 !important;
  color: #CA4300 !important;
  font-weight: 600 !important;
  border-radius: 8px !important;
  padding: 0.6rem 1.2rem !important;
  transition: all 0.2s ease-in-out;
}
.stDownloadButton button:hover, .stButton button:hover {
  background-color: #CA4300 !important;
  color: #ffffff !important;
  border-color: #A13700 !important;
}
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Estado
# ───────────────────────────────────────────────
for k in ["processing", "finished", "uploads", "results", "zip_bytes", "zip_name"]:
    if k not in st.session_state:
        st.session_state[k] = False if k in ["processing", "finished"] else []

# ───────────────────────────────────────────────
# Ecrã inicial
# ───────────────────────────────────────────────
if not st.session_state.processing and not st.session_state.finished:
    uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True)
    if uploads:
        st.session_state.uploads = uploads
        if st.button(f"📄 Processar {len(uploads)} ficheiro(s) de Input", type="primary"):
            st.session_state.processing = True
            st.rerun()
    else:
        st.info("💡 Carrega ficheiros PDF para ativar o processamento.")

# ───────────────────────────────────────────────
# Ecrã de processamento
# ───────────────────────────────────────────────
elif st.session_state.processing:
    uploads = st.session_state.uploads
    total = len(uploads)
    st.markdown('<div class="info-box">⏳ A processar ficheiros... aguarde até o processo terminar.</div>', unsafe_allow_html=True)

    progress = st.progress(0)
    all_excel = []
    session_dir = tempfile.mkdtemp(prefix="xylella_session_")

    try:
        for i, up in enumerate(uploads, start=1):
            st.markdown(f"### 📘 A processar ficheiro {i}/{total}: `{up.name}`")

            tmpdir = tempfile.mkdtemp(dir=session_dir)
            tmp_path = os.path.join(tmpdir, up.name)
            with open(tmp_path, "wb") as f:
                f.write(up.getbuffer())

            os.environ["OUTPUT_DIR"] = tmpdir
            results = process_pdf(tmp_path)

            if not results:
                st.markdown(f'<div class="warning-box">⚠️ Nenhum ficheiro gerado para <b>{up.name}</b>.</div>', unsafe_allow_html=True)
            else:
                st.markdown("#### 📄 Ficheiros gerados")
                for r in results:
                    fp = r.get("path")
                    samples = r.get("samples", 0)
                    declared = r.get("declared", samples)
                    diff = r.get("diff", 0)
                    all_excel.append(fp)

                    if diff and diff != 0:
                        st.markdown(f'<div class="warning-box">⚠️ {Path(fp).name}: ficheiro gerado. ({samples} vs {declared} — discrepância {diff:+d})</div>', unsafe_allow_html=True)
                    else:
                        st.markdown(f'<div class="success-box">✅ {Path(fp).name}: ficheiro gerado. ({samples} amostras OK)</div>', unsafe_allow_html=True)

            progress.progress(i / total)
            time.sleep(0.3)

        # Finalização
        if all_excel:
            st.session_state.results = all_excel
            st.session_state.finished = True
            st.session_state.zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
            st.session_state.zip_bytes = build_zip(all_excel)
        else:
            st.warning("⚠️ Nenhum ficheiro Excel foi detetado.")

    except Exception as e:
        st.error(f"❌ Erro inesperado: {e}")
    finally:
        shutil.rmtree(session_dir, ignore_errors=True)
        st.session_state.processing = False

# ───────────────────────────────────────────────
# Ecrã final — painel verde com resumo e botões
# ───────────────────────────────────────────────
if st.session_state.finished and st.session_state.results:
    results = st.session_state.results

    # Estatísticas
    total_files = len(results)
    total_samples = 0
    total_discrep = 0

    for r in results:
        total_samples += r.get("samples", 0)
        if r.get("diff", 0):
            total_discrep += 1

    st.markdown(
        f"""
        <div style="background:#E8F5E9; border-left:6px solid #2E7D32; border-radius:10px;
                    padding:1.2rem 1.6rem; margin-top:1.4rem;">
            <h4 style="color:#2E7D32; font-weight:600; margin:.2rem 0 .3rem 0;">✅ Processamento concluído</h4>
            <p style="color:#2E7D32; margin:0;">📊 Total: {total_files} ficheiro(s) Excel</p>
            <p style="color:#2E7D32; margin:0;">🧪 Total de amostras processadas: {total_samples}</p>
            <p style="color:#2E7D32; margin:0;">⚠️ {total_discrep} ficheiro(s) com discrepâncias</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    # Botões lado a lado
    st.markdown('<div class="button-row">', unsafe_allow_html=True)
    col1, col2 = st.columns([1, 1])

    with col1:
        st.download_button(
            "⬇️ Descarregar resultados (ZIP)",
            data=st.session_state.zip_bytes,
            file_name=st.session_state.zip_name,
            mime="application/zip",
            key="zip_download_final"
        )

    with col2:
        if st.button("🔁 Novo processamento", key="btn_new_run"):
            with st.spinner("🔄 A reiniciar..."):
                for k in ["processing", "finished", "uploads", "results", "zip_bytes", "zip_name"]:
                    if k in st.session_state:
                        del st.session_state[k]
                time.sleep(0.6)
                st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)
