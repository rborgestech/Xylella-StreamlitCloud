# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, shutil, time
from pathlib import Path
from datetime import datetime
from xylella_processor import process_pdf, build_zip

st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 Excel por requisição.")

# ───────────────────────────────────────────────
# CSS (laranja SGS + ajustes visuais)
# ───────────────────────────────────────────────
st.markdown("""
<style>
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
[data-testid="stFileUploader"] > div:first-child {
  border: 2px dashed #CA4300 !important;
  border-radius: 10px !important;
  padding: 1rem !important;
}
/* Desativar botão "Browse files" enquanto processa */
.processing [data-testid="stFileUploader"] button {
  display: none !important;
}
/* Blocos de feedback */
.success-box {
  background-color: #E8F5E9;
  border-left: 5px solid #2E7D32;
  padding: 0.7rem 1rem;
  border-radius: 6px;
  margin-bottom: 0.4rem;
}
.warning-box {
  background-color: #FFF3E0;
  border-left: 5px solid #F57C00;
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
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Estado global
# ───────────────────────────────────────────────
if "processing" not in st.session_state:
    st.session_state.processing = False
if "finished" not in st.session_state:
    st.session_state.finished = False
if "all_excel" not in st.session_state:
    st.session_state.all_excel = []
if "uploads" not in st.session_state:
    st.session_state.uploads = []

# ───────────────────────────────────────────────
# File uploader — limpa automaticamente uploads antigos
# ───────────────────────────────────────────────
container_class = "processing" if st.session_state.processing else ""
st.markdown(f'<div class="{container_class}">', unsafe_allow_html=True)

uploads = st.file_uploader(
    "📂 Carrega um ou vários PDFs",
    type=["pdf"],
    accept_multiple_files=True,
    key="file_uploader"
)

# Se o utilizador carregar novos ficheiros, limpar estado anterior
if uploads and uploads != st.session_state.uploads:
    for key in ["finished", "all_excel"]:
        st.session_state[key] = False if key == "finished" else []
    st.session_state.uploads = uploads

st.markdown("</div>", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Interface principal
# ───────────────────────────────────────────────
if not st.session_state.processing and not st.session_state.finished:
    if st.session_state.uploads:
        start = st.button(
            f"📄 Processar {len(st.session_state.uploads)} ficheiro(s) de Input",
            type="primary"
        )
        if start:
            st.session_state.processing = True
            st.rerun()
    else:
        st.info("💡 Carrega ficheiros PDF para ativar o processamento.")

# ───────────────────────────────────────────────
# Execução principal (processamento)
# ───────────────────────────────────────────────
if st.session_state.processing and st.session_state.uploads:
    uploads = st.session_state.uploads
    st.markdown('<div class="info-box">⏳ A processar ficheiros... aguarde até o processo terminar.</div>', unsafe_allow_html=True)

    session_dir = tempfile.mkdtemp(prefix="xylella_session_")
    all_excel = []
    total = len(uploads)
    progress = st.progress(0)
    status_text = st.empty()

    try:
        for i, up in enumerate(uploads, start=1):
            status_text.markdown(
                f'<div class="info-box">📄 <b>A processar ficheiro {i}/{total}:</b> {up.name}</div>',
                unsafe_allow_html=True
            )

            tmpdir = tempfile.mkdtemp(dir=session_dir)
            tmp_path = os.path.join(tmpdir, up.name)
            with open(tmp_path, "wb") as f:
                f.write(up.getbuffer())

            os.environ["OUTPUT_DIR"] = tmpdir
            result = process_pdf(tmp_path)

            if isinstance(result, tuple) and len(result) == 3:
                created, n_amostras, discrepancias = result
            else:
                created, n_amostras, discrepancias = result, None, None

            if not created:
                st.markdown(
                    f'<div class="warning-box">⚠️ Nenhum ficheiro gerado para <b>{up.name}</b>.</div>',
                    unsafe_allow_html=True
                )
            else:
                for fp in created:
                    all_excel.append(fp)
                    msg = f"🟢 <b>{Path(fp).name}</b> processado com sucesso"
                    detalhes = []
                    if n_amostras is not None:
                        detalhes.append(f"{n_amostras} amostras")
                    if discrepancias is not None:
                        detalhes.append(f"{discrepancias} discrepâncias")
                    if detalhes:
                        msg += " — " + ", ".join(detalhes)
                    st.markdown(f'<div class="success-box">{msg}</div>', unsafe_allow_html=True)

            progress.progress(i / total)
            time.sleep(0.2)

        # ✅ Só mostra mensagem final se houver Excel gerados
        if len(all_excel) > 0:
            st.session_state.all_excel = all_excel
            st.session_state.finished = True
            st.markdown(
                f'<div class="success-box">✅ Processamento concluído '
                f'({len(all_excel)} ficheiro{"s" if len(all_excel)>1 else ""} Excel gerado{"s" if len(all_excel)>1 else ""}).</div>',
                unsafe_allow_html=True
            )
        else:
            st.warning("⚠️ Nenhum ficheiro Excel foi detetado.")

    except Exception as e:
        st.error(f"❌ Erro inesperado: {e}")
    finally:
        shutil.rmtree(session_dir, ignore_errors=True)
        st.session_state.processing = False

# ───────────────────────────────────────────────
# Download final (sem botão de novo processamento)
# ───────────────────────────────────────────────
if st.session_state.finished:
    all_excel = st.session_state.all_excel
    zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
    zip_bytes = build_zip(all_excel)

    st.download_button(
        "⬇️ Descarregar resultados (ZIP)",
        data=zip_bytes,
        file_name=zip_name,
        mime="application/zip",
        key="download_zip"
    )
