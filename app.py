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
# CSS — tema SGS + animações
# ───────────────────────────────────────────────
st.markdown("""
<style>
/* Botão primário laranja SGS */
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

/* Caixas de estado */
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

/* Loader animado "..." em laranja SGS */
.st-processing-dots::after {
  content: ' ';
  animation: dots 1.2s steps(4, end) infinite;
  color: #CA4300;
  font-weight: 700;
  margin-left: .15rem;
}
@keyframes dots {
  0%, 20%   { content: ''; }
  40%       { content: '.'; }
  60%       { content: '..'; }
  80%, 100% { content: '...'; }
}

/* Linha de botões finais lado a lado */
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

/* Fade-in no painel final */
@keyframes fadeIn {
  from { opacity: 0; transform: translateY(10px); }
  to { opacity: 1; transform: translateY(0); opacity: 1; }
}
.fade-in {
  animation: fadeIn 0.8s ease-in-out;
}
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Estado
# ───────────────────────────────────────────────
for k in ["processing", "finished", "uploads", "all_excel", "zip_bytes", "zip_name"]:
    if k not in st.session_state:
        st.session_state[k] = False if k in ["processing", "finished"] else []

# ───────────────────────────────────────────────
# Ecrã inicial
# ───────────────────────────────────────────────
if not st.session_state.processing and not st.session_state.finished:
    uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True)

    if uploads:
        st.session_state.uploads = uploads
        start = st.button(f"📄 Processar {len(uploads)} ficheiro(s) de Input", type="primary")
        if start:
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

    with st.expander("📄 Ficheiros em processamento", expanded=True):
        for up in uploads:
            st.markdown(f"- {up.name}")

    generated_panel = st.expander("📄 Ficheiros gerados", expanded=True)
    progress = st.progress(0)
    status_text = st.empty()
    all_excel = []
    total_amostras = 0
    ficheiros_com_discrepancias = 0
    session_dir = tempfile.mkdtemp(prefix="xylella_session_")

    try:
        for i, up in enumerate(uploads, start=1):
            status_text.markdown(
                f'<div class="info-box">📘 <b>A processar ficheiro {i}/{total}</b>'
                f'<span class="st-processing-dots"></span><br>{up.name}</div>',
                unsafe_allow_html=True
            )

            tmpdir = tempfile.mkdtemp(dir=session_dir)
            tmp_path = os.path.join(tmpdir, up.name)
            with open(tmp_path, "wb") as f:
                f.write(up.getbuffer())

            os.environ["OUTPUT_DIR"] = tmpdir
            result = process_pdf(tmp_path)

            if not result:
                generated_panel.markdown(
                    f'<div class="warning-box">⚠️ Nenhum ficheiro gerado para <b>{up.name}</b>.</div>',
                    unsafe_allow_html=True
                )
            else:
                # Cada item pode ser (path, n_amostras, discrepancias)
                for item in result:
                    if isinstance(item, tuple):
                        fp, n_amostras, discrepancias = item
                    else:
                        fp, n_amostras, discrepancias = item, None, None

                    all_excel.append(fp)
                    base_name = Path(fp).name
                    if n_amostras:
                        total_amostras += n_amostras
                    if discrepancias:
                        ficheiros_com_discrepancias += 1

                    # Mensagem
                    if discrepancias:
                        msg = (
                            f"⚠️ <b>{base_name}</b>: ficheiro gerado. "
                            f"<span style='color:#F57C00;'>⚠️ discrepância detectada ({discrepancias})</span>"
                        )
                        css_class = "warning-box"
                    else:
                        amostras_txt = (
                            f"({n_amostras} amostra{'s' if n_amostras != 1 else ''} OK)"
                            if n_amostras else ""
                        )
                        msg = f"✅ <b>{base_name}</b>: ficheiro gerado. {amostras_txt}"
                        css_class = "success-box"

                    generated_panel.markdown(f'<div class="{css_class}">{msg}</div>', unsafe_allow_html=True)

            progress.progress(i / total)
            time.sleep(0.2)

        status_text.empty()

        # Criar ZIP e finalizar
        if all_excel:
            with st.spinner("🧩 A gerar ficheiro ZIP… aguarde alguns segundos."):
                zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
                zip_bytes = build_zip(all_excel)
                st.session_state.update({
                    "finished": True,
                    "processing": False,
                    "all_excel": all_excel,
                    "zip_name": zip_name,
                    "zip_bytes": zip_bytes,
                    "total_amostras": total_amostras,
                    "ficheiros_com_discrepancias": ficheiros_com_discrepancias
                })
                st.rerun()
        else:
            st.warning("⚠️ Nenhum ficheiro Excel foi detetado.")

    except Exception as e:
        st.error(f"❌ Erro inesperado: {e}")
    finally:
        shutil.rmtree(session_dir, ignore_errors=True)

# ───────────────────────────────────────────────
# Ecrã final — painel de sucesso + botões lado a lado
# ───────────────────────────────────────────────
if st.session_state.finished and st.session_state.all_excel:
    num_files = len(st.session_state.all_excel)
    total_amostras = st.session_state.get("total_amostras", 0)

    st.markdown(
        f"""
        <div class="fade-in" style="
          background:#E8F5E9; border-left:6px solid #2E7D32; border-radius:10px;
          padding:1.2rem 1.6rem; margin-top:1.4rem; text-align:center;
        ">
          <h4 style="color:#2E7D32; font-weight:600; margin:.2rem 0 .3rem 0;">
            ✅ Processamento concluído
          </h4>
          <p style="color:#2E7D32; margin:.2rem 0 0 0;">
            {num_files} ficheiro{'s' if num_files>1 else ''} Excel gerado{'s' if num_files>1 else ''} · 
            {total_amostras} amostra{'s' if total_amostras != 1 else ''} no total
          </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    zip_name = st.session_state.zip_name
    zip_bytes = st.session_state.zip_bytes

    st.markdown('<div class="button-row">', unsafe_allow_html=True)
    col1, col2 = st.columns([1, 1])

    with col1:
        st.download_button(
            "⬇️ Descarregar resultados (ZIP)",
            data=zip_bytes,
            file_name=zip_name,
            mime="application/zip",
            key="zip_download_final"
        )

    with col2:
        if st.button("🔁 Novo processamento", key="btn_new_run"):
            with st.spinner("🔄 A reiniciar..."):
                for k in ["processing", "finished", "uploads", "all_excel", "zip_bytes", "zip_name"]:
                    if k in st.session_state:
                        del st.session_state[k]
                time.sleep(0.6)
                st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)
