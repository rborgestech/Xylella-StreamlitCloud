# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, shutil, time
from pathlib import Path
from datetime import datetime
from xylella_processor import process_pdf, build_zip

# ───────────────────────────────────────────────
# Configuração base
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🦢", layout="centered")
st.title("\ud83e\udda2 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 Excel por requisição.")

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

    st.markdown('''<div style="background-color:#FFF3E0; border-left: 5px solid #F57C00; padding:0.7rem 1rem; border-radius: 6px; margin-bottom: 0.4rem;">
    ⏳ A processar ficheiros... aguarde até o processo terminar.</div>''', unsafe_allow_html=True)

    with st.expander("📄 Ficheiros em processamento", expanded=True):
        for up in uploads:
            st.markdown(f"- {up.name}")

    progress = st.progress(0)
    status_text = st.empty()
    all_excel = []
    session_dir = tempfile.mkdtemp(prefix="xylella_session_")

    try:
        for i, up in enumerate(uploads, start=1):
            status_text.markdown(
                f'''<div style="background-color:#E3F2FD; border-left:5px solid #1E88E5; padding:0.7rem 1rem; border-radius:6px; margin-bottom:0.4rem;">
                📘 <b>A processar ficheiro {i}/{total}</b>: {up.name}</div>''',
                unsafe_allow_html=True
            )

            tmpdir = tempfile.mkdtemp(dir=session_dir)
            tmp_path = os.path.join(tmpdir, up.name)
            with open(tmp_path, "wb") as f:
                f.write(up.getbuffer())

            os.environ["OUTPUT_DIR"] = tmpdir
            result = process_pdf(tmp_path)

            created = [r["path"] for r in result if isinstance(r, dict) and r.get("path")]
            n_amostras = sum(r.get("processed", 0) for r in result)
            discrepancias = sum(1 for r in result if r.get("discrepancy"))

            if not created:
                st.markdown(f'''<div style="background-color:#FFF3E0; border-left:5px solid #F57C00; padding:0.7rem 1rem; border-radius:6px; margin-bottom:0.4rem;">
                ⚠️ Nenhum ficheiro gerado para <b>{up.name}</b>.</div>''', unsafe_allow_html=True)
            else:
                for fp in created:
                    all_excel.append(fp)
                st.markdown(f'''<div style="background-color:#E8F5E9; border-left:5px solid #2E7D32; padding:0.7rem 1rem; border-radius:6px; margin-bottom:0.4rem;">
                ✅ <b>{Path(up.name).stem}</b>: {n_amostras} amostras, {discrepancias} discrepâncias</div>''', unsafe_allow_html=True)

            progress.progress(i / total)
            time.sleep(0.2)

        status_text.empty()
        with st.spinner("🧩 A gerar ficheiro ZIP… aguarde alguns segundos."):
            time.sleep(0.5)
            if all_excel:
                st.session_state.all_excel = all_excel
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
# Ecrã final — painel de sucesso + botões lado a lado
# ───────────────────────────────────────────────
if st.session_state.finished and st.session_state.all_excel:
    num_files = len(st.session_state.all_excel)

    st.success(f"Processamento concluído: {num_files} ficheiro(s) Excel gerado(s).")

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
            for k in ["processing", "finished", "uploads", "all_excel", "zip_bytes", "zip_name"]:
                if k in st.session_state:
                    del st.session_state[k]
            st.rerun()
