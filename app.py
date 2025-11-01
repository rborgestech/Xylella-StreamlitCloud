# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, shutil, time
from pathlib import Path
from datetime import datetime
from xylella_processor import process_pdf, build_zip

st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 Excel por requisição.")

# ————— CSS —————
st.markdown("""
<style>
.success-box{background:#E8F5E9;border-left:5px solid #2E7D32;padding:.7rem 1rem;border-radius:6px;margin:.35rem 0}
.warning-box{background:#FFF8E1;border-left:5px solid #FBC02D;padding:.7rem 1rem;border-radius:6px;margin:.35rem 0}
.info-box{background:#E3F2FD;border-left:5px solid #1E88E5;padding:.7rem 1rem;border-radius:6px;margin:.35rem 0}
.button-row{display:flex;gap:1rem;justify-content:center;margin-top:1rem}
.stDownloadButton button,.stButton button{background:#fff!important;border:1.5px solid #CA4300!important;color:#CA4300!important;font-weight:600!important;border-radius:8px!important;padding:.6rem 1.2rem!important}
.stDownloadButton button:hover,.stButton button:hover{background:#CA4300!important;color:#fff!important}
</style>
""", unsafe_allow_html=True)

# ————— Ecrã principal —————
uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True)

if uploads:
    if st.button(f"📄 Processar {len(uploads)} ficheiro(s) de Input"):
        progress = st.progress(0)
        st.markdown('<div class="info-box">⏳ A processar... aguarde até o processo terminar.</div>', unsafe_allow_html=True)

        all_results = []
        session_dir = tempfile.mkdtemp(prefix="xylella_session_")

        try:
            for i, up in enumerate(uploads, start=1):
                tmp_path = os.path.join(session_dir, up.name)
                with open(tmp_path, "wb") as f:
                    f.write(up.getbuffer())

                st.markdown(f"### 📄 {up.name}")
                st.write("⏳ Início de processamento...")

                results = process_pdf(tmp_path)
                if not results:
                    st.markdown(f'<div class="warning-box">⚠️ Nenhum ficheiro gerado para {up.name}</div>', unsafe_allow_html=True)
                else:
                    for r in results:
                        base = Path(r["path"]).name
                        proc = r.get("processed") or 0
                        disc = r.get("discrepancy")
                        if disc:
                            st.markdown(f'<div class="warning-box">⚠️ {base}: discrepância detectada.</div>', unsafe_allow_html=True)
                        else:
                            st.markdown(f'<div class="success-box">✅ {base}: {proc} amostras OK.</div>', unsafe_allow_html=True)
                        all_results.append(r)

                progress.progress(i / len(uploads))
                time.sleep(0.3)

            if all_results:
                zip_bytes = build_zip(all_results)
                zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"

                st.success(f"🏁 Processamento concluído ({len(all_results)} ficheiros Excel gerados).")
                st.download_button("⬇️ Descarregar resultados (ZIP)",
                                   data=zip_bytes,
                                   file_name=zip_name,
                                   mime="application/zip")
                st.balloons()
            else:
                st.error("⚠️ Nenhum ficheiro Excel foi detetado.")

        except Exception as e:
            st.error(f"❌ Erro inesperado: {e}")

        finally:
            shutil.rmtree(session_dir, ignore_errors=True)
else:
    st.info("💡 Carrega ficheiros PDF e clica em **Processar ficheiros de Input**.")
