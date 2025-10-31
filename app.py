# app.py — versão final consolidada (Streamlit Cloud)
import streamlit as st
import tempfile, os, shutil, time, traceback
from pathlib import Path
from datetime import datetime
from xylella_processor import process_pdf, build_zip

# ───────────────────────────────────────────────
# Configuração base
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente um Excel por requisição.")

# ───────────────────────────────────────────────
# CSS — tema laranja (#CA4300) sem vermelhos
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
.stButton > button[kind="primary"]:hover,
.stButton > button[kind="primary"]:focus,
.stButton > button[kind="primary"]:active {
  background-color: #A13700 !important;
  border: 1px solid #A13700 !important;
  color: #fff !important;
  box-shadow: none !important;
}
.stButton > button[kind="primary"][disabled],
.stButton > button[kind="primary"][disabled]:hover {
  background-color: #b3b3b3 !important;
  border: 1px solid #b3b3b3 !important;
  color: #f2f2f2 !important;
  cursor: not-allowed !important;
}
[data-testid="stFileUploader"] > div:first-child {
  border: 2px dashed #CA4300 !important;
  border-radius: 10px !important;
  padding: 1rem !important;
}
[data-testid="stFileUploader"] > div:first-child:hover {
  border-color: #A13700 !important;
}
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Interface de upload
# ───────────────────────────────────────────────
uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True)

if "processing" not in st.session_state:
    st.session_state.processing = False

start = st.button("📄 Processar ficheiros de Input", type="primary", disabled=st.session_state.processing or not uploads)

# ───────────────────────────────────────────────
# Execução principal
# ───────────────────────────────────────────────
if start and uploads:
    st.session_state.processing = True
    try:
        st.info("⚙️ A processar... aguarda alguns segundos.")
        all_excel = []
        summary_lines = []
        total_amostras, total_reqs = 0, 0

        final_dir = Path.cwd() / "output_final"
        debug_dir = final_dir / "debug"
        final_dir.mkdir(exist_ok=True)
        debug_dir.mkdir(exist_ok=True)

        progress = st.progress(0)
        total_pdfs = len(uploads)

        for i, up in enumerate(uploads, start=1):
            st.markdown(f"### 📄 {up.name}")
            st.write("⏳ Início de processamento...")

            tmpdir = tempfile.mkdtemp()
            tmp_path = os.path.join(tmpdir, up.name)
            with open(tmp_path, "wb") as f:
                f.write(up.getbuffer())

            os.environ["OUTPUT_DIR"] = str(tmpdir)
            created = process_pdf(tmp_path)

            if not created:
                st.warning(f"⚠️ Nenhum ficheiro gerado para {up.name}")
                summary_lines.append(f"⚠️ {up.name}: sem ficheiros gerados.")
            else:
                st.success(f"✅ {up.name}: {len(created)} ficheiro(s) Excel criado(s).")
                summary_lines.append(f"{up.name}: {len(created)} requisição(ões).")
                total_reqs += len(created)

                for fp in created:
                    fname = Path(fp).name
                    dest = final_dir / fname
                    shutil.copy(fp, dest)
                    all_excel.append(str(dest))
                    st.write(f" • {fname}")

            # Mover ficheiros de debug (txt, csv) para subpasta
            for f in Path(tmpdir).glob("*_ocr_debug.txt"):
                shutil.move(f, debug_dir / f.name)
            for f in Path(tmpdir).glob("*.csv"):
                shutil.move(f, debug_dir / f.name)

            progress.progress(i / total_pdfs)
            time.sleep(0.2)

        # ── ZIP final com timestamp ───────────────────────────────
        if all_excel:
            zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
            zip_bytes = build_zip(all_excel)

            st.success(f"🏁 Processamento concluído ({len(all_excel)} ficheiros Excel gerados).")
            st.download_button("⬇️ Descarregar resultados (ZIP)", data=zip_bytes,
                               file_name=zip_name, mime="application/zip")

            # ── Sumário final ───────────────────────────────
            st.markdown("### 🧾 Resumo de execução")
            summary_text = "\n".join(summary_lines)
            summary_text += f"\n\n📊 Total: {total_reqs} requisições | {len(all_excel)} ficheiros Excel"
            st.code(summary_text, language="markdown")

        else:
            st.error("⚠️ Nenhum ficheiro Excel foi detetado para incluir no ZIP.")

    except Exception as e:
        st.error(f"❌ Erro inesperado: {e}")
        st.exception(e)

    finally:
        st.session_state.processing = False

else:
    st.info("💡 Carrega ficheiros PDF e clica em **Processar ficheiros de Input**.")
