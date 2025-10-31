# app.py — versão final estável (Streamlit Cloud)
import streamlit as st
import tempfile, os, shutil, time
from pathlib import Path
from datetime import datetime
from xylella_processor import process_pdf_with_stats, build_zip_with_summary

# ───────────────────────────────────────────────
# Configuração base e estilo
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 ficheiro Excel por requisição.")

# CSS — tons laranja (#CA4300)
st.markdown("""
<style>
/* Botão principal */
.stButton > button[kind="primary"] {
  background-color: #CA4300 !important;
  border: 1px solid #CA4300 !important;
  color: #fff !important;
  font-weight: 600 !important;
  border-radius: 6px !important;
  transition: background-color 0.2s ease-in-out !important;
}
/* Hover / Focus */
.stButton > button[kind="primary"]:hover,
.stButton > button[kind="primary"]:focus {
  background-color: #A13700 !important;
  border: 1px solid #A13700 !important;
}
/* Disabled */
.stButton > button[kind="primary"][disabled],
.stButton > button[kind="primary"][disabled]:hover {
  background-color: #b3b3b3 !important;
  border: 1px solid #b3b3b3 !important;
  color: #f2f2f2 !important;
  cursor: not-allowed !important;
}
/* File uploader */
[data-testid="stFileUploader"] > div:first-child {
  border: 2px dashed #CA4300 !important;
  border-radius: 10px !important;
  padding: 1rem !important;
}
[data-testid="stFileUploader"] > div:first-child:hover {
  border-color: #A13700 !important;
}
[data-testid="stFileUploader"] > div:focus-within {
  border-color: #CA4300 !important;
  box-shadow: none !important;
}
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Interface de upload
# ───────────────────────────────────────────────
uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True)

if "processing" not in st.session_state:
    st.session_state.processing = False

start = st.button("📄 Processar ficheiros de Input", type="primary", disabled=st.session_state.processing)

# ───────────────────────────────────────────────
# Execução principal
# ───────────────────────────────────────────────
if start and uploads:
    st.session_state.processing = True
    try:
        st.info("⚙️ A processar... aguarda alguns segundos.")
        all_excel, all_debug, all_stats = [], [], []

        # diretório de saída
        final_dir = Path.cwd() / "output_final"
        final_dir.mkdir(exist_ok=True)

        progress = st.progress(0)
        total = len(uploads)
        summary_lines = []

        for i, up in enumerate(uploads, start=1):
            st.markdown(f"### 📄 {up.name}")
            st.write("⏳ Início de processamento...")

            tmpdir = tempfile.mkdtemp()
            tmp_path = os.path.join(tmpdir, up.name)
            with open(tmp_path, "wb") as f:
                f.write(up.getbuffer())

            os.environ["OUTPUT_DIR"] = tmpdir
            created, stats, debug = process_pdf_with_stats(tmp_path)

            # copiar resultados
            for fp in created:
                if os.path.exists(fp):
                    dest = final_dir / Path(fp).name
                    shutil.copy(fp, dest)
                    all_excel.append(str(dest))

            for dbg in debug:
                if os.path.exists(dbg):
                    all_debug.append(dbg)

            all_stats.append(stats)
            reqs = stats["req_count"]
            samples = stats["samples_total"]
            st.success(f"✅ {up.name}: {reqs} requisição(ões), {samples} amostras.")

            # detalhar por requisição
            for p in stats["per_req"]:
                line = f"• Requisição {p['req']}: {p['samples']} amostras → {Path(p['file']).name}"
                if p["diff"] is not None and p["diff"] != 0:
                    sign = "+" if p["diff"] > 0 else ""
                    line += f" ⚠️ discrepância {sign}{p['diff']} ({p['samples']} processadas / {p['expected']} declaradas)"
                st.write(line)

            summary_lines.append(
                f"{up.name}: {reqs} requisição(ões), {samples} amostras."
            )

            progress.progress(i / total)
            time.sleep(0.3)

        # ───────────────────────────────────────────────
        # ZIP final
        # ───────────────────────────────────────────────
        if all_excel:
            summary_text = "\n".join(summary_lines) + f"\n\n📊 Total: {sum(s['req_count'] for s in all_stats)} requisições | {len(all_excel)} ficheiros Excel"
            zip_bytes, zip_name = build_zip_with_summary(all_excel, all_debug, summary_text)
            st.success(f"🏁 Processamento concluído ({len(all_excel)} ficheiros Excel gerados).")
            st.download_button("⬇️ Descarregar resultados (ZIP)", data=zip_bytes,
                               file_name=f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip",
                               mime="application/zip")
        else:
            st.error("⚠️ Nenhum ficheiro Excel foi criado.")

    finally:
        st.session_state.processing = False
else:
    st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")
