import streamlit as st
import tempfile, os, shutil, time
from pathlib import Path
from datetime import datetime
from xylella_processor import process_pdf_with_stats, build_zip_with_summary

# ───────────────────────────────────────────────
# Configuração
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor (Cloud)")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente um Excel por requisição.")

# ───────────────────────────────────────────────
# CSS — Tema laranja (#CA4300)
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
.stButton > button[kind="primary"]:focus {
  background-color: #A13700 !important;
  border: 1px solid #A13700 !important;
  color: #fff !important;
}
.stButton > button[kind="primary"][disabled] {
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
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Interface
# ───────────────────────────────────────────────
uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True)
if "processing" not in st.session_state:
    st.session_state.processing = False
start = st.button("📄 Processar ficheiros de Input", type="primary", disabled=st.session_state.processing or not uploads)

# ───────────────────────────────────────────────
# Execução
# ───────────────────────────────────────────────
if start and uploads:
    st.session_state.processing = True
    try:
        st.info("⚙️ A processar... isto pode demorar alguns segundos.")
        progress = st.progress(0)
        total = len(uploads)
        all_excel, all_debug, summary_lines = [], [], []
        total_amostras, total_reqs = 0, 0

        for i, up in enumerate(uploads, start=1):
            st.markdown(f"### 📄 {up.name}")
            st.write("⏳ Início de processamento...")
            tmpdir = tempfile.mkdtemp()
            tmp_path = Path(tmpdir) / up.name
            tmp_path.write_bytes(up.getbuffer())
            os.environ["OUTPUT_DIR"] = str(tmpdir)

            created, stats, dbg = process_pdf_with_stats(str(tmp_path))
            all_excel.extend(created)
            all_debug.extend(dbg)

            st.success(f"✅ {stats['pdf_name']}: {stats['req_count']} requisições, {stats['samples_total']} amostras.")
            for r in stats["per_req"]:
                diff = r.get("diff")
                msg = f" • Requisição {r['req']}: {r['processed']} processadas"
                if r["expected"] is not None:
                    msg += f" vs {r['expected']} declaradas"
                if diff and diff != 0:
                    msg += f" ⚠️ (diferença {diff:+d})"
                msg += f" → {Path(r['file']).name}"
                st.write(msg)

            total_amostras += stats["samples_total"]
            total_reqs += stats["req_count"]
            summary_lines.append(f"{stats['pdf_name']}: {stats['req_count']} requisições • {stats['samples_total']} amostras")
            progress.progress(i / total)
            time.sleep(0.3)

        # ZIP final
        if all_excel:
            summary_text = "\n".join(summary_lines)
            summary_text += f"\n\n📊 Total: {total_reqs} requisições | {total_amostras} amostras | {len(all_excel)} ficheiros Excel"
            zip_bytes, zip_name = build_zip_with_summary(all_excel, all_debug, summary_text)
            st.success("🏁 Processamento concluído.")
            st.download_button("⬇️ Descarregar resultados (ZIP)", data=zip_bytes,
                               file_name=f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip",
                               mime="application/zip")
            st.markdown("### 🧾 Resumo de execução")
            st.code(summary_text, language="markdown")
        else:
            st.error("⚠️ Nenhum ficheiro Excel foi criado.")

    except Exception as e:
        st.error(f"❌ Erro inesperado: {e}")
        st.exception(e)
    finally:
        st.session_state.processing = False
else:
    st.info("💡 Carrega ficheiros PDF e clica em **Processar ficheiros de Input**.")
