import streamlit as st
import tempfile, os, traceback
from datetime import datetime
from xylella_processor import process_pdf, build_zip_with_summary

# ───────────────────────────────────────────────
# Configuração base
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor (Cloud)")
st.caption("Faz upload de um ou vários PDFs. O sistema gera automaticamente um Excel por requisição.")

# ───────────────────────────────────────────────
# CSS — botão laranja, hover escuro, sem vermelho
# ───────────────────────────────────────────────
st.markdown("""
<style>
.stButton > button[kind="primary"] {
  background-color: #CA4300 !important;
  border: 1px solid #CA4300 !important;
  color: white !important;
  font-weight: 600 !important;
  border-radius: 6px !important;
}
.stButton > button[kind="primary"]:hover,
.stButton > button[kind="primary"]:focus,
.stButton > button[kind="primary"]:active {
  background-color: #A13700 !important;
  border-color: #A13700 !important;
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
# Upload
# ───────────────────────────────────────────────
uploaded = st.file_uploader("📤 Carrega os PDFs", type=["pdf"], accept_multiple_files=True)
start = st.button("📄 Processar ficheiros de Input", type="primary", disabled=not uploaded)

# ───────────────────────────────────────────────
# Processamento
# ───────────────────────────────────────────────
if start:
    st.session_state["processing"] = True
    with st.spinner("⚙️ A processar... isto pode demorar alguns segundos."):
        tmp = tempfile.mkdtemp()
        outdir = os.path.join(tmp, "output")
        os.makedirs(outdir, exist_ok=True)
        os.environ["OUTPUT_DIR"] = outdir

        logs, ok, fail = [], 0, 0
        created_all = []
        summary_data = []

        for up in uploaded:
            try:
                in_path = os.path.join(tmp, up.name)
                with open(in_path, "wb") as f:
                    f.write(up.read())

                st.markdown(f"### 🧾 {up.name}")
                st.write("⏳ Início de processamento...")

                req_files = process_pdf(in_path)
                if not req_files:
                    st.warning(f"⚠️ Nenhum ficheiro gerado para {up.name}")
                    continue

                created_all.extend(req_files)

                per_req = []
                for i, fpath in enumerate(req_files, start=1):
                    st.success(f"✅ {os.path.basename(fpath)} gravado")
                    per_req.append({
                        "req": i,
                        "samples": len(req_files[i-1]) if hasattr(req_files[i-1], "__len__") else "—",
                        "expected": None,  # valor pode vir do parser no futuro
                        "file": fpath
                    })

                summary_data.append({
                    "pdf": up.name,
                    "req_count": len(req_files),
                    "samples_total": sum([len(req_files[i-1]) for i in range(1, len(req_files)+1)
                                          if hasattr(req_files[i-1], "__len__")]),
                    "per_req": per_req
                })

                ok += 1
            except Exception as e:
                err = traceback.format_exc()
                logs.append(f"❌ {up.name}:\n{err}")
                st.error(f"❌ Erro ao processar {up.name}: {e}")
                fail += 1

        # ZIP final com summary.txt detalhado
        if created_all:
            zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
            zip_bytes = build_zip_with_summary(created_all, summary_data)
            st.success(f"🏁 Processamento concluído • {ok} ok, {fail} com erro.")
            st.download_button("⬇️ Descarregar resultados (ZIP)",
                               data=zip_bytes,
                               file_name=zip_name,
                               mime="application/zip")
            st.balloons()
        else:
            st.error("❌ Nenhum ficheiro .xlsx foi criado.")

    st.session_state["processing"] = False

else:
    st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")
