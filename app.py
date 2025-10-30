import streamlit as st
import tempfile, os, shutil
from pathlib import Path
from xylella_processor import process_pdf_with_stats, build_zip

# ───────────────────────────────────────────────
# Configuração base
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 Excel por requisição.")

# ───────────────────────────────────────────────
# CSS — estilo laranja (#CA4300) sem vermelhos
# ───────────────────────────────────────────────
st.markdown("""
<style>
/* 🔸 Botão principal laranja */
.stButton > button[kind="primary"] {
  background: #CA4300 !important;
  border-color: #CA4300 !important;
  color: #ffffff !important;
  box-shadow: none !important;
  border-radius: 6px !important;
  font-weight: 600 !important;
}

/* 🔸 Hover / Focus / Active mais escuro */
.stButton > button[kind="primary"]:hover,
.stButton > button[kind="primary"]:focus,
.stButton > button[kind="primary"]:active {
  background: #A13700 !important;
  border-color: #A13700 !important;
  color: #ffffff !important;
  outline: none !important;
  box-shadow: none !important;
}

/* 🔸 Estado desativado = cinzento */
.stButton > button[kind="primary"][disabled],
.stButton > button[kind="primary"][disabled]:hover {
  background: #b3b3b3 !important;
  border-color: #b3b3b3 !important;
  color: #f2f2f2 !important;
  cursor: not-allowed !important;
  box-shadow: none !important;
}

/* 🔸 File uploader (sem vermelho nem foco) */
[data-testid="stFileUploader"] > div:first-child {
  border: 2px dashed #CA4300 !important;
  border-radius: 10px !important;
  padding: 1rem !important;
  transition: border-color 0.3s ease-in-out;
}

[data-testid="stFileUploader"] > div:first-child:hover {
  border-color: #A13700 !important;
}

[data-testid="stFileUploader"] > div:focus-within {
  border-color: #CA4300 !important;
  box-shadow: none !important;
  outline: none !important;
}

/* 🔸 Remover foco vermelho global */
:root {
  --primary-color: #CA4300 !important;
  --text-selection-color: #CA4300 !important;
  --accent-color: #CA4300 !important;
  --focus-ring: 0 0 0 0 rgba(0,0,0,0) !important;
}
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Interface de Upload
# ───────────────────────────────────────────────
uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True)

if "processing" not in st.session_state:
    st.session_state.processing = False

btn = st.button("📄 Processar ficheiros de Input", type="primary", disabled=st.session_state.processing)

# ───────────────────────────────────────────────
# Execução principal
# ───────────────────────────────────────────────
if btn and uploads:
    st.session_state.processing = True
    try:
        st.info("⚙️ A processar... aguarda alguns segundos.")
        all_excel = []
        all_stats = []

        # Criar diretório persistente
        final_dir = Path.cwd() / "output_final"
        final_dir.mkdir(exist_ok=True)

        for up in uploads:
            st.markdown(f"### 📄 {up.name}")
            st.write("⏳ Início de processamento...")

            tmpdir = tempfile.mkdtemp()
            tmp_path = os.path.join(tmpdir, up.name)
            with open(tmp_path, "wb") as f:
                f.write(up.getbuffer())

            os.environ["OUTPUT_DIR"] = tmpdir
            files, stats = process_pdf_with_stats(tmp_path)

            # copiar ficheiros Excel para diretório persistente
            for fp in files:
                dest = final_dir / Path(fp).name
                if os.path.exists(fp):
                    shutil.copy(fp, dest)
                    all_excel.append(str(dest))

            all_stats.append(stats)
            st.write(f"✅ {up.name}: {stats['req_count']} requisições, {stats['samples_total']} amostras.")
            for item in stats["per_req"]:
                msg = f" • Requisição {item['req']}: {item['samples']} amostras → {Path(item['file']).name}"
                if item["diff"]:
                    sign = "+" if item["diff"] > 0 else ""
                    msg += f" ⚠️ discrepância {sign}{item['diff']} (decl={item['expected']})"
                st.write(msg)

        if all_excel:
            zip_bytes = build_zip(all_excel, all_stats)
            st.success("🏁 Processamento concluído.")
            st.download_button("⬇️ Descarregar resultados (ZIP)", data=zip_bytes,
                               file_name="xylella_resultados.zip", mime="application/zip")

    finally:
        st.session_state.processing = False

else:
    st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")
