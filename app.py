# app.py — versão final (Streamlit Cloud)

import streamlit as st
import tempfile, os, traceback
from datetime import datetime
from xylella_processor import process_pdf, build_zip

# ───────────────────────────────────────────────
# Configuração base do Streamlit
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor (Cloud)")
st.caption("Faz upload de um ou vários PDFs. O sistema gera automaticamente um Excel por requisição.")

# ───────────────────────────────────────────────
# Interface de Upload
# ───────────────────────────────────────────────
uploaded = st.file_uploader("📤 Carrega os PDFs", type=["pdf"], accept_multiple_files=True)
start = st.button("📄 Processar ficheiros de Input", type="primary", disabled=not uploaded)

# ───────────────────────────────────────────────
# Processamento principal
# ───────────────────────────────────────────────
if start:
    with st.spinner("⚙️ A processar os ficheiros... Isto pode demorar alguns segundos."):

        # Cria diretório temporário e de saída
        tmp = tempfile.mkdtemp()
        outdir = os.path.join(tmp, "output")
        os.makedirs(outdir, exist_ok=True)
        os.environ["OUTPUT_DIR"] = outdir

        logs, ok, fail = [], 0, 0
        created_all = []

        # ── Loop pelos PDFs carregados ───────────────────────────────
        for up in uploaded:
            try:
                in_path = os.path.join(tmp, up.name)
                with open(in_path, "wb") as f:
                    f.write(up.read())

                st.markdown(f"### 🧾 {up.name}")
                st.write("⏳ Início de processamento...")

                # Processa PDF → devolve lista de ficheiros gerados (.xlsx)
                req_files = process_pdf(in_path)
                if not req_files:
                    st.warning(f"⚠️ Nenhum ficheiro gerado para {up.name}")
                    continue

                created_all.extend(req_files)
                for fpath in req_files:
                    fname = os.path.basename(fpath)
                    st.success(f"✅ {fname} gravado")

                ok += 1

            except Exception as e:
                err = traceback.format_exc()
                logs.append(f"❌ {up.name}:\n{err}")
                st.error(f"❌ Erro ao processar {up.name}: {e}")
                fail += 1

        # ── Criação do ZIP final ─────────────────────────────────────
        if created_all:
            zip_bytes = build_zip(created_all)
            zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
            zip_path = os.path.join(tmp, zip_name)
            with open(zip_path, "wb") as f:
                f.write(zip_bytes)

            st.success(f"🏁 Processamento concluído • {ok} ok, {fail} com erro.")
            with open(zip_path, "rb") as f:
                st.download_button("⬇️ Descarregar resultados (ZIP)", f, file_name=os.path.basename(zip_path))

        else:
            st.error("❌ Nenhum ficheiro .xlsx foi criado.")

    # ───────────────────────────────────────────────
    # Log final (expansível)
    # ───────────────────────────────────────────────
    with st.expander("🧾 Registo de execução"):
        if logs:
            st.code("\n".join(logs))
        else:
            st.info("Sem erros a reportar.")
else:
    st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")
