import streamlit as st
import tempfile, os, zipfile, traceback   # ← adiciona traceback
from datetime import datetime

from xylella_processor import process_pdf, write_to_template

st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor (Cloud)")
st.caption("Faz upload de um ou vários PDFs. Vou gerar um Excel por requisição.")

uploaded = st.file_uploader("Carrega os PDFs", type=["pdf"], accept_multiple_files=True)
expected = st.text_input("Nº de requisições esperadas (opcional)", "")
start = st.button("📄 Processar ficheiros de Input", type="primary", disabled=not uploaded)

if start:
    with st.spinner("A processar…"):
        tmp = tempfile.mkdtemp()
        outdir = os.path.join(tmp, "output")
        os.makedirs(outdir, exist_ok=True)
        logs, ok, fail = [], 0, 0

        for up in uploaded:
            try:
                in_path = os.path.join(tmp, up.name)
                with open(in_path, "wb") as f:
                    f.write(up.read())
                rows = process_pdf(in_path)
                exp = int(expected) if expected.strip().isdigit() else None
                base = os.path.splitext(up.name)[0]
                write_to_template(rows, os.path.join(outdir, base), expected_count=exp, source_pdf=up.name)
                logs.append(f"✅ {up.name}: concluído")
                ok += 1
            except Exception:
                logs.append(f"❌ {up.name}:\n{traceback.format_exc()}")
                fail += 1

        zip_path = os.path.join(tmp, f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(outdir):
                for fn in files:
                    full = os.path.join(root, fn)
                    z.write(full, os.path.relpath(full, outdir))

    st.success(f"Concluído • {ok} ok, {fail} com erro.")
    with open(zip_path, "rb") as f:
        st.download_button("⬇️ Descarregar resultados (ZIP)", f, file_name=os.path.basename(zip_path))
    with st.expander("Registo de execução"):
        st.code("\n".join(logs) if logs else "Sem logs a apresentar.")
