import streamlit as st
import tempfile, os, zipfile, traceback
from datetime import datetime
from pathlib import Path

# ⚙️ Garante que o core grava dentro da pasta temporária
os.environ["OUTPUT_DIR"] = tempfile.mkdtemp()

from xylella_processor import process_pdf, write_to_template

# ───────────────────────────────────────────────
# Configuração base do Streamlit
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor (Cloud)")
st.caption("Faz upload de um ou vários PDFs. Vou gerar automaticamente um Excel por requisição.")

# ───────────────────────────────────────────────
# Interface de Upload
# ───────────────────────────────────────────────
uploaded = st.file_uploader("📤 Carrega os PDFs", type=["pdf"], accept_multiple_files=True)
start = st.button("📄 Processar ficheiros de Input", type="primary", disabled=not uploaded)

# ───────────────────────────────────────────────
# Processamento
# ───────────────────────────────────────────────
# ───────────────────────────────────────────────
# Processamento
# ───────────────────────────────────────────────
if start:
    with st.spinner("⚙️ A processar os ficheiros... Isto pode demorar alguns segundos."):

        tmp = tempfile.mkdtemp()
        outdir = os.path.join(tmp, "output")
        os.makedirs(outdir, exist_ok=True)

        logs, ok, fail = [], 0, 0

        # ⚙️ Garante que o core grava dentro da pasta temporária
        os.environ["OUTPUT_DIR"] = outdir  

        for up in uploaded:
            try:
                in_path = os.path.join(tmp, up.name)
                with open(in_path, "wb") as f:
                    f.write(up.read())

                # processa PDF (pode conter várias requisições)
                rows_per_req = process_pdf(in_path)

                base = os.path.splitext(up.name)[0]
                total_amostras = 0

                # ───────────────
                # Cria 1 ficheiro Excel por requisição
                # ───────────────
                for i, req_rows in enumerate(rows_per_req, start=1):
                    if not req_rows:
                        continue

                    out_name = f"{base}_req{i}.xlsx"
                    out_path = os.path.join(outdir, out_name)

                    # 🔄 Limpeza do template antes de cada escrita
                    from shutil import copyfile
                    template_copy = os.path.join(outdir, f"_tmp_template_req{i}.xlsx")
                    copyfile(os.environ.get("TEMPLATE_PATH", "TEMPLATE_PXf_SGSLABIP1056.xlsx"), template_copy)

                    write_to_template(req_rows, out_path, source_pdf=up.name)
                    total_amostras += len(req_rows)

                logs.append(f"✅ {up.name}: concluído ({total_amostras} amostras, {len(rows_per_req)} requisições)")
                ok += 1

            except Exception as e:
                logs.append(f"❌ {up.name}:\n{traceback.format_exc()}")
                fail += 1

        # ───────────────
        # Gera ZIP final
        # ───────────────
        zip_path = os.path.join(tmp, f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(outdir):
                for fn in files:
                    if fn.startswith("_tmp_template"):
                        continue  # ignora cópias temporárias
                    full = os.path.join(root, fn)
                    z.write(full, os.path.relpath(full, outdir))

    # ───────────────────────────────────────────────
    # Resultado final
    # ───────────────────────────────────────────────
    st.success(f"🏁 Processamento concluído • {ok} ok, {fail} com erro.")
    with open(zip_path, "rb") as f:
        st.download_button("⬇️ Descarregar resultados (ZIP)", f, file_name=os.path.basename(zip_path))

    with st.expander("🧾 Registo de execução"):
        st.code("\n".join(logs) if logs else "Sem logs a apresentar.")
else:
    st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")
