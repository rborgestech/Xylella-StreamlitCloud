import streamlit as st
import tempfile, os, zipfile, traceback
from datetime import datetime
from pathlib import Path
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
if start:
    with st.spinner("⚙️ A processar os ficheiros... Isto pode demorar alguns segundos."):
        tmp = tempfile.mkdtemp()
        outdir = os.path.join(tmp, "output")
        os.makedirs(outdir, exist_ok=True)

        logs, ok, fail = [], 0, 0

        for up in uploaded:
            try:
                # grava o PDF carregado
                in_path = os.path.join(tmp, up.name)
                with open(in_path, "wb") as f:
                    f.write(up.read())

                # preparar TEMPLATE e OUTPUT_DIR
                TEMPLATE_FILENAME = "TEMPLATE_PXf_SGSLABIP1056.xlsx"
                template_local = Path(__file__).with_name(TEMPLATE_FILENAME)
                if not template_local.exists():
                    st.error(f"❌ TEMPLATE não encontrado: {template_local}")
                    st.stop()

                template_tmp = os.path.join(tmp, TEMPLATE_FILENAME)
                if not os.path.exists(template_tmp):
                    with open(template_local, "rb") as src, open(template_tmp, "wb") as dst:
                        dst.write(src.read())

                os.environ["TEMPLATE_PATH"] = template_tmp
                os.environ["OUTPUT_DIR"] = outdir

                # processa PDF (OCR + parser + split requisições)
                rows = process_pdf(in_path)

                # nome base do ficheiro
                base = os.path.splitext(up.name)[0]

               # escreve 1 ficheiro por requisição
                for i, req_rows in enumerate(rows, start=1):
                    req_name = f"{base}_req{i}"
                    write_to_template(req_rows, os.path.join(outdir, req_name), source_pdf=up.name)

                # total de amostras e requisições
                total_amostras = sum(len(r) for r in rows)
                logs.append(f"✅ {up.name}: concluído ({total_amostras} amostras, {len(rows)} requisições)")
                ok += 1

            except Exception:
                logs.append(f"❌ {up.name}:\n{traceback.format_exc()}")
                fail += 1

        # gera ZIP de resultados
        zip_path = os.path.join(tmp, f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(outdir):
                for fn in files:
                    full = os.path.join(root, fn)
                    z.write(full, os.path.relpath(full, outdir))

    # ───────────────────────────────────────────────
    # Resultado final
    # ───────────────────────────────────────────────
    st.success(f"🏁 Processamento concluído • {ok} ok, {fail} com erro.")
    with open(zip_path, "rb") as f:
        st.download_button("⬇️ Descarregar resultados (ZIP)", f, file_name=os.path.basename(zip_path))

    # Mostrar OCR debug se existir
    debug_files = [fn for fn in os.listdir(outdir) if fn.endswith("_ocr_debug.txt")]
    if debug_files:
        with open(os.path.join(outdir, debug_files[0]), "rb") as f:
            st.download_button("📄 Ver texto OCR extraído", f, file_name=debug_files[0])

    # Logs de execução
    with st.expander("🧾 Registo de execução"):
        st.code("\n".join(logs) if logs else "Sem logs a apresentar.")

else:
    st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")
