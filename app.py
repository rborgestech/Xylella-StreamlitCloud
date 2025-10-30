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
                # Guarda o PDF carregado
                in_path = os.path.join(tmp, up.name)
                with open(in_path, "wb") as f:
                    f.write(up.read())

                base = os.path.splitext(up.name)[0]

                # 🔧 Garante que o core_xylella usa a pasta temporária correta
                os.environ["OUTPUT_DIR"] = outdir

                # 🔍 Processa o PDF (OCR + parser)
                rows = process_pdf(in_path)

                # ✍️ Gera um ficheiro Excel por requisição
                for i, req_rows in enumerate(rows, start=1):
                    req_name = f"{base}_req{i}"
                    write_to_template(req_rows, os.path.join(outdir, req_name), source_pdf=up.name)

                # 📊 Estatísticas
                total_amostras = sum(len(r) for r in rows)
                logs.append(f"✅ {up.name}: concluído ({total_amostras} amostras, {len(rows)} requisições)")
                ok += 1

            except Exception:
                logs.append(f"❌ {up.name}:\n{traceback.format_exc()}")
                fail += 1

        # ───────────────────────────────────────────────
        # Geração do ZIP
        # ───────────────────────────────────────────────
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

    # 🧾 Logs de execução
    with st.expander("🧾 Registo de execução"):
        st.code("\n".join(logs) if logs else "Sem logs a apresentar.")

    # 📄 Opcional — botão para descarregar texto OCR
    debug_files = [fn for fn in os.listdir(outdir) if fn.endswith("_ocr_debug.txt")]
    if debug_files:
        with open(os.path.join(outdir, debug_files[0]), "rb") as f:
            st.download_button("📄 Ver texto OCR extraído", f, file_name=debug_files[0])

else:
    st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")
