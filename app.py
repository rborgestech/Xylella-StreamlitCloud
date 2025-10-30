import streamlit as st
import tempfile, os, traceback
from datetime import datetime
from xylella_processor import process_pdf, build_zip

# ───────────────────────────────────────────────
# Configuração
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")

st.markdown("""
<style>
div.stButton > button:first-child {
    background-color: #004080;
    color: white;
    font-weight: bold;
}
div.stButton > button:first-child:disabled {
    background-color: #708090;
    color: white;
}
</style>
""", unsafe_allow_html=True)

st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 Excel por requisição.")

uploaded = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True)
start = st.button("Processar ficheiros de Input", type="primary", disabled=not uploaded)

# ───────────────────────────────────────────────
# Processamento
# ───────────────────────────────────────────────
if start:
    st.session_state["processing"] = True
    with st.spinner("⚙️ A processar... aguarda alguns segundos."):

        tmp = tempfile.mkdtemp()
        outdir = os.path.join(tmp, "output")
        os.makedirs(outdir, exist_ok=True)
        os.environ["OUTPUT_DIR"] = outdir

        logs, ok, fail = [], 0, 0
        created_all = []
        log_lines = []

        for up in uploaded:
            try:
                in_path = os.path.join(tmp, up.name)
                with open(in_path, "wb") as f:
                    f.write(up.read())

                st.markdown(f"### 📄 {up.name}")
                st.write("⏳ Início de processamento...")

                req_files = process_pdf(in_path)
                if not req_files:
                    st.warning(f"⚠️ Nenhum ficheiro gerado para {up.name}")
                    continue

                created_all.extend(req_files)

                # Contar amostras totais (lidas do ficheiro)
                total_amostras = 0
                for fpath in req_files:
                    fname = os.path.basename(fpath)
                    st.success(f"✅ {fname} gravado")
                    # Leitura rápida de contagem
                    try:
                        import openpyxl
                        wb = openpyxl.load_workbook(fpath)
                        ws = wb.active
                        vals = [c.value for c in ws["A"] if c.value]
                        n_amostras = len(vals) - 3 if len(vals) > 3 else 0
                        total_amostras += n_amostras
                    except Exception:
                        pass

                resumo = f"{len(req_files)} requisições, ~{total_amostras} amostras."
                st.info(f"📊 {up.name}: {resumo}")
                log_lines.append(f"{up.name}: {resumo}")
                ok += 1

            except Exception as e:
                err = traceback.format_exc()
                logs.append(f"❌ {up.name}:\n{err}")
                st.error(f"❌ Erro ao processar {up.name}: {e}")
                fail += 1

        # Criar log de execução
        log_path = os.path.join(outdir, "log_processamento.txt")
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(f"Log de execução — {datetime.now():%d/%m/%Y %H:%M}\n\n")
            f.write("\n".join(log_lines or ["Sem ficheiros processados."]))
            if logs:
                f.write("\n\nErros:\n" + "\n".join(logs))
        created_all.append(log_path)

        # ZIP final
        if created_all:
            zip_bytes = build_zip(created_all)
            zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
            zip_path = os.path.join(tmp, zip_name)
            with open(zip_path, "wb") as f:
                f.write(zip_bytes)

            st.success(f"🏁 Concluído: {ok} ok, {fail} com erro.")
            with open(zip_path, "rb") as f:
                st.download_button(
                    "⬇️ Descarregar resultados (ZIP)",
                    f,
                    file_name=os.path.basename(zip_path),
                    mime="application/zip"
                )
        else:
            st.error("❌ Nenhum ficheiro .xlsx criado.")

    # Registo expandido
    with st.expander("🧾 Registo detalhado"):
        if logs:
            st.code("\n".join(logs))
        else:
            st.info("Sem erros reportados.")

    st.session_state["processing"] = False

else:
    st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")
