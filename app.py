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

if start:
    tmp = tempfile.mkdtemp()
    outdir = os.path.join(tmp, "output")
    os.makedirs(outdir, exist_ok=True)

    logs, ok, fail = [], 0, 0
    os.environ["OUTPUT_DIR"] = outdir  # Garante saída no diretório temporário

    progress = st.progress(0, text="Inicializando...")

    total_files = len(uploaded)
    processed_files = 0

    for up in uploaded:
        try:
            in_path = os.path.join(tmp, up.name)
            with open(in_path, "wb") as f:
                f.write(up.read())

            progress.progress(processed_files / total_files, text=f"📄 A processar {up.name}...")
            st.write(f"🧪 **{up.name}** — início de processamento...")

            # 📘 Processa PDF → várias requisições
            rows_per_req = process_pdf(in_path)
            base = os.path.splitext(up.name)[0]
            total_amostras = 0
            discrepancias = []

            # 🔄 Cria 1 ficheiro por requisição
            for i, req_rows in enumerate(rows_per_req, start=1):
                if not req_rows:
                    continue

                out_name = f"{base}_req{i}.xlsx"
                out_path = os.path.join(outdir, out_name)

                # Copiar template limpo
                from shutil import copyfile
                template_src = os.environ.get("TEMPLATE_PATH", "TEMPLATE_PXf_SGSLABIP1056.xlsx")
                if os.path.exists(template_src):
                    template_copy = os.path.join(outdir, f"_tmp_template_req{i}.xlsx")
                    copyfile(template_src, template_copy)

                # Calcular amostras declaradas/processadas
                expected = None
                if len(req_rows) > 0 and "declared_samples" in req_rows[0]:
                    expected = req_rows[0]["declared_samples"]

                write_to_template(req_rows, out_path, expected_count=expected, source_pdf=up.name)

                if expected and expected != len(req_rows):
                    discrepancias.append((i, expected, len(req_rows)))

                total_amostras += len(req_rows)
                st.write(f"✅ Requisição {i}: {len(req_rows)} amostras gravadas → {out_name}")

            # Resumo por PDF
            msg = f"✅ {up.name}: {len(rows_per_req)} requisições, {total_amostras} amostras"
            if discrepancias:
                msg += f" ⚠️ ({len(discrepancias)} discrepâncias detectadas)"
                st.warning(msg)
            else:
                st.success(msg)

            logs.append(msg)
            ok += 1

        except Exception:
            logs.append(f"❌ {up.name}:\n{traceback.format_exc()}")
            st.error(f"❌ Erro ao processar {up.name}.")
            fail += 1

        processed_files += 1
        progress.progress(processed_files / total_files, text=f"Concluído {processed_files}/{total_files}")

    # Finalização
    progress.progress(1.0, text="🏁 Todos os ficheiros processados.")

    # Gera ZIP
    zip_path = os.path.join(tmp, f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(outdir):
            for fn in files:
                if fn.startswith("_tmp_template"):
                    continue
                z.write(os.path.join(root, fn), os.path.relpath(os.path.join(root, fn), outdir))

    st.success(f"🏁 Processamento concluído • {ok} ok, {fail} com erro.")
    with open(zip_path, "rb") as f:
        st.download_button("⬇️ Descarregar resultados (ZIP)", f, file_name=os.path.basename(zip_path))

    # Log detalhado
    with st.expander("🧾 Registo de execução"):
        st.code("\n".join(logs) if logs else "Sem logs a apresentar.")
else:
    st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")
