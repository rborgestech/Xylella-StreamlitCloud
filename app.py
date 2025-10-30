# -*- coding: utf-8 -*-
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
# Processamento principal
# ───────────────────────────────────────────────
if start:
    with st.spinner("⚙️ A processar os ficheiros... Isto pode demorar alguns segundos."):
        # cria diretório temporário
        tmp = tempfile.mkdtemp()
        outdir = os.path.join(tmp, "output")
        os.makedirs(outdir, exist_ok=True)

        # garante que o TEMPLATE existe e copia para o tmp
        TEMPLATE_FILENAME = "TEMPLATE_PXf_SGSLABIP1056.xlsx"
        template_local = Path(__file__).with_name(TEMPLATE_FILENAME)

        if not template_local.exists():
            st.error(f"❌ TEMPLATE não encontrado: {template_local}")
            st.stop()

        template_tmp = os.path.join(tmp, TEMPLATE_FILENAME)
        if not os.path.exists(template_tmp):
            with open(template_local, "rb") as src, open(template_tmp, "wb") as dst:
                dst.write(src.read())

        # define as variáveis de ambiente para o core
        os.environ["TEMPLATE_PATH"] = template_tmp
        os.environ["OUTPUT_DIR"] = outdir

        logs, ok, fail = [], 0, 0

        # ───────────────────────────────────────────────
        # Loop de processamento de PDFs
        # ───────────────────────────────────────────────
        for up in uploaded:
            try:
                # guarda PDF no tmp
                in_path = os.path.join(tmp, up.name)
                with open(in_path, "wb") as f:
                    f.write(up.read())

                # processa PDF com OCR + parser (core_xylella)
                rows = process_pdf(in_path)

                # nome base do ficheiro (sem extensão)
                base = os.path.splitext(up.name)[0]

                # escreve os resultados no template (1 ficheiro por requisição)
                write_to_template(rows, os.path.join(outdir, base), source_pdf=up.name)

                # contagem e logs
                total_amostras = sum(len(r) for r in rows)
                logs.append(f"✅ {up.name}: concluído ({total_amostras} amostras, {len(rows)} requisições)")
                ok += 1
            except Exception:
                logs.append(f"❌ {up.name}:\n{traceback.format_exc()}")
                fail += 1

        # ───────────────────────────────────────────────
        # Criação do ZIP final
        # ───────────────────────────────────────────────
        zip_path = os.path.join(tmp_
