# -*- coding: utf-8 -*-
"""
app.py — Interface Streamlit do Processador Xylella

Executa o OCR, parsing e geração de Excel a partir de PDFs SGS/DGAV.
Usa core_xylella.py e xylella_processor.py como backend.
"""

import streamlit as st
import time
import io
import os
from datetime import datetime
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

import xylella_processor as processor
from core_xylella import read_e1_counts  # caso precises de validar o E1

# ───────────────────────────────────────────────
# Configuração inicial
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", layout="centered")
st.title("🧪 Processador Xylella")
st.markdown("Carrega um ou mais ficheiros PDF para processar:")

uploaded_files = st.file_uploader("Seleciona ficheiros PDF", type=["pdf"], accept_multiple_files=True)

# ───────────────────────────────────────────────
# Funções auxiliares
# ───────────────────────────────────────────────
def process_single_pdf(uploaded_file):
    """Guarda o PDF temporariamente e processa-o via xylella_processor."""
    temp_dir = Path("/tmp/xylella_input")
    temp_dir.mkdir(exist_ok=True, parents=True)
    temp_path = temp_dir / uploaded_file.name

    with open(temp_path, "wb") as f:
        f.write(uploaded_file.read())

    # Mostra box azul enquanto processa
    with st.container():
        st.info(f"📄 A processar: **{uploaded_file.name}** ...")
        start_time = time.time()

        try:
            created_files = processor.process_pdf(str(temp_path))
        except Exception as e:
            st.error(f"❌ Erro ao processar {uploaded_file.name}: {e}")
            return [], 0, 0, 0

        elapsed = time.time() - start_time

        # Análise dos ficheiros gerados
        total_samples = 0
        discrepancies = []
        details = []

        for p in created_files:
            exp, proc = read_e1_counts(p)
            total_samples += proc or 0
            if exp is not None and proc is not None and exp != proc:
                discrepancies.append((p, exp, proc))
            details.append((p, exp, proc))

        return created_files, total_samples, len(discrepancies), elapsed, details


def build_summary(results, total_elapsed):
    """Constrói texto do summary.txt no formato solicitado."""
    lines = []
    total_excels = 0
    total_samples = 0
    total_discrep_files = 0

    for r in results:
        pdf = r["name"]
        nreq = len(r["files"])
        nsamples = r["total_samples"]
        ndisc = r["discrepancies"]
        details = r["details"]

        line = f"📄 {pdf}: {nreq} requisição(ões), {nsamples} amostras"
        if ndisc > 0:
            line += f" ⚠️ {ndisc} discrepância(s)"
        lines.append(line)

        for (p, exp, proc) in details:
            base = os.path.basename(p)
            if exp is not None and proc is not None and exp != proc:
                lines.append(f"   ↳ ⚠️ {base} (processadas: {proc} / declaradas: {exp})")
            else:
                lines.append(f"   ↳ {base}")

        total_excels += len(details)
        total_samples += nsamples
        total_discrep_files += ndisc

    lines.append(f"📊 Total: {total_excels} ficheiro(s) Excel")
    lines.append(f"🧪 Total de amostras: {total_samples}")
    lines.append(f"⏱️ Tempo total: {total_elapsed:.1f} segundos")
    lines.append(f"📅 Executado em: {datetime.now().strftime('%d/%m/%Y às %H:%M:%S')}")
    if total_discrep_files > 0:
        lines.append(f"⚠️ {total_discrep_files} ficheiro(s) com discrepâncias")
    else:
        lines.append("✅ Nenhum ficheiro com discrepâncias")

    return "\n".join(lines)


# ───────────────────────────────────────────────
# Botão de execução
# ───────────────────────────────────────────────
if uploaded_files:
    if st.button("🚀 Processar ficheiros de Input"):
        st.write("⏳ A processar ficheiros... aguarde até o processo terminar.")
        start_global = time.time()
        results = []

        for uploaded_file in uploaded_files:
            files, total_samples, ndisc, elapsed, details = process_single_pdf(uploaded_file)
            results.append({
                "name": uploaded_file.name,
                "files": files,
                "total_samples": total_samples,
                "discrepancies": ndisc,
                "details": details
            })
            time.sleep(0.1)  # pausa curta para evitar sobreposição visual

        total_elapsed = time.time() - start_global
        summary_txt = build_summary(results, total_elapsed)

        # Exibe resumo
        st.markdown("### 📊 Resumo final")
        st.text(summary_txt)

        # Guarda o summary.txt
        out_summary = Path("/tmp/summary.txt")
        out_summary.write_text(summary_txt, encoding="utf-8")

        # ZIP com os ficheiros Excel
        all_excels = [f for r in results for f in r["files"] if os.path.exists(f)]
        if all_excels:
            import zipfile, io
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for p in all_excels:
                    zf.write(p, arcname=os.path.basename(p))
            st.download_button(
                label="📦 Descarregar ZIP",
                data=zip_buffer.getvalue(),
                file_name="xylella_excels.zip",
                mime="application/zip"
            )

        st.download_button(
            label="🧾 Descarregar summary.txt",
            data=summary_txt.encode("utf-8"),
            file_name="summary.txt",
            mime="text/plain"
        )

        st.success("🏁 Processamento concluído!")

        # botão novo processamento
        if st.button("🔄 Novo processamento"):
            st.session_state.clear()
            st.experimental_rerun()
