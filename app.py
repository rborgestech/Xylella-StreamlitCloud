# -*- coding: utf-8 -*-
import streamlit as st
import os, time, base64, zipfile, io
from datetime import datetime
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from core_xylella import read_e1_counts
import xylella_processor as processor

# Configuração base
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 ficheiro Excel por requisição.")

if "processing" not in st.session_state:
    st.session_state.processing = False

# Upload
uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True)

def build_summary(results, total_time):
    lines = []
    total_excels = 0
    total_samples = 0
    discrep_files = 0

    for res in results:
        pdf = res["name"]
        reqs = res["reqs"]
        samples = res["samples"]
        discrep = res["discrepancies"]
        lines.append(f"📄 {pdf}: {reqs} requisição(ões), {samples} amostras" + (f" ⚠️ {discrep} discrepância(s)" if discrep else ""))
        for d in res["details"]:
            base = os.path.basename(d["file"])
            if d["disc"]:
                lines.append(f"   ↳ ⚠️ {base} (processadas: {d['proc']} / declaradas: {d['exp']})")
            else:
                lines.append(f"   ↳ {base}")
        total_excels += len(res["details"])
        total_samples += samples
        if discrep:
            discrep_files += 1

    lines.append("")
    lines.append(f"📊 Total: {total_excels} ficheiro(s) Excel")
    lines.append(f"🧪 Total de amostras: {total_samples}")
    lines.append(f"⏱️ Tempo total: {total_time:.1f} segundos")
    lines.append(f"📅 Executado em: {datetime.now().strftime('%d/%m/%Y às %H:%M:%S')}")
    if discrep_files:
        lines.append(f"⚠️ {discrep_files} ficheiro(s) com discrepâncias")
    else:
        lines.append("✅ Nenhum ficheiro com discrepâncias")
    return "\n".join(lines)


def process_pdf_file(file):
    temp_path = Path("/tmp") / file.name
    with open(temp_path, "wb") as f:
        f.write(file.read())
    created = processor.process_pdf(str(temp_path))
    details = []
    total_samples = 0
    discrepancies = 0

    for path in created:
        exp, proc = read_e1_counts(path)
        total_samples += proc or 0
        is_disc = (exp is not None and proc is not None and exp != proc)
        if is_disc:
            discrepancies += 1
        details.append({"file": path, "exp": exp, "proc": proc, "disc": is_disc})
    return {
        "name": file.name,
        "reqs": len(created),
        "samples": total_samples,
        "discrepancies": discrepancies,
        "details": details
    }


if uploads and st.button("🚀 Processar ficheiros"):
    st.session_state.processing = True
    start = time.time()
    results = []
    placeholders = []

    for file in uploads:
        ph = st.empty()
        ph.info(f"📄 {file.name} — a processar...")
        placeholders.append(ph)

    with ThreadPoolExecutor(max_workers=min(4, len(uploads))) as ex:
        futures = {ex.submit(process_pdf_file, f): i for i, f in enumerate(uploads)}
        for fut in as_completed(futures):
            idx = futures[fut]
            try:
                res = fut.result()
                results.append(res)
                placeholders[idx].success(f"✅ {res['name']} — {res['reqs']} requisição(ões), {res['samples']} amostras.")
            except Exception as e:
                placeholders[idx].error(f"❌ {uploads[idx].name}: {e}")

    total_time = time.time() - start
    summary = build_summary(results, total_time)
    st.markdown("### 📊 Resumo final")
    st.text(summary)

    # ZIP
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for res in results:
            for d in res["details"]:
                if os.path.exists(d["file"]):
                    zf.write(d["file"], arcname=os.path.basename(d["file"]))
    zip_buffer.seek(0)
    st.download_button("📦 Descarregar ZIP", data=zip_buffer, file_name="xylella_output.zip", mime="application/zip")

    st.download_button("🧾 Descarregar summary.txt", data=summary.encode("utf-8"), file_name="summary.txt", mime="text/plain")

    st.success("🏁 Processamento concluído!")
