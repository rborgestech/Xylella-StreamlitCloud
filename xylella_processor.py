# -*- coding: utf-8 -*-
"""
Xylella Processor Wrapper
-------------------------
Faz a ponte entre o core_xylella.py e a app Streamlit.
Garante leitura do ficheiro summary gerado e devolve
os dados prontos para o painel do utilizador.
"""

import os
from pathlib import Path
from datetime import datetime
from core_xylella import process_pdf_sync


# ───────────────────────────────────────────────
#  Função principal
# ───────────────────────────────────────────────
def process_pdf(pdf_path):
    """
    Processa um PDF e devolve:
    [
      {"path": "ficheiro.xlsx", "samples": 30, "declared": 30, "diff": 0},
      ...
    ]
    """
    entries = process_pdf_sync(pdf_path)
    if not entries:
        print("⚠️ Nenhum resultado devolvido pelo core.")
        return []

    # Identifica o ficheiro summary correspondente
    pdf_stem = Path(pdf_path).stem
    summary_path = Path("debug") / f"{pdf_stem}_summary.txt"
    if not summary_path.exists():
        print(f"⚠️ Ficheiro summary não encontrado: {summary_path}")
        return entries

    # Lê e interpreta o resumo
    parsed = []
    total_amostras = 0
    ficheiros_discrep = 0

    with open(summary_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    for line in lines:
        line = line.strip()
        if not line or line.startswith(("🧾", "PDF:", "📊", "🧪", "⚠️ 0 ficheiro", "Total:")):
            continue

        # Exemplo: ✅ ficheiro.xlsx: ficheiro gerado. (30 amostras OK)
        if line.startswith("✅") or line.startswith("⚠️"):
            item = {"path": None, "samples": 0, "declared": 0, "diff": 0}

            # Nome do ficheiro
            try:
                name = line.split(":")[0][2:].strip()
                item["path"] = name
            except Exception:
                continue

            # Amostras e discrepâncias
            if "amostras OK" in line:
                try:
                    num = int(line.split("(")[1].split()[0])
                    item["samples"] = num
                    item["declared"] = num
                except Exception:
                    pass
            elif "vs" in line:
                # ⚠️ ... (12 vs 10 — discrepância +2)
                try:
                    left = int(line.split("(")[1].split("vs")[0].strip())
                    right = int(line.split("vs")[1].split("—")[0].strip())
                    diff = int(line.split("discrepância")[1].split(")")[0].replace("+", "").strip())
                    item["samples"], item["declared"], item["diff"] = left, right, diff
                    ficheiros_discrep += 1
                except Exception:
                    pass

            total_amostras += item.get("samples", 0)
            parsed.append(item)

    print(f"📊 {len(parsed)} ficheiros processados, {total_amostras} amostras, {ficheiros_discrep} discrepância(s).")
    return parsed


# ───────────────────────────────────────────────
#  Função build_zip
# ───────────────────────────────────────────────
import io, zipfile

def build_zip(file_paths):
    """Cria um ZIP em memória com os ficheiros fornecidos."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zipf:
        for fp in file_paths:
            fp = Path(fp)
            if fp.exists():
                zipf.write(fp, arcname=fp.name)
    buf.seek(0)
    return buf.getvalue()


# ───────────────────────────────────────────────
#  Execução direta (teste)
# ───────────────────────────────────────────────
if __name__ == "__main__":
    pdf = "INPUT/20231023_ReqX02_X03_X04_Lab SGS 23 10 2025.pdf"
    result = process_pdf(pdf)
    print("\n🧾 Resultado interpretado:")
    for r in result:
        print(r)
