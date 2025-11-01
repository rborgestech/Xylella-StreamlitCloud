# -*- coding: utf-8 -*-
from pathlib import Path
from datetime import datetime
from core_xylella import process_pdf_sync

pdf_path = "INPUT/20231023_ReqX02_X03_X04_Lab SGS 23 10 2025.pdf"  # substitui pelo teu caminho real
debug_dir = Path(__file__).resolve().parent / "debug"
debug_dir.mkdir(parents=True, exist_ok=True)

print(f"\n🧪 A testar o parser real diretamente em: {pdf_path}")
result = process_pdf_sync(pdf_path)
print("\n🧾 Resultado devolvido:")
print(result)
print("\n──────────────────────────────")

# Normalização simples só para inspecionar
entries = []
if not result:
    print("⚠️ Nenhum resultado devolvido pelo parser.")
else:
    if isinstance(result, list):
        for item in result:
            if isinstance(item, dict):
                entries.append(item)
            elif isinstance(item, (tuple, list)):
                entries.append({"path": item[0], "samples": item[1:]})
            else:
                entries.append({"path": str(item)})
    elif isinstance(result, tuple):
        entries.append({"tuple": result})
    else:
        entries.append({"raw": result})

# Grava um summary simples
summary_path = debug_dir / f"{Path(pdf_path).stem}_summary.txt"
with open(summary_path, "w", encoding="utf-8") as f:
    f.write(f"🧾 TESTE DIRECTO — {datetime.now():%Y-%m-%d %H:%M:%S}\n")
    f.write(f"PDF: {Path(pdf_path).name}\n\n")
    for e in entries:
        f.write(f"{e}\n")

print(f"✅ Ficheiro summary criado em {summary_path}")
