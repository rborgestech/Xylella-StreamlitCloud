# xylella_processor.py
# -*- coding: utf-8 -*-
"""
Adaptador leve entre a app Streamlit e o core_xylella.py

Responsabilidades:
- Garante que o TEMPLATE Excel é encontrado;
- Importa as funções reais do core;
- Expõe uma API estável esperada pela UI:
    • process_pdf(pdf_path) -> rows
    • write_to_template(rows, pdf_name)
"""

from __future__ import annotations
from pathlib import Path
import os
import sys
import importlib
import openpyxl

# ───────────────────────────────────────────────
# Localização robusta do TEMPLATE
# ───────────────────────────────────────────────
TEMPLATE_FILENAME = "TEMPLATE_PXF_SGSLABIP1056.xlsx"
TEMPLATE_PATH = Path(__file__).with_name(TEMPLATE_FILENAME)

if not TEMPLATE_PATH.exists():
    print("⚠️ TEMPLATE não encontrado — a criar dummy temporário.")
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Avaliação pré registo"
        ws.append(["Data Receção", "Data Colheita", "Código", "Espécie", "Natureza", "Zona", "Responsável"])
        wb.save(TEMPLATE_PATH)
        print(f"✅ TEMPLATE dummy criado em {TEMPLATE_PATH}")
    except Exception as e:
        raise FileNotFoundError(f"❌ Falha ao criar TEMPLATE: {e}")

os.environ.setdefault("TEMPLATE_PATH", str(TEMPLATE_PATH))
print(f"📂 TEMPLATE_PATH final: {TEMPLATE_PATH}")

# ───────────────────────────────────────────────
# Importa o core real
# ───────────────────────────────────────────────
sys.path.append(str(Path(__file__).resolve().parent))
_CORE_MODULE_NAME = "core_xylella"

try:
    core = importlib.import_module(_CORE_MODULE_NAME)
except Exception as e:
    raise ImportError(
        f"❌ Não foi possível importar '{_CORE_MODULE_NAME}'. "
        f"Verifica se 'core_xylella.py' existe e compila. Detalhe: {e!r}"
    )

# Mapeia funções reais (com nomes adaptados)
if not hasattr(core, "process_pdf_sync"):
    raise AttributeError("O core_xylella.py não contém a função 'process_pdf_sync(pdf_path)'.")

if not hasattr(core, "write_to_template"):
    raise AttributeError("O core_xylella.py não contém a função 'write_to_template(rows, pdf_name)'.")

_core_process_pdf = getattr(core, "process_pdf_sync")
_core_write_to_template = getattr(core, "write_to_template")

# ───────────────────────────────────────────────
# API pública usada pela app Streamlit
# ───────────────────────────────────────────────
def process_pdf(pdf_path: str):
    """Recebe o caminho para um PDF e devolve as linhas (rows)."""
    return _core_process_pdf(pdf_path)


def write_to_template(rows, out_base_path, expected_count=None, source_pdf=None):
    """
    Redireciona para a função real no core.
    Ignora parâmetros não usados (expected_count, source_pdf).
    """
    return _core_write_to_template(rows, out_base_path)


# ───────────────────────────────────────────────
# Execução direta (teste local)
# ───────────────────────────────────────────────
if __name__ == "__main__":
    import sys, traceback

    if len(sys.argv) < 2:
        print("Uso: python xylella_processor.py <ficheiro.pdf>")
        sys.exit(1)

    pdf = sys.argv[1]
    try:
        rows = process_pdf(pdf)
        write_to_template(rows, Path(pdf).stem)
        print("✅ Processado com sucesso.")
    except Exception:
        print("❌ Erro ao processar:\n" + traceback.format_exc())
        sys.exit(2)
