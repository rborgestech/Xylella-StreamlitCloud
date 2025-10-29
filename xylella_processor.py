# xylella_processor.py
# -*- coding: utf-8 -*-
"""
Adaptador leve para a app Streamlit.

Responsabilidades:
- Importa o motor real a partir de `core_xylella.py`;
- Garante que o TEMPLATE Excel é encontrado por caminho relativo;
- Expõe unicamente as funções esperadas pela UI:
    • process_pdf(pdf_path) -> rows
    • write_to_template(ocr_rows, out_base_path, expected_count=None, source_pdf=None)
"""

from __future__ import annotations

from pathlib import Path
import os
import importlib
from typing import Any, Optional

# ───────────────────────────────────────────────────────────────
# Localização robusta do TEMPLATE (ao lado deste ficheiro)
# ───────────────────────────────────────────────────────────────
# Caminho robusto do TEMPLATE (funciona local e no Streamlit Cloud)
# 🔍 Caminho robusto para o TEMPLATE (funciona em local e no Streamlit Cloud)
# Caminho robusto do TEMPLATE (funciona local, GitHub Codespaces e Streamlit Cloud)
# Nome do ficheiro de template
TEMPLATE_FILENAME = "TEMPLATE_PXF_SGSLABIP1056.xlsx"
TEMPLATE_PATH = Path(__file__).with_name(TEMPLATE_FILENAME)

# 🩹 Caso o ficheiro não exista no ambiente Streamlit Cloud, tenta obtê-lo via GitHub raw
if not TEMPLATE_PATH.exists():
    print("⚠️ TEMPLATE não encontrado localmente — a tentar descarregar do GitHub...")
    url = "https://github.com/rborgestech/Xylella-StreamlitCloud/raw/main/TEMPLATE_PXF_SGSLABIP1056.xlsx"
    try:
        r = requests.get(url)
        r.raise_for_status()
        TEMPLATE_PATH.write_bytes(r.content)
        print(f"✅ TEMPLATE descarregado para {TEMPLATE_PATH}")
    except Exception as e:
        raise FileNotFoundError(f"❌ Não foi possível obter o TEMPLATE: {e}")

# Exporta o caminho final para o ambiente
os.environ.setdefault("TEMPLATE_PATH", str(TEMPLATE_PATH))
print(f"📂 TEMPLATE_PATH final: {TEMPLATE_PATH}")


# ───────────────────────────────────────────────────────────────
# Import do motor (core)
# ───────────────────────────────────────────────────────────────
_CORE_MODULE_NAME = "core_xylella"

try:
    core = importlib.import_module(_CORE_MODULE_NAME)
except Exception as e:  # erro de import claro para a UI
    raise ImportError(
        f"Não foi possível importar '{_CORE_MODULE_NAME}'. "
        f"Garante que o ficheiro 'core_xylella.py' existe na raiz do projeto "
        f"e que compila sem erros. Detalhe: {e!r}"
    )

# Verificações suaves de interface
if not hasattr(core, "process_pdf"):
    raise AttributeError(
        "O módulo 'core_xylella' não expõe a função 'process_pdf(pdf_path)'."
    )
if not hasattr(core, "write_to_template"):
    raise AttributeError(
        "O módulo 'core_xylella' não expõe a função "
        "'write_to_template(ocr_rows, out_base_path, expected_count=None, source_pdf=None)'."
    )

_core_process_pdf = getattr(core, "process_pdf")
_core_write_to_template = getattr(core, "write_to_template")

# ───────────────────────────────────────────────────────────────
# API pública para a app Streamlit
# ───────────────────────────────────────────────────────────────
def process_pdf(pdf_path: str) -> Any:
    """
    Recebe o caminho para 1 PDF e devolve 'rows' (estrutura compreendida pelo core).
    """
    pdf_path = str(pdf_path)
    if not Path(pdf_path).exists():
        raise FileNotFoundError(f"PDF não encontrado: {pdf_path}")
    return _core_process_pdf(pdf_path)


def write_to_template(
    ocr_rows: Any,
    out_base_path: str,
    expected_count: Optional[int] = None,
    source_pdf: Optional[str] = None,
) -> Any:
    """
    Grava 1+ ficheiros Excel com base no TEMPLATE.
    - out_base_path: caminho base sem extensão. O core deve gravar:
         <out_base_path>_req1.xlsx, _req2.xlsx, ...
    - expected_count: nº de requisições esperado (opcional). Pode ser usado para validação.
    - source_pdf: nome do PDF de origem (opcional, útil para metadata/log).
    """
    base = Path(out_base_path)
    base.parent.mkdir(parents=True, exist_ok=True)

    # Garante que o TEMPLATE existe (melhor falhar cedo)
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"TEMPLATE não encontrado em {TEMPLATE_PATH}. "
            f"Confirma que '{TEMPLATE_FILENAME}' está ao lado do 'xylella_processor.py'."
        )

    # Se o core aceitar TEMPLATE via argumento/kwargs, podes passar aqui:
    # return _core_write_to_template(ocr_rows, base.as_posix(), expected_count, source_pdf, template_path=str(TEMPLATE_PATH))

    # Caso contrário, a maioria dos cores lê TEMPLATE_PATH do ambiente (já definido acima)
    return _core_write_to_template(
        ocr_rows,
        base.as_posix(),
        expected_count=expected_count,
        source_pdf=source_pdf,
    )


# ───────────────────────────────────────────────────────────────
# Execução direta para teste rápido (opcional)
#   Ex.: python xylella_processor.py /caminho/fich.pdf
# ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys, zipfile, datetime, traceback

    if len(sys.argv) < 2:
        print("Uso: python xylella_processor.py <ficheiro.pdf>")
        sys.exit(1)

    pdf = sys.argv[1]
    try:
        rows = process_pdf(pdf)
        out_base = Path(pdf).with_suffix("")  # sem .pdf
        write_to_template(rows, str(out_base), expected_count=None, source_pdf=Path(pdf).name)
        print("✅ Processado com sucesso.")
    except Exception:
        print("❌ Erro ao processar:\n" + traceback.format_exc())
        sys.exit(2)
