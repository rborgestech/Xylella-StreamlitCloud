# -*- coding: utf-8 -*-
"""
core_xylella.py — Cloud/Streamlit (OCR Azure direto + Parser Colab + Writer por requisição)
"""

import os
import re
import time
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import load_workbook

# Diretório de saída seguro
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", "output_final"))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# ───────────────────────────────────────────────
# OCR Azure (PDF direto)
# ───────────────────────────────────────────────
def azure_analyze_pdf(pdf_path: str) -> Dict[str, Any]:
    """
    Envia o PDF para o Azure Form Recognizer e devolve o resultado JSON.
    Requer variáveis de ambiente:
      - AZURE_API_KEY
      - AZURE_ENDPOINT
      - AZURE_MODEL_ID (opcional)
    """
    AZURE_API_KEY = os.getenv("AZURE_API_KEY", "")
    AZURE_ENDPOINT = os.getenv("AZURE_ENDPOINT", "")
    MODEL_ID = os.getenv("AZURE_MODEL_ID", "prebuilt-document")

    if not AZURE_API_KEY or not AZURE_ENDPOINT:
        raise RuntimeError("Azure não configurado (AZURE_API_KEY/AZURE_ENDPOINT).")

    url = f"{AZURE_ENDPOINT.rstrip('/')}/formrecognizer/documentModels/{MODEL_ID}:analyze?api-version=2023-07-31"
    headers = {"Ocp-Apim-Subscription-Key": AZURE_API_KEY, "Content-Type": "application/pdf"}

    with open(pdf_path, "rb") as f:
        resp = requests.post(url, data=f.read(), headers=headers, timeout=120)
    if resp.status_code != 202:
        raise RuntimeError(f"Azure analyze falhou: {resp.status_code} {resp.text}")

    op = resp.headers.get("Operation-Location")
    if not op:
        raise RuntimeError("Azure não devolveu Operation-Location.")

    start = time.time()
    while True:
        r = requests.get(op, headers={"Ocp-Apim-Subscription-Key": AZURE_API_KEY}, timeout=60)
        j = r.json()
        st = j.get("status")
        if st == "succeeded":
            return j
        if st == "failed":
            raise RuntimeError(f"OCR Azure falhou: {j}")
        if time.time() - start > 180:
            raise RuntimeError("Timeout a aguardar OCR Azure.")
        time.sleep(1.2)


# ───────────────────────────────────────────────
# Extrai texto completo do JSON Azure
# ───────────────────────────────────────────────
def extract_all_text(result_json: Dict[str, Any]) -> str:
    """Concatena todo o texto linha a linha de todas as páginas."""
    lines = []
    for pg in result_json.get("analyzeResult", {}).get("pages", []):
        for ln in pg.get("lines", []):
            txt = (ln.get("content") or ln.get("text") or "").strip()
            if txt:
                lines.append(txt)
    return "\n".join(lines)


# ───────────────────────────────────────────────
# Placeholder: parser simplificado (substituído pelo Colab parser)
# ───────────────────────────────────────────────
def parse_all_requisitions(result_json, pdf_name, txt_path):
    """
    Placeholder simplificado — em produção usa o parser Colab.
    Aqui devolve apenas 1 requisição com 1 linha dummy.
    """
    print(f"⚠️ Parser simplificado ativo para {os.path.basename(pdf_name)}")
    return [{"rows": [{"referencia": "dummy", "datarececao": "01/01/2025"}], "expected": 1}]


# ───────────────────────────────────────────────
# Escrita do Excel (simulada)
# ───────────────────────────────────────────────
def write_to_template(ocr_rows, out_name, expected_count=None, source_pdf=None):
    """
    Escreve dados reais no template XLSX.
    Cada elemento em ocr_rows é um dicionário com os campos extraídos do PDF.
    """
    try:
        TEMPLATE_PATH = os.getenv("TEMPLATE_PATH", "TEMPLATE_PXf_SGSLABIP1056.xlsx")
        if not os.path.exists(TEMPLATE_PATH):
            raise FileNotFoundError(f"Template não encontrado: {TEMPLATE_PATH}")

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        out_path = OUTPUT_DIR / out_name

        # Abre o template original
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active

        # Linha inicial para escrita (ajusta conforme o teu template)
        start_row = 6

        # Mapeia nomes de colunas (linha 5)
        col_map = {str(c.value).strip().lower(): i for i, c in enumerate(ws[5], start=1) if c.value}

        # Preenche as linhas de amostras
        for r_idx, sample in enumerate(ocr_rows, start=start_row):
            for key, value in sample.items():
                col_name = key.strip().lower()
                if col_name in col_map:
                    col = col_map[col_name]
                    ws.cell(row=r_idx, column=col, value=value)

        # Atualiza contagens e metadados
        processed_count = len(ocr_rows)
        ws["E1"] = f"Nº Amostras: {expected_count or processed_count} / {processed_count}"
        ws["F1"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws["G1"] = source_pdf or ""

        # Guarda o ficheiro
        wb.save(out_path)
        print(f"🟢 Gravado (com template real): {out_path}")
        return str(out_path)

    except Exception as e:
        print(f"❌ Erro ao escrever no template: {e}")
        return None


# ───────────────────────────────────────────────
# API pública usada pela app Streamlit
# ───────────────────────────────────────────────
def process_pdf_sync(pdf_path: str) -> List[Dict[str, Any]]:
    """
    Executa o OCR Azure direto ao PDF e o parser Colab integrado, em paralelo por requisição.
    Devolve: lista de dicionários:
        [
            {"rows": [...], "declared": int},
            {"rows": [...], "declared": int},
        ]
    """
    base = os.path.basename(pdf_path)
    print(f"\n🧪 Início de processamento: {base}")

    # Diretório de output e ficheiro de debug
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    txt_path = OUTPUT_DIR / f"{os.path.splitext(base)[0]}_ocr_debug.txt"

    # 1️⃣ OCR Azure direto
    result_json = azure_analyze_pdf(pdf_path)

    # 2️⃣ Guardar texto OCR global (debug)
    txt_path.write_text(extract_all_text(result_json), encoding="utf-8")
    print(f"📝 Texto OCR bruto guardado em: {txt_path}")

    # 3️⃣ Dividir em requisições
    requisitions = parse_all_requisitions(result_json, pdf_path, str(txt_path))
    total_reqs = len(requisitions)
    print(f"🔍 {total_reqs} requisição(ões) detetada(s).")

    if total_reqs == 0:
        print(f"⚠️ {base}: nenhum bloco de requisição encontrado.")
        return []

    # 4️⃣ Processamento paralelo
    results: List[Dict[str, Any]] = []
    start_time = datetime.now()
    with ThreadPoolExecutor(max_workers=min(4, total_reqs)) as executor:
        futures = {
            executor.submit(_process_single_req, i, req, base, pdf_path): i
            for i, req in enumerate(requisitions, 1)
        }
        for future in as_completed(futures):
            i = futures[future]
            try:
                result_item = future.result()
                if result_item and result_item.get("rows"):
                    results.append(result_item)
            except Exception as e:
                print(f"❌ Erro na requisição {i}: {e}")

    total_amostras = sum(len(r["rows"]) for r in results)
    elapsed = (datetime.now() - start_time).total_seconds()
    print(f"✅ {base}: {len(results)} requisições processadas ({total_amostras} amostras) em {elapsed:.1f}s.")
    return results


def _process_single_req(i: int, req: Dict[str, Any], base: str, pdf_path: str) -> Dict[str, Any]:
    """
    Processa uma única requisição (subfunção auxiliar paralela).
    Retorna {"rows": [...], "declared": expected}
    """
    try:
        rows = req.get("rows", [])
        expected = req.get("expected") or req.get("declared") or 0

        if not rows:
            print(f"⚠️ Requisição {i}: sem amostras — ignorada.")
            return {"rows": [], "declared": expected}

        diff = len(rows) - expected
        if expected and diff != 0:
            print(f"⚠️ Requisição {i}: {len(rows)} processadas vs {expected} declaradas ({diff:+d}).")
        else:
            print(f"✅ Requisição {i}: {len(rows)} amostras processadas (declaradas: {expected}).")

        return {"rows": rows, "declared": expected}

    except Exception as e:
        print(f"❌ Erro interno na requisição {i}: {e}")
        return {"rows": [], "declared": 0}


# ───────────────────────────────────────────────
# Função utilitária: leitura de E1 (nº amostras declaradas/processadas)
# ───────────────────────────────────────────────
def read_e1_counts(xlsx_path: str):
    """
    Lê o valor da célula E1/F1 para obter nº de amostras declaradas/processadas.
    Retorna (expected, processed) — None se não for possível ler.
    """
    try:
        wb = load_workbook(xlsx_path, data_only=True)
        ws = wb.active
        cell = ws["E1"].value
        if not cell or not isinstance(cell, str):
            return (None, None)
        m = re.search(r"(\d+)\s*/\s*(\d+)", cell)
        if m:
            return (int(m.group(1)), int(m.group(2)))
        e_val = ws["E1"].value
        f_val = ws["F1"].value
        if isinstance(e_val, (int, float)) and isinstance(f_val, (int, float)):
            return (int(e_val), int(f_val))
    except Exception as e:
        print(f"⚠️ Erro ao ler E1/F1 em {os.path.basename(xlsx_path)}: {e}")
    return (None, None)

