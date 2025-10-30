# -*- coding: utf-8 -*-
"""
core_xylella.py — versão consolidada (Cloud + Parser Colab)
"""

from __future__ import annotations
import os, re, io, json, time, requests
from pathlib import Path
from datetime import datetime
from typing import Any, Dict, List

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ───────────────────────────────────────────────
# Configuração base
# ───────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", BASE_DIR / "Output"))
OUTPUT_DIR.mkdir(exist_ok=True)
TEMPLATE_PATH = Path(os.environ.get("TEMPLATE_PATH", BASE_DIR / "TEMPLATE_PXf_SGSLABIP1056.xlsx"))

AZURE_API_KEY = os.environ.get("AZURE_API_KEY", "")
AZURE_ENDPOINT = os.environ.get("AZURE_ENDPOINT", "")
MODEL_ID = os.environ.get("AZURE_MODEL_ID", "prebuilt-document")

# ───────────────────────────────────────────────
# Azure OCR (PDF direto)
# ───────────────────────────────────────────────
def azure_analyze_pdf(pdf_path: str) -> Dict[str, Any]:
    if not AZURE_API_KEY or not AZURE_ENDPOINT:
        raise RuntimeError("Azure não configurado (AZURE_API_KEY/AZURE_ENDPOINT).")
    url = f"{AZURE_ENDPOINT.rstrip('/')}/formrecognizer/documentModels/{MODEL_ID}:analyze?api-version=2023-07-31"
    headers = {"Ocp-Apim-Subscription-Key": AZURE_API_KEY, "Content-Type": "application/pdf"}
    with open(pdf_path, "rb") as f:
        resp = requests.post(url, data=f.read(), headers=headers, timeout=90)
    if resp.status_code != 202:
        raise RuntimeError(f"Azure analyze falhou: {resp.status_code} {resp.text}")
    op = resp.headers.get("Operation-Location")
    for _ in range(60):
        time.sleep(1.2)
        r = requests.get(op, headers={"Ocp-Apim-Subscription-Key": AZURE_API_KEY}, timeout=30)
        j = r.json()
        if j.get("status") == "succeeded":
            return j
        if j.get("status") == "failed":
            raise RuntimeError(f"OCR Azure falhou: {j}")
    raise RuntimeError("Timeout a aguardar OCR Azure.")

def extract_all_text(result_json: Dict[str, Any]) -> str:
    lines = []
    for pg in result_json.get("analyzeResult", {}).get("pages", []):
        for ln in pg.get("lines", []):
            txt = (ln.get("content") or ln.get("text") or "").strip()
            if txt:
                lines.append(txt)
    return "\n".join(lines)

# ───────────────────────────────────────────────
# Funções de parsing (secções 8 e 9 do Colab)
# ───────────────────────────────────────────────
NATUREZA_KEYWORDS = [
    "ramos","folhas","ramosefolhas","ramosc/folhas","material",
    "materialherbalho","materialherbário","materialherbalo",
    "natureza","insetos","sementes","solo"
]
REF_SLASH_RE = re.compile(r"\b\d{2,4}/\d{2,4}/[A-Z0-9]{1,8}(?:/[A-Z0-9]{1,8})*\b", re.I)
REF_NUM_RE   = re.compile(r"\b\d{7,8}\b")

def _clean_ref(raw: str) -> str:
    s = raw.strip()
    s = re.sub(r"\s*/\s*", "/", s)
    s = re.sub(r"/{2,}", "/", s)
    s = re.sub(r"[A-Za-z]+", lambda m: m.group(0).upper(), s)
    s = s.replace("LUT","LVT")
    s = re.sub(r"[^A-Z0-9/]+$","",s)
    return s

def _looks_like_natureza(txt: str) -> bool:
    t = re.sub(r"\s+","",txt.lower())
    return any(k in t for k in NATUREZA_KEYWORDS)

def detect_requisicoes(full_text: str):
    pattern = re.compile(r"PROGRAMA\s+NACIONAL\s+DE\s+PROSPE[ÇC][AÃ]O", re.I)
    matches = list(pattern.finditer(full_text))
    if not matches:
        print("🔍 Nenhum cabeçalho encontrado — assumido 1 requisição.")
        return 1,[]
    print(f"🔍 Detetadas {len(matches)} requisições.")
    return len(matches),[m.start() for m in matches]

def split_if_multiple_requisicoes(full_text: str):
    text = re.sub(r"[ \t]+"," ",full_text)
    text = re.sub(r"\n{2,}","\n",text)
    pattern = re.compile(r"PROGRAMA\s+NACIONAL\s+DE\s+PROSPE[ÇC][AÃ]O", re.I)
    marks=[m.start() for m in pattern.finditer(text)]
    if not marks or len(marks)==1: return [text]
    marks.append(len(text))
    blocos=[]
    for i in range(len(marks)-1):
        blocos.append(text[marks[i]:marks[i+1]].strip())
    print(f"📄 Documento dividido em {len(blocos)} requisições.")
    return blocos

def extract_context_from_text(full_text: str) -> Dict[str,Any]:
    ctx={}
    m_zona=re.search(r"Xylella\s+fastidiosa\s*\(([^)]+)\)",full_text,re.I)
    ctx["zona"]=m_zona.group(1).strip() if m_zona else "Zona Isenta"
    m_env=re.search(r"Data\s+(?:do|de)\s+envio.*?([\d/]{8,10})",full_text,re.I)
    ctx["data_envio"]=m_env.group(1) if m_env else datetime.now().strftime("%d/%m/%Y")
    ctx["dgav"]="DGAV"
    ctx["responsavel_colheita"]=""
    ctx["default_colheita"]=ctx["data_envio"]
    ctx["colheita_map"]={}
    return ctx

def parse_xylella_tables(result_json: Dict[str,Any], context: Dict[str,Any], req_id=None)->List[Dict[str,Any]]:
    out=[]
    tables=result_json.get("analyzeResult",{}).get("tables",[]) or []
    if not tables: return out
    for t in tables:
        for c in t.get("cells",[]):
            val=str(c.get("content","")).strip()
            if REF_NUM_RE.match(val) or REF_SLASH_RE.match(val):
                ref=_clean_ref(val)
                out.append({
                    "requisicao_id":req_id,
                    "datarececao":context["data_envio"],
                    "datacolheita":context["default_colheita"],
                    "referencia":ref,
                    "hospedeiro":"",
                    "tipo":"",
                    "zona":context["zona"],
                    "responsavelamostra":context["dgav"],
                    "responsavelcolheita":context["responsavel_colheita"],
                    "observacoes":"",
                    "procedure":"XYLELLA",
                    "datarequerido":context["data_envio"],
                    "Score":""
                })
    print(f"✅ {len(out)} amostras extraídas (req {req_id}).")
    return out

def parse_xylella_from_result(result_json, pdf_name, txt_path=None):
    if txt_path and os.path.exists(txt_path):
        full_text = Path(txt_path).read_text(encoding="utf-8")
    else:
        full_text = extract_all_text(result_json)

    count,_=detect_requisicoes(full_text)
    blocos = split_if_multiple_requisicoes(full_text)
    all_tables=result_json.get("analyzeResult",{}).get("tables",[]) or []
    all_samples=[]
    for i,bloco in enumerate(blocos,start=1):
        context=extract_context_from_text(bloco)
        refs=re.findall(r"\b\d{7,8}\b|\b\d{2,4}/\d{2,4}/[A-Z0-9\-]+\b",bloco,re.I)
        tables_filtradas=[]
        for t in all_tables:
            joined=" ".join(c.get("content","") for c in t.get("cells",[]))
            if any(r in joined for r in refs):
                tables_filtradas.append(t)
        if not tables_filtradas:
            print(f"⚠️ Sem tabelas para requisição {i}.")
            continue
        local={"analyzeResult":{"tables":tables_filtradas}}
        amostras=parse_xylella_tables(local,context,req_id=i)
        if amostras: all_samples.append(amostras)
    print(f"🏁 Concluído: {len(all_samples)} requisições, {sum(len(x) for x in all_samples)} amostras.")
    return all_samples

# ───────────────────────────────────────────────
# Escrita no template SGS (limpeza completa)
# ─────────────────────────────────
