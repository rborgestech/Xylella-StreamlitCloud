# -*- coding: utf-8 -*-
"""
core_xylella.py — Cloud/Streamlit (OCR Azure direto + Parser Colab + Writer por requisição)

API exposta e usada pela UI (xylella_processor.py):
    • process_pdf_sync(pdf_path) -> List[List[Dict]]]   # devolve lista de requisições; cada requisição = lista de amostras (dict)
    • write_to_template(rows, out_name, expected_count=None, source_pdf=None) -> str  # escreve 1 XLSX com base no template

Requer:
  - AZURE_API_KEY, AZURE_ENDPOINT (env)
  - TEMPLATE_PATH (env) ou ficheiro 'TEMPLATE_PXf_SGSLABIP1056.xlsx' ao lado do core
  - OUTPUT_DIR (env) — diretório onde guardar .xlsx e _ocr_debug.txt
"""

# -*- coding: utf-8 -*-
import os
import re
import time
import tempfile
import importlib
import requests
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional
import zipfile
import shutil

# 🟢 Biblioteca Excel
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

from datetime import datetime, timedelta
from workalendar.europe import Portugal
from openpyxl.formula.translate import Translator

# ───────────────────────────────────────────────
# Diretório de saída seguro
# ───────────────────────────────────────────────
try:
    OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", tempfile.gettempdir()))
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
except Exception as e:
    OUTPUT_DIR = Path(tempfile.gettempdir())
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"[WARN] Não foi possível criar diretório de output definido: {e}. Usando {OUTPUT_DIR}")

# ───────────────────────────────────────────────
# Diretório base e template
# ───────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = Path(os.environ.get("TEMPLATE_PATH", BASE_DIR / "TEMPLATE_PXf_SGSLABIP1056.xlsx"))
if not TEMPLATE_PATH.exists():
    print(f"ℹ️ Aviso: TEMPLATE não encontrado em {TEMPLATE_PATH}. Será verificado no momento da escrita.")

# ───────────────────────────────────────────────
# Carregamento do módulo principal (seguro)
# ───────────────────────────────────────────────
try:
    import core_xylella_main as core
except ModuleNotFoundError:
    try:
        import core_xylella_base as core
    except ModuleNotFoundError:
        core = None
        print("⚠️ Nenhum módulo core_xylella_* encontrado — funcionalidade limitada.")

# ───────────────────────────────────────────────
# Azure OCR — credenciais
# ───────────────────────────────────────────────
AZURE_API_KEY = os.environ.get("AZURE_API_KEY", "")
AZURE_ENDPOINT = os.environ.get("AZURE_ENDPOINT", "")
MODEL_ID = os.environ.get("AZURE_MODEL_ID", "prebuilt-document")

# ───────────────────────────────────────────────
# Estilos Excel
# ───────────────────────────────────────────────
GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
RED   = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
GRAY  = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
BOLD  = Font(bold=True, color="000000")
ITALIC= Font(italic=True, color="555555")

# ───────────────────────────────────────────────
# Utilitários genéricos
# ───────────────────────────────────────────────

from datetime import datetime, timedelta
def integrate_logic_and_generate_name(source_pdf: str) -> tuple[str, str]:
    """
    Função utilitária que:
    - Extrai a data (YYYYMMDD) do início do nome do ficheiro PDF
    - Adiciona 1 dia útil (PT)
    - Substitui esse prefixo de data pelo novo (YYYYMMDD) no nome final do ficheiro Excel

    Retorna:
        (data_ddmm, novo_nome_base)
    """
    base_name = os.path.splitext(os.path.basename(source_pdf))[0]

    # 1. Extrair prefixo de data (YYYYMMDD)
    m = re.match(r"(\d{8})_", base_name)
    if not m:
        return "0000", base_name

    try:
        data_envio = datetime.strptime(m.group(1), "%Y%m%d").date()
        cal = Portugal()
        data_util = cal.add_working_days(data_envio, 1)
        data_ddmm = data_util.strftime("%d%m")
        data_util_str = data_util.strftime("%Y%m%d")
    except Exception:
        return "0000", base_name

    # 2. Substituir prefixo no nome do ficheiro
    novo_nome = re.sub(r"^\d{8}_", f"{data_util_str}_", base_name)
    return data_ddmm, novo_nome
# Feriados fixos em Portugal
FERIADOS_FIXOS = [
    "01-01", "25-04", "01-05", "10-06", "15-08",
    "05-10", "01-11", "01-12", "08-12", "25-12"
]



def _is_valid_date(v) -> bool:
    try:
        datetime.strptime(str(v).strip(), "%d/%m/%Y")
        return True
    except Exception:
        return False

def _to_dt(v):
    try:
        return datetime.strptime(str(v).strip(), "%d/%m/%Y")
    except Exception:
        return v

def clean_value(s: str) -> str:
    if s is None:
        return ""
    if isinstance(s, (int, float)):
        return str(s)
    s = re.sub(r"[\u200b\t\r\f\v]+", " ", str(s))
    s = (s.strip()
           .replace("N/A", "")
           .replace("%", "")
           .replace("\n", " ")
           .replace("  ", " "))
    return s.strip()

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
# OCR Azure (PDF direto)
# ───────────────────────────────────────────────
def azure_analyze_pdf(pdf_path: str) -> Dict[str, Any]:
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

    # Polling
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
# Parser — blocos do Colab (integrado)
# ───────────────────────────────────────────────
NATUREZA_KEYWORDS = [
    "ramos","folhas","ramosefolhas","ramosc/folhas","material","materialherbalho",
    "materialherbário","materialherbalo","natureza","insetos","sementes","solo"
]
TIPO_RE = re.compile(r"\b(Simples|Composta|Composto|Individual)\b", re.I)

def _looks_like_natureza(txt: str) -> bool:
    t = re.sub(r"\s+", "", (txt or "").lower())
    return any(k in t for k in NATUREZA_KEYWORDS)

def _clean_ref(raw: str) -> str:
    s = (raw or "").strip()
    s = re.sub(r"\s*/\s*", "/", s)
    s = re.sub(r"/{2,}", "/", s)
    s = re.sub(r"[A-Za-z]+", lambda m: m.group(0).upper(), s)
    s = s.replace("LUT", "LVT")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^A-Z0-9/]+$", "", s)
    return s

def detect_requisicoes(full_text: str):
    """Conta quantas requisições DGAV→SGS existem no texto OCR de um PDF."""
    pattern = re.compile(
        r"PROGRAMA\s+NACIONAL\s+DE\s+PROSPE[ÇC][AÃ]O\s+DE\s+PRAGAS\s+DE\s+QUARENTENA",
        re.IGNORECASE,
    )
    matches = list(pattern.finditer(full_text))
    count = len(matches)
    positions = [m.start() for m in matches]
    if count == 0:
        print("🔍 Nenhum cabeçalho encontrado — assumido 1 requisição.")
        count = 1
    else:
        print(f"🔍 Detetadas {count} requisições no ficheiro (posições: {positions})")
    return count, positions

def split_if_multiple_requisicoes(full_text: str) -> List[str]:
    """Divide o texto OCR em blocos distintos, um por requisição DGAV→SGS."""
    # Limpeza leve (como no Colab) para juntar tokens partidos por \n
    text = full_text.replace("\r", "")
    text = re.sub(r"(\w)[\n\s]+(\w)", r"\1 \2", text)              # junta palavras quebradas
    text = re.sub(r"(\d+)\s*/\s*([Xx][Ff])", r"\1/\2", text)       # "01 /Xf" → "01/Xf"
    text = re.sub(r"([Dd][Gg][Aa][Vv])[\s\n]*-", r"\1-", text)     # "DGAV -" → "DGAV-"
    text = re.sub(r"([Ee][Dd][Mm])\s*/\s*(\d+)", r"\1/\2", text)   # "EDM /25" → "EDM/25"
    text = re.sub(r"[ \t]+", " ", text)                            # espaços múltiplos
    text = re.sub(r"\n{2,}", "\n", text)

    pattern = re.compile(
        r"(?:PROGRAMA\s+NACIONAL\s+DE\s+PROSPE[ÇC][AÃ]O\s+DE\s+PRAGAS\s+DE\s+QUARENTENA)",
        re.IGNORECASE,
    )
    marks = [m.start() for m in pattern.finditer(text)]

    if not marks:
        print("🔍 Nenhum cabeçalho encontrado — tratado como 1 requisição.")
        return [text]
    if len(marks) == 1:
        print("🔍 Apenas 1 cabeçalho — 1 requisição detectada.")
        return [text]

    marks.append(len(text))
    blocos = []
    for i in range(len(marks) - 1):
        start = max(0, marks[i] - 200)            # padding antes
        end = min(len(text), marks[i + 1] + 200)  # padding depois
        bloco = text[start:end].strip()
        if len(bloco) > 400:
            blocos.append(bloco)
        else:
            print(f"⚠️ Bloco {i+1} demasiado pequeno ({len(bloco)} chars) — possivelmente OCR truncado.")
    print(f"🔍 Detetadas {len(blocos)} requisições distintas (por cabeçalho).")
    return blocos



def normalize_date_str(val: str) -> str:
    """
    Corrige datas OCR partidas/coladas:
    - remove quebras/espaços (mantendo '/')
    - respeita 3.º e 6.º caráter quando existirem
    - reconstrói dd/mm/yyyy a partir de dígitos
    """
    if not val:
        return ""
    txt = str(val).strip().replace("-", "/").replace(".", "/")
    # remove espaços, tabs e quebras de linha, mantendo '/'
    txt = re.sub(r"[\u00A0\s]+", "", txt)

    # já em dd/mm/yyyy?
    m_std = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", txt)
    if m_std:
        d, m_, y = map(int, m_std.groups())
        if 1 <= d <= 31 and 1 <= m_ <= 12 and 1900 <= y <= 2100:
            return f"{d:02d}/{m_:02d}/{y:04d}"

    # se 3º e 6º carater forem '/', tentar leitura posicional direta
    if len(txt) >= 10 and txt[2] == "/" and txt[5] == "/":
        try:
            d, m_, y = int(txt[:2]), int(txt[3:5]), int(txt[6:10])
            if 1 <= d <= 31 and 1 <= m_ <= 12 and 1900 <= y <= 2100:
                return f"{d:02d}/{m_:02d}/{y:04d}"
        except Exception:
            pass

    # remover tudo exceto dígitos para reconstrução
    digits = re.sub(r"\D", "", txt)

    # 8 dígitos: ddmmyyyy
    if len(digits) == 8:
        d, m_, y = int(digits[:2]), int(digits[2:4]), int(digits[4:])
        if 1 <= d <= 31 and 1 <= m_ <= 12 and 1900 <= y <= 2100:
            return f"{d:02d}/{m_:02d}/{y:04d}"

    # 9 dígitos (caso típico 23110/2025 → 23/10/2025)
    if len(digits) == 9:
        # heurística: se os dígitos 3..5 forem '110' → mês 10
        if digits[2:5] == "110":
            d, m_, y = int(digits[:2]), 10, int(digits[-4:])
            return f"{d:02d}/{m_:02d}/{y:04d}"
        # fallback: ddmmyyyy nos primeiros 8
        d, m_, y = int(digits[:2]), int(digits[2:4]), int(digits[4:8])
        if 1 <= d <= 31 and 1 <= m_ <= 12:
            return f"{d:02d}/{m_:02d}/{y:04d}"

    # flexível: d/m/aa ou d/m/aaaa
    m_flex = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})$", txt)
    if m_flex:
        d, m_, y = m_flex.groups()
        y = int(y) + (2000 if len(y) == 2 else 0)
        d, m_ = int(d), int(m_)
        if 1 <= d <= 31 and 1 <= m_ <= 12 and 1900 <= y <= 2100:
            return f"{d:02d}/{m_:02d}/{y:04d}"

    return ""


def _is_valid_date(value: str) -> bool:
    if isinstance(value, datetime):
        return True
    norm = normalize_date_str(value)
    if not norm:
        return False
    try:
        dt = datetime.strptime(norm, "%d/%m/%Y")
        return 1900 <= dt.year <= 2100
    except Exception:
        return False


def _to_datetime(value: str):
    if isinstance(value, datetime):
        return value
    norm = normalize_date_str(value)
    if not norm:
        return None
    try:
        dt = datetime.strptime(norm, "%d/%m/%Y")
        return dt if dt.year >= 1900 else None
    except Exception:
        return None




def extract_context_from_text(full_text: str):
    """Extrai informações gerais da requisição (zona, DGAV, datas, nº de amostras)."""
    ctx = {}

    # ───────────────────────────────────────────────
    # Zona
    # ───────────────────────────────────────────────
    m_zona = re.search(r"Xylella\s+fastidiosa\s*\(([^)]+)\)", full_text, re.I)
    ctx["zona"] = m_zona.group(1).strip() if m_zona else "Zona Isenta"

    # ───────────────────────────────────────────────
    # DGAV / Responsável pela colheita
    # ───────────────────────────────────────────────
    responsavel, dgav = None, None
    m_hdr = re.search(
        r"Amostra(?:s|\(s\))?\s*colhida(?:s|\(s\))?\s*por\s*DGAV\s*[:\-]?\s*(.*)",
        full_text,
        re.IGNORECASE,
    )
    if m_hdr:
        tail = full_text[m_hdr.end():]
        linhas = [m_hdr.group(1)] + tail.splitlines()
        for ln in linhas[:4]:
            ln = (ln or "").strip()
            if ln:
                responsavel = ln
                break
        if responsavel:
            responsavel = re.sub(r"\S+@dgav\.pt|\S+@\S+", "", responsavel, flags=re.I)
            responsavel = re.sub(r"PROGRAMA.*|Data.*|N[º°].*", "", responsavel, flags=re.I)
            responsavel = re.sub(r"[:;,.\-–—]+$", "", responsavel).strip()

    if responsavel:
        dgav = f"DGAV {responsavel}".strip() if not re.match(r"^DGAV\b", responsavel, re.I) else responsavel
    else:
        m_d = re.search(r"\bDGAV(?:\s+[A-Za-zÀ-ÿ?]+){1,4}", full_text)
        dgav = re.sub(r"[:;,.\-–—]+$", "", m_d.group(0)).strip() if m_d else None

    ctx["dgav"] = dgav
    ctx["responsavel_colheita"] = None

    # ───────────────────────────────────────────────
    # Datas de colheita (mapeamento com asteriscos, se existir)
    # ───────────────────────────────────────────────
    colheita_map = {}
    for m in re.finditer(r"(\d{1,2}/\d{1,2}/\d{4})\s*\(\s*(\*+)\s*\)", full_text):
        colheita_map[f"({m.group(2).replace(' ', '')})"] = m.group(1)
    if not colheita_map:
        m_simple = re.search(r"Data\s+de\s+colheita\s*[:\-\s]*([0-9/\-\s]+)", full_text, re.I)
        if m_simple:
            only_date = re.sub(r"\s+", "", m_simple.group(1))
            for key in ("(*)", "(**)", "(***)"):
                colheita_map[key] = only_date
    default_colheita = normalize_date_str(next(iter(colheita_map.values()), ""))
    ctx["colheita_map"] = colheita_map
    ctx["default_colheita"] = default_colheita

    # ───────────────────────────────────────────────
    # Data de envio ao laboratório
    # ───────────────────────────────────────────────
    m_envio = re.search(
        r"Data\s+(?:do|de)\s+envio(?:\s+ao\s+laborat[oó]rio)?[:\-\s]*([0-9/\-\s]+)",
        full_text,
        re.I,
    )
    if m_envio:
        ctx["data_envio"] = normalize_date_str(m_envio.group(1))
    elif default_colheita:
        ctx["data_envio"] = default_colheita
    else:
        ctx["data_envio"] = datetime.now().strftime("%d/%m/%Y")

    # ───────────────────────────────────────────────
    # Nº de amostras declaradas (debug + robusto a OCR e placeholders)
    # ───────────────────────────────────────────────
    print("\n──────── OCR RAW EXCERPT ────────")
    sample_zone = re.findall(r"(N.?amostras?.{0,40})", full_text, flags=re.I)
    for s in sample_zone:
        print("👉", s)
    print("────────────────────────────────\n")

    flat = re.sub(r"[\u00A0_\s]+", " ", full_text)  # normaliza espaços e underscores
    flat = flat.replace("–", "-").replace("—", "-")

    # aceita variações e ruído OCR (env1o, II, ll, _, etc.)
    patterns = [
        r"N[º°o]?\s*de\s*amostras(?:\s+neste\s+env[i1]o)?[\s:.\-]*([0-9OoQIl]{1,4})\b",
        r"N[º°o]?\s*amostras.*?([0-9OoQIl]{1,4})\b",
        r"amostras\s*(?:neste\s+env[i1]o)?\s*[:\-]?\s*([0-9OoQIl]{1,4})\b",
        r"n\s*[º°o]?\s*de\s*amostras.*?([0-9OoQIl]{1,4})\b",
        r"N\s*amostras.*?([0-9OoQIl]{1,4})\b",
        r"N.*?amostras.*?([0-9OoQIl]{1,4})\b"
    ]

    found = None
    for pat in patterns:
        m_decl = re.search(pat, flat, re.I)
        if m_decl:
            found = m_decl.group(1)
            break

    if found:
        raw = found.strip()
        # corrige distorções típicas do OCR
        raw = (
            raw.replace("O", "0").replace("o", "0")
               .replace("Q", "0").replace("q", "0")
               .replace("I", "1").replace("l", "1")
               .replace("|", "1").replace("B", "8")
        )
        try:
            ctx["declared_samples"] = int(raw)
        except ValueError:
            ctx["declared_samples"] = 0
    else:
        # fallback adicional: tenta linha completa com "Nº de amostras"
        m_line = re.search(r"(N[º°o]?\s*de\s*amostras[^\n]*)", full_text, re.I)
        if m_line:
            line = re.sub(r"[_\s]+", " ", m_line.group(1))
            m_num = re.search(r"([0-9OoQIl]{1,4})(?!\s*/)\b", line)
            if m_num:
                raw = m_num.group(1)
                raw = (raw.replace("O", "0").replace("o", "0")
                             .replace("Q", "0").replace("q", "0")
                             .replace("I", "1").replace("l", "1"))
                try:
                    ctx["declared_samples"] = int(raw)
                except ValueError:
                    ctx["declared_samples"] = 0
            else:
                ctx["declared_samples"] = 0
        else:
            ctx["declared_samples"] = 0

    print(f"📊 Nº de amostras declaradas detetadas: {ctx['declared_samples']}")
    return ctx


def parse_xylella_tables(result_json, context, req_id=None) -> List[Dict[str, Any]]:
    """Extrai as amostras das tabelas Azure OCR, aplicando o contexto da requisição."""
    out: List[Dict[str, Any]] = []
    tables = result_json.get("analyzeResult", {}).get("tables", [])
    if not tables:
        print("⚠️ Nenhuma tabela encontrada.")
        return out

    for t in tables:
        nc = max(c.get("columnIndex", 0) for c in t.get("cells", [])) + 1
        nr = max(c.get("rowIndex", 0) for c in t.get("cells", [])) + 1
        grid = [[""] * nc for _ in range(nr)]
        for c in t.get("cells", []):
            grid[c["rowIndex"]][c["columnIndex"]] = clean_value(c.get("content", ""))

        for row in grid:
            if not row or not any(row):
                continue
            ref = _clean_ref(row[0]) if len(row) > 0 else ""
            if not ref or re.match(r"^\D+$", ref):
                continue

            hospedeiro = row[2] if len(row) > 2 else ""
            obs        = row[3] if len(row) > 3 else ""

            if _looks_like_natureza(hospedeiro):
                hospedeiro = ""

            tipo = ""
            joined = " ".join([x for x in row if isinstance(x, str)])
            m_tipo = re.search(r"\b(Simples|Composta|Composto|Individual)\b", joined, re.I)
            if m_tipo:
                tipo = m_tipo.group(1).capitalize()
                # 🔧 Correção do erro comum: "Composto" → "Composta"
                if tipo.lower() == "composto":
                    tipo = "Composta"
                obs = re.sub(r"\b(Simples|Composta|Composto|Individual)\b", "", obs, flags=re.I).strip()

            datacolheita = context.get("default_colheita", "")
            m_ast = re.search(r"\(\s*\*+\s*\)", joined)
            if m_ast:
                mark = re.sub(r"\s+", "", m_ast.group(0))
                datacolheita = context.get("colheita_map", {}).get(mark, datacolheita)

            if obs.strip().lower() in ("simples", "composta", "composto", "individual"):
                obs = ""

            out.append({
                "requisicao_id": req_id,
                "datarececao": context["data_envio"],
                "datacolheita": datacolheita,
                "referencia": ref,
                "hospedeiro": hospedeiro,
                "tipo": tipo,
                "zona": context["zona"],
                "responsavelamostra": context["dgav"],
                "responsavelcolheita": context["responsavel_colheita"],
                "observacoes": obs.strip(),
                "procedure": "XYLELLA",
                "datarequerido": context["data_envio"],
                "Score": ""
            })
    # 🧩 Fallback — tentar extrair linhas da tabela via regex se Azure não devolveu cells válidas
    if not out:
        full_text = extract_all_text(result_json)
        # procura padrões tipo "63020099" ou "01/LVT/DGAV-23/..." etc.
        pattern = re.compile(r"(\d{5,8}|[0-9]{1,3}/[A-Z]{1,3}/DGAV[-/]?\d{0,4})", re.I)
        matches = pattern.findall(full_text)
        if matches:
            for ref in matches:
                out.append({
                    "requisicao_id": req_id,
                    "datarececao": context["data_envio"],
                    "datacolheita": context.get("default_colheita", ""),
                    "referencia": ref.strip(),
                    "hospedeiro": "",
                    "tipo": "",
                    "zona": context["zona"],
                    "responsavelamostra": context["dgav"],
                    "responsavelcolheita": context["responsavel_colheita"],
                    "observacoes": "",
                    "procedure": "XYLELLA",
                    "datarequerido": context["data_envio"],
                    "Score": ""
                })
            print(f"🔍 Fallback regex: {len(matches)} amostras detetadas.")

    print(f"✅ {len(out)} amostras extraídas no total (req_id={req_id}).")
    return out

# ───────────────────────────────────────────────
# Dividir em requisições e extrair por bloco
# ───────────────────────────────────────────────
def parse_all_requisitions(result_json: Dict[str, Any], pdf_name: str, txt_path: str | None) -> List[Dict[str, Any]]:
    """
    Divide o documento em blocos (requisições) e devolve uma lista onde cada elemento
    é um dicionário: { "rows": [...amostras...], "expected": nº_declarado }.
    Suporta múltiplas requisições e atribuição exclusiva de tabelas por bloco.
    """
    # Texto global OCR
    if txt_path and os.path.exists(txt_path):
        full_text = Path(txt_path).read_text(encoding="utf-8")
        print(f"📝 Contexto extraído de {os.path.basename(txt_path)}")
    else:
        full_text = extract_all_text(result_json)

    # Detetar nº de requisições
    count, _ = detect_requisicoes(full_text)
    all_tables = result_json.get("analyzeResult", {}).get("tables", []) or []

    # Caso simples (1 requisição)
    if count <= 1:
        context = extract_context_from_text(full_text)
        amostras = parse_xylella_tables(result_json, context, req_id=1)
        expected = context.get("declared_samples", 0)
        return [{"rows": amostras, "expected": expected}] if amostras else []

    # Múltiplas requisições — segmentar por cabeçalhos
    blocos = split_if_multiple_requisicoes(full_text)
    num_blocos = len(blocos)
    out: List[List[Dict[str, Any]]] = [[] for _ in range(num_blocos)]

    # Extrair referências por bloco
    refs_por_bloco: List[List[str]] = []
    for i, bloco in enumerate(blocos, start=1):
        refs_bloco = re.findall(
            r"\b\d{1,3}/[A-Z]{0,2}/DGAV(?:-[A-Z0-9/]+)?|\b\d{2,4}/\d{2,4}/[A-Z0-9\-]+",
            bloco, re.I
        )
        refs_bloco = [r.strip() for r in refs_bloco if len(r.strip()) > 4]
        print(f"   ↳ Bloco {i}: {len(refs_bloco)} referências detectadas")
        refs_por_bloco.append(refs_bloco)

    # Pré-calcular texto de cada tabela
    table_texts = [" ".join(c.get("content", "") for c in t.get("cells", [])) for t in all_tables]

    # Atribuição exclusiva de tabelas por bloco
    assigned_to: List[int] = [-1] * len(all_tables)
    for ti, ttxt in enumerate(table_texts):
        scores = []
        for bi, refs in enumerate(refs_por_bloco):
            if not refs:
                scores.append(0)
                continue
            cnt = sum(1 for r in refs if r in ttxt)
            scores.append(cnt)
        best = max(scores) if scores else 0
        if best > 0:
            bi = scores.index(best)
            assigned_to[ti] = bi

    # fallback: tabelas não atribuídas → distribuição uniforme
    unassigned = [i for i, b in enumerate(assigned_to) if b < 0]
    if unassigned:
        for k, ti in enumerate(unassigned):
            assigned_to[ti] = k % num_blocos

    # Construir amostras por bloco com base na atribuição
    for bi in range(num_blocos):
        try:
            context = extract_context_from_text(blocos[bi])
            tables_filtradas = [all_tables[ti] for ti in range(len(all_tables)) if assigned_to[ti] == bi]
            if not tables_filtradas:
                print(f"⚠️ Bloco {bi+1}: sem tabelas atribuídas (usar todas como fallback).")
                tables_filtradas = all_tables

            local = {"analyzeResult": {"tables": tables_filtradas}}
            amostras = parse_xylella_tables(local, context, req_id=bi+1)
            out[bi] = amostras or []
        except Exception as e:
            print(f"❌ Erro no bloco {bi+1}: {e}")
            out[bi] = []

    # Remover blocos vazios no fim (mantém ordenação)
    out = [req for req in out if req]
    print(f"\n🏁 Concluído: {len(out)} requisições com amostras extraídas (atribuição exclusiva).")

    # 🔹 NOVO: devolve [{rows, expected}] para validação esperadas/processadas
    results = []
    for bi, bloco in enumerate(blocos[:len(out)], start=1):
        ctx = extract_context_from_text(bloco)
        expected = ctx.get("declared_samples", 0)
        results.append({
            "rows": out[bi - 1],
            "expected": expected
        })
    return results

# ───────────────────────────────────────────────
# Escrita no TEMPLATE — 1 ficheiro por requisição
# ───────────────────────────────────────────────


def gerar_nome_excel_corrigido(source_pdf: str, data_envio: str) -> str:
    """
    Substitui a data inicial do nome do PDF pela nova data (com +1 útil).
    Ex: 20251030_ReqX19_27-10 Formulário.pdf -> 20251031_ReqX19_27-10 Formulário.xlsx
    """
    base_pdf = Path(source_pdf).name
    nova_data = get_next_business_day(data_envio)  # YYYYMMDD
    nome_corrigido = re.sub(r"^\d{8}_", f"{nova_data}_", base_pdf)
    return nome_corrigido.replace(".pdf", ".xlsx")


def write_to_template (ocr_rows, out_name, expected_count=None, source_pdf=None):
    if not ocr_rows:
        print(f"⚠️ {out_name}: sem linhas para escrever.")
        return None

    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Template não encontrado: {TEMPLATE_PATH}")

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.worksheets[0]
    start_row = 4

    # Estilos
    yellow_fill = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")
    green_fill  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    gray_fill   = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    bold_center = Font(bold=True, color="000000")

    # Limpa linhas antigas
    for row in range(start_row, 201):
        for col in range(1, 13):
            cell = ws.cell(row=row, column=col)
            cell.value = None
            cell.fill = PatternFill(fill_type=None)
        ws[f"I{row}"].value = None

    # Funções auxiliares
    def normalize_date_str(val: str) -> str:
        if not val:
            return ""
        s = re.sub(r"\D", "", str(val))
        if len(s) >= 8:
            d, m, y = int(s[:2]), int(s[2:4]), int(s[4:8])
            if 1 <= d <= 31 and 1 <= m <= 12:
                return f"{d:02d}/{m:02d}/{y:04d}"
        m = re.match(r"^\s*(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})\s*$", str(val))
        if m:
            d, m_, y = map(int, m.groups())
            return f"{d:02d}/{m_:02d}/{y:04d}"
        return str(val).strip()

    def to_excel_date(val: str):
        s = normalize_date_str(val)
        try:
            return datetime.strptime(s, "%d/%m/%Y")
        except Exception:
            return None

    # Extração do req_id do nome do ficheiro PDF
    base = Path(source_pdf or out_name).name
    m = re.search(r"(X\d{2,3})", base, flags=re.I)
    req_id = m.group(1).upper() if m else "X??"
    
        
    # Processar linhas
    for idx, row in enumerate(ocr_rows, start=start_row):
        # Extrair valores da linha OCR
        rececao_val = row.get("datarececao", "")
        colheita_val = row.get("datacolheita", "")
        
        # ───────────────────────────────────────────────
        # 🧭 Coluna A — Data de receção + 1 dia útil
        # ───────────────────────────────────────────────
        base_date = normalize_date_str(rececao_val)
        if base_date and re.match(r"\d{2}/\d{2}/\d{4}", str(base_date)):
            try:
                cal = Portugal()
                dt = datetime.strptime(base_date, "%d/%m/%Y").date()
                next_bd = cal.add_working_days(dt, 1)
                ws[f"A{idx}"].value = next_bd
                ws[f"A{idx}"].number_format = "dd/mm/yyyy"
            except Exception:
                ws[f"A{idx}"].value = base_date
                ws[f"A{idx}"].fill = red_fill
        else:
            ws[f"A{idx}"].value = str(rececao_val or "").strip()
            ws[f"A{idx}"].fill = red_fill
    
        # ───────────────────────────────────────────────
        # 🧭 Coluna B — Data de colheita (valor direto)
        # ───────────────────────────────────────────────
        cell_B = ws[f"B{idx}"]
        dt_colheita = to_excel_date(colheita_val)
        if dt_colheita:
            cell_B.value = dt_colheita
            cell_B.number_format = "dd/mm/yyyy"
        else:
            norm = normalize_date_str(colheita_val)
            cell_B.value = norm or str(colheita_val).strip()
            cell_B.fill = red_fill
    
        # ───────────────────────────────────────────────
        # 📄 Restantes colunas
        # ───────────────────────────────────────────────
        ws[f"C{idx}"] = row.get("referencia", "")
        ws[f"D{idx}"] = row.get("hospedeiro", "")
        ws[f"E{idx}"] = row.get("tipo", "")
        ws[f"F{idx}"] = row.get("zona", "")
        ws[f"G{idx}"] = row.get("responsavelamostra", "")
        ws[f"H{idx}"] = row.get("responsavelcolheita", "")
        ws[f"I{idx}"] = ""
    
        # 🧩 Coluna J — Código interno Lab (sem @)
        ws[f"J{idx}"] = f'=TEXT(A{idx},"ddmm")&"{req_id}."&TEXT(ROW()-3,"000")'
    
        # Coluna K — Procedimento
        ws[f"K{idx}"] = row.get("procedure", "")
    
        # ───────────────────────────────────────────────
        # 🚨 Validação visual
        # ───────────────────────────────────────────────
        for col in ("A", "B", "C", "D", "E", "F", "G"):
            c = ws[f"{col}{idx}"]
            if not c.value or str(c.value).strip() == "":
                c.fill = red_fill
    
        if row.get("WasCorrected") or row.get("ValidationStatus") in ("review", "unknown", "no_list"):
            ws[f"D{idx}"].fill = yellow_fill


    # Validação E1:F1
    processed = len(ocr_rows)
    expected = expected_count
    ws.merge_cells("E1:F1")
    cell = ws["E1"]
    val_str = f" {expected or 0} / {processed}"
    cell.value = f"Nº Amostras (Dec./Proc.): {val_str}"
    cell.font = bold_center
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.fill = red_fill if (expected is not None and expected != processed) else green_fill

    # Origem do PDF
    ws.merge_cells("G1:J1")
    pdf_orig_name = Path(source_pdf).name if source_pdf else "(desconhecida)"
    ws["G1"].value = f"Origem: {pdf_orig_name}"
    ws["G1"].font = Font(italic=True, color="555555")
    ws["G1"].alignment = Alignment(horizontal="left", vertical="center")
    ws["G1"].fill = gray_fill

    # Timestamp
    ws.merge_cells("K1:L1")
    timestamp = datetime.now().strftime("%d/%m/%Y %H:%M")
    ws["K1"].value = f"Processado em: {timestamp}"
    ws["K1"].font = Font(italic=True, color="555555")
    ws["K1"].alignment = Alignment(horizontal="right", vertical="center")
    ws["K1"].fill = gray_fill

    # ───────────────────────────────────────────────
    # 💾 Nome final baseado na data_envio (data_rececao + 1 dia útil)
    # ───────────────────────────────────────────────
    try:
        # Tenta usar a última data calculada (coluna A)
        data_envio = next_bd
    except NameError:
        # Fallback se a variável não existir
        data_envio = datetime.now().date()
    
    # Converter para datetime se necessário
    if not isinstance(data_envio, datetime):
        data_envio = datetime.combine(data_envio, datetime.min.time())
    
    # Extrair data como YYYYMMDD
    data_util = data_envio.strftime("%Y%m%d")
    
    # Nome base sem prefixo de data anterior
    base_name = Path(out_name).stem
    base_name = re.sub(r"^\d{8}_", "", base_name)
    
    # Novo nome → YYYYMMDD_restante.xlsx
    new_name = f"{data_util}_{base_name}.xlsx"
    
    out_path = Path(OUTPUT_DIR) / new_name
    wb.save(out_path)
    
    print(f"📁 Ficheiro gravado: {out_path}")
    return str(out_path)





# ───────────────────────────────────────────────
# Log opcional (compatível com o teu Colab)
# ───────────────────────────────────────────────
def append_process_log(pdf_name, req_id, processed, expected, out_path=None, status="OK", error_msg=None):
    log_path = os.path.join(OUTPUT_DIR, "process_log.csv")
    today_str = datetime.now().strftime("%Y-%m-%d")
    summary_path = os.path.join(OUTPUT_DIR, f"process_summary_{today_str}.txt")

    exists = os.path.exists(log_path)
    with open(log_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        if not exists:
            writer.writerow(["DataHora","PDF","ReqID","Processadas","Requisitadas","OutputExcel","Status","Mensagem"])
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        writer.writerow([ts, os.path.basename(pdf_name), req_id, processed, expected or "", out_path or "", status, error_msg or ""])

    try:
        with open(summary_path, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().strftime('%H:%M:%S')}] {os.path.basename(pdf_name)} | Req {req_id} | {processed}/{expected or '?'} | {status} {os.path.basename(out_path or '')}\n")
    except Exception:
        pass

# ───────────────────────────────────────────────
# API pública usada pela app Streamlit
# ───────────────────────────────────────────────
def process_pdf_sync(pdf_path: str) -> List[Dict[str, Any]]:
    """
    Executa o OCR Azure direto ao PDF e o parser Colab integrado.
    Devolve: lista de requisições, cada uma no formato:
      {
        "rows": [ {dados da amostra}, ... ],
        "expected": nº_declarado
      }
    A escrita do Excel é feita a jusante (no xylella_processor.py),
    1 ficheiro por requisição, com validação esperadas/processadas.
    """
    base = os.path.basename(pdf_path)
    print(f"\n🧪 Início de processamento: {base}")

    # 1️⃣ Executar OCR Azure
    result_json = azure_analyze_pdf(pdf_path)

    # 2️⃣ Guardar texto OCR global para debug
    txt_path = OUTPUT_DIR / f"{os.path.splitext(base)[0]}_ocr_debug.txt"
    txt_path.write_text(extract_all_text(result_json), encoding="utf-8")
    print(f"📝 Texto OCR bruto guardado em: {txt_path}")

    # 3️⃣ Parser — dividir em requisições e extrair amostras
    req_results = parse_all_requisitions(result_json, pdf_path, str(txt_path))

    # 4️⃣ Log e resumo de validação
    total_amostras = sum(len(req["rows"]) for req in req_results)
    print(f"✅ {base}: {len(req_results)} requisições, {total_amostras} amostras extraídas.")

    # 5️⃣ Escrever ficheiros Excel diretamente (para compatibilidade cloud)
    created_files = []
    for i, req in enumerate(req_results, start=1):
        rows = req.get("rows", [])
        expected = req.get("expected", 0)

        if not rows:
            print(f"⚠️ Requisição {i}: sem amostras — ignorada.")
            continue

        base_name = os.path.splitext(base)[0]
        out_name = f"{base_name}_req{i}.xlsx" if len(req_results) > 1 else f"{base_name}.xlsx"

        out_path = write_to_template(rows, out_name, expected_count=expected, source_pdf=pdf_path)
        created_files.append(out_path)

        diff = len(rows) - (expected or 0)
        if len(rows) > 0 and diff != 0:
            if expected == 0:
                print(f"⚠️ Requisição {i}: {len(rows)} amostras vs ausente/0 declaradas (diferença {diff:+d}).")
            else:
                print(f"⚠️ Requisição {i}: {len(rows)} amostras vs {expected} declaradas (diferença {diff:+d}).")
        else:
            print(f"✅ Requisição {i}: {len(rows)} amostras gravadas → {out_path}")

    print(f"🏁 {base}: {len(created_files)} ficheiro(s) Excel gerado(s).")
    # Guardar excerto OCR para debug de "Nº de amostras"
    try:
        ocr_text_path = OUTPUT_DIR / f"{Path(pdf_path).stem}_ocr_debug_excerpt.txt"
        with open(ocr_text_path, "w", encoding="utf-8") as dbg:
            with open(OUTPUT_DIR / f"{Path(pdf_path).stem}_ocr_debug.txt", "r", encoding="utf-8") as full:
                text = full.read()
                # guarda apenas 400 caracteres à volta de "amostra" para ver o contexto real
                match = re.search(r".{0,200}amostra.{0,200}", text, re.I)
                dbg.write(match.group(0) if match else text[:400])
        print(f"🪶 Excerto OCR guardado em: {ocr_text_path}")
    except Exception as e:
        print(f"[WARN] Não foi possível gerar excerto OCR: {e}")

    # 6️⃣ Garantir que o PDF existe e copiá-lo para /tmp (compatível com Streamlit Cloud)
    try:
        pdf_src = Path(pdf_path)
        pdf_copy = Path("/tmp") / pdf_src.name
        if not pdf_copy.exists():
            shutil.copy2(pdf_src, pdf_copy)
            print(f"📂 PDF copiado para /tmp: {pdf_copy}")
    except Exception as e:
        print(f"[WARN] Falha ao copiar PDF para /tmp ({pdf_src}) → {e}")

    # Incluir o PDF (cópia ou original) na lista final
    pdf_final = pdf_copy if pdf_copy.exists() else pdf_src
    if pdf_final.exists():
        created_files.append(str(pdf_final))
        print(f"📄 PDF incluído na lista final: {pdf_final}")
    else:
        print(f"[WARN] PDF não encontrado: {pdf_final}")

    # 7️⃣ Gerar summary.txt formatado
    try:
        # Obter prefixo de data do primeiro Excel
        first_excel = next((f for f in created_files if f.endswith(".xlsx")), None)
        data_prefix = ""
        if first_excel:
            match = re.match(r"^(\d{8})_", Path(first_excel).stem)
            if match:
                data_prefix = match.group(1)

        # Definir nome do summary
        summary_name = f"{data_prefix}_summary.txt" if data_prefix else "summary.txt"
        summary_path = Path("/tmp") / summary_name

        # Escrever conteúdo
        with open(summary_path, "w", encoding="utf-8") as s:
            s.write(f"Resumo de processamento — {datetime.now():%d/%m/%Y %H:%M}\n")
            s.write(f"Ficheiro original: {base}\n")
            s.write(f"Total de requisições válidas: {len(valid_reqs)}\n")
            s.write(f"Total de amostras: {total_amostras}\n\n")
            s.write("Ficheiros incluídos:\n")
            for f in created_files:
                s.write(f"  - {Path(f).name}\n")

        created_files.append(str(summary_path))
        print(f"🪶 Summary criado: {summary_path}")

    except Exception as e:
        print(f"[WARN] Falha ao gerar summary.txt: {e}")



