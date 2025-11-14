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

# 🔎 Reconstruir referências quando vêm partidas entre contador e /XF/
def merge_counter_and_ref(row, next_row=None):
    """
    Reconstroi refs partidas:
    Ex:
        ["3", ""] + ["", "/XF/ICNF-C/..."] → "3/XF/ICNF-C/..."
        ["3", "/XF/..."] → "3/XF/..."
        ["", "/XF/..."] → usa contador da linha anterior
    """
    # pega contador se existir
    contador = None
    if row and re.fullmatch(r"\d{1,3}", row[0].strip()):
        contador = row[0].strip()

    # primeira parte da referência (se estiver na mesma linha)
    ref_raw = None
    if len(row) > 1 and "/XF/" in row[1]:
        ref_raw = row[1].strip()

    # referência pode vir na linha seguinte se estiver vazia aqui
    if not ref_raw and next_row and len(next_row) > 0:
        if "/XF/" in next_row[0]:
            ref_raw = next_row[0].strip()
        elif len(next_row) > 1 and "/XF/" in next_row[1]:
            ref_raw = next_row[1].strip()

    if not ref_raw:
        return None

    # juntar contador quando existir
    if contador:
        final = f"{contador}{ref_raw}"
        # normalizar “1/XF/...” → garante a barra correta
        final = re.sub(r"^(\d+)\s*/?", r"\1/", final)
        return final

    return ref_raw


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
    """
    Conta quantas requisições existem no texto OCR.
    Suporta:
      - 'Programa nacional de Prospeção de pragas de quarentena'
      - 'Prospeção de: Xylella fastidiosa em Zonas Demarcadas'
    """
    pattern = re.compile(
        r"(?:PROGRAMA\s+NACIONAL\s+DE\s+PROSPE[ÇC][AÃ]O\s+DE\s+PRAGAS\s+DE\s+QUARENTENA"
        r"|PROSPE[ÇC][AÃ]O\s*DE:?\s*XYLELLA\s+FASTIDIOSA\s+EM\s+ZONAS\s+DEMARCADAS)",
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
        r"(?:PROGRAMA\s+NACIONAL\s+DE\s+PROSPE[ÇC][AÃ]O\s+DE\s+PRAGAS\s+DE\s+QUARENTENA"
        r"|PROSPE[ÇC][AÃ]O\s*DE:?\s*XYLELLA\s+FASTIDIOSA\s+EM\s+ZONAS\s+DEMARCADAS)",
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

def normalize_ocr_line(ln: str) -> str:
    """
    Normalização agressiva de uma linha OCR:
    - remove caracteres invisíveis
    - colapsa espaços
    - normaliza barras (/)
    - normaliza /XF/ independente de maiúsculas
    - corrige espaço entre nº e primeira barra
    """
    if not ln:
        return ""

    # remover caracteres invisíveis / estranhos
    ln = re.sub(r"[\u200b\u00A0\r\t\f\v]", "", str(ln))

    # normalizar espaços múltiplos
    ln = re.sub(r"\s+", " ", ln).strip()

    # normalizar espaços à volta de barras -> " / XF / " -> "/XF/"
    ln = re.sub(r"\s*/\s*", "/", ln)

    # normalizar /XF/ (XF, Xf, xf, xF)
    ln = re.sub(r"/x[fF]/", "/XF/", ln, flags=re.IGNORECASE)

    # corrigir espaço entre nº e primeira barra: "91 /XF" -> "91/XF"
    ln = re.sub(r"^(\d{1,3})\s+/", r"\1/", ln)

    # se ainda houver espaços imediatamente antes de "/", remove
    ln = re.sub(r"\s+/", "/", ln)

    return ln



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
    """Extrai informações gerais da requisição (zona, DGAV/ICNF, datas, nº de amostras)."""
    ctx: Dict[str, Any] = {}

    # ───────────────────────────────────────────────
    # 🟩 1. Zona (ICNF "Zona demarcada" ENTRE "Zona demarcada" e "Entidade:")
    # ───────────────────────────────────────────────
    zona = None

    # Caso típico OCR: "Zona demarcada: Centro/Covilhã-Fundão Entidade: ICNF ..."
    m_zd_span = re.search(
        r"Zona\s+demarcada\s*:?\s*(.+?)(?=\s+Entidade\s*:)",
        full_text,
        flags=re.I | re.S,      # DOTALL para apanhar se houver quebra de linha
    )
    if m_zd_span:
        zona = m_zd_span.group(1).strip()
    else:
        # Fallback antigo: tudo depois de "Zona demarcada:" até ao fim da linha
        m_zd_line = re.search(
            r"Zona\s+demarcada\s*:?\s*(.+)",
            full_text,
            flags=re.I,
        )
        if m_zd_line:
            zona = m_zd_line.group(1).strip()

    if zona:
        # normalizar espaços, mas NÃO mexer em caracteres como "/"
        zona = re.sub(r"\s+", " ", zona)
        ctx["zona"] = zona
        ctx["template_tipo"] = "ZONAS_DEMARCADAS"
    else:
        # Ex: "Prospeção de: Xylella fastidiosa (Zona Isenta)"
        m_zona = re.search(r"Xylella\s+fastidiosa\s*\(([^)]+)\)", full_text, re.I)
        ctx["zona"] = m_zona.group(1).strip() if m_zona else "Zona Isenta"
        ctx["template_tipo"] = "PROGRAMA_NACIONAL"


    # ───────────────────────────────────────────────
    # 🟩 2. Entidade (ICNF / DGAV)
    # Ex: "Entidade: ICNF"
    # ───────────────────────────────────────────────
    entidade = None
    m_ent = re.search(r"Entidade\s*:\s*(.+)", full_text, re.I)
    if m_ent:
        entidade = m_ent.group(1).strip()
        entidade = re.sub(r"T[ée]cnico\s+respons[aá]vel.*$", "", entidade, flags=re.I).strip()

    # ───────────────────────────────────────────────
    # 🟩 3. Técnico responsável (limpeza robusta)
    # Ex OCR: "Técnico responsável: António Cabanas Data de envio..."
    # ───────────────────────────────────────────────
    tecnico_resp = None
    m_tecnico = re.search(
        r"T[ée]cnico\s+respons[aá]vel\s*:\s*(.+?)(?:\n|$|Data\s+envio|Data\s+de\s+envio|Ref[ºª]|Hospedeiro|Tipo\s+amostra)",
        full_text,
        flags=re.I | re.S,
    )
    if m_tecnico:
        tecnico_resp = m_tecnico.group(1).strip()
        # Remover lixo colado pelo OCR
        tecnico_resp = re.sub(r"(Data\s+.*)$", "", tecnico_resp, flags=re.I).strip()
        tecnico_resp = re.sub(r"(Ref[ºª].*)$", "", tecnico_resp, flags=re.I).strip()
        tecnico_resp = re.sub(r"(Hospedeiro.*)$", "", tecnico_resp, flags=re.I).strip()
        tecnico_resp = re.sub(r"(Tipo\s+amostra.*)$", "", tecnico_resp, flags=re.I).strip()

    if entidade:
        ctx["dgav"] = entidade
    ctx["responsavel_colheita"] = tecnico_resp

    # ───────────────────────────────────────────────
    # 🟩 4. Fallback DGAV antigo ("Amostra colhida por DGAV: X")
    # ───────────────────────────────────────────────
    if not entidade:
        m_hdr = re.search(
            r"Amostra(?:s|\(s\))?\s*colhida(?:s|\(s\))?\s*por\s*DGAV\s*[:\-]?\s*(.*)",
            full_text,
            re.IGNORECASE,
        )
        if m_hdr:
            responsavel = m_hdr.group(1).strip()
            responsavel = re.sub(r"\S+@\S+", "", responsavel).strip()
            responsavel = re.sub(r"Data.*", "", responsavel, flags=re.I).strip()
            if responsavel:
                ctx["dgav"] = f"DGAV {responsavel}".strip()

        if not tecnico_resp:
            ctx["responsavel_colheita"] = None

    # Garante que "dgav" existe sempre (evita KeyError)
    if "dgav" not in ctx:
        ctx["dgav"] = ""

    # ───────────────────────────────────────────────
    # 🟩 5. Data de colheita das amostras (inclui "Datas de recolha...")
    # ───────────────────────────────────────────────
    m_col = None

    # Novo formato ICNF: "Datas de recolha de amostras: 04-11-2025"
    if not m_col:
        m_col = re.search(
            r"Datas?\s+de\s+recolha\s+de\s+amostras\s*[:\-]?\s*([0-9/\-\s]+)",
            full_text,
            re.I,
        )

    # Outro formato: "Data colheita das amostras: 03/11/2025"
    if not m_col:
        m_col = re.search(
            r"Data\s+colheita\s+das\s+amostras\s*[:\-]?\s*([0-9/\-\s]+)",
            full_text,
            re.I,
        )

    # Fallback antigo: "Data de colheita: 03/11/2025"
    if not m_col:
        m_col = re.search(
            r"Data\s+de\s+colheita\s*[:\-\s]*([0-9/\-\s]+)",
            full_text,
            re.I,
        )

    if m_col:
        ctx["default_colheita"] = normalize_date_str(m_col.group(1))
    else:
        ctx["default_colheita"] = ""

    # ───────────────────────────────────────────────
    # 🟩 6. Data de envio ao laboratório
    # Suporta:
    #   "Data de envio das amostras ao laboratório: 07/11/2025"
    #   "Data envio amostras ao laboratório: 07/11/2025"
    #   "Data de envio: 07/11/2025"
    # ───────────────────────────────────────────────
    m_envio = re.search(
        r"Data\s+(?:do|de)?\s*envio.*?([0-9]{1,2}[\/\-\s][0-9]{1,2}[\/\-\s][0-9]{2,4})",
        full_text,
        re.I | re.S,
    )

    if m_envio:
        ctx["data_envio"] = normalize_date_str(m_envio.group(1))
    else:
        # fallback: se não houver data de envio, usar colheita ou hoje
        if ctx.get("default_colheita"):
            ctx["data_envio"] = ctx["default_colheita"]
        else:
            ctx["data_envio"] = datetime.now().strftime("%d/%m/%Y")

    # ───────────────────────────────────────────────
    # 🟩 7. Nº de amostras — vários formatos
    #  a) Novo formato DGAV: "Total: 27/35 amostras"
    #  b) ICNF: "Total:\n30\nAmostras" ou "Total: 30"
    #  c) Fallback antigo: "Nº de amostras ..."
    # ───────────────────────────────────────────────
    ctx["declared_samples"] = 0

    # a) "Total: 27/35 amostras" → usar SEMPRE o ÚLTIMO match do bloco
    matches_frac = re.findall(
        r"Total\s*[:\-]?\s*\d+\s*/\s*([0-9]{1,4})\s*amostras",
        full_text,
        re.I,
    )
    if matches_frac:
        try:
            ctx["declared_samples"] = int(matches_frac[-1])
        except Exception:
            ctx["declared_samples"] = 0

    if ctx["declared_samples"] == 0:
        # b) "Total:\n30\nAmostras" ou "Total: 30" → usar o ÚLTIMO match
        matches_simple = re.findall(
            r"Total\s*[:\-]?\s*([0-9]{1,4})\s*(?:amostras)?\b",
            full_text,
            re.I,
        )
        if matches_simple:
            try:
                ctx["declared_samples"] = int(matches_simple[-1])
            except Exception:
                ctx["declared_samples"] = 0

    # c) Fallback antigo (Nº de amostras ...)
    if ctx["declared_samples"] == 0:
        flat = re.sub(r"[\u00A0_\s]+", " ", full_text)
        patterns = [
            r"N[º°o]?\s*de\s*amostras.*?([0-9OoQIl]{1,4})\b",
            r"N[º°o]?\s*amostras.*?([0-9OoQIl]{1,4})\b",
        ]
        for pat in patterns:
            m2 = re.search(pat, flat, re.I)
            if m2:
                raw = m2.group(1)
                raw = (
                    raw.replace("O", "0")
                    .replace("o", "0")
                    .replace("Q", "0")
                    .replace("I", "1")
                )
                try:
                    ctx["declared_samples"] = int(raw)
                except Exception:
                    pass
                break

    return ctx


def parse_xylella_from_text_block(block_text: str, context: Dict[str, Any], req_id: int = 1) -> List[Dict[str, Any]]:
    """
    Parser genérico baseado em texto (linhas) que funciona para:
      - Template antigo DGAV (com linha de 'Natureza da amostra')
      - Template novo ICNF (sem 'Natureza da amostra')

    Estrutura típica por amostra (vertical):

      [ref]
      [natureza?]      ← opcional, só no template antigo
      [hospedeiro]
      [tipo]           ← 'simples', 'composta', 'Amostra composta', ...

    O objectivo é:
      • reconhecer ref (63020090 ou 91/Xf/DGAVC/MRLRA/25)
      • saltar natureza se existir
      • ler hospedeiro
      • ler tipo
    """

    lines_raw = block_text.splitlines()
    lines: List[str] = []
    for ln in lines_raw:
        ln = ln.replace("\u00A0", " ")
        ln = re.sub(r"\s+", " ", ln).strip()
        if ln:
            lines.append(ln)

    # contexto de datas
    data_envio = context.get("data_envio", datetime.now().strftime("%d/%m/%Y"))
    data_colheita = context.get("default_colheita", data_envio)

    results: List[Dict[str, Any]] = []

    # referência: ou só dígitos (63020090) ou padrão 91/Xf/DGAV...
    ref_re = re.compile(
        r"^(\d{5,8}|\d{1,3}\s*/\s*X[fF][^ ]*)$"
    )

    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]

        m_ref = ref_re.match(line)
        if not m_ref:
            i += 1
            continue

        # referência normalizada
        ref = m_ref.group(1)
        ref = re.sub(r"\s*/\s*", "/", ref).strip()

        i += 1

        # ── opcional: natureza da amostra (aparece no template antigo)
        natureza = None
        while i < n and not lines[i].strip():
            i += 1
        if i < n:
            low = lines[i].lower()
            if _looks_like_natureza(low) or "partes de vegetais" in low or "insetos" in low:
                natureza = lines[i].strip()
                i += 1

        # ── hospedeiro
        while i < n and not lines[i].strip():
            i += 1
        hospedeiro = ""
        if i < n:
            hospedeiro = lines[i].strip()
            i += 1

        # ── tipo (simples/composta/amostra simples/amostra composta)
        while i < n and not lines[i].strip():
            i += 1
        tipo = ""
        if i < n:
            tipo_line = lines[i].strip()
            m_tipo = re.search(r"(simples|composta|amostra simples|amostra composta)", tipo_line, re.I)
            if m_tipo:
                t = m_tipo.group(1).lower()
                if "composta" in t:
                    tipo = "Composta"
                else:
                    tipo = "Simples"
                i += 1
            else:
                # se não reconhecer tipo, pode ser continuação do hospedeiro
                hospedeiro = (hospedeiro + " " + tipo_line).strip()
                i += 1

        results.append({
            "requisicao_id": req_id,
            "datarececao": data_envio,
            "datacolheita": data_colheita,
            "referencia": ref,
            "hospedeiro": hospedeiro,
            "tipo": tipo,
            "zona": context.get("zona", ""),
            "responsavelamostra": context.get("dgav", ""),
            "responsavelcolheita": context.get("responsavel_colheita", ""),
            "observacoes": "",
            "procedure": "XYLELLA",
            "datarequerido": data_envio,
            "Score": "",
        })

    print(f"✅ [fallback texto] Extraídas {len(results)} amostras no bloco (req_id={req_id}).")
    return results


# ───────────────────────────────────────────────
# Parser ICNF – "Prospeção de: Xylella fastidiosa em Zonas Demarcadas"
# ───────────────────────────────────────────────
def parse_icnf_requisition(result_json, full_text, pdf_name, txt_path=None):
    ctx = extract_context_from_text(full_text)
    data_envio = ctx["data_envio"]
    data_colheita = ctx["default_colheita"] or data_envio

    raw_lines = [ln.strip() for ln in full_text.splitlines() if ln.strip()]
    clean_lines = [normalize_ocr_line(ln) for ln in raw_lines]

    rows = []

    # 1) Combinar linhas (caso venham partidas)
    combined = []
    i = 0
    while i < len(clean_lines):
        ln = clean_lines[i]
        if re.match(r"^\d{1,3}\s+/XF/", ln) and i + 1 < len(clean_lines):
            combined.append(ln + " " + clean_lines[i + 1])
            i += 2
        else:
            combined.append(ln)
            i += 1

    # 2) Padrões
    pattern = re.compile(
        r"""
        ^\s*
        (?P<num>\d{1,3})\s+
        (?P<ref>/XF/[A-Z0-9\-\/]+)\s+
        (?P<hosp>[A-Za-zÀ-ÿ\s\.\-]+?)\s+
        (?P<tipo>Simples|Composta|Individual)(?:\s*\(\d+\))?
        """,
        re.I | re.VERBOSE,
    )

    for ln in combined:
        m = pattern.search(ln)
        if not m:
            continue

        tipo = m.group("tipo").capitalize()
        rows.append({
            "requisicao_id": 1,
            "datarececao": data_envio,
            "datacolheita": data_colheita,
            "referencia": m.group("ref").strip(),
            "hospedeiro": m.group("hosp").strip(),
            "tipo": tipo,
            "zona": ctx["zona"],
            "responsavelamostra": ctx["dgav"],
            "responsavelcolheita": ctx["responsavel_colheita"],
            "observacoes": "",
            "procedure": "XYLELLA",
            "datarequerido": data_envio,
            "Score": "",
        })

    expected = ctx["declared_samples"] or len(rows)
    return [{"rows": rows, "expected": expected}] if rows else []

def parse_xylella_tables(
    result_json,
    context,
    req_id: int | None = None,
    col_ref: int = 0,
    col_hosp: int = 1,
    col_obs: int = 2,
) -> List[Dict[str, Any]]:
    """
    Extrai as amostras das tabelas Azure OCR, aplicando o contexto da requisição.
    Suporta:
      • Template DGAV antigo  → 4 colunas: ref, natureza, hosp, tipo
      • Template ICNF novo    → 3 colunas: ref, hospedeiro, tipo
    """

    # Detectar template ICNF → não tem coluna de observações
    if context.get("template_tipo") == "ZONAS_DEMARCADAS":
        col_ref = 0
        col_hosp = 1
        col_obs = -1   # não existe observações

    out: List[Dict[str, Any]] = []
    tables = result_json.get("analyzeResult", {}).get("tables", [])
    if not tables:
        return out

    for t in tables:
        # reconstrução da grelha
        nc = max(c.get("columnIndex", 0) for c in t.get("cells", [])) + 1
        nr = max(c.get("rowIndex", 0) for c in t.get("cells", [])) + 1
        grid = [[""] * nc for _ in range(nr)]
        for c in t.get("cells", []):
            grid[c["rowIndex"]][c["columnIndex"]] = clean_value(c.get("content", ""))

        for row in grid:
            if not row or not any(row):
                continue

            # referência
            ref = merge_counter_and_ref(row, next_row=grid[row_index+1] if row_index+1 < nr else None)
            if not ref:
                continue
            ref = _clean_ref(ref)

            # hospedeiro
            hospedeiro = row[col_hosp] if len(row) > col_hosp else ""
            if _looks_like_natureza(hospedeiro):
                hospedeiro = ""

            # observações (só DGAV antigo)
            obs = ""
            if col_obs >= 0 and len(row) > col_obs:
                obs = row[col_obs]

            # tipo (extraído por regex)
            joined = " ".join(x for x in row if isinstance(x, str))
            tipo = ""
            m_tipo = re.search(r"\b(Simples|Composta|Individual|Composto)\b", joined, re.I)
            if m_tipo:
                tipo = m_tipo.group(1).capitalize()
                if tipo.lower() == "composto":
                    tipo = "Composta"

            # data colheita
            datacolheita = context.get("default_colheita", "")

            out.append({
                "requisicao_id": req_id,
                "datarececao": context.get("data_envio", ""),
                "datacolheita": datacolheita,
                "referencia": ref,
                "hospedeiro": hospedeiro,
                "tipo": tipo,
                "zona": context.get("zona", ""),   # sempre do contexto
                "responsavelamostra": context.get("dgav", ""),
                "responsavelcolheita": context.get("responsavel_colheita", ""),
                "observacoes": obs.strip(),
                "procedure": "XYLELLA",
                "datarequerido": context.get("data_envio", ""),
                "Score": "",
            })

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

    # 🔎 DETEÇÃO DO NOVO TEMPLATE ICNF / ZONAS DEMARCADAS
    icnf_pattern = re.compile(
        r"prospe[cç][aã]o\s*de:?\s*xylella\s+fastidiosa\s+em\s+zonas\s+demarcadas",
        re.IGNORECASE,
    )
    is_icnf = icnf_pattern.search(full_text) is not None

    # ───────────────────────────────────────────────
    # Caminho normal (template antigo DGAV / ICNF)
    # ───────────────────────────────────────────────
    # Detetar nº de requisições
    count, _ = detect_requisicoes(full_text)
    all_tables = result_json.get("analyzeResult", {}).get("tables", []) or []

    # Caso simples (1 requisição)
    if count <= 1:
        context = extract_context_from_text(full_text)

        # 1º tentar via tabelas (comportamento original)
        amostras = parse_xylella_tables(result_json, context, req_id=1)

        # Fallback: se não vier nada das tabelas, usar parser baseado em texto (linhas)
        if not amostras:
            print("⚠️ Nenhuma amostra via tables — a usar fallback de texto.")
            amostras = parse_xylella_from_text_block(full_text, context, req_id=1)

        expected = context.get("declared_samples", len(amostras))
        return [{"rows": amostras, "expected": expected}] if amostras else []

    # ───────────────────────────────────────────────
    # Múltiplas requisições — segmentar por cabeçalhos
    # ───────────────────────────────────────────────
    blocos = split_if_multiple_requisicoes(full_text)
    num_blocos = len(blocos)
    out: List[List[Dict[str, Any]]] = [[] for _ in range(num_blocos)]

    # Extrair referências por bloco
    refs_por_bloco: List[List[str]] = []
    for i, bloco in enumerate(blocos, start=1):
        refs_bloco = re.findall(
            r"\b\d{1,3}/[A-Z]{0,2}/DGAV(?:-[A-Z0-9/]+)?|\b\d{2,4}/\d{2,4}/[A-Z0-9\-]+",
            bloco,
            re.I,
        )
        refs_bloco = [r.strip() for r in refs_bloco if len(r.strip()) > 4]
        print(f"   ↳ Bloco {i}: {len(refs_bloco)} referências detectadas")
        refs_por_bloco.append(refs_bloco)

    # Pré-calcular texto de cada tabela
    table_texts = [
        " ".join(c.get("content", "") for c in t.get("cells", [])) for t in all_tables
    ]

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
            tables_filtradas = [
                all_tables[ti]
                for ti in range(len(all_tables))
                if assigned_to[ti] == bi
            ]
            if not tables_filtradas:
                print(f"⚠️ Bloco {bi+1}: sem tabelas atribuídas (usar todas como fallback).")
                tables_filtradas = all_tables

            local = {"analyzeResult": {"tables": tables_filtradas}}
            amostras = parse_xylella_tables(local, context, req_id=bi + 1)

            if not amostras:
                print(f"⚠️ Bloco {bi+1}: tables vazias — a usar fallback de texto.")
                amostras = parse_xylella_from_text_block(
                    blocos[bi], context, req_id=bi + 1
                )

            out[bi] = amostras or []
        except Exception as e:
            print(f"❌ Erro no bloco {bi+1}: {e}")
            out[bi] = []

    # Remover blocos vazios no fim (mantém ordenação)
    out = [req for req in out if req]
    print(f"\n🏁 Concluído: {len(out)} requisições com amostras extraídas (atribuição exclusiva).")

    # 🔹 Devolve [{rows, expected}] para validação esperadas/processadas
    results: List[Dict[str, Any]] = []
    for bi, bloco in enumerate(blocos[: len(out)], start=1):
        ctx = extract_context_from_text(bloco)
        expected = ctx.get("declared_samples") or len(out[bi - 1])
        results.append(
            {
                "rows": out[bi - 1],
                "expected": expected,
            }
        )

    return results

# ───────────────────────────────────────────────
# Parser ICNF – "Prospeção de: Xylella fastidiosa em Zonas Demarcadas"
# ───────────────────────────────────────────────
def parse_icnf_requisition(
    result_json: Dict[str, Any],
    full_text: str,
    pdf_name: str,
    txt_path: Optional[str] = None
) -> List[Dict[str, Any]]:
    """
    Parser dedicado ao template ICNF / XF / Zonas Demarcadas.

    Assume que o documento inteiro corresponde a 1 requisição
    e que as linhas das amostras têm formato aproximado:

        1 /XF/ICNFC/COV-FND/AC/25 Accacia dealbata Individual
        2 /XF/ICNFC/COV-FND/AC/25 Pteridium aquilinium Composta 3
    """

    # Contexto base (data_envio, zona, etc.). Se não encontrar nada, usa defaults da função.
    ctx = extract_context_from_text(full_text)
    data_envio = ctx.get("data_envio", datetime.now().strftime("%d/%m/%Y"))
    data_colheita = ctx.get("default_colheita", data_envio)

    rows: List[Dict[str, Any]] = []

    # Normalização leve de linhas
    lines = full_text.replace("\t", " ").splitlines()

    pattern = re.compile(
        r"""
        ^\s*
        (?P<num>\d{1,3})                    # nº amostra
        \s+
        (?P<ref>/XF/[A-Z0-9\-\/]+)          # referência /XF/ICNFC/...
        \s+
        (?P<hosp>[A-Za-zÀ-ÿ\s\.\-]+?)       # hospedeiro
        \s+
        (?P<tipo>Individual|Composta\s*\d+) # tipo
        \s*$
        """,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    for ln in lines:
        m = pattern.match(ln)
        if not m:
            continue

        hosp = m.group("hosp").strip()
        ref = m.group("ref").strip()
        tipo_raw = m.group("tipo").strip()

        tipo = "Individual"
        m_comp = re.match(r"Composta\s*(\d+)", tipo_raw, flags=re.I)
        if m_comp:
            tipo = "Composta"

        rows.append(
            {
                "requisicao_id": 1,
                "datarececao": data_envio,
                "datacolheita": data_colheita,
                "referencia": ref,
                "hospedeiro": hosp,
                "tipo": tipo,
                "zona": ctx.get("zona", ""),
                "responsavelamostra": ctx.get("dgav", ""),
                "responsavelcolheita": ctx.get("responsavel_colheita", ""),
                "observacoes": "",
                "procedure": "XYLELLA",
                "datarequerido": data_envio,
                "Score": "",
            }
        )

    # nº de amostras declaradas → tenta usar lógica atual; se não encontrar, assume len(rows)
    expected = ctx.get("declared_samples") or len(rows)

    print(f"✅ [ICNF] Extraídas {len(rows)} amostras (esperadas: {expected}).")

    return [{"rows": rows, "expected": expected}] if rows else []

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
       # ───────────────────────────────────────────────
    # 🔁 Processar linhas OCR
    # ───────────────────────────────────────────────
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
                ws[f"L{idx}"].value = f"=A{idx}+30"
                ws[f"L{idx}"].number_format = "dd/mm/yyyy"
            except Exception:
                ws[f"A{idx}"].value = base_date
                ws[f"A{idx}"].fill = red_fill
                ws[f"L{idx}"].value = ""
                ws[f"L{idx}"].fill = red_fill
        else:
            ws[f"A{idx}"].value = str(rececao_val or "").strip()
            ws[f"A{idx}"].fill = red_fill
            ws[f"L{idx}"].value = ""
            ws[f"L{idx}"].fill = red_fill
    
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

         # 📅 Coluna L — Data requerido (+30 dias após receção)
        ws[f"L{idx}"].value = f"=A{idx}+30"
        ws[f"L{idx}"].number_format = "dd/mm/yyyy"
        
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
def process_pdf_sync(pdf_path: str) -> list[str]:
    """
    Processa um único PDF:
      - executa OCR Azure,
      - extrai requisições e amostras,
      - gera 1 ficheiro Excel por requisição.
    Retorna: lista de caminhos absolutos dos ficheiros Excel criados.
    """
    base = os.path.basename(pdf_path)
    print(f"\n🧪 Início de processamento: {base}")

    # 1️⃣ Executar OCR Azure
    result_json = azure_analyze_pdf(pdf_path)

    # 2️⃣ Guardar texto OCR para debug
    txt_path = OUTPUT_DIR / f"{Path(base).stem}_ocr_debug.txt"
    txt_path.write_text(extract_all_text(result_json), encoding="utf-8")
    print(f"📝 Texto OCR bruto guardado em: {txt_path}")

    # 3️⃣ Parser — dividir em requisições e extrair amostras
    #    (inclui a lógica interna para ICNF / Zonas Demarcadas)
    req_results = parse_all_requisitions(result_json, pdf_path, str(txt_path))

    valid_reqs = [req for req in req_results if req.get("rows")]
    total_amostras = sum(len(req["rows"]) for req in valid_reqs)
    print(f"✅ {base}: {len(valid_reqs)} requisição(ões) válidas, {total_amostras} amostras extraídas.")

    created_files = []
    for i, req in enumerate(valid_reqs, start=1):
        rows = req.get("rows", [])
        expected = req.get("expected", 0)

        if not rows:
            continue

        # Nome base para o Excel (mantém a data original)
        base_name = Path(pdf_path).stem
        out_name = f"{base_name}_req{i}.xlsx" if len(valid_reqs) > 1 else f"{base_name}.xlsx"

        out_path = write_to_template(rows, out_name, expected_count=expected, source_pdf=pdf_path)
        created_files.append(out_path)
        print(f"💾 Excel criado: {out_path}")

    print(f"🏁 {base}: {len(created_files)} ficheiro(s) Excel gerado(s).")
    return [str(f) for f in created_files if Path(f).exists()]

# ───────────────────────────────────────────────
# API pública usada pela app Streamlit
# ───────────────────────────────────────────────

def process_folder_async(input_dir: str = "/tmp") -> str:
    """
    Processa todos os PDFs em `input_dir` chamando `process_pdf_sync(pdf_path)`.
    Cria:
      - ficheiros Excel (um por requisição)
      - summary.txt
      - ZIP final apenas com XLSX + summary.txt
    Retorna o caminho completo do ZIP criado.
    """
    start_time = time.time()
    input_path = Path(input_dir)
    pdf_files = sorted(input_path.glob("*.pdf"))

    if not pdf_files:
        print("⚠️ Nenhum PDF encontrado na pasta.")
        return ""

    print(f"📂 Início do processamento: {input_path} ({len(pdf_files)} PDF(s))")

    all_excels = []

    # Processar cada PDF → gerar Excels
    for pdf_path in pdf_files:
        base = pdf_path.name
        print(f"\n🔹 A processar: {base}")
        try:
            created = process_pdf_sync(str(pdf_path))
            excels = [f for f in created if str(f).lower().endswith(".xlsx")]
            all_excels.extend(excels)
            print(f"✅ {base}: {len(excels)} ficheiro(s) Excel.")
        except Exception as e:
            print(f"❌ Erro ao processar {base}: {e}")

    elapsed_time = time.time() - start_time

    # Criar summary.txt
    summary_path = input_path / "summary.txt"
    with open(summary_path, "w", encoding="utf-8") as f:
        for pdf_path in pdf_files:
            base = pdf_path.name
            related_excels = [e for e in all_excels if Path(base).stem in Path(e).stem]
            f.write(f"{base}: {len(related_excels)} requisição(ões)\n")
            for e in related_excels:
                f.write(f"   ↳ {Path(e).name}\n")
            f.write("\n")

        f.write("──────────────────────────────\n")
        f.write(f"📊 Total de ficheiros Excel: {len(all_excels)}\n")
        f.write(f"⏱️ Tempo total: {elapsed_time:.1f} segundos\n")
        f.write(f"📅 Executado em: {datetime.now():%d/%m/%Y às %H:%M:%S}\n")

    print(f"🧾 Summary criado: {summary_path}")

    # Criar ZIP apenas com XLSX e summary.txt
    first_pdf = pdf_files[0]
    base_name = Path(first_pdf).stem
    zip_name = f"{base_name}_output.zip"
    zip_path = Path("/tmp") / zip_name  # usa o /tmp global (não apagado pela sessão)

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        # Adiciona todos os Excel
        for e in all_excels:
            e_path = Path(e)
            if e_path.exists():
                zipf.write(e_path, e_path.name)

        # Adiciona summary.txt
        if summary_path.exists():
            zipf.write(summary_path, summary_path.name)

    print(f"📦 ZIP final criado: {zip_path}")
    print(f"✅ Processamento completo ({elapsed_time:.1f}s). ZIP contém {len(all_excels)} Excel(s) + summary.txt")

    return str(zip_path)






















































