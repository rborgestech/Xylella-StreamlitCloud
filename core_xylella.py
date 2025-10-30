# -*- coding: utf-8 -*-
"""
CORE_XYLELLA.PY — Motor principal do processamento Xylella

Responsável por:
 - Executar OCR via Azure (página a página)
 - Extrair tabelas e texto global
 - Analisar blocos de requisição
 - Exportar resultados para o TEMPLATE Excel

Funções principais expostas:
    • process_pdf_sync(pdf_path) -> list[list[dict]]
    • write_to_template(ocr_rows, out_base_name, expected_count=None, source_pdf=None)
"""

from __future__ import annotations

import os
import re
import shutil
import tempfile
from pathlib import Path
from datetime import datetime

from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment

# OCR Azure (módulo externo do projeto)
# Estas funções devem existir em azure_ocr.py:
#   • pdf_to_images(pdf_path)
#   • extract_text_from_image_azure(image_path)
#   • get_analysis_result_azure(result_url)
#   • (opcional) extract_all_text(result_json)
try:
    from azure_ocr import pdf_to_images, extract_text_from_image_azure, get_analysis_result_azure
except Exception:
    # import tardio dentro das funções (permite testes unitários sem deps)
    pdf_to_images = extract_text_from_image_azure = get_analysis_result_azure = None


# ── Caminhos globais ────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", BASE_DIR / "Output"))
OUTPUT_DIR.mkdir(exist_ok=True)
TEMPLATE_PATH = BASE_DIR / "TEMPLATE_PXf_SGSLABIP1056.xlsx"  # mantém o teu template por defeito

# Garantir que a pasta Output existe
OUTPUT_DIR.mkdir(exist_ok=True)


# ----------------------------------------------------------------
#  Utilitários
# ----------------------------------------------------------------

def _now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def _to_list_row(item):
    """Aceita dict (mapeado para colunas template) ou lista já preparada."""
    if isinstance(item, dict):
        # Ordem de colunas esperada no template (linha 6+):
        # 1 Data Receção | 2 Data Colheita | 3 Referência | 4 Hospedeiro
        # 5 Tipo | 6 Zona | 7 Responsável Amostra | 8 Responsável Colheita
        # 9 Observações | 10 Procedure | 11 Data Requerido | 12 Score
        return [
            item.get("datarececao", ""),
            item.get("datacolheita", ""),
            item.get("referencia", ""),
            item.get("hospedeiro", ""),
            item.get("tipo", ""),
            item.get("zona", ""),
            item.get("responsavelamostra", ""),
            item.get("responsavelcolheita", ""),
            item.get("observacoes", ""),
            item.get("procedure", ""),
            item.get("datarequerido", ""),
            item.get("Score", ""),
        ]
    return list(item)


# ----------------------------------------------------------------
#  Escrever resultado no TEMPLATE Excel
# ----------------------------------------------------------------
def write_to_template(ocr_rows_per_req, out_base_name, expected_count=None, source_pdf=None):
    """
    Escreve as requisições no TEMPLATE_PXf_SGSLABIP1056.xlsx,
    mantendo fórmulas/validações/formatos.

    Parâmetros:
      - ocr_rows_per_req: list[list[dict|list]]  → uma lista por requisição
      - out_base_name: str                       → base do nome do ficheiro de saída
      - expected_count: Optional[int]            → nº esperado de amostras (para alerta)
      - source_pdf: Optional[str]                → nome do PDF de origem
    """
    template_path = Path(os.environ.get("TEMPLATE_PATH", TEMPLATE_PATH))
    if not template_path.exists():
        raise FileNotFoundError(f"TEMPLATE não encontrado: {template_path}")

    sheet_name = "Avaliação pré registo"
    start_row = 6  # Mantém as 5 primeiras linhas do template

    out_files = []

    # Garante nome base limpo
    base = Path(out_base_name).stem

    for idx, req_rows in enumerate(ocr_rows_per_req, start=1):
        # Copia o template
        out_path = OUTPUT_DIR / f"{base}_req{idx}.xlsx"
        shutil.copy(template_path, out_path)

        wb = load_workbook(out_path)
        if sheet_name not in wb.sheetnames:
            wb.close()
            raise KeyError(f"Folha '{sheet_name}' não encontrada no template.")
        ws = wb[sheet_name]

        # 1) Metadados (não sobrescrever fórmulas de E1/F1)
        #    - G1:J1 → origem (PDF / req)
        #    - K1    → timestamp
        #    - L1    → resumo contagem
        origem_text = f"Origem: {source_pdf or base} | Req #{idx}"
        ws["G1"].value = origem_text
        ws["K1"].value = _now_str()
        # L1 com sumário da contagem
        try:
            count_rows = len(req_rows) if req_rows else 0
        except Exception:
            count_rows = 0
        ws["L1"].value = f"Amostras: {count_rows}" + (f" (esperado: {expected_count})" if expected_count else "")

        # 2) Escrita das linhas – respeitando validações do template
        row_idx = start_row
        for row in (req_rows or []):
            values = _to_list_row(row)
            for col, value in enumerate(values, start=1):
                ws.cell(row=row_idx, column=col).value = value
            row_idx += 1

        # 3) Validação opcional: se houver expected_count e mismatch, destacar E1/F1
        if isinstance(expected_count, int) and expected_count >= 0:
            if count_rows != expected_count:
                fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                for cell_addr in ("E1", "F1"):
                    ws[cell_addr].fill = fill
                # Adiciona comentário (sem apagar fórmulas)
                try:
                    ws["E1"].comment = Comment(
                        f"Atenção: processadas {count_rows} amostras; esperado {expected_count}.", "NeoLab"
                    )
                except Exception:
                    pass  # comentários podem falhar em versões antigas do Excel

        wb.save(out_path)
        wb.close()
        print(f"🟢 Gravado com validações: {out_path}")
        out_files.append(str(out_path))

    return out_files


# ───────────────────────────────────────────────────────────────
#  PARSER — deteção simples via regex (auxiliar / fallback)
# ───────────────────────────────────────────────────────────────
ONLY_DATE_RE = re.compile(r"^\d{1,2}/\d{1,2}/\d{4}$")
REF_SLASH_RE = re.compile(r"\b\d{1,4}\s*/\s*[A-Za-z0-9]{1,10}(?:\s*/\s*[A-Za-z0-9\-]{1,12})+\b", re.I)
REF_NUM_RE = re.compile(r"\b\d{7,8}\b")
TIPO_RE = re.compile(r"\b(Composta|Simples|Individual)\b", re.I)
NATUREZA_RE = re.compile(r"\bpartes?\s+de?\s*vegetais?\b|\bnatureza\s+da\s+amostra\b", re.I)
NATUREZA_KEYWORDS = [
    "ramos", "folhas", "ramosefolhas", "ramosc/folhas",
    "material", "materialherbalho", "materialherbário", "materialherbalo",
    "natureza", "insetos", "sementes", "solo"
]


def _clean_ref(raw: str) -> str:
    s = raw.strip()
    s = re.sub(r"[\s\t]+", "", s)
    s = re.sub(r"[|,;]+", "/", s)
    s = re.sub(r"[^\w/]", "", s)
    s = re.sub(r"/{2,}", "/", s)
    return s.upper()

def _is_natureza_line(s: str) -> bool:
    t = re.sub(r"\s+", " ", s.strip().lower())
    if NATUREZA_RE.search(t):
        return True
    return any(k in t.replace(" ", "") for k in NATUREZA_KEYWORDS)

def _merge_host(lines, j):
    def ok_host(txt):
        return (re.search(r"[A-Za-zÀ-ÿ]", txt)
                and not _is_natureza_line(txt)
                and not TIPO_RE.search(txt))
    a = lines[j].strip()
    if not ok_host(a):
        return "", 0
    if j + 1 < len(lines):
        b = lines[j+1].strip()
        if ok_host(b) and re.match(r"^[A-Za-zÀ-ÿ\.\-]+$", b):
            return f"{a} {b}".strip(), 2
    return a, 1


def parse_with_regex(text: str):
    """
    Extrai blocos de amostras simples via regex tolerante a OCR.
    Devolve lista de listas no formato do template.
    """
    text = re.sub(r"\s+", " ", text)
    padrao = re.compile(
        r"(?P<data_rec>\d{1,2}\s*[/\-]?\s*\d{1,2}\s*[/\-]?\s*\d{2,4}).*?"
        r"(?P<data_col>\d{1,2}\s*[/\-]?\s*\d{1,2}\s*[/\-]?\s*\d{2,4}).*?"
        r"(?P<codigo>\d{2,5}\/\d{4}\/[A-Z]{2,}|[0-9]{5,})?.*?"
        r"(?P<especie>[A-Z][a-zç]+(?: [a-zç]+){0,2}).*?"
        r"(?P<natureza>Simples|Composta).*?"
        r"(?P<zona>Isenta|Contida|Desconhec[ia]do|Zona [A-Za-z]+)?.*?"
        r"(?P<responsavel>DGAV|INIAV|INSA|Outros)?",
        re.S,
    )
    resultados = []
    for m in padrao.finditer(text):
        resultados.append([
            (m.group("data_rec") or "").replace(" ", ""),
            (m.group("data_col") or "").replace(" ", ""),
            m.group("codigo") or "",
            m.group("especie") or "",
            m.group("natureza") or "",
            m.group("zona") or "",
            m.group("responsavel") or "",
            "",  # responsável colheita (não preencher por agora)
            "",  # observações
            "XYLELLA",
            (m.group("data_rec") or "").replace(" ", ""),
            "",  # Score
        ])
    return resultados


# ----------------------------------------------------------------
#  SPLIT E CONTEXTO GLOBAL (robusto para múltiplas requisições)
# ----------------------------------------------------------------
def split_if_multiple_requisicoes(full_text: str):
    """
    Divide o texto OCR em blocos de 'requisição', usando marcadores
    típicos detetados nos formulários DGAV/SGS.
    """
    text = re.sub(r"[ \t]+", " ", full_text)
    text = re.sub(r"\n{2,}", "\n", text)
    pattern = re.compile(
        r"(?:(?:^|\n)\s*(?:PROGRAMA\s+DE\s+PROSPE|Amostra\s+colhida\s+por\s+DGAV|Refer[eê]ncia\s+da\s+amostra))",
        re.IGNORECASE
    )
    matches = list(pattern.finditer(text))
    if not matches:
        print("🔍 Nenhum marcador de nova requisição detetado.")
        return [text]

    positions = []
    last_pos = -9999
    for m in matches:
        if m.start() - last_pos > 1200:  # espaçamento mínimo para separar formulários no mesmo PDF
            positions.append(m.start())
            last_pos = m.start()

    if len(positions) == 1:
        print(f"🔍 Detetada 1 requisição (posições: {positions})")
        return [text]

    blocos = []
    for i, start in enumerate(positions):
        end = positions[i + 1] if i + 1 < len(positions) else len(text)
        blocos.append(text[start:end].strip())
    print(f"📄 Documento dividido em {len(blocos)} requisições distintas.")
    return blocos


def extract_context_from_text(full_text: str):
    """
    Extrai contexto global do bloco (zona, DGAV, datas de colheita/envio).
    """
    context = {}

    # Zona
    m_zona = re.search(r"Xylella\s+fastidiosa\s*\(([^)]+)\)", full_text, re.I)
    context["zona"] = (m_zona.group(1).strip() if m_zona else "Zona Isenta")

    # Responsável colheita (linha a seguir ao rótulo DGAV)
    responsavel = None
    m_base = re.search(
        r"Amostra.{0,40}?colhida.{0,15}?por.{0,10}?DGAV\s*[:\-]?",
        full_text, re.IGNORECASE | re.DOTALL,
    )
    if m_base:
        sub = full_text[m_base.end():]
        for ln in sub.strip().splitlines():
            ln = ln.strip()
            if not ln or re.search(r"Data|Refer", ln, re.I):
                break
            responsavel = re.sub(r"[\.:\-;,]+$", "", ln.strip())
            break
    context["responsavel_colheita"] = responsavel

    # Campo 'dgav' (quem assina/entidade)
    if responsavel and re.match(r"^DGAV\b", responsavel, re.I):
        context["dgav"] = responsavel
    elif responsavel:
        context["dgav"] = f"DGAV {responsavel}"
    else:
        m_dgav = re.search(r"DGAV\s+[A-ZÀ-ÿ\- ]{2,30}", full_text, re.I)
        context["dgav"] = (m_dgav.group(0).strip() if m_dgav else "DGAV")

    # Mapa de datas de colheita marcadas com (*), (**), etc.
    colheita_map = {}
    text_norm = re.sub(r"\s+(?:e|ou)\s+", " ", full_text)
    text_norm = text_norm.replace(",", " ")
    for m in re.finditer(r"(\d{1,2}/\d{1,2}/\d{4})\s*\(\s*(\*+)\s*\)", text_norm):
        mark = "(" + m.group(2).replace(" ", "") + ")"
        colheita_map[mark] = m.group(1)
    if not colheita_map:
        m_simple = re.search(r"Data\s+de\s+colheita.*?([\d/]{8,10})", full_text, re.I)
        if m_simple:
            d = m_simple.group(1)
            colheita_map["(*)"] = d
            colheita_map["(**)"] = d

    context["colheita_map"] = colheita_map
    context["default_colheita"] = next(iter(colheita_map.values()), "")

    # Data de envio
    m_envio = re.search(r"Data\s+(?:do|de)\s+envio.*?([\d/]{8,10})", full_text, re.I)
    context["data_envio"] = m_envio.group(1) if m_envio else context["default_colheita"] or datetime.now().strftime("%d/%m/%Y")

    print(f"🌍 Zona de origem: {context['zona']}")
    print(f"👤 Responsável DGAV: {context['dgav']}")
    print(f"👷 Responsável pela colheita: {context['responsavel_colheita'] or '(não identificado)'}")
    print(f"📅 Datas de colheita: {colheita_map or '(nenhuma)'} (padrão: {context['default_colheita'] or 'nenhuma'})")
    print(f"📬 Data do envio ao laboratório: {context['data_envio']}")
    return context


def parse_xylella_tables_from_text(full_text: str, context: dict, req_id=None):
    """
    Parser robusto: deteta várias amostras consecutivas no bloco.
    Considera cada linha com referência válida (xxxx/aaaa/LAB/...) como nova amostra.
    """
    out = []
    lines = [l.strip() for l in full_text.splitlines() if l.strip()]
    n = len(lines)
    for i, line in enumerate(lines):
        # referência válida (padrões SGS: 123/2025/LVT/1, 63020083, etc.)
        mref = re.search(r"\b(\d{1,4}/\d{4}/[A-Z]{2,4}/?\d*|\d{7,8})\b", line)
        if not mref:
            continue

        ref = _clean_ref(mref.group(1))
        hospedeiro, tipo = "", ""
        datacolheita = context.get("default_colheita", "")

        # procura hospedeiro nas 3 linhas seguintes
        for j in range(1, 4):
            if i + j >= n: break
            ln = lines[i + j]
            if TIPO_RE.search(ln):
                tipo = TIPO_RE.search(ln).group(1).capitalize()
            elif not hospedeiro and re.search(r"[A-Za-zÀ-ÿ]", ln) and not _is_natureza_line(ln):
                hospedeiro = re.sub(r"\s{2,}", " ", ln.strip())

        out.append({
            "requisicao_id": req_id,
            "datarececao": context.get("data_envio", ""),
            "datacolheita": datacolheita,
            "referencia": ref,
            "hospedeiro": hospedeiro,
            "tipo": tipo,
            "zona": context.get("zona", ""),
            "responsavelamostra": context.get("dgav", ""),
            "responsavelcolheita": "",
            "observacoes": "",
            "procedure": "XYLELLA",
            "datarequerido": context.get("data_envio", ""),
            "Score": ""
        })

    print(f"✅ {len(out)} amostras extraídas (req_id={req_id}) do texto OCR.")
    return out



# ----------------------------------------------------------------
#  Processamento do PDF (síncrono)
# ----------------------------------------------------------------
def process_pdf_sync(pdf_path: str):
    """
    Extrai texto de um PDF usando OCR Azure (se configurado) ou OCR local como fallback.
    Divide automaticamente por requisições e devolve:
        list_de_requisicoes -> cada uma é list[dict] (linhas já no formato do template)
    """
    import pytesseract
    from PIL import Image

    if pdf_to_images is None or extract_text_from_image_azure is None or get_analysis_result_azure is None:
        # tentativa de import tardio (se falhou no topo)
        from azure_ocr import pdf_to_images as _p2i, extract_text_from_image_azure as _ext, get_analysis_result_azure as _get
        globals()["pdf_to_images"] = _p2i
        globals()["extract_text_from_image_azure"] = _ext
        globals()["get_analysis_result_azure"] = _get

    pdf_path = str(pdf_path)
    text_total = ""

    # 1) PDF → imagens
    try:
        images = pdf_to_images(pdf_path)  # devolve lista de PIL.Image
        print(f"📄 PDF convertido em {len(images)} imagem(ns).")
    except Exception as e:
        raise RuntimeError(f"Falha ao converter PDF em imagens: {e}")

    # 2) OCR por página (Azure, com fallback para Tesseract local)
    for idx, img in enumerate(images, start=1):
        tmp_path = os.path.join(tempfile.gettempdir(), f"page_{idx}.png")
        img.save(tmp_path, "PNG")

        try:
            result = extract_text_from_image_azure(tmp_path)
            data = get_analysis_result_azure(result)

            # Extrair linhas de 'pages'
            pages = data.get("analyzeResult", {}).get("pages", [])
            if pages:
                for page in pages:
                    for ln in page.get("lines", []):
                        content = ln.get("content") or ln.get("text") or ""
                        if content:
                            text_total += content + "\n"
            else:
                # compatibilidade com 'readResult'
                read_blocks = data.get("analyzeResult", {}).get("readResult", [])
                for block in read_blocks:
                    for line in block.get("lines", []):
                        text_total += (line.get("text", "") or "") + "\n"

        except Exception as e:
            print(f"⚠️ Erro no OCR Azure (página {idx}: {e}) — a usar Tesseract local.")
            text_total += pytesseract.image_to_string(img) + "\n"

    if not text_total.strip():
        raise RuntimeError(f"Não foi possível extrair texto de {os.path.basename(pdf_path)}")

    # Guardar debug OCR bruto (útil)
    base_name = Path(pdf_path).stem
    debug_txt = OUTPUT_DIR / f"{base_name}_ocr_debug.txt"
    try:
        with open(debug_txt, "w", encoding="utf-8") as f:
            f.write(text_total)
        print(f"📝 Texto OCR bruto guardado em: {debug_txt}")
    except Exception:
        pass

    # 3) Split em requisições + parsing
    blocos = split_if_multiple_requisicoes(text_total)
    todas_reqs = []

    for i, bloco in enumerate(blocos, start=1):
        print(f"\n🔹 A processar requisição {i}/{len(blocos)}...")
        context = extract_context_from_text(bloco)
        amostras = parse_xylella_tables_from_text(bloco, context, req_id=i)
        if amostras:
            todas_reqs.append(amostras)
        else:
            print(f"⚠️ Requisição {i} ignorada — sem referências válidas.")

    if not todas_reqs:
        print("⚠️ Nenhuma requisição válida detetada.")

    return todas_reqs


# ----------------------------------------------------------------
#  Execução direta (opcional)
# ----------------------------------------------------------------
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Processador Xylella (OCR + Parser + Export para Template)")
    ap.add_argument("pdf", help="Caminho do PDF a processar")
    ap.add_argument("--expected", type=int, default=None, help="Nº de amostras esperado por requisição (para alerta E1/F1)")
    args = ap.parse_args()

    reqs = process_pdf_sync(args.pdf)
    base = Path(args.pdf).stem
    files = write_to_template(
        reqs,
        out_base_name=base,
        expected_count=args.expected,
        source_pdf=Path(args.pdf).name
    )

    print("\n📂 Saídas geradas:")
    for f in files:
        print("   -", f)


