# azure_ocr.py
# Cliente leve para Azure Computer Vision (OCR)
# Compatível com Streamlit Cloud (usa secrets automaticamente)

import os, requests

# Tenta ler as variáveis de ambiente (não falha se não existirem)
AZURE_KEY = os.environ.get("AZURE_KEY")
AZURE_ENDPOINT = os.environ.get("AZURE_ENDPOINT")

# Só avisa, não falha — permite fallback para OCR local
if not AZURE_KEY or not AZURE_ENDPOINT:
    print("⚠️ AVISO: AZURE_KEY ou AZURE_ENDPOINT não configurados — OCR Azure será ignorado.")
else:
    print(f"🔗 Azure endpoint ativo: {AZURE_ENDPOINT}")

# URL base do serviço OCR (Azure Cognitive Services)
READ_URL = f"{AZURE_ENDPOINT}/computervision/imageanalysis:analyze?api-version=2023-02-01-preview&features=read"

def pdf_to_images(pdf_path):
    """
    Converte um PDF em lista de imagens (PIL.Image) sem precisar de poppler.
    Usa PyMuPDF (fitz), compatível com Streamlit Cloud.
    """
    import fitz  # PyMuPDF
    from PIL import Image
    import io

    images = []
    with fitz.open(pdf_path) as doc:
        for page in doc:
            pix = page.get_pixmap(dpi=200)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            images.append(img)
    return images
    
def extract_text_from_image_azure(image_path: str):
    """Envia a imagem para o endpoint OCR da Azure e retorna o resultado JSON."""
    if not AZURE_KEY or not AZURE_ENDPOINT:
        raise RuntimeError("⚠️ OCR Azure não configurado. Usa OCR local.")
        
    headers = {
        "Ocp-Apim-Subscription-Key": AZURE_KEY,
        "Content-Type": "application/octet-stream"
    }
    with open(image_path, "rb") as f:
        img_data = f.read()

    response = requests.post(READ_URL, headers=headers, data=img_data)
    if response.status_code not in (200, 202):
        raise RuntimeError(f"Erro no envio OCR: {response.status_code} - {response.text}")

    return response.json()

def get_analysis_result_azure(result_json):
    """Normaliza o JSON devolvido pela Azure para o formato esperado."""
    if "analyzeResult" in result_json:
        return result_json
    return {"analyzeResult": result_json}

def extract_all_text(pdf_path):
    """
    Extrai todo o texto de um PDF usando OCR Azure (se configurado)
    ou OCR local como fallback.
    """
    from pdf2image import convert_from_path
    import pytesseract

    if not AZURE_KEY or not AZURE_ENDPOINT:
        print("⚠️ Azure OCR não configurado — a usar OCR local.")
        images = convert_from_path(pdf_path)
        text = ""
        for img in images:
            text += pytesseract.image_to_string(img, lang="por")
        return text

    # Azure configurado — envia página a página
    text_total = ""
    images = convert_from_path(pdf_path)
    for idx, img in enumerate(images, start=1):
        tmp_path = f"/tmp/page_{idx}.jpg"
        img.save(tmp_path, "JPEG")
        try:
            result = extract_text_from_image_azure(tmp_path)
            data = get_analysis_result_azure(result)
            for block in data.get("analyzeResult", {}).get("readResult", []):
                for line in block.get("lines", []):
                    text_total += line.get("text", "") + "\n"
        except Exception as e:
            print(f"⚠️ Erro OCR página {idx}: {e}")
    return text_total


