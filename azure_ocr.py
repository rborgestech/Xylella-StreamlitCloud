# azure_ocr.py
# Cliente leve para Azure Computer Vision (OCR)
# Usa a AZURE_KEY e AZURE_ENDPOINT definidas nos secrets do Streamlit

import os, time, requests
from dotenv import load_dotenv

# Carrega variáveis locais (se existir .env) ou cloud (Streamlit secrets)
load_dotenv()
AZURE_KEY = os.getenv("AZURE_KEY") or os.environ.get("AZURE_KEY")
AZURE_ENDPOINT = os.getenv("AZURE_ENDPOINT") or os.environ.get("AZURE_ENDPOINT")

if not AZURE_KEY or not AZURE_ENDPOINT:
    raise ValueError("❌ Falta configurar AZURE_KEY e AZURE_ENDPOINT nos secrets do Streamlit.")

# URL base do serviço OCR (Azure Cognitive Services)
READ_URL = f"{AZURE_ENDPOINT}/computervision/imageanalysis:analyze?api-version=2023-02-01-preview&features=read"

def extract_text_from_image_azure(image_path: str):
    """Envia a imagem para o endpoint OCR da Azure e retorna o URL de operação."""
    headers = {
        "Ocp-Apim-Subscription-Key": AZURE_KEY,
        "Content-Type": "application/octet-stream"
    }
    with open(image_path, "rb") as f:
        img_data = f.read()

    response = requests.post(READ_URL, headers=headers, data=img_data)
    if response.status_code not in (200, 202):
        raise RuntimeError(f"Erro no envio OCR: {response.status_code} - {response.text}")

    # Azure Vision devolve resultado direto (sem polling)
    return response.json()

def get_analysis_result_azure(result_json):
    """Normaliza o JSON devolvido pela Azure para o formato esperado."""
    # Para manter compatibilidade com core_xylella
    if "analyzeResult" in result_json:
        return result_json
    return {"analyzeResult": result_json}
