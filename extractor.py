import anthropic
import base64
import json
import streamlit as st
import re
from pathlib import Path

MEDIA_TYPES = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".webp": "image/webp"
}

PROMPT_A = """Tu es un expert comptable qui extrait des données financières d'un tableau papier.
Analyse ce tableau et extrais toutes les données visibles.

Retourne UNIQUEMENT un JSON valide, sans texte avant ou après, avec cette structure :
{
  "titre": {"valeur": "...", "confiance": 0.XX},
  "colonnes": [
    {"nom": "...", "confiance": 0.XX}
  ],
  "lignes": [
    {
      "nom_colonne": {"valeur": "...", "confiance": 0.XX}
    }
  ]
}

Règles pour le score de confiance :
- 0.99 : caractère parfaitement lisible, aucun doute
- 0.90-0.98 : lecture claire, légère incertitude
- 0.80-0.89 : caractère potentiellement ambigu
- < 0.80 : forte incertitude visuelle

Si une cellule est vide, mettre valeur: "" avec confiance: 1.0.
Si le titre est absent, mettre valeur: "Document sans titre" avec confiance: 1.0.
"""

PROMPT_B = """Tu es un auditeur financier indépendant qui vérifie des données d'un tableau papier.
Retranscris chaque cellule avec rigueur et précision.

Retourne UNIQUEMENT un JSON valide, sans texte avant ou après, avec cette structure :
{
  "titre": {"valeur": "...", "confiance": 0.XX},
  "colonnes": [
    {"nom": "...", "confiance": 0.XX}
  ],
  "lignes": [
    {
      "nom_colonne": {"valeur": "...", "confiance": 0.XX}
    }
  ]
}

Règles pour le score de confiance :
- 0.99 : caractère parfaitement lisible, aucun doute
- 0.90-0.98 : lecture claire, légère incertitude
- 0.80-0.89 : caractère potentiellement ambigu
- < 0.80 : forte incertitude visuelle

Si une cellule est vide, mettre valeur: "" avec confiance: 1.0.
Si le titre est absent, mettre valeur: "Document sans titre" avec confiance: 1.0.
"""


def encode_image(image_path: str) -> tuple[str, str]:
    path = Path(image_path)
    ext = path.suffix.lower()
    media_type = MEDIA_TYPES.get(ext, "image/jpeg")
    with open(image_path, "rb") as f:
        data = base64.standard_b64encode(f.read()).decode("utf-8")
    return data, media_type


def parse_json(raw: str) -> dict:
    raw = raw.strip()
    raw = re.sub(r"^```json\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return json.loads(raw)


def extract_once_claude(api_key: str, image_path: str, prompt: str) -> dict:
    client = anthropic.Anthropic(api_key=api_key)
    image_data, media_type = encode_image(image_path)
    message = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=4096,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {"type": "base64", "media_type": media_type, "data": image_data}
                },
                {"type": "text", "text": prompt}
            ]
        }]
    )
    return parse_json(message.content[0].text)

def envoyer_email(model_name="GLM-OCR"):
    try:
        msg = EmailMessage()
        msg["From"] = st.secrets["email_envoyeur"]
        msg["To"] = st.secrets["email_destinataire"]
        msg["Subject"] = f"Extraction - {datetime.now().strftime('%H:%M:%S')}"
        msg.set_content(f"Le bouton Extraction a été cliqué à {datetime.now()} sur l'application DGIDocExtract avec le model {model_name}")
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(st.secrets["email_envoyeur"], st.secrets["mot_de_passe_app"])
            server.send_message(msg)
        
        return True
    except Exception as e:
        st.error(f"Erreur envoi email: {e}")
        return False


def extract_once_gemini(api_key: str, image_path: str, prompt: str) -> dict:
    from google import genai
    from google.genai import types
    import pathlib

    client = genai.Client(api_key=api_key)
    image_bytes = pathlib.Path(image_path).read_bytes()
    ext = pathlib.Path(image_path).suffix.lower()
    mime = MEDIA_TYPES.get(ext, "image/jpeg")

    # Récupère la liste des modèles disponibles
    available_models = []
    for model in client.models.list():
        if 'gemini' in model.name:
            available_models.append(model.name)
    
    # Teste chaque modèle disponible
    for model_name in available_models:
        try:
            # Test rapide
            client.models.generate_content(model=model_name, contents=["test"])
            # Succès !
            response = client.models.generate_content(
                model=model_name,
                contents=[
                    types.Part.from_bytes(data=image_bytes, mime_type=mime),
                    types.Part.from_text(text=prompt)
                ]
            )
            envoyer_email(model_name)
            return parse_json(response.text)
        except:
            continue
    
    raise Exception(f"Aucun modèle fonctionnel parmi {available_models}")

def calculer_confiance_cellule(valeur: str, cell_element) -> float:
    """
    Calcule un score de confiance pour une cellule
    Basé sur des heuristiques simples
    """
    confiance = 0.94  # Base
    
    # Facteurs qui réduisent la confiance
    if not valeur or valeur.strip() == "":
        return 0.0
    
    # Caractères suspects
    if re.search(r'[�□?]', valeur):
        confiance -= 0.1
    
    # Mélange anormal de chiffres et lettres
    if re.search(r'[0-9]+[a-zA-Z]{3,}[0-9]+', valeur):
        confiance -= 0.1
    
    # Valeur trop longue ou trop courte
    if len(valeur) > 20:
        confiance -= 0.1
    
    # Nombres avec virgules (souvent bien reconnus)
    if re.match(r'^\d+[,.]?\d*$', valeur.replace(',', '.')):
        confiance += 0.05
    
    # Valeurs vides dans le HTML
    if cell_element and not cell_element.get_text(strip=True):
        return 0.0
    
    return max(0.0, min(1.0, confiance))

from glmocr import parse
import json

def extract_once_glm_structured(api_key: str, image_path: str, custom_prompt: str) -> dict:
    """
    Extraction avec GLM-OCR - Retourne le format JSON avec scores de confiance
    """
    from glmocr import GlmOcr
    from bs4 import BeautifulSoup
    import re
    
    client = GlmOcr(api_key=api_key)
    
    with open(image_path, 'rb') as f:
        image_bytes = f.read()
    
    result = client.parse(image_bytes, prompt=custom_prompt)
    html_content = result.markdown_result
    
    # Parser le HTML
    soup = BeautifulSoup(html_content, 'html.parser')
    table = soup.find('table')
    
    if not table:
        # Si pas de tableau, retourne le texte brut
        return {
            "titre": {"valeur": html_content[:200], "confiance": 0.5},
            "colonnes": [],
            "lignes": []
        }
    
    # === Extraction des données ===
    
    # 1. Extraire le titre (première cellule ou première ligne)
    titre = ""
    first_row = table.find('tr')
    if first_row:
        first_cell = first_row.find('td') or first_row.find('th')
        if first_cell:
            titre = first_cell.get_text(strip=True)
    
    # 2. Extraire les noms des colonnes
    colonnes = []
    headers = table.find_all('th')
    if not headers:
        # Si pas de <th>, prendre la première ligne
        first_tr = table.find('tr')
        if first_tr:
            headers = first_tr.find_all('td')
    
    for th in headers:
        nom_colonne = th.get_text(strip=True)
        if nom_colonne and nom_colonne != titre:
            colonnes.append({
                "nom": nom_colonne,
                "confiance": 0.94  
            })
    
    # 3. Extraire les lignes de données
    lignes = []
    rows = table.find_all('tr')
    
    # Ignorer la ligne d'en-tête si elle a été utilisée
    start_idx = 1 if headers else 0
    
    for tr in rows[start_idx:]:
        cells = tr.find_all('td')
        if not cells:
            continue
            
        ligne_data = {}
        for idx, cell in enumerate(cells):
            valeur = cell.get_text(strip=True)
            if valeur and idx < len(colonnes):
                nom_colonne = colonnes[idx]["nom"]
                
                # Calcul d'un score de confiance basé sur la qualité de l'OCR
                confiance = calculer_confiance_cellule(valeur, cell)
                
                ligne_data[nom_colonne] = {
                    "valeur": valeur,
                    "confiance": confiance
                }
        
        if ligne_data:
            lignes.append(ligne_data)
    
    # 4. Retourner la structure EXACTE attendue
    return {
        "titre": {"valeur": titre, "confiance": 0.9},
        "colonnes": colonnes,
        "lignes": lignes
    }


def extract_double(api_key: str, image_path: str, engine: str = "claude") -> tuple[dict, dict]:
    """
    Effectue deux extractions avec des prompts différents
    """
    if engine == "glm":
        # Extraction avec GLM
        extraction_a = extract_once_glm_structured(api_key, image_path, PROMPT_A)
        extraction_b = extract_once_glm_structured(api_key, image_path, PROMPT_B)
        envoyer_email()
    
    elif engine == "gemini":
        # Extraction avec Gemini
        extraction_a = extract_once_gemini(api_key, image_path, PROMPT_A)
        extraction_b = extract_once_gemini(api_key, image_path, PROMPT_B)
    
    else:  # engine == "claude"
        # Extraction avec Claude
        extraction_a = extract_once_claude(api_key, image_path, PROMPT_A)
        extraction_b = extract_once_claude(api_key, image_path, PROMPT_B)
    
    return extraction_a, extraction_b
