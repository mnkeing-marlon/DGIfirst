import anthropic
import base64
import streamlit as st
import json
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

def envoyer_email(model_name):
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
        return 1
    
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
        return 1
    
    return max(0.0, min(1.0, confiance))

from glmocr import parse
import json

def extract_once_glm_structured(api_key: str, image_path: str, custom_prompt: str) -> dict:
    """
    Extraction avec GLM-OCR - Gère les cellules fusionnées (colspan/rowspan)
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
        return {
            "titre": {"valeur": html_content[:200], "confiance": 0.5},
            "colonnes": [],
            "lignes": []
        }
    
    # === 1. Extraire le titre ===
    titre = ""
    first_row = table.find('tr')
    if first_row:
        first_cell = first_row.find('td') or first_row.find('th')
        if first_cell:
            titre = first_cell.get_text(strip=True)
    
    # === 2. Construire la matrice avec conservation des infos de fusion ===
    
    all_rows = table.find_all('tr')
    
    # D'abord, collecter toutes les cellules avec leurs attributs
    cells_info = []  # Liste de listes : (row, col, valeur, rowspan, colspan)
    
    for row_idx, tr in enumerate(all_rows):
        col_idx = 0
        for cell in tr.find_all(['td', 'th']):
            valeur = cell.get_text(strip=True)
            rowspan = int(cell.get('rowspan', 1))
            colspan = int(cell.get('colspan', 1))
            
            cells_info.append({
                "row": row_idx,
                "col": col_idx,
                "valeur": valeur,
                "rowspan": rowspan,
                "colspan": colspan
            })
            
            col_idx += colspan
    
    # Trouver le nombre max de colonnes
    max_cols = max([c["col"] + c["colspan"] for c in cells_info]) if cells_info else 0
    
    # Créer la matrice
    matrix = [[None for _ in range(max_cols)] for _ in range(len(all_rows))]
    rowspan_map = {}
    
    for cell_info in cells_info:
        row = cell_info["row"]
        col = cell_info["col"]
        valeur = cell_info["valeur"]
        rowspan = cell_info["rowspan"]
        colspan = cell_info["colspan"]
        
        # Remplir la matrice
        for r in range(rowspan):
            for c in range(colspan):
                if row + r < len(matrix) and col + c < max_cols:
                    # Ne pas écraser si déjà rempli par un rowspan précédent
                    if matrix[row + r][col + c] is None:
                        matrix[row + r][col + c] = valeur
    
    # === 3. Extraire les noms des colonnes ===
    colonnes = []
    if matrix and len(matrix) > 0:
        header_row = matrix[0]
        for idx, valeur in enumerate(header_row):
            if valeur and valeur != titre:
                colonnes.append({
                    "nom": valeur,
                    "confiance": 0.94
                })
    
    # === 4. Extraire les lignes avec les infos de fusion ===
    lignes = []
    start_idx = 1  # Ignorer l'en-tête
    
    for row_idx in range(start_idx, len(matrix)):
        row = matrix[row_idx]
        ligne_data = {}
        
        # Trouver les infos de fusion pour cette ligne
        fusions_ligne = [c for c in cells_info if c["row"] <= row_idx < c["row"] + c["rowspan"]]
        
        for col_idx, valeur in enumerate(row):
            if col_idx < len(colonnes):
                nom_colonne = colonnes[col_idx]["nom"]
                
                # Nettoyer la valeur : ignorer les chiffres parasites dans les cellules vides
                if valeur and len(str(valeur)) < 3 and str(valeur).isdigit():
                    # Vérifier si c'est un vrai nombre ou un parasite
                    # Si la cellule voisine est vide, probablement parasite
                    voisine_gauche = row[col_idx - 1] if col_idx > 0 else None
                    if voisine_gauche is None or voisine_gauche == "":
                        valeur = ""  # Ignorer le parasite
                
                # Récupérer rowspan et colspan depuis les infos de fusion
                rowspan_val = 1
                colspan_val = 1
                for f in fusions_ligne:
                    if f["col"] <= col_idx < f["col"] + f["colspan"]:
                        rowspan_val = f["rowspan"]
                        colspan_val = f["colspan"]
                        break
                
                if valeur:  # Ne créer la cellule que si elle a une valeur
                    confiance = calculer_confiance_cellule(valeur, None)
                    ligne_data[nom_colonne] = {
                        "valeur": valeur,
                        "confiance": confiance,
                        "rowspan": rowspan_val,
                        "colspan": colspan_val
                    }
        
        if ligne_data:
            lignes.append(ligne_data)
    
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
    
    elif engine == "gemini":
        # Extraction avec Gemini
        extraction_a = extract_once_gemini(api_key, image_path, PROMPT_A)
        extraction_b = extract_once_gemini(api_key, image_path, PROMPT_B)
    
    else:  # engine == "claude"
        # Extraction avec Claude
        extraction_a = extract_once_claude(api_key, image_path, PROMPT_A)
        extraction_b = extract_once_claude(api_key, image_path, PROMPT_B)
    
    return extraction_a, extraction_b
