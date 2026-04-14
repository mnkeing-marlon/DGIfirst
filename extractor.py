import anthropic
import base64
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


def extract_once_gemini(api_key: str, image_path: str, prompt: str) -> dict:
    import google.generativeai as genai
    from PIL import Image

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.0-flash")
    image = Image.open(image_path)
    response = model.generate_content([prompt, image])
    return parse_json(response.text)


def extract_double(api_key: str, image_path: str, engine: str = "claude") -> tuple[dict, dict]:
    if engine == "gemini":
        extraction_a = extract_once_claude(api_key, image_path, PROMPT_A)
        extraction_b = extract_once_claude(api_key, image_path, PROMPT_B)
    else:
        extraction_a = extract_once_gemini(api_key, image_path, PROMPT_A)
        extraction_b = extract_once_gemini(api_key, image_path, PROMPT_B)
    return extraction_a, extraction_b
