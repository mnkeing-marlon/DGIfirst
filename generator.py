"""
Phase 3 — Générateur Excel + Rapport TXT
"""

from datetime import datetime
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

COULEURS = {
    "vert":   {"bg": "C6EFCE", "font": "276221"},
    "orange": {"bg": "FFEB9C", "font": "9C6500"},
    "rouge":  {"bg": "FFC7CE", "font": "9C0006"},
    "header": {"bg": "2F4F8F", "font": "FFFFFF"},
    "meta":   {"bg": "D9E1F2", "font": "1F3864"},
    "legende":{"bg": "F2F2F2", "font": "000000"},
}

def appliquer_style(cell, statut: str = None, header: bool = False, meta: bool = False):
    if header:
        key = "header"
    elif meta:
        key = "meta"
    else:
        key = statut or "vert"

    couleur = COULEURS.get(key, COULEURS["vert"])
    cell.fill = PatternFill("solid", start_color=couleur["bg"])
    cell.font = Font(color=couleur["font"], bold=header or meta, name="Arial", size=10)
    cell.alignment = Alignment(horizontal="center", vertical="center")


def generer_excel(resultat: dict, image_path: str) -> Workbook:
    wb = Workbook()
    ws = wb.active

    titre_doc = resultat["titre"]["valeur"] or "Extraction"
    ws.title = titre_doc[:31]

    colonnes = resultat["colonnes"]
    lignes = resultat["lignes"]
    noms_colonnes = [col["valeur"] for col in colonnes]
    n_cols = len(noms_colonnes)

    ORANGE = PatternFill("solid", start_color="FFEB9C")
    ROUGE  = PatternFill("solid", start_color="FFC7CE")
    NORMAL = Font(name="Arial", size=10)
    BOLD   = Font(name="Arial", size=10, bold=True)

    def style(cell, statut=None, bold=False):
        cell.font = BOLD if bold else NORMAL
        if statut == "orange":
            cell.fill = ORANGE
        elif statut == "rouge":
            cell.fill = ROUGE

    # --- Ligne 1 : en-têtes de colonnes ---
    for j, col in enumerate(colonnes):
        cell = ws.cell(row=1, column=j + 1)
        cell.value = col["valeur"]
        style(cell, statut=col["statut"], bold=True)
    ws.row_dimensions[1].height = 22

    # --- Lignes de données ---
    for i, ligne in enumerate(lignes):
        for j, nom_col in enumerate(noms_colonnes):
            cell = ws.cell(row=i + 2, column=j + 1)
            cellule = ligne.get(nom_col, {"valeur": "", "statut": "vert"})
            val = cellule["valeur"]
            try:
                val = int(val.replace(' ', '').replace(',', '.')) if '.' not in val.replace(',', '.') else float(val.replace(' ', '').replace(',', '.'))
            except (ValueError, AttributeError):
                pass
            cell.value = val
            style(cell, statut=cellule["statut"])
        ws.row_dimensions[i + 2].height = 18

    # --- Largeur des colonnes ---
    for j in range(1, n_cols + 1):
        col_letter = ws.cell(row=1, column=j).column_letter
        max_len = max(
            (len(str(ws.cell(row=i, column=j).value or "")) for i in range(1, len(lignes) + 3)),
            default=10
        )
        ws.column_dimensions[col_letter].width = min(max_len + 4, 40)

    return wb


def generer_rapport_txt(resultat: dict, image_path: str) -> str:
    rapport = resultat["rapport"]
    lignes = resultat["lignes"]
    colonnes = resultat["colonnes"]
    noms_colonnes = [col["valeur"] for col in colonnes]
    now = datetime.now().strftime("%d/%m/%Y %H:%M")

    lignes_txt = [
        "=" * 60,
        "RAPPORT DE CONTRÔLE QUALITÉ",
        "=" * 60,
        f"Document  : {Path(image_path).name}",
        f"Titre     : {resultat['titre']['valeur']}",
        f"Date      : {now}",
        "",
        "RÉSUMÉ",
        "-" * 40,
        f"Score global     : {rapport['score_global']}%",
        f"Cellules vertes  : {rapport['verts']}",
        f"Cellules oranges : {rapport['oranges']}",
        f"Cellules rouges  : {rapport['rouges']}",
        f"Total cellules   : {rapport['total_cellules']}",
        "",
    ]

    # Titre
    t = resultat["titre"]
    if t["statut"] != "vert":
        lignes_txt += [
            "ALERTES — TITRE",
            "-" * 40,
            f"  Statut   : {t['statut'].upper()}",
            f"  Valeur A : {t['valeur']}",
            f"  Valeur B : {t['valeur_b']}",
            f"  Confiance: A={t['confiance_a']} | B={t['confiance_b']}",
            "",
        ]

    # En-têtes
    alertes_entetes = [col for col in colonnes if col["statut"] != "vert"]
    if alertes_entetes:
        lignes_txt += ["ALERTES — EN-TÊTES DE COLONNES", "-" * 40]
        for col in alertes_entetes:
            lignes_txt += [
                f"  Colonne  : {col['valeur']}",
                f"  Statut   : {col['statut'].upper()}",
                f"  Valeur B : {col['valeur_b']}",
                f"  Confiance: A={col['confiance_a']} | B={col['confiance_b']}",
                "",
            ]

    # Données
    alertes_data = []
    for i, ligne in enumerate(lignes):
        for nom_col in noms_colonnes:
            cell = ligne.get(nom_col, {"statut": "vert"})
            if cell["statut"] != "vert":
                alertes_data.append((i + 1, nom_col, cell))

    if alertes_data:
        lignes_txt += ["ALERTES — DONNÉES", "-" * 40]
        for num_ligne, nom_col, cell in alertes_data:
            lignes_txt += [
                f"  Ligne {num_ligne} | {nom_col}",
                f"  Statut   : {cell['statut'].upper()}",
                f"  Valeur A : {cell['valeur']}",
                f"  Valeur B : {cell['valeur_b']}",
                f"  Confiance: A={cell['confiance_a']} | B={cell['confiance_b']}",
                "",
            ]

    if not alertes_entetes and not alertes_data and t["statut"] == "vert":
        lignes_txt.append("✅ Aucune alerte — toutes les cellules sont fiables.")

    lignes_txt += ["=" * 60]
    return "\n".join(lignes_txt)


def sauvegarder(resultat: dict, image_path: str, dossier_sortie: str = "."):
    base = Path(dossier_sortie) / Path(image_path).stem
    wb = generer_excel(resultat, image_path)
    path_excel = f"{base}_extraction.xlsx"
    path_txt = f"{base}_rapport.txt"
    wb.save(path_excel)
    rapport_txt = generer_rapport_txt(resultat, image_path)
    with open(path_txt, "w", encoding="utf-8") as f:
        f.write(rapport_txt)
    return path_excel, path_txt
