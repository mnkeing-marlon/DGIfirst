"""
Phase 2 — Moteur de contrôle qualité
Compare les deux extractions et attribue un statut à chaque cellule.

Statuts :
  vert   : confiance OK + extractions identiques
  orange : confiance faible (mais identiques)
  rouge  : divergence entre A et B (priorité absolue)
"""

SEUIL_DONNEES = 0.90
SEUIL_ENTETES = 0.93


def normaliser(valeur: str) -> str:
    return str(valeur).strip().lower().replace(" ", "").replace(",", ".")


def comparer_cellule(valeur_a: str, confiance_a: float,
                     valeur_b: str, confiance_b: float,
                     seuil: float) -> dict:

    identiques = normaliser(valeur_a) == normaliser(valeur_b)
    confiance_min = min(confiance_a, confiance_b)

    if not identiques:
        statut = "rouge"
    elif confiance_min < seuil:
        statut = "orange"
    else:
        statut = "vert"

    return {
        "valeur": valeur_a,
        "valeur_b": valeur_b,
        "confiance_a": confiance_a,
        "confiance_b": confiance_b,
        "identiques": identiques,
        "statut": statut
    }


def comparer_titre(ext_a: dict, ext_b: dict) -> dict:
    titre_a = ext_a.get("titre", {})
    titre_b = ext_b.get("titre", {})
    return comparer_cellule(
        titre_a.get("valeur", ""),
        titre_a.get("confiance", 0.0),
        titre_b.get("valeur", ""),
        titre_b.get("confiance", 0.0),
        SEUIL_DONNEES
    )


def comparer_colonnes(ext_a: dict, ext_b: dict) -> list:
    cols_a = ext_a.get("colonnes", [])
    cols_b = ext_b.get("colonnes", [])
    n = max(len(cols_a), len(cols_b))
    resultats = []

    for i in range(n):
        col_a = cols_a[i] if i < len(cols_a) else {"nom": "", "confiance": 0.0}
        col_b = cols_b[i] if i < len(cols_b) else {"nom": "", "confiance": 0.0}
        resultat = comparer_cellule(
            col_a.get("nom", ""),
            col_a.get("confiance", 0.0),
            col_b.get("nom", ""),
            col_b.get("confiance", 0.0),
            SEUIL_ENTETES
        )
        resultats.append(resultat)

    return resultats


def comparer_lignes(ext_a: dict, ext_b: dict, colonnes: list) -> list:
    lignes_a = ext_a.get("lignes", [])
    lignes_b = ext_b.get("lignes", [])
    n = max(len(lignes_a), len(lignes_b))
    noms_colonnes = [col["valeur"] for col in colonnes]
    resultats = []

    for i in range(n):
        ligne_a = lignes_a[i] if i < len(lignes_a) else {}
        ligne_b = lignes_b[i] if i < len(lignes_b) else {}
        ligne_result = {}

        for nom_col in noms_colonnes:
            cell_a = ligne_a.get(nom_col, {"valeur": "", "confiance": 0.0})
            cell_b = ligne_b.get(nom_col, {"valeur": "", "confiance": 0.0})

            # Cellule vide inattendue dans une colonne non vide
            valeurs_col = [l.get(nom_col, {}).get("valeur", "") for l in lignes_a]
            col_a_non_vide = any(v != "" for v in valeurs_col)
            vide_inattendue = col_a_non_vide and cell_a.get("valeur", "") == ""

            resultat = comparer_cellule(
                cell_a.get("valeur", ""),
                cell_a.get("confiance", 0.0),
                cell_b.get("valeur", ""),
                cell_b.get("confiance", 0.0),
                SEUIL_DONNEES
            )

            confiance_vide = cell_a.get("confiance", 0.0)
            if vide_inattendue and confiance_vide < 1.0 and resultat["statut"] == "vert":
                resultat["statut"] = "orange"

            ligne_result[nom_col] = resultat

        resultats.append(ligne_result)

    return resultats


def generer_rapport(titre: dict, colonnes: list, lignes: list) -> dict:
    toutes_cellules = [titre] + colonnes
    for ligne in lignes:
        toutes_cellules.extend(ligne.values())

    total = len(toutes_cellules)
    rouges = sum(1 for c in toutes_cellules if c["statut"] == "rouge")
    oranges = sum(1 for c in toutes_cellules if c["statut"] == "orange")
    verts = total - rouges - oranges
    score = round((verts / total) * 100, 1) if total > 0 else 0

    return {
        "total_cellules": total,
        "verts": verts,
        "oranges": oranges,
        "rouges": rouges,
        "score_global": score
    }


def comparer(extraction_a: dict, extraction_b: dict) -> dict:
    titre = comparer_titre(extraction_a, extraction_b)
    colonnes = comparer_colonnes(extraction_a, extraction_b)
    lignes = comparer_lignes(extraction_a, extraction_b, colonnes)
    rapport = generer_rapport(titre, colonnes, lignes)

    return {
        "titre": titre,
        "colonnes": colonnes,
        "lignes": lignes,
        "rapport": rapport
    }
