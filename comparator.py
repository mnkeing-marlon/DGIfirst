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
    """Normalise une valeur pour la comparaison"""
    return str(valeur).strip().lower().replace(" ", "").replace(",", ".")


def comparer_cellule(valeur_a: str, confiance_a: float,
                     valeur_b: str, confiance_b: float,
                     seuil: float) -> dict:
    """Compare deux cellules et retourne leur statut"""
    
    # Normaliser les valeurs
    val_a_norm = normaliser(valeur_a)
    val_b_norm = normaliser(valeur_b)
    
    # ✅ Cas 1 : Les DEUX sont vides → VERT (rien à signaler)
    if val_a_norm == "" and val_b_norm == "":
        return {
            "valeur": valeur_a,
            "valeur_b": valeur_b,
            "confiance_a": confiance_a,
            "confiance_b": confiance_b,
            "identiques": True,
            "statut": "vert"
        }
    
    # ✅ Cas 2 : L'un est vide, l'autre non → ROUGE (alerte)
    if (val_a_norm == "" and val_b_norm != "") or (val_a_norm != "" and val_b_norm == ""):
        return {
            "valeur": valeur_a,
            "valeur_b": valeur_b,
            "confiance_a": confiance_a,
            "confiance_b": confiance_b,
            "identiques": False,
            "statut": "rouge"
        }
    
    # ✅ Cas 3 : Les deux ont des valeurs → comparaison normale
    identiques = val_a_norm == val_b_norm
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
    """Compare les titres des deux extractions"""
    titre_a = ext_a.get("titre", {})
    titre_b = ext_b.get("titre", {})
    
    return comparer_cellule(
        titre_a.get("valeur", ""),
        titre_a.get("confiance", 0.0),
        titre_b.get("valeur", ""),
        titre_b.get("confiance", 0.0),
        SEUIL_DONNEES
    )


def comparer_lignes(ext_a: dict, ext_b: dict, colonnes: list) -> list:
    """
    Compare les lignes des deux extractions en gérant les cellules fusionnées
    """
    lignes_a = ext_a.get("lignes", [])
    lignes_b = ext_b.get("lignes", [])
    
    # Récupérer les noms de colonnes (peuvent être dupliqués si fusion)
    noms_colonnes = []
    for col in colonnes:
        nom = col.get("valeur", col.get("nom", ""))
        noms_colonnes.append(nom)
    
    # Si des colonnes ont le même nom (fusion), on les différencie
    compteur_colonnes = {}
    noms_colonnes_uniques = []
    for nom in noms_colonnes:
        if nom in compteur_colonnes:
            compteur_colonnes[nom] += 1
            nom_unique = f"{nom}_{compteur_colonnes[nom]}"
        else:
            compteur_colonnes[nom] = 1
            nom_unique = nom
        noms_colonnes_uniques.append(nom_unique)
    
    n = max(len(lignes_a), len(lignes_b))
    resultats = []

    for i in range(n):
        ligne_a = lignes_a[i] if i < len(lignes_a) else {}
        ligne_b = lignes_b[i] if i < len(lignes_b) else {}
        ligne_result = {}

        for idx, nom_col in enumerate(noms_colonnes_uniques):
            # Nom original sans suffixe
            nom_original = noms_colonnes[idx]
            
            # Gérer les cellules fusionnées (peuvent être absentes)
            cell_a = ligne_a.get(nom_original, {})
            cell_b = ligne_b.get(nom_original, {})
            
            # Si la cellule est manquante à cause d'une fusion, propager la valeur
            if not cell_a and idx > 0:
                # Chercher la valeur dans la cellule précédente (fusion à gauche)
                prev_col = noms_colonnes_uniques[idx - 1]
                if prev_col in ligne_result:
                    cell_a = ligne_result[prev_col]
                    cell_a = {"valeur": cell_a.get("valeur", ""), 
                              "confiance": cell_a.get("confiance_a", 0.0)}
            
            if not isinstance(cell_a, dict):
                cell_a = {"valeur": str(cell_a), "confiance": 0.0}
            if not isinstance(cell_b, dict):
                cell_b = {"valeur": str(cell_b), "confiance": 0.0}

            # ✅ SUPPRIMÉ : la vérification des cellules vides inattendues
            # Plus de transformation des cellules vides en orange

            resultat = comparer_cellule(
                cell_a.get("valeur", ""),
                cell_a.get("confiance", 0.0),
                cell_b.get("valeur", ""),
                cell_b.get("confiance", 0.0),
                SEUIL_DONNEES
            )

            # ❌ LIGNES SUPPRIMÉES :
            # if vide_inattendue and resultat["statut"] == "vert":
            #     resultat["statut"] = "orange"

            ligne_result[nom_col] = resultat

        resultats.append(ligne_result)

    return resultats


def aligner_colonnes_avec_fusion(ext_a: dict, ext_b: dict) -> list:
    """
    Aligne les colonnes en tenant compte des fusions
    """
    cols_a = ext_a.get("colonnes", [])
    cols_b = ext_b.get("colonnes", [])
    
    # Extraire les noms
    noms_a = [c.get("nom", c.get("valeur", "")) for c in cols_a]
    noms_b = [c.get("nom", c.get("valeur", "")) for c in cols_b]
    
    # ✅ Détecter les fusions (colonnes avec mêmes noms)
    from collections import Counter
    counter_a = Counter(noms_a)
    counter_b = Counter(noms_b)
    
    # Créer une liste alignée
    colonnes_alignees = []
    i, j = 0, 0
    
    while i < len(cols_a) or j < len(cols_b):
        if i < len(cols_a) and j < len(cols_b):
            nom_a = noms_a[i]
            nom_b = noms_b[j]
            
            if nom_a == nom_b:
                # Même colonne
                colonnes_alignees.append({
                    "nom": nom_a,
                    "confiance_a": cols_a[i].get("confiance", 0.0),
                    "confiance_b": cols_b[j].get("confiance", 0.0),
                    "fusion": counter_a[nom_a] > 1 or counter_b[nom_b] > 1
                })
                i += 1
                j += 1
            elif counter_a.get(nom_a, 0) > 1:
                # Fusion dans A
                colonnes_alignees.append({
                    "nom": f"{nom_a}_part{len([c for c in colonnes_alignees if c['nom'].startswith(nom_a)]) + 1}",
                    "confiance_a": cols_a[i].get("confiance", 0.0),
                    "confiance_b": 0.0,
                    "fusion": True
                })
                i += 1
            else:
                # Colonne normale
                colonnes_alignees.append({
                    "nom": nom_a,
                    "confiance_a": cols_a[i].get("confiance", 0.0),
                    "confiance_b": 0.0,
                    "fusion": False
                })
                i += 1
        elif i < len(cols_a):
            nom_a = noms_a[i]
            colonnes_alignees.append({
                "nom": nom_a,
                "confiance_a": cols_a[i].get("confiance", 0.0),
                "confiance_b": 0.0,
                "fusion": counter_a.get(nom_a, 0) > 1
            })
            i += 1
        else:
            nom_b = noms_b[j]
            colonnes_alignees.append({
                "nom": nom_b,
                "confiance_a": 0.0,
                "confiance_b": cols_b[j].get("confiance", 0.0),
                "fusion": counter_b.get(nom_b, 0) > 1
            })
            j += 1
    
    return colonnes_alignees


def comparer_colonnes(ext_a: dict, ext_b: dict) -> list:
    """
    Compare les colonnes en gérant les fusions
    """
    cols_a = ext_a.get("colonnes", [])
    cols_b = ext_b.get("colonnes", [])
    
    # ✅ Utiliser l'alignement intelligent
    cols_alignees = aligner_colonnes_avec_fusion(ext_a, ext_b)
    resultats = []

    for col in cols_alignees:
        resultat = comparer_cellule(
            col.get("nom", ""),
            col.get("confiance_a", 0.0),
            col.get("nom", ""),
            col.get("confiance_b", 0.0),
            SEUIL_ENTETES
        )
        # Ajouter l'information de fusion
        resultat["fusion"] = col.get("fusion", False)
        resultats.append(resultat)

    return resultats


def generer_rapport(titre: dict, colonnes: list, lignes: list) -> dict:
    """Génère un rapport statistique global"""
    toutes_cellules = [titre] + colonnes
    
    for ligne in lignes:
        toutes_cellules.extend(ligne.values())

    total = len(toutes_cellules)
    rouges = sum(1 for c in toutes_cellules if c.get("statut") == "rouge")
    oranges = sum(1 for c in toutes_cellules if c.get("statut") == "orange")
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
    """
    Compare deux extractions et retourne un rapport détaillé
    
    Args:
        extraction_a: Premier dictionnaire d'extraction
        extraction_b: Second dictionnaire d'extraction
    
    Returns:
        Dictionnaire contenant les comparaisons et le rapport
    """
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
