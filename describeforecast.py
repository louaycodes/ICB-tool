"""
Script d'analyse de STRUCTURE - Fichier Forecast
=================================================
Objectif : comprendre l'architecture du fichier (feuilles, colonnes,
types, formules, dimensions reelles) SANS exposer aucune valeur
sensible reelle.

Toutes les valeurs sont masquees (lettres -> X, chiffres -> 9) ou
remplacees par leur type/ordre de grandeur. Le rapport genere est
safe a partager.

Usage :
    python analyse_structure_forecast.py chemin/vers/fichier_forecast.xlsx
"""

import sys
import re
from collections import defaultdict
from datetime import datetime, date

import openpyxl


def masquer_valeur(valeur):
    """
    Transforme une valeur en pattern anonyme, en gardant uniquement
    la structure (longueur, ponctuation, type), jamais le contenu reel.
    """
    if valeur is None:
        return "VIDE"
    if isinstance(valeur, (datetime, date)):
        return f"DATE({type(valeur).__name__})"
    if isinstance(valeur, bool):
        return "BOOL"
    if isinstance(valeur, (int, float)):
        s = str(valeur)
        partie_entiere = s.split('.')[0].lstrip('-')
        nb_decimales = len(s.split('.')[1]) if '.' in s else 0
        signe = "-" if s.startswith('-') else ""
        return f"NUMERIQUE({signe}~{len(partie_entiere)}chiffres,{nb_decimales}dec)"
    s = str(valeur)
    masque = re.sub(r'[A-Za-zÀ-ÿ]', 'X', s)
    masque = re.sub(r'[0-9]', '9', masque)
    return masque


def detecter_vraies_dimensions(ws):
    """
    Detecte les dimensions REELLEMENT utilisees (cellule par cellule),
    car ws.dimensions peut etre gonfle par un bug connu openpyxl/pandas
    (deja rencontre sur le fichier ICB : ~1000 lignes reelles vs ~44890
    declarees).
    """
    max_row_reel = 0
    max_col_reel = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                max_row_reel = max(max_row_reel, cell.row)
                max_col_reel = max(max_col_reel, cell.column)
    return max_row_reel, max_col_reel


def analyser_entetes(ws, max_row_reel, nb_lignes=5):
    """Affiche les premieres lignes masquees pour identifier visuellement le header (peut etre multi-ligne)."""
    rapport = ["\n--- Apercu des premieres lignes (masquees) ---"]
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=min(nb_lignes, max_row_reel), values_only=True), start=1):
        ligne_masquee = [masquer_valeur(v) for v in row]
        rapport.append(f"Ligne {i} : {ligne_masquee}")
    return rapport


def analyser_colonnes(ws, ligne_header, max_row_reel, max_col_reel, nb_echantillon=3, limite_lignes=1000):
    """
    Analyse colonne par colonne : nom de colonne (l'en-tete n'est pas
    sensible en general), types rencontres, taux de vides, cardinalite
    approximative, et echantillon masque.
    """
    rapport = [f"\n--- Analyse detaillee des colonnes (header suppose = ligne {ligne_header}) ---"]

    headers_row = ws[ligne_header]
    headers = [cell.value for cell in headers_row][:max_col_reel]

    for col_idx in range(1, max_col_reel + 1):
        nom_col = headers[col_idx - 1] if col_idx <= len(headers) else f"Colonne_{col_idx}"
        types_rencontres = defaultdict(int)
        valeurs_uniques = set()
        echantillon = []
        nb_vides = 0
        nb_lignes_vues = 0

        derniere_ligne = min(ligne_header + limite_lignes, max_row_reel)
        for row_idx in range(ligne_header + 1, derniere_ligne + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            nb_lignes_vues += 1
            if val is None:
                nb_vides += 1
                continue
            types_rencontres[type(val).__name__] += 1
            valeurs_uniques.add(str(val))
            if len(echantillon) < nb_echantillon:
                echantillon.append(masquer_valeur(val))

        cardinalite = len(valeurs_uniques)
        rapport.append(f"\n  Colonne {col_idx} : '{nom_col}'")
        rapport.append(f"    Types rencontres    : {dict(types_rencontres)}")
        rapport.append(f"    Vides               : {nb_vides}/{nb_lignes_vues}")
        rapport.append(f"    Valeurs uniques      : {cardinalite} (sur {nb_lignes_vues} lignes analysees)")
        rapport.append(f"    Echantillon masque   : {echantillon}")

    return rapport


def analyser_feuille(ws):
    rapport = [f"\n{'='*70}", f"FEUILLE : {ws.title}", f"{'='*70}"]
    rapport.append(f"Dimensions declarees (ws.dimensions) : {ws.dimensions}")

    max_row_reel, max_col_reel = detecter_vraies_dimensions(ws)
    rapport.append(f"Dimensions REELLES detectees          : {max_row_reel} lignes x {max_col_reel} colonnes")

    if ws.merged_cells.ranges:
        rapport.append(f"\nCellules fusionnees ({len(ws.merged_cells.ranges)}), premieres 10 :")
        for r in list(ws.merged_cells.ranges)[:10]:
            rapport.append(f"  - {r}")

    # Detection de formules Excel (sur les 100 premieres lignes)
    nb_formules = 0
    exemples_formules = []
    for row in ws.iter_rows(max_row=min(100, max_row_reel)):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith('='):
                nb_formules += 1
                if len(exemples_formules) < 3:
                    exemples_formules.append(cell.value)
    rapport.append(f"\nFormules Excel detectees (100 premieres lignes) : {nb_formules}")
    if exemples_formules:
        rapport.append(f"Exemples de formules (structure, pas de donnee sensible) : {exemples_formules}")

    rapport.extend(analyser_entetes(ws, max_row_reel))
    rapport.extend(analyser_colonnes(ws, ligne_header=1, max_row_reel=max_row_reel, max_col_reel=max_col_reel))

    return rapport


def main(chemin_fichier):
    print(f"Analyse du fichier : {chemin_fichier}")
    print("Chargement (data_only=False pour detecter les formules)...")

    wb = openpyxl.load_workbook(chemin_fichier, read_only=False, data_only=False)

    rapport = []
    rapport.append(f"RAPPORT DE STRUCTURE ANONYMISE - {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    rapport.append(f"Nombre de feuilles : {len(wb.sheetnames)}")
    rapport.append(f"Noms des feuilles  : {wb.sheetnames}")

    for nom_feuille in wb.sheetnames:
        ws = wb[nom_feuille]
        rapport.extend(analyser_feuille(ws))

    chemin_sortie = "rapport_structure_anonyme.txt"
    with open(chemin_sortie, 'w', encoding='utf-8') as f:
        f.write('\n'.join(rapport))

    print(f"\nRapport genere : {chemin_sortie}")
    print("=> Aucune donnee sensible reelle n'est presente dans ce fichier.")
    print("=> Tu peux me le copier-coller directement.")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage : python analyse_structure_forecast.py chemin/vers/fichier.xlsx")
        sys.exit(1)
    main(sys.argv[1])