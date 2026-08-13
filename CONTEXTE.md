# Contexte du Projet : ICB Analysis Tool (FinOps Agent)

Ce document centralise tout le contexte métier et technique du projet. Il est conçu pour permettre à un autre agent ou développeur de reprendre facilement le projet et d'en comprendre l'architecture, la logique, et l'historique des fonctionnalités implémentées.

---

## 1. But du Projet
**ICB Analysis Tool** est une application hybride (Desktop/Web) conçue pour aider les équipes FinOps (Capgemini/ICB) à analyser, filtrer et exporter des données de facturation inter-compagnies ("Inter-Company Billing" ou "Expenditures"). 
Le but est de charger de volumineux fichiers Excel de facturation, d'agréger les données, de détecter d'éventuelles anomalies financières, de les comparer à un fichier Forecast, et de fournir un export structuré.

---

## 2. Architecture Technique
L'application repose sur une architecture simple et légère sans base de données (traitement "In-Memory" pour la session courante) :

* **Backend (Python / Flask)** : `icb_app.py`
  * Gère le serveur web (port 5050 par défaut).
  * Traite l'ingestion de fichiers `.xlsx` avec `openpyxl` et `pandas` (pour un chargement rapide).
  * Gère l'import du fichier Forecast (`/api/upload_forecast`) avec détection dynamique de la structure.
  * Fournit des API RESTful (`/api/hierarchy`, `/api/transactions`, `/api/months`, `/api/export`, `/api/upload_forecast`) pour alimenter l'interface.
* **Frontend (Vanilla HTML/CSS/JS)** : `templates/index.html`
  * Interface réactive construite sans framework lourd.
  * Manipulation directe du DOM pour la création du tableau hiérarchique expansible (Projet > Consultant > Transactions).
* **Wrapper Desktop (Python / PyWebview)** : `icb_desktop.py`
  * Lance le serveur Flask dans un thread séparé.
  * Ouvre une fenêtre native OS qui charge `http://localhost:5050`, donnant l'illusion d'une vraie application Desktop (au lieu d'un onglet de navigateur).
* **Lancement** : Script Bash `launch.sh` qui démarre l'application.

---

## 3. Modèle Sémantique (Logique Métier)

### Hiérarchie des Données
1. **Niveau 1 : CC Project** (Identifié par `CC Project Number` et `Project Name`).
2. **Niveau 2 : Consultant** (Identifié par `Employee Number` et `Full Name`).
3. **Niveau 3 : Transactions (Expenditures)** (Chaque ligne de dépense individuelle du consultant dans ce projet).

### Règles de Calculs Financiers
* **Working Days** : Calculé en divisant la quantité d'heures brutes (`Quantity`) par **8**.
* **NCR (Taux journalier)** : Calculé en multipliant le `Bill Rate` existant (qui est à l'heure) par **8**.
* **Raw Cost** : Conservé tel quel depuis le fichier source (colonne `Raw Cost`, col AA).
* **Cost (anciennement "Bill Amount")** : Source ICB provenant de la colonne `Project Transfer Price` du fichier ICB. La colonne est détectée dynamiquement par son nom dans les en-têtes (comparaison insensible à la casse et aux espaces). Si la colonne est introuvable, une erreur explicite est loggée dans le terminal. **Important** : la variable interne reste nommée `bill_amount` / `total_bill_amount` dans le code Python et le JSON API, mais l'affichage UI et l'export Excel utilisent le libellé "Cost".
* **NCR au niveau projet** : Calculé comme la **moyenne simple** des NCR des consultants du projet (pas une somme). Idem pour le NCR Forecast au niveau projet.

### Détection des Anomalies (NCR)
* Basée sur la colonne source `FULL NCR to Direct NCR Correlation`.
* **Cible (Target)** : 0.75 (soit 75%).
* **Tolérance** : ± 0.05.
* **Règle métier** : Une ligne est considérée comme une anomalie si la valeur absolue de cette corrélation n'est **PAS** comprise entre `[0.70, 0.80]`. 

---

## 4. Fonctionnalités Implémentées (Historique de développement)

### UI & Expérience Utilisateur
* **Design "Dark/Corporate"** : Entête fixée (sticky) avec décalage dynamique pour que les en-têtes complexes (multi-lignes) du tableau restent visibles lors du scroll sans se superposer.
* **Mise en forme du tableau en grille** :
  * Séparations verticales ajoutées sur toutes les lignes pour améliorer la lisibilité.
  * **Colonnes Groupées** : Les métriques clés (`NCR`, `Working Days`, `Cost`) sont des en-têtes parents. Sous chacune d'elles se trouvent deux sous-colonnes centrées : **`ICB`** (valeur réelle provenant du backend) et **`FORECAST`** (valeur extraite du fichier Forecast importé, remplie automatiquement par consultant et par projet après import).
* **Surlignage des écarts (Gap Highlight)** : Pour chaque paire de cellules ICB / FORECAST :
  * **NCR et Cost** : Surlignage rouge si les valeurs diffèrent de plus que le seuil de tolérance configurable (défaut : 2). Ce seuil est modifiable en temps réel via le champ "Tolérance Ecart (Montant/NCR)" de l'interface.
  * **Working Days** : Comparaison **stricte** (sans tolérance) — tout écart, même de 0.01 jour, déclenche le surlignage.
  * Le surlignage est recalculé dynamiquement à chaque affichage ou changement de tolérance.
* **État des tableaux par défaut** : Lors du chargement, tous les projets sont "repliés" (collapsed) pour ne pas surcharger la vue.
* **Recherche Dynamique** :
  * Recherche textuelle (sans tenir compte de la casse) sur le `Project Name`, le `CC Number`, l'ID Consultant ou son Nom.
  * **Persistance de l'état** : Effacer la recherche (ou la vider) restaure l'état exact des projets (ouverts ou fermés) tels qu'ils étaient avant ou pendant la recherche, sans tout forcer à s'ouvrir.
* **Filtres de Date (Global)** :
  * Deux champs input `Du` et `Au` permettent de filtrer tout le fichier sur une période précise.
  * **Barre des mois dynamique** : Se met à jour en temps réel pour n'afficher que les mois inclus dans l'intervalle "Du/Au" et contenant des données.

### Import et Intégration du Fichier Forecast
* **Bouton "Importer Forecast"** : Apparaît dans le header après le chargement du fichier ICB. Permet d'uploader un fichier `.xlsx` de Forecast. Le bouton disparaît après import (un seul import par session). Le mécanisme de chargement est aligné sur celui de l'ICB : sauvegarde disque dans `uploads/`, lecture via openpyxl, fermeture explicite du handle fichier.
* **Parsing dynamique de la structure** (`load_forecast` dans `ICBPipeline`) :
  * **Détection automatique de la ligne d'en-têtes** : Scan des 20 premières lignes, identification par présence d'au moins 3 marqueurs parmi `Project Num`, `Project Name`, `GGID`, `Employee Name`, `Grade`.
  * **Colonnes d'identité trouvées par nom** (pas par position fixe) : `Project Num` = clé de mapping projet (correspond à `CC Project Number` de l'ICB), `GGID` = clé de mapping consultant (correspond à `Employee Number` de l'ICB, après retrait du préfixe `6427_` des deux côtés).
  * **Blocs mensuels détectés dynamiquement** : Les noms de mois (FR + EN : Janvier/January/Jan, … , Décembre/December/Dec) sont cherchés dans les lignes au-dessus du header. Chaque bloc couvre 6 sous-colonnes : `Grade`, `F.RC`, `Vaca`, `Prod`, `Cost`, `FTE`. Si l'ordre est non standard, fallback par recherche de nom dans une plage élargie.
  * **Exclusion automatique des blocs Total/TR** : Les blocs dont le label contient `Total` ou commence par `TR ` (TR Q1, TR Q2, etc.) sont ignorés.
  * **Détection des blocs projets** : Via plages fusionnées (merged cells) dans la colonne `Project Num` détectée, ou par rupture de valeur en fallback. La propagation du numéro de projet gère les cellules fusionnées (seule la cellule d'ancrage porte la valeur, les suivantes sont `None`).
  * **Lignes Total exclues** : Toute ligne où `GGID` est vide est considérée comme une ligne résumé/total et ignorée.
  * **Garde-fous** : Lignes parasites détectées (valeur de Project Num correspondant à un libellé d'en-tête ou un nom de mois), arrêt propre sans crash si colonnes obligatoires introuvables.
* **Mapping Forecast → colonnes du tableau** :
  * `F.RC` → **NCR FORECAST** (taux journalier, moyenne des F.RC non nulles sur les mois actifs ; si aucune F.RC, calcul dérivé : `(Cost / Prod) × 8`)
  * `Prod` → **Working Days FORECAST**
  * `Cost` → **Cost FORECAST** (alias Forecast bill amount dans l'export)
* **Sérialisation complète** : Les champs `forecast_ncr`, `forecast_working_days`, `forecast_bill_amount` sont inclus dans le JSON de chaque consultant renvoyé par `/api/hierarchy`.
* **Stockage in-memory** : `pipeline.forecast_data` est un dict `{(proj_norm, ggid_norm, month_key): {ncr, working_days, bill_amount}}` et `pipeline.forecast_names` stocke les noms extraits du Forecast.

### Logs de Diagnostic (Forecast Import)
* **Conservés volontairement en production** — visibles uniquement dans le terminal (`print()` côté serveur), aucun impact UI.
* **6 sections structurées** affichées lors de chaque import Forecast :
  1. **Détection de la ligne d'en-têtes** : ligne trouvée, nombre de colonnes
  2. **Colonnes d'identité** : position Excel de chaque colonne trouvée
  3. **Blocs mensuels** : chaque mois avec ses colonnes F.RC/Prod/Cost, blocs Total/TR ignorés
  4. **Blocs projets** : merged cells ou rupture de valeur, lignes Total détectées
  5. **Mapping ICB** : projets matchés/non matchés, consultants matchés/non matchés (avec GGID brut et normalisé)
  6. **Résultat final** : résumé en une ligne `[FORECAST IMPORT] X projets, Y/Z consultants matchés, W cellules remplies`
* **Aucun blocage si anomalie** : les problèmes sont loggés et le traitement continue avec ce qui peut être mappé.
* **Note** : tous les logs de debug temporaires (préfixes `[DIAG-BACKEND]`, `[DIAG-FRONTEND]`, compteurs de clics, dumps JSON, traces d'insertion DOM) ont été supprimés. Seul ce bloc structurel de Forecast Import reste.

### Mécanisme d'Exportation Excel
* **Export "Ce que tu vois est ce que tu exportes"** : Le bouton Exporter prend les filtres actuels de dates (Global + Barre de mois) et le seuil de tolérance configurable, et exporte ces données.
* **Dialogue Natif (Pywebview)** : Le backend intercepte la demande et ouvre une fenêtre de sauvegarde native (`pywebview.FileDialog.SAVE`) pour choisir précisément où stocker le fichier. Fallback sur le Bureau si pywebview n'est pas disponible.
* **Format de l'Export** :
  * Uniquement les agrégats de projets (pas le détail des consultants).
  * Les onglets (sheets) sont divisés par **Mois**.
  * **Vrai tableau Excel structuré** : filtres automatiques, bandes alternées (style `TableStyleMedium9`).
  * **10 Colonnes générées** : 
    1. `EM_Lead` — vide, réservée pour saisie manuelle future
    2. `Project Name`
    3. `CC Project Number`
    4. `Working Days`
    5. `Cost` — alimentée par la colonne `Project Transfer Price` de l'ICB
    6. `Forecast bill amount` — somme automatique du Forecast Cost pour le mois et projet donnés. Vide si aucune donnée forecast disponible.
    7. `Gap` — formule Excel active (`=E{row}-F{row}` soit `Cost - Forecast bill amount`) qui se met à jour seule à l'ouverture d'Excel.
    8. `Raw Cost`
    9. `Issue Type` — vide, réservée pour saisie manuelle future
    10. `Generated comment` — commentaire automatique en cas d'écart (voir ci-dessous)
* **Logique de génération des commentaires (`Generated comment`)** :
  * Déclenché si un écart `Gap` dépasse le seuil de tolérance configurable (défaut : 2).
  * Analyse les écarts en partant **exclusivement** des consultants présents dans l'ICB (pas du Forecast).
  * Détail nominatif par consultant (ID + nom si disponible) pour chaque cause identifiée :
    * **Écart de jours travaillés** (`Working Days`) : comparaison stricte
    * **Écart de taux NCR** : si la différence dépasse le seuil de tolérance
    * **Écart de Cost** (fallback) : si aucun écart WD ou NCR n'est identifié, les 5 consultants contribuant le plus à la différence de Cost sont listés avec le libellé "Probleme avec Cost".
    * **Fallback final** : "Probleme avec Cost (a verifier manuellement)" si la cause reste diffuse.
  * Maximum 5 consultants affichés par cause, avec indication du nombre d'autres consultants restants.

### Gap signé (sans valeur absolue)
* **Changement** : La colonne **GAP** dans le tableau UI et dans l'export Excel affiche désormais la **valeur signée** de l'écart (`Cost ICB - Forecast`), et non plus la valeur absolue.
  * **Positif** : ICB > Forecast (on facture plus que prévu).
  * **Négatif** : ICB < Forecast (on facture moins que prévu).
* **Logique de surlignage** : Le surlignage rouge reste déclenché si `|gap| > seuil de tolérance` (la détection utilise toujours la valeur absolue pour identifier les écarts significatifs, mais la valeur affichée montre le sens).
* **Export Excel** : La formule `=E{row}-F{row}` dans la colonne Gap était déjà signée par nature. Le commentaire du code a été mis à jour pour refléter cette intention.

### Feuille KPI Dashboard dans l'Export Excel
* **Nouvelle feuille** ajoutée automatiquement à la fin du fichier Excel d'export, intitulée **"KPI Dashboard"**.
* **5 sections analytiques** stylisées avec couleurs, bordures, et formats numériques :

  1. **Synthèse Mensuelle** : Tableau récapitulatif par mois avec Cost ICB, Forecast, Gap (signé), Working Days, Raw Cost, Nb Projets, Nb Consultants. Ligne TOTAL en bas avec fond bleu accentué. Les écarts négatifs sont surlignés en rouge clair, les positifs en vert clair.

  2. **Bill Amount Cumulé par Mois** *(exigence spécifique)* : Montre la progression du Bill Amount (Cost ICB) mois après mois :
     * Mois 1 : valeur du mois 1
     * Mois 2 : valeur cumulée du mois 1 + mois 2
     * Mois 3 : valeur cumulée du mois 1 + mois 2 + mois 3
     * Et ainsi de suite...
     * Inclut aussi le Forecast cumulé, le Gap cumulé, et le pourcentage de progression du cost d'un mois à l'autre.
     * La colonne Cost ICB Cumulé est en **gras vert** sur fond vert clair, la colonne Forecast Cumulé en **gras bleu**.

  3. **Synthèse par Projet** : Vue consolidée de chaque projet sur toutes les périodes, triée par Cost ICB décroissant. Inclut : Cost ICB Total, Forecast Total, Gap (signé), Gap en %, Working Days, NCR Moyen, Nb Consultants uniques. Alternance de couleur de fond pour la lisibilité.

  4. **Top Consultants par Cost** : Les 50 premiers consultants triés par Cost ICB total décroissant. Inclut : Employee ID, Nom, Projet principal, Cost ICB Total, Working Days Total, NCR moyen, Nb Mois Actifs, Cost Moyen par Mois.

  5. **Détail Consultant par Mois** : Tableau croisé avec un consultant par ligne et un mois par colonne, affichant le Cost ICB de chaque consultant pour chaque mois. Colonne TOTAL en fin de ligne (en gras vert sur fond vert clair). Trié par nom alphabétique.

* **Méthode `_build_kpi_sheet`** : Méthode séparée dans `ICBPipeline` pour la maintenabilité. Appelée automatiquement par `export_to_excel` après la création des onglets mensuels.
* **Styles** : Palette cohérente avec le thème de l'application (bleu foncé/corporate). Utilise `openpyxl.styles` (Font, PatternFill, Border, Alignment) et des formats numériques Excel natifs (`#,##0.00`, `0.0%`).

---

## 5. Notes pour le prochain Agent / Développeur
* **Structure In-Memory** : Actuellement, le fichier chargé est stocké dans la variable globale `pipeline` dans `icb_app.py`. Cela convient à une utilisation Desktop mono-utilisateur.
* **Normalisation des clés** : Attention dans `icb_app.py`, la fonction `_normalize_key` est très importante. Le `CC Project Number` peut parfois être lu comme un `int` (ex: `1234`) ou une `str` (ex: `P-1234`). La normalisation empêche les plantages lors des tris côté Python.
* **Normalisation GGID / Employee Number** : Les deux fichiers (ICB et Forecast) utilisent un préfixe `6427_` devant l'ID consultant. Ce préfixe est **retiré des deux côtés** avant comparaison (`.replace("6427_", "")`). Si un fichier de production utilise un préfixe différent, cette logique devra être adaptée.
* **Détection dynamique de la colonne "Project Transfer Price"** : Dans `ICBPipeline.load()`, le code parcourt dynamiquement les en-têtes du fichier ICB pour trouver la colonne dont le nom correspond à `"Project Transfer Price"` (comparaison insensible à la casse, espaces ignorés). Si la colonne est introuvable, une erreur explicite est affichée dans le terminal avec la liste des colonnes disponibles. Ce mécanisme évite de coder en dur un index de colonne (ex: AF) qui pourrait changer d'un fichier à l'autre.
* **Naming interne vs affichage** : La variable `self.bill_amount` et les clés JSON (`bill_amount`, `total_bill_amount`, `forecast_bill_amount`) conservent leur nom historique pour éviter de casser les références croisées backend/frontend. Seuls les libellés affichés (UI, export) utilisent le nom "Cost".
* **Dates Excel** : La fonction `excel_serial_to_date` est essentielle, car `pandas/openpyxl` lisent parfois les dates sous forme de numéros de série (float) à partir de 1899.
* **Pywebview Threading** : Les boîtes de dialogue natives (comme `create_file_dialog`) nécessitent que la fenêtre de pywebview soit accessible et gérée sur le bon thread. C'est le cas actuellement.
* **Parsing Forecast robuste** : Le parser Forecast (`load_forecast`) est conçu pour fonctionner même si la structure du fichier est décalée (ligne d'en-têtes pas à la ligne 6, colonnes d'identité pas en A-D, blocs mensuels pas à partir de la colonne G). Tout est détecté par nom. En cas de fichier de production avec une structure inconnue, les logs de diagnostic dans le terminal permettent d'identifier rapidement le problème.
* **Double ouverture du fichier Forecast** : `load_forecast` ouvre le fichier deux fois — une première fois en mode normal (`data_only=True`) pour la détection de structure (merged cells, headers), puis une seconde fois en `read_only=True` pour le parsing rapide des données. C'est voulu pour la fiabilité (merged cells non accessibles en read_only).
* **Nettoyage des NaN (JSON strict)** : L'utilisation de Pandas génère parfois des valeurs `NaN` (float) pour les cellules vides. Pour empêcher que `jsonify` ne produise un token littéral `NaN` (ce qui fait crasher `JSON.parse` côté frontend), toute valeur est passée dans des filtres locaux (ex: `_clean()` dans `ExpenditureItem`, ou `_to_float()`) qui convertissent les `NaN`/`Infinity` en `None` ou `0.0`.
* **Logique des écarts (Commentaires Export)** : Le rapport Excel n'itère que sur les consultants ICB pour chercher leurs équivalents dans le Forecast. Toute absence est considérée comme un turn-over normal et est ignorée.
* **Fichier de référence** : Le fichier `structureDetailleICBandForecast.txt` contient la documentation technique détaillée de la structure des deux fichiers Excel (ICB et Forecast), colonnes, types, et règles de mapping.
* **KPI Dashboard** : La méthode `_build_kpi_sheet` est appelée par `export_to_excel` et collecte toutes les données nécessaires en itérant sur chaque mois via `hierarchy_data()`. Les données sont d'abord agrégées en mémoire (`monthly_stats`, `project_stats`, `consultant_monthly`) puis écrites section par section. La performance est acceptable car le volume de données KPI est faible comparé aux données brutes.
