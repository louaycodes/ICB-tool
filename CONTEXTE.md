# Contexte du Projet : ICB Analysis Tool (FinOps Agent)

Ce document centralise tout le contexte métier et technique du projet. Il est conçu pour permettre à un autre agent ou développeur de reprendre facilement le projet et d'en comprendre l'architecture, la logique, et l'historique des fonctionnalités implémentées.

---

## 1. But du Projet
**ICB Analysis Tool** est une application hybride (Desktop/Web) conçue pour aider les équipes FinOps (Capgemini/ICB) à analyser, filtrer et exporter des données de facturation inter-compagnies ("Inter-Company Billing" ou "Expenditures"). 
Le but est de charger de volumineux fichiers Excel de facturation, d'agréger les données, de détecter d'éventuelles anomalies financières, et de fournir un export structuré.

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
* **Raw Cost** et **Bill Amount** : Conservés tels quels depuis le fichier source.

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
  * **Colonnes Groupées** : Les métriques clés (`NCR`, `Working Days`, `Bill Amount`) sont désormais des en-têtes parents. Sous chacune d'elles se trouvent deux sous-colonnes centrées : **`ICB`** (valeur réelle provenant du backend) et **`FORECAST`** (valeur extraite du fichier Forecast importé, remplie automatiquement par consultant et par projet après import).
* **Surlignage des écarts (Gap Highlight)** : Pour chaque paire de cellules ICB / FORECAST (NCR, Working Days, Bill Amount), si les deux valeurs sont présentes et diffèrent (tolérance > 0.01), elles sont automatiquement surlignées en rouge translucide (recalculé dynamiquement à chaque affichage).
* **État des tableaux par défaut** : Lors du chargement, tous les projets sont "repliés" (collapsed) pour ne pas surcharger la vue.
* **Recherche Dynamique** :
  * Recherche textuelle (sans tenir compte de la casse) sur le `Project Name`, le `CC Number`, l'ID Consultant ou son Nom.
  * **Persistance de l'état** : Effacer la recherche (ou la vider) restaure l'état exact des projets (ouverts ou fermés) tels qu'ils étaient avant ou pendant la recherche, sans tout forcer à s'ouvrir.
* **Filtres de Date (Global)** :
  * Deux champs input `Du` et `Au` permettent de filtrer tout le fichier sur une période précise.
  * **Barre des mois dynamique** : Se met à jour en temps réel pour n'afficher que les mois inclus dans l'intervalle "Du/Au" et contenant des données.

### Import et Intégration du Fichier Forecast
* **Bouton "Importer Forecast"** : Apparaît dans le header après le chargement du fichier ICB. Permet d'uploader un fichier `.xlsx` de Forecast. Le bouton disparaît après import (un seul import par session).
* **Parsing dynamique de la structure** (`load_forecast` dans `ICBPipeline`) :
  * **Détection automatique de la ligne d'en-têtes** : Scan des 20 premières lignes, identification par présence d'au moins 3 marqueurs parmi `Project Num`, `Project Name`, `GGID`, `Employee Name`, `Grade`.
  * **Colonnes d'identité trouvées par nom** (pas par position fixe) : `Project Num` = clé de mapping projet (correspond à `CC Project Number` de l'ICB), `GGID` = clé de mapping consultant (correspond à `Employee Number` de l'ICB, après retrait du préfixe `6427_` des deux côtés).
  * **Blocs mensuels détectés dynamiquement** : Les noms de mois (FR + EN : Janvier/January/Jan, … , Décembre/December/Dec) sont cherchés dans les lignes au-dessus du header. Chaque bloc couvre 6 sous-colonnes : `Grade`, `F.RC`, `Vaca`, `Prod`, `Cost`, `FTE`. Si l'ordre est non standard, fallback par recherche de nom dans une plage élargie.
  * **Exclusion automatique des blocs Total/TR** : Les blocs dont le label contient `Total` ou commence par `TR ` (TR Q1, TR Q2, etc.) sont ignorés.
  * **Détection des blocs projets** : Via plages fusionnées (merged cells) dans la colonne `Project Num` détectée, ou par rupture de valeur en fallback. La propagation du numéro de projet gère les cellules fusionnées (seule la cellule d'ancrage porte la valeur, les suivantes sont `None`).
  * **Lignes Total exclues** : Toute ligne où `GGID` est vide est considérée comme une ligne résumé/total et ignorée.
  * **Garde-fous** : Lignes parasites détectées (valeur de Project Num correspondant à un libellé d'en-tête ou un nom de mois), arrêt propre sans crash si colonnes obligatoires introuvables.
* **Mapping Forecast → colonnes du tableau** :
  * `F.RC` → **NCR FORECAST** (taux journalier, moyenne pondérée si plusieurs mois)
  * `Prod` → **Working Days FORECAST**
  * `Cost` → **Bill Amount FORECAST**
  * Le NCR Forecast est la moyenne des F.RC non nulles sur les mois actifs. Si aucune F.RC, calcul dérivé : `(Cost / Prod) × 8`.
* **Sérialisation complète** : Les champs `forecast_ncr`, `forecast_working_days`, `forecast_bill_amount` sont inclus dans le JSON de chaque consultant renvoyé par `/api/hierarchy`.
* **Stockage in-memory** : `pipeline.forecast_data` est un dict `{(proj_norm, ggid_norm, month_key): {ncr, working_days, bill_amount}}` et `pipeline.forecast_names` stocke les noms extraits du Forecast.

### Logs de Diagnostic (Forecast)
* **Visibles uniquement dans le terminal** (`print()` côté serveur), aucun impact UI.
* **6 sections structurées** affichées lors de chaque import Forecast :
  1. **Détection de la ligne d'en-têtes** : ligne trouvée, nombre de colonnes
  2. **Colonnes d'identité** : position Excel de chaque colonne trouvée
  3. **Blocs mensuels** : chaque mois avec ses colonnes F.RC/Prod/Cost, blocs Total/TR ignorés
  4. **Blocs projets** : merged cells ou rupture de valeur, lignes Total détectées
  5. **Mapping ICB** : projets matchés/non matchés, consultants matchés/non matchés (avec GGID brut et normalisé)
  6. **Résultat final** : résumé en une ligne `[FORECAST IMPORT] X projets, Y/Z consultants matchés, W cellules remplies`
* **Aucun blocage si anomalie** : les problèmes sont loggés et le traitement continue avec ce qui peut être mappé.

### Mécanisme d'Exportation Excel
* **Export "Ce que tu vois est ce que tu exportes"** : Le bouton Exporter n'ouvre plus de modale. Il prend les filtres actuels de dates (Global + Barre de mois) et exporte ces données.
* **Dialogue Natif (Pywebview)** : Le backend intercepte la demande et ouvre une fenêtre de sauvegarde native (`pywebview.FileDialog.SAVE`) pour choisir précisément où stocker le fichier.
* **Format de l'Export** :
  * Uniquement les agrégats de projets (pas le détail des consultants).
  * Les onglets (sheets) sont divisés par **Mois**.
  * **8 Colonnes générées** : 
    1. `Project Name`
    2. `CC Project Number` (ces deux premières ont été séparées pour plus de clarté)
    3. `Working Days`
    4. `Bill Amount`
    5. **`Forecast bill amount`** : Somme automatique du Forecast Bill Amount pour le mois et projet donnés. Vide si aucune donnée forecast n'est disponible.
    6. **`Gap`** : Contient une formule Excel active (ex: `=D2-E2` soit `Bill Amount - Forecast`) qui se met à jour seule à l'ouverture d'Excel.
    7. `Raw Cost`
    8. **`Comment`** : Généré automatiquement en cas d'écart (Gap > 0.01). Analyse les écarts (Jours travaillés, Taux NCR, Consultants absents de l'ICB ou du Forecast) et produit un commentaire nominatif détaillé, listant jusqu'à 5 consultants en cause avec leurs ID et montants.

---

## 5. Notes pour le prochain Agent / Développeur
* **Structure In-Memory** : Actuellement, le fichier chargé est stocké dans la variable globale `pipeline` dans `icb_app.py`. Cela convient à une utilisation Desktop mono-utilisateur.
* **Normalisation des clés** : Attention dans `icb_app.py`, la fonction `_normalize_key` est très importante. Le `CC Project Number` peut parfois être lu comme un `int` (ex: `1234`) ou une `str` (ex: `P-1234`). La normalisation empêche les plantages lors des tris côté Python.
* **Normalisation GGID / Employee Number** : Les deux fichiers (ICB et Forecast) utilisent un préfixe `6427_` devant l'ID consultant. Ce préfixe est **retiré des deux côtés** avant comparaison (`.replace("6427_", "")`). Si un fichier de production utilise un préfixe différent, cette logique devra être adaptée.
* **Dates Excel** : La fonction `excel_serial_to_date` est essentielle, car `pandas/openpyxl` lisent parfois les dates sous forme de numéros de série (float) à partir de 1899.
* **Pywebview Threading** : Garde à l'esprit que les boîtes de dialogue natives (comme `create_file_dialog`) nécessitent que la fenêtre de pywebview soit accessible et gérée sur le bon thread. C'est le cas actuellement.
* **Parsing Forecast robuste** : Le parser Forecast (`load_forecast`) est conçu pour fonctionner même si la structure du fichier est décalée (ligne d'en-têtes pas à la ligne 6, colonnes d'identité pas en A-D, blocs mensuels pas à partir de la colonne G). Tout est détecté par nom. En cas de fichier de production avec une structure inconnue, les logs de diagnostic dans le terminal permettent d'identifier rapidement le problème.
* **Double ouverture du fichier Forecast** : `load_forecast` ouvre le fichier deux fois — une première fois en mode normal (`data_only=True`) pour la détection de structure (merged cells, headers), puis une seconde fois en `read_only=True` pour le parsing rapide des données. C'est voulu pour la fiabilité (merged cells non accessibles en read_only).
* **Fichier de référence** : Le fichier `structureDetailleICBandForecast.txt` contient la documentation technique détaillée de la structure des deux fichiers Excel (ICB et Forecast), colonnes, types, et règles de mapping.
