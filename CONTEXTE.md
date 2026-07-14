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
  * Fournit des API RESTful (`/api/hierarchy`, `/api/transactions`, `/api/months`, `/api/export`) pour alimenter l'interface.
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
  * **Colonnes Groupées** : Les métriques clés (`NCR`, `Working Days`, `Bill Amount`) sont désormais des en-têtes parents. Sous chacune d'elles se trouvent deux sous-colonnes centrées : **`ICB`** (valeur réelle provenant du backend) et **`FORECAST`** (colonne vide pour la prévision).
* **État des tableaux par défaut** : Lors du chargement, tous les projets sont "repliés" (collapsed) pour ne pas surcharger la vue.
* **Recherche Dynamique** :
  * Recherche textuelle (sans tenir compte de la casse) sur le `Project Name`, le `CC Number`, l'ID Consultant ou son Nom.
  * **Persistance de l'état** : Effacer la recherche (ou la vider) restaure l'état exact des projets (ouverts ou fermés) tels qu'ils étaient avant ou pendant la recherche, sans tout forcer à s'ouvrir.
* **Filtres de Date (Global)** :
  * Deux champs input `Du` et `Au` permettent de filtrer tout le fichier sur une période précise.
  * **Barre des mois dynamique** : Se met à jour en temps réel pour n'afficher que les mois inclus dans l'intervalle "Du/Au" et contenant des données.

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
    5. **`Forecast bill amount`** : Colonne laissée vide pour la saisie manuelle.
    6. **`Gap`** : Contient une formule Excel active (ex: `=D2-E2` soit `Bill Amount - Forecast`) qui se met à jour seule à l'ouverture d'Excel.
    7. `Raw Cost`
    8. **`Comment`** : Colonne vide pour ajout de notes manuelles.

---

## 5. Notes pour le prochain Agent / Développeur
* **Structure In-Memory** : Actuellement, le fichier chargé est stocké dans la variable globale `pipeline` dans `icb_app.py`. Cela convient à une utilisation Desktop mono-utilisateur.
* **Normalisation des clés** : Attention dans `icb_app.py`, la fonction `_normalize_key` est très importante. Le `CC Project Number` peut parfois être lu comme un `int` (ex: `1234`) ou une `str` (ex: `P-1234`). La normalisation empêche les plantages lors des tris côté Python.
* **Dates Excel** : La fonction `excel_serial_to_date` est essentielle, car `pandas/openpyxl` lisent parfois les dates sous forme de numéros de série (float) à partir de 1899.
* **Pywebview Threading** : Garde à l'esprit que les boîtes de dialogue natives (comme `create_file_dialog`) nécessitent que la fenêtre de pywebview soit accessible et gérée sur le bon thread. C'est le cas actuellement.
