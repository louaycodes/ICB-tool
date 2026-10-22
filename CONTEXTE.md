# Contexte du Projet : ICB Analysis Tool (FinOps Agent)

Ce document centralise tout le contexte métier et technique du projet. Il est conçu pour permettre à un autre agent ou développeur de reprendre facilement le projet et d'en comprendre l'architecture, la logique, et l'historique des fonctionnalités implémentées.

Structure du document :
1. But du projet
2. Architecture technique
3. Modèle sémantique (fichiers d'entrée, règles métier, mappings)
4. Fonctionnalités actuelles (état courant de l'application)
5. Historique détaillé des modifications (ce qui a été changé, et pourquoi)
6. Bugs résolus et leçons
7. Notes pour le prochain agent / développeur
8. Points à vérifier et pistes futures

---

## 1. But du Projet
**ICB Analysis Tool** est une application hybride (Desktop/Web) conçue pour aider les équipes FinOps (Capgemini/ICB) à analyser, filtrer et exporter des données de facturation inter-compagnies ("Inter-Company Billing" ou "Expenditures").

Le but est de :
* charger de volumineux fichiers Excel de facturation (ICB) ;
* agréger les données par projet / consultant ;
* les comparer à un fichier **Forecast** (prévisionnel), qui est la **source de vérité** ;
* détecter les écarts (gaps) et les expliquer automatiquement ;
* fournir un export Excel structuré (avec commentaires d'écart générés et un dashboard KPI).

---

## 2. Architecture Technique
L'application repose sur une architecture simple et légère sans base de données (traitement "In-Memory" pour la session courante) :

* **Backend (Python / Flask)** : `icb_app.py`
  * Serveur web (port 5050 par défaut).
  * Ingestion des fichiers `.xlsx` avec `openpyxl` et `pandas`.
  * Import du fichier Forecast avec détection dynamique de la structure.
  * API RESTful : `/api/hierarchy`, `/api/transactions`, `/api/months`, `/api/export`, `/api/upload_forecast`.
* **Frontend (Vanilla HTML/CSS/JS)** : `templates/index.html` + `static/style.css`
  * Interface réactive sans framework lourd.
  * Manipulation directe du DOM pour le tableau hiérarchique expansible (Projet > Consultant > Transactions).
* **Wrapper Desktop (Python / PyWebview)** : `icb_desktop.py`
  * Lance le serveur Flask dans un thread séparé.
  * Ouvre une fenêtre native OS qui charge `http://localhost:5050`.
  * Moteurs de rendu : WebKit sur macOS (dev), Edge WebView2 sur Windows (prod) : des différences de rendu CSS sont possibles.
* **Lancement** : script Bash `launch.sh`.

### Organisation du dépôt (après réorganisation)
```
.
├── README.md
├── icb_app.py
├── icb_desktop.py
├── launch.sh
├── requirements.txt
├── .gitignore            (ignore *.exe, dist/, build/, venv, uploads...)
├── static/style.css
├── templates/index.html
└── docs/
    ├── CONTEXTE.md
    └── structureDetailleICBandForecast.txt
```
* `icb_app.py`, `icb_desktop.py`, `launch.sh`, `templates/` et `static/` restent **à la racine** : Flask résout `templates/` et `static/` relativement au script, et un futur build PyInstaller dépend de chemins stables.
* L'artefact de build `icb_desktop 2.exe` ne doit plus être versionné.
* Branche de travail Git : `forecast` (remote `louaycodes/ICB-tool`). En cas de branches divergentes : `git pull origin forecast --no-rebase` (merge), `git config pull.rebase false`.

---

## 3. Modèle Sémantique (Logique Métier)

### Hiérarchie des Données
1. **Niveau 1 : CC Project** (`CC Project Number` + `Project Name`).
2. **Niveau 2 : Consultant** (`Employee Number` + `Full Name`).
3. **Niveau 3 : Transactions (Expenditures)** (chaque ligne de dépense individuelle du consultant dans ce projet).

### Fichier ICB (entrée principale)
* Classeur Excel complet **multi-feuilles** ; la feuille utile s'appelle exactement **`ICB_DB`** (nom comparé exactement, erreur claire si absente).
* Cette feuille a ses **5 premières lignes vides** et sa **colonne A entièrement vide** : la ligne d'en-têtes et la colonne de départ sont détectées dynamiquement (scan des premières lignes à la recherche des en-têtes attendus), jamais supposées en ligne 1 / colonne A.
* Colonnes clés lues **par nom** : `Employee Number`, `CC Project Number`, `Project Name`, `Full Name`, `Expenditure Type`, `Expenditure Item Date`, `Quantity`, `UOM`, `Bill Rate`, `Raw Cost`, `FULL NCR to Direct NCR correlation`, et la colonne de prix de transfert (voir ci-dessous).
* Colonne de prix de transfert : détectée par nom, dans cet ordre, parmi 3 variantes possibles (comparaison insensible à la casse et aux espaces) :
  1. `Project Transfer Price`
  2. `Project Revenue Price`
  3. `Revenue Project TP`
  Si aucune n'existe : erreur explicite dans le terminal (jamais de repli silencieux sur 0).
* `Employee Number` : format `6427_XXX` (préfixe `6427_` + ID).
* Filtres métier usuels : `Expenditure Type` = `RC_Time Std` (et variantes `RC_Overtime Std Cost`, `x1.25`, `x1.5`, `x2`), `UOM` = `HOURS`.

### Fichier Forecast (entrée secondaire)
Classeur Excel à **plusieurs feuilles** :
* **Feuille principale** (nom variable, ex. `Sheet1` ou `Forecast26`) : une ligne par consultant et par projet, 12 blocs mensuels.
  * Colonnes d'identité : `Project Num`, `Project Name`, `GGID`, `Employee Name`.
  * 12 blocs mensuels de 6 sous-colonnes : `Grade`, `F.RC`, `Vaca`, `Prod`, `Cost`, `FTE`.
  * Blocs `Total` (annuel) et `TR Q1..TR Q4` (trimestriels) : ignorés.
  * Aucune année dans le fichier : les mois sont rattachés à **l'année système actuelle**.
  * Relations internes : `Cost = F.RC × Prod` ; `FTE = Prod / 21`.
  * En production : plusieurs projets empilés (bloc consultants + ligne Total par projet, cellules `Project Num`/`Project Name` fusionnées ou non).
* **Feuille `EDM`** (matching Projet → Engagement Manager) :
  * Détection dynamique des colonnes par nom : `Project Num` (ou variantes `Project Number` / `project_num`) et `Engagement Manager` (avec tolérance typo `Engagemenet Manager` / `engagement_manager`).
  * Plus d'obligation de cibler la colonne C ni d'ignorer la colonne A : la colonne détectée par son nom est directement utilisée.
  * Remplace l'ancienne feuille `Project-EM` (Project Code / Project Name / EM Name), qui n'est plus utilisée.

### Règles de Calculs Financiers
* **Working Days** : `Quantity` (heures) / **8**.
* **NCR (taux journalier)** : `Bill Rate` (horaire) × **8**, calculé **exclusivement sur les lignes standard `RC_Time Std`** (si un consultant a des heures supplémentaires, ex: x1.5 ou x2, le NCR affiché reste son taux standard et n'est pas faussé par les majorations). Si un consultant a des heures supplémentaires, un drapeau `OT` (badge ambré) est affiché sur sa ligne. La colonne GAP reste nette et numérique.
* **Working Days & Transfer Price** : la logique reste intacte et cumule l'ensemble des types d'heures (`RC_Time Std` et variantes `RC_Overtime*`).
* **Raw Cost** : conservé tel quel (colonne `Raw Cost`). **Masqué dans l'UI**, présent dans l'export (optionnel). Lorsqu'il est inclus dans le rapport Excel, la corrélation `|Raw Cost| / |Project Transfer Price| = 0.75` est calculée (tolérance 2%) : toute anomalie surbrille la cellule en **rouge clair** (`#FECACA` / texte foncé `#991B1B`) avec règle de mise en forme conditionnelle Excel active.
* **Transfer Price (ex-"Bill Amount", puis "Cost")** : source ICB = colonne `Project Transfer Price` (ou variantes). Variables internes inchangées : `bill_amount`, `total_bill_amount`, `forecast_bill_amount`. Seuls les libellés affichés changent.
* **NCR au niveau projet** (UI) : **moyenne simple** des NCR standards valides (`ncr > 0`) des consultants (ICB et Forecast séparément).

### Détection des Anomalies (NCR)
* Basée sur `FULL NCR to Direct NCR Correlation` : cible 0.75, tolérance ± 0.05.
* Anomalie si la valeur absolue n'est **pas** dans `[0.70, 0.80]`.

### Règles de Mapping ICB ↔ Forecast
| Élément | ICB | Forecast |
|---|---|---|
| Projet | `CC Project Number` | `Project Num` |
| Consultant | `Employee Number` | `GGID` |
| Normalisation | retirer le préfixe `6427_` des deux côtés avant comparaison | idem |
| Clés | `_normalize_key` appliquée aux numéros de projet | idem |

Mapping des métriques Forecast → colonnes du tableau :
* `F.RC` → **NCR FORECAST** (si absent : `(Cost / Prod) × 8`)
* `Prod` → **Working Days FORECAST**
* `Cost` → **Transfer Price FORECAST**

---

## 4. Fonctionnalités actuelles (état courant)

### UI & Expérience Utilisateur
* **Design "Dark/Corporate"**, en-têtes sticky sur deux lignes, grille avec séparations verticales.
* **Application entièrement en anglais** (libellés, messages, dialogues, logs). Les mots-clés de matching (en-têtes de fichiers, noms de feuilles, noms de mois, valeurs filtrées, identifiants techniques) n'ont **pas** été traduits.
* **Colonnes groupées** : `NCR`, `Working Days`, `Transfer Price`, chacune avec deux sous-colonnes `ICB` et `FORECAST`. Plus de colonne Raw Cost dans l'app.
* **Colonne `EM Name`** juste à droite de `Projet`, au niveau projet uniquement, avec un petit bouton ▼ dans l'en-tête :
  * ouvre une liste déroulante : `All (N)`, les EM triés alphabétiquement avec leur nombre de projets, et `(Unassigned)` si nécessaire ;
  * sélection unique, fermeture au clic extérieur et au scroll ;
  * se **combine** avec la recherche textuelle (moteur unique `applyTableFilters()`), sans toucher à l'état ouvert/fermé des projets ;
  * le bouton prend la couleur d'accent quand un filtre est actif ;
  * n'a **aucun** impact sur l'export ni sur la colonne `EM_Lead`.
* **Surlignage des écarts** (Forecast = source de vérité) : seule la cellule **ICB** est surlignée en rouge quand elle diverge du Forecast ; la cellule FORECAST n'est **jamais** rouge.
  * NCR et Transfer Price : écart > seuil de tolérance configurable.
  * Working Days : comparaison **stricte** (toute différence).
  * Aucun surlignage si l'une des deux valeurs est absente, ni avant le chargement de l'ICB.
* **Seuil de tolérance configurable** : champ de saisie dans l'UI (défaut **2**, unité monétaire), source de vérité unique pour le surlignage **et** pour la génération des commentaires d'export. Modification appliquée immédiatement, sans rechargement.
* **Gap signé** : `Transfer Price ICB − Forecast` (positif = ICB > Forecast). Le surlignage se déclenche sur `|gap| > seuil`.
* **Union ICB/Forecast** : un consultant présent dans le Forecast mais absent de l'ICB (ou l'inverse) reste affiché ; un indicateur discret signale ces cas pour faciliter l'investigation.
* **Consultants quittés** : un consultant absent de l'ICB pour un projet/mois **et** dont toutes les valeurs Forecast sont à 0/vides n'apparaît **nulle part** pour ce projet/mois (app, export, commentaires, totaux). Filtre appliqué mois par mois.
* **État par défaut** : tous les projets repliés. **Recherche** (Project Name, CC Number, ID consultant, nom) : filtre sans ouvrir les projets ; effacer la recherche restaure l'état exact d'avant.
* **Filtres de date** `From` / `To` et **barre des mois dynamique**.
* **Boutons de chargement** : le bouton de chargement ICB disparaît après un chargement réussi ; le bouton `Importer Forecast` aussi (un seul import par session). Chaque bouton est indépendant.

### Import du Fichier Forecast
* Bouton dans le header (apparaît après le chargement de l'ICB). Fichier sauvegardé dans `uploads/current_forecast.xlsx`, lu via openpyxl, workbook fermé explicitement.
* **Feuille principale trouvée par détection**, pas par `wb.active` : parcours de **tous** les onglets, en appliquant les marqueurs d'en-têtes (au moins 3 parmi `Project Num`, `Project Name`, `GGID`, `Employee Name`, `Grade`, dans les 20 premières lignes). La feuille `EDM` est **exclue** de cette recherche. Indépendant de l'ordre des onglets.
* **Parsing dynamique** (`load_forecast`) :
  * ligne d'en-têtes trouvée par scan ;
  * colonnes d'identité trouvées par nom ;
  * blocs mensuels détectés par nom de mois (FR + EN, avec/sans accents), avec fallback si l'ordre des 6 sous-colonnes est non standard ;
  * blocs `Total` / `TR ` exclus ;
  * blocs projets via cellules fusionnées, sinon rupture de valeur ; gestion de la cellule d'ancrage des fusions ;
  * ligne Total = `GGID` vide, ignorée ; lignes parasites ignorées ;
  * arrêt propre (sans crash) si une colonne obligatoire est introuvable.
* **Lecture de la feuille `EDM`** : en-têtes scannés dans les 20 premières lignes ; `Project Num` et `Engagement Manager` détectés dynamiquement par nom ; lignes avec valeur vide ignorées ; doublons avec EM différents → warning, première valeur conservée ; absence de feuille = map vide silencieuse. Le champ `em_name` est ajouté à chaque projet dans `/api/hierarchy`.
* **Stockage** : `pipeline.forecast_data` = `{(proj_norm, ggid_norm, month_key): {ncr, working_days, bill_amount}}` ; `pipeline.forecast_names` pour les noms.

### Logs de Diagnostic (conservés volontairement en production)
Visibles uniquement dans le terminal (`print()`), aucun impact UI. Bloc `[FORECAST IMPORT]` structuré :
1. détection de la ligne d'en-têtes ;
2. colonnes d'identité (position Excel) ;
3. blocs mensuels (Total/TR ignorés) ;
4. blocs projets (fusions ou rupture de valeur, lignes Total) ;
5. mapping ICB (projets et consultants matchés/non matchés, GGID brut et normalisé) ;
6. résultat final : `[FORECAST IMPORT] X projets, Y/Z consultants matchés, W cellules remplies`.

S'y ajoutent les logs `EDM` : colonnes retenues (ex. « Project Num = colonne C, Engagement Manager = colonne E, colonne A ignorée »), nombre de projets lus, matchés et sans correspondance.
Tous les logs de debug temporaires (`[DIAG-BACKEND]`, `[DIAG-FRONTEND]`, compteurs de clics, dumps JSON, traces DOM) ont été supprimés.

### Export Excel
* "Ce que tu vois est ce que tu exportes" : respecte les filtres de dates (Global + barre des mois) et le seuil de tolérance. Dialogue de sauvegarde natif pywebview (fallback Bureau).
* Un onglet par **mois**, chacun en **vrai tableau Excel** (`TableStyleMedium9` : filtres automatiques, bandes alternées).
* **Ligne de Totaux au sommet (Ligne 1)** : au-dessus des en-têtes (Ligne 2), une ligne de synthèse avec fond sombre (`#1E293B`) et texte blanc en gras affiche les sommes dynamiques via `=SUBTOTAL(9, ...)` (respectant les filtres de tableau Excel) pour `Working Days`, `ICB Transfer Price`, `Forecast Transfer Price`, `Gap` (`=ICB - Forecast`) et `Raw Cost`.
* **Volets figés (`Freeze Panes A3`)** : les lignes 1 (Totaux) et 2 (En-têtes) restent épinglées en haut de l'écran lors du défilement.
* **Dialogue d'export** : liste de cases à cocher, même style visuel qu'avant :
  1. `Working Days` : décochée par défaut
  2. `Raw Cost` : décochée par défaut
  3. `Generated comment` : cochée par défaut
  4. `Controlling comment` : cochée par défaut, **dernière** de la liste
  Une case décochée retire complètement la colonne du tableau (elle n'est pas laissée vide). Les autres colonnes sont toujours présentes.
* **Colonnes de l'export** (dans l'ordre) :
  1. `EM_Lead` : vide, saisie manuelle (indépendante de `EM Name`)
  2. `Project Name`
  3. `CC Project Number`
  4. `Working Days` *(optionnelle)*
  5. `ICB Transfer Price`
  6. `Forecast Transfer Price` : somme du Forecast pour le projet/mois ; vide si aucune donnée Forecast (distinction « pas de donnée » / « 0 » faite via `forecast_data`)
  7. `Gap` : formule Excel active `ICB Transfer Price − Forecast Transfer Price` (signée)
  8. `Raw Cost` *(optionnelle)*
  9. `Issue Type` : vide, saisie manuelle
  10. `Generated comment` *(optionnelle)*
  11. `Controlling comment` *(optionnelle)* : vide, saisie manuelle
* Feuille **KPI Dashboard** ajoutée à la fin (voir plus bas).

### Génération des commentaires (`Generated comment`)
* Écrits en **anglais**.
* Déclenchés si le gap dépasse le seuil de tolérance configurable.
* Analyse partant **de l'ICB** (source), en cherchant l'équivalent dans le Forecast ; un consultant sans équivalent Forecast n'est jamais signalé (turn-over normal). Les catégories « absent de l'ICB » / « absent du Forecast » **n'existent plus**.
* **Regroupés par consultant** (et non par métrique). Un consultant n'apparaît qu'une fois, avec une ligne par gap détecté pour lui :
  ```
  30065302 (Name)
  Gap due to working days: ICB=95.0 D/Forecast=19.0 D
  Gap due to rate (NCR): ICB=160.8/Forecast=143.11

  50612319 (Name)
  Gap due to working days: ICB=100.0 D/Forecast=20.0 D
  ```
* Code consultant normalisé (sans `6427_`) + nom. Jours affichés avec le suffixe ` D` (ex. `95.0 D`).
* **Aucune troncature** : tous les consultants concernés sont listés (plus de « ... and X other(s) »).
* Une cause par ligne dans la **même cellule** : `"\n".join(...)`, `Alignment(wrap_text=True, vertical='top')` et hauteur de ligne recalculée (~15 pts par ligne), car openpyxl ne l'ajuste pas seul.
* Causes indépendantes : jours travaillés, taux (NCR) et `Cost issue` peuvent coexister (le `Cost issue` n'est plus un repli exclusif).
* Repli final `Cost issue (requires manual verification)` : seulement si un gap de Transfer Price existe au niveau projet sans consultant individuel qui l'explique.

### Feuille KPI Dashboard (export)
Feuille **"KPI Dashboard"** ajoutée automatiquement, générée par `_build_kpi_sheet` :
1. **Synthèse mensuelle** : Cost ICB, Forecast, Gap signé, Working Days, Raw Cost, Nb Projets, Nb Consultants, ligne TOTAL.
2. **Cumul par mois** : Cost ICB cumulé mois après mois (mois 1, mois 1+2, mois 1+2+3...), Forecast cumulé, Gap cumulé, % de progression.
3. **Synthèse par projet** : triée par Cost ICB décroissant (Cost, Forecast, Gap, Gap %, Working Days, NCR moyen, Nb consultants).
4. **Top 50 consultants** par Cost ICB.
5. **Détail consultant par mois** : tableau croisé, colonne TOTAL.
Palette cohérente avec le thème, formats numériques Excel natifs.

---

## 5. Historique détaillé des modifications

Regroupé par thème, dans l'ordre approximatif de réalisation.

### A. Intégration du fichier Forecast
1. **Ajout du bouton d'import Forecast**, chargement du tableau avec les valeurs Forecast : `Prod` → Working Days, `Cost` → Forecast du Bill Amount, `F.RC` → NCR. Mapping projet par `Project Num` = `CC Project Number`, consultants par `GGID` = `Employee Number`.
2. **Année implicite** : le fichier ne contient pas d'année, les mois sont rattachés à l'année système actuelle.
3. **Normalisation du préfixe `6427_`** retiré des deux côtés avant comparaison.
4. **Import limité à un seul par session** ; l'ICB principal n'est jamais rechargé.
5. **Règle de non-correspondance** : projets/consultants du Forecast sans équivalent ICB ignorés silencieusement. Le Forecast ne sert qu'à remplir les cellules mappées.
6. **Parsing rendu dynamique** après constat que le fichier de production avait une structure décalée (feuille `Forecast26`, en-têtes pas en ligne 6, 104 colonnes, pas de cellules fusionnées). Détection par nom de la ligne d'en-têtes, des colonnes d'identité, des blocs mensuels et des blocs projets, avec fallback par rupture de valeur.
7. **Logs de diagnostic** ajoutés pour analyser les fichiers de production ; faux positif « colonnes détectées 16383 vs 87 » corrigé (le contrôle se base maintenant sur le contenu réel et non sur `ws.max_column`).

### B. Fiabilité du chargement
8. **WinError 32 sous Windows** : l'upload Forecast passait par un fichier temporaire système dont le handle restait ouvert. Alignement sur le mécanisme de l'ICB : sauvegarde dans `uploads/current_forecast.xlsx`, lecture par chemin, `wb.close()` dans un `finally`, plus aucune suppression de fichier temporaire.
9. **IndentationError** introduite par une modification de l'agent dans `icb_app.py` : corrigée.
10. **Nouvelle structure du fichier ICB** : classeur multi-feuilles, feuille `ICB_DB`, 5 premières lignes et colonne A vides → sélection explicite de la feuille et détection dynamique de la ligne d'en-têtes et de la colonne de départ.
11. **Détection de la feuille Forecast principale** : `wb.active` pointait sur la feuille d'appariement EM après son ajout et annulait tout l'import. Remplacé par un parcours de tous les onglets avec exclusion de la feuille EM.

### C. Colonne Bill Amount → Cost → Transfer Price
12. **Remplacement de la source ICB** : la colonne Bill Amount ICB est remplacée par `Project Transfer Price`. Première hypothèse erronée (`Revenue Project TP`, col AF) → toutes les valeurs à 0 ; corrigé par détection dynamique par nom exact.
13. **Trois noms possibles** pour cette colonne : `Project Transfer Price`, `Project Revenue Price`, `Revenue Project TP`. Erreur explicite si aucun n'existe.
14. **Renommages** : "Bill Amount" → "Cost" → **"Transfer Price"** (en-tête de la grande colonne dans l'app) ; export : `ICB Transfer Price` et `Forecast Transfer Price`. Les noms internes de variables restent `bill_amount`.

### D. Écarts, surlignage et seuils
14. **Surlignage rouge** des écarts ICB vs FORECAST (NCR, Working Days, Transfer Price).
15. **Seuil de tolérance relevé** (de 0.01 à 2) pour ignorer les écarts de quelques centimes, puis rendu **configurable depuis l'UI** (champ de saisie, défaut 2). Plus aucune constante en dur.
16. **Working Days** : plus aucune tolérance, comparaison stricte.
17. **Aucun surlignage avant le chargement de l'ICB.**
18. **Forecast = source de vérité** : seule la cellule ICB est surlignée en rouge (plus les deux).
19. **Gap signé** dans l'UI et l'export.
20. **NCR projet** : moyenne simple au lieu de la somme.

### E. Export Excel
21. **Colonnes Forecast Bill Amount / Gap / Comment remplies automatiquement** à l'export (auparavant vides pour saisie manuelle).
22. **Export en vrai tableau Excel** avec filtres et bandes alternées ; **nouvelles colonnes** `EM_Lead` (début) et `Issue Type` (avant le commentaire), vides ; "Comment" renommée `Generated comment`.
23. **Commentaires** : détail nominatif (code + nom) ; une cause par ligne dans la cellule ; sans troncature ; écrits en anglais ; suffixe ` D` pour les jours ; regroupés par consultant ; ajout de `Cost issue` avec détail nominatif.
24. **Inversion de la logique d'analyse** : on part de l'ICB et non du Forecast ; suppression des commentaires « absent de l'ICB » / « absent du Forecast ».
25. **Dialogue d'export** : d'abord une case pour inclure/exclure les commentaires, puis liste de cases (`Working Days`, `Raw Cost`, `Generated comment`, `Controlling comment`) avec les défauts décrits plus haut. `Controlling comment` est une nouvelle colonne vide de saisie manuelle.
26. **Raw Cost** retiré de l'affichage de l'app, conservé (optionnel) dans l'export.
27. **Feuille KPI Dashboard** ajoutée à l'export.

### F. Données : union, zéros et quittés
28. **Union au lieu d'intersection** des expenditure items ICB/Forecast (l'ICB généré par la plateforme source omet parfois des items pourtant pointés), avec indicateur discret dans le tableau.
29. **Consultants quittés** : absent de l'ICB + toutes valeurs Forecast à 0 → non affiché (voir section 4).
30. **Cas ICB = 0 / Forecast = 5** absent des commentaires : cause = tests de vérité implicites (`if wd_icb and wd_fc and ...`, `if ncr_icb and ncr_fc and ...`) qui traitaient 0 comme « pas de donnée ». Corrigés par comparaison explicite (`!=` / `abs(...) > seuil`), et assemblage des causes rendu indépendant.

### G. Colonne EM Name
31. **Première version** : feuille `Project-EM` (Project Code, Project Name, EM Name), colonne `EM Name` à côté du projet, filtre déroulant combinable avec la recherche.
32. **Remplacement par la feuille `EDM`** : `Project Num` = colonne C (la colonne A, homonyme, est ignorée), EM = colonne `Engagement Manager` (colonne E). Détection dynamique de la ligne d'en-têtes.
33. **Correction d'une anomalie UI au scroll** : l'en-tête de la colonne avait `position: relative` (style inline + classe) qui écrasait `position: sticky`, et un `<tr>` ouvrant manquait dans le `<thead>`. Corrigé (sticky rétabli, `z-index` supérieur, fermeture du menu au scroll).

### H. Ergonomie et nettoyage
34. **Recherche** : liste les projets correspondants sans les ouvrir.
35. **Boutons de chargement** : disparaissent après un chargement réussi, indépendamment.
36. **Traduction de toute l'application en anglais** (UI, messages, dialogues, logs, commentaires), sans toucher aux mots-clés de matching.
37. **Nettoyage** : suppression des logs de debug temporaires et du code mort ; conservation du bloc `[FORECAST IMPORT]`.
38. **Réorganisation du dépôt** : `docs/` pour la documentation, `.gitignore` (`*.exe`, `dist/`, `build/`), suppression de l'exécutable versionné.

---

## 6. Bugs résolus et leçons

| Symptôme | Cause réelle | Correction |
|---|---|---|
| `WinError 32` à l'upload Forecast (Windows) | handle du fichier temporaire non libéré | fichier dans `uploads/`, lecture par chemin, `wb.close()` en `finally` |
| Certains consultants n'affichent pas leurs expenditure items au clic | JSON invalide : `NaN` littéral (ex. `"invoice_type": NaN`) → `JSON.parse` plante côté navigateur | nettoyage des `NaN`/`Infinity` en `None` à la source et avant sérialisation |
| Hypothèses écartées pendant ce diagnostic | type `int`/`float` d'`Employee Number`, parité de lignes (`alt`), collisions d'ID | les logs backend montraient toujours des items non vides ; la cause était dans la réponse HTTP elle-même |
| Toutes les valeurs Transfer Price à 0 | mauvaise colonne supposée (AF / `Revenue Project TP`) | détection par nom exact + erreur explicite |
| Import Forecast annulé après ajout de la feuille EM | `wb.active` pointait sur la feuille EM | parcours de tous les onglets, exclusion de la feuille EM |
| Cas ICB=0 / Forecast>0 absent des commentaires | `if valeur:` traite 0 comme absent | comparaisons explicites |
| `Cost issue` masqué dès qu'un autre consultant avait un gap WD/NCR | repli exclusif (`if not parts`) | causes indépendantes |
| En-tête `EM` qui défile avec le corps du tableau | `position: relative` écrase `sticky` + `<tr>` manquant | CSS et HTML corrigés |

**Leçon générale** : ne jamais deviner un nom de colonne ou une structure de fichier ; vérifier sur le fichier réel (ou via les logs de production), et toujours loguer clairement quand une colonne est introuvable plutôt que de retomber silencieusement sur 0.

---

## 7. Notes pour le prochain Agent / Développeur
* **Structure In-Memory** : le fichier chargé est stocké dans la variable globale `pipeline` (`icb_app.py`), adapté à un usage Desktop mono-utilisateur.
* **`_normalize_key`** : fonction centrale. Un `CC Project Number` peut être lu en `int`, `float` (`123.0`) ou `str` ; la normalisation évite les plantages de tri et les non-correspondances.
* **Préfixe `6427_`** : retiré des deux côtés (`.replace("6427_", "")`). À adapter si la production utilise un autre préfixe.
* **Dates Excel** : `excel_serial_to_date` est essentielle (numéros de série lus en float).
* **`NaN` et JSON strict** : tout champ passe par des filtres (`_clean()`, `_to_float()`) ; ne jamais renvoyer un `float('nan')` à `jsonify`.
* **Tests de vérité** : n'utiliser jamais `if valeur:` pour une valeur numérique pouvant légitimement valoir 0 ; comparer explicitement avec `None`.
* **Naming interne vs affichage** : `bill_amount`, `total_bill_amount`, `forecast_bill_amount` gardent leur nom historique ; seuls les libellés affichés disent « Transfer Price ».
* **Pywebview** : les dialogues natifs exigent une fenêtre accessible et gérée sur le bon thread.
* **Double ouverture du Forecast** : une ouverture normale (`data_only=True`) pour les fusions et en-têtes, puis `read_only=True` pour la lecture rapide (les cellules fusionnées ne sont pas accessibles en `read_only`).
* **Cellules fusionnées** : seule la cellule d'ancrage porte la valeur ; les autres renvoient `None`. Ne jamais supposer `Project Num` présent sur chaque ligne.
* **Feuilles Forecast** : jamais `wb.active` ; détecter la feuille principale par marqueurs en excluant `EDM`.
* **Colonnes `EDM`** : détection dynamique par nom (`Project Num` et `Engagement Manager`).
* **Colonnes optionnelles de l'export** : la formule `Gap` doit référencer dynamiquement les colonnes `ICB Transfer Price` et `Forecast Transfer Price`, car leurs lettres changent selon les cases cochées.
* **Seuil de tolérance** : valeur unique pilotée par l'UI et transmise au backend ; ne jamais réintroduire de constante en dur (risque de désynchronisation surlignage / export).
* **Commentaires multi-lignes** : `\n` + `wrap_text` + hauteur de ligne explicite (openpyxl ne calcule pas la hauteur).
* **Fichier de référence** : `docs/structureDetailleICBandForecast.txt` décrit en détail la structure des fichiers ICB et Forecast.
* **Logs de production** : le bloc `[FORECAST IMPORT]` est volontairement conservé pour diagnostiquer les futurs fichiers.

---



