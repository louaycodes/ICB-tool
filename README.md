# ICB Tool — Intercompany Billing Analysis

Application desktop pour charger, analyser, filtrer et valider les heures et les coûts des consultants facturés en Intercompany Billing (ICB).

## 🚀 Fonctionnalités

- **Chargement optimisé** : Lecture ultra-rapide des fichiers Excel volumineux (évite le bug des feuilles Excel à dimensions gonflées).
- **Tableau hiérarchique** : Visualisation claire et structurée par `Project Number` ➔ `CC Project Number` ➔ `Consultant`. Les lignes de projets sont pliables/dépliables.
- **Détail en ligne (Inline Toggle)** : Affichage instantané du détail des transactions d'un consultant directement sous sa ligne d'un simple clic (et se referme en recliquant).
- **Vérification automatique (FinOps)** : Validation automatique de la règle `Raw Cost = 0.75 * Bill Amount` (pour les heures standards `RC_Time Std`).
- **Détection des anomalies** : Mise en surbrillance automatique en **rouge** des consultants et des transactions individuelles présentant un écart de plus de **2$** (ex: problèmes de devises ou de conversion).
- **Recherche temps réel** : Filtrage instantané du tableau principal par nom, matricule, grade, ou numéro de projet.
- **Filtres temporels** : Sélecteur de mois toujours visible au-dessus du tableau pour naviguer d'une période à l'autre en un clic.

---

## 🛠️ Installation et Lancement

### 1. Prérequis
Assurez-vous d'avoir Python 3 installé sur votre machine.

### 2. Cloner le projet
```bash
git clone https://github.com/louaycodes/ICB-tool.git
cd ICB-tool
```

### 3. Lancement rapide (un seul clic)
```bash
./launch.sh
```
Ce script crée automatiquement l'environnement virtuel, installe les dépendances, et lance l'application dans une fenêtre native.

### 4. Lancement manuel
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python3 icb_desktop.py
```

### 5. Utiliser l'application
Une fenêtre native s'ouvre automatiquement. Cliquez sur **Charger un fichier ICB** et sélectionnez votre fichier `.xlsx`.

---

## 📁 Structure du projet

- `icb_desktop.py` : Point d'entrée principal — lance l'application dans une fenêtre desktop native.
- `icb_app.py` : Backend Flask et pipeline de traitement des données Excel.
- `launch.sh` : Script de lancement rapide (un clic).
- `requirements.txt` : Dépendances Python du projet.
- `templates/index.html` : Interface utilisateur HTML5 et logique applicative JavaScript.
- `static/style.css` : Design system et styles CSS professionnels (police Inter, thème corporate).
- `.gitignore` : Exclusion des fichiers lourds/sensibles (fichiers `.xlsx`, environnements virtuels `venv`).

