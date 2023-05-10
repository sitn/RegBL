# RegBL
Boîte à outils pour la gestion du RegBL

## Prérequis
* Python 3.10 ou ultérieur + git version 2.37.3 ou ultérieur
* Dossier git initialisé pour le projet
```
git init
```
* Projet récupéré de github et à jour
```
git clone https://github.com/sitn/RegBL.git
```
* Environnement virtuel dans le dossier du projet
```
python -m venv venv
```
* Copier le fichier .env.sample et renseigner les variables d'environnement
```
cp .env.sample .env
```


## Installation
Dans un terminal avec l'environnement virtuel activé:
```
./venv/scripts/activate

(venv): pip install -r requirements.txt
```

## Pour générer les fichiers feedback des communes
```
(venv): cd ./apurement/feedback_communes
(venv): python main.py
```

## Pour compléter les listes d'extension des communes
Déposer les listes excel dans `\regbl_toolbox\extension\traitement_communes\input`

Les listes complétées seront déposées dans `regbl_toolbox\extension\traitement_communes\output`
```
(venv): cd ./extension/traitement_communes
(venv): python main.py
```

## Pour générer des fichiers KML
Déposer les listes excel dans `\regbl_toolbox\extension\kml_generator\input`

Les fichiers KML seront déposées dans `regbl_toolbox\extension\kml_generator\output`
```
(venv): cd ./extension/kml_generator
(venv): python main.py
```
## Pour générer les fichiers KML des rapports des communes
Déposer les listes excel dans `\regbl_toolbox\extension\analyse_rapport\rapports`

Les fichiers KML seront déposées dans le même dossier.
```
(venv): cd ./extension/analyse_rapport
(venv): python main.py
```