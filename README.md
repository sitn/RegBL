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