# Envoi de mails automatisés pour le piaulage de la maison des mines

## Installation 
Le lancement de ce programme requiert Python (version 2.6 ou plus) et pip (utiliser WSL pour travailler sous environnement virtuel avec Linux)

Tout d'abord créer un environnement virtuel pour les installations : "*virtualenv .venv*" et l'activer "*source .venv/bin/activate*".
Faire les installations nécessaires avec : "*pip install -r requirements.txt*"
Se connecter sur https://console.cloud.google.com et créer un nouveau projet.
Dans l'onglet "API et services" (sur la gauche), aller à "API et services activés" et cliquer "+ activer les API et les services" et ajouter Google Sheets et Gmail.
Dans l'onglet "API et services" (sur la gauche), aller à "Identifiants" et cliquer "+ créer des identifiants"->ID client OAuth.
Choisir "Application de bureau" (si demandé indiquer l'adresse mail dans les testeurs)
Télécharger l'ID client sous format JSON, renommer le fichier et le placer à un endroit sûr (pas sur le repo git, ni à un endroit public)
Dans le fichier .env du repo, remplacer PATH_TO_JSON_CREDENTIALS par le chemin du JSON

## Configuration
Modifier SPREADSHEET_ID dans le fichier .venv avec l'identifiant du tableau (docs.google.com/spreadsheets/d/**17rMaQTu7CjJ-I6p7jxXAB1Nf_ijkRMW2meNj7UHDMjk**/edit#gid=0) et RANGE si nécessaire (nom des colonnes à extraire du fichier)
Modifier EMAIL_ADDRESS si nécessaire
Modifier TEXTFILE_PATH : c'est le chemin relatif du fichier texte contenant le corps du mail avec les champs correspondant aux colonnes du fichier Excel à remplacer par les valeurs des cellules correspondantes
Modifier EMAIL_SUBJECT : c'est le sujet du mail à envoyer