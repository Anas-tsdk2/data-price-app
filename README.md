# Data Price App

Application web permettant de traiter des fichiers Excel et CSV pour le calcul de prix, avec un focus particulier sur le traitement des tags d'applications et la déduplication des données.

## Fonctionnalités principales

### Traitement des tags d'applications
- Séparation automatique des tags multiples
- Calcul du nombre total de tags par serveur
- Génération d'un coefficient par tag (1/nombre total de tags)
- Conservation des données d'origine dans une colonne "App tags (raw)"

### Déduplication intelligente
- Création d'une ligne distincte pour chaque tag d'application
- Conservation de toutes les informations du serveur (nom, statut, environnement, etc.)
- Maintien de la traçabilité avec les données sources

### Exemple de transformation
Pour un serveur comme avec les tags "a,b,c" :
- Création de 3 lignes distinctes
- Calcul du coefficient (0.333333 = 1/3)
- Conservation des métadonnées (statut, environnement, localisation, etc.)

## Technologies utilisées

- HTML5
- CSS3
- JavaScript
- ExcelJS pour le traitement des fichiers Excel
- SheetJS pour la manipulation des données tabulaires

## Installation et utilisation

1. Clonez le repository
2. Ouvrez index.html dans votre navigateur
3. Importez votre fichier Excel ou CSV via l'interface drag & drop
4. Le traitement se fait automatiquement
5. Téléchargez le fichier traité dans son format d'origine

## Format des données

### Colonnes d'entrée requises
- Server name
- Server status
- Environment
- Administration company
- Location
- App tags (raw)
- Service criticality

### Colonnes générées
- App tags (CT)
- Nbr App Tags
- Coef AppTags
- Split App tags (raw)

## Auteur

Anastasia Tsundyk - 2025
