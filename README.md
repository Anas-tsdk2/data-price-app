# Data Price App - Guide Complet

## 1. Présentation générale

### Objectif de l'application
Cette application résout un problème courant en gestion IT : la répartition des coûts des serveurs entre différentes applications. Quand un serveur héberge plusieurs applications, comment répartir équitablement son coût ?

### Principe de fonctionnement
- L'application prend en entrée un fichier contenant la liste des serveurs
- Pour chaque serveur ayant plusieurs applications (tags), elle crée une ligne par application
- Elle calcule automatiquement un coefficient de répartition des coûts

### Exemple concret
Si un serveur "SRV" héberge 3 applications (P1, P2, P3) :
- L'application créera 3 lignes distinctes
- Chaque ligne aura un coefficient de 0.333333 (1/3)
- La somme des coefficients égale toujours 1 pour garantir une répartition à 100%

## 2. Utilisation pas à pas

### Étape 1 : Préparation du fichier
Format accepté :
- Excel (.xlsx, .xls)
- CSV (séparateur point-virgule)

### Étape 2 : Import du fichier
Deux méthodes :
- Glisser-déposer le fichier dans la zone prévue
- Cliquer sur la zone pour sélectionner le fichier

### Étape 3 : Traitement
1. Cliquer sur "Traiter"
2. L'application va :
   - Séparer les tags multiples
   - Calculer les coefficients
   - Valider les données
   - Afficher les statistiques

### Étape 4 : Validation
L'application vérifie automatiquement :
- Le nombre total de serveurs d'origine
- La somme des coefficients
- Le nombre de serveurs sans tags
- L'intégrité des données

### Étape 5 : Export
- Cliquer sur "Exporter"
- Le fichier généré conserve le format d'origine
- Les nouvelles colonnes sont ajoutées :
  ```
  App tags (CT)      : Tags traités
  Nbr App Tags       : Nombre de tags par serveur
  Coef AppTags       : Coefficient calculé
  Split App tags     : Tag individuel
  ```

## 3. Cas particuliers

### Serveur sans tags
- Une seule ligne est créée
- "Split App tags (raw)" = "No App Tags"
- Coefficient = 1.000000000

### Tags dans App tags (CT)
Si App tags (raw) est vide mais App tags (CT) contient des valeurs :
- Les tags de App tags (CT) sont utilisés
- Le traitement reste identique

## 4. Aspects techniques

### Traitement des données
- Séparation des tags par virgules
- Nettoyage des espaces
- Calcul précis des coefficients (9 décimales)
- Validation mathématique des résultats

### Format Excel
- Conservation des filtres automatiques
- Mise en forme conditionnelle
- En-têtes en couleur
- Alternance de couleurs pour les lignes

### Format CSV
- Séparateur point-virgule
- Encodage UTF-8
- Protection des valeurs avec guillemets

## 5. Messages d'erreur courants

- "Veuillez sélectionner un fichier" : Aucun fichier sélectionné
- "Erreur de validation" : La somme des coefficients ne correspond pas
- "Erreur lors du traitement" : Format de fichier incorrect

## 6. Bonnes pratiques

1. Vérifier le format des données avant import
2. Attendre la fin du traitement avant export
3. Contrôler les statistiques de validation
4. Sauvegarder le fichier d'origine

## 7. Support et maintenance

Développé par : Anastasia Tsundyk
Version : 2025
Technologies : HTML5, CSS3, JavaScript, ExcelJS, SheetJS
