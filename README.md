# CSP Registry Mapper

## Description
Outil PowerShell pour analyser et mapper les chemins de Configuration Service Provider (CSP) de Mobile Device Management (MDM) Windows avec leurs entrées correspondantes dans le registre Windows. Cet utilitaire facilite le dépannage et la vérification des politiques MDM en identifiant où et comment elles sont stockées dans le registre.

## Fonctionnalités
- Parse les chemins CSP de MDM Windows (format "./Device/Vendor/MSFT/Policy/...")
- Recherche automatiquement les entrées de registre correspondantes
- Supporte la correspondance exacte et par préfixe
- Prend en compte les différentes structures de stockage dans le registre Windows
- Génère un rapport Excel détaillé des correspondances trouvées

## Prérequis
- Windows PowerShell 5.1 ou supérieur
- Module PowerShell ImportExcel (installé automatiquement si nécessaire)
- Droits d'administrateur pour accéder à certaines zones du registre

## Installation
Aucune installation nécessaire. Téléchargez simplement le script et exécutez-le avec PowerShell en tant qu'administrateur.

## Utilisation
```powershell
.\MDM-CSP-Registry-Analyzer.ps1
```

Le script vous demandera interactivement :
1. Le chemin du fichier Excel contenant les chemins CSP
2. Le nom de la feuille Excel (optionnel)
3. Le nom de la colonne contenant les chemins CSP (optionnel - détection automatique disponible)

## Format d'entrée
Fichier Excel avec une colonne contenant des chemins CSP au format :
- `./Device/Vendor/MSFT/Policy/Config/{area}/{setting}`
- `Device/Vendor/MSFT/Policy/Config/{area}/{setting}`

## Sortie
Un fichier Excel est généré avec les correspondances entre les chemins CSP et les entrées du registre, incluant :
- Le chemin CSP original
- Le paramètre extrait
- Le chemin dans le registre
- Si l'entrée existe
- Le nom réel et la valeur de la politique

## Utilisation avancée
Le script peut être intégré dans d'autres outils PowerShell en appelant directement la fonction :
```powershell
Analyze-CSPPaths -ExcelFilePath "C:\Chemin\vers\fichier.xlsx" -WorksheetName "Feuil1" -CSPColumnName "CheminCSP"
```

## Limites
- Recherche limitée à une profondeur maximale de 2 niveaux dans le registre
- Certains paramètres CSP peuvent utiliser des structures de stockage particulières non couvertes

## Licence
GNU General Public License v3 (GPL-3.0)

## Contributions
Les contributions sont les bienvenues via pull requests. Veuillez documenter clairement les modifications apportées.
