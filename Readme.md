# Add-in Elyse Energy

## Description
Add-in Excel pour la gestion et l'analyse des données énergétiques d'Elyse Energy. Cet add-in fournit une interface unifiée pour accéder et manipuler les données techniques, financières et de projet, avec une architecture robuste incluant la gestion des erreurs, le logging, et l'intégration SharePoint.

## Guide de Démarrage Rapide

### Installation pour Utilisateurs
1. Copier le fichier `Addin Elyse Energy.xlam` dans le dossier des add-ins Excel
   - Windows : `C:\Users\[Username]\AppData\Roaming\Microsoft\AddIns`
2. Ouvrir Excel et aller dans Fichier > Options > Add-ins
3. Cliquer sur "Gérer : Add-ins Excel" puis "Atteindre..."
4. Cocher "Addin Elyse Energy" et cliquer sur OK
5. Le ruban "Elyse Energy" apparaîtra dans l'interface Excel

### Installation pour Développeurs
1. Cloner le repository
2. Ouvrir le fichier `Addin Elyse Energy.xlsm` dans Excel
3. Configurer les variables d'environnement dans `CFG_04_EnvironmentVariables.bas`
4. Exécuter la procédure `InitializeElyseSystem` dans `APP_MainOrchestrator.bas`
5. Pour déboguer, utiliser le mode DEBUG_MODE :
   ```vb
   InitializeElyseSystem DEBUG_MODE
   ```

## Structure du Projet

### Modules Système
- `elyse_main_orchestrator.bas` : Module principal de coordination du système
- `elyse_error_handler.bas` : Gestion des erreurs
- `elyse_logger_module.bas` : Système de journalisation
- `SYS_CoreSystem.bas` : Fonctionnalités système de base
- `SYS_MessageBox.bas` : Système de boîtes de dialogue
- `SYS_TicketSystem.bas` : Gestion des tickets de support
- `SYS_SharePointIntegration.bas` : Intégration SharePoint

### Modules de Configuration
- `CFG_01_CategoryDefinitions.bas` : Définitions des catégories de données
- `CFG_02_RagicFieldDictionary.bas` : Dictionnaire des champs Ragic
- `CFG_03_CustomDataTypes.bas` : Types de données personnalisés
- `CFG_04_EnvironmentVariables.bas` : Variables d'environnement

### Modules de Données
- `DAT_01_DataLoadManager.bas` : Gestionnaire de chargement des données
- `DAT_02_QueryLoader.bas` : Chargement des requêtes
- `DAT_03_PowerQueryManager.bas` : Gestion des requêtes Power Query
- `DAT_04_PowerQueryDebug.bas` : Débogage Power Query

### Modules Interface Utilisateur
- `UI_01_RibbonDefinition.xml` : Définition du ruban Excel
- `UI_02_RibbonLogic.bas` : Logique du ruban
- `UI_03_RibbonActions.bas` : Actions du ruban

### Modules de Sécurité
- `SEC_01_AccessProfiles.bas` : Gestion des profils d'accès

### Utilitaires
- `UTL_01_GeneralUtilities.bas` : Utilitaires généraux

## Catégories de Données

### Technologies
- Compression
- CO2 Capture
- H2 Waters Electrolysis
- MeOH Synthesis
- SAF Synthesis

### Utilitaires
- Chiller
- Cooling Water Production
- Heat Production
- Water Treatment
- Waste Water Treatment

### Métriques
- Métriques de base
- Métriques expert
- Timings de référence
- Métriques RED III
- Emissions

### Finances
- Budget Corpo
- Détails Budgets
- DIB
- Demandes d'achat
- Réceptions

### Projets
- Scénarios techniques
- Plannings
- Budget Projet
- Devex
- Capex

## Installation

1. Copier le fichier `Addin Elyse Energy.xlam` dans le dossier des add-ins Excel
2. Activer l'add-in dans Excel via Fichier > Options > Add-ins
3. Le ruban "Elyse Energy" apparaîtra dans l'interface Excel

## Développement

### Prérequis
- Excel 2016 ou supérieur
- Visual Basic Editor (VBE)
- Accès aux sources de données Ragic

### Scripts Utilitaires
- `ConvertToXLAM.vbs` : Conversion du projet en add-in
- `Inject XML (not tested).vbs` : Injection du XML du ruban

## Maintenance

### Logs
Les logs sont gérés par le module `elyse_logger_module.bas` et peuvent être consultés via :
- Interface utilisateur dédiée
- Fichiers de log dans le dossier de l'application

### Gestion des Erreurs
Le système de gestion des erreurs est centralisé via `elyse_error_handler.bas` et permet :
- Capture des erreurs
- Journalisation
- Création de tickets de support automatiques

## Support
Pour toute question ou problème, veuillez utiliser le système de tickets intégré via le ruban "Elyse Energy".
