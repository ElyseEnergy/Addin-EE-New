# Add-in Elyse Energy

## Description


Add-in Excel pour la gestion et l'analyse des données énergétiques d'Elyse Energy. Cet add-in fournit une interface unifiée pour accéder et manipuler les données techniques, financières et de projet, avec une architecture robuste incluant la gestion des erreurs, le logging, et l'intégration SharePoint.

## Prérequis et Dépendances

### Système

- Windows 10/11
- Excel 2016 ou supérieur
- Visual Basic for Applications (VBA)
- Droits d'administration pour l'installation

### Accès et Authentification

- Compte Elyse Energy actif
- Accès aux tableaux Ragic
- Accès au SharePoint d'entreprise
- Permissions VBA activées dans Excel

### Configuration Système Excel

1. Activer le mode développeur dans Excel
   - Fichier > Options > Personnaliser le ruban
   - Cocher "Développeur" dans la liste à droite

2. Configurer les paramètres de sécurité
   - Onglet Développeur > Sécurité des macros
   - Activer "Faire confiance à l'accès au modèle d'objet des projets VBA"
   - Autoriser les macros pour ce document

3. Installer les références requises   - Microsoft Scripting Runtime
   - Microsoft XML
   - Microsoft Forms 2.0 Object Library
   - Microsoft Office Object Library
   - Microsoft SharePoint Type Library (IMPORTANT : sélectionner "Microsoft SharePoint Type Library", PAS "Microsoft SharePoint Object Library" ni "Microsoft SharePoint Plugin")

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

## Architecture et Implémentation

### Composants Système Principal

1. **Application Principale (`APP_MainOrchestrator.bas`)**
   - Point d'entrée du système
   - Coordination des modules
   - Interface publique unifiée
   - Gestion du cycle de vie

2. **Système Core (`SYS_CoreSystem.bas`)**
   - Configuration de base
   - Gestion des sessions
   - Constantes et énumérations
   - Utilitaires système

3. **Gestion des Erreurs (`SYS_ErrorHandler.bas`)**
   - Capture d'erreurs centralisée
   - Création automatique de tickets
   - Recovery intelligent
   - Journalisation des erreurs

4. **Système de Logs (`SYS_Logger.bas`)**
   - Logs multiniveaux
   - Stockage Ragic
   - Buffer avec auto-flush
   - Métriques système

5. **Gestion des Messages (`SYS_MessageBox.bas`)**
   - Interface utilisateur
   - Messages contextuels
   - Support multilingue
   - Styles personnalisés

6. **SharePoint (`SYS_SharePointIntegration.bas`)**
   - Synchronisation des fichiers
   - Métadonnées SharePoint
   - Gestion des versions
   - Cache local

7. **Support Utilisateur (`SYS_TicketSystem.bas`)**
   - Création de tickets
   - Suivi des incidents
   - Base de connaissances
   - Assistance automatisée

### Composants Interface Utilisateur

1. **Ruban Excel**
   - `UI_01_RibbonDefinition.xml` : Structure du ruban
   - `UI_02_RibbonLogic.bas` : Logique des contrôles
   - `UI_03_RibbonActions.bas` : Actions des boutons

2. **Formulaires**
   - `frmCustomMessageBox_vbacode.bas` : Messages personnalisés
   - `frmListSelection_vbacode.bas` : Sélection de listes
   - `frmMarkdownDisplay_vbacode.bas` : Affichage Markdown
   - `frmRangeSelector_vbacode.bas` : Sélection de plages
   - `frmTicketInput_vbacode.bas` : Création de tickets

### Composants Fonctionnels

1. **Configuration**
   - `CFG_01_CategoryDefinitions.bas` : Catégories de données
   - `CFG_02_RagicFieldDictionary.bas` : Mapping Ragic
   - `CFG_03_CustomDataTypes.bas` : Types personnalisés
   - `CFG_04_EnvironmentVariables.bas` : Variables d'environnement

2. **Gestion des Données**
   - `DAT_01_DataLoadManager.bas` : Chargement des données
   - `DAT_02_QueryLoader.bas` : Gestion des requêtes
   - `DAT_03_PowerQueryManager.bas` : Power Query
   - `DAT_04_PowerQueryDebug.bas` : Débogage

3. **Sécurité et Utilitaires**
   - `SEC_01_AccessProfiles.bas` : Profils d'accès
   - `UTL_01_GeneralUtilities.bas` : Fonctions utilitaires

### Guide d'Implémentation

1. **Initialisation du Système**

```vb
' Dans un nouveau module
Public Sub InitializeElyseSystem()
    ' Initialisation en mode DEBUG
    If Not APP_MainOrchestrator.InitializeElyseSystem(DEBUG_MODE) Then
        Debug.Print "Échec de l'initialisation"
        Exit Sub
    End If
    
    ' Vérification du statut
    If SYS_CoreSystem.IsSystemInitialized() Then
        Debug.Print "Système initialisé avec succès"
    End If
End Sub
```

2. **Gestion des Erreurs et Logging**

```vb
' Exemple de procédure avec gestion d'erreurs
Public Sub ExempleProcedure()
    On Error GoTo ErrorHandler
    
    ' Log de début
    SYS_Logger.LogEvent "procedure_start", "Démarrage de ExempleProcedure", INFO_LEVEL
    
    ' Votre code ici...
    
    ' Log de fin
    SYS_Logger.LogEvent "procedure_end", "Fin de ExempleProcedure", INFO_LEVEL
    Exit Sub

ErrorHandler:
    ' Gestion de l'erreur
    SYS_ErrorHandler.HandleError MODULE_NAME, "ExempleProcedure"
End Sub
```

3. **Interactions Utilisateur**

```vb
' Exemple d'utilisation des formulaires
Public Sub AfficherMessage()
    ' Message personnalisé
    SYS_MessageBox.ShowCustomMessage _
        "Titre", _
        "Message avec formatage Markdown", _
        "OK", _
        INFO_MESSAGE
    
    ' Création d'un ticket
    If SYS_TicketSystem.CreateSupportTicket("Sujet", "Description") Then
        Debug.Print "Ticket créé avec succès"
    End If
End Sub
```

### Meilleures Pratiques

1. **Structure des Modules**
   - Déclarer `Option Explicit`
   - Documenter les dépendances en en-tête
   - Grouper les procédures par fonctionnalité
   - Utiliser des régions avec commentaires

2. **Gestion des Erreurs**
   - Toujours utiliser `On Error GoTo ErrorHandler`
   - Logger les erreurs via l'orchestrateur
   - Nettoyer les ressources dans les blocs d'erreur
   - Utiliser les contextes d'erreur pour le diagnostic

3. **Performance**
   - Désactiver les mises à jour d'écran pour les opérations longues
   - Utiliser le buffer de logs judicieusement
   - Éviter les appels inutiles à l'API Ragic
   - Optimiser les boucles et les collections

## Maintenance

### Logs System

Les logs sont stockés à plusieurs niveaux :

- Ragic (logs système et erreurs)
- Fichiers locaux (debug et diagnostics)
- Buffer mémoire (performance)

Configuration des logs :

```vb
' Définir le niveau de log
ElyseMain_Orchestrator.SetLogLevel DEBUG_LEVEL  ' Pour le développement
ElyseMain_Orchestrator.SetLogLevel INFO_LEVEL   ' Pour la production
```

### Gestion des Erreurs

Le système comprend :

- Capture automatique des erreurs
- Recovery intelligent
- Messages utilisateur contextuels
- Création de tickets support

### Mise à Jour

Pour mettre à jour l'add-in :

1. Modifier le code dans la version `.xlsm`
2. Tester en mode DEBUG
3. Exécuter `ConvertToXLAM.vbs`
4. Distribuer la nouvelle version `.xlam`

### Support et Maintenance

- Tickets automatiques via Ragic
- Logs détaillés pour le diagnostic
- Système de santé intégré
- Métriques de performance

## Support

Pour toute question ou problème, utiliser le système de tickets intégré via le ruban "Elyse Energy".
