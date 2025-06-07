# Guide de développement VBA – Addin Elyse Energy

Ce document sert de référence pour toutes les conventions de codage, les patterns d'architecture et les bonnes pratiques à suivre dans ce projet.

---

## 1. Définition et Usage des Types Personnalisés

### Centralisation des Types
La règle est de centraliser les types de données qui sont partagés ou qui représentent des objets métier fondamentaux dans un unique module : `Types.bas`.

**Exemple (`Types.bas`) :**
```vba
' Type pour les informations de catégorie, utilisé dans tout le projet
Public Type CategoryInfo
    CategoryName As String
    FilterLevel As String
    SecondaryFilterLevel As String
    DisplayName As String
    URL As String
    PowerQueryName As String
    CategoryGroup As String
    SheetName As String
End Type

' Type pour les informations de chargement de données
Public Type DataLoadInfo
    Category As CategoryInfo
    SelectedValues As Collection
    ModeTransposed As Boolean
    FinalDestination As Range
    PreviewRows As Long
End Type
```
**Exceptions :** Un type "helper" très spécifique qui n'est utilisé que par un seul module peut être défini au sein de ce module pour plus de clarté.
*Exemple : `FormattedCellOutput` dans `DataFormatter.bas`.*

### Manipulation des Types (Copie Profonde)
Lorsqu'une fonction retourne un tableau de types personnalisés, il est **impératif** d'effectuer une **copie profonde** pour éviter les problèmes de référence et garantir l'encapsulation. Ne jamais assigner directement le tableau interne.

**Exemple (`CategoryManager.bas`) :**
```vba
Public Function GetAllCategories() As CategoryInfo()
    ' ... (gestion d'erreur)
    
    ' Créer un nouveau tableau de la bonne taille
    Dim result() As CategoryInfo
    ReDim result(LBound(Categories) To UBound(Categories))
    
    ' Copier chaque élément champ par champ
    Dim i As Long
    For i = LBound(Categories) To UBound(Categories)
        With Categories(i)
            result(i).CategoryName = .CategoryName
            result(i).FilterLevel = .FilterLevel
            result(i).SecondaryFilterLevel = .SecondaryFilterLevel
            result(i).DisplayName = .DisplayName
            result(i).URL = .URL
            result(i).PowerQueryName = .PowerQueryName
            result(i).CategoryGroup = .CategoryGroup
            result(i).SheetName = .SheetName
        End With
    Next i
    
    GetAllCategories = result ' Assigner le nouveau tableau
    ' ...
End Function
```

---

## 2. Structure et Conventions des Modules

1.  **Modularité** : La logique est séparée par domaine fonctionnel.
    *   `CategoryManager` : Gère la définition et l'accès aux catégories.
    *   `DataLoaderManager` : Orchestre le processus de chargement des données.
    *   `PQQueryManager` : Gère la création et la maintenance des requêtes PowerQuery.
    *   `RibbonVisibility` : Gère l'état et la visibilité du ruban.
    *   `SYS_Logger` / `SYS_ErrorHandler` : Modules système pour le logging et les erreurs.

2.  **Constantes** : Les constantes partagées sont définies en haut du module le plus pertinent.
    *   Les variables d'environnement (clés API, URLs) sont dans `env.bas`.
    *   La version de l'addin est dans `Utilities.bas`.

---

## 3. Bonnes Pratiques Générales du Code VBA

### Structure d'un Module
Chaque module doit **obligatoirement** respecter la structure suivante pour la cohérence et la fiabilité, notamment pour l'import/export via des outils externes comme Git.

```vba
Attribute VB_Name = "NomDuModule"
Option Explicit

' --- Constantes et Enums du module ---
' --- Variables privées du module ---

' --- Procédures publiques ---
' --- Procédures privées ---
```

1.  **`Attribute VB_Name = "NomDuModule"`** : Toujours en **première ligne**. Ne jamais l'oublier, sinon l'import peut échouer ou créer un module au nom incorrect (ex: `Module1`).
2.  **`Option Explicit`** : Toujours en **deuxième ligne**. Force la déclaration de toutes les variables.

### Déclaration des Variables et Constantes
- **Déclaration en haut de procédure** : Toutes les variables locales doivent être déclarées au début de la fonction ou de la `Sub` pour une meilleure lisibilité.
- **Déclaration des constantes** : Une constante (`Const`) doit être initialisée avec une **valeur littérale** (ex: `123`, `"texte"`) ou une autre constante. **Les appels de fonction (comme `RGB()`) sont interdits** car ils sont évalués à l'exécution, et non à la compilation.

**Pattern correct pour les valeurs dynamiques :**
Pour une valeur qui nécessite un calcul, utilisez une fonction publique qui retourne la valeur.

**Exemple :**
```vba
' Interdit dans une déclaration Const :
Private Const FORBIDDEN_COLOR As Long = RGB(128, 128, 128)

' --- PATTERNS CORRECTS ---

' 1. Utiliser la valeur littérale si elle est connue et fixe :
Public Const CORRECT_COLOR As Long = 8421504 ' La valeur de RGB(128, 128, 128)

' 2. Ou, pour garder la lisibilité, utiliser une fonction :
Public Function GetMediumGrayColor() As Long
    GetMediumGrayColor = RGB(128, 128, 128)
End Function
```

### Conventions de Nommage
- **Procédures (`Sub`, `Function`)** : `PascalCase` (ex: `ProcessCategory`, `GetLastColumn`).
- **Variables locales** : `camelCase` (ex: `nextRow`, `sanitizedName`).
- **Variables de niveau module (privées)** : `m_camelCase` (ex: `m_currentProfile`).
- **Constantes** : `ALL_CAPS_WITH_UNDERSCORES` (ex: `PROC_NAME`, `LOG_SHEET_NAME`).
- **Enums et Types** : `PascalCase` (ex: `LogLevel`, `CategoryInfo`).

### Utilisation des `Enum`
Pour des ensembles de constantes liées, utilisez une `Enum` pour améliorer la lisibilité et bénéficier de l'IntelliSense.

**Exemple (`SYS_Logger.bas`) :**
```vba
Public Enum LogLevel
    DEBUG_LEVEL = 0
    INFO_LEVEL = 1
    WARNING_LEVEL = 2
    ERROR_LEVEL = 3
End Enum
```
Utilisez `LogLevel.INFO_LEVEL` plutôt que le chiffre `1`.

---

## 4. Gestion des Erreurs et des Logs

### Pattern de Gestion d'Erreur
Chaque fonction ou sub susceptible de planter doit implémenter le pattern suivant pour une gestion centralisée et robuste.

**Exemple (`n'importe quel module`) :**
```vba
Public Sub MaFonctionCritique()
    Const PROC_NAME As String = "MaFonctionCritique"
    Const MODULE_NAME As String = "MonModule"
    On Error GoTo ErrorHandler
    
    ' ... Code métier ...
    
    Exit Sub
ErrorHandler:
    ' Appel au gestionnaire centralisé
    HandleError MODULE_NAME, PROC_NAME, "Une erreur spécifique est survenue ici."
End Sub
```
La fonction `HandleError` (de `SYS_ErrorHandler.bas`) se charge de logger l'erreur et d'informer l'utilisateur si nécessaire.

### Système de Logging
Utiliser la fonction `Log` (de `SYS_Logger.bas`) pour tracer les événements importants.

**Exemple :**
```vba
Log "ribbon_load", "Le ruban a été chargé.", INFO_LEVEL, PROC_NAME, MODULE_NAME
```
*   `actionCode`: Un code court et unique pour filtrer les logs.
*   `message`: Le message descriptif.
*   `level`: `DEBUG_LEVEL`, `INFO_LEVEL`, `WARNING_LEVEL`, `ERROR_LEVEL`.
*   `procedureName`, `moduleName`: Pour le contexte (fournis par les constantes locales).

---

## 5. Callbacks du Ruban (customUI)

La logique des callbacks est séparée en deux catégories :

1.  **Visibilité et État de l'UI (`RibbonVisibility.bas`)**
    Toutes les callbacks qui gèrent l'apparence, la visibilité, les labels dynamiques ou l'état général de l'interface sont centralisées dans `RibbonVisibility.bas`.
    
    **Exemple (`RibbonVisibility.bas`) :**
    ```vba
    ' Callback pour la visibilité du groupe Technologies
    Public Sub GetTechnologiesVisibility(control As IRibbonControl, ByRef visible As Variant)
        ' Délègue la logique, ne la contient pas
        visible = HasAccess("Engineering")
    End Sub
    
    ' Callback pour changer de profil utilisateur
    Public Sub OnSelectDemoProfile(control As IRibbonControl)
        Select Case control.id
            Case "btnEngineerBasic": SetCurrentProfile AccessProfiles.Engineer_Basic
            ' ...
        End Select
        InvalidateRibbon ' Rafraîchit l'UI
    End Sub
    ```

2.  **Actions Métier (`*Manager.bas`)**
    Les callbacks `onAction` qui déclenchent un processus métier sont placées dans le module manager approprié. Elles servent de simple point d'entrée.

    **Exemple (`Technologies_Manager.bas`) :**
    ```xml
    <!-- customUI.xml -->
    <ns0:button id="btnCompression" label="Compression" onAction="ProcessCompression" />
    ```
    ```vba
    ' Technologies_Manager.bas
    Public Sub ProcessCompression(ByVal control As IRibbonControl)
        ' Délègue immédiatement toute la logique au manager compétent
        DataLoaderManager.ProcessCategory "Compression", "Erreur lors du traitement..."
    End Sub
    ```

---

## 6. Normalisation des Requêtes PowerQuery

La création et la gestion des requêtes PQ sont entièrement automatisées et normalisées via `PQQueryManager.bas` pour garantir la cohérence.

- **Nommage** : Le nom de la requête et de la table associée est généré via `Utilities.SanitizeTableName`.
- **Création/Mise à jour** : `PQQueryManager.EnsurePQQueryExists` vérifie si une requête existe, si sa formule (URL) a changé, et la crée ou la met à jour au besoin.
- **Template de Requête** : `GeneratePQQueryTemplate` crée le code M standard, qui inclut le typage de la colonne ID et sa mise en première position.
- **Chargement** : `LoadQueries.LoadQuery` est la seule fonction à utiliser pour charger une requête dans une feuille.

---

## 7. Le Dictionnaire de Données (Ragic Dictionary)

Le "Ragic Dictionary" est un mécanisme clé de l'addin. Il s'agit d'une table de correspondance qui fournit des **méta-informations sur les champs de données Ragic**, comme leur type de données réel ("DATE", "NUMBER", etc.) ou si un champ doit être masqué dans l'interface.

### Rôle et Utilité
Il permet de décorréler la logique de l'addin des données brutes. Par exemple, au lieu de coder en dur qu'un champ nommé "Date de début" doit être formaté comme une date, on consulte le dictionnaire pour connaître son type.

### Source et Mise en Cache
- **Source** : Les données proviennent d'un fichier CSV centralisé sur Ragic (`matching-matrix/6.csv`).
- **Chargement** : Au démarrage ou sur demande, une requête PowerQuery charge ces données dans une feuille de cache masquée (`PQ_DICT`).
- **Mise en cache** : Le dictionnaire en mémoire (`RagicFieldDict`) est peuplé à partir de cette feuille. Le rechargement depuis le réseau n'a lieu qu'une fois par jour ou si le cache est inexistant, mais peut être forcé via le bouton "Update Data Dictionary" du ruban.

### Structure et Utilisation
Le dictionnaire est un `Scripting.Dictionary` où :
- La **clé** est une chaîne composite : `NomFeuilleNormalisé & "|" & NomDuChamp`.
- La **valeur** contient les méta-informations.

**Exemple d'utilisation (`DataFormatter.bas`) :**
Le code interroge le dictionnaire pour savoir comment traiter une cellule en fonction du nom de sa colonne (le champ) et de sa feuille (la catégorie).

```vba
' Dans DataFormatter.bas
Public Function GetCellProcessingInfo(...) As FormattedCellOutput
    ' ...
    ' Récupère le type de champ depuis le dictionnaire
    ragicType = RagicDictionary.GetFieldRagicType(categorySheetName, fieldName)
    
    ' Applique une logique en fonction du type retourné
    Select Case ragicType
        Case "DATE":
            ' ... formater comme une date
        Case "NUMBER":
            ' ... formater comme un nombre
        ' ...
    End Select
    ' ...
End Function

' Dans RagicDictionary.bas
Public Function GetFieldRagicType(categorySheetName As String, fieldName As String) As String
    ' Construit la clé et recherche dans le dictionnaire
    Dim key As String
    key = NormalizeSheetName(categorySheetName) & "|" & fieldName
    
    If RagicFieldDict.Exists(key) Then
        ' Extrait l'info de la valeur...
    End If
End Function
```
Ce mécanisme rend l'addin beaucoup plus flexible, car une modification dans le fichier CSV de Ragic suffit à changer le comportement de l'addin sans redéploiement.

---

## 8. Guide de Contribution

### Comment ajouter une nouvelle fonctionnalité (ex: un nouveau bouton) ?

1.  **XML (`customUI.xml`)** : Ajouter le bouton et lier son `onAction` à un nouveau nom de Sub (ex: `ProcessNewFeature`).
2.  **Module Manager** : Créer la `Public Sub ProcessNewFeature(control as IRibbonControl)` dans le module manager le plus pertinent. Cette Sub doit être un simple "wrapper" qui appelle d'autres fonctions pour faire le vrai travail.
3.  **Logique Métier** : Implémenter la logique dans les modules appropriés (`DataLoaderManager`, etc.).
4.  **Types** : Si de nouvelles structures de données sont nécessaires, les ajouter dans `Types.bas`.
5.  **Documentation** : Si un nouveau pattern est introduit, le documenter dans ce fichier.

**Ce guide doit être suivi à la lettre pour toute évolution du projet.**