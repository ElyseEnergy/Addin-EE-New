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

1.  **Attribut `VB_Name`** : Chaque module exporté (`.bas`, `.cls`) doit **obligatoirement** commencer par la ligne `Attribute VB_Name = "NomDuModule"`. Cela garantit un import/export fiable.

2.  **Modularité** : La logique est séparée par domaine fonctionnel.
    *   `CategoryManager` : Gère la définition et l'accès aux catégories.
    *   `DataLoaderManager` : Orchestre le processus de chargement des données.
    *   `PQQueryManager` : Gère la création et la maintenance des requêtes PowerQuery.
    *   `RibbonVisibility` : Gère l'état et la visibilité du ruban.
    *   `SYS_Logger` / `SYS_ErrorHandler` : Modules système pour le logging et les erreurs.

3.  **Constantes** : Les constantes partagées sont définies en haut du module le plus pertinent.
    *   Les variables d'environnement (clés API, URLs) sont dans `env.bas`.
    *   La version de l'addin est dans `Utilities.bas`.

---

## 3. Gestion des Erreurs et des Logs

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

## 4. Callbacks du Ruban (customUI)

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

## 5. Normalisation des Requêtes PowerQuery

La création et la gestion des requêtes PQ sont entièrement automatisées et normalisées via `PQQueryManager.bas` pour garantir la cohérence.

- **Nommage** : Le nom de la requête et de la table associée est généré via `Utilities.SanitizeTableName`.
- **Création/Mise à jour** : `PQQueryManager.EnsurePQQueryExists` vérifie si une requête existe, si sa formule (URL) a changé, et la crée ou la met à jour au besoin.
- **Template de Requête** : `GeneratePQQueryTemplate` crée le code M standard, qui inclut le typage de la colonne ID et sa mise en première position.
- **Chargement** : `LoadQueries.LoadQuery` est la seule fonction à utiliser pour charger une requête dans une feuille.

---

## 6. Guide de Contribution

### Comment ajouter une nouvelle fonctionnalité (ex: un nouveau bouton) ?

1.  **XML (`customUI.xml`)** : Ajouter le bouton et lier son `onAction` à un nouveau nom de Sub (ex: `ProcessNewFeature`).
2.  **Module Manager** : Créer la `Public Sub ProcessNewFeature(control as IRibbonControl)` dans le module manager le plus pertinent. Cette Sub doit être un simple "wrapper" qui appelle d'autres fonctions pour faire le vrai travail.
3.  **Logique Métier** : Implémenter la logique dans les modules appropriés (`DataLoaderManager`, etc.).
4.  **Types** : Si de nouvelles structures de données sont nécessaires, les ajouter dans `Types.bas`.
5.  **Documentation** : Si un nouveau pattern est introduit, le documenter dans ce fichier.

**Ce guide doit être suivi à la lettre pour toute évolution du projet.**