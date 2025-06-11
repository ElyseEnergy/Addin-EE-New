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

### Assignation des Types Personnalisés
Lorsqu'un type personnalisé contient des objets (comme `Collection`, `Range`, ou un autre type personnalisé), il est **impératif** d'utiliser le mot-clé `Set` pour l'assignation. Ne jamais assigner directement sans `Set`.

**Exemple incorrect :**
```vba
Dim loadInfo As DataLoadInfo
loadInfo = DeserializeLoadInfo(metadata) ' Erreur : Type d'argument ByRef incompatible
```

**Exemple correct :**
```vba
Dim loadInfo As DataLoadInfo
Set loadInfo = DeserializeLoadInfo(metadata) ' Correct : utilise Set pour l'assignation
```

Cette règle s'applique car les types contenant des objets sont passés par référence (ByRef). L'omission de `Set` provoquera une erreur de compilation "Type d'argument ByRef incompatible".

### Passage des Types en Paramètres (ByRef vs ByVal)
Une règle fondamentale et souvent source d'erreurs en VBA est que les **types définis par l'utilisateur (UDT) ne peuvent pas être passés avec `ByVal`** dans les procédures `Public` d'un module standard (`.bas`). Ils sont toujours passés par référence (`ByRef`).

Tenter de forcer le passage par valeur (`ByVal`) résultera en une erreur de compilation : *"Un type défini par l'utilisateur ne peut pas être passé ByVal"*.

**Exemple incorrect :**
```vba
' Dans un module standard (.bas)
Public Sub ProcessMyType(ByVal myData As CategoryInfo) ' ERREUR DE COMPILATION
    ' ...
End Sub
```

**Exemple correct :**
```vba
' Le mot-clé ByRef est implicite et peut être omis
Public Sub ProcessMyType(info As CategoryInfo) ' OK
    ' ...
End Sub

' Ou de manière explicite :
Public Sub ProcessMyTypeExplicit(ByRef info As CategoryInfo) ' OK
    ' ...
End Sub
```
**Important :** Cette règle s'applique aux modules standards. Le comportement est différent pour les procédures `Public` dans des modules de classe.

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

## 3. Interaction avec les Données Externes et API

### Variables d'Environnement (`env.bas`)
Le module `env.bas` centralise toutes les constantes et fonctions liées à l'environnement externe. Ne jamais coder en dur une URL ou une clé API ailleurs.
- **`RAGIC_BASE_URL`**: URL de base pour toutes les requêtes vers l'API Ragic.
- **`GetRagicApiKey()`**: Fonction qui retourne la clé API Ragic, chargée depuis une variable d'environnement.
- **`GetRagicApiParams()`**: Fonction qui retourne les paramètres standards pour les URLs de lecture (`GET`), incluant la clé API.

### Bonne Pratique : Lire des Données depuis une URL CSV
Pour lire des données, le pattern standard est d'utiliser le système Power Query intégré pour bénéficier de la mise en cache, de la performance et de la cohérence.

1.  **Définir une `CategoryInfo`** : Créer ou utiliser une catégorie existante qui définit l'URL de la ressource.
2.  **Assurer l'existence de la requête** : Appeler `PQQueryManager.EnsurePQQueryExists(maCategorie)`. Cette fonction crée ou met à jour la requête Power Query en mémoire sans la rafraîchir.
3.  **Charger les données** : Appeler `LoadQueries.LoadQuery(maCategorie.PowerQueryName, ...)`. Cette fonction exécute la requête et charge les données dans une table d'une feuille de cache (`PQ_DATA`).
4.  **Manipuler les données locales** : Votre code doit ensuite lire les données depuis la table locale, et non directement depuis le web.

### Bonne Pratique : Créer une Entrée dans Ragic (POST)
Pour envoyer des données à Ragic (créer une entrée), il faut forger une requête `POST` avec un corps en JSON.

**Exemple de la fonction de log (`SYS_Logger.bas`) :**
```vba
Private Sub LogToRagic(ByVal logMessage As String)
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0") ' Ou une version fallback
    
    ' 1. Construire le JSON
    Dim jsonPayload As String
    jsonPayload = "{" & _
        """" & RAGIC_FIELD_ID_EMAIL & """: """ & JsonEscape(userEmail) & """, " & _
        """" & RAGIC_FIELD_ID_LOG & """: """ & JsonEscape(logMessage) & """" & _
    "}"

    ' 2. Préparer la requête POST
    Dim ragicUrl As String
    ragicUrl = RAGIC_LOG_URL & "?APIKey=" & env.RAGIC_API_KEY
    http.Open "POST", ragicUrl, True ' True = Asynchrone

    ' 3. Définir le header
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    
    ' 4. Envoyer (de manière asynchrone)
    http.send jsonPayload
End Sub
```
**Structure du JSON pour les cas avancés :**
- **Sélection multiple** : Utiliser un tableau de chaînes.
  `"field_id": ["valeur1", "valeur2"]`
- **Sous-table** : Utiliser un objet dont le nom est `_subtable_FIELD_ID` et qui contient des objets pour chaque ligne, identifiés par un ID négatif unique.
  ```json
  {
      "2000123": "Dunder Mifflin",
      "_subtable_2000154": {
          "-1": { "2000147": "Bill", "2000148": "Manager" },
          "-2": { "2000147": "Satya", "2000148": "VP" }
      }
  }
  ```

---

## 4. Bonnes Pratiques Générales du Code VBA

### Structure d'un Module
Chaque module doit **obligatoirement** respecter la structure suivante pour la cohérence et la fiabilité, notamment pour l'import/export via des outils externes comme Git.

```vba
Attribute VB_Name = "NomDuModule"
Option Explicit

' --- Déclarations API Windows ---
' Toutes les déclarations d'API Windows doivent être au niveau du module
#If VBA7 Then
    Private Declare PtrSafe Function MaFonctionAPI Lib "kernel32" ...
#Else
    Private Declare Function MaFonctionAPI Lib "kernel32" ...
#End If

' --- Constantes et Enums du module ---
' --- Variables privées du module ---
' --- Procédures publiques ---
' --- Procédures privées ---
```

1.  **`Attribute VB_Name = "NomDuModule"`** : Toujours en **première ligne**. Ne jamais l'oublier, sinon l'import peut échouer ou créer un module au nom incorrect (ex: `Module1`).
2.  **`Option Explicit`** : Toujours en **deuxième ligne**. Force la déclaration de toutes les variables.
3.  **Déclarations API Windows** : Toujours au niveau du module, jamais dans une procédure. Les déclarations d'API doivent :
    - Être placées dans une section dédiée après `Option Explicit`
    - Utiliser la directive de compilation conditionnelle `#If VBA7 Then` pour la compatibilité 32/64 bits
    - Être déclarées comme `Private` sauf si elles doivent être partagées

**⚠️ AVERTISSEMENT :** Ne jamais déclarer des API Windows à l'intérieur d'une fonction ou d'une procédure. Cela peut causer des erreurs de compilation ou des comportements imprévisibles.

### Règle : Pas de `Continue For` en VBA
- **Important :** L'instruction `Continue For` n'existe pas en VBA. Pour passer à l'itération suivante d'une boucle, il faut utiliser une structure `If...Then...Else` et placer le code à exécuter uniquement en l'absence d'erreur dans le bloc `Else`.
- **Exemple incorrect :**
```vba
For i = 1 To 10
    If erreur Then Continue For ' <-- Syntaxe invalide en VBA
    ' ...
Next i
```
- **Exemple correct :**
```vba
For i = 1 To 10
    If erreur Then
        ' Gérer l'erreur ou logger
    Else
        ' ... code normal ...
    End If
Next i
```

### Règle : Join ne fonctionne qu'avec des Arrays
- **Important :** La fonction `Join` ne fonctionne qu'avec des tableaux (arrays), pas avec des Collections. Pour utiliser `Join` avec une Collection, il faut d'abord convertir la Collection en array.
- **Exemple incorrect :**
```vba
Dim maCollection As Collection
Set maCollection = New Collection
maCollection.Add "A"
maCollection.Add "B"
Debug.Print Join(maCollection, ", ") ' <-- Erreur !
```
- **Exemple correct :**
```vba
Dim maCollection As Collection
Set maCollection = New Collection
maCollection.Add "A"
maCollection.Add "B"

' Convertir en array avant d'utiliser Join
Dim monArray() As String
ReDim monArray(1 To maCollection.Count)
Dim i As Long
For i = 1 To maCollection.Count
    monArray(i) = CStr(maCollection(i))
Next i
Debug.Print Join(monArray, ", ") ' OK : affiche "A, B"
```

### Déclaration des Variables et Constantes
- **Déclaration en haut de procédure** : Toutes les variables locales doivent être déclarées au début de la fonction ou de la `Sub` pour une meilleure lisibilité.
- **Déclaration des constantes** : Une constante (`Const`) doit être initialisée avec une **valeur littérale** (ex: `123`, `"texte"`) ou une autre constante. **Les appels de fonction (comme `RGB()`) sont interdits** car ils sont évalués à l'exécution, et non à la compilation.
- **Éviter les doubles déclarations** : Toujours vérifier qu'une variable n'est pas déjà déclarée dans la même portée, que ce soit :
  - Dans la même procédure
  - Dans une boucle `For` (ex : ne pas redéclarer l'index `i` dans une boucle imbriquée)
  - Dans les paramètres de la fonction
  - En tant que variable de module

**Exemple de ce qu'il ne faut PAS faire :**
```vba
Public Sub ProcessData(ByVal i As Long)  ' i est déjà un paramètre
    Dim result As String
    Dim i As Long  ' ERREUR : i est déjà déclaré comme paramètre !
    
    For i = 1 To 10
        Dim j As Long
        For j = 1 To 5
            Dim i As Long  ' ERREUR : i est déjà l'index de la boucle externe !
            ' ...
        Next j
    Next i
End Sub
```

**Pattern correct :**
```vba
Public Sub ProcessData(ByVal inputIndex As Long)  ' Nom clair et unique
    Dim result As String
    Dim outerIndex As Long, innerIndex As Long  ' Indices distincts pour les boucles
    
    For outerIndex = 1 To 10
        For innerIndex = 1 To 5
            ' Utilisation des variables avec des noms uniques et explicites
            result = result & outerIndex & "," & innerIndex
        Next innerIndex
    Next outerIndex
End Sub
```

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

## 5. Gestion des Erreurs et des Logs

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
    ' IMPORTANT : Ne jamais utiliser "On Error Resume Next" dans un ErrorHandler
    ' car cela masquerait les erreurs suivantes et empêcherait leur gestion correcte
    
    ' Appel au gestionnaire centralisé
    HandleError MODULE_NAME, PROC_NAME, "Une erreur spécifique est survenue ici."
End Sub
```

**⚠️ AVERTISSEMENT :** Ne jamais utiliser `On Error Resume Next` dans un bloc `ErrorHandler`. Cette instruction masquerait toutes les erreurs suivantes et empêcherait leur gestion correcte. Si vous avez besoin de gérer des erreurs spécifiques dans le bloc `ErrorHandler`, utilisez des conditions sur `Err.Number` à la place.
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

## 6. Callbacks du Ruban (customUI)

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

2.  **Actions Métier (`*.bas`)**
    Les callbacks `onAction` qui déclenchent un processus métier sont placées dans le module  approprié. Elles servent de simple point d'entrée.

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

## 7. Normalisation des Requêtes PowerQuery

La création et la gestion des requêtes PQ sont entièrement automatisées et normalisées via `PQQueryManager.bas` pour garantir la cohérence.

### Convention de Nommage
La cohérence des noms est essentielle pour que les automatisations fonctionnent correctement.

- **Requêtes Power Query** :
  - **Préfixe** : Toutes les requêtes gérées par l'addin commencent par `PQ_`.
  - **Nom de base** : Le nom est dérivé du `PowerQueryName` de la `CategoryInfo`, qui est lui-même basé sur le `CategoryName`.
  - **Nettoyage** : Le nom de base est "nettoyé" par la fonction `Utilities.SanitizeTableName` qui :
    - Remplace les espaces et caractères spéciaux (`-`, `.`, `/`, etc.) par des underscores `_`.
    - Supprime tous les accents (diacritiques).
    - Ne conserve que les caractères alphanumériques et les underscores.
  - **Exemple** : Une catégorie nommée "Coûts (détails)" aura une requête nommée `PQ_Couts_details`.

- **Tables Excel** :
  - **Préfixe** : Chaque table créée à partir d'une requête Power Query est préfixée par `Table_`.
  - **Nom** : Le reste du nom est le nom complet de la requête qui la génère.
  - **Exemple** : La requête `PQ_Couts_details` chargera ses données dans une table nommée `Table_PQ_Couts_details`.

### Processus de gestion
- **Création/Mise à jour** : `PQQueryManager.EnsurePQQueryExists` vérifie si une requête existe, si sa formule (URL) a changé, et la crée ou la met à jour au besoin.
- **Template de Requête** : `GeneratePQQueryTemplate` crée le code M standard, qui inclut le typage de la colonne ID et sa mise en première position.
- **Chargement** : `LoadQueries.LoadQuery` est la seule fonction à utiliser pour charger une requête dans une feuille.

---

## 8. Le Dictionnaire de Données (Ragic Dictionary)

Le "Ragic Dictionary" est un mécanisme clé de l'addin. Il s'agit d'une table de correspondance qui fournit des **méta-informations sur les champs de données Ragic**, comme leur type de données réel ("DATE", "NUMBER", etc.) ou si un champ doit être masqué dans l'interface.

### Rôle et Utilité
Il permet de décorréler la logique de l'addin des données brutes. Par exemple, au lieu de coder en dur qu'un champ nommé "Date de début" doit être formaté comme une date, on consulte le dictionnaire pour connaître son type.

### Structure Détaillée du Dictionnaire

Le dictionnaire en mémoire (`RagicFieldDict`) est un objet `Scripting.Dictionary` dont la structure clé/valeur est très spécifique.

- **La Clé** :
  - C'est une chaîne de caractères composite qui identifie un champ de manière unique dans tout l'addin.
  - **Format** : `NomDeFeuilleNormalisé & "|" & NomDuChamp`
  - **`NomDeFeuilleNormalisé`** : Le nom de la feuille de la catégorie (ex: "Synthèse Coûts") est "normalisé" par la fonction `NormalizeSheetName` en ne gardant que les lettres et les chiffres (`SyntheseCouts`).
  - **`NomDuChamp`** : Le nom de la colonne tel qu'il apparaît dans les données (ex: "Date de validation").
  - **Exemple de clé** : `SyntheseCouts|Date de validation`

- **La Valeur** :
  - C'est une chaîne de caractères simple lue depuis la colonne **"Field Type"** du fichier CSV source du dictionnaire.
  - Cette chaîne définit le **type de donnée sémantique** du champ et peut contenir des **indicateurs (flags)**.
  - **Exemples de valeurs** :
    - `"DATE"` : Indique que le champ doit être traité comme une date.
    - `"NUMBER"` : Doit être traité comme un nombre.
    - `"TEXT"`
    - `"PERCENT"`
    - `