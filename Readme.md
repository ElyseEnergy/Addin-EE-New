# Add-in Elyse Energy pour Excel

Cet add-in a pour but de faciliter l'interaction avec les données de la société Elyse Energy directement depuis Excel. Il permet de charger, mettre à jour et manipuler des données provenant de diverses sources via des requêtes PowerQuery.

## Architecture du Code

Le code VBA est structuré en plusieurs modules, chacun ayant une responsabilité unique. Cette approche modulaire vise à améliorer la maintenabilité, la lisibilité et la testabilité du code.

### Modules Principaux

-   **`DataLoaderManager.bas`**: C'est le module orchestrateur principal. Il coordonne le processus complet de chargement des données, depuis la sélection de la catégorie jusqu'à l'affichage final dans la feuille Excel. Il délègue les tâches spécifiques aux autres modules.

-   **`CategoryManager.bas`**: Gère la définition et la récupération des différentes catégories de données que l'utilisateur peut charger (ex: "Coûts", "Projets"). Chaque catégorie contient des méta-informations comme son nom, son URL de source de données, et les niveaux de filtre.

-   **`PQQueryManager.bas`**: Centralise la création, la mise à jour et la gestion des requêtes PowerQuery. Il s'assure que chaque requête est correctement formatée et nommée selon les conventions du projet.

-   **`RibbonVisibility.bas`**: Gère toute la logique d'affichage du ruban personnalisé (Custom UI). Il contrôle la visibilité, l'état (activé/désactivé) et les labels dynamiques des boutons en fonction du contexte (ex: profil utilisateur, tableaux présents dans le classeur).

-   **`Types.bas`**: Définit les types de données personnalisés (`Public Type`) utilisés à travers l'application, comme `CategoryInfo` et `DataLoadInfo`. La centralisation des types garantit la cohérence des structures de données.

### Modules Spécialisés (Helpers)

La logique complexe est décomposée en modules spécialisés pour une meilleure séparation des responsabilités :

-   **`TableMetadata.bas`**: Gère la **sérialisation** et la **désérialisation** des métadonnées des tableaux. Lorsqu'un tableau est créé, ses informations de chargement (catégorie, filtres, etc.) sont stockées dans le commentaire de la première cellule du tableau. Ce module se charge de convertir ces informations en une chaîne de caractères et vice-versa.

-   **`TableManager.bas`**: Centralise la manipulation des objets `ListObject` (tableaux Excel). Il fournit des fonctions pour générer des noms de tableaux uniques, lister tous les tableaux gérés par l'addin, et vérifier leur existence.

-   **`DataPasting.bas`**: Contient la logique de **collage** des données. Une fois les données récupérées par PowerQuery, ce module s'occupe de les copier depuis la feuille de cache `PQ_DATA` vers leur destination finale, en gérant les modes normal et transposé.

-   **`SheetManager.bas`**: Fournit des fonctions utilitaires pour la gestion des feuilles de calcul, comme la création de la feuille de cache `PQ_DATA`, la protection et la déprotection des feuilles.

-   **`DataInteraction.bas`**: Isole toutes les fonctions qui interagissent avec l'utilisateur via des boîtes de dialogue (`InputBox`, `MsgBox`), comme la sélection des fiches à charger ou le choix de la cellule de destination.

### Modules Système

-   **`SYS_ErrorHandler.bas`**: Module centralisé pour la gestion des erreurs. Toutes les erreurs inattendues sont capturées et traitées par la fonction `HandleError` de ce module.
-   **`SYS_Logger.bas`**: Fournit un système de logging pour tracer les événements importants de l'application.
-   **`Diagnostics.bas`**: Outils pour mesurer les performances et faciliter le débogage.

### Fichiers de Configuration

-   **`env.bas`**: Contient les constantes liées à l'environnement, comme les URLs de base des API et les clés d'authentification. **Ce fichier ne doit pas être versionné avec des informations sensibles.**
-   **`customUI.xml`**: Définit la structure et l'apparence du ruban personnalisé dans Excel.

## Configuration de l'Environnement

### Configuration de la Clé API Ragic

Pour des raisons de sécurité, la clé API Ragic doit être configurée via une variable d'environnement Windows. Voici comment procéder :

1. Ouvrir les Paramètres système avancés :
   - Clic droit sur "Ce PC" > Propriétés
   - Paramètres système avancés > Variables d'environnement

2. Dans la section "Variables d'environnement utilisateur", cliquer sur "Nouvelle" et ajouter :
   - Nom de la variable : `RAGIC_API_KEY`
   - Valeur : Votre clé API Ragic (sans le '&' à la fin)

3. Redémarrer Excel pour que les changements prennent effet.

**Note** : Si la variable d'environnement n'est pas configurée, l'add-in utilisera une clé de développement par défaut. Cette configuration n'est pas recommandée pour un usage en production.

## Flux de Données (Exemple)

1.  L'utilisateur clique sur un bouton du ruban pour charger des données (ex: "Charger Coûts").
2.  Le callback `onAction` dans le module métier (ex: `Technologies_Manager.bas`) est appelé.
3.  Ce callback délègue immédiatement le travail à `DataLoaderManager.ProcessCategory`.
4.  `DataLoaderManager` récupère les informations de la catégorie via `CategoryManager`.
5.  Il s'assure que la requête PowerQuery existe avec `PQQueryManager.EnsurePQQueryExists`.
6.  Il télécharge les données dans la feuille cachée `PQ_DATA` via `LoadQueries.LoadQuery`.
7.  Il demande à l'utilisateur de sélectionner les fiches et la destination via les fonctions de `DataInteraction.bas`.
8.  Il appelle `DataPasting.PasteData` pour coller les données à l'emplacement choisi.
9.  `DataPasting` crée un `ListObject` (tableau) en utilisant `TableManager.GetUniqueTableName`.
10. `DataPasting` sauvegarde les métadonnées du chargement dans le commentaire du tableau via `TableMetadata.SerializeLoadInfo`.
11. En cas d'erreur à n'importe quelle étape, `SYS_ErrorHandler` est appelé.
12. Des informations sont enregistrées tout au long du processus par `SYS_Logger`.
