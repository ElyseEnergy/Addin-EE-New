# Addin Elyse Energy

Ce dépôt contient le code de l'addin Excel **Elyse Energy**. L'objectif de cet add-in est de charger depuis Ragic (base de données en ligne) les fiches techniques ou financières nécessaires aux projets et de les mettre en forme dans Excel via Power Query.

## Prérequis

- Excel 2016 ou version plus récente avec Power Query.
- Un fichier `ENV.bas` (non fourni) à placer à la racine du projet. Il doit définir au minimum les constantes suivantes :
  - `RAGIC_BASE_URL` : URL de base de l'API Ragic.
  - `RAGIC_API_PARAMS` : paramètres communs (clé API, format CSV...).

## Fichiers principaux

- `Addin Elyse Energy.xlsm` : version de développement contenant tous les modules VBA.
- `Addin Elyse Energy.xlam` : add-in généré à partir du fichier précédent.
- `ConvertToXLAM.vbs` : script pour convertir le `xlsm` en `xlam`.
- `customUI.xml` : définition du ruban personnalisé (onglet *Elyse Energy*).
- `icons/` : icônes utilisées par le ruban.

## Modules VBA

- `Types.bas` : types personnalisés (`CategoryInfo`, `DataLoadInfo`, `AccessProfile`, ...).
- `CategoryManager.bas` : définition de toutes les catégories à charger depuis Ragic.
- `AccessProfiles.bas` : profils de démonstration et droits d'accès.
- `RibbonVisibility.bas` : callbacks du ruban et logique de visibilité.
- `Technologies_Manager.bas` : macros appelées par les boutons du ruban.
- `DataLoaderManager.bas` : logique principale de chargement/filtrage des données et création des tableaux Excel.
- `PQQueryManager.bas` : création et mise à jour des requêtes Power Query en fonction des catégories.
- `LoadQueries.bas` : insertion des données Power Query dans les feuilles Excel.
- `PQDebugTools.bas` : outils de débogage (injection/nettoyage de toutes les requêtes).
- `Utilities.bas` : fonctions utilitaires (sanitisation de noms, gestion de la feuille `PQ_DATA`, ...).

## Fonctionnement général

1. L'utilisateur clique sur un bouton du ruban *Elyse Energy*.
2. La macro associée dans `Technologies_Manager.bas` appelle `DataLoaderManager.ProcessCategory` avec la catégorie voulue.
3. `DataLoaderManager` crée ou met à jour la requête Power Query via `PQQueryManager`, charge les données dans `PQ_DATA` grâce à `LoadQueries`, puis demande éventuellement à l'utilisateur quelles fiches coller.
4. Les données sélectionnées sont copiées dans la feuille cible sous forme de tableau protégé (`EE_...`).

## Développement

1. Ouvrir `Addin Elyse Energy.xlsm` pour modifier ou ajouter des modules.
2. Créer un fichier `ENV.bas` adapté à votre environnement (API Ragic, clés...).
3. Exécuter `ConvertToXLAM.vbs` pour obtenir la version `xlam` distribuable.
4. Pour mettre à jour le ruban après modification de `customUI.xml`, utiliser le script `Inject XML (not tested).vbs`.
5. Les macros de `PQDebugTools.bas` peuvent injecter ou supprimer toutes les requêtes afin de tester rapidement l'import des données.

## Notes

Le fichier `ENV.bas` est volontairement ignoré par git afin de ne pas exposer les paramètres sensibles (URLs et clés API). Chaque développeur doit créer ce fichier localement.
