﻿2025.05.21_09.53

Je remarque que dans Ragic, si sélection de pourcentage, alors le % doit être stocké en string => quand on consomme la data, il faudrait le reconvertir en pourcentages

Pour l'instant dans les datasheet, pas de typage fort des chiffres

J'aimerais pouvoir lire le html du ragic dictionnary (et plus tard dans la base ragic en fait) pour avoir le format à afficher

quand il y a des min max comme dans CO2 capture, Specific Electricity Consumption (SEC) [MWhe/tonCO2], il vaut mieux créer deux champs ragic pour permettre le format des nombres


dans PQ_data, faut il unloader à chaque fois les requêtes non utilisées ?

2025.05.21_11.06

Il faut que tous les formulaires aient un champ ID
=> remettre sur le CO2

encore des tests à effectuer sur le fait que quand je mets 1 pour CO2 parameters (une seule fiche disponible) il me jette une erreur

2025.05.21_17.11

Bon je crois que je m'arrête là, c'est déjà pas mal.
Normalement c'est fonctionnel


2025.05.2_12.46

Les groupes de visibilité semblent fonctionner
il y a encore quelques problèmes d'icones et les groupes ne sont pas fonctionnelement corrects
maintenant nous pouvons loader plus de data

2025.05.22_09.00

Pour les plannings, la sélection se fait désormais en deux étapes :
1. Choix du projet (filtre principal)
2. Choix de la version de planning
Toutes les lignes correspondant à ces choix sont chargées automatiquement sans nouvelle sélection.
