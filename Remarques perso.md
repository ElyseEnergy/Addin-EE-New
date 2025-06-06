2025.05.21_09.53

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

2025.06.06_12.03

Logs dans Ragic
	identifier les user par leurs emails => créer l'email à partir du username
	implémenter les mêmes messages dans les 3 loggings (immediate window, 
	
Import :
	importer par script vbs ? avec les fonctions d'import directement intégré dans vb ?
	frm, cls, bas

Logging :
	logger l'ensemble des accès boutons
	logger chaque étape de demande de database
	
Performance :
	comment ne pas requêter le dictionnaire à chaque fois ?
	demander une mise à jour du dictionnaire explicite la première fois par session/par jour ?
	
UX :
	Supprimer les msgbox après sélection des userform
	UserFormSelection
		Récupérer la userform de Nicolas
		Câbler
		Reformater àla volée en fonction de la taille du texte
	Loading Form

Changelog user:
	Ragic :
		Ajouter une database contenant les messages de changelog en markdown
Ajouter un fichier/database markdown

tester les new userform

mettre à jour un script d'update => comment réaliser la liaison ? => dans la database de logs ?

problème pour télécharger les CAPEX

je veux créer une message box générique sous la forme d'une box avec

encore un problème de gestion des annulations qui crée une erreur à la fin de chaque processus

le "premier lancement" ... c'est relou et ça devrait être avec un loader

gérer les champs hidden => sauter une ligner => ajouter dans la database dictionnary

changer le log du user avec un email