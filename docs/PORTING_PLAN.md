# Plan de portage Linux et Windows

L'objectif est un jeu natif sur Fedora et les Windows récents, sans runtime VB6, composants OCX ni DirectX 7. La cible proposée est Godot 4.x.

## Jalon intermédiaire : twinBASIC Win32

Avant le portage natif, une compilation twinBASIC Win32 sert à retrouver rapidement une version Windows/Proton testable sans `MSVBVM60.DLL` ni `VB6FR.DLL`. Cette cible conserve temporairement DirectX 7 et les Common Controls ; elle ne remplace donc pas les étapes ci-dessous et ne constitue pas encore le portage Linux/Win64.

## 0. Figer la référence originale

- conserver cette branche comme photographie vérifiable des sources 3.0.0.17 ;
- documenter les fichiers, dépendances et points d'entrée ;
- ne pas modifier mécaniquement l'encodage des fichiers VB6 ;
- cataloguer les comportements visibles et les formats persistants.

## 1. Extraire les données

- écrire des importeurs en lecture seule pour `Donnees`, `Cartes`, `IAs`, `Mods` et les sauvegardes ;
- ajouter des tests avec des fichiers d'origine comme cas de référence ;
- convertir uniquement vers des structures internes typées, sans altérer les originaux ;
- établir les règles de correspondance des images, sons et musiques.

## 2. Porter le cœur de simulation

- isoler le temps de jeu, le monde, les personnages, objets, effets, peuples et ressources ;
- remplacer les variables globales par un état explicite ;
- rendre la simulation déterministe lorsque c'est possible ;
- couvrir chaque sous-système par des tests avant d'ajouter l'affichage.

## 3. Construire le client Godot

- rendu 2D, caméra et cycles jour/nuit ;
- clavier, souris et configuration des commandes ;
- audio et musique ;
- interface, fenêtres et éditeurs nécessaires ;
- compatibilité des résolutions modernes et mise à l'échelle.

## 4. Porter l'IA et les scripts

- caractériser l'usage de Microsoft Script Control ;
- définir une API de script limitée et sûre ;
- convertir les comportements progressivement ;
- vérifier les décisions de l'IA sur des scénarios reproductibles.

## 5. Remplacer le réseau

- documenter les messages et règles de synchronisation DirectPlay ;
- choisir une autorité claire pour la partie ;
- créer un protocole versionné et testable ;
- ajouter le multijoueur seulement après stabilisation du mode local.

## 6. Livrer

- exports Godot x86_64 pour Fedora/Linux et Windows 10/11 ;
- contrôles automatiques sur les deux plateformes ;
- migration versionnée des sauvegardes ;
- paquet contenant uniquement les ressources redistribuables et les licences requises.

## Premier jalon jouable

Le premier jalon doit charger une carte et ses données d'origine, afficher le monde, déplacer un personnage et sauvegarder/recharger son état. Le son, l'IA avancée, les éditeurs et le réseau viendront ensuite : ce découpage permet de valider rapidement les formats et la boucle centrale sans reproduire d'emblée tout le programme VB6.
