# Vendetta Reborn

Ce dépôt contient désormais les sources originales de **Vendetta 3.0.0.17**, reçues sous la forme de l'archive `Vendetta-3.0.0.17-src.zip`.

Il s'agit d'un jeu complet écrit en **Visual Basic 6**, et non du prototype C++/SFML précédemment étudié. Ce dernier reste disponible sur la branche `prototype-sfml`.

## État du projet

- sources VB6 et ressources du jeu importées depuis l'archive originale ;
- adaptations de compatibilité isolées sur la branche `port/twinbasic-win32-win64` ;
- exécutable historique `Vendetta.exe` conservé uniquement pour les tests de compatibilité Wine/Proton ; il provient de l'archive originale et n'a pas été exécuté pendant l'import ;
- SHA-256 de `Vendetta.exe` : `914cacd4fea78f82b91e5480e52a0bb619005718d5eeb257c2b4451211936898` ;
- point d'entrée : `Sub Main` dans `ModMain.bas` ;
- projet VB6 : `PrjImperator.vbp` ;
- version déclarée : 3.0.0, révision 17 dans le nom de l'archive ;
- licence principale : GNU GPL v2 ou ultérieure, selon les en-têtes des sources et `license.txt`.

L'empreinte de l'archive d'origine est conservée dans `SOURCE_ARCHIVE.sha256`.

## Compatibilité actuelle

La branche de portage supprime les références de compilation à Microsoft Script Control, MSXML et `MSINET.OCX`. Le moteur VBScript devient facultatif, le HTTP passe par MSXML chargé dynamiquement et les appels Win32 sont préparés pour `PtrSafe`/`LongPtr` sous twinBASIC.

DirectX 7 pour VB (`dx7vb.dll`) et `COMCTL32.OCX` restent nécessaires. La première cible réaliste est donc un exécutable **Win32** produit par twinBASIC, autonome vis-à-vis de `MSVBVM60.DLL` et `VB6FR.DLL`. Le Win64 ne deviendra exécutable qu'après remplacement du rendu, des entrées, de l'audio et du réseau DirectX 7 ainsi que des contrôles OCX restants.

Voir [docs/TWINBASIC_BUILD.md](docs/TWINBASIC_BUILD.md) pour construire et tester la cible intermédiaire, [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) pour l'inventaire et [docs/PORTING_PLAN.md](docs/PORTING_PLAN.md) pour le portage natif à long terme.

## Contenu principal

- `ClsJeu*.cls` : règles et état du jeu ;
- `ClsAff*.cls` : affichage DirectDraw ;
- `ClsInt*.cls` : interface utilisateur ;
- `ClsPer*.cls` : clavier, souris et autres périphériques ;
- `ClsCom*.cls` : communications et réseau ;
- `Mod*.bas` : modules globaux et initialisation ;
- `Frm*.frm`, `MDIFrmMain.frm` : formulaires VB6 et éditeurs ;
- `Donnees`, `Cartes`, `IAs`, `IAScripts`, `Mods`, `Images`, `Sons`, `Musiques` : données et ressources du jeu.

## Cibles

- court terme : twinBASIC Win32 pour retrouver un jeu lançable sans runtime VB6 ni DLL de langue ;
- long terme : **Godot 4.x** pour supprimer DirectX 7/COM et produire des versions Linux et Windows x86_64 natives.
