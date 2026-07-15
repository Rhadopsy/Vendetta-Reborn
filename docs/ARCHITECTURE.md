# Architecture de Vendetta 3.0.0.17

## Vue d'ensemble

Le programme est une application monolithique Visual Basic 6. `ModMain.bas` contient `Sub Main`, initialise DirectX 7, charge la configuration et les langues, puis construit l'affichage, le son, les périphériques, l'IA, les sauvegardes et le réseau.

Le fichier `PrjImperator.vbp` est la source de vérité pour les composants, classes, formulaires et références COM du projet.

## Couches identifiées

| Préfixe ou emplacement | Responsabilité observée |
|---|---|
| `ClsJeu*` | Monde, personnages, objets, effets, ressources, trésors, partie et règles de jeu |
| `ClsAff*` | Rendu, fenêtres, personnages, objets et ressources via DirectDraw |
| `ClsInt*` | Fenêtres, boutons, touches et interactions de l'interface |
| `ClsPer*` | Abstraction historique du clavier, de la souris et des périphériques DirectInput |
| `ClsCom*` | Réseau, échanges et représentation distante du monde |
| `ClsSav*`, `ModSav*` | Chargement et sauvegarde des données |
| `Mod*` | Initialisation et services globaux, dont décors et réseau |
| `Frm*`, `MDIFrmMain*` | Formulaires VB6, fenêtres de paramétrage et éditeurs |
| `IAs`, `IAScripts` | Définitions et scripts de comportement des personnages non joueurs |
| `Donnees`, `Cartes`, `Mods`, `Sauvegardes` | Configuration, cartes, extensions et états persistants |
| `Images`, `Sons`, `Musiques` | Ressources graphiques et audio |

## Dépendances historiques et état du portage

Le projet historique référence le runtime VB6 et des technologies Windows anciennes : DirectX 7 pour VB, DirectDraw7, DirectInput, DirectSound, DirectMusic, DirectPlay4, Script Control, MSXML 3, Common Controls et Internet Transfer Control.

Sur la branche de portage, Script Control et MSXML sont chargés tardivement, Internet Transfer Control est supprimé, et les appels à `user32`, `gdi32`, `kernel32`, `shell32`, `winmm` et `iphlpapi` disposent de déclarations twinBASIC compatibles 32/64 bits. DirectX 7 pour VB et `COMCTL32.OCX` restent les deux blocages structurants.

Ces dépendances forment la frontière à remplacer. Les règles du jeu et les formats de données restent la partie la plus réutilisable.

## Risques de migration

- les fichiers VB6 utilisent probablement un encodage Windows historique ; ils doivent rester intacts pendant l'analyse ;
- les fichiers `.frm` référencent leurs compagnons binaires `.frx`, qui ne doivent pas être dissociés ;
- le comportement dépend de variables globales et de l'ordre d'initialisation de `Sub Main` ;
- les temporisations et boucles de jeu peuvent être liées à la cadence de l'ancien moteur ;
- DirectPlay et les contrôles COM n'ont pas d'équivalent moderne direct ;
- les formats de sauvegarde, cartes et scripts doivent être caractérisés avant toute conversion.
