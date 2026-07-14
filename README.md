# Vendetta Reborn

Ce dépôt contient désormais les sources originales de **Vendetta 3.0.0.17**, reçues sous la forme de l'archive `Vendetta-3.0.0.17-src.zip`.

Il s'agit d'un jeu complet écrit en **Visual Basic 6**, et non du prototype C++/SFML précédemment étudié. Ce dernier reste disponible sur la branche `prototype-sfml`.

## État de cette branche

- code source VB6 original et ressources de jeu importés sans conversion ;
- exécutable historique `Vendetta.exe` volontairement exclu : il n'est pas nécessaire pour conserver les sources et n'a pas été exécuté ;
- point d'entrée : `Sub Main` dans `ModMain.bas` ;
- projet VB6 : `PrjImperator.vbp` ;
- version déclarée : 3.0.0, révision 17 dans le nom de l'archive ;
- licence principale : GNU GPL v2 ou ultérieure, selon les en-têtes des sources et `license.txt`.

L'empreinte de l'archive d'origine est conservée dans `SOURCE_ARCHIVE.sha256`.

## Compatibilité actuelle

Le projet dépend notamment du runtime VB6, de DirectX 7 pour VB (`dx7vb.dll`), de Microsoft Script Control, MSXML 3, `COMCTL32.OCX` et `MSINET.OCX`. Il utilise DirectDraw, DirectInput, DirectSound, DirectMusic et DirectPlay, ainsi que plusieurs API Win32.

Il ne peut donc pas être compilé nativement sur Fedora ni être considéré comme compatible avec les Windows récents en l'état. Wine et d'anciens composants Windows peuvent servir à l'archéologie ou à une comparaison ponctuelle, mais la remise au goût du jour nécessite un portage progressif.

Voir [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) pour l'inventaire et [docs/PORTING_PLAN.md](docs/PORTING_PLAN.md) pour la stratégie proposée.

## Contenu principal

- `ClsJeu*.cls` : règles et état du jeu ;
- `ClsAff*.cls` : affichage DirectDraw ;
- `ClsInt*.cls` : interface utilisateur ;
- `ClsPer*.cls` : clavier, souris et autres périphériques ;
- `ClsCom*.cls` : communications et réseau ;
- `Mod*.bas` : modules globaux et initialisation ;
- `Frm*.frm`, `MDIFrmMain.frm` : formulaires VB6 et éditeurs ;
- `Donnees`, `Cartes`, `IAs`, `IAScripts`, `Mods`, `Images`, `Sons`, `Musiques` : données et ressources du jeu.

## Prochaine cible

La cible recommandée est **Godot 4.x**, avec une simulation portée indépendamment de l'affichage. Ce choix offre un export natif Linux et Windows et permet de réutiliser les données et ressources progressivement sans reproduire les dépendances COM/DirectX 7.
