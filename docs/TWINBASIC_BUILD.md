# Construction twinBASIC et test Proton

## Résultat visé

La cible intermédiaire est un nouveau `Vendetta.exe` **Win32** compilé par twinBASIC. Ce binaire incorpore son propre runtime et ne doit donc plus demander `MSVBVM60.DLL` ni `VB6FR.DLL`.

Le 64 bits est préparé dans les déclarations d'API Windows, mais ne doit pas encore être activé : la bibliothèque Visual Basic de DirectX 7 et les contrôles `COMCTL32.OCX` sont toujours 32 bits.

## Changements déjà appliqués

| Zone | Avant | Branche de portage |
|---|---|---|
| Runtime et langue VB6 | `MSVBVM60.DLL` et `VB6FR.DLL` nécessaires | inclus dans un exécutable twinBASIC |
| Scripts d'IA/scénario | référence obligatoire à `MSSCRIPT.OCX` | création tardive ; repli sur l'IA intégrée si absent |
| Accès HTTP | contrôle `MSINET.OCX` dans `FrmParam` | `MSXML2.XMLHTTP` chargé dynamiquement |
| XML | référence de compilation MSXML 3 | MSXML 6 puis 3 chargé dynamiquement |
| API Windows | handles stockés en `Long` | `PtrSafe` et `LongPtr` sous twinBASIC, code VB6 conservé |
| DirectX 7 | `dx7vb.dll` obligatoire | migration vers les interfaces natives de WinDevLib en cours |
| Common Controls | `COMCTL32.OCX` | encore obligatoire, Win32 uniquement |

## Produire le premier exécutable

La compilation nécessite actuellement l'IDE twinBASIC sous Windows ; elle n'est pas automatisable depuis ce dépôt Linux.

1. Télécharger et extraire la version courante de l'IDE twinBASIC sur Windows.
2. Dans l'assistant de démarrage, choisir **Import from VBP** et ouvrir `PrjImperator.vbp`.
3. Désactiver l'option qui force `Option Explicit` à l'échelle du projet pendant ce premier import : plusieurs modules historiques utilisent des variables implicites.
4. Sélectionner une cible **Win32** dans les paramètres du projet.
5. Vérifier l'import des formulaires et de leurs fichiers `.frx` associés.
6. Dans **Available Packages**, activer **Windows Development Library for twinBASIC** (symbole `WinDevLib`). Ne pas activer `WinDevLib for Implements`.
7. Désactiver la référence manquante **DirectX 7 for Visual Basic Type Library**. `ModDirectXAliases.bas` raccorde progressivement les anciens noms `dx7vb` aux interfaces natives de WinDevLib.
8. Laisser provisoirement la référence `COMCTL32.OCX` : son remplacement est un chantier distinct limité à quatre formulaires.
9. Compiler un exécutable nommé `Vendetta.exe`.
10. Placer le binaire à la racine du jeu, à côté de `Donnees`, `Images`, `Cartes`, `Sons` et des autres répertoires : le code utilise `App.Path` pour les retrouver.

Ne pas remplacer l'exécutable historique dans Git avant d'avoir validé le nouveau binaire. Le publier d'abord sous un nom distinct, par exemple `Vendetta-twinBASIC-win32.exe`, avec son SHA-256 et le numéro de version twinBASIC utilisé.

## Test minimal sous Proton

Créer une nouvelle entrée Steam non-Steam pointant vers le binaire twinBASIC, afin d'obtenir un préfixe propre. Le test est réussi si :

1. aucune boîte de dialogue ne réclame `MSVBVM60.DLL` ou `VB6FR.DLL` ;
2. le menu principal s'affiche ;
3. une partie locale charge ses données et sa carte ;
4. le clavier, la souris, l'affichage et l'audio fonctionnent ;
5. quitter le jeu ne produit pas d'erreur non gérée.

Conserver un journal Proton pour chaque échec. Une erreur liée à `dx7vb.dll`, DirectDraw, DirectInput, DirectSound, DirectMusic ou DirectPlay appartient au chantier DirectX ; une erreur liée à un Slider, TreeView, TabStrip ou ProgressBar appartient au chantier Common Controls.

## Conditions pour Win64

Un binaire Win64 ne sera fiable qu'après ces remplacements :

- DirectDraw7 par un moteur de rendu moderne ;
- DirectInput par les événements clavier/souris natifs ;
- DirectSound/DirectMusic et MCI par une couche audio moderne ;
- DirectPlay4 par un protocole réseau maintenu ;
- les Slider, TreeView, TabStrip et ProgressBar de `COMCTL32.OCX` par des contrôles twinBASIC natifs ou une nouvelle interface.

Les règles du jeu et les formats de données peuvent rester partagés pendant cette migration. La cible Godot décrite dans `PORTING_PLAN.md` reste la voie la plus solide pour obtenir ensuite Linux et Windows x86_64 natifs.
