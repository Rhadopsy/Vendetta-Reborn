# Vendetta — modernisation communautaire

Vendetta est un ancien jeu libre de gestion, stratégie et RPG créé par Quentin Santos. Ce dépôt repart de la [dernière révision SVN officielle, r268](https://sourceforge.net/p/vendetta2/code/HEAD/tree/), en conservant les données du jeu et sa licence GPL-3.0-or-later.

Cette branche de modernisation remplace le socle SFML 1.x par SFML 2.6, adopte C++17 et CMake, et supprime les dépendances POSIX qui empêchaient une compilation native sous Windows. La couche de gameplay Python 2.6 est isolée et désactivée par défaut pendant sa migration vers Python 3 ; le moteur C++ reste utilisable sans elle.

## Fedora

Prérequis (Fedora 44 ou version récente) :

```bash
sudo dnf install gcc-c++ cmake ninja-build SFML-devel
```

Compilation et lancement :

```bash
cmake -S . -B build -G Ninja -DCMAKE_BUILD_TYPE=Release
cmake --build build
./build/vendetta
```

## Windows 10/11

Installez Visual Studio 2022 avec le composant « Développement Desktop en C++ », puis [vcpkg](https://learn.microsoft.com/vcpkg/get_started/get-started). Dans un terminal Developer PowerShell :

```powershell
vcpkg install sfml:x64-windows
cmake -S . -B build -A x64 `
  -DCMAKE_TOOLCHAIN_FILE="$env:VCPKG_ROOT/scripts/buildsystems/vcpkg.cmake"
cmake --build build --config Release
./build/Release/vendetta.exe
```

Le répertoire `Data` est copié automatiquement à côté de l'exécutable. Les builds Fedora et Windows sont aussi vérifiés par GitHub Actions.

## État du portage

- SFML 2.6 et C++17 : portage initial terminé ;
- chemins et parcours de répertoires : portables via `std::filesystem` ;
- système de build : CMake, Fedora et Visual Studio ;
- scripts de personnages : désactivés par défaut, migration Python 3 en cours ;
- prochaine étape : tests d'exécution, corrections fonctionnelles et empaquetage des versions.

## Origine et licence

Le code original provient du projet [Vendetta sur SourceForge](https://sourceforge.net/projects/vendetta2/). Les sources restent distribuées selon les en-têtes d'origine, sous GNU GPL version 3 ou ultérieure. Consultez [LICENCE](LICENCE) et [THIRD_PARTY_ASSETS.md](THIRD_PARTY_ASSETS.md).
