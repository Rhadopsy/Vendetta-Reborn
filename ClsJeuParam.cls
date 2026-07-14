VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsJeuParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Vendetta - Standalone Role Playing game, simulate a MMORPG with AI
' Copyright (C) 2003-2007  S. HARLAUT (sharlaut@netcourrier.com)
'
' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
Option Explicit

Const FichierINI_Globales = "Parametres_Globales"
Const SectionINI_Globales = "Parametres_Globales"
Const FichierINI_Persos = "Parametres_Persos"
Const FichierINI_Fiefs = "Parametres_Fiefs"
Const FichierINI_Ressources = "Parametres_Ressources"
Const FichierINI_Batiments = "Parametres_Batiments"
Const SectionINI_Batiments = "Batiments"
Const FichierINI_Peuples = "Parametres_Peuples"
Const SectionINI_Peuples = "Peuples"
Const FichierINI_Tresors = "Tresors"
Const FichierINI_Services = "Services"
Const FichierINI_Competences_Speciales = "Competences_Speciales"
Const SectionINI_Competences_Speciales = "Competences_Speciales"
Const FichierINI_Resistances = "Parametres_Resistances"
Const SectionINI_Resistances = "Resistances"
Const SectionINI_Statuts = "Statuts"
Const FichierINI_CompetencesObjets = "Parametres_CompetencesObjets"
Const SectionINI_CompetencesObjets = "CompetencesObjets"
Const FichierINI_Objets = "Parametres_Objets"
Const SectionINI_Objets = "Objets"
Const FichierINI_Decors = "Parametres_Decors"
Const SectionINI_Decors = "Decors"
Const FichierINI_Effets = "Parametres_Effets"
Const SectionINI_Effets = "Effets"
Const FichierINI_Races = "Parametres_Races"
Const SectionINI_Races = "Races"

'Paramètres globales.
Public EtiquetteEpoque As String

Public EtiquetteNom As String
Public EtiquetteRace As String
Public EtiquetteRacesJouables As String
Public EtiquetteSexe As String
Public EtiquetteMasculin As String
Public EtiquetteFeminin As String
Public EtiquetteMetier As String
Public EtiquetteNiveau As String
Public EtiquetteNivPlus As String
Public EtiquetteEquipe As String

Public EtiquetteVieRestante As String
Public EtiquetteViesRestantes As String
Public EtiquetteVie As String
Public EtiquetteEnergie As String
Public EtiquetteMagie As String
Public EtiquetteMoral As String
Public EtiquetteAttaque As String
Public EtiquetteDefense As String
Public EtiquetteArmure As String
Public EtiquetteVitesse As String
Public EtiquetteDeplacement As String
Public EtiquetteVitalite As String

Public EtiquetteCompetence As String
Public EtiquetteCompetences As String
Public EtiquetteBestiaire As String
Public EtiquetteTresor As String
Public EtiquetteTresors As String
Public EtiquetteFortuneTotale As String
Public EtiquetteFortuneRestante As String
Public MessageFortuneManquante As String
Public EtiquetteMinimum As String
Public EtiquetteMinimumAbr As String
Public EtiquetteMaximum As String
Public EtiquetteMaximumAbr As String
Public EtiquetteMoyenne As String
Public EtiquetteTotal As String

Public EtiquetteMonstre As String
Public EtiquetteRegeneration As String
Public EtiquetteRegenerationVie As String
Public EtiquetteRegenerationEnergie As String
Public EtiquetteRegenerationMagie As String
Public EtiquetteRegenerationMoral As String
Public EtiquetteParole As String
Public EtiquetteMuet As String
Public EtiquetteVolant As String
Public EtiquettePasdExperience As String
Public EtiquetteCommentairesSpecifique As String
Public EtiquetteArme2mains As String
Public EtiquettePortee As String
Public EtiquettePorteeAttaque As String
Public EtiquetteSurfaceAttaque As String
Public EtiquetteUniteDistance As String
Public EtiquetteEffetAttaque As String
Public EtiquettePietinement As String
Public EtiquetteInfatiguable As String
Public EtiquetteRestaurationComplete As String
Public EtiquetteKamikaze As String
Public EtiquetteEquipement As String
Public EtiquetteSEquipe As String
Public EtiquetteMange As String
Public EtiquetteManger As String
Public EtiquettePacifique As String
Public EtiquetteTemeraire As String
Public EtiquetteBerserk As String
Public EtiquetteBerserker As String
Public EtiquetteRessuscite As String
Public EtiquetteCapacite As String
Public EtiquetteExperience As String
Public EtiquetteXP As String
Public EtiquetteIA As String
Public EtiquetteParticularite As String
Public EtiquetteResistances As String
Public EtiquetteUtilisationMontureImpossible As String
Public EtiquetteUtilisationMontureImpossible2 As String

Public EtiquetteObjets As String
Public EtiquetteArgent As String
Public EtiquettePieces As String
Public EtiquettePrix As String
Public EtiquetteIndestructible As String
Public EtiquetteImage As String
Public EtiquetteDisponibilite As String
Public EtiquetteCout As String
Public EtiquetteBonus As String
Public EtiquetteEffet As String
Public EtiquetteInvoque As String
Public EtiquetteInvoqueChance1 As String
Public EtiquetteInvoqueChance2 As String
Public EtiquetteInvoqueMax1 As String
Public EtiquetteInvoqueMax1Pluriel As String
Public EtiquetteInvoqueMax2 As String
Public EtiquetteInvoqueMax2Pluriel As String
Public EtiquettePersonne As String
Public EtiquettePersonnes As String
Public EtiquetteParUtilisation As String
Public EtiquetteGenere As String
Public EtiquetteRegenere As String
Public EtiquetteRedonne As String
Public EtiquetteAffecte As String
Public EtiquetteRessusciteSoi As String
Public EtiquetteTranport As String
Public EtiquetteRessourcesSupplementaires As String
Public EtiquettePourcent As String
Public EtiquetteUtilisation As String
Public EtiquetteUtilisationPossible As String
Public EtiquetteUtilisationsPossibles As String
Public EtiquetteUtilisationMunition As String
Public EtiquetteUtilisationMunitions As String
Public EtiquetteUtilisationsRessourcesInsuffisantes As String
Public EtiquetteUtilisationsMunitionsInsuffisantes As String
Public EtiquetteProjectilesSupplementaires As String
Public EtiquettePoison As String
Public EtiquetteObjetInvendable As String
Public EtiquetteObjetIndeposable As String
Public EtiquetteObjetAcheter As String
Public EtiquetteObjetVendre As String
Public EtiquetteObjetVente As String
Public EtiquetteObjetVenteImpossible As String
Public EtiquetteObjetDepot As String
Public EtiquetteObjetDepotImpossible As String
Public EtiquetteDureeRestant As String
Public EtiquetteEffetAnnulationImpossible As String
Public EtiquetteUtilisationUnique As String
Public EtiquetteUtilisationMort As String
Public EtiquetteDuree As String
Public EtiquetteDureeRecharge As String
Public EtiquetteDureeRechargeAttaque As String
Public EtiquetteRechargeMultiple As String
Public EtiquetteRechargesMultiples As String
Public EtiquetteUniteDuree As String
Public EtiquetteMonture As String

Public EtiquetteTombeMasculin As String
Public EtiquetteTombeFeminin As String

Public EtiquetteTaille As String
Public EtiquetteStock As String
Public EtiquetteFabrique As String
Public EtiquetteSecret As String
Public EtiquetteBatimentEmpereur As String
Public EtiquetteFermeM As String
Public EtiquetteFermeF As String
Public EtiquetteAbandonneM As String
Public EtiquetteAbandonneF As String

Public EtiquetteRessources As String
Public EtiquetteMatieresPremieres As String
Public EtiquetteValeur As String
Public EtiquetteVitesseRecolte As String
Public EtiquetteRegain As String
Public EtiquetteRessourceOrigine As String
Public EtiquetteRessourceObjets As String

Public EtiquettePeuple As String
Public EtiquetteForteresse As String
Public EtiquetteCreatures As String

Public EtiquettePersonnage As String
Public EtiquettePersonnages As String
Public EtiquetteBatiment As String
Public EtiquetteBatiments As String
Public EtiquetteChateau As String
Public EtiquetteChateaux As String
Public EtiquetteFrags As String
Public EtiquetteMorts As String
Public EtiquetteGainCaracs As String
Public EtiquetteGainCompetences As String

Public TitreErreur As String
Public MessageErreurFichier As String
Public MessageErreurFichierManquant As String
Public MessageErreurLigne As String

Private EtiquettesChargement() As String

'Ce paramètre est le rapport entre le prix d'achat et de vente.
Public CoefficientAchat As Double

'Parametres des personnages.
Public PersosLargeurDefaut As Integer
Public PersosHauteurDefaut As Integer

Public PersosMaxVie As Long
Public PersosMaxEnergie As Long
Public PersosMaxMagie As Long
Public PersosMaxMoral As Long

Public PersosNiveauSpecialesDepart As Long
Public PersosMaxRessources As Long
Public PersosNiveauServicesDepart As Long
Public PersosNiveauObjetsDepart As Long

Public PersosAttaque As Long
Public PersosPorteeAttaque As Long
Public PersosDefense As Long

Public PersosVitesse As Integer
Public PersosAcceleration As Integer

Public PersosNombreObjetsEquipes As Integer
Public PersosNombreObjetsInventaire As Integer
Public PersosNombreObjetsInventaireMaison As Integer

Public NombreEpoques As Integer
Private Epoques() As Epoque

Private Type Epoque
    Nom As String
    Duree As Long
    BonusVieChateau As Long
    EmpereurNomM As String
    EmpereurNomF As String
    EmpereurPrix As Long
End Type

Public NombreRessourcesDepart As Integer
Private RessourcesDeparts() As RessourcesDepart
Private Type RessourcesDepart
    Nom As String
    ChateauxRessourcesDepart As Double
    PersosRessourcesDepart As Double
    PersosArgentDepart As Double
End Type

Public NombreVitesseEpoques As Integer
Public VitesseEpoqueDefaut As Integer
Public VitesseEpoqueSelectionnee As Integer
Private VitesseEpoques() As VitesseEpoque
Private Type VitesseEpoque
    Nom As String
    Coef As Single
End Type

Private BatimentsTri() As Integer 'Enregistre les indices des bâtiments dans un ordre précis.
Private Type ParamBatiments
    ID As String
    Nom As String
    Feminin As Boolean
    Jouable As Boolean
    Image As String
    CheminImage As String
    ImageBouton As String
    CheminImageBouton As String
    MetierM As String
    MetierF As String
    Description As String
    Empereur As Boolean 'Seul l'empereur peut construire ce bâtiment.
    Limite As Integer 'Limite du nombre de bâtiments de ce type par fief.
    CreerProprietaire As Integer 'A la création, c'est une invocation qui est le propriétaire du bâtiment au lieu du joueur.
    Roi As Boolean 'On devient le roi quand on a construit ce bâtiment.
    RestrictionRoi As Boolean 'Seul le roi peut construire ce type de bâtiment.
    GainObjet As Integer 'Offre un objet au propriétaire dès que le bâtiment est construit.
    PerteObjet As Integer 'Supprime tous les objets de ce type au propriétaire.
    Tir As Boolean 'Ce batiment permet de tirer de l'intérieur sans être obligé de sortir.
    Marche As Boolean 'On peut vendre et acheter des ressources dans ce bâtiment.
    Prix As Long 'Coût total du bâtiment. (Utilisée pour trier l'ordre d'affichage de ceux-ci)
    PrixArgent As Long 'Prix de fabrication du bâtiment en pièces d'argent.
    PrixRessources() As Long
    MaxVie As Long
    CoefVitesseConstruction As Single
    RessourceFabrique As Long
    NombreObjetsFabriques As Integer
    ObjetsFabriques() As Integer
    Magasin As Boolean
    Fabrique As Boolean
    Service As Integer
    MaxStock As Double
    Niveau As Long 'Degrés qui permet à l'IA de connaître l'importance de la maison.
    NoEpoque As Integer 'Numéro de l'époque à laquelle on peut construire le batiment.
    Largeur As Integer
    Hauteur As Integer
    EntreePositionX As Long
    EntreePositionY As Long
    NombreEtapeFabrication As Integer
    NombreEtapeAnimation As Integer
    ApparencePerso As Integer 'Inutilisé.
    FichierMod As Boolean 'Si vrai, cet élément appartient à un fichier mod.
End Type
Private Type ParamRessources
    Nom As String
    NomCompetence As String
    Image As String 'Nom du fichier image.
    CheminImage As String 'Indique le chemin complet de l'image
    VitesseExtraction As Double
    TypeMatierePremiere As Long
    QuantiteMatierePremiere As Double
    PrixVente As Double
    NecessairePourConstruire As Boolean
    NecessairePourAttaquer As Boolean
    Recoltable As Boolean
    TypeTerrainRessource As Integer
    Mangeable As Boolean
    GainVie As Double
    GainEnergie As Double
    GainMagie As Double
    GainMoral As Double
    Epoque As Integer
    FichierMod As Boolean 'Si vrai, cet élément appartient à un fichier mod.
End Type
Private Type ParamTerrainRessources
    Nom As String
    TypeRessource As Integer
    Quantite As Double
End Type
Private Type ParamPeuples
    Nom As String
    Image As String
    CheminImage As String 'Indique le chemin complet de l'image
    NombreMonstres As Integer
    MonstreIndice() As Integer
    MonstreChance() As Single
    MonstreFichierIA() As String
    NomChateau As String
    ChateauMaxVie As Long
    ChateauLargeur As Integer
    ChateauHauteur As Integer
    ChateauEntreePositionX As Integer
    ChateauEntreePositionY As Integer
    FichierMod As Boolean 'Si vrai, cet élément appartient à un fichier mod.
End Type
Private Type ParamCompetencesObjets
    ID As String 'Identifiant unique.
    SAVPerso As String
    NoListe As Integer
    Nom As String
    FichierMod As Boolean 'Si vrai, cet élément appartient à un fichier mod.
End Type
Private Type ParamObjets
    Nom As String
    NomEffetM As String 'Intitulé de l'effet temporaire pour les hommes.
    NomEffetF As String 'Intitulé de l'effet temporaire pour les femmes.
    Image As String 'Nom du fichier image.
    CheminImage As String 'Indique le chemin complet de l'image
    NomPluriel As String
    Description As String 'Intitulé supplémentaire dans l'info bulle.
    Secret As Boolean 'L'objet ne fait pas parti de la liste des objets fabricables.
    Unique As Boolean 'L'objet n'est pas sauvegardé dans la partie.
    Indestructible As Boolean 'L'objet ne peut être détruit ou pris.
    Inseparable As Boolean 'L'objet ne peut être donné, déposé ou vendu.
    Superposable As Integer 'Nombre d'objets que l'on peut superposer au même endroit.
    Empereur As Boolean 'Si vrai, c'est un objet spécifique à l'empereur.
    Categorie As Integer
    Competence As Integer
    PrixRessources() As Integer
    Epoque As Integer
    DureeFabrication As Integer
    CoefXPFabrication As Single
    PrixVente As Integer
    BonusVie As Integer
    BonusEnergie As Integer
    BonusMagie As Integer
    BonusMoral As Integer
    BonusAttaque As Integer
    BonusPorteeAttaque As Integer
    LanceProjectile As Integer
    ProjectileSupplementaire As Single 'Nombre de projectiles supplémentaires.
    EffetAttaque As Integer
    AttaqueResistance As Integer
    SurfaceAttaque As Integer 'L'arme fait toute une zone de dégâts.
    AttaqueAutomatique As Boolean 'L'arme attaque sans arrêt.
    AttaqueCauseEffet As Integer 'En touchant, l'arme cause l'effet temporaire d'indice x.
    BonusDefense As Integer
    BonusDeplacement As Single 'Augmente la vitesse pour les non volant uniquement.
    BonusVitesse As Single 'Augmente la vitesse pour tout le monde.
    BonusVitalite As Single 'Augmente la vitalité en pourcentage.
    BonusCompetencesSpeciales() As Integer
    BonusCompetencesRessources() As Integer
    BonusCompetencesServices() As Integer
    BonusCompetencesObjets() As Integer
    BonusResistances() As Single 'Augmente les résistances.
    BonusArmure As Integer
    BonusMaxRessources As Integer
    BonusVitesseRegenerationVie As Single
    BonusVitesseRegenerationEnergie As Single
    BonusVitesseRegenerationMagie As Single
    BonusVitesseRegenerationMoral As Single
    PerteVie As Double
    PerteEnergie As Double
    PerteMagie As Double
    PerteMoral As Double
    PerteArgent As Long
    PerteRessourceIndice As Integer
    PerteRessourceQuantite As Single
    Poison As Single
    ApparencePerso As Integer
    FichierApparenceM As String
    FichierApparenceF As String
    NumeroApparenceM As Integer 'Variable temporaire créé automatiquement.
    NumeroApparenceF As Integer 'Variable temporaire créé automatiquement.
    ApparenceLargeur As Integer 'Largeur de l'image du fichier apparence.
    ApparenceHauteur As Integer 'Hauteur de l'image du fichier apparence.
    ApparenceDecalageX() As Integer
    ApparenceDecalageY() As Integer
    ApparenceAvantPlan() As Boolean 'Place l'apparence en question en arrière plan quand le personange monte.
    ChangerApparencePerso As Boolean 'Métamorphose le personnage.
    Utilisable As Boolean 'Si vrai, l'objet s'utilise au lieu d'être équipé.
    UtilisationUnique As Boolean 'L'objet disparaît après utilisation.
    UtilisationMort As Boolean 'Utilisable même par un personnage mort.
    PorteeUtilisation As Integer 'Portee des effets après utilisation.
    Experience As Double 'Nombre de points d'expérience que l'utilisation rapporte.
    Invocation As Integer
    InvocationIA_Fichier As String 'Charge un fichier d'IA pour l'invocation en question.
    Quantite As Single
    QuantiteMax As Single 'Indique un seuil à la quantité.
    GainVie As Double
    GainEnergie As Double
    GainMagie As Double
    GainMoral As Double
    GainArgent As Single
    Impot As Boolean
    GainRessource() As Single
    Resurrection As Long 'Ressuscite X personnes selon la portée.
    Berserk As Integer 'Le porteur de l'objet passe en berserk.
    SelfResurrection As Boolean 'Ressuscute l'utilisateur de l'objet.
    CreerObjet As Integer 'Crée un objet ou un effet temporaire selon la durée.
    CreerObjetAllie As Integer 'Crée un objet ou un effet temporaire sur des alliés proches selon la durée et la portée.
    CreerObjetEnnemi As Integer 'Crée un objet ou un effet temporaire sur des adversaires proches selon la durée et la portée.
    Duree As Long 'Duree de l'effet temporaire. Si 0 : l'objet est permanent.
    DureeRecharge As Long 'Duree avant que l'on puisse réutiliser cet objet.
    DureeRechargeAttaque As Long 'Durée avant que l'on puisse attaquer à nouveau.
    RechargesMultiples As Single 'Nombre de fois que l'on peut lancer ce pouvoir de suite.
    EffetAnnuleRestaureObjet As Boolean 'Restaure l'objet en question si l'effet est annulé.
    EffetAnnulationImpossible As Boolean 'Le joueur ne peut pas mettre fin à cet effet manuellement.
    EffetAnnulationMort As Boolean 'L'effet disparaît à la mort du personnage.
    UtilisationAutomatique As Boolean 'L'objet se lance automatiquement en début de partie.
    Monture As Boolean 'Si vrai, l'objet est considéré comme une monture avec ces restrictions éventuelles.
    Volant As Boolean
    PrixTalents As Integer 'Coût en points de trésors.
    FichierMod As Boolean 'Si vrai, cet élément appartient à un fichier mod.
End Type
Private Type ParamTresors
    Nom As String
    ValeurMin As Integer
End Type
Private Type ParamDecors
    Nom As String
    Image As String
    Largeur As Integer
    Hauteur As Integer
    NombreEtapeAnimation As Integer
    Duree As Integer
End Type
Private Type ParamServices
    Nom As String
    Description As String
    VitesseCreation As Double
    TypeMatierePremiere As Integer
    QuantiteMatierePremiere As Double
    Prix As String
    Coefficient As Single
    NomCompetence As String
    CoefExp As Double
    Effet As Integer
    FichierMod As Boolean 'Si vrai, cet élément appartient à un fichier mod.
End Type
Private Type ParamEffets
    Nom As String
    Fichier As String
    Largeur As Integer
    Hauteur As Integer
    NombreEtapes As Integer
    DureeEtape As Double
    Son As String
End Type
Private Type ParamRaces
    NomM As String
    NomF As String
    NomMPluriel As String
    NomFPluriel As String
    NomAleatoire As Boolean
    Jouable As Boolean 'Si vrai, les joueurs peuvent jouer avec.
    Description As String
    RestrictionMaculin As Boolean
    RestrictionFeminin As Boolean
    Largeur As Integer 'Largeur du sprite.
    Hauteur As Integer 'Hauteur du sprite.
    Portrait As String 'Fichier image du portait.
    NombreCategoriesApparences As Integer
    DossiersCategoriesF() As String
    DossiersCategoriesM() As String
    DefinirApparence As Boolean
    CheminApparenceM As String
    CheminApparenceF As String
    CheminAbsolu As Boolean 'Si vrai, le chemin de l'apparence est complet.
    ApparenceVerouillee As Boolean 'Si vrai, l'apparence du personnage ne change pas durant le jeu.
    Vie As Long
    Energie As Long
    Magie As Long
    Moral As Long
    Attaque As Long
    Defense As Long
    CoefVie As Single
    CoefEnergie As Single
    CoefMagie As Single
    CoefMoral As Single
    CoefAttaque As Single
    CoefDefense As Single
    VitesseRegenerationVie As Single
    VitesseRegenerationEnergie As Single
    VitesseRegenerationMagie As Single
    VitesseRegenerationMoral As Single
    Parole As Boolean
    Experience As Boolean
    Equiper As Integer
    CoefCompSpeciales() As Single 'Compétences spéciales.
    CoefCompRessources() As Single 'Compétences des ressources.
    CoefCompServices() As Single 'Compétences des Services.
    CoefCompObjets() As Single 'Compétences des objets.
    Resistances() As Single 'Résistances.
    Volant As Boolean
    CommentairesSpecifique As Integer 'Utilise qu'un commentaire pour toutes ces phrases.
    BonusMaxRessources As Integer 'Transporte plus de ressources.
    PorteeAttaque As Integer
    EffetAttaque As Integer
    AttaqueCauseEffet As Integer
    AttaqueResistance As Integer 'Indique quelle résistance est opposée à l'attaque.
    Armure As Integer
    Vitesse As Integer
    Acceleration As Integer
    'Projectile As Boolean
    Pietinement As Integer
    Inattaquable As Boolean
    AutoResurrection As Boolean
    IA_Pacifique As Boolean
    IA_Manger As Boolean
    IA_Temeraire As Boolean
    IA_Protecteur As Boolean
    IA_Egoiste As Boolean
    IA_Berserk As Integer
    IA_Incommandable As Boolean
    IA_Selectionner As Boolean 'Vrai, on sélectionne une IA parmi les fichiers d'IA.
    IA_Fichier As String 'Indique le fichier d'IA sélectionné par défaut.
    Infatigable As Boolean 'Si vrai, le personnage ne perd plus d'énergie ou de moral.
    Kamikaze As Boolean
    RestaurationComplete As Boolean 'Récupère toutes ses caractéristiques quand il passe au château.
    Argent As Double
    Ressources() As Double
    DepartSansRessources As Boolean
    NombreObjets As Integer
    Objet() As Integer
    ObjetRace As Collection 'Mémorise les objets uniquement propres à la race (N'éxiste pas les fichiers ini)
    NombreEquipement As Integer
    Equipement() As Integer
    Statut As Collection 'Indique tous les statuts en rapport aux resistances.
    UtilisationMonture As Boolean 'Si vrai, le personnage peut monter sur une monture.
    IndiceBestiaire As Integer 'Numéro de bestiaire enregistrer dans le fichier personnage.
    EffetMort As Integer 'Effet visuel quand le personnage est tué.
    Tombe As Integer 'Numéro du décor qui correspond à sa tombe.
    FichierMod As Boolean 'Si vrai, cet élément appartient à un fichier mod.
End Type

'Paramètres des résistances.
Public ResistanceDefaut As Single
Public ResistanceMinimum As Single
Public ResistanceMaximum As Single
Private Type ParamResistances
    Nom As String
    NomComplet As String
    Attaque As String
    Abreviation As String
End Type
Public NombreResistances As Integer
Private Resistances() As ParamResistances
Private Type ParamStatuts
    Nom As String
    Resistances() As Single
End Type
Public NombreStatuts As Integer
Private Statuts() As ParamStatuts

Public NombreCompetencesSpeciales As Integer
Private CompetencesSpeciales() As String

Public NombreRessources As Integer
Private Ressources() As ParamRessources

Public NombreTerrainRessource As Integer
Private TerrainRessources() As ParamTerrainRessources

Public NombreTypeBatiments As Integer
Public NombreTypeBatimentsJouables As Integer
Public BatimentsCheminImageDefaut As String
Private Batiments() As ParamBatiments

Public NombrePeuples As Integer
Public PeuplesCheminImageDefaut As String
Private Peuples() As ParamPeuples

Public NombreCompetencesObjets As Integer
Private CompetencesObjets() As ParamCompetencesObjets

Public NombreObjets As Integer
Private Objets() As ParamObjets

Public NombreTresors As Integer
Private Tresors() As ParamTresors
Public NombreDecors As Integer
Private Decors() As ParamDecors

Public NombreServices As Integer
Private Services() As ParamServices

Public NombreEffets As Integer
Private Effets() As ParamEffets
Public EffetClique As Integer
Public EffetMaisonEnFeu As Integer
Public EffetMaisonDetruite As Integer
Public EffetChateauDetruit As Integer
Public EffetNouvelObjectif As Integer
Public EffetDureeRestante As Integer
Public EffetDureeRestante2 As Integer

Public NombreRaces As Integer
Public NombreRacesJouables As Integer
Public ChancesRaceAleatoire As Single
Public JouerToutesLesRaces As Boolean
Private Races() As ParamRaces
Private NombreObjetsCommmuns As Integer
Private ObjetCommmuns() As Integer

Private BestiaireOrdreAlpha() As Integer 'Contient l'ordre alphabétique du bestiaire.

Private Sub Class_Initialize()
    Dim i As Integer
    
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI_Globales
    FicIni.Section = SectionINI_Globales
    'Paramètres globales.
    EtiquetteEpoque = FicIni.Parametre("EtiquetteEpoque")
    
    EtiquetteNom = FicIni.Parametre("EtiquetteNom")
    EtiquetteRace = FicIni.Parametre("EtiquetteRace")
    EtiquetteRacesJouables = FicIni.Parametre("EtiquetteRacesJouables")
    EtiquetteSexe = FicIni.Parametre("EtiquetteSexe")
    EtiquetteMasculin = FicIni.Parametre("EtiquetteMasculin")
    EtiquetteFeminin = FicIni.Parametre("EtiquetteFeminin")
    EtiquetteMetier = FicIni.Parametre("EtiquetteMetier")
    EtiquetteNiveau = FicIni.Parametre("EtiquetteNiveau")
    EtiquetteNivPlus = FicIni.Parametre("EtiquetteNivPlus")
    EtiquetteEquipe = FicIni.Parametre("EtiquetteEquipe")
    
    EtiquetteVieRestante = FicIni.Parametre("EtiquetteVieRestante")
    EtiquetteViesRestantes = FicIni.Parametre("EtiquetteViesRestantes")
    EtiquetteVie = FicIni.Parametre("EtiquetteVie")
    EtiquetteEnergie = FicIni.Parametre("EtiquetteEnergie")
    EtiquetteMagie = FicIni.Parametre("EtiquetteMagie")
    EtiquetteMoral = FicIni.Parametre("EtiquetteMoral")
    EtiquetteAttaque = FicIni.Parametre("EtiquetteAttaque")
    EtiquetteDefense = FicIni.Parametre("EtiquetteDefense")
    EtiquetteArmure = FicIni.Parametre("EtiquetteArmure")
    EtiquetteVitesse = FicIni.Parametre("EtiquetteVitesse")
    EtiquetteDeplacement = FicIni.Parametre("EtiquetteDeplacement")
    EtiquetteVitalite = FicIni.Parametre("EtiquetteVitalite")
    
    EtiquetteCompetence = FicIni.Parametre("EtiquetteCompetence")
    EtiquetteCompetences = FicIni.Parametre("EtiquetteCompetences")
    EtiquetteBestiaire = FicIni.Parametre("EtiquetteBestiaire")
    EtiquetteTresor = FicIni.Parametre("EtiquetteTresor")
    EtiquetteTresors = FicIni.Parametre("EtiquetteTresors")
    EtiquetteFortuneTotale = FicIni.Parametre("EtiquetteFortuneTotale")
    EtiquetteFortuneRestante = FicIni.Parametre("EtiquetteFortuneRestante")
    MessageFortuneManquante = FicIni.Parametre("MessageFortuneManquante")
    EtiquetteMinimum = FicIni.Parametre("EtiquetteMinimum")
    EtiquetteMinimumAbr = FicIni.Parametre("EtiquetteMinimumAbr")
    EtiquetteMaximum = FicIni.Parametre("EtiquetteMaximum")
    EtiquetteMaximumAbr = FicIni.Parametre("EtiquetteMaximumAbr")
    EtiquetteMoyenne = FicIni.Parametre("EtiquetteMoyenne")
    EtiquetteTotal = FicIni.Parametre("EtiquetteTotal")
    
    EtiquetteMonstre = FicIni.Parametre("EtiquetteMonstre")
    EtiquetteRegeneration = FicIni.Parametre("EtiquetteRegeneration")
    EtiquetteRegenerationVie = FicIni.Parametre("EtiquetteRegenerationVie")
    EtiquetteRegenerationEnergie = FicIni.Parametre("EtiquetteRegenerationEnergie")
    EtiquetteRegenerationMagie = FicIni.Parametre("EtiquetteRegenerationMagie")
    EtiquetteRegenerationMoral = FicIni.Parametre("EtiquetteRegenerationMoral")
    EtiquetteParole = FicIni.Parametre("EtiquetteParole")
    EtiquetteMuet = FicIni.Parametre("EtiquetteMuet")
    EtiquetteVolant = FicIni.Parametre("EtiquetteVolant")
    EtiquettePasdExperience = FicIni.Parametre("EtiquettePasdExperience")
    EtiquetteCommentairesSpecifique = FicIni.Parametre("EtiquetteCommentairesSpecifique")
    EtiquetteArme2mains = FicIni.Parametre("EtiquetteArme2mains")
    EtiquettePortee = FicIni.Parametre("EtiquettePortee")
    EtiquettePorteeAttaque = FicIni.Parametre("EtiquettePorteeAttaque")
    EtiquetteSurfaceAttaque = FicIni.Parametre("EtiquetteSurfaceAttaque")
    EtiquetteUniteDistance = FicIni.Parametre("EtiquetteUniteDistance")
    EtiquetteEffetAttaque = FicIni.Parametre("EtiquetteEffetAttaque")
    EtiquettePietinement = FicIni.Parametre("EtiquettePietinement")
    EtiquetteInfatiguable = FicIni.Parametre("EtiquetteInfatiguable")
    EtiquetteRestaurationComplete = FicIni.Parametre("EtiquetteRestaurationComplete")
    EtiquetteKamikaze = FicIni.Parametre("EtiquetteKamikaze")
    EtiquetteEquipement = FicIni.Parametre("EtiquetteEquipement")
    EtiquetteSEquipe = FicIni.Parametre("EtiquetteSEquipe")
    EtiquetteMange = FicIni.Parametre("EtiquetteMange")
    EtiquetteManger = FicIni.Parametre("EtiquetteManger")
    EtiquettePacifique = FicIni.Parametre("EtiquettePacifique")
    EtiquetteTemeraire = FicIni.Parametre("EtiquetteTemeraire")
    EtiquetteBerserk = FicIni.Parametre("EtiquetteBerserk")
    EtiquetteBerserker = FicIni.Parametre("EtiquetteBerserker")
    EtiquetteRessuscite = FicIni.Parametre("EtiquetteRessuscite")
    EtiquetteCapacite = FicIni.Parametre("EtiquetteCapacite")
    EtiquetteExperience = FicIni.Parametre("EtiquetteExperience")
    EtiquetteXP = FicIni.Parametre("EtiquetteXP")
    EtiquetteIA = FicIni.Parametre("EtiquetteIA")
    EtiquetteParticularite = FicIni.Parametre("EtiquetteParticularite")
    EtiquetteResistances = FicIni.Parametre("EtiquetteResistances")
    EtiquetteUtilisationMontureImpossible = FicIni.Parametre("EtiquetteUtilisationMontureImpossible")
    EtiquetteUtilisationMontureImpossible2 = FicIni.Parametre("EtiquetteUtilisationMontureImpossible2")
    
    EtiquetteObjets = FicIni.Parametre("EtiquetteObjets")
    EtiquetteArgent = FicIni.Parametre("EtiquetteArgent")
    EtiquettePieces = FicIni.Parametre("EtiquettePieces")
    EtiquettePrix = FicIni.Parametre("EtiquettePrix")
    EtiquetteIndestructible = FicIni.Parametre("EtiquetteIndestructible")
    EtiquetteImage = FicIni.Parametre("EtiquetteImage")
    EtiquetteDisponibilite = FicIni.Parametre("EtiquetteDisponibilite")
    EtiquetteCout = FicIni.Parametre("EtiquetteCout")
    EtiquetteBonus = FicIni.Parametre("EtiquetteBonus")
    EtiquetteEffet = FicIni.Parametre("EtiquetteEffet")
    EtiquetteInvoque = FicIni.Parametre("EtiquetteInvoque")
    EtiquetteInvoqueChance1 = FicIni.Parametre("EtiquetteInvoqueChance1")
    EtiquetteInvoqueChance2 = FicIni.Parametre("EtiquetteInvoqueChance2")
    EtiquetteInvoqueMax1 = FicIni.Parametre("EtiquetteInvoqueMax1")
    EtiquetteInvoqueMax1Pluriel = FicIni.Parametre("EtiquetteInvoqueMax1Pluriel")
    EtiquetteInvoqueMax2 = FicIni.Parametre("EtiquetteInvoqueMax2")
    EtiquetteInvoqueMax2Pluriel = FicIni.Parametre("EtiquetteInvoqueMax2Pluriel")
    EtiquettePersonne = FicIni.Parametre("EtiquettePersonne")
    EtiquettePersonnes = FicIni.Parametre("EtiquettePersonnes")
    EtiquetteParUtilisation = FicIni.Parametre("EtiquetteParUtilisation")
    EtiquetteGenere = FicIni.Parametre("EtiquetteGenere")
    EtiquetteRegenere = FicIni.Parametre("EtiquetteRegenere")
    EtiquetteRedonne = FicIni.Parametre("EtiquetteRedonne")
    EtiquetteAffecte = FicIni.Parametre("EtiquetteAffecte")
    EtiquetteRessusciteSoi = FicIni.Parametre("EtiquetteRessusciteSoi")
    EtiquetteTranport = FicIni.Parametre("EtiquetteTranport")
    EtiquetteRessourcesSupplementaires = FicIni.Parametre("EtiquetteRessourcesSupplementaires")
    EtiquettePourcent = FicIni.Parametre("EtiquettePourcent")
    EtiquetteUtilisation = FicIni.Parametre("EtiquetteUtilisation")
    EtiquetteUtilisationPossible = FicIni.Parametre("EtiquetteUtilisationPossible")
    EtiquetteUtilisationsPossibles = FicIni.Parametre("EtiquetteUtilisationsPossibles")
    EtiquetteUtilisationMunition = FicIni.Parametre("EtiquetteUtilisationMunition")
    EtiquetteUtilisationMunitions = FicIni.Parametre("EtiquetteUtilisationMunitions")
    EtiquetteUtilisationsRessourcesInsuffisantes = FicIni.Parametre("EtiquetteUtilisationsRessourcesInsuffisantes")
    EtiquetteUtilisationsMunitionsInsuffisantes = FicIni.Parametre("EtiquetteUtilisationsMunitionsInsuffisantes")
    EtiquetteProjectilesSupplementaires = FicIni.Parametre("EtiquetteProjectilesSupplementaires")
    EtiquettePoison = FicIni.Parametre("EtiquettePoison")
    EtiquetteObjetInvendable = FicIni.Parametre("EtiquetteObjetInvendable")
    EtiquetteObjetIndeposable = FicIni.Parametre("EtiquetteObjetIndeposable")
    EtiquetteObjetAcheter = FicIni.Parametre("EtiquetteObjetAcheter")
    EtiquetteObjetVendre = FicIni.Parametre("EtiquetteObjetVendre")
    EtiquetteObjetVente = FicIni.Parametre("EtiquetteObjetVente")
    EtiquetteObjetVenteImpossible = FicIni.Parametre("EtiquetteObjetVenteImpossible")
    EtiquetteObjetDepot = FicIni.Parametre("EtiquetteObjetDepot")
    EtiquetteObjetDepotImpossible = FicIni.Parametre("EtiquetteObjetDepotImpossible")
    EtiquetteDureeRestant = FicIni.Parametre("EtiquetteDureeRestant")
    EtiquetteEffetAnnulationImpossible = FicIni.Parametre("EtiquetteEffetAnnulationImpossible")
    EtiquetteUtilisationUnique = FicIni.Parametre("EtiquetteUtilisationUnique")
    EtiquetteUtilisationMort = FicIni.Parametre("EtiquetteUtilisationMort")
    EtiquetteDuree = FicIni.Parametre("EtiquetteDuree")
    EtiquetteDureeRecharge = FicIni.Parametre("EtiquetteDureeRecharge")
    EtiquetteDureeRechargeAttaque = FicIni.Parametre("EtiquetteDureeRechargeAttaque")
    EtiquetteRechargeMultiple = FicIni.Parametre("EtiquetteRechargeMultiple")
    EtiquetteRechargesMultiples = FicIni.Parametre("EtiquetteRechargesMultiples")
    EtiquetteUniteDuree = FicIni.Parametre("EtiquetteUniteDuree")
    EtiquetteMonture = FicIni.Parametre("EtiquetteMonture")
    
    EtiquetteTombeMasculin = FicIni.Parametre("EtiquetteTombeMasculin")
    EtiquetteTombeFeminin = FicIni.Parametre("EtiquetteTombeFeminin")
    
    EtiquetteTaille = FicIni.Parametre("EtiquetteTaille")
    EtiquetteStock = FicIni.Parametre("EtiquetteStock")
    EtiquetteFabrique = FicIni.Parametre("EtiquetteFabrique")
    EtiquetteSecret = FicIni.Parametre("EtiquetteSecret")
    EtiquetteBatimentEmpereur = FicIni.Parametre("EtiquetteBatimentEmpereur")
    EtiquetteFermeM = FicIni.Parametre("EtiquetteFermeM")
    EtiquetteFermeF = FicIni.Parametre("EtiquetteFermeF")
    EtiquetteAbandonneM = FicIni.Parametre("EtiquetteAbandonneM")
    EtiquetteAbandonneF = FicIni.Parametre("EtiquetteAbandonneF")
    
    EtiquetteRessources = FicIni.Parametre("EtiquetteRessources")
    EtiquetteMatieresPremieres = FicIni.Parametre("EtiquetteMatieresPremieres")
    EtiquetteValeur = FicIni.Parametre("EtiquetteValeur")
    EtiquetteVitesseRecolte = FicIni.Parametre("EtiquetteVitesseRecolte")
    EtiquetteRegain = FicIni.Parametre("EtiquetteRegain")
    EtiquetteRessourceOrigine = FicIni.Parametre("EtiquetteRessourceOrigine")
    EtiquetteRessourceObjets = FicIni.Parametre("EtiquetteRessourceObjets")
    
    EtiquettePeuple = FicIni.Parametre("EtiquettePeuple")
    EtiquetteForteresse = FicIni.Parametre("EtiquetteForteresse")
    EtiquetteCreatures = FicIni.Parametre("EtiquetteCreatures")
    
    EtiquettePersonnage = FicIni.Parametre("EtiquettePersonnage")
    EtiquettePersonnages = FicIni.Parametre("EtiquettePersonnages")
    EtiquetteBatiment = FicIni.Parametre("EtiquetteBatiment")
    EtiquetteBatiments = FicIni.Parametre("EtiquetteBatiments")
    EtiquetteChateau = FicIni.Parametre("EtiquetteChateau")
    EtiquetteChateaux = FicIni.Parametre("EtiquetteChateaux")
    EtiquetteFrags = FicIni.Parametre("EtiquetteFrags")
    EtiquetteMorts = FicIni.Parametre("EtiquetteMorts")
    EtiquetteGainCaracs = FicIni.Parametre("EtiquetteGainCaracs")
    EtiquetteGainCompetences = FicIni.Parametre("EtiquetteGainCompetences")
    
    TitreErreur = FicIni.Parametre("TitreErreur")
    MessageErreurFichier = FicIni.Parametre("MessageErreurFichier")
    MessageErreurFichierManquant = FicIni.Parametre("MessageErreurFichierManquant")
    MessageErreurLigne = FicIni.Parametre("MessageErreurLigne")
    
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierChargement
    FicIni.Section = SectionChargement2
    ReDim EtiquettesChargement(13)
    For i = 0 To UBound(EtiquettesChargement())
        EtiquettesChargement(i) = FicIni.Parametre("EtiquetteChargement" & i + 1)
    Next i
    
    FicIni.Section = SectionChargement3
    CheminNotePadIntrouvable = FicIni.Parametre("CheminNotePadIntrouvable")
    CheminExplorerIntrouvable = FicIni.Parametre("CheminExplorerIntrouvable")
    CheminIExplorerIntrouvable = FicIni.Parametre("CheminIExplorerIntrouvable")
    CheminFireFoxIntrouvable = FicIni.Parametre("CheminFireFoxIntrouvable")
    
    NombreObjets = -1
End Sub

Public Sub Init(ByVal Index As Integer)
    'On Error GoTo Erreur
    Dim i As Long, j As Long, k As Long
    Select Case Index
    Case 0:
        FicIni.Fichier = FicIni.Chemin & FichierINI_Persos
        
        FicIni.Section = "Transactions"
        CoefficientAchat = FicIni.Parametre("CoefficientAchat")
        
        'Chargement des paramètres des personnages.
        FicIni.Section = "Personnages"
        
        PersosLargeurDefaut = Val(FicIni.Parametre("PersosLargeurDefaut"))
        PersosHauteurDefaut = Val(FicIni.Parametre("PersosHauteurDefaut"))
        
        PersosMaxVie = FicIni.Parametre("PersosMaxVie")
        PersosMaxEnergie = FicIni.Parametre("PersosMaxEnergie")
        PersosMaxMagie = FicIni.Parametre("PersosMaxMagie")
        PersosMaxMoral = FicIni.Parametre("PersosMaxMoral")
        
        PersosNiveauSpecialesDepart = FicIni.Parametre("PersosNiveauSpecialesDepart")
        PersosMaxRessources = FicIni.Parametre("PersosMaxRessources")
        PersosNiveauServicesDepart = FicIni.Parametre("PersosNiveauServicesDepart")
        PersosNiveauObjetsDepart = FicIni.Parametre("PersosNiveauObjetsDepart")
        
        PersosAttaque = FicIni.Parametre("PersosAttaque")
        'Portée de l'attaque à mains-nues.
        PersosPorteeAttaque = FicIni.Parametre("PersosPorteeAttaque")
        PersosDefense = FicIni.Parametre("PersosDefense")
        
        PersosVitesse = FicIni.Parametre("PersosVitesse")
        PersosAcceleration = FicIni.Parametre("PersosAcceleration")
        
        PersosNombreObjetsEquipes = FicIni.Parametre("PersosNombreObjetsEquipes")
        PersosNombreObjetsInventaire = FicIni.Parametre("PersosNombreObjetsInventaire")
        PersosNombreObjetsInventaireMaison = FicIni.Parametre("PersosNombreObjetsInventaireMaison")
        
        FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI_Competences_Speciales
        'Chargement des compétences spéciales.
        FicIni.Section = SectionINI_Competences_Speciales
        NombreCompetencesSpeciales = FicIni.Parametre("NombreCompetencesSpeciales")
        ReDim CompetencesSpeciales(NombreCompetencesSpeciales - 1)
        For i = 0 To UBound(CompetencesSpeciales())
            CompetencesSpeciales(i) = FicIni.Parametre("CompetencesSpeciales" & i + 1)
        Next i
        
    Case 1:
        FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI_Fiefs
        'Chargement des chateaux.
        'FicIni.Section = "Chateaux"
        'ChateauxMaxVie = FicIni.Parametre("ChateauxMaxVie")
        'ChateauxLargeur = FicIni.Parametre("ChateauxLargeur")
        'ChateauxHauteur = FicIni.Parametre("ChateauxHauteur")
        
        'Chargement des epoques.
        FicIni.Section = "Epoques"
        NombreEpoques = FicIni.Parametre("NombreEpoques")
        ReDim Epoques(NombreEpoques - 1)
        For i = 0 To NombreEpoques - 1
            FicIni.Section = "Epoque_" & i + 1
            Epoques(i).Nom = FicIni.Parametre("Nom")
            Epoques(i).Duree = FicIni.Parametre("Duree")
            Epoques(i).BonusVieChateau = FicIni.Parametre("BonusVieChateau")
            Epoques(i).EmpereurNomM = FicIni.Parametre("EmpereurNomM")
            Epoques(i).EmpereurNomF = FicIni.Parametre("EmpereurNomF")
            Epoques(i).EmpereurPrix = FicIni.Parametre("EmpereurPrix")
        Next i
        
        'Chargement des ressources de départ.
        FicIni.Section = "RessourcesDeparts"
        NombreRessourcesDepart = FicIni.Parametre("NombreRessourcesDeparts")
        ReDim RessourcesDeparts(NombreRessourcesDepart - 1)
        For i = 0 To NombreRessourcesDepart - 1
            FicIni.Section = "RessourcesDepart_" & i + 1
            RessourcesDeparts(i).Nom = FicIni.Parametre("Nom")
            RessourcesDeparts(i).ChateauxRessourcesDepart = Val(FicIni.Parametre("ChateauxRessourcesDepart"))
            RessourcesDeparts(i).PersosRessourcesDepart = Val(FicIni.Parametre("PersosRessourcesDepart"))
            RessourcesDeparts(i).PersosArgentDepart = Val(FicIni.Parametre("PersosArgentDepart"))
        Next i
        
        'Chargement des vitesses d'époques.
        FicIni.Section = "VitesseEpoques"
        NombreVitesseEpoques = FicIni.Parametre("NombreVitesseEpoques")
        ReDim VitesseEpoques(NombreVitesseEpoques - 1)
        For i = 0 To NombreVitesseEpoques - 1
            FicIni.Section = "VitesseEpoque_" & i + 1
            VitesseEpoques(i).Nom = FicIni.Parametre("Nom")
            VitesseEpoques(i).Coef = Val(FicIni.Parametre("Coef"))
        Next i
        
    Case 2:
        'FicIni.Fichier = FichierINI_Ressources
    
        'Chargement des paramètres des ressources.
        Parametre_Lire_Fichier App.Path & FicIni.Chemin & FichierINI_Ressources & FicIni.ExtensionFichierINI
        'FicIni.Section = "Ressources"
        'NombreRessources = FicIni.Parametre("NombreRessources")
        'ReDim Ressources(NombreRessources - 1)
        
        'For i = 0 To UBound(Ressources())
        '    With Ressources(i)
'            FicIni.Section = "Ressource_" & i + 1
'            .Nom = FicIni.Parametre("Nom" & Langues.Suffixe)
'            .NomCompetence = FicIni.Parametre("NomCompetence" & Langues.Suffixe)
'            .VitesseExtraction = Val(FicIni.Parametre("VitesseExtraction"))
'            .TypeMatierePremiere = Val(FicIni.Parametre("TypeMatierePremiere")) - 1
'            .QuantiteMatierePremiere = Val(FicIni.Parametre("QuantiteMatierePremiere"))
'            .PrixVente = Val(FicIni.Parametre("PrixVente"))
'            .Mangeable = Val(FicIni.Parametre("Mangeable"))
'            .GainVie = Val(FicIni.Parametre("GainVie"))
'            .GainEnergie = Val(FicIni.Parametre("GainEnergie"))
'            .GainMagie = Val(FicIni.Parametre("GainMagie"))
'            .GainMoral = Val(FicIni.Parametre("GainMoral"))
'            .NecessairePourConstruire = Val(FicIni.Parametre("NecessairePourConstruire"))
'            .Epoque = Val(FicIni.Parametre("Epoque")) - 1
        '    End With
        'Next i
        
        'FicIni.Section = "Terrain_Ressources"
        'Chargement des ressources disponibles sur le terrain.
        'NombreTerrainRessource = FicIni.Parametre("NombreTerrainRessource")
        'ReDim TerrainRessources(NombreTerrainRessource - 1)
        
        'For i = 0 To UBound(TerrainRessources())
        '    FicIni.Section = "TerrainRessource_" & i + 1
        '    TerrainRessources(i).Nom = FicIni.Parametre("Nom" & Langues.Suffixe)
        '    TerrainRessources(i).TypeRessource = FicIni.Parametre("TypeRessource") - 1
        '    TerrainRessources(i).Quantite = Val(FicIni.Parametre("Quantite"))
        'Next i
        
    Case 3:
        'FicIni.Fichier = FichierINI_Batiments
    
        'Chargement des paramètres des batiments.
        Parametre_Lire_Fichier App.Path & FicIni.Chemin & FichierINI_Batiments & FicIni.ExtensionFichierINI
    
    Case 4:
        FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI_Tresors

        'Chargement des paramètres des trésors.
        FicIni.Section = "Tresors"
        NombreTresors = Val(FicIni.Parametre("NombreTresors"))
        ReDim Tresors(NombreTresors - 1)
        For i = 0 To UBound(Tresors())
            FicIni.Section = "Tresor_" & i + 1
            Tresors(i).Nom = FicIni.Parametre("Nom")
            Tresors(i).ValeurMin = FicIni.Parametre("ValeurMin")
        Next i
        
        FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI_Services
        
        'Chargement des paramètres des services.
        FicIni.Section = "Services"
        NombreServices = FicIni.Parametre("NombreServices")
        ReDim Services(NombreServices - 1)
        For i = 0 To UBound(Services())
            FicIni.Section = "Service_" & i + 1
            Services(i).Nom = FicIni.Parametre("Nom")
            Services(i).Description = FicIni.Parametre("Description")
            Services(i).VitesseCreation = Val(FicIni.Parametre("VitesseCreation"))
            Services(i).TypeMatierePremiere = Val(FicIni.Parametre("TypeMatierePremiere"))
            Services(i).QuantiteMatierePremiere = Val(FicIni.Parametre("QuantiteMatierePremiere"))
            Services(i).Prix = Val(FicIni.Parametre("Prix"))
            Services(i).Coefficient = Val(FicIni.Parametre("Coefficient"))
            Services(i).NomCompetence = FicIni.Parametre("NomCompetence")
            Services(i).Effet = Val(FicIni.Parametre("Effet"))
            Services(i).CoefExp = Val(FicIni.Parametre("CoefExp"))
        Next i
        
        'Chargement des paramètres des compétences d'objets.
        Parametre_Lire_Fichier App.Path & FicIni.Chemin & FichierINI_CompetencesObjets & FicIni.ExtensionFichierINI
'        FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI_Competences_Objets
'        FicIni.Section = "CompetencesObjets"
'        NombreCompetencesObjets = FicIni.Parametre("NombreCompetencesObjets")
'        ReDim CompetencesObjets(NombreCompetencesObjets - 1)
'        For i = 0 To UBound(CompetencesObjets())
'            FicIni.Section = "CompetenceObjet_" & i + 1
'            CompetencesObjets(i).Numero = FicIni.Parametre("Numero")
'            CompetencesObjets(i).NoListe = FicIni.Parametre("NoListe") - 1
'            CompetencesObjets(i).Nom = FicIni.Parametre("Nom")
'        Next i
        
        'Chargement des paramètres des résistances.
        Parametre_Lire_Fichier App.Path & FicIni.Chemin & FichierINI_Resistances & FicIni.ExtensionFichierINI
        
        'Chargement des paramètres des objets.
        Parametre_Lire_Fichier App.Path & FicIni.Chemin & FichierINI_Objets & FicIni.ExtensionFichierINI
        
        'Chargement des décors.
        Parametre_Lire_Fichier App.Path & FicIni.Chemin & FichierINI_Decors & FicIni.ExtensionFichierINI
        'FicIni.Fichier = FichierINI_Decors
        'FicIni.Section = "Decors"
        'NombreDecors = FicIni.Parametre("NombreDecors")
        'ReDim Decors(NombreDecors - 1)
        'For i = 0 To UBound(Decors())
        '    FicIni.Section = "Decor_" & i + 1
        '    Decors(i).Nom = FicIni.Parametre("Nom")
        '    Decors(i).Image = FicIni.Parametre("Image")
        '    Decors(i).Largeur = FicIni.Parametre("Largeur")
        '    Decors(i).Hauteur = FicIni.Parametre("Hauteur")
        '    Decors(i).NombreEtapeAnimation = Val(FicIni.Parametre("NombreEtapeAnimation"))
        'Next i
        
    Case 5:
        'FicIni.Fichier = FichierINI_Peuples
        
        'Chargement des paramètres des peuples.
        Parametre_Lire_Fichier App.Path & FicIni.Chemin & FichierINI_Peuples & FicIni.ExtensionFichierINI
'        FicIni.Section = "Peuples"
'        NombrePeuples = FicIni.Parametre("NombrePeuples")
'        ReDim Peuples(NombrePeuples - 1)
'
        For i = 0 To UBound(Peuples())
'            FicIni.Section = "Peuple_" & i + 1
            With Peuples(i)
            If .CheminImage = "" Then .CheminImage = PeuplesCheminImageDefaut
'            .Nom = FicIni.Parametre("Nom")
'            .NombreMonstres = Val(FicIni.Parametre("NombreMonstres"))
'            If .NombreMonstres > 0 Then
'                ReDim .MonstreIndice(.NombreMonstres - 1)
'                ReDim .MonstreChance(.NombreMonstres - 1)
'                ReDim .MonstreFichierIA(.NombreMonstres - 1)
'                For j = 0 To .NombreMonstres - 1
'                    .MonstreIndice(j) = Val(FicIni.Parametre("Monstre(" & j + 1 & ")Indice")) - 1
'                    .MonstreChance(j) = Val(FicIni.Parametre("Monstre(" & j + 1 & ")Chance"))
'                    .MonstreFichierIA(j) = FicIni.Parametre("Monstre(" & j + 1 & ")FichierIA")
'                Next j
'            End If
'            .NomChateau = FicIni.Parametre("NomChateau")
'            .ChateauMaxVie = Val(FicIni.Parametre("ChateauMaxVie"))
'            .ChateauLargeur = Val(FicIni.Parametre("ChateauLargeur"))
'            .ChateauHauteur = Val(FicIni.Parametre("ChateauHauteur"))
'            .ChateauEntreePositionX = Val(FicIni.Parametre("ChateauEntreePositionX"))
'            .ChateauEntreePositionY = Val(FicIni.Parametre("ChateauEntreePositionY"))
            End With
        Next i
        
    Case 6:
        'FicIni.Fichier = FichierINI_Effets
        
        'Chargement des paramètres visuels.
        Parametre_Lire_Fichier App.Path & FicIni.Chemin & FichierINI_Effets & FicIni.ExtensionFichierINI
        'FicIni.Section = "Effets"
        'NombreEffets = FicIni.Parametre("NombreEffets")
        'ReDim Effets(NombreEffets - 1)
        'EffetMaisonEnFeu = Val(FicIni.Parametre("EffetMaisonEnFeu")) - 1
        'EffetMaisonDetruite = Val(FicIni.Parametre("EffetMaisonDetruite")) - 1
        'EffetChateauDetruit = Val(FicIni.Parametre("EffetChateauDetruit")) - 1
        'EffetNouvelObjectif = Val(FicIni.Parametre("EffetNouvelObjectif")) - 1
        
        For i = 0 To UBound(Effets())
        '    FicIni.Section = "Effet_" & i + 1
        '    Effets(i).Nom = FicIni.Parametre("Nom" & Langues.Suffixe)
        '    Effets(i).Fichier = FicIni.Parametre("Fichier")
        '    Effets(i).Largeur = Val(FicIni.Parametre("Largeur"))
        '    Effets(i).Hauteur = Val(FicIni.Parametre("Hauteur"))
        '    Effets(i).NombreEtapes = Val(FicIni.Parametre("NombreEtapes"))
        '    Effets(i).DureeEtape = Val(FicIni.Parametre("DureeEtape"))
            If Effets(i).DureeEtape = 0 Then Effets(i).DureeEtape = 1
        '    Effets(i).Son = FicIni.Parametre("Son")
        Next i
    
    Case 7:
        'Chargement des informations des races.
        Parametre_Lire_Fichier App.Path & FicIni.Chemin & FichierINI_Races & FicIni.ExtensionFichierINI
        
    Case 8:
        'Verification_Donnees
        'Charge les fichiers de mod officiels.
        Parametre_Lire_Mods App.Path & FicIni.CheminModsOfficiels
        'Charge les fichiers de mod propres à la langue choisie.
        Parametre_Lire_Mods App.Path & FicIni.CheminModsOfficiels & Langues.Dossier
        
        'Charge les fichiers de mod joueurs.
        Parametre_Lire_Mods App.Path & FicIni.CheminMods, True, True
        'Charge les fichiers de mod propres à la langue choisie.
        'Parametre_Lire_Mods App.Path & FicIni.CheminMods & Langues.Dossier
        Verification_Donnees
    End Select
    Exit Sub
Erreur:
    MsgBox "Une erreur s'est produit pendant le chargement de :" & FicIni.DernierChargement & "(" & FicIni.Fichier & ").", vbCritical, "Erreur"
    End
End Sub

Private Sub Bestiaire_Tri_Alphabetique()
    Dim i As Integer, Temp1 As Integer, Temp2 As Integer
    If NombreRaces > 1 Then
        'Trie le bestiaire par ordre alhpabétique.
        'Système de tri par bulle.
        ReDim BestiaireOrdreAlpha(NombreRaces - 1)
        For i = 0 To NombreRaces - 1
            BestiaireOrdreAlpha(i) = i
        Next i
        Do
            Temp1 = False
            For i = 1 To NombreRaces - 1
                If Races(BestiaireOrdreAlpha(i - 1)).NomMPluriel > Races(BestiaireOrdreAlpha(i)).NomMPluriel Then
                    Temp2 = BestiaireOrdreAlpha(i - 1)
                    BestiaireOrdreAlpha(i - 1) = BestiaireOrdreAlpha(i)
                    BestiaireOrdreAlpha(i) = Temp2
                    Temp1 = True
                End If
            Next i
        Loop Until Not Temp1
    End If
End Sub

Private Sub Parametre_Lire_Mods(ByVal Chemin As String, Optional ByVal SousDossiers As Boolean, Optional ByVal FichierMod As Boolean)
    'Charge tous les fichiers de mods situés dans les dossiers de mod.
    Dim i As Long
    Dim Temp As String
    Dim Dossiers As Collection
    
    Temp = Dir(Chemin, vbNormal)
    While Temp <> ""
        If Right(Temp, Len(FicIni.ExtensionFichierINI)) = FicIni.ExtensionFichierINI Then
            Parametre_Lire_Fichier Chemin & Temp, FichierMod
        End If
        Temp = Dir()
    Wend
    
    'Vérifie les sous répertoires.
    If SousDossiers Then
        Set Dossiers = New Collection
        Temp = Dir(Chemin, vbDirectory)
        While Temp <> ""
            If InStr(Temp, ".") < 1 Then
                Dossiers.Add Temp
            End If
            Temp = Dir()
        Wend
        If Dossiers.Count > 0 Then
            For i = 1 To Dossiers.Count
                Parametre_Lire_Mods Chemin & Dossiers(i) & "\", , FichierMod
            Next i
        End If
    End If
End Sub

Private Sub Parametre_Lire_Fichier(ByVal Fichier As String, Optional ByVal FichierMod As Boolean)
    'On Error GoTo Erreur
    Dim Indice As Long
    Dim Temp As String, Section As String
    Open Fichier For Input As #1
    Do
        Line Input #1, Temp
        Parametre_Analyser Temp, Section, Indice, Fichier, FichierMod
    Loop Until EOF(1)
    Close #1
    Exit Sub
Erreur:
    'Message d'erreur précisant son emplacment.
    MsgBox MessageErreurFichierManquant & Chr(10) & Chr(13) & Fichier, vbCritical, TitreErreur
    End
End Sub

Private Sub Parametre_Analyser(ByVal Ligne As String, ByRef Section As String, ByRef i As Long, ByVal Fichier As String, Optional ByVal FichierMod As Boolean)
    'On Error GoTo Erreur
    Dim Valeur As String
    Dim j As Long
    Dim Temp As String
    Temp = Parametre_Filtre(Ligne)
    If Temp <> "" And Left(Temp, 1) <> "#" And Left(Temp, 1) <> ";" Then
        If InStr(Temp, "=") = 0 Then
            If Left(Temp, 11) = "[Ressource_" Then
                i = Val(Mid(Temp, 12, Len(Temp) - 12)) - 1
                If i >= NombreRessources Then
                    NombreRessources = NombreRessources + 1
                    ReDim Preserve Ressources(NombreRessources - 1)
                    Ressources(NombreRessources - 1).FichierMod = FichierMod
                End If
            ElseIf Left(Temp, 18) = "[TerrainRessource_" Then
                i = Val(Mid(Temp, 19, Len(Temp) - 19)) - 1
                If i >= NombreTerrainRessource Then
                    NombreTerrainRessource = NombreTerrainRessource + 1
                    ReDim Preserve TerrainRessources(NombreTerrainRessource - 1)
                End If
            ElseIf Left(Temp, 10) = "[Batiment_" Then
                i = Gestion_Cle_Batiments(Mid(Temp, 11, Len(Temp) - 11), FichierMod)
                Section = SectionINI_Batiments
            ElseIf Left(Temp, 17) = "[CompetenceObjet_" Then
                i = Gestion_Cle_CompetencesObjets(Mid(Temp, 18, Len(Temp) - 18), FichierMod)
                Section = SectionINI_CompetencesObjets
            ElseIf Left(Temp, 7) = "[Objet_" Then
                If Mid(Temp, 8, Len(Temp) - 8) = "+" Then
                    i = NombreObjets
                Else
                    i = Val(Mid(Temp, 8, Len(Temp) - 8)) - 1
                End If
                If i >= NombreObjets Then
                    NombreObjets = i + 1
                    ReDim Preserve Objets(NombreObjets - 1)
                    'For j = 0 To NombreObjets - 1
                    ReDim Preserve Objets(NombreObjets - 1).PrixRessources(NombreRessources - 1)
                    ReDim Preserve Objets(NombreObjets - 1).BonusCompetencesSpeciales(NombreCompetencesSpeciales - 1)
                    ReDim Preserve Objets(NombreObjets - 1).BonusCompetencesRessources(NombreRessources - 1)
                    ReDim Preserve Objets(NombreObjets - 1).BonusCompetencesServices(NombreServices - 2)
                    ReDim Preserve Objets(NombreObjets - 1).BonusCompetencesObjets(NombreCompetencesObjets - 1)
                    ReDim Preserve Objets(NombreObjets - 1).BonusResistances(NombreResistances - 1)
                    ReDim Preserve Objets(NombreObjets - 1).GainRessource(NombreRessources - 1)
                    ReDim Preserve Objets(NombreObjets - 1).ApparenceDecalageX(4)  'Quatre directions différentes.
                    ReDim Preserve Objets(NombreObjets - 1).ApparenceDecalageY(4)  'Quatre directions différentes.
                    ReDim Preserve Objets(NombreObjets - 1).ApparenceAvantPlan(4) 'Quatre directions différentes.
                    'Next j
                    'Objets(i).PerteRessourceIndice = -1
                    Objets(NombreObjets - 1).EffetAttaque = -1
                    Objets(NombreObjets - 1).CoefXPFabrication = 1
                    Objets(NombreObjets - 1).Invocation = -1
                    Objets(NombreObjets - 1).CreerObjet = -1
                    Objets(NombreObjets - 1).CreerObjetAllie = -1
                    Objets(NombreObjets - 1).CreerObjetEnnemi = -1
                    Objets(NombreObjets - 1).AttaqueCauseEffet = -1
                    Objets(NombreObjets - 1).AttaqueResistance = -1
                    Objets(NombreObjets - 1).FichierMod = FichierMod
                End If
                Section = SectionINI_Objets
            ElseIf Left(Temp, 8) = "[Peuple_" Then
                If Mid(Temp, 9, Len(Temp) - 9) = "+" Then
                    i = NombrePeuples
                Else
                    i = Val(Mid(Temp, 9, Len(Temp) - 9)) - 1
                End If
                If i >= NombrePeuples Then
                    NombrePeuples = NombrePeuples + 1
                    ReDim Preserve Peuples(NombrePeuples - 1)
                    Peuples(NombrePeuples - 1).FichierMod = FichierMod
                End If
                Section = SectionINI_Peuples
            ElseIf Left(Temp, 7) = "[Effet_" Then
                If Mid(Temp, 8, Len(Temp) - 8) = "+" Then
                    i = NombreEffets
                Else
                    i = Val(Mid(Temp, 8, Len(Temp) - 8)) - 1
                End If
                If i >= NombreEffets Then
                    NombreEffets = NombreEffets + 1
                    ReDim Preserve Effets(NombreEffets - 1)
                End If
                Section = SectionINI_Effets
            ElseIf Left(Temp, 6) = "[Race_" Then
                If Mid(Temp, 7, Len(Temp) - 7) = "+" Then
                    i = NombreRaces
                Else
                    i = Val(Mid(Temp, 7, Len(Temp) - 7)) - 1
                End If
                If i >= NombreRaces Then
                    NombreRaces = i + 1
                    ReDim Preserve Races(NombreRaces - 1)
                    ReDim Preserve Races(NombreRaces - 1).DossiersCategoriesF(0)
                    ReDim Preserve Races(NombreRaces - 1).DossiersCategoriesM(0)
                    Races(NombreRaces - 1).Vitesse = -1
                    ReDim Preserve Races(NombreRaces - 1).CoefCompSpeciales(NombreCompetencesSpeciales - 1)
                    ReDim Preserve Races(NombreRaces - 1).CoefCompRessources(NombreRessources - 1) 'Compétences des ressources.
                    ReDim Preserve Races(NombreRaces - 1).CoefCompServices(NombreServices - 1) 'Compétences des Services.
                    ReDim Preserve Races(NombreRaces - 1).CoefCompObjets(NombreCompetencesObjets - 1)  'Compétences des objets.
                    ReDim Preserve Races(NombreRaces - 1).Resistances(NombreResistances - 1)
                    For j = 0 To NombreResistances - 1
                        Races(NombreRaces - 1).Resistances(j) = ResistanceDefaut 'Charge la valeur pas défaut.
                    Next j
                    ReDim Preserve Races(NombreRaces - 1).Ressources(NombreRessources - 1)
                    Set Races(NombreRaces - 1).Statut = New Collection
                    Races(NombreRaces - 1).Tombe = -1
                    Races(NombreRaces - 1).AttaqueCauseEffet = -1
                    Set Races(NombreRaces - 1).ObjetRace = New Collection
                    Races(NombreRaces - 1).EffetMort = -1
                    Races(NombreRaces - 1).FichierMod = FichierMod
                End If
                Section = SectionINI_Races
            ElseIf Left(Temp, 7) = "[Decor_" Then
                If Mid(Temp, 8, Len(Temp) - 8) = "+" Then
                    i = NombreDecors
                Else
                    i = Val(Mid(Temp, 8, Len(Temp) - 8)) - 1
                End If
                If i >= NombreDecors Then
                    NombreDecors = NombreDecors + 1
                    ReDim Preserve Decors(NombreDecors - 1)
                End If
            ElseIf Left(Temp, 12) = "[Resistance_" Then
                If Mid(Temp, 13, Len(Temp) - 13) = "+" Then
                    i = NombreResistances
                Else
                    i = Val(Mid(Temp, 13, Len(Temp) - 13)) - 1
                End If
                If i >= NombreResistances Then
                    NombreResistances = i + 1
                    ReDim Preserve Resistances(NombreResistances - 1)
                    For j = 0 To NombreRaces - 1
                        ReDim Preserve Races(NombreRaces - 1).Resistances(NombreResistances - 1)
                    Next j
                    For j = 0 To NombreObjets - 1
                        ReDim Preserve Objets(NombreObjets - 1).BonusResistances(NombreResistances - 1)
                    Next j
                End If
                Section = SectionINI_Resistances
            ElseIf Left(Temp, 8) = "[Statut_" Then
                If Mid(Temp, 9, Len(Temp) - 9) = "+" Then
                    i = NombreStatuts
                Else
                    i = Val(Mid(Temp, 9, Len(Temp) - 9)) - 1
                End If
                If i >= NombreStatuts Then
                    NombreStatuts = i + 1
                    ReDim Preserve Statuts(NombreStatuts - 1)
                    ReDim Preserve Statuts(NombreStatuts - 1).Resistances(NombreResistances - 1)
                    For j = 0 To NombreResistances - 1
                        Statuts(NombreStatuts - 1).Resistances(j) = ResistanceDefaut 'Charge la valeur pas défaut.
                    Next j
                End If
                Section = SectionINI_Statuts
            Else
                Select Case Temp
                Case "[Ressources]"
                    Section = "Ressources"
                Case "[Terrain_Ressources]"
                    Section = "Terrain_Ressources"
                Case "[" & SectionINI_Batiments & "]"
                    Section = SectionINI_Batiments
                Case "[CompetencesObjets]":
                    Section = SectionINI_CompetencesObjets
                Case "[Objets]":
                    Section = SectionINI_Objets
                Case "[" & SectionINI_Peuples & "]"
                    Section = SectionINI_Peuples
                Case "[" & SectionINI_Effets & "]"
                    Section = SectionINI_Effets
                Case "[Races]"
                    Section = SectionINI_Races
                    i = -1
                Case "[" & SectionINI_Decors & "]"
                    Section = SectionINI_Decors
                Case "[" & SectionINI_Resistances & "]"
                    Section = SectionINI_Resistances
                Case "[" & SectionINI_Statuts & "]"
                    Section = SectionINI_Statuts
                End Select
            End If
        Else
            Valeur = Parametre_Droite(Temp)
            Temp = Parametre_Gauche(Temp)
            j = Parametre_Indice(Temp)
            Temp = Parametre_Gauche2(Temp)
            Select Case Section
            Case "Ressources":
                With Ressources(i)
                Select Case Temp
                    Case "Nom": .Nom = Valeur
                    Case "NomCompetence": .NomCompetence = Valeur
                    Case "Image": .Image = Valeur
                    Case "CheminImage": .CheminImage = Valeur
                    Case "VitesseExtraction": .VitesseExtraction = Val(Valeur)
                    Case "TypeMatierePremiere": .TypeMatierePremiere = Val(Valeur) - 1
                    Case "QuantiteMatierePremiere": .QuantiteMatierePremiere = Val(Valeur)
                    Case "PrixVente": .PrixVente = Val(Valeur)
                    Case "Mangeable": .Mangeable = Val(Valeur)
                    Case "GainVie": .GainVie = Val(Valeur)
                    Case "GainEnergie": .GainEnergie = Val(Valeur)
                    Case "GainMagie": .GainMagie = Val(Valeur)
                    Case "GainMoral": .GainMoral = Val(Valeur)
                    Case "NecessairePourConstruire": .NecessairePourConstruire = Val(Valeur)
                    Case "Epoque": .Epoque = Val(Valeur) - 1
                End Select
                End With
            Case "Terrain_Ressources"
                With TerrainRessources(i)
                Select Case Temp
                    Case "Nom": .Nom = Valeur
                    Case "TypeRessource": .TypeRessource = Val(Valeur) - 1
                    Case "Quantite": .Quantite = Val(Valeur)
                End Select
                End With
            Case SectionINI_Batiments:
                'Chargement des paramètres des bâtiments.
                Select Case Temp
                    Case "ID": i = Gestion_Cle_Batiments(Valeur, FichierMod)
                    Case "Nom": Batiments(i).Nom = Valeur
                    Case "Nom" & Langues.Suffixe: Batiments(i).Nom = Valeur
                    Case "Feminin" & Langues.Suffixe: Batiments(i).Feminin = Val(Valeur)
                    Case "Image": Batiments(i).Image = Valeur
                    Case "CheminImage":
                        Batiments(i).CheminImage = Valeur
                    Case "CheminImageDefaut":
                        BatimentsCheminImageDefaut = Valeur
                    Case "ImageBouton": Batiments(i).ImageBouton = Valeur
                    Case "CheminImageBouton": Batiments(i).CheminImageBouton = Right(Valeur, Len(Valeur) - 1)
                    Case "Metier": Batiments(i).MetierM = Valeur
                                   Batiments(i).MetierF = Valeur
                    Case "MetierM": Batiments(i).MetierM = Valeur
                    Case "MetierF": Batiments(i).MetierF = Valeur
                    Case "Metier" & Langues.Suffixe: Batiments(i).MetierM = Valeur
                                   Batiments(i).MetierF = Valeur
                    Case "MetierM" & Langues.Suffixe: Batiments(i).MetierM = Valeur
                    Case "MetierF" & Langues.Suffixe: Batiments(i).MetierF = Valeur
                    Case "Description": Batiments(i).Description = Valeur
                    Case "Description" & Langues.Suffixe: Batiments(i).Description = Valeur
                    Case "Empereur": Batiments(i).Empereur = Val(Valeur)
                    
                    Case "Limite": Batiments(i).Limite = Val(Valeur)
                    Case "CreerProprietaire":
                        If Valeur = "+" Or Valeur = "-" Then
                            Valeur = NombreRaces - 1
                        ElseIf Left(Valeur, 1) = "+" Then
                            Valeur = Right(Valeur, Len(Valeur) - 1) + NombreRaces - 1
                        ElseIf Left(Valeur, 1) = "-" Then
                            Valeur = Right(Valeur, Len(Valeur) - 1) - NombreRaces - 1
                        Else
                            Valeur = Val(Valeur) - 1
                        End If
                        Batiments(i).CreerProprietaire = Val(Valeur)
                    Case "Roi": Batiments(i).Roi = Val(Valeur)
                    Case "RestrictionRoi": Batiments(i).RestrictionRoi = Val(Valeur)
                    Case "GainObjet": Batiments(i).GainObjet = Val(Valeur) - 1
                    Case "PerteObjet": Batiments(i).PerteObjet = Val(Valeur) - 1
                    Case "Tir": Batiments(i).Tir = Val(Valeur)
                    Case "Marche": Batiments(i).Marche = Val(Valeur)
                    Case "Prix": Batiments(i).Prix = Val(Valeur)
                    Case "PrixArgent": Batiments(i).PrixArgent = Val(Valeur)
                    Case "PrixRessources": Batiments(i).PrixRessources(j - 1) = Val(Valeur)
                    Case "MaxVie": Batiments(i).MaxVie = Val(Valeur)
                    Case "CoefVitesseConstruction": Batiments(i).CoefVitesseConstruction = Val(Valeur)
                    Case "RessourceFabrique": Batiments(i).RessourceFabrique = Val(Valeur)
                    Case "NombreObjetsfabriques": Batiments(i).NombreObjetsFabriques = Val(Valeur)
                        If Batiments(i).NombreObjetsFabriques > 0 Then
                            ReDim Preserve Batiments(i).ObjetsFabriques(Batiments(i).NombreObjetsFabriques - 1)
                        End If
                    Case "ObjetsFabriques": Batiments(i).ObjetsFabriques(j - 1) = Val(Valeur) - 1
                    Case "Magasin": Batiments(i).Magasin = Val(Valeur)
                    Case "Fabrique": Batiments(i).Fabrique = Val(Valeur)
                    Case "Service": Batiments(i).Service = Val(Valeur)
                    Case "MaxStock": Batiments(i).MaxStock = Val(Valeur)
                    Case "Niveau": Batiments(i).Niveau = Val(Valeur)
                    Case "NoEpoque": Batiments(i).NoEpoque = Val(Valeur)
                    Case "Largeur": Batiments(i).Largeur = Val(Valeur)
                    Case "Hauteur": Batiments(i).Hauteur = Val(Valeur)
                    Case "EntreePositionX": Batiments(i).EntreePositionX = Val(Valeur)
                    Case "EntreePositionY": Batiments(i).EntreePositionY = Val(Valeur)
                    Case "NombreEtapeFabrication": Batiments(i).NombreEtapeFabrication = Val(Valeur)
                    Case "NombreEtapeAnimation": Batiments(i).NombreEtapeAnimation = Val(Valeur)
                    Case "ApparencePerso": Batiments(i).ApparencePerso = Val(Valeur)
                End Select
            Case SectionINI_CompetencesObjets:
                Select Case Temp
                    Case "ID": i = Gestion_Cle_CompetencesObjets(Valeur, FichierMod)
                    Case "SAVPerso": CompetencesObjets(i).SAVPerso = Valeur
                    Case "NoListe": CompetencesObjets(i).NoListe = Val(Valeur) - 1
                    Case "Nom": CompetencesObjets(i).Nom = Valeur
                    Case "Nom" & Langues.Suffixe: CompetencesObjets(i).Nom = Valeur
                End Select
            Case SectionINI_Objets:
                'Chargement des paramètres des objets.
                Select Case Temp
                    Case "Nom": Objets(i).Nom = Valeur
                    Case "Nom" & Langues.Suffixe: Objets(i).Nom = Valeur
                    Case "NomEffet": Objets(i).NomEffetM = Valeur
                                     Objets(i).NomEffetF = Valeur
                    Case "NomEffetM": Objets(i).NomEffetM = Valeur
                    Case "NomEffetF": Objets(i).NomEffetF = Valeur
                    Case "NomEffet" & Langues.Suffixe: Objets(i).NomEffetM = Valeur
                                                       Objets(i).NomEffetF = Valeur
                    Case "NomEffetM" & Langues.Suffixe: Objets(i).NomEffetM = Valeur
                    Case "NomEffetF" & Langues.Suffixe: Objets(i).NomEffetF = Valeur
                    Case "Image": Objets(i).Image = Valeur
                    Case "CheminImage":
                        Objets(i).CheminImage = Valeur
                    Case "NomPluriel": Objets(i).NomPluriel = Valeur
                    Case "NomPluriel" & Langues.Suffixe: Objets(i).NomPluriel = Valeur
                    Case "Description": Objets(i).Description = Valeur
                    Case "Description" & Langues.Suffixe: Objets(i).Description = Valeur
                    Case "Secret": Objets(i).Secret = Val(Valeur)
                    Case "Unique": Objets(i).Unique = Val(Valeur)
                    Case "Indestructible": Objets(i).Indestructible = Val(Valeur)
                    Case "Inseparable": Objets(i).Inseparable = Val(Valeur)
                    Case "Superposable": Objets(i).Superposable = Val(Valeur)
                    Case "Empereur": Objets(i).Empereur = Val(Valeur)
                    Case "Categorie": Objets(i).Categorie = Valeur
                    Case "Competence":
                        'L'objet est lié à la clé de la compétence en paramètre.
                        'Valeur = Gestion_Cle_CompetencesObjets(Valeur, FichierMod)
                        Objets(i).Competence = Gestion_Cle_CompetencesObjets(Valeur, FichierMod)
                    Case "PrixRessources": Objets(i).PrixRessources(j - 1) = Val(Valeur)
                    Case "Epoque": Objets(i).Epoque = Val(Valeur)
                    Case "DureeFabrication": Objets(i).DureeFabrication = Val(Valeur)
                    Case "CoefXPFabrication": Objets(i).CoefXPFabrication = Val(Valeur)
                    Case "FabricationBatiment":
                        'Rajoute l'objet dans le bâtiment donné en paramètre.
                        Valeur = Gestion_Cle_Batiments(Valeur, FichierMod)
                        Batiments(Valeur).NombreObjetsFabriques = Batiments(Valeur).NombreObjetsFabriques + 1
                        ReDim Preserve Batiments(Valeur).ObjetsFabriques(Batiments(Valeur).NombreObjetsFabriques - 1)
                        Batiments(Valeur).ObjetsFabriques(Batiments(Valeur).NombreObjetsFabriques - 1) = i
                    Case "EquipementRace":
                        'Rajoute l'objet à l'équipement de la race concernée.
                        If Valeur = "+" Or Valeur = "-" Then
                            Valeur = NombreRaces - 1
                        ElseIf Left(Valeur, 1) = "+" Then
                            Valeur = Right(Valeur, Len(Valeur) - 1) + NombreRaces - 1
                        ElseIf Left(Valeur, 1) = "-" Then
                            Valeur = Right(Valeur, Len(Valeur) - 1) - NombreRaces - 1
                        Else
                            Valeur = Val(Valeur) - 1
                        End If
                        Races(Valeur).NombreEquipement = Races(Valeur).NombreEquipement + 1
                        ReDim Preserve Races(Valeur).Equipement(Races(Valeur).NombreEquipement - 1)
                        Races(Valeur).Equipement(Races(Valeur).NombreEquipement - 1) = i
                    Case "TresorRace":
                        'Rajoute l'objet à la race concernée.
                        If Valeur = "+" Or Valeur = "-" Then
                            Valeur = NombreRaces - 1
                        ElseIf Left(Valeur, 1) = "+" Then
                            Valeur = Right(Valeur, Len(Valeur) - 1) + NombreRaces - 1
                        ElseIf Left(Valeur, 1) = "-" Then
                            Valeur = Right(Valeur, Len(Valeur) - 1) - NombreRaces - 1
                        Else
                            Valeur = Val(Valeur) - 1
                        End If
                        If Valeur = -1 Then 'Si 0, on donne l'objet à toutes les races.
                            NombreObjetsCommmuns = NombreObjetsCommmuns + 1
                            ReDim Preserve ObjetCommmuns(NombreObjetsCommmuns - 1)
                            ObjetCommmuns(NombreObjetsCommmuns - 1) = i
                        Else
                            Races(Valeur).NombreObjets = Races(Valeur).NombreObjets + 1
                            ReDim Preserve Races(Valeur).Objet(Races(Valeur).NombreObjets - 1)
                            Races(Valeur).Objet(Races(Valeur).NombreObjets - 1) = i
                        End If
                    Case "PrixVente": Objets(i).PrixVente = Val(Valeur)
                    Case "BonusVie": Objets(i).BonusVie = Val(Valeur)
                    Case "BonusEnergie": Objets(i).BonusEnergie = Val(Valeur)
                    Case "BonusMagie": Objets(i).BonusMagie = Val(Valeur)
                    Case "BonusMoral": Objets(i).BonusMoral = Val(Valeur)
                    Case "BonusAttaque": Objets(i).BonusAttaque = Val(Valeur)
                    Case "BonusPorteeAttaque": Objets(i).BonusPorteeAttaque = Val(Valeur)
                    Case "LanceProjectile":
                        If Jeu.Definir_AfficherProjectiles Then
                            If Valeur = "+" Then
                                Objets(i).LanceProjectile = NombreRaces - 1
                            Else
                                Objets(i).LanceProjectile = Val(Valeur) - 1
                            End If
                        End If
                    Case "EffetAttaque":
                        If Valeur = "+" Then
                            Objets(i).EffetAttaque = NombreEffets - 1
                        Else
                            Objets(i).EffetAttaque = Val(Valeur) - 1
                        End If
                    Case "ProjectileSupplementaire":
                        Objets(i).ProjectileSupplementaire = Val(Valeur)
                    Case "AttaqueResistance":
                        Objets(i).AttaqueResistance = Val(Valeur) - 1
                    Case "BonusResistance": Objets(i).BonusResistances(j - 1) = Val(Valeur)
                    Case "SurfaceAttaque": Objets(i).SurfaceAttaque = Val(Valeur)
                    Case "AttaqueAutomatique": Objets(i).AttaqueAutomatique = Val(Valeur)
                    Case "AttaqueCauseEffet":
                        If Valeur = "+" Or Valeur = "-" Then
                            Objets(i).AttaqueCauseEffet = NombreRaces - 1
                        ElseIf Left(Valeur, 1) = "+" Then
                            Objets(i).AttaqueCauseEffet = Right(Valeur, Len(Valeur) - 1) + NombreObjets - 1
                        ElseIf Left(Valeur, 1) = "-" Then
                            Objets(i).AttaqueCauseEffet = Right(Valeur, Len(Valeur) - 1) - NombreObjets - 1
                        Else
                            Objets(i).AttaqueCauseEffet = Val(Valeur) - 1
                        End If
                    Case "BonusDefense": Objets(i).BonusDefense = Val(Valeur)
                    Case "BonusDeplacement": Objets(i).BonusDeplacement = Val(Valeur)
                    Case "BonusVitesse":
                        Objets(i).BonusVitesse = Val(Valeur)
                    Case "BonusVitalite":
                        Objets(i).BonusVitalite = Val(Valeur)
                    Case "BonusCompetencesSpeciales": Objets(i).BonusCompetencesSpeciales(j - 1) = Val(Valeur)
                    Case "BonusCompetencesRessources": Objets(i).BonusCompetencesRessources(j - 1) = Val(Valeur)
                    Case "BonusCompetencesServices": Objets(i).BonusCompetencesServices(j - 2) = Val(Valeur)
                    Case "BonusCompetencesObjets": Objets(i).BonusCompetencesObjets(j - 1) = Val(Valeur)
                    Case "BonusArmure": Objets(i).BonusArmure = Val(Valeur)
                    Case "BonusMaxRessources": Objets(i).BonusMaxRessources = Val(Valeur)
                    Case "BonusVitesseRegenerationVie": Objets(i).BonusVitesseRegenerationVie = Val(Valeur)
                    Case "BonusVitesseRegenerationEnergie": Objets(i).BonusVitesseRegenerationEnergie = Val(Valeur)
                    Case "BonusVitesseRegenerationMagie": Objets(i).BonusVitesseRegenerationMagie = Val(Valeur)
                    Case "BonusVitesseRegenerationMoral": Objets(i).BonusVitesseRegenerationMoral = Val(Valeur)
                    Case "PerteVie": Objets(i).PerteVie = Val(Valeur)
                    Case "PerteEnergie": Objets(i).PerteEnergie = Val(Valeur)
                    Case "PerteMagie": Objets(i).PerteMagie = Val(Valeur)
                    Case "PerteMoral": Objets(i).PerteMoral = Val(Valeur)
                    Case "PerteArgent": Objets(i).PerteArgent = Val(Valeur)
                    Case "PerteRessourceIndice": Objets(i).PerteRessourceIndice = Val(Valeur) - 1
                    Case "PerteRessourceQuantite": Objets(i).PerteRessourceQuantite = Val(Valeur)
                    Case "Poison": Objets(i).Poison = Val(Valeur)
                    Case "ApparencePerso": Objets(i).ApparencePerso = Val(Valeur)
                    Case "CheminApparence":
                        AffApparence.Ajouter_Dossier Valeur
                    Case "FichierApparence":
                        Objets(i).FichierApparenceM = Valeur
                        Objets(i).FichierApparenceF = Valeur
                    Case "FichierApparenceM":
                        Objets(i).FichierApparenceM = Valeur
                    Case "FichierApparenceF":
                        Objets(i).FichierApparenceF = Valeur
                    Case "NumeroApparenceM": Objets(i).NumeroApparenceM = Val(Valeur)
                    Case "NumeroApparenceF": Objets(i).NumeroApparenceF = Val(Valeur)
                    Case "ApparenceLargeur": Objets(i).ApparenceLargeur = Val(Valeur)
                    Case "ApparenceHauteur": Objets(i).ApparenceHauteur = Val(Valeur)
                    Case "ApparenceDecalageXH": Objets(i).ApparenceDecalageX(1) = Val(Valeur)
                    Case "ApparenceDecalageYH": Objets(i).ApparenceDecalageY(1) = Val(Valeur)
                    Case "ApparenceDecalageXD": Objets(i).ApparenceDecalageX(2) = Val(Valeur)
                    Case "ApparenceDecalageYD": Objets(i).ApparenceDecalageY(2) = Val(Valeur)
                    Case "ApparenceDecalageXB": Objets(i).ApparenceDecalageX(3) = Val(Valeur)
                    Case "ApparenceDecalageYB": Objets(i).ApparenceDecalageY(3) = Val(Valeur)
                    Case "ApparenceDecalageXG": Objets(i).ApparenceDecalageX(4) = Val(Valeur)
                    Case "ApparenceDecalageYG": Objets(i).ApparenceDecalageY(4) = Val(Valeur)
                    Case "ApparenceAvantPlanH": Objets(i).ApparenceAvantPlan(1) = Not (CBool(Val(Valeur)))
                    Case "ApparenceAvantPlanD": Objets(i).ApparenceAvantPlan(2) = Not (CBool(Val(Valeur)))
                    Case "ApparenceAvantPlanB": Objets(i).ApparenceAvantPlan(3) = Not (CBool(Val(Valeur)))
                    Case "ApparenceAvantPlanG": Objets(i).ApparenceAvantPlan(4) = Not (CBool(Val(Valeur)))
                    Case "ChangerApparencePerso": Objets(i).ChangerApparencePerso = Val(Valeur)
                    Case "Utilisable": Objets(i).Utilisable = Val(Valeur)
                    Case "UtilisationUnique": Objets(i).UtilisationUnique = Val(Valeur)
                    Case "UtilisationMort": Objets(i).UtilisationMort = Val(Valeur)
                    Case "PorteeUtilisation": Objets(i).PorteeUtilisation = Val(Valeur)
                    Case "Experience": Objets(i).Experience = Val(Valeur)
                    Case "Invocation":
                        If Valeur = "+" Or Valeur = "-" Then
                            Objets(i).Invocation = NombreRaces - 1
                        ElseIf Left(Valeur, 1) = "+" Then
                            Objets(i).Invocation = NombreRaces - 1 + Val(Right(Valeur, Len(Valeur) - 1))
                        ElseIf Left(Valeur, 1) = "-" Then
                            Objets(i).Invocation = NombreRaces - 1 - Val(Right(Valeur, Len(Valeur) - 1))
                        Else
                            Objets(i).Invocation = Val(Valeur) - 1
                        End If
                    Case "InvocationIA_Fichier": Objets(i).InvocationIA_Fichier = Valeur
                    Case "Quantite": Objets(i).Quantite = Val(Valeur)
                    Case "QuantiteMax": Objets(i).QuantiteMax = Val(Valeur)
                    Case "GainVie": Objets(i).GainVie = Val(Valeur)
                    Case "GainEnergie": Objets(i).GainEnergie = Val(Valeur)
                    Case "GainMagie": Objets(i).GainMagie = Val(Valeur)
                    Case "GainMoral": Objets(i).GainMoral = Val(Valeur)
                    Case "GainArgent": Objets(i).GainArgent = Val(Valeur)
                    Case "Impot": Objets(i).Impot = Val(Valeur)
                    Case "GainRessource": Objets(i).GainRessource(j - 1) = Val(Valeur)
                    Case "Resurrection": Objets(i).Resurrection = Val(Valeur)
                    Case "Berserk": Objets(i).Berserk = Val(Valeur)
                    Case "SelfResurrection": Objets(i).SelfResurrection = Val(Valeur)
                    Case "CreerObjet":
                        If Valeur = "+" Or Valeur = "-" Then
                            Objets(i).CreerObjet = NombreObjets - 1
                        ElseIf Left(Valeur, 1) = "+" Then
                            Objets(i).CreerObjet = Right(Valeur, Len(Valeur) - 1) + Val(i)
                        ElseIf Left(Valeur, 1) = "-" Then
                            Objets(i).CreerObjet = Right(Valeur, Len(Valeur) - 1) - Val(i)
                        Else
                            Objets(i).CreerObjet = Val(Valeur) - 1
                        End If
                    Case "CreerObjetAllie":
                        If Valeur = "+" Or Valeur = "-" Then
                            Objets(i).CreerObjetAllie = NombreObjets - 1
                        ElseIf Left(Valeur, 1) = "+" Then
                            Objets(i).CreerObjetAllie = Right(Valeur, Len(Valeur) - 1) + Val(i)
                        ElseIf Left(Valeur, 1) = "-" Then
                            Objets(i).CreerObjetAllie = Right(Valeur, Len(Valeur) - 1) - Val(i)
                        Else
                            Objets(i).CreerObjetAllie = Val(Valeur) - 1
                        End If
                    Case "CreerObjetEnnemi":
                        If Valeur = "+" Or Valeur = "-" Then
                            Objets(i).CreerObjetEnnemi = NombreObjets - 1
                        ElseIf Left(Valeur, 1) = "+" Then
                            Objets(i).CreerObjetEnnemi = Right(Valeur, Len(Valeur) - 1) + Val(i)
                        ElseIf Left(Valeur, 1) = "-" Then
                            Objets(i).CreerObjetEnnemi = Right(Valeur, Len(Valeur) - 1) - Val(i)
                        Else
                            Objets(i).CreerObjetEnnemi = Val(Valeur) - 1
                        End If
                    Case "Duree": Objets(i).Duree = Val(Valeur)
                    Case "DureeRecharge":
                        If Jeu.Regle_TempsRechargement_Objets Then
                            Objets(i).DureeRecharge = Val(Valeur)
                        End If
                    Case "DureeRechargeAttaque": Objets(i).DureeRechargeAttaque = Val(Valeur)
                    Case "RechargesMultiples": Objets(i).RechargesMultiples = Val(Valeur)
                    Case "EffetAnnuleRestaureObjet": Objets(i).EffetAnnuleRestaureObjet = Val(Valeur)
                    Case "EffetAnnulationImpossible": Objets(i).EffetAnnulationImpossible = Val(Valeur)
                    Case "EffetAnnulationMort": Objets(i).EffetAnnulationMort = Val(Valeur)
                    Case "UtilisationAutomatique": Objets(i).UtilisationAutomatique = Val(Valeur)
                    Case "Monture": Objets(i).Monture = Val(Valeur)
                    Case "Volant": Objets(i).Volant = Val(Valeur)
                    Case "PrixTalent": Objets(i).PrixTalents = Val(Valeur)
                    Case "PrixTalents": Objets(i).PrixTalents = Val(Valeur)
                End Select
            
            Case SectionINI_Peuples:
                Select Case Temp
                    Case "Nom": Peuples(i).Nom = Valeur
                    Case "Nom" & Langues.Suffixe: Peuples(i).Nom = Valeur
                    Case "Image": Peuples(i).Image = Valeur
                    Case "CheminImage": Peuples(i).CheminImage = Valeur
                    Case "CheminImageDefaut": PeuplesCheminImageDefaut = Valeur
                    Case "NombreMonstres": Peuples(i).NombreMonstres = Val(Valeur)
                        ReDim Preserve Peuples(i).MonstreIndice(Peuples(i).NombreMonstres - 1)
                        ReDim Preserve Peuples(i).MonstreChance(Peuples(i).NombreMonstres - 1)
                        ReDim Preserve Peuples(i).MonstreFichierIA(Peuples(i).NombreMonstres - 1)
                    Case "MonstreIndice":
                        If Valeur = "+" Or Valeur = "-" Then
                            Peuples(i).MonstreIndice(j - 1) = NombreRaces - 1
                        ElseIf Left(Valeur, 1) = "+" Then
                            Peuples(i).MonstreIndice(j - 1) = NombreRaces + Right(Valeur, Len(Valeur) - 1) - 1
                        ElseIf Left(Valeur, 1) = "-" Then
                            Peuples(i).MonstreIndice(j - 1) = NombreRaces - Right(Valeur, Len(Valeur) - 1) - 1
                        Else
                            Peuples(i).MonstreIndice(j - 1) = Val(Valeur) - 1
                        End If
                    Case "MonstreChance": Peuples(i).MonstreChance(j - 1) = Val(Valeur)
                    Case "MonstreFichierIA": Peuples(i).MonstreFichierIA(j - 1) = Valeur
                    Case "NomChateau": Peuples(i).NomChateau = Valeur
                    Case "NomChateau" & Langues.Suffixe: Peuples(i).NomChateau = Valeur
                    Case "ChateauMaxVie": Peuples(i).ChateauMaxVie = Val(Valeur)
                    Case "ChateauLargeur": Peuples(i).ChateauLargeur = Val(Valeur)
                    Case "ChateauHauteur": Peuples(i).ChateauHauteur = Val(Valeur)
                    Case "ChateauEntreePositionX": Peuples(i).ChateauEntreePositionX = Val(Valeur)
                    Case "ChateauEntreePositionY": Peuples(i).ChateauEntreePositionY = Val(Valeur)
                End Select
            
            Case SectionINI_Effets:
                Select Case Temp
                    Case "EffetClique": EffetClique = Val(Valeur) - 1
                    Case "EffetMaisonEnFeu": EffetMaisonEnFeu = Val(Valeur) - 1
                    Case "EffetMaisonDetruite": EffetMaisonDetruite = Val(Valeur) - 1
                    Case "EffetChateauDetruit": EffetChateauDetruit = Val(Valeur) - 1
                    Case "EffetNouvelObjectif": EffetNouvelObjectif = Val(Valeur) - 1
                    Case "EffetDureeRestante": EffetDureeRestante = Val(Valeur) - 1
                    Case "EffetDureeRestante2": EffetDureeRestante2 = Val(Valeur) - 1
                    
                    Case "Nom": Effets(i).Nom = Valeur
                    Case "Fichier": Effets(i).Fichier = Valeur
                    Case "Largeur": Effets(i).Largeur = Val(Valeur)
                    Case "Hauteur": Effets(i).Hauteur = Val(Valeur)
                    Case "NombreEtapes": Effets(i).NombreEtapes = Val(Valeur)
                    Case "DureeEtape": Effets(i).DureeEtape = Val(Valeur)
                    Case "Son": Effets(i).Son = Valeur
                End Select
                
            Case SectionINI_Races:
                'Chargement des paramètres des races.
                Select Case Temp
                    Case "ChancesRaceAleatoire": ChancesRaceAleatoire = Val(Valeur)
                    Case "JouerToutesLesRaces": JouerToutesLesRaces = Val(Valeur)
                    Case "Nom": Races(i).NomM = Valeur
                                Races(i).NomF = Valeur
                    Case "Nom" & Langues.Suffixe: Races(i).NomM = Valeur
                                                  Races(i).NomF = Valeur
                    Case "NomM": Races(i).NomM = Valeur
                    Case "NomM" & Langues.Suffixe: Races(i).NomM = Valeur
                    Case "NomF": Races(i).NomF = Valeur
                    Case "NomF" & Langues.Suffixe: Races(i).NomF = Valeur
                    Case "NomPluriel": Races(i).NomMPluriel = Valeur
                                       Races(i).NomFPluriel = Valeur
                    Case "NomPluriel" & Langues.Suffixe: Races(i).NomMPluriel = Valeur
                                       Races(i).NomFPluriel = Valeur
                    Case "NomMPluriel": Races(i).NomMPluriel = Valeur
                    Case "NomMPluriel" & Langues.Suffixe: Races(i).NomMPluriel = Valeur
                    Case "NomFPluriel": Races(i).NomFPluriel = Valeur
                    Case "NomFPluriel" & Langues.Suffixe: Races(i).NomFPluriel = Valeur
                    Case "NomAleatoire": Races(i).NomAleatoire = Val(Valeur)
                    Case "Jouable": Races(i).Jouable = Val(Valeur)
                    Case "Description": Races(i).Description = Valeur
                    Case "Description" & Langues.Suffixe: Races(i).Description = Valeur
                    
                    Case "RestrictionMaculin": Races(i).RestrictionMaculin = Val(Valeur)
                    Case "RestrictionFeminin": Races(i).RestrictionFeminin = Val(Valeur)
                    Case "Largeur": Races(i).Largeur = Val(Valeur)
                    Case "Hauteur": Races(i).Hauteur = Val(Valeur)
                    Case "Portrait": Races(i).Portrait = Valeur
                    Case "NombreCategoriesApparences":
                        Races(i).NombreCategoriesApparences = Val(Valeur)
                        ReDim Races(i).DossiersCategoriesF(Val(Valeur - 1))
                        ReDim Races(i).DossiersCategoriesM(Val(Valeur - 1))
                    Case "DossiersCategoriesF":
                        Races(i).DossiersCategoriesF(j - 1) = Valeur
                        If Left(Valeur, 1) = "." Then
                            AffApparence.Ajouter_Dossier Races(i).DossiersCategoriesF(j - 1)
                        End If
                    Case "DossiersCategoriesM":
                        Races(i).DossiersCategoriesM(j - 1) = Valeur
                        If Left(Valeur, 1) = "." Then
                            AffApparence.Ajouter_Dossier Races(i).DossiersCategoriesM(j - 1)
                        End If
                    Case "DossiersCategories":
                        Races(i).DossiersCategoriesF(j - 1) = Valeur
                        Races(i).DossiersCategoriesM(j - 1) = Valeur
                        If Left(Valeur, 1) = "." Then
                            AffApparence.Ajouter_Dossier Races(i).DossiersCategoriesF(j - 1)
                        End If
                    Case "DefinirApparence": Races(i).DefinirApparence = Val(Valeur)
                    Case "CheminApparenceM":
                        If Left(Valeur, 1) = "." Then
                            Races(i).CheminApparenceM = Right(Valeur, Len(Valeur) - 1)
                            AffPerso.Ajouter_Dossier Races(i).CheminApparenceM
                            Races(i).CheminAbsolu = True
                        Else
                            Races(i).CheminApparenceM = Valeur
                            Races(i).CheminAbsolu = False
                        End If
                    Case "CheminApparenceF":
                        If Left(Valeur, 1) = "." Then
                            Races(i).CheminApparenceF = Right(Valeur, Len(Valeur) - 1)
                            AffPerso.Ajouter_Dossier Races(i).CheminApparenceF
                            Races(i).CheminAbsolu = True
                        Else
                            Races(i).CheminApparenceF = Valeur
                            Races(i).CheminAbsolu = False
                        End If
                    Case "CheminApparence":
                        If Left(Valeur, 1) = "." Then
                            Races(i).CheminApparenceM = Right(Valeur, Len(Valeur) - 1)
                            Races(i).CheminApparenceF = Right(Valeur, Len(Valeur) - 1)
                            AffPerso.Ajouter_Dossier Races(i).CheminApparenceM
                            Races(i).CheminAbsolu = True
                        Else
                            Races(i).CheminApparenceF = Valeur
                            Races(i).CheminApparenceM = Valeur
                            Races(i).CheminAbsolu = False
                        End If
                    Case "ApparenceVerouillee":
                        Races(i).ApparenceVerouillee = Val(Valeur)
                    Case "Vie": Races(i).Vie = Val(Valeur)
                    Case "Energie": Races(i).Energie = Val(Valeur)
                    Case "Magie": Races(i).Magie = Val(Valeur)
                    Case "Moral": Races(i).Moral = Val(Valeur)
                    Case "Attaque": Races(i).Attaque = Val(Valeur)
                    Case "Defense": Races(i).Defense = Val(Valeur)
                    Case "CoefVie": Races(i).CoefVie = Val(Valeur)
                    Case "CoefEnergie": Races(i).CoefEnergie = Val(Valeur)
                    Case "CoefMagie": Races(i).CoefMagie = Val(Valeur)
                    Case "CoefMoral": Races(i).CoefMoral = Val(Valeur)
                    Case "CoefAttaque": Races(i).CoefAttaque = Val(Valeur)
                    Case "CoefDefense": Races(i).CoefDefense = Val(Valeur)
                    Case "VitesseRegenerationVie": Races(i).VitesseRegenerationVie = Val(Valeur)
                    Case "VitesseRegenerationEnergie": Races(i).VitesseRegenerationEnergie = Val(Valeur)
                    Case "VitesseRegenerationMagie": Races(i).VitesseRegenerationMagie = Val(Valeur)
                    Case "VitesseRegenerationMoral": Races(i).VitesseRegenerationMoral = Val(Valeur)
                    
                    Case "Parole": Races(i).Parole = Val(Valeur)
                    Case "Experience": Races(i).Experience = Val(Valeur)
                    Case "Equiper": Races(i).Equiper = Val(Valeur)
                    
                    Case "CoefCompSpeciale": Races(i).CoefCompSpeciales(j - 1) = Val(Valeur)
                    Case "CoefCompRessource": Races(i).CoefCompRessources(j - 1) = Val(Valeur)
                    Case "CoefCompService": Races(i).CoefCompServices(j - 1) = Val(Valeur)
                    Case "CoefCompObjet": Races(i).CoefCompObjets(j - 1) = Val(Valeur)

                    Case "Volant": Races(i).Volant = Val(Valeur)
                    Case "CommentairesSpecifique": Races(i).CommentairesSpecifique = Val(Valeur) - 1
                    Case "BonusMaxRessources": Races(i).BonusMaxRessources = Val(Valeur)
                    Case "PorteeAttaque": Races(i).PorteeAttaque = Val(Valeur)
                    Case "EffetAttaque":
                        If Valeur = "+" Then
                            Races(i).EffetAttaque = NombreEffets - 1
                        Else
                            Races(i).EffetAttaque = Val(Valeur) - 1
                        End If
                    Case "AttaqueCauseEffet": Races(i).AttaqueCauseEffet = Val(Valeur) - 1
                    Case "AttaqueResistance": Races(i).AttaqueResistance = Val(Valeur) - 1
                    Case "Armure": Races(i).Armure = Val(Valeur)
                    Case "Vitesse": Races(i).Vitesse = Val(Valeur)
                    Case "Acceleration": Races(i).Acceleration = Val(Valeur)
                    Case "Pietinement": Races(i).Pietinement = Val(Valeur)
                    'Case "Projectile": Races(i).Projectile = Val(Valeur)
                    Case "Inattaquable": Races(i).Inattaquable = Val(Valeur)
                    
                    Case "AutoResurrection": Races(i).AutoResurrection = Val(Valeur)
                    
                    Case "IA_Pacifique": Races(i).IA_Pacifique = Val(Valeur)
                    Case "IA_Manger": Races(i).IA_Manger = Val(Valeur)
                    Case "IA_Temeraire": Races(i).IA_Temeraire = Val(Valeur)
                    Case "IA_Protecteur": Races(i).IA_Protecteur = Val(Valeur)
                    Case "IA_Egoiste": Races(i).IA_Egoiste = Val(Valeur)
                    Case "IA_Berserk": Races(i).IA_Berserk = Val(Valeur)
                    Case "IA_Incommandable": Races(i).IA_Incommandable = Val(Valeur)
                    Case "IA_Selectionner": Races(i).IA_Selectionner = Val(Valeur)
                    Case "IA_Fichier": Races(i).IA_Fichier = Valeur
                    Case "Infatigable": Races(i).Infatigable = Val(Valeur)
                    Case "Kamikaze": Races(i).Kamikaze = Val(Valeur)
                    Case "RestaurationComplete": Races(i).RestaurationComplete = Val(Valeur)
                    Case "Argent": Races(i).Argent = Val(Valeur)
                    
                    Case "Resistance": Races(i).Resistances(j - 1) = Val(Valeur)
                    Case "Resistances": Races(i).Resistances(j - 1) = Val(Valeur)
                    Case "Ressource": Races(i).Ressources(j - 1) = Val(Valeur)
                    Case "Ressources": Races(i).Ressources(j - 1) = Val(Valeur)
                    Case "DepartSansRessource": Races(i).DepartSansRessources = Val(Valeur)
                    Case "DepartSansRessources": Races(i).DepartSansRessources = Val(Valeur)
                    
                    Case "NombreObjets":
                        If i = -1 Then 'Charge les objets communs à chaque race.
                            NombreObjetsCommmuns = Val(Valeur)
                            ReDim ObjetCommmuns(NombreObjetsCommmuns - 1)
                        Else 'Charge les objets propres à chaque race.
                            Races(i).NombreObjets = Val(Valeur)
                            If Races(i).NombreObjets > 0 Then
                                ReDim Races(i).Objet(Races(i).NombreObjets - 1)
                            End If
                        End If
                    Case "Objet":
                        If Valeur = "+" Or Valeur = "-" Then
                            If i = -1 Then
                                ObjetCommmuns(j - 1) = NombreObjets - 1
                            Else
                                Races(i).Objet(j - 1) = NombreObjets - 1
                                Races(i).ObjetRace.Add NombreObjets - 1
                            End If
                        ElseIf Left(Valeur, 1) = "+" Then
                            If i = -1 Then
                                ObjetCommmuns(j - 1) = Right(Valeur, Len(Valeur) - 1) + NombreObjets - 1
                            Else
                                Races(i).Objet(j - 1) = Right(Valeur, Len(Valeur) - 1) + NombreObjets - 1
                                Races(i).ObjetRace.Add Right(Valeur, Len(Valeur) - 1) + NombreObjets - 1
                            End If
                        ElseIf Left(Valeur, 1) = "-" Then
                            If i = -1 Then
                                ObjetCommmuns(j - 1) = -Right(Valeur, Len(Valeur) - 1) + NombreObjets - 1
                            Else
                                Races(i).Objet(j - 1) = -Right(Valeur, Len(Valeur) - 1) + NombreObjets - 1
                                Races(i).ObjetRace.Add -Right(Valeur, Len(Valeur) - 1) + NombreObjets - 1
                            End If
                        Else
                            If i = -1 Then 'Charge les objets communs à chaque race.
                                ObjetCommmuns(j - 1) = Val(Valeur) - 1
                            Else 'Charge les objets propres à chaque race.
                                Races(i).Objet(j - 1) = Val(Valeur) - 1
                                'ReDim Races(i).ObjetRace(UBound(Races(i).ObjetRace()) + 1)
                                'Races(i).ObjetRace(UBound(Races(i).ObjetRace())) = Val(Valeur) - 1
                                Races(i).ObjetRace.Add Val(Valeur) - 1
                            End If
                        End If
                        
                    Case "NombreEquipement":
                        Races(i).NombreEquipement = Val(Valeur)
                        If Races(i).NombreEquipement > 0 Then
                            ReDim Races(i).Equipement(Races(i).NombreEquipement - 1)
                        End If
                    Case "Equipement":
                        If Valeur = "+" Or Valeur = "-" Then
                            Races(i).Equipement(j - 1) = NombreObjets - 1
                        ElseIf Left(Valeur, 1) = "+" Then
                            Races(i).Equipement(j - 1) = Right(Valeur, Len(Valeur) - 1) + NombreObjets - 1
                        ElseIf Left(Valeur, 1) = "-" Then
                            Races(i).Equipement(j - 1) = -Right(Valeur, Len(Valeur) - 1) + NombreObjets - 1
                        Else
                            Races(i).Equipement(j - 1) = Val(Valeur) - 1
                        End If
                    Case "Statut":
                        If Val(Valeur) = 0 Then
                            'Si 0, on vide tous les statuts.
                            Collection_Vider Races(i).Statut
                        Else
                            Races(i).Statut.Add Val(Valeur) - 1
                        End If
                    Case "IndiceBestiaire":
                        If Valeur = "+" Or Valeur = "-" Then
                            Races(i).IndiceBestiaire = i + 1
                        ElseIf Left(Valeur, 1) = "+" Then
                            Races(i).IndiceBestiaire = Right(Valeur, Len(Valeur) - 1) + Val(i) + 1
                        ElseIf Left(Valeur, 1) = "-" Then
                            Races(i).IndiceBestiaire = Right(Valeur, Len(Valeur) - 1) - Val(i) + 1
                        Else
                            Races(i).IndiceBestiaire = Val(Valeur)
                        End If
                    Case "UtilisationMonture":
                        Races(i).UtilisationMonture = Val(Valeur)
                    Case "Tombe":
                        If Valeur = "+" Or Valeur = "-" Then
                            Races(i).Tombe = NombreDecors - 1
                        ElseIf Left(Valeur, 1) = "+" Then
                            Races(i).Tombe = Right(Valeur, Len(Valeur) - 1) + NombreDecors - 1
                        ElseIf Left(Valeur, 1) = "-" Then
                            Races(i).Tombe = -Right(Valeur, Len(Valeur) - 1) + NombreDecors - 1
                        Else
                            Races(i).Tombe = Val(Valeur) - 1
                        End If
                    Case "EffetMort":
                        If Valeur = "+" Then
                            Races(i).EffetMort = NombreEffets - 1
                        Else
                            Races(i).EffetMort = Val(Valeur) - 1
                        End If
                End Select
            Case SectionINI_Decors:
                With Decors(i)
                Select Case Temp
                    Case "Nom": .Nom = Valeur
                    Case "Image": .Image = Valeur
                    Case "Largeur": .Largeur = Val(Valeur)
                    Case "Hauteur": .Hauteur = Val(Valeur)
                    Case "NombreEtapeAnimation": .NombreEtapeAnimation = Val(Valeur)
                    Case "Duree": .Duree = Val(Valeur)
                End Select
                End With
            Case SectionINI_Resistances:
                Select Case Temp
                    Case "Defaut": ResistanceDefaut = Val(Valeur)
                    Case "Minimum": ResistanceMinimum = Val(Valeur)
                    Case "Maximum": ResistanceMaximum = Val(Valeur)
                    Case "Nom": Resistances(i).Nom = Valeur
                    Case "NomComplet": Resistances(i).NomComplet = Valeur
                    Case "Attaque": Resistances(i).Attaque = Valeur
                    Case "Abreviation": Resistances(i).Abreviation = Valeur
                End Select
            Case SectionINI_Statuts:
                Select Case Temp
                    Case "Nom": Statuts(i).Nom = Valeur
                    Case "Resistance": Statuts(i).Resistances(j - 1) = Val(Valeur)
                    Case "Resistances": Statuts(i).Resistances(j - 1) = Val(Valeur)
                End Select
            End Select
        End If
    End If
    Exit Sub
Erreur:
    'Message d'erreur précisant son emplacment.
    MsgBox MessageErreurFichier & Chr(10) & Chr(13) & Fichier & Chr(10) & Chr(13) & MessageErreurLigne & Ligne, vbCritical, TitreErreur
    End
End Sub

Private Function Gestion_Cle_Batiments(ByVal ID As String, ByVal FichierMod As Boolean) As Integer
    'i est l'indice du bâtiment correspondant à la clé.
    Dim i As Integer
    Dim j As Integer
    If ID <> "" Then
        i = -1
        For j = 0 To NombreTypeBatiments - 1
            If Batiments(j).ID = ID Then
                i = j
                Exit For
            End If
        Next j
        If i = -1 Then 'Si aucune clé ne correspond à la clé donnée, on crée un nouveau bâtiment.
            i = NombreTypeBatiments
            NombreTypeBatiments = NombreTypeBatiments + 1
            ReDim Preserve Batiments(i)
            ReDim Preserve Batiments(i).PrixRessources(NombreRessources - 1)
            Batiments(i).CreerProprietaire = -1
            Batiments(i).GainObjet = -1
            Batiments(i).PerteObjet = -1
            Batiments(i).ID = ID
            Batiments(i).FichierMod = FichierMod
        End If
    End If
    Gestion_Cle_Batiments = i
End Function
Private Function Gestion_Cle_CompetencesObjets(ByVal ID As String, ByVal FichierMod As Boolean) As Integer
    'i est l'indice de la compétence objet correspondant à la clé.
    Dim i As Integer
    Dim j As Integer
    If ID <> "" Then
        i = -1
        For j = 0 To NombreCompetencesObjets - 1
            If CompetencesObjets(j).ID = ID Then
                i = j
                Exit For
            End If
        Next j
        If i = -1 Then 'Si aucune clé ne correspond à la clé donnée, on crée un nouveau bâtiment.
            i = NombreCompetencesObjets
            NombreCompetencesObjets = NombreCompetencesObjets + 1
            ReDim Preserve CompetencesObjets(i)
            CompetencesObjets(i).ID = ID
            CompetencesObjets(i).FichierMod = FichierMod
            CompetencesObjets(i).NoListe = i
            'Redimentionne les tableaux des objets et des races.
            For j = 0 To NombreObjets - 1
                ReDim Preserve Objets(j).BonusCompetencesObjets(NombreCompetencesObjets - 1)
            Next j
            For j = 0 To NombreRaces - 1
                ReDim Preserve Races(j).CoefCompObjets(NombreCompetencesObjets - 1)
            Next j
        End If
    End If
    Gestion_Cle_CompetencesObjets = i
End Function

Private Function Parametre_Filtre(ByVal Ligne As String) As String
    Ligne = Trim(Ligne)
    While Left(Ligne, 1) = Chr(9)
        Ligne = Right(Ligne, Len(Ligne) - 1)
    Wend
    Parametre_Filtre = Trim(Ligne)
End Function
Private Function Parametre_Indice(ByVal Ligne As String) As Long
    'Renvoie l'indice situé entre les parentèses.
    If InStr(Ligne, "(") > 0 And InStr(Ligne, ")") > 0 Then
        Parametre_Indice = Trim(Mid(Ligne, InStr(Ligne, "(") + 1, InStr(Ligne, ")") - InStr(Ligne, "(") - 1))
    End If
End Function
Private Function Parametre_Gauche(ByVal Ligne As String) As String
    'Renvoie le paramètre à gauche du "=" de la ligne lue.
    Parametre_Gauche = Trim(Left(Ligne, InStr(Ligne, "=") - 1))
End Function
Private Function Parametre_Gauche2(ByVal Ligne As String) As String
    'Renvoie le paramètre à gauche du "=" sans les paranthèses.
    If InStr(Ligne, "(") > 0 And InStr(Ligne, ")") > 0 Then
        Parametre_Gauche2 = Trim(Left(Ligne, InStr(Ligne, "(") - 1))
    Else
        Parametre_Gauche2 = Ligne
    End If
End Function
Private Function Parametre_Droite(ByVal Ligne As String) As String
    'Renvoie le paramètre à droite du "=" de la ligne lue.
    Parametre_Droite = Trim(Right(Ligne, Len(Ligne) - InStr(Ligne, "=")))
End Function

Private Sub Verification_Donnees()
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim TriOk As Boolean 'Sert à vérifier les tris des éléments.
    Dim Temp  As Collection
    Set Temp = New Collection
    
    'VERIFICATION DES BATIMENTS.
    For i = 0 To NombreTypeBatiments - 1
        With Batiments(i)
        If .Prix = 0 Then
            'Affecte un prix en fonction de son coût en ressources.
            For j = 0 To NombreRessources - 1
                .Prix = .Prix + .PrixRessources(j) * Ressources(j).PrixVente
            Next j
            If .Empereur Then .Prix = .Prix + 100000
        End If
        .Prix = .Prix + .PrixArgent
        'Debug.Print .Nom & " : " & .Prix
        'Si il n'y a pas d'image en paramètres, c'est le nom qui est pris comme fichier image.
        If .Image = "" Then .Image = .Nom
        If .CheminImage = "" Then .CheminImage = BatimentsCheminImageDefaut
        If .ImageBouton = "" Then .ImageBouton = .Image
        If .CoefVitesseConstruction = 0 Then .CoefVitesseConstruction = 1
        .RessourceFabrique = .RessourceFabrique - 1
        .Jouable = .NoEpoque <= NombreEpoques
        If .Jouable Then
            'Si l'époque d'un batiment dépasse les époques officielles, le batiment n'est pas jouable.
            NombreTypeBatimentsJouables = NombreTypeBatimentsJouables + 1
        End If
        Temp.Add i 'Mémorise les bâtiments pour la suite.
        End With
    Next i
    
    'Tri les bâtiments selon leur époque et leur prix.
    ReDim BatimentsTri(NombreTypeBatiments - 1)
    For i = 1 To NombreTypeBatiments
        k = 1
        For j = 1 To Temp.Count
            If (j = 1 Or _
               (Batiment_Prix(Temp(j)) < Batiment_Prix(Temp(k)) And Batiment_NoEpoque(Temp(j)) <= Batiment_NoEpoque(Temp(k))) Or _
                Batiment_NoEpoque(Temp(j)) < Batiment_NoEpoque(Temp(k))) And _
               (Not Batiment_Empereur(Temp(j)) Or Batiment_NoEpoque(Temp(k)) = NombreEpoques + 1) Then
                k = j
            End If
        Next j
        BatimentsTri(i - 1) = Temp(k)
        'Debug.Print i - 1 & " : " & Batiment_NoEpoque(BatimentsTri(i - 1)) & " : " & Batiment_Nom(BatimentsTri(i - 1))
        Temp.Remove k 'Evite de réutiliser 2 fois le même bâtiment.
        If Temp.Count = 0 Then Exit For
    Next i
    
    'VERIFICATION DES RESSOURCES.
    For i = 0 To UBound(Ressources())
        With Ressources(i)
        If .Image = "" Then .Image = .Nom 'Si il n'y a pas d'image en paramètres, c'est le nom qui est pris comme fichier image.
        End With
    Next i
    For i = 0 To UBound(TerrainRessources())
        Ressources(TerrainRessources(i).TypeRessource).Recoltable = True
        Ressources(TerrainRessources(i).TypeRessource).TypeTerrainRessource = i
    Next i
    'VERIFICATION DES OBJETS.
    For i = 0 To NombreObjets - 1
        'Affecte un fichier image aux objets qui n'en n'ont pas.
        If Objets(i).NomPluriel = "" Then Objets(i).NomPluriel = Objets(i).Nom
        If Objets(i).Image = "" Then Objets(i).Image = Objets(i).Nom
        'Mémorise que la ressource est nécessaire pour attaquer.
        If Objets(i).PerteRessourceQuantite > 0 Then
            Ressources(Objets(i).PerteRessourceIndice).NecessairePourAttaquer = True
        End If
        'Recherche à quelle époque est fabriquée l'objet.
        If Objets(i).Secret Then
            Objets(i).Epoque = -1
        Else
            For j = 0 To NombreTypeBatiments - 1
                If Batiments(j).Fabrique Then
                    'Recherche le batiment le plus précoce qui fabrique l'objet
                    For k = 0 To Batiments(j).NombreObjetsFabriques - 1
                        If Batiments(j).ObjetsFabriques(k) = i Then
                            If Objets(i).Epoque < Batiments(j).NoEpoque Then
                                Objets(i).Epoque = Batiments(j).NoEpoque
                            End If
                            k = Batiments(j).NombreObjetsFabriques
                            j = NombreTypeBatiments
                        End If
                    Next k
                End If
            Next j
            'Vérifie si l'on peut trouver cette ressource à cette époque.
            For j = 0 To NombreRessources - 1
                If Objets(i).Epoque >= 0 And _
                   Objets(i).PrixRessources(j) > 0 And _
                   Ressources(j).Epoque > Objets(i).Epoque - 1 Then
                    Objets(i).Epoque = Ressources(j).Epoque + 1
                End If
            Next j
        End If
    Next i
    'Tri les objets à fabriquer des bâtiments selon leur époque.
    For i = 0 To NombreTypeBatiments - 1
        If Batiment_NombreObjetsFabriques(i) > 0 Then
            Do
            TriOk = True
            For j = 1 To Batiment_NombreObjetsFabriques(i) - 1
                k = Batiment_ObjetFabrique(i, j - 1)
                l = Batiment_ObjetFabrique(i, j)
                If Objet_Epoque(l) < Objet_Epoque(k) Then
                    Batiment_ObjetFabrique(i, j - 1) = l
                    Batiment_ObjetFabrique(i, j) = k
                    'Debug.Print Objet_Nom(l) & "<->" & Objet_Nom(k)
                    TriOk = False
                End If
            Next j
            Loop Until TriOk
        End If
    Next i
    
    'VERIFICATION DES RACES.
    For i = 0 To NombreRaces - 1
        With Races(i)
        
        If .Jouable Then NombreRacesJouables = NombreRacesJouables + 1
        
        'Rectifie les valeurs qui n'ont pas été chargés.
        If .Largeur = 0 Then .Largeur = PersosLargeurDefaut
        If .Hauteur = 0 Then .Hauteur = PersosHauteurDefaut
        If .Vie = 0 Then .Vie = PersosMaxVie
        If .Energie = 0 Then .Energie = PersosMaxEnergie
        If .Magie = 0 Then .Magie = PersosMaxMagie
        If .Moral = 0 Then .Moral = PersosMaxMoral
        If .Attaque = 0 Then .Attaque = PersosAttaque
        If .Defense = 0 Then .Defense = PersosDefense
        If .CoefVie = 0 Then .CoefVie = 1
        If .CoefEnergie = 0 Then .CoefEnergie = 1
        If .CoefMagie = 0 Then .CoefMagie = 1
        If .CoefMoral = 0 Then .CoefMoral = 1
        If .CoefAttaque = 0 Then .CoefAttaque = 1
        If .CoefDefense = 0 Then .CoefDefense = 1
        
        'Compétences.
        For j = 0 To NombreCompetencesSpeciales - 1
            If .CoefCompSpeciales(j) = 0 Then .CoefCompSpeciales(j) = 1
        Next j
        For j = 0 To NombreRessources - 1
            If .CoefCompRessources(j) = 0 Then .CoefCompRessources(j) = 1
        Next j
        For j = 0 To NombreServices - 2
            If .CoefCompServices(j) = 0 Then .CoefCompServices(j) = 1
        Next j
        For j = 0 To NombreCompetencesObjets - 1
            If .CoefCompObjets(j) = 0 Then .CoefCompObjets(j) = 1
        Next j
        
        If .Infatigable Then
            .VitesseRegenerationEnergie = 0.01
            .VitesseRegenerationMoral = 0.01
        End If
        If .Vitesse < 0 Then .Vitesse = PersosVitesse
        If .Acceleration <= 0 Then .Acceleration = PersosAcceleration
        'Charge les objets communs à chaque race.
        k = 0
        .NombreObjets = .NombreObjets + NombreObjetsCommmuns
        If .NombreObjets > 0 Then
            ReDim Preserve .Objet(.NombreObjets - 1)
            For j = .NombreObjets - NombreObjetsCommmuns To .NombreObjets - 1
                .Objet(j) = ObjetCommmuns(k)
                k = k + 1
            Next j
        End If
        
        'Charge les statuts et les résistances associées.
        For j = 1 To .Statut.Count
            For k = 0 To NombreResistances - 1
                If Statuts(.Statut.Item(j)).Resistances(k) <> ResistanceDefaut Then
                    .Resistances(k) = Statuts(.Statut.Item(j)).Resistances(k)
                End If
            Next k
        Next j
        For k = 0 To NombreResistances - 1
            If .Resistances(k) < ResistanceMinimum Then
                .Resistances(k) = ResistanceMinimum
            ElseIf .Resistances(k) > ResistanceMaximum Then
                   .Resistances(k) = ResistanceMaximum
            End If
        Next k
        
        End With
    Next i
    Bestiaire_Tri_Alphabetique 'Trie le bestiaire par ordre alhpabétique
End Sub

' ************************************************************************************
' *                              Propriétés de la classe                             *
' ************************************************************************************

'Etiquettes de chargement.
Public Property Get Etiquette_Chargement(ByVal Index As Integer) As String
    Etiquette_Chargement = EtiquettesChargement(Index)
End Property

'Propriétés des époques.
Public Property Get Epoque_Nom(ByVal Index As Integer) As String
    Epoque_Nom = Epoques(Index).Nom
End Property
Public Property Get Epoque_Duree(ByVal Index As Integer) As Long
    Epoque_Duree = Epoques(Index).Duree
End Property
Public Property Get Epoque_BonusVieChateau(ByVal Index As Integer) As Long
    Epoque_BonusVieChateau = Epoques(Index).BonusVieChateau
End Property
Public Property Get Epoque_EmpereurNom(ByVal Index As Integer, ByVal Feminin As Boolean) As String
    If Feminin Then
        Epoque_EmpereurNom = Epoques(Index).EmpereurNomF
    Else
        Epoque_EmpereurNom = Epoques(Index).EmpereurNomM
    End If
End Property
Public Property Get Epoque_EmpereurPrix(ByVal Index As Integer) As Long
    Epoque_EmpereurPrix = Epoques(Index).EmpereurPrix
End Property

'Pripriétés des ressources de départ.
Public Property Get RessourcesDeparts_Nom(ByVal Index As Integer) As String
    RessourcesDeparts_Nom = RessourcesDeparts(Index).Nom
End Property
Public Property Get RessourcesDeparts_ChateauxRessourcesDepart(ByVal Index As Integer) As Double
    RessourcesDeparts_ChateauxRessourcesDepart = RessourcesDeparts(Index).ChateauxRessourcesDepart
End Property
Public Property Get RessourcesDeparts_PersosRessourcesDepart(ByVal Index As Integer) As Double
    RessourcesDeparts_PersosRessourcesDepart = RessourcesDeparts(Index).PersosRessourcesDepart
End Property
Public Property Get RessourcesDeparts_PersosArgentDepart(ByVal Index As Integer) As Double
    RessourcesDeparts_PersosArgentDepart = RessourcesDeparts(Index).PersosArgentDepart
End Property

'Propriétés des vitesse d'époques.
Public Property Get VitesseEpoque_Nom(ByVal Index As Integer) As String
    VitesseEpoque_Nom = VitesseEpoques(Index).Nom
End Property
Public Property Get VitesseEpoque_Coef() As Single
    VitesseEpoque_Coef = VitesseEpoques(VitesseEpoqueSelectionnee).Coef
End Property

'Propriétés des compétences spéciales.
Public Property Get Speciales_NomCompetences(ByVal Index As Integer) As String
    Speciales_NomCompetences = CompetencesSpeciales(Index)
End Property

'Propriétés des ressources.
Public Property Get Ressources_Nom(ByVal Index As Long) As String
    Ressources_Nom = Ressources(Index).Nom
End Property
Public Property Get Ressources_NomCompetence(ByVal Index As Long) As String
    Ressources_NomCompetence = Ressources(Index).NomCompetence
End Property
Public Property Get Ressources_Image(ByVal Index As Integer) As String
    Ressources_Image = Ressources(Index).Image
End Property
Public Property Get Ressources_CheminImage(ByVal Index As Integer) As String
    Ressources_CheminImage = Ressources(Index).CheminImage
End Property
Public Property Get Ressources_VitesseExtraction(ByVal Index As Long) As Double
    Ressources_VitesseExtraction = Ressources(Index).VitesseExtraction
End Property
Public Property Get Ressources_TypeMatierePremiere(ByVal Index As Long) As Long
    Ressources_TypeMatierePremiere = Ressources(Index).TypeMatierePremiere
End Property
Public Property Get Ressources_QuantiteMatierePremiere(ByVal Index As Long) As Double
    Ressources_QuantiteMatierePremiere = Ressources(Index).QuantiteMatierePremiere
End Property
Public Property Get Ressources_PrixVente(ByVal Index As Long) As Double
    Ressources_PrixVente = Ressources(Index).PrixVente
End Property
Public Property Get Ressources_NecessairePourConstruire(ByVal Index As Long) As Boolean
    Ressources_NecessairePourConstruire = Ressources(Index).NecessairePourConstruire
End Property
Public Property Get Ressources_NecessairePourAttaquer(ByVal Index As Long) As Boolean
    Ressources_NecessairePourAttaquer = Ressources(Index).NecessairePourAttaquer
End Property
Public Property Get Ressources_Recoltable(ByVal Index As Integer) As Boolean
    If Index >= 0 Then
        Ressources_Recoltable = Ressources(Index).Recoltable
    End If
End Property
Public Property Get Ressources_TypeTerrainRessource(ByVal Index As Integer) As Integer
    Ressources_TypeTerrainRessource = Ressources(Index).TypeTerrainRessource
End Property
Public Property Get Ressources_Mangeable(ByVal Index As Long) As Boolean
    Ressources_Mangeable = Ressources(Index).Mangeable
End Property
Public Property Get Ressources_GainVie(ByVal Index As Long) As Double
    Ressources_GainVie = Ressources(Index).GainVie
End Property
Public Property Get Ressources_GainEnergie(ByVal Index As Long) As Double
    Ressources_GainEnergie = Ressources(Index).GainEnergie
End Property
Public Property Get Ressources_GainMagie(ByVal Index As Long) As Double
    Ressources_GainMagie = Ressources(Index).GainMagie
End Property
Public Property Get Ressources_GainMoral(ByVal Index As Long) As Double
    Ressources_GainMoral = Ressources(Index).GainMoral
End Property
Public Property Get Ressources_Epoque(ByVal Index As Long) As Double
    Ressources_Epoque = Ressources(Index).Epoque
End Property

'Propriété des ressources diponibles sur le terrain.
Public Property Get TerrainRessources_Nom(ByVal Index As Long) As String
    TerrainRessources_Nom = TerrainRessources(Index).Nom
End Property
Public Property Get TerrainRessources_TypeRessource(ByVal Index As Long) As Integer
    TerrainRessources_TypeRessource = TerrainRessources(Index).TypeRessource
End Property
Public Property Get TerrainRessources_Quantite(ByVal Index As Long) As Double
    TerrainRessources_Quantite = TerrainRessources(Index).Quantite
End Property

'Propriétés des batiments.
Public Property Get Batiment_Tri(ByVal Index As Integer)
    Batiment_Tri = BatimentsTri(Index)
End Property
Public Property Get Batiment_TriInverse(ByVal Index As Integer)
    Dim i As Integer
    For i = 0 To NombreTypeBatiments - 1
        If BatimentsTri(i) = Index Then
            Batiment_TriInverse = i
            Exit For
        End If
    Next i
End Property
Public Property Get Batiment_Numero(ByVal ID As String, _
                                    Optional ByVal MessageErreur As String, _
                                    Optional ByVal FichierErreur As String) As Integer
    'Renvoie le numéro en fonction de l'ID entrée.
    Dim i As Integer
    i = -1
    Batiment_Numero = -1
    For i = 0 To NombreTypeBatiments - 1
        If Batiments(i).ID = ID Then
            Batiment_Numero = i
            Exit For
        End If
    Next i
    If Batiment_Numero = -1 Then
        Editer_Texte FichierErreur
        MsgBox MessageErreur, vbCritical, "Erreur"
        End
    End If
End Property
Public Property Get Batiment_ID(ByVal Index As Long) As String
    Batiment_ID = Batiments(Index).ID
End Property
Public Property Get Batiment_Nom(ByVal Index As Long) As String
    Batiment_Nom = Batiments(Index).Nom
End Property
Public Property Get Batiment_Feminin(ByVal Index As Long) As Boolean
    Batiment_Feminin = Batiments(Index).Feminin
End Property
Public Property Get Batiment_Jouable(ByVal Index As Long) As Boolean
    Batiment_Jouable = Batiments(Index).Jouable
End Property
Public Property Get Batiment_Image(ByVal Index As Long) As String
    Batiment_Image = Batiments(Index).Image
End Property
Public Property Get Batiment_CheminImage(ByVal Index As Long) As String
    Batiment_CheminImage = Batiments(Index).CheminImage
End Property
Public Property Get Batiment_ImageBouton(ByVal Index As Long) As String
    Batiment_ImageBouton = Batiments(Index).ImageBouton
End Property
Public Property Get Batiment_CheminImageBouton(ByVal Index As Long) As String
    Batiment_CheminImageBouton = Batiments(Index).CheminImageBouton
End Property
Public Property Get Batiment_Metier(ByVal Index As Long, ByVal Feminin As Boolean) As String
    If Feminin Then
        Batiment_Metier = Batiments(Index).MetierF
    Else
        Batiment_Metier = Batiments(Index).MetierM
    End If
End Property
Public Property Get Batiment_Description(ByVal Index As Integer) As String
    Batiment_Description = Batiments(Index).Description
End Property
Public Property Get Batiment_Empereur(ByVal Index As Integer) As Boolean
    Batiment_Empereur = Batiments(Index).Empereur
End Property
Public Property Get Batiment_Limite(ByVal Index As Integer) As Integer
    Batiment_Limite = Batiments(Index).Limite
End Property
Public Property Get Batiment_CreerProprietaire(ByVal Index As Integer) As Integer
    Batiment_CreerProprietaire = Batiments(Index).CreerProprietaire
End Property
Public Property Get Batiment_Roi(ByVal Index As Integer) As Boolean
    Batiment_Roi = Batiments(Index).Roi
End Property
Public Property Get Batiment_RestrictionRoi(ByVal Index As Integer) As Boolean
    Batiment_RestrictionRoi = Batiments(Index).RestrictionRoi
End Property
Public Property Get Batiment_GainObjet(ByVal Index As Integer) As Integer
    Batiment_GainObjet = Batiments(Index).GainObjet
End Property
Public Property Get Batiment_PerteObjet(ByVal Index As Integer) As Integer
    Batiment_PerteObjet = Batiments(Index).PerteObjet
End Property
Public Property Get Batiment_Tir(ByVal Index As Integer) As Boolean
    Batiment_Tir = Batiments(Index).Tir
End Property
Public Property Get Batiment_Marche(ByVal Index As Integer) As Boolean
    Batiment_Marche = Batiments(Index).Marche
End Property
Public Property Get Batiment_Prix(ByVal Index As Long) As Long
    Batiment_Prix = Batiments(Index).Prix
End Property
Public Property Get Batiment_PrixArgent(ByVal Index As Long) As Long
    Batiment_PrixArgent = Batiments(Index).PrixArgent
End Property
Public Property Get Batiment_PrixRessource(ByVal Index As Long, ByVal Ressource As Long) As Long
    Batiment_PrixRessource = Batiments(Index).PrixRessources(Ressource)
End Property
Public Property Get Batiment_MaxVie(ByVal Index As Long) As Long
    Batiment_MaxVie = Batiments(Index).MaxVie
End Property
Public Property Get Batiment_CoefVitesseConstruction(ByVal Index As Integer) As Single
    Batiment_CoefVitesseConstruction = Batiments(Index).CoefVitesseConstruction
End Property
Public Property Get Batiment_RessourceFabrique(ByVal Index As Long) As Long
    Batiment_RessourceFabrique = Batiments(Index).RessourceFabrique
End Property
Public Property Get Batiment_NombreObjetsFabriques(ByVal Index As Long) As Long
    Batiment_NombreObjetsFabriques = Batiments(Index).NombreObjetsFabriques
End Property
Public Property Get Batiment_ObjetFabrique(ByVal Index As Integer, ByVal Objet As Integer) As Integer
    Batiment_ObjetFabrique = Batiments(Index).ObjetsFabriques(Objet)
End Property
Public Property Let Batiment_ObjetFabrique(ByVal Index As Integer, ByVal Objet As Integer, ByVal Valeur As Integer)
    Batiments(Index).ObjetsFabriques(Objet) = Valeur
End Property
Public Property Get Batiment_TypeMatierePremiere(ByVal Index As Long) As Integer
    Dim i As Integer
    Dim j As Long
    Dim Temp As Integer
    If Batiments(Index).Fabrique Then
        Temp = -1
        Batiment_TypeMatierePremiere = -1
        For i = 0 To NombreRessources - 1
            For j = 0 To Batiment_NombreObjetsFabriques(Index) - 1
                If Objets(Batiment_ObjetFabrique(Index, j)).PrixRessources(i) > 0 Then
                    Temp = i
                    If Rnd > 0.5 Then
                        Batiment_TypeMatierePremiere = i 'Objets(Batiment_ObjetFabrique(index, 0)).PrixRessources(i)
                        j = Batiment_NombreObjetsFabriques(Index)
                        i = NombreRessources
                    End If
                End If
            Next j
        Next i
        If Temp >= 0 And Batiment_TypeMatierePremiere = -1 Then
            Batiment_TypeMatierePremiere = Temp
        End If
    ElseIf Batiments(Index).Service Then
        Batiment_TypeMatierePremiere = Services(Batiments(Index).Service).TypeMatierePremiere
    ElseIf Batiments(Index).RessourceFabrique >= 0 Then
        Batiment_TypeMatierePremiere = Ressources(Batiments(Index).RessourceFabrique).TypeMatierePremiere
    Else
        Batiment_TypeMatierePremiere = -1
    End If
End Property
Public Property Get Batiment_Magasin(ByVal Index As Long) As Boolean
    Batiment_Magasin = Batiments(Index).Magasin
End Property
Public Property Get Batiment_Fabrique(ByVal Index As Integer) As Boolean
    Batiment_Fabrique = Batiments(Index).Fabrique
End Property
Public Property Get Batiment_Service(ByVal Index As Integer) As Integer
    Batiment_Service = Batiments(Index).Service
End Property
Public Property Get Batiment_MaxStockVente(ByVal Index As Long) As Double
    Batiment_MaxStockVente = Batiments(Index).MaxStock
End Property
Public Property Get Batiment_Niveau(ByVal Index As Long) As Long
    Batiment_Niveau = Batiments(Index).Niveau
End Property
Public Property Get Batiment_NoEpoque(ByVal Index As Long) As Integer
    Batiment_NoEpoque = Batiments(Index).NoEpoque
End Property
Public Property Get Batiment_Largeur(ByVal Index As Long) As Integer
    Batiment_Largeur = Batiments(Index).Largeur
End Property
Public Property Get Batiment_Hauteur(ByVal Index As Long) As Integer
    Batiment_Hauteur = Batiments(Index).Hauteur
End Property
Public Property Get Batiment_EntreePositionX(ByVal Index As Long) As Long
    Batiment_EntreePositionX = Batiments(Index).EntreePositionX
End Property
Public Property Get Batiment_EntreePositionY(ByVal Index As Long) As Long
    Batiment_EntreePositionY = Batiments(Index).EntreePositionY
End Property
Public Property Get Batiment_NombreEtapeFabrication(ByVal Index As Long) As Integer
    Batiment_NombreEtapeFabrication = Batiments(Index).NombreEtapeFabrication
End Property
Public Property Get Batiment_NombreEtapeAnimation(ByVal Index As Long) As Integer
    Batiment_NombreEtapeAnimation = Batiments(Index).NombreEtapeAnimation
End Property
Public Property Get Batiment_ApparencePerso(ByVal Index As Long) As Integer
    Batiment_ApparencePerso = Batiments(Index).ApparencePerso
End Property

'Paramètres des peuples.
Public Property Get Peuples_Nom(ByVal Index As Integer) As String
    Peuples_Nom = Peuples(Index).Nom
End Property
Public Property Get Peuples_Image(ByVal Index As Integer) As String
    Peuples_Image = Peuples(Index).Image
End Property
Public Property Get Peuples_CheminImage(ByVal Index As Integer) As String
    Peuples_CheminImage = Peuples(Index).CheminImage
End Property
Public Property Get Peuples_NombreMonstres(ByVal Index As Integer) As Integer
    Peuples_NombreMonstres = Peuples(Index).NombreMonstres
End Property
Public Property Get Peuples_MonstreIndice(ByVal Index As Integer, ByVal IndiceMonstre As Integer) As Integer
    Peuples_MonstreIndice = Peuples(Index).MonstreIndice(IndiceMonstre)
End Property
Public Property Get Peuples_MonstreChance(ByVal Index As Integer, ByVal IndiceMonstre As Integer) As Single
    Peuples_MonstreChance = Peuples(Index).MonstreChance(IndiceMonstre)
End Property
Public Property Get Peuples_MonstreFichierIA(ByVal Index As Integer, ByVal IndiceMonstre As Integer) As String
    Peuples_MonstreFichierIA = Peuples(Index).MonstreFichierIA(IndiceMonstre)
End Property
Public Property Get Peuples_NomChateau(ByVal Index As Integer) As String
    Peuples_NomChateau = Peuples(Index).NomChateau
End Property
Public Property Get Peuples_ChateauMaxVie(ByVal Index As Integer) As Long
    Peuples_ChateauMaxVie = Peuples(Index).ChateauMaxVie
End Property
Public Property Get Peuples_ChateauLargeur(ByVal Index As Integer) As Integer
    Peuples_ChateauLargeur = Peuples(Index).ChateauLargeur
End Property
Public Property Get Peuples_ChateauHauteur(ByVal Index As Integer) As Integer
    Peuples_ChateauHauteur = Peuples(Index).ChateauHauteur
End Property
Public Property Get Peuples_ChateauEntreePositionX(ByVal Index As Integer) As Integer
    Peuples_ChateauEntreePositionX = Peuples(Index).ChateauEntreePositionX
End Property
Public Property Get Peuples_ChateauEntreePositionY(ByVal Index As Integer) As Integer
    Peuples_ChateauEntreePositionY = Peuples(Index).ChateauEntreePositionY
End Property

'Propriétées des compétences objets.
Public Property Get CompetenceObjet_SAVPerso(ByVal Index As Integer) As String
    CompetenceObjet_SAVPerso = CompetencesObjets(Index).SAVPerso
End Property
Public Property Get CompetenceObjet_NoListe(ByVal Index As Integer) As Integer
    CompetenceObjet_NoListe = CompetencesObjets(Index).NoListe
End Property
Public Property Get CompetenceObjet_Nom(ByVal Index As Integer) As String
    CompetenceObjet_Nom = CompetencesObjets(Index).Nom
End Property
Public Property Get CompetenceObjet_FichierMod(ByVal Index As Integer) As Boolean
    CompetenceObjet_FichierMod = CompetencesObjets(Index).FichierMod
End Property

'Propriétés des objets.
Public Property Get Objet_Nom(ByVal Index As Integer) As String
    Objet_Nom = Objets(Index).Nom
End Property
Public Property Get Objet_NomEffet(ByVal Index As Integer, Optional ByVal Feminin As Boolean) As String
    If Feminin Then
        Objet_NomEffet = Objets(Index).NomEffetF
    Else
        Objet_NomEffet = Objets(Index).NomEffetM
    End If
End Property
Public Property Get Objet_Image(ByVal Index As Integer) As String
    Objet_Image = Objets(Index).Image
End Property
Public Property Get Objet_CheminImage(ByVal Index As Integer) As String
    Objet_CheminImage = Objets(Index).CheminImage
End Property
Public Property Get Objet_NomPluriel(ByVal Index As Integer) As String
    Objet_NomPluriel = Objets(Index).NomPluriel
End Property
Public Property Get Objet_Description(ByVal Index As Integer) As String
    Objet_Description = Objets(Index).Description
End Property
Public Property Get Objet_Secret(ByVal Index As Integer) As Boolean
    Objet_Secret = Objets(Index).Secret
End Property
Public Property Get Objet_Unique(ByVal Index As Integer) As Boolean
    If Index = -1 Then
        Objet_Unique = False
    Else
        Objet_Unique = Objets(Index).Unique
    End If
End Property
Public Property Get Objet_Indestructible(ByVal Index As Integer) As Boolean
    Objet_Indestructible = Objets(Index).Indestructible
End Property
Public Property Get Objet_Inseparable(ByVal Index As Integer) As Boolean
    Objet_Inseparable = Objets(Index).Inseparable
End Property
Public Property Get Objet_Superposable(ByVal Index As Integer) As Integer
    Objet_Superposable = Objets(Index).Superposable
End Property
Public Property Get Objet_Empereur(ByVal Index As Integer) As Boolean
    Objet_Empereur = Objets(Index).Empereur
End Property
Public Property Get Objet_Categorie(ByVal Index As Integer) As Integer
    Objet_Categorie = Objets(Index).Categorie
End Property
Public Property Get Objet_Competence(ByVal Index As Integer) As Integer
    Objet_Competence = Objets(Index).Competence
End Property
Public Property Get Objet_PrixRessources(ByVal Index As Integer, ByVal Ressource As Integer) As Integer
    Objet_PrixRessources = Objets(Index).PrixRessources(Ressource)
End Property
Public Property Get Objet_Epoque(ByVal Index As Integer) As Integer
    Objet_Epoque = Objets(Index).Epoque
End Property
Public Property Get Objet_DureeFabrication(ByVal Index As Integer) As Integer
    Objet_DureeFabrication = Objets(Index).DureeFabrication
End Property
Public Property Get Objet_CoefXPFabrication(ByVal Index As Integer) As Single
    Objet_CoefXPFabrication = Objets(Index).CoefXPFabrication
End Property
Public Property Get Objet_PrixVente(ByVal Index As Integer) As Integer
    Objet_PrixVente = Objets(Index).PrixVente
End Property
Public Property Get Objet_BonusVie(ByVal Index As Integer) As Integer
    Objet_BonusVie = Objets(Index).BonusVie
End Property
Public Property Get Objet_BonusEnergie(ByVal Index As Integer) As Integer
    Objet_BonusEnergie = Objets(Index).BonusEnergie
End Property
Public Property Get Objet_BonusMagie(ByVal Index As Integer) As Integer
    Objet_BonusMagie = Objets(Index).BonusMagie
End Property
Public Property Get Objet_BonusMoral(ByVal Index As Integer) As Integer
    Objet_BonusMoral = Objets(Index).BonusMoral
End Property
Public Property Get Objet_BonusAttaque(ByVal Index As Integer) As Integer
    Objet_BonusAttaque = Objets(Index).BonusAttaque
End Property
Public Property Get Objet_BonusPorteeAttaque(ByVal Index As Integer) As Integer
    Objet_BonusPorteeAttaque = Objets(Index).BonusPorteeAttaque
End Property
Public Property Get Objet_LanceProjectile(ByVal Index As Integer) As Integer
    Objet_LanceProjectile = Objets(Index).LanceProjectile
End Property
Public Property Get Objet_ProjectileSupplementaire(ByVal Index As Integer) As Single
    Objet_ProjectileSupplementaire = Objets(Index).ProjectileSupplementaire
End Property
Public Property Get Objet_EffetAttaque(ByVal Index As Integer) As Integer
    Objet_EffetAttaque = Objets(Index).EffetAttaque
End Property
Public Property Get Objet_AttaqueResistance(ByVal Index As Integer) As Integer
    Objet_AttaqueResistance = Objets(Index).AttaqueResistance
End Property
Public Property Get Objet_SurfaceAttaque(ByVal Index As Integer) As Integer
    Objet_SurfaceAttaque = Objets(Index).SurfaceAttaque
End Property
Public Property Get Objet_AttaqueAutomatique(ByVal Index As Integer) As Boolean
    Objet_AttaqueAutomatique = Objets(Index).AttaqueAutomatique
End Property
Public Property Get Objet_AttaqueCauseEffet(ByVal Index As Integer) As Integer
    Objet_AttaqueCauseEffet = Objets(Index).AttaqueCauseEffet
End Property
Public Property Get Objet_BonusDefense(ByVal Index As Integer) As Integer
    Objet_BonusDefense = Objets(Index).BonusDefense
End Property
Public Property Get Objet_BonusDeplacement(ByVal Index As Integer) As Single
    Objet_BonusDeplacement = Objets(Index).BonusDeplacement
End Property
Public Property Get Objet_BonusVitesse(ByVal Index As Integer) As Single
    Objet_BonusVitesse = Objets(Index).BonusVitesse
End Property
Public Property Get Objet_BonusVitalite(ByVal Index As Integer) As Single
    Objet_BonusVitalite = Objets(Index).BonusVitalite
End Property
Public Property Get Objet_BonusCompetencesSpeciales(ByVal Index As Integer, ByVal Ressource As Integer) As Integer
    Objet_BonusCompetencesSpeciales = Objets(Index).BonusCompetencesSpeciales(Ressource)
End Property
Public Property Get Objet_BonusCompetencesRessources(ByVal Index As Integer, ByVal Ressource As Integer) As Integer
    Objet_BonusCompetencesRessources = Objets(Index).BonusCompetencesRessources(Ressource)
End Property
Public Property Get Objet_BonusCompetencesServices(ByVal Index As Integer, ByVal Ressource As Integer) As Integer
    Objet_BonusCompetencesServices = Objets(Index).BonusCompetencesServices(Ressource)
End Property
Public Property Get Objet_BonusCompetencesObjets(ByVal Index As Integer, ByVal Ressource As Integer) As Integer
    Objet_BonusCompetencesObjets = Objets(Index).BonusCompetencesObjets(Ressource)
End Property
Public Property Get Objet_BonusArmure(ByVal Index As Integer) As Integer
    Objet_BonusArmure = Objets(Index).BonusArmure
End Property
Public Property Get Objet_BonusMaxRessources(ByVal Index As Integer) As Integer
    Objet_BonusMaxRessources = Objets(Index).BonusMaxRessources
End Property
Public Property Get Objet_BonusResistance(ByVal Index As Integer, ByVal IndiceResistance As Integer) As Single
    Objet_BonusResistance = Objets(Index).BonusResistances(IndiceResistance)
End Property
Public Property Get Objet_BonusVitesseRegenerationVie(ByVal Index As Integer) As Single
    Objet_BonusVitesseRegenerationVie = Objets(Index).BonusVitesseRegenerationVie
End Property
Public Property Get Objet_BonusVitesseRegenerationEnergie(ByVal Index As Integer) As Single
    Objet_BonusVitesseRegenerationEnergie = Objets(Index).BonusVitesseRegenerationEnergie
End Property
Public Property Get Objet_BonusVitesseRegenerationMagie(ByVal Index As Integer) As Single
    Objet_BonusVitesseRegenerationMagie = Objets(Index).BonusVitesseRegenerationMagie
End Property
Public Property Get Objet_BonusVitesseRegenerationMoral(ByVal Index As Integer) As Single
    Objet_BonusVitesseRegenerationMoral = Objets(Index).BonusVitesseRegenerationMoral
End Property
Public Property Get Objet_PerteVie(ByVal Index As Integer) As Double
    Objet_PerteVie = Objets(Index).PerteVie
End Property
Public Property Get Objet_PerteEnergie(ByVal Index As Integer) As Double
    Objet_PerteEnergie = Objets(Index).PerteEnergie
End Property
Public Property Get Objet_PerteMagie(ByVal Index As Integer) As Double
    Objet_PerteMagie = Objets(Index).PerteMagie
End Property
Public Property Get Objet_PerteMoral(ByVal Index As Integer) As Double
    Objet_PerteMoral = Objets(Index).PerteMoral
End Property
Public Property Get Objet_PerteArgent(ByVal Index As Integer) As Long
    Objet_PerteArgent = Objets(Index).PerteArgent
End Property
Public Property Get Objet_PerteRessourceIndice(ByVal Index As Integer) As Integer
    Objet_PerteRessourceIndice = Objets(Index).PerteRessourceIndice
End Property
Public Property Get Objet_PerteRessourceQuantite(ByVal Index As Integer) As Single
    Objet_PerteRessourceQuantite = Objets(Index).PerteRessourceQuantite
End Property
Public Property Get Objet_Poison(ByVal Index As Integer) As Single
    Objet_Poison = Objets(Index).Poison
End Property
Public Property Get Objet_ApparencePerso(ByVal Index As Integer) As Integer
    Objet_ApparencePerso = Objets(Index).ApparencePerso
End Property
Public Property Get Objet_ChangerApparence(ByVal Index As Integer) As Boolean
    If Objets(Index).FichierApparenceM <> "" Then
        Objet_ChangerApparence = True
    End If
End Property
Public Property Get Objet_FichierApparence(ByVal Index As Integer, _
                                           ByVal Feminin As Boolean) As String
    If Feminin Then
        Objet_FichierApparence = Objets(Index).FichierApparenceF
    Else
        Objet_FichierApparence = Objets(Index).FichierApparenceM
    End If
End Property
Public Property Get Objet_NumeroApparence(ByVal Index As Integer, _
                                          ByVal Feminin As Boolean) As Integer
    If Feminin Then
        Objet_NumeroApparence = Objets(Index).NumeroApparenceF
    Else
        Objet_NumeroApparence = Objets(Index).NumeroApparenceM
    End If
End Property
Public Property Let Objet_NumeroApparence(ByVal Index As Integer, _
                                          ByVal Feminin As Boolean, _
                                          ByVal Valeur As Integer)
    If Feminin Then
        Objets(Index).NumeroApparenceF = Valeur
    Else
        Objets(Index).NumeroApparenceM = Valeur
    End If
End Property
Public Property Get Objet_ApparenceLargeur(ByVal Index As Integer) As Integer
    Objet_ApparenceLargeur = Objets(Index).ApparenceLargeur
End Property
Public Property Get Objet_ApparenceHauteur(ByVal Index As Integer) As Integer
    Objet_ApparenceHauteur = Objets(Index).ApparenceHauteur
End Property
Public Property Get Objet_ApparenceDecalageX(ByVal Index As Integer, ByVal Direction As Integer) As Integer
    Objet_ApparenceDecalageX = Objets(Index).ApparenceDecalageX(Direction)
End Property
Public Property Get Objet_ApparenceDecalageY(ByVal Index As Integer, ByVal Direction As Integer) As Integer
    Objet_ApparenceDecalageY = Objets(Index).ApparenceDecalageY(Direction)
End Property
Public Property Get Objet_ApparenceAvantPlan(ByVal Index As Integer, ByVal Direction As Integer) As Boolean
    Objet_ApparenceAvantPlan = Objets(Index).ApparenceAvantPlan(Direction)
End Property
Public Property Get Objet_ChangerApparencePerso(ByVal Index As Integer) As Boolean
    Objet_ChangerApparencePerso = Objets(Index).ChangerApparencePerso
End Property
Public Property Get Objet_Utilisable(ByVal Index As Integer) As Boolean
    Objet_Utilisable = Objets(Index).Utilisable
End Property
Public Property Get Objet_UtilisationUnique(ByVal Index As Integer) As Boolean
    Objet_UtilisationUnique = Objets(Index).UtilisationUnique
End Property
Public Property Get Objet_UtilisationMort(ByVal Index As Integer) As Boolean
    Objet_UtilisationMort = Objets(Index).UtilisationMort
End Property
Public Property Get Objet_PorteeUtilisation(ByVal Index As Integer) As Integer
    Objet_PorteeUtilisation = Objets(Index).PorteeUtilisation
End Property
Public Property Get Objet_Experience(ByVal Index As Integer) As Double
    Objet_Experience = Objets(Index).Experience
End Property
Public Property Get Objet_Invocation(ByVal Index As Integer) As Integer
    Objet_Invocation = Objets(Index).Invocation
End Property
Public Property Get Objet_InvocationIA_Fichier(ByVal Index As Integer) As String
    Objet_InvocationIA_Fichier = Objets(Index).InvocationIA_Fichier
End Property
Public Property Get Objet_Quantite(ByVal Index As Integer) As Single
    Objet_Quantite = Objets(Index).Quantite
End Property
Public Property Get Objet_QuantiteMax(ByVal Index As Integer) As Single
    Objet_QuantiteMax = Objets(Index).QuantiteMax
End Property
Public Property Get Objet_GainVie(ByVal Index As Integer) As Double
    Objet_GainVie = Objets(Index).GainVie
End Property
Public Property Get Objet_GainEnergie(ByVal Index As Integer) As Double
    Objet_GainEnergie = Objets(Index).GainEnergie
End Property
Public Property Get Objet_GainMagie(ByVal Index As Integer) As Double
    Objet_GainMagie = Objets(Index).GainMagie
End Property
Public Property Get Objet_GainMoral(ByVal Index As Integer) As Double
    Objet_GainMoral = Objets(Index).GainMoral
End Property
Public Property Get Objet_GainArgent(ByVal Index As Integer) As Single
    Objet_GainArgent = Objets(Index).GainArgent
End Property
Public Property Get Objet_Impot(ByVal Index As Integer) As Boolean
    Objet_Impot = Objets(Index).Impot
End Property
Public Property Get Objet_GainRessource(ByVal Index As Integer, ByVal Ressource As Integer) As Single
    Objet_GainRessource = Objets(Index).GainRessource(Ressource)
End Property
Public Property Get Objet_Resurrection(ByVal Index As Integer) As Long
    Objet_Resurrection = Objets(Index).Resurrection
End Property
Public Property Get Objet_Berserk(ByVal Index As Integer) As Integer
    Objet_Berserk = Objets(Index).Berserk
End Property
Public Property Get Objet_SelfResurrection(ByVal Index As Integer) As Long
    Objet_SelfResurrection = Objets(Index).SelfResurrection
End Property
Public Property Get Objet_CreerObjet(ByVal Index As Integer) As Integer
    Objet_CreerObjet = Objets(Index).CreerObjet
End Property
Public Property Get Objet_CreerObjetAllie(ByVal Index As Integer) As Integer
    Objet_CreerObjetAllie = Objets(Index).CreerObjetAllie
End Property
Public Property Get Objet_CreerObjetEnnemi(ByVal Index As Integer) As Integer
    Objet_CreerObjetEnnemi = Objets(Index).CreerObjetEnnemi
End Property
Public Property Get Objet_Duree(ByVal Index As Integer) As Long
    Objet_Duree = Objets(Index).Duree
End Property
Public Property Get Objet_DureeRecharge(ByVal Index As Integer) As Long
    Objet_DureeRecharge = Objets(Index).DureeRecharge
End Property
Public Property Get Objet_DureeRechargeAttaque(ByVal Index As Integer) As Long
    Objet_DureeRechargeAttaque = Objets(Index).DureeRechargeAttaque
End Property
Public Property Get Objet_RechargesMultiples(ByVal Index As Integer) As Single
    Objet_RechargesMultiples = Objets(Index).RechargesMultiples
End Property
Public Property Get Objet_EffetAnnuleRestaureObjet(ByVal Index As Integer) As Boolean
    Objet_EffetAnnuleRestaureObjet = Objets(Index).EffetAnnuleRestaureObjet
End Property
Public Property Get Objet_EffetAnnulationImpossible(ByVal Index As Integer) As Boolean
    Objet_EffetAnnulationImpossible = Objets(Index).EffetAnnulationImpossible
End Property
Public Property Get Objet_EffetAnnulationMort(ByVal Index As Integer) As Boolean
    Objet_EffetAnnulationMort = Objets(Index).EffetAnnulationMort
End Property
Public Property Get Objet_UtilisationAutomatique(ByVal Index As Integer) As Boolean
    Objet_UtilisationAutomatique = Objets(Index).UtilisationAutomatique
End Property
Public Property Get Objet_Monture(ByVal Index As Integer) As Boolean
    Objet_Monture = Objets(Index).Monture
End Property
Public Property Get Objet_Volant(ByVal Index As Integer) As Boolean
    Objet_Volant = Objets(Index).Volant
End Property
Public Property Get Objet_PrixTalents(ByVal Index As Integer) As Long
    Objet_PrixTalents = Objets(Index).PrixTalents
End Property

'Propriétés des trésors.
Public Property Get Tresor_Nom(ByVal Index As Integer) As String
    Tresor_Nom = Tresors(Index).Nom
End Property
Public Property Get Tresor_ValeurMin(ByVal Index As Integer) As Integer
    Tresor_ValeurMin = Tresors(Index).ValeurMin
End Property

'Propriétés des décors.
Public Property Get Decor_Nom(ByVal Index As Integer) As String
    Decor_Nom = Decors(Index).Nom
End Property
Public Property Get Decor_Image(ByVal Index As Integer) As String
    Decor_Image = Decors(Index).Image
End Property
Public Property Get Decor_Largeur(ByVal Index As Integer) As Integer
    Decor_Largeur = Decors(Index).Largeur
End Property
Public Property Get Decor_Hauteur(ByVal Index As Integer) As Integer
    Decor_Hauteur = Decors(Index).Hauteur
End Property
Public Property Get Decor_NombreEtapeAnimation(ByVal Index As Integer) As Integer
    Decor_NombreEtapeAnimation = Decors(Index).NombreEtapeAnimation
End Property
Public Property Get Decor_Duree(ByVal Index As Integer) As Integer
    Decor_Duree = Decors(Index).Duree
End Property

'Propriétés des compétences de service.
Public Property Get Service_Nom(ByVal Index As Integer) As String
    Service_Nom = Services(Index).Nom
End Property
Public Property Get Service_Description(ByVal Index As Integer) As String
    Service_Description = Services(Index).Description
End Property
Public Property Get Service_VitesseCreation(ByVal Index As Integer) As String
    Service_VitesseCreation = Services(Index).VitesseCreation
End Property
Public Property Get Service_MatierePremiere(ByVal Index As Integer) As String
    Service_MatierePremiere = Services(Index).TypeMatierePremiere
End Property
Public Property Get Service_QuantiteMatierePremiere(ByVal Index As Integer) As String
    Service_QuantiteMatierePremiere = Services(Index).QuantiteMatierePremiere
End Property
Public Property Get Service_Prix(ByVal Index As Integer) As Double
    Service_Prix = Services(Index).Prix
End Property
Public Property Get Service_Coefficient(ByVal Index As Integer) As Single
    Service_Coefficient = Services(Index).Coefficient
End Property
Public Property Get Service_NomCompetence(ByVal Index As Integer) As String
    Service_NomCompetence = Services(Index).NomCompetence
End Property
Public Property Get Service_CoefExp(ByVal Index As Integer) As String
    Service_CoefExp = Services(Index).CoefExp
End Property
Public Property Get Service_Effet(ByVal Index As Integer) As String
    Service_Effet = Services(Index).Effet
End Property

'Propriétés des effets.
Public Property Get Effet_Nom(ByVal Index As Integer) As String
    Effet_Nom = Effets(Index).Nom
End Property
Public Property Get Effet_Fichier(ByVal Index As Integer) As String
    Effet_Fichier = Effets(Index).Fichier
End Property
Public Property Get Effet_Largeur(ByVal Index As Integer) As Integer
    Effet_Largeur = Effets(Index).Largeur
End Property
Public Property Get Effet_Hauteur(ByVal Index As Integer) As Integer
    Effet_Hauteur = Effets(Index).Hauteur
End Property
Public Property Get Effet_NombreEtapes(ByVal Index As Integer) As Integer
    Effet_NombreEtapes = Effets(Index).NombreEtapes
End Property
Public Property Get Effet_DureeEtape(ByVal Index As Integer) As Double
    Effet_DureeEtape = Effets(Index).DureeEtape
End Property
Public Property Get Effet_Son(ByVal Index As Integer) As String
    Effet_Son = Effets(Index).Son
End Property

'Propriétés des races.
Public Property Get Race_Indice(ByVal Nom As String) As Integer
    'Renvoie l'indice de la race en fonction du nom entré en paramètre.
    Dim i As Integer
    For i = 0 To NombreRaces - 1
        If Nom = Race_Nom(i) Then
            Race_Indice = i
            i = Parametres.NombreRaces
        End If
    Next i
End Property
Public Property Get Race_Nom(ByVal Index As Integer, _
                             Optional ByVal Feminin As Boolean, _
                             Optional ByVal Pluriel As Boolean) As String
    If Index >= 0 Then
        If Feminin Then
            If Pluriel Then
                Race_Nom = Races(Index).NomFPluriel
            Else
                Race_Nom = Races(Index).NomF
            End If
        Else
            If Pluriel Then
                Race_Nom = Races(Index).NomMPluriel
            Else
                Race_Nom = Races(Index).NomM
            End If
        End If
    End If
End Property
Public Property Get Race_NomAleatoire(ByVal Index As Integer) As Boolean
    Race_NomAleatoire = Races(Index).NomAleatoire
End Property
Public Property Get Race_Jouable(ByVal Index As Integer) As Boolean
    Race_Jouable = Races(Index).Jouable
End Property
Public Property Get Race_Description(ByVal Index As Integer) As String
    Race_Description = Races(Index).Description
End Property
Public Property Get Race_Largeur(ByVal Index As Integer) As Integer
    Race_Largeur = Races(Index).Largeur
End Property
Public Property Get Race_Hauteur(ByVal Index As Integer) As Integer
    Race_Hauteur = Races(Index).Hauteur
End Property
Public Property Get Race_Portrait(ByVal Index As Integer) As String
    Race_Portrait = Races(Index).Portrait
End Property
Public Property Get Race_RestrictionSexe(ByVal Index As Integer, Optional ByVal Feminin As Boolean) As Boolean
    If Feminin Then
        Race_RestrictionSexe = Races(Index).RestrictionFeminin
    Else
        Race_RestrictionSexe = Races(Index).RestrictionMaculin
    End If
End Property
Public Property Get Race_DefinirNombreCategoriesApparences(ByVal Index As Integer) As Integer
    Race_DefinirNombreCategoriesApparences = Races(Index).NombreCategoriesApparences
End Property
Public Property Get Race_DefinirDossiersCategories(ByVal Index As Integer, _
                                                   ByVal Categorie As Integer, _
                                                   Optional ByVal Feminin As Boolean) As String
    If Feminin Then
        Race_DefinirDossiersCategories = Races(Index).DossiersCategoriesF(Categorie)
    Else
        Race_DefinirDossiersCategories = Races(Index).DossiersCategoriesM(Categorie)
    End If
End Property
Public Property Get Race_DefinirApparence(ByVal Index As Integer) As Boolean
    Race_DefinirApparence = Races(Index).DefinirApparence
End Property
Public Property Get Race_CheminApparence(ByVal Index As Integer, Optional ByVal Feminin As Boolean) As String
    If Feminin Then
        Race_CheminApparence = Races(Index).CheminApparenceF
    Else
        Race_CheminApparence = Races(Index).CheminApparenceM
    End If
End Property
Public Property Get Race_CheminAbsolu(ByVal Index As Integer) As Boolean
    Race_CheminAbsolu = Races(Index).CheminAbsolu
End Property
Public Property Get Race_ApparenceVerouillee(ByVal Index As Integer) As Boolean
    Race_ApparenceVerouillee = Races(Index).ApparenceVerouillee
End Property
Public Property Get Race_Vie(ByVal Index As Integer) As Long
    Race_Vie = Races(Index).Vie
End Property
Public Property Get Race_Energie(ByVal Index As Integer) As Long
    Race_Energie = Races(Index).Energie
End Property
Public Property Get Race_Magie(ByVal Index As Integer) As Long
    Race_Magie = Races(Index).Magie
End Property
Public Property Get Race_Moral(ByVal Index As Integer) As Long
    Race_Moral = Races(Index).Moral
End Property
Public Property Get Race_Attaque(ByVal Index As Integer) As Long
    Race_Attaque = Races(Index).Attaque
End Property
Public Property Get Race_Defense(ByVal Index As Integer) As Long
    Race_Defense = Races(Index).Defense
End Property
Public Property Get Race_CoefVie(ByVal Index As Integer) As Single
    Race_CoefVie = Races(Index).CoefVie
End Property
Public Property Get Race_CoefEnergie(ByVal Index As Integer) As Single
    Race_CoefEnergie = Races(Index).CoefEnergie
End Property
Public Property Get Race_CoefMagie(ByVal Index As Integer) As Single
    Race_CoefMagie = Races(Index).CoefMagie
End Property
Public Property Get Race_CoefMoral(ByVal Index As Integer) As Single
    Race_CoefMoral = Races(Index).CoefMoral
End Property
Public Property Get Race_CoefAttaque(ByVal Index As Integer) As Single
    Race_CoefAttaque = Races(Index).CoefAttaque
End Property
Public Property Get Race_CoefDefense(ByVal Index As Integer) As Single
    Race_CoefDefense = Races(Index).CoefDefense
End Property
Public Property Get Race_VitesseRegenerationVie(ByVal Index As Integer) As Single
    Race_VitesseRegenerationVie = Races(Index).VitesseRegenerationVie
End Property
Public Property Get Race_VitesseRegenerationEnergie(ByVal Index As Integer) As Single
    Race_VitesseRegenerationEnergie = Races(Index).VitesseRegenerationEnergie
End Property
Public Property Get Race_VitesseRegenerationMagie(ByVal Index As Integer) As Single
    Race_VitesseRegenerationMagie = Races(Index).VitesseRegenerationMagie
End Property
Public Property Get Race_VitesseRegenerationMoral(ByVal Index As Integer) As Single
    Race_VitesseRegenerationMoral = Races(Index).VitesseRegenerationMoral
End Property
Public Property Get Race_CoefCompSpeciales(ByVal Index As Integer, ByVal Competence As Integer) As Single
    Race_CoefCompSpeciales = Races(Index).CoefCompSpeciales(Competence)
End Property
Public Property Get Race_CoefCompRessources(ByVal Index As Integer, ByVal Competence As Integer) As Single
    Race_CoefCompRessources = Races(Index).CoefCompRessources(Competence)
End Property
Public Property Get Race_CoefCompServices(ByVal Index As Integer, ByVal Competence As Integer) As Single
    Race_CoefCompServices = Races(Index).CoefCompServices(Competence)
End Property
Public Property Get Race_CoefCompObjets(ByVal Index As Integer, ByVal Competence As Integer) As Single
    Race_CoefCompObjets = Races(Index).CoefCompObjets(Competence)
End Property
Public Property Get Race_Equiper(ByVal Index As Integer) As Integer
    Race_Equiper = Races(Index).Equiper
End Property
Public Property Get Race_Parole(ByVal Index As Integer) As Boolean
    Race_Parole = Races(Index).Parole
End Property
Public Property Get Race_Experience(ByVal Index As Integer) As Boolean
    Race_Experience = Races(Index).Experience
End Property
Public Property Get Race_Volant(ByVal Index As Integer) As Boolean
    Race_Volant = Races(Index).Volant
End Property
Public Property Get Race_CommentairesSpecifique(ByVal Index As Integer) As Integer
    Race_CommentairesSpecifique = Races(Index).CommentairesSpecifique
End Property
Public Property Get Race_BonusMaxRessources(ByVal Index As Integer) As Integer
    Race_BonusMaxRessources = Races(Index).BonusMaxRessources
End Property
Public Property Get Race_PorteeAttaque(ByVal Index As Integer) As Integer
    Race_PorteeAttaque = Races(Index).PorteeAttaque
End Property
Public Property Get Race_EffetAttaque(ByVal Index As Integer) As Integer
    Race_EffetAttaque = Races(Index).EffetAttaque
End Property
Public Property Get Race_AttaqueCauseEffet(ByVal Index As Integer) As Integer
    Race_AttaqueCauseEffet = Races(Index).AttaqueCauseEffet
End Property
Public Property Get Race_AttaqueResistance(ByVal Index As Integer) As Integer
    Race_AttaqueResistance = Races(Index).AttaqueResistance
End Property
Public Property Get Race_Armure(ByVal Index As Integer) As Integer
    Race_Armure = Races(Index).Armure
End Property
Public Property Get Race_Vitesse(ByVal Index As Integer) As Integer
    Race_Vitesse = Races(Index).Vitesse
End Property
Public Property Get Race_Acceleration(ByVal Index As Integer) As Integer
    Race_Acceleration = Races(Index).Acceleration
End Property
'Public Property Get Race_Projectile(ByVal Index As Integer) As Boolean
'    Race_Projectile = Races(Index).Projectile
'End Property
Public Property Get Race_Pietinement(ByVal Index As Integer) As Integer
    Race_Pietinement = Races(Index).Pietinement
End Property
Public Property Get Race_Inattaquable(ByVal Index As Integer) As Boolean
    Race_Inattaquable = Races(Index).Inattaquable
End Property
Public Property Get Race_AutoResurrection(ByVal Index As Integer) As Boolean
    Race_AutoResurrection = Races(Index).AutoResurrection
End Property
Public Property Get Race_IA_Pacifique(ByVal Index As Integer) As Boolean
    Race_IA_Pacifique = Races(Index).IA_Pacifique
End Property
Public Property Get Race_IA_Manger(ByVal Index As Integer) As Boolean
    Race_IA_Manger = Races(Index).IA_Manger
End Property
Public Property Get Race_IA_Temeraire(ByVal Index As Integer) As Boolean
    Race_IA_Temeraire = Races(Index).IA_Temeraire
End Property
Public Property Get Race_IA_Protecteur(ByVal Index As Integer) As Boolean
    Race_IA_Protecteur = Races(Index).IA_Protecteur
End Property
Public Property Get Race_IA_Egoiste(ByVal Index As Integer) As Boolean
    Race_IA_Egoiste = Races(Index).IA_Egoiste
End Property
Public Property Get Race_IA_Berserk(ByVal Index As Integer) As Integer
    Race_IA_Berserk = Races(Index).IA_Berserk
End Property
Public Property Get Race_IA_Incommandable(ByVal Index As Integer) As Boolean
    Race_IA_Incommandable = Races(Index).IA_Incommandable
End Property
Public Property Get Race_IA_Selectionner(ByVal Index As Integer) As Boolean
    Race_IA_Selectionner = Races(Index).IA_Selectionner
End Property
Public Property Get Race_IA_Fichier(ByVal Index As Integer) As String
    Race_IA_Fichier = Races(Index).IA_Fichier
End Property
Public Property Get Race_Infatigable(ByVal Index As Integer) As Boolean
    Race_Infatigable = Races(Index).Infatigable
End Property
Public Property Get Race_Kamikaze(ByVal Index As Integer) As Boolean
    Race_Kamikaze = Races(Index).Kamikaze
End Property
Public Property Get Race_RestaurationComplete(ByVal Index As Integer) As Boolean
    Race_RestaurationComplete = Races(Index).RestaurationComplete
End Property
Public Property Get Race_Argent(ByVal Index As Long) As Double
    Race_Argent = Races(Index).Argent
End Property
Public Property Get Race_Ressources(ByVal Index As Long, ByVal IndiceRessource As Long) As Double
    Race_Ressources = Races(Index).Ressources(IndiceRessource)
End Property
Public Property Get Race_Resistances(ByVal Index As Long, ByVal IndiceResistance As Long) As Single
    Race_Resistances = Races(Index).Resistances(IndiceResistance)
End Property
Public Property Get Race_DepartSansRessources(ByVal Index As Integer) As Boolean
    Race_DepartSansRessources = Races(Index).DepartSansRessources
End Property
Public Property Get Race_NombreObjets(ByVal Index As Long) As Integer
    Race_NombreObjets = Races(Index).NombreObjets
End Property
Public Property Get Race_Objet(ByVal Index As Long, ByVal Objet As Long) As Integer
    Race_Objet = Races(Index).Objet(Objet)
End Property
Public Property Get Race_NombreObjetsRace(ByVal Index As Long) As Integer
    Race_NombreObjetsRace = Races(Index).ObjetRace.Count
End Property
Public Property Get Race_ObjetRace(ByVal Index As Long, ByVal Objet As Long) As Integer
    Race_ObjetRace = Races(Index).ObjetRace(Objet)
End Property
Public Property Get Race_NombreEquipement(ByVal Index As Long) As Integer
    Race_NombreEquipement = Races(Index).NombreEquipement
End Property
Public Property Get Race_Equipement(ByVal Index As Long, ByVal Objet As Long) As Integer
    Race_Equipement = Races(Index).Equipement(Objet)
End Property
Public Property Get Race_IndiceBestiaire(ByVal Index As Integer) As Integer
    Race_IndiceBestiaire = Races(Index).IndiceBestiaire
End Property
Public Property Get Race_UtilisationMonture(ByVal Index As Integer) As Boolean
    Race_UtilisationMonture = Races(Index).UtilisationMonture
End Property
Public Property Get Race_Tombe(ByVal Index As Integer) As Integer
    Race_Tombe = Races(Index).Tombe
End Property
Public Property Get Race_EffetMort(ByVal Index As Integer) As Integer
    Race_EffetMort = Races(Index).EffetMort
End Property
Public Property Get Race_FichierMod(ByVal Index As Integer) As Boolean
    Race_FichierMod = Races(Index).FichierMod
End Property

Public Property Get Bestiaire_Ordre_Alpha(ByVal Index As Integer) As Integer
    Bestiaire_Ordre_Alpha = BestiaireOrdreAlpha(Index)
End Property

'Propriétés des résistances.
Public Property Get Resistance_Nom(ByVal Index As Integer) As String
    Resistance_Nom = Resistances(Index).Nom
End Property
Public Property Get Resistance_NomComplet(ByVal Index As Integer) As String
    Resistance_NomComplet = Resistances(Index).NomComplet
End Property
Public Property Get Resistance_Attaque(ByVal Index As Integer) As String
    Resistance_Attaque = Resistances(Index).Attaque
End Property
Public Property Get Resistance_Abreviation(ByVal Index As Integer) As String
    Resistance_Abreviation = Resistances(Index).Abreviation
End Property
Public Property Get Statut_Nom(ByVal Index As Integer) As String
    Statut_Nom = Statuts(Index).Nom
End Property
Public Property Get Statut_Resistance(ByVal Index As Integer, ByVal Resistance As Integer) As Single
    Statut_Resistance = Statuts(Index).Resistances(Resistance)
End Property
