VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsJeuPerso"
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

Const IntervalesActions = 10
Const VitesseAction = 400  '20 '400 normal '20 rapide.
Const VitesseConstruction = 3 / (VitesseAction / 400) '75
Const VitesseFabricationObjet = 0.1 / (VitesseAction / 400) '2

Const VitesseFatigueDeplacement = 0.001 / (VitesseAction / 400) * IntervalesActions
Const VitesseDemoralisationDeplacement = 0 / (VitesseAction / 400) * IntervalesActions

Const VitesseFatigueTravail = 10 / (VitesseAction * 4) / 20 * IntervalesActions * 0.8
Const VitesseDemoralisationTravail = 2 / (VitesseAction * 4) / 20 * IntervalesActions * 0.8

Const VitesseFatigueConstruction = 10 / (VitesseAction * 4) / 20 * IntervalesActions * 0.8
Const VitesseDemoralisationConstruction = 2 / (VitesseAction * 4) / 20 * IntervalesActions * 0.8

Const VitesseFatigueFabrication = 10 / (VitesseAction * 4) / 20 * IntervalesActions * 0.8
Const VitesseDemoralisationFabrication = 2 / (VitesseAction * 4) / 20 * IntervalesActions * 0.8

Const VitesseFatigueCombat = 0.0125 / (VitesseAction / 400) * IntervalesActions
Const VitesseDemoralisationCombat = 0.0025 / (VitesseAction / 400) * IntervalesActions

Const VitesseGenerationRessources = 1 / (VitesseAction / 400)

Const VitesseRecuperationVie = 0.075 / 20 / (VitesseAction / 400) * IntervalesActions '0.05 / 20
Const VitesseRecuperationEnergie = 0.1 / 20 / (VitesseAction / 400) * IntervalesActions
Const VitesseRecuperationMagie = 0.2 / 20 / 2 / (VitesseAction / 400) * IntervalesActions
Const VitesseRecuperationMoral = 0.05 / 20 / (VitesseAction / 400) * IntervalesActions
Const VitesseRecuperationAttaque = 1.25
Const VitesseRecuperationDefense = 0.75

Public VitesseRegenerationVie As Single
Public VitesseRegenerationEnergie As Single
Public VitesseRegenerationMagie As Single
Public VitesseRegenerationMoral As Single
Public BonusVitesseRegenerationVie As Single
Public BonusVitesseRegenerationEnergie As Single
Public BonusVitesseRegenerationMagie As Single
Public BonusVitesseRegenerationMoral As Single
Private RestaurationComplete As Boolean
Const VitesseRegenerationVieEnAttente = 0.001 / 20 / (VitesseAction / 400) * IntervalesActions
Const VitesseRegenerationEnergieEnAttente = 0.001 / 20 / (VitesseAction / 400) * IntervalesActions
Const VitesseRegenerationMagieEnAttente = 0 '0.001 / 20 / (VitesseAction / 400) * IntervalesActions
Const VitesseRegenerationMoralEnAttente = 0 / 20 / (VitesseAction / 400) * IntervalesActions
Const VitesseRegenerationVieEnRepos = 0.002 / 20 / (VitesseAction / 400) * IntervalesActions
Const VitesseRegenerationEnergieEnRepos = 0.002 / 20 / (VitesseAction / 400) * IntervalesActions
Const VitesseRegenerationMagieEnRepos = 0 '0.002 / 20 / (VitesseAction / 400) * IntervalesActions
Const VitesseRegenerationMoralEnRepos = 0.0005 / 20 / (VitesseAction / 400) * IntervalesActions

Const VitesseDeplacementFantome = 6
Const VitesseResurrectionDepart = 0.25 * IntervalesActions

Const NiveauCoefCarac = 0.25
Const NiveauMultiplicateur = 1

Const CoefExpVie = 20 'Multiplicateur d'expérience nécessaire pour changer de niveau de la vie.
Const CoefExpEnergie = 20 'Multiplicateur d'expérience nécessaire pour changer de niveau de l'énergie.
Const CoefExpMagie = 20 'Multiplicateur d'expérience nécessaire pour changer de niveau de la magie.
Const CoefExpMoral = 20 'Multiplicateur d'expérience nécessaire pour changer de niveau du moral.
Const CoefExpAttaque = 400 'Multiplicateur d'expérience nécessaire pour changer de niveau de l'attaque.
Const CoefExpDefense = 250 'Multiplicateur d'expérience nécessaire pour changer de niveau de la défense.
Const CoefExpCompSpeciales = 20
Const CoefExpCompConstruction = 0.00075 '0.001
Const CoefExpCompConstructionObjet = 0.05
Const CoefExpCompDestruction = 0.0025
Const CoefExpCompDestructionObjet = 1
Const CoefExpCompCoupCritique = 0.00125 * 4 * 4
Const CoefExpCompCoupCritiqueObjet = 1 '0.05
Const CoefExpCompMarchandage = 0.02 / 20
Const CoefExpCompMarchandageObjet = 1
Const CoefExpCompRessources = 15
Const CoefExpCompObjets = 20
Const CoefExpCompObjetsFabrication = 0.0025
Const CoefExpCompObjetsVie = 0.2 'Multiplicateur d'expérience supplémentaire pour les objets donnant un bonus de vie.
Const CoefExpCompObjetsEnergie = 2
Const CoefExpCompObjetsMagie = 0.2
Const CoefExpCompObjetsMoral = 10
Const CoefExpCompObjetsAttaque = 0.1 'Multiplicateur d'expérience supplémentaire pour les objets de combat.
Const CoefExpCompObjetsDefense = 0.16 'Multiplicateur d'expérience supplémentaire pour les objets de combat.
Const CoefExpCompObjetsResistance = 0.2 'Multiplicateur d'expérience supplémentaire pour les objets offrant des résistances supplémentaires.
Const CoefExpCompObjetsBonus = 1 'Multiplicateur de bonus supplémentaire pour l'arme en fonction de la compétence.
Const CoefExpCompPorteeBonus = 0.25 'Multiplicateur de bonus supplémentaire pour la portee des armes en fonction de la compétence.
Const CoefExpCompSurfaceBonus = 0.25 'Multiplicateur de bonus supplémentaire pour la surface de dégâts des armes en fonction de la compétence.
Const CoefExpCompArmureBonus = 0 'Multiplicateur de bonus supplémentaire pour l'armure.
Const CoefExpCompServices = 20

Const NombreAttitudes = 5
Const NombreFormations = 5

'Position actuelle du personnage.
Public PositionX As Long
Public PositionY As Long
Public LongueurPas As Long
Public LongueurPasMax As Long
Public Acceleration As Integer
Public BonusDeplacement As Long 'Bonus accordé normalement.
Private BonusVitesse As Long 'Bonus accordés magiquement.
Public NumeroPas As Integer
Public Apparence As Integer

'Position où doit se rendre le personnage.
Public DirectionX As Long
Public DirectionY As Long
'Position où le personnage tire en se déplaçant.
Public CibleX As Long
Public CibleY As Long

'Caractéristiques :
Public Protege As Boolean 'Si vrai, la clé de protection pour ce personnage est active.
Public Numero As Long
Public Nom As String
Public Largeur As Integer
Public Hauteur As Integer
Public Joueur As Long
Public NumeroFief As Long
Public NumeroEquipe As Long
Public NumeroChef As Long 'Est supérieur ou égal à 0 s'il s'agit d'une invocation.
Public Incommandable As Boolean
Public ChefDecalageX As Long 'Position relative par rapport au chef.
Public ChefDecalageY As Long
Public ChefPositionX As Long 'Position où le chef veut que l'unité soit placé.
Public ChefPositionY As Long
Public Attitude As Integer
Public Formation As Integer
Public Empereur As Boolean
Public NombreSoldats As Long
Public Race As Integer
Public Parole As Boolean
Public Volant As Boolean
Private BonusVolant As Boolean
Public ChoisirApparence As Boolean
Public FichierApparence As String
Public ApparenceChoisie As Integer 'Mémorise l'apparence d'origine.
Private FichierApparenceBase() As String
Private ApparenceBase() As Integer
Private ApparenceBaseAfficher() As Boolean
Public ApparenceObjetAfficher As Boolean
Private ApparenceObjet As Collection
Private ApparenceLargeur As Integer 'Largeur de la monture.
Private ApparenceHauteur As Integer 'Hauteur de la monture.
Public ApparenceDecalageXTemp As Integer 'Décale l'apparence de la monture.
Public ApparenceDecalageYTemp As Integer 'Décale l'apparence de la monture.
Private ApparenceDecalageX() As Integer 'Décale l'apparence du personnage horizontalement.
Private ApparenceDecalageY() As Integer 'Décale l'apparence du personnage verticalement.
Private ApparenceAvantPlan() As Boolean 'Place l'apparence en question en arrière plan selon la direction du personnage comme indice.
Private ApparenceVerouillee As Boolean 'Si vrai, l'équipement ne change pas l'apparence du personnage. (sauf pour les montures)
Private NombreApparenceBase As Integer
Public Titre As Integer
Public Feminin As Boolean
Public GainExperience As Boolean 'Si vrai, le personnage gagne de l'expérience.
Public DureeRessurection As Double 'Temps restant pour ressusciter.
Public ResurrectionNombre As Integer 'Nombre de fois que le personnage peut ressusciter au château. "-1" pour illimité.
Public VitesseResurrection As Single
Private Commentaires As String
Public CommentairesSpecifique As Boolean
Public CommentairesSpecifiqueIndice As Integer 'Indique le type de commentaire que le joueur prononce systématiquement.
Public TempsDernierCommentaire As Double

Public Vie As Double
Public MaxVie As Long
Public BonusMaxVie As Long
Public RecuperationVie As Double
Public ExpVie As Double
Public NivVie As Long
Public ExpNivVie As Double
Public ExpNivPlusVie As Double

Public Energie As Double
Public MaxEnergie As Long
Public BonusMaxEnergie As Long
Public RecuperationEnergie As Double
Public ExpEnergie As Double
Public NivEnergie As Long
Public ExpNivEnergie As Double
Public ExpNivPlusEnergie As Double

Public Magie As Double
Public MaxMagie As Long
Public BonusMaxMagie As Long
Public RecuperationMagie As Double
Public ExpMagie As Double
Public NivMagie As Long
Public ExpNivMagie As Double
Public ExpNivPlusMagie As Double

Public Moral As Double
Public MaxMoral As Long
Public BonusMaxMoral As Long
Public RecuperationMoral As Double
Public ExpMoral As Double
Public NivMoral As Long
Public ExpNivMoral As Double
Public ExpNivPlusMoral As Double

Public Attaque As Double
Public MaxAttaque As Long
Public BonusMaxAttaque As Long
Public ExpAttaque As Double
Public NivAttaque As Long
Public ExpNivAttaque As Double
Public ExpNivPlusAttaque As Double
Public PorteeAttaque As Integer
Public LanceProjectile As Integer
Public NombreProjectiles As Integer
Public EffetAttaque As Integer
Public AttaqueResistance As Integer
Public BonusAttaqueResistance As Integer
Private SurfaceAttaque As Integer
Public BonusSurfaceAttaque As Integer
Private AttaqueAutomatique As Boolean
Public AttaqueCauseEffet As Integer
Public BonusAttaqueCauseEffet As Collection
Public Kamikaze As Boolean
Private BonusPorteeAttaque As Integer
'Public PerteAttaque As Long
Public ForcerAttaque As Boolean
Public PerteVie As Double
Public PerteEnergie As Double
Public PerteMagie As Double
Public PerteMoral As Double
Public PerteArgent As Double
Public PerteRessourceIndice As Integer
Public PerteRessourceQuantite As Double
Public Poison As Integer
Public PoisonLanceur As Integer
Public DureeRechargeAttaque As Long 'Temps de rechargement avant de pouvoir réattaquer.
Public DureeRechargeAttaqueMax As Long
Public DureeRechargeAttaqueObjet As Long 'Temps donné par l'objet

Public Defense As Double
Public MaxDefense As Long
Public BonusMaxDefense As Long
Public ExpDefense As Double
Public NivDefense As Long
Public ExpNivDefense As Double
Public ExpNivPlusDefense As Double
Public Armure As Integer
Private BonusArmure As Integer

Public Projectile As Boolean
Public ProjectileCible As Integer '0: cible un personnage, 1: cible une maison, 2: cible un château.
Public Pietinement As Integer
Public Inattaquable As Boolean

Public IA As Boolean
Public ChoisirIA As Boolean
Public ChoisirIAScript As Boolean
Public IA_Equiper As Integer
Public IA_Manger As Boolean
Public IA_Pacifique As Boolean
Public IA_Temeraire As Boolean
Public IA_Protecteur As Boolean 'Protège ses alliés avec ses soldats.
Public IA_Egoiste As Boolean 'Si vrai, ne s'occupe pas du château.
Private Berserk As Integer 'Si 1, attaque automatiquement, si 2, aide les fantômes.
Private BonusBerserk As Integer 'Berserk accordé par l'équipement.
Public IA_Ordre As Integer '1=Attaquer perso, 2=Attaquer maison, 3 attaquer château.
Public IA_Cible As Integer 'Indice de la cible qu'il est demandé d'attaquer.
Public FichierIA As String
Public FichierIAScript As String
Public NumeroIA As Integer

Public JeuIA3 As ClsJeuIA3

Public ControleJoueur As Boolean
Private Argent As Double
Public Action As Integer
Public Direction As Integer
Public EnDeplacement As Boolean
Public DureeImmobilisation As Double 'La latence est le temps en nombre de pas où le personnage est immobilisé.
Public DansUneMaison As Boolean 'Renvoie vrai si le personnage est à l'intérieur d'une maison.
Public DansSaMaison As Boolean
Public IndiceMaison As Long 'Mémorise l'indice de la maison vers laquelle le joueur se dirige.
Public DansUnChateau As Boolean 'Renvoie vrai si le personnage est à l'intérieur d'un chateau.
Public DansSonChateau As Boolean

'Mémorise les indices des destinations vers lesquels le joueur se dirige.
Public IndiceRessource As Long
Public IndiceChateau As Long 'Mémorise l'indice du chateau vers lequel le joueur se dirige.
Public IndicePerso As Long
Public IndiceTresor As Long

'Variables de ressources.
'Indique l'indice+1 de la ressource que l'on est en train de récolter
'0 indique que le perso ne récolte pas de ressource.
Public RessourceCourante As Long
Public ServiceCourant As Long
'Stock de ressources.
Private Ressources() As Double
Private MaxRessources() As Long
Public BonusMaxRessourcesRace As Long
Private BonusMaxRessourcesObjet As Long

'Compétences spéciales.
Private CaracSpeciales() As Long
Private ExpSpeciales() As Double
Private NivSpeciales() As Long
Private BonusNivSpeciales() As Long
Private ExpNivSpeciales() As Double
Private ExpNivPlusSpeciales() As Double
'Compétences des ressources.
Private ExpRessources() As Double
Private NivRessources() As Long
Private BonusNivRessources() As Long
Private ExpNivRessources() As Double
Private ExpNivPlusRessources() As Double
'Compétences des Services.
Private CaracServices() As Long
Private ExpServices() As Double
Private NivServices() As Long
Private BonusNivServices() As Long
Private ExpNivServices() As Double
Private ExpNivPlusServices() As Double
'Compétences des objets.
Public CaracObjetsDepart As Long
Private CaracObjets() As Long
Private ExpObjets() As Double
Private NivObjets() As Long
Private BonusNivObjets() As Long
Private ExpNivObjets() As Double
Private ExpNivPlusObjets() As Double

'Résitances.
Private Resistances() As Single 'Score de base.
Private BonusResistances() As Single 'Score en comptant les bonus de l'équipement.

'Modificateurs de la race
Public CoefRaceVie As Single
Public CoefRaceEnergie As Single
Public CoefRaceMagie As Single
Public CoefRaceMoral As Single
Public CoefRaceAttaque As Single
Public CoefRaceDefense As Single
Private CoefRaceCompSpeciales() As Single 'Compétences spéciales.
Private CoefRaceCompRessources() As Single 'Compétences des ressources.
Private CoefRaceCompServices() As Single 'Compétences des Services.
Private CoefRaceCompObjets() As Single 'Compétences des objets.

'Objets.
Public IndiceObjet As Integer 'Indice de l'objet que le joueur fabrique.
Public ObjetSelectionne As Integer 'Objet saisie avec le clique gauche.
Public RessourceSelectionne As Boolean 'Vrai si le personnage a saisi des ressources pour les donner ou les déposer.
Private RessourcesSelectionnes() As Double
Public ArgentSelectionne As Double
Public EtatFabricationObjet As Double
Private ObjetsEquipes() As ClsJeuObjet
Private ObjetsInventaires() As ClsJeuObjet
Public NumeroProjectileLance As Integer 'Mémorise l'indice de l'objet lancé.
Public ProjectileEnCours As Integer 'Nombre de projectiles lancés par le personnage.
Public NumeroObjetUtilise As Integer 'Mémorise l'indice de l'objet utilisé.
Private GenerationRessources As Boolean
Private GainArgent As Single
Private GainRessource() As Single
Public Monture As Integer 'Si le personange a une monture, cette valeur indique le numéro de l'apparence de cette monture.

Private EffetsTemp As Collection
Private EffetsRecharge As Collection

'Bestiaire.
Private Bestiaire() As Long

'Variable temporaire.
Public Vitalite As Double
Public BonusVitalite As Double
Public Longueur_Pas As Long
Public Longueur_PasTemp As Long 'Variable temporaire utilisée pour calculer l'accélération
Private Immobile As Boolean
Public CheminX As Collection 'Utilisé pour calculer le chemin le plus court.
Public CheminY As Collection 'Utilisé pour calculer le chemin le plus court.
Private CheminDirectionX As Long
Private CheminDirectionY As Long
Const CheminNombres = 3
Const CheminAngle = 3.14 / CheminNombres
Private TempCheminX As Long
Private TempCheminY As Long
Private CheminComplet As Boolean 'Si vrai, le chemin complet a été calculé.

'Stats de la partie
Public FragsPersos As Long
Public FragsBatiments As Long
Public FragsChateaux As Long
Public Morts As Long

'Variables temporaires utilisées par les IAs scriptées
'Public Temp As Collection

Private Sub Class_Initialize()
    Set ApparenceObjet = New Collection
    Set EffetsTemp = New Collection
    Set EffetsRecharge = New Collection
    Set BonusAttaqueCauseEffet = New Collection
    Set CheminX = New Collection
    Set CheminY = New Collection
    ReDim ApparenceDecalageX(4)  'Quatre directions différentes.
    ReDim ApparenceDecalageY(4)  'Quatre directions différentes.
    ReDim ApparenceAvantPlan(4) 'Quatre directions différentes.
End Sub

Public Sub Init(ByVal Indice As Long, ByVal Parametres As ClsJeuParam)
    Dim i As Integer
    
    Numero = Indice
    
    'Choisit une apparence au hasard.
    NumeroPas = Int(Rnd * 3)
    
    ChoisirApparence = False
    FichierApparence = ""
    
    Argent = 0
    
    Attitude = 0
    NumeroChef = -1
    NumeroObjetUtilise = -1
    NumeroProjectileLance = -1
    ObjetSelectionne = -1
    RessourceSelectionne = False
    ReDim RessourcesSelectionnes(Parametres.NombreRessources - 1)
    ArgentSelectionne = 0
    
    SurfaceAttaque = 0
    
    FragsPersos = 0
    FragsBatiments = 0
    FragsChateaux = 0
    Morts = 0
    
    ResurrectionNombre = -1
    VitesseResurrection = VitesseResurrectionDepart
    
    Kamikaze = False
    Projectile = False
    Pietinement = 0
    Inattaquable = False
    DureeRechargeAttaque = 0
    NombreProjectiles = 1
    
    IndiceRessource = 0
    IndiceChateau = NumeroFief
    IndicePerso = Numero
    
    Init_Carac Parametres
    Vie = MaxVie
    Energie = MaxEnergie
    Magie = MaxMagie
    Moral = MaxMoral
    Attaque = MaxAttaque
    Defense = MaxDefense
    
    ReDim Ressources(Parametres.NombreRessources - 1)
    GenerationRessources = False
    GainArgent = 0
    ReDim GainRessource(Parametres.NombreRessources - 1)
    
    ReDim Resistances(Parametres.NombreResistances - 1)
    ReDim BonusResistances(Parametres.NombreResistances - 1)
    
    Init_Competences
    
    ReDim ObjetsEquipes(Parametres.PersosNombreObjetsEquipes - 1)
    ReDim ObjetsInventaires(Parametres.PersosNombreObjetsInventaire - 1)
    For i = 0 To UBound(ObjetsEquipes())
        Set ObjetsEquipes(i) = New ClsJeuObjet
        ObjetsEquipes(i).Init Parametres
    Next i
    For i = 0 To UBound(ObjetsInventaires())
        Set ObjetsInventaires(i) = New ClsJeuObjet
        ObjetsInventaires(i).Init Parametres
    Next i
        
    EffetTemp_Effacer_Tout
    EffetRecharge_Effacer_Tout

    Changer_Race_Personnage 0, Parametres 'Humain par défaut.
    
    Ajuster_Bonus
    Calculer_Vitalite
    
    ReDim Bestiaire(Parametres.NombreRaces - 1)
    
    Arreter
    
    Set JeuIA3 = New ClsJeuIA3
    IA = False
    ChoisirIA = False
    ChoisirIAScript = False
    'Set Temp = New Collection
End Sub

Public Sub Init_Carac(ByVal Parametres As ClsJeuParam)
    Dim CoefVie As Single
    Dim CoefEnergie As Single
    Dim CoefMagie As Single
    Dim CoefMoral As Single
    Dim CoefAttaque As Single
    Dim CoefDefense As Single
    
    If Definir_MaxVie > 0 And Definir_MaxEnergie > 0 And Definir_MaxMagie > 0 And Definir_MaxMoral > 0 And Definir_MaxAttaque > 0 And Definir_MaxDefense > 0 Then
        CoefVie = Vie / Definir_MaxVie
        CoefEnergie = Energie / Definir_MaxEnergie
        CoefMagie = Magie / Definir_MaxMagie
        CoefMoral = Moral / Definir_MaxMoral
        CoefAttaque = Attaque / Definir_MaxAttaque
        CoefDefense = Defense / Definir_MaxDefense
    Else
        CoefVie = 1
        CoefEnergie = 1
        CoefMagie = 1
        CoefMoral = 1
        CoefAttaque = 1
        CoefDefense = 1
    End If
    
    'Réinitialise les caractériques.
    Parole = True
    Volant = False
    GainExperience = True
    ChoisirIA = False
    IA_Equiper = 0
    IA_Manger = True
    IA_Pacifique = False
    IA_Temeraire = False
    IA_Protecteur = False
    IA_Egoiste = False
    IA_Berserk = 0

    'Chargement des caractéristiques par défaut.
    MaxVie = Parametres.PersosMaxVie
    Vie = Definir_MaxVie * CoefVie
    MaxEnergie = Parametres.PersosMaxEnergie
    Energie = Definir_MaxEnergie * CoefEnergie
    MaxMagie = Parametres.PersosMaxMagie
    Magie = Definir_MaxMagie * CoefMagie
    MaxMoral = Parametres.PersosMaxMoral
    Moral = Definir_MaxMoral * CoefMoral
        
    ExpVie = 0: NivVie = 0: ExpNivVie = 0: ExpNivPlusVie = CoefExpVie
    ExpEnergie = 0: NivEnergie = 0: ExpNivEnergie = 0: ExpNivPlusEnergie = CoefExpEnergie
    ExpMagie = 0: NivMagie = 0: ExpNivMagie = 0: ExpNivPlusMagie = CoefExpMagie
    ExpMoral = 0: NivMoral = 0: ExpNivMoral = 0: ExpNivPlusMoral = CoefExpMoral
    ExpAttaque = 0: NivAttaque = 0: ExpNivAttaque = 0: ExpNivPlusAttaque = CoefExpAttaque
    ExpDefense = 0: NivDefense = 0: ExpNivDefense = 0: ExpNivPlusDefense = CoefExpDefense
        
    MaxAttaque = Parametres.PersosAttaque
    Attaque = Definir_MaxAttaque * CoefAttaque
    PorteeAttaque = Parametres.PersosPorteeAttaque
    MaxDefense = Parametres.PersosDefense
    Defense = Definir_MaxDefense * CoefDefense
    
    LongueurPas = Parametres.PersosVitesse
    LongueurPasMax = 16
    Acceleration = Parametres.PersosAcceleration
    Init_Competences
End Sub

Private Sub Init_Competences()
    Dim i As Integer
    ReDim CaracSpeciales(Parametres.NombreCompetencesSpeciales - 1)
    ReDim ExpSpeciales(Parametres.NombreCompetencesSpeciales - 1)
    ReDim NivSpeciales(Parametres.NombreCompetencesSpeciales - 1)
    ReDim BonusNivSpeciales(Parametres.NombreCompetencesSpeciales - 1)
    ReDim ExpNivSpeciales(Parametres.NombreCompetencesSpeciales - 1)
    ReDim ExpNivPlusSpeciales(Parametres.NombreCompetencesSpeciales - 1)
    For i = 0 To UBound(CaracSpeciales())
        CaracSpeciales(i) = Parametres.PersosNiveauSpecialesDepart
        ExpNivPlusSpeciales(i) = CoefExpCompSpeciales
    Next i
    
    ReDim MaxRessources(Parametres.NombreRessources - 1)
    ReDim ExpRessources(Parametres.NombreRessources - 1)
    ReDim NivRessources(Parametres.NombreRessources - 1)
    ReDim BonusNivRessources(Parametres.NombreRessources - 1)
    ReDim ExpNivRessources(Parametres.NombreRessources - 1)
    ReDim ExpNivPlusRessources(Parametres.NombreRessources - 1)
    For i = 0 To UBound(MaxRessources())
        MaxRessources(i) = Parametres.PersosMaxRessources
        ExpNivPlusRessources(i) = CoefExpCompRessources
    Next i
    
    ReDim CaracServices(Parametres.NombreServices - 2)
    ReDim ExpServices(Parametres.NombreServices - 2)
    ReDim NivServices(Parametres.NombreServices - 2)
    ReDim BonusNivServices(Parametres.NombreServices - 2)
    ReDim ExpNivServices(Parametres.NombreServices - 2)
    ReDim ExpNivPlusServices(Parametres.NombreServices - 2)
    For i = 0 To UBound(CaracServices())
        CaracServices(i) = Parametres.PersosNiveauServicesDepart
        ExpNivPlusServices(i) = CoefExpCompServices
    Next i
    
    CaracObjetsDepart = Parametres.PersosNiveauObjetsDepart
    ReDim CaracObjets(Parametres.NombreCompetencesObjets - 1)
    ReDim ExpObjets(Parametres.NombreCompetencesObjets - 1)
    ReDim NivObjets(Parametres.NombreCompetencesObjets - 1)
    ReDim BonusNivObjets(Parametres.NombreCompetencesObjets - 1)
    ReDim ExpNivObjets(Parametres.NombreCompetencesObjets - 1)
    ReDim ExpNivPlusObjets(Parametres.NombreCompetencesObjets - 1)
    For i = 0 To UBound(CaracObjets())
        CaracObjets(i) = CaracObjetsDepart
        ExpNivPlusObjets(i) = CoefExpCompObjets
    Next i
End Sub

' ************************************************************************************
' *                              Fonctions de déplacement                            *
' ************************************************************************************

Public Sub Aller_A(ByVal X As Long, ByVal Y As Long, Optional ByVal SansCommentaire As Boolean)
    DirectionX = X
    DirectionY = Y
    Orienter_Direction DirectionX, DirectionY
    Action = 1
    Visible = True
    If NombreSoldats > 0 Then
        If Vivant Then Ordonner_Chargez Numero, , , , , , SansCommentaire
    End If
End Sub

Public Sub Arreter()
    DirectionX = PositionX
    DirectionY = PositionY
    Decelerer
    If DansSaMaison Then
        Action = 100
    Else
        Action = 0
    End If
    RessourceCourante = -1
    If IA And Visible Then 'Si c'est une IA qui s'arrête, elle refait quelque chose.
        Comportement_IA_Persos Numero, True
    End If
    'If NombreSoldats > 0 Then Ordonner_Arret Numero
End Sub

Private Sub Decelerer()
    'Longueur_PasTemp = Longueur_PasTemp - Acceleration
    'If Longueur_PasTemp < 0 Then
        Longueur_PasTemp = 0
    'End If
End Sub

Public Sub Rentrer_Maison(ByVal Maison As ClsJeuBatiment, Optional ByVal Rentrer As Boolean)
    If Maison.Visible Then
        DirectionX = Maison.PositionX + Maison.Largeur / 2 - Largeur / 2 + Maison.EntreePositionX
        DirectionY = Maison.PositionY + Maison.Hauteur / 2 - Hauteur / 2 + Maison.EntreePositionY
        Orienter_Direction DirectionX, DirectionY
        If Maison.Construit And Maison.Definir_Vie = Maison.MaxVie Or Not Vivant Or Rentrer Then
            Action = 101
        Else
            Action = 7
            'DansSaMaison = False
        End If
        IndiceMaison = Maison.Numero
        Visible = True
    End If
    If NombreSoldats > 0 Then
        If Vivant Then Ordonner_Chargez Numero
    End If
End Sub

Public Sub Aller_Maison(ByVal Maison As ClsJeuBatiment, Optional ByVal Reparer As Boolean)
    With Maison
    If .Visible Then
        DirectionX = .PositionX + .Largeur / 2 - Largeur / 2 + .EntreePositionX
        DirectionY = .PositionY + .Hauteur / 2 - Hauteur / 2 + .EntreePositionY
        Orienter_Direction DirectionX, DirectionY
        If Maison.Construit And Maison.Definir_Vie = Maison.MaxVie Or Not Vivant Or Not Reparer Then
            Action = 201
        Else
            Action = 7
        End If
        IndiceMaison = .Numero
        Visible = True
    End If
    End With
    If NombreSoldats > 0 Then
        If Vivant Then Ordonner_Chargez Numero
    End If
End Sub

Public Sub Rentrer_Chateau(ByVal Chateau As ClsJeuBatiment)
    If Not Hostile Then 'Les personnes hostiles ne peuvent pas rentrer dans leur château.
        If Chateau.Visible Then
            DirectionX = Chateau.PositionX + Chateau.Largeur / 2 - Largeur / 2 + Chateau.EntreePositionX
            DirectionY = Chateau.PositionY + Chateau.Hauteur / 2 - Hauteur / 2 + Chateau.EntreePositionY
            Orienter_Direction DirectionX, DirectionY
            Action = 202
            Visible = True
        End If
        If NombreSoldats > 0 Then
            If Vivant Then Ordonner_Chargez Numero
        End If
    End If
End Sub

Public Sub Aller_Chateau(ByVal Chateau As ClsJeuBatiment, _
                         ByVal Indice As Long)
    If Chateau.Visible Then
        DirectionX = Chateau.PositionX + Chateau.Largeur / 2 - Largeur / 2 + Chateau.EntreePositionX
        DirectionY = Chateau.PositionY + Chateau.Hauteur / 2 - Hauteur / 2 + Chateau.EntreePositionY
        Orienter_Direction DirectionX, DirectionY
        Action = 203
        IndiceChateau = Indice
        Visible = True
    End If
    If NombreSoldats > 0 Then Ordonner_Chargez Numero
End Sub

Public Sub Aller_Ressource(ByVal Ressource As ClsJeuRessources, _
                           ByVal Indice As Long, _
                           ByVal LargeurRessource As Long, _
                           ByVal HauteurRessource As Long)
    DirectionX = Ressource.PositionX + LargeurRessource / 2 - Largeur / 2
    DirectionY = Ressource.PositionY + LargeurRessource / 2 - Hauteur / 2
    Orienter_Direction DirectionX, DirectionY
    IndiceRessource = Indice
    Action = 3
    Visible = True
    If NombreSoldats > 0 Then Ordonner_Chargez Numero
End Sub

Public Sub Aller_Tresor(ByVal Tresor As ClsJeuTresors, _
                        ByVal Indice As Long, _
                        ByVal LargeurTresor As Long, _
                        ByVal HauteurTresor As Long)
    DirectionX = Tresor.X + LargeurTresor / 2 - Largeur / 2
    DirectionY = Tresor.Y + LargeurTresor / 2 - Hauteur / 2
    Orienter_Direction DirectionX, DirectionY
    IndiceTresor = Indice
    Action = 204
    Visible = True
    If NombreSoldats > 0 Then Ordonner_Chargez Numero
End Sub

Public Sub Aller_Donner(ByVal Indice As Long)
    IndicePerso = Indice
    'ObjetSelectionne = TypeObjet
    Action = 205
    Visible = True
End Sub

Public Sub Aller_Vendre(ByVal Indice As Long, ByVal Marche As Boolean)
    If Marche Then
        If DansUneMaison And IndiceMaison = Indice Then
            Fiefs(Persos(Indice).NumeroFief).Vendre_ObjetsSelectionnees Me, Parametres
        Else
            Aller_Maison Maisons(Indice)
            Action = 207
            Visible = True
        End If
    Else
        If NumeroFief = Indice Then
            If DansSonChateau Then
                Fiefs(Indice).Vendre_ObjetsSelectionnees Me, Parametres
            Else
                Rentrer_Chateau Chateaux(Indice)
                Action = 206
                Visible = True
            End If
        Else
            If DansUnChateau And IndiceChateau = Indice Then
                Fiefs(Indice).Vendre_ObjetsSelectionnees Me, Parametres
            Else
                Aller_Chateau Chateaux(Indice), Indice
                Action = 206
                Visible = True
            End If
        End If
    End If
End Sub

' ************************************************************************************
' *                              Fonctions d'ordres d'attaque                        *
' ************************************************************************************

Public Sub Attaquer_Perso(ByVal Indice As Long)
    Action = 300
    IndicePerso = Indice
    If DansUneMaison And Maisons(IndiceMaison).Tir And Definir_PorteeAttaque > 0 Then
        With Persos(Indice)
        If Not (.PositionX + .Largeur >= PositionX - Definir_PorteeAttaque And _
           .PositionY + .Hauteur >= PositionY - Definir_PorteeAttaque And _
           .PositionX <= PositionX + Largeur + Definir_PorteeAttaque And _
           .PositionY <= PositionY + Hauteur + Definir_PorteeAttaque) Then
            Visible = True
        End If
        End With
    Else
        Visible = True
    End If
    DureeRechargeAttaque = DureeRechargeAttaqueObjet
    DureeRechargeAttaqueMax = DureeRechargeAttaqueObjet
    ForcerAttaque = True
    If NombreSoldats > 0 Then Ordonner_Chargez Numero
End Sub

Public Sub Attaquer_Maison(ByVal Indice As Long)
    Action = 301
    IndiceMaison = Indice
    Visible = True
    DureeRechargeAttaque = DureeRechargeAttaqueObjet
    DureeRechargeAttaqueMax = DureeRechargeAttaqueObjet
    ForcerAttaque = True
    If NombreSoldats > 0 Then Ordonner_Chargez Numero
End Sub

Public Sub Attaquer_Chateau(ByVal Indice As Long)
    Action = 302
    IndiceChateau = Indice
    Visible = True
    DureeRechargeAttaque = DureeRechargeAttaqueObjet
    DureeRechargeAttaqueMax = DureeRechargeAttaqueObjet
    ForcerAttaque = True
    If NombreSoldats > 0 Then Ordonner_Chargez Numero
End Sub

Public Sub Attaquer_Terrain(ByVal X As Long, ByVal Y As Long, Optional ByVal TypeCible As Integer)
    If LanceProjectile >= 0 Then
        ProjectileCible = TypeCible
        'Arreter
        Action = 303
        If Not ((DansSaMaison Or DansUneMaison) And Maisons(IndiceMaison).Tir) Then
            Visible = True
        End If
        DureeRechargeAttaque = DureeRechargeAttaqueObjet
        DureeRechargeAttaqueMax = DureeRechargeAttaqueObjet
        ForcerAttaque = True
        'DirectionX = X
        'DirectionY = Y
        CibleX = X
        CibleY = Y
        Orienter_Direction CibleX, CibleY
    End If
End Sub

Public Sub Attaquer_Allies(ByVal Valeur As Boolean)
    'Le perso devient hostile à tout le monde.
    Dim i As Long
    If Valeur Then
        NumeroEquipe = -Numero
        Arreter
    Else
        NumeroEquipe = Fiefs(NumeroFief).NumeroEquipe
    End If
    'Toutes les invocations changent également de numéro d'équipe.
    If NombreSoldats > 0 Then
        For i = 0 To LPerso
            If Persos(i).NumeroChef = Numero Then
                Persos(i).NumeroEquipe = NumeroEquipe
                Persos(i).Arreter
            End If
        Next i
    End If
End Sub

' ************************************************************************************
' *                              Fonctions de récuparation                           *
' ************************************************************************************

Public Sub Recuperer_AutoVie(ByVal Parametres As ClsJeuParam)
    If Definir_Vie < Definir_MaxVie * 1 Then
        Recuperer_Vie Parametres
    End If
End Sub
Public Sub Recuperer_Vie(ByVal Parametres As ClsJeuParam)
    'Mange tout ce qu'il peut pour récuperer de la vie.
    Dim X As Integer
    For X = 0 To Parametres.NombreRessources - 1
        If Parametres.Ressources_Mangeable(X) And _
           Parametres.Ressources_GainVie(X) > 0 And _
           Definir_ressources(X) >= 1 Then
            Do
                Manger_Ressources X, Parametres
            Loop Until Definir_Vie >= Definir_MaxVie - Parametres.Ressources_GainVie(X) * Manger_Ressources_Coef(0) Or _
                       Definir_ressources(X) < 1
            X = Parametres.NombreRessources
        End If
    Next X
End Sub
Public Sub Recuperer_AutoEnergie(ByVal Parametres As ClsJeuParam)
    If Definir_Energie < Definir_MaxEnergie * 0.5 Then
        Recuperer_Energie Parametres
    End If
End Sub
Public Sub Recuperer_Energie(ByVal Parametres As ClsJeuParam)
    'Mange tout ce qu'il peut pour récuperer de l'énergie.
    Dim X As Integer
    For X = 0 To Parametres.NombreRessources - 1
        If Parametres.Ressources_Mangeable(X) And _
           Parametres.Ressources_GainEnergie(X) > 0 And _
           Definir_ressources(X) >= 1 Then
            Do
                Manger_Ressources X, Parametres
            Loop Until Definir_Energie >= Definir_MaxEnergie - Parametres.Ressources_GainEnergie(X) * Manger_Ressources_Coef(1) Or _
                       Definir_ressources(X) < 1
            X = Parametres.NombreRessources
        End If
    Next X
End Sub
Public Sub Recuperer_AutoMagie(ByVal Parametres As ClsJeuParam)
    If Definir_Magie < Definir_MaxMagie * 0.5 Then
        Recuperer_Magie Parametres
    End If
End Sub
Public Sub Recuperer_Magie(ByVal Parametres As ClsJeuParam)
    'Mange tout ce qu'il peut pour récuperer de la magie.
    Dim X As Integer
    For X = 0 To Parametres.NombreRessources - 1
        If Parametres.Ressources_Mangeable(X) And _
           Parametres.Ressources_GainMagie(X) > 0 And _
           Definir_ressources(X) >= 1 Then
            Do
                Me.Manger_Ressources X, Parametres
            Loop Until Definir_Magie >= Definir_MaxMagie - Parametres.Ressources_GainMagie(X) * Manger_Ressources_Coef(2) Or _
                       Definir_ressources(X) < 1
            X = Parametres.NombreRessources
        End If
    Next X
End Sub
Public Sub Recuperer_AutoMoral(ByVal Parametres As ClsJeuParam)
    If Definir_Moral < Definir_MaxMoral * 0.5 Then
        Recuperer_Moral Parametres
    End If
End Sub
Public Sub Recuperer_Moral(ByVal Parametres As ClsJeuParam)
    'Mange tout ce qu'il peut pour récuperer de la magie.
    Dim X As Integer
    For X = 0 To Parametres.NombreRessources - 1
        If Parametres.Ressources_Mangeable(X) And _
           Parametres.Ressources_GainMoral(X) > 0 And _
           Definir_ressources(X) >= 1 Then
            Do
                Manger_Ressources X, Parametres
            Loop Until Definir_Moral >= Definir_MaxMoral - Parametres.Ressources_GainMoral(X) * Manger_Ressources_Coef(3) Or _
                       Definir_ressources(X) < 1
            X = Parametres.NombreRessources
        End If
    Next X
End Sub

' ************************************************************************************
' *                               Fonction principale                                *
' ************************************************************************************

Public Sub Deplacement1(ByRef RessourceCible As ClsJeuRessources, _
                        ByRef PersoCible As ClsJeuPerso, _
                        ByRef MaisonCible As ClsJeuBatiment, _
                        ByRef ChateauCible As ClsJeuBatiment, _
                        ByRef Fief As ClsJeuFief, _
                        ByRef NoCommentaire As Integer, _
                        ByRef Commentaires As ClsJeuCommentaires, _
                        ByVal Parametres As ClsJeuParam)
    Calculer_Vitalite
    
    If EnCombat Or Action = 205 Then
        Select Case Action
        Case 205:
            If (PersoCible.Visible Or PersoCible.DansSaMaison) And PersoCible.Vivant Then
                DirectionX = PersoCible.PositionX + PersoCible.Largeur / 2 - Largeur / 2
                DirectionY = PersoCible.PositionY + PersoCible.Hauteur / 2 - Hauteur / 2
            Else
                Arreter
            End If
            Orienter_Direction DirectionX, DirectionY
        Case 300:
            If PersoCible.Visible And PersoCible.Vivant Then
                DirectionX = PersoCible.PositionX + PersoCible.Largeur / 2 - Largeur / 2
                DirectionY = PersoCible.PositionY + PersoCible.Hauteur / 2 - Hauteur / 2
            ElseIf PersoCible.Vivant Then
                If IA Then 'Les IA poursuivent le réfugié.
                    If PersoCible.DansSaMaison Then
                        Attaquer_Maison PersoCible.Numero
                    ElseIf PersoCible.DansUneMaison Then
                        Attaquer_Maison PersoCible.IndiceMaison
                    ElseIf PersoCible.DansSonChateau Then
                        Attaquer_Chateau PersoCible.NumeroFief
                    ElseIf PersoCible.DansUnChateau Then
                        Attaquer_Chateau PersoCible.IndiceChateau
                    End If
                Else
                    Arreter
                End If
                If Not (Maisons(PersoCible.Numero).Tir) Then
                    Definir_Commentaires = Commentaires.Message(15, Me)
                End If
            Else
                Arreter
                If Projectile Then Vivant = False
            End If
            Orienter_Direction DirectionX, DirectionY
        Case 301:
            If MaisonCible.Visible Then
                DirectionX = MaisonCible.PositionX + MaisonCible.Largeur / 2 - Largeur / 2
                DirectionY = MaisonCible.PositionY + MaisonCible.Hauteur / 2 - Hauteur / 2
            Else
                Arreter
                If Projectile Then Vivant = False
            End If
            Orienter_Direction DirectionX, DirectionY
        Case 302:
            If Chateaux(IndiceChateau).Visible Then
                DirectionX = ChateauCible.PositionX + ChateauCible.Largeur / 2 - Largeur / 2
                DirectionY = ChateauCible.PositionY + ChateauCible.Hauteur / 2 - Hauteur / 2
            Else
                Arreter
                If Projectile Then Vivant = False
            End If
            Orienter_Direction DirectionX, DirectionY
        End Select
    End If
End Sub
Public Sub Deplacement2(ByRef RessourceCible As ClsJeuRessources, _
                        ByRef PersoCible As ClsJeuPerso, _
                        ByRef MaisonCible As ClsJeuBatiment, _
                        ByRef ChateauCible As ClsJeuBatiment, _
                        ByRef Fief As ClsJeuFief, _
                        ByRef NoCommentaire As Integer, _
                        ByRef Commentaires As ClsJeuCommentaires, _
                        ByVal Parametres As ClsJeuParam)
    If Action = 205 And _
        PersoCible.PositionX + PersoCible.Largeur >= PositionX And _
        PersoCible.PositionY + PersoCible.Hauteur >= PositionY And _
        PersoCible.PositionX <= PositionX + Largeur And _
        PersoCible.PositionY <= PositionY + Hauteur _
       Or Action = 300 And _
        PersoCible.PositionX + PersoCible.Largeur >= PositionX - Definir_PorteeAttaque And _
        PersoCible.PositionY + PersoCible.Hauteur >= PositionY - Definir_PorteeAttaque And _
        PersoCible.PositionX <= PositionX + Largeur + Definir_PorteeAttaque And _
        PersoCible.PositionY <= PositionY + Hauteur + Definir_PorteeAttaque _
       Or Action = 301 And _
        MaisonCible.PositionX + MaisonCible.Largeur >= PositionX - Definir_PorteeAttaque And _
        MaisonCible.PositionY + MaisonCible.Hauteur >= PositionY - Definir_PorteeAttaque And _
        MaisonCible.PositionX <= PositionX + Largeur + Definir_PorteeAttaque And _
        MaisonCible.PositionY <= PositionY + Hauteur + Definir_PorteeAttaque _
       Or Action = 302 And _
        ChateauCible.PositionX + ChateauCible.Largeur >= PositionX - Definir_PorteeAttaque And _
        ChateauCible.PositionY + ChateauCible.Hauteur >= PositionY - Definir_PorteeAttaque And _
        ChateauCible.PositionX <= PositionX + Largeur + Definir_PorteeAttaque And _
        ChateauCible.PositionY <= PositionY + Hauteur + Definir_PorteeAttaque _
       Or Action = 303 _
       Then
        If Action = 205 Then
            'Donne un objet.
            If ObjetSelectionne >= 0 Then
                If Not Parametres.Objet_Inseparable(ObjetSelectionne) Then
                    PersoCible.Objet_Ajouter_Inventaire ObjetSelectionne, Parametres, Commentaires, True
                    ObjetSelectionne = -1
                End If
            End If
            'Donne des ressources.
            Ressources_Donner PersoCible
            If Parole Then PersoCible.Definir_Commentaires = Commentaires.Message(46, Me)
            Arreter
        Else
            'If Attaque >= 1 And _
               Not (ProjectileEnCours And LanceProjectile >= 0) And _
               (LanceProjectile = -1 Or Attaque >= (PersoCible.Vie + 1) / CoefDegatPerso + PersoCible.Definir_MaxDefense Or Attaque = Definir_MaxAttaque) And _
               (Attaque = Definir_MaxAttaque Or _
               ForcerAttaque Or _
               (Action = 300 And _
                IIf(ObjetsEquipes(0).PerteSiAttaque Or ObjetsEquipes(1).PerteSiAttaque, Attaque >= PersoCible.Vie / CoefDegatPerso + PersoCible.Defense, Attaque > PersoCible.Vie + PersoCible.Defense + 1))) Then
            If DureeRechargeAttaque = 0 Then
            If Attaque >= 1 And _
               (Not (ProjectileEnCours > 0 And LanceProjectile >= 0) And Attaque >= (PersoCible.Vie + 1) / CoefDegatPerso / PersoCible.Definir_Resistance(Definir_AttaqueResistance) + PersoCible.Definir_MaxDefense Or Attaque = Definir_MaxAttaque) And _
               IIf(ProjectileEnCours > 0 And LanceProjectile >= 0, False, Attaque = Definir_MaxAttaque) Or _
               ForcerAttaque Or _
               (Not (ProjectileEnCours > 0 And LanceProjectile >= 0) And Action = 300 And _
                IIf(ObjetsEquipes(0).PerteSiAttaque Or ObjetsEquipes(1).PerteSiAttaque, Attaque >= PersoCible.Vie / CoefDegatPerso + PersoCible.Defense, Attaque > PersoCible.Vie + PersoCible.Defense + 1)) Or _
                AttaqueAutomatique Then
                'S'il peut achever la cible, il l'achève.
                If Vie > PerteVie Then
                    If Energie > PerteEnergie Then
                        If Magie >= PerteMagie Then
                            If Moral > PerteMoral Then
                                If Definir_Argent >= PerteArgent Then
                                    'If Ressources(PerteRessourceIndice) >= PerteRessourceQuantite * NombreProjectiles Then
                                    If Ressources(PerteRessourceIndice) >= PerteRessourceQuantite Then
                                        'Le personnage attaque.
                                        If LanceProjectile >= 0 Then
                                            NumeroProjectileLance = LanceProjectile
                                            'If Action = 303 Then Aller_A DirectionX, DirectionY
                                            If Action = 303 Then
                                                Deplacement3 RessourceCible, PersoCible, MaisonCible, ChateauCible, Fief, NoCommentaire, Commentaires, Parametres
                                                'If Direction > 0 Or Volant Then
                                                '    NumeroPas = NumeroPas + 1
                                                '    If NumeroPas > 3 Then NumeroPas = 0
                                                'End If
                                            End If
                                        Else
                                            If BonusSurfaceAttaque > 0 Then
                                                Select Case Action
                                                Case 300: Perso_Attaquer_Surface Me, PersoCible.PositionX, PersoCible.PositionY, BonusSurfaceAttaque, ControleJoueur
                                                Case 301: Perso_Attaquer_Surface Me, MaisonCible.PositionX, MaisonCible.PositionY, BonusSurfaceAttaque, ControleJoueur
                                                Case 302: Perso_Attaquer_Surface Me, ChateauCible.PositionX, ChateauCible.PositionY, BonusSurfaceAttaque, ControleJoueur, IndiceChateau
                                                End Select
                                            Else
                                                Select Case Action
                                                Case 300: Perso_Attaquer_Perso Me, PersoCible
                                                Case 301: Perso_Attaquer_Batiment Me, MaisonCible, Persos(IndiceMaison)
                                                Case 302: Perso_Attaquer_Batiment Me, ChateauCible, Persos(IndicePerso), IndiceChateau
                                                End Select
                                            End If
                                        End If
                                        If Objet_Equipes_Actif(0) Then
                                            Definir_Vie = Vie - ObjetsEquipes(0).PerteVie - ObjetsEquipes(1).PerteVie
                                            Definir_Energie = Energie - ObjetsEquipes(0).PerteEnergie - ObjetsEquipes(1).PerteEnergie
                                            Definir_Magie = Magie - ObjetsEquipes(0).PerteMagie - ObjetsEquipes(1).PerteMagie
                                            Definir_Moral = Moral - ObjetsEquipes(0).PerteMoral - ObjetsEquipes(1).PerteMoral
                                            'Les objets a utilisation unique disparaissent.
                                            If ObjetsEquipes(0).UtilisationUnique Then
                                                ObjetsEquipes(0).Decharger
                                            End If
                                            If ObjetsEquipes(1).UtilisationUnique Then
                                                ObjetsEquipes(1).Decharger
                                            End If
                                        End If
                                        ForcerAttaque = False
                                        Definir_Argent = Definir_Argent - PerteArgent
                                        'Definir_ressources(PerteRessourceIndice) = Definir_ressources(PerteRessourceIndice) - PerteRessourceQuantite * NombreProjectiles
                                        Definir_ressources(PerteRessourceIndice) = Definir_ressources(PerteRessourceIndice) - PerteRessourceQuantite
                                        DureeRechargeAttaque = DureeRechargeAttaqueObjet
                                        DureeRechargeAttaqueMax = DureeRechargeAttaqueObjet
                                        'Les kamikazes meurent après avoir attaqué.
                                        If Kamikaze Then Vivant = False
                                    Else
                                        'N'a plus de ressources pour attaquer.
                                        Definir_Commentaires = Commentaires.Message(34, Me)
                                        Objet_Ranger_Armes Commentaires, Parametres
                                    End If
                                Else
                                    'N'a plus d'argent pour attaquer.
                                    Definir_Commentaires = Commentaires.Message(45, Me)
                                    Objet_Ranger_Armes Commentaires, Parametres
                                End If
                            Else
                                'N'a plus de point de moral.
                                Definir_Commentaires = Commentaires.Message(20, Me)
                                Objet_Ranger_Armes Commentaires, Parametres
                            End If
                        Else
                            'N'a plus de point de magie.
                            Definir_Commentaires = Commentaires.Message(19, Me)
                            Objet_Ranger_Armes Commentaires, Parametres
                        End If
                    Else
                        'N'a plus de point d'énergie.
                        Definir_Commentaires = Commentaires.Message(18, Me)
                        Objet_Ranger_Armes Commentaires, Parametres
                    End If
                Else
                    'N'a plus de point de vie.
                    Definir_Commentaires = Commentaires.Message(17, Me)
                    Objet_Ranger_Armes Commentaires, Parametres
                End If
            End If
            End If
        End If
        If Definir_Volant Then
            NumeroPas = NumeroPas + 1
            If NumeroPas > 3 Then NumeroPas = 0
        End If
    Else
        Deplacement3 RessourceCible, PersoCible, MaisonCible, ChateauCible, Fief, NoCommentaire, Commentaires, Parametres
        If EnDeplacement Or Definir_Volant Then
            NumeroPas = NumeroPas + 1
            If NumeroPas > 3 Then NumeroPas = 0
        End If
    End If
End Sub

Private Sub Deplacement3(ByRef RessourceCible As ClsJeuRessources, _
                         ByRef PersoCible As ClsJeuPerso, _
                         ByRef MaisonCible As ClsJeuBatiment, _
                         ByRef ChateauCible As ClsJeuBatiment, _
                         ByRef Fief As ClsJeuFief, _
                         ByRef NoCommentaire As Integer, _
                         ByRef Commentaires As ClsJeuCommentaires, _
                         ByVal Parametres As ClsJeuParam)
    If DureeImmobilisation <= 0 Then
        Calculer_Longueur_Pas

        If Marcher3 Then
            If Projectile Then
                Collision_Projectile
            End If
            If Action > 99 Then
                If Action Mod 2 = 0 And Action < 200 Then
                    Action = Action + 1
                End If
                If Action < 100 Then
                    Action = 1
                End If
            ElseIf Action = 7 And Not MaisonCible.Visible Then
                Arreter
            End If
            MaisonCible.EnConstruction = False
            'Visible = True
            'RessourceCourante = -1
        Else
            If Projectile And Vivant Then Vivant = False 'Les projectiles disparaissent à leur arrivée.
            If Definir_Volant Then
                Direction = 3
                'EnDeplacement = True
                'EnDeplacement = False
            Else
                'EnDeplacement = False
                Direction = 0
            End If
            'EnDeplacement = False
            DirectionX = PositionX
            DirectionY = PositionY
            Decelerer
            If Action Mod 2 = 1 And _
               Action < 100 Then
                Action = Action - 1
            End If
            If Action = 100 Then
                If EtatFabricationObjet > 0 Then 'reprend la construction d'un objet.
                    If PeutConstruire_Objet(IndiceObjet, MaisonCible, Parametres) Then
                        Creer_Objet IndiceObjet, MaisonCible, Parametres
                    End If
                End If
            End If
            'Travailler
        End If
    Else
        DureeImmobilisation = DureeImmobilisation - 1
        If Projectile Then Vivant = False
    End If
    'Direction = Deplacement
    'Changer_Pas
End Sub

Public Sub Changer_Pas()
    'On incrémente le pas pour le prochain affichage.
    If EnDeplacement Or Definir_Volant Then
        NumeroPas = NumeroPas + 1
        If NumeroPas > 3 Then NumeroPas = 0
    End If
End Sub

Public Sub Travailler(ByRef RessourceCible As ClsJeuRessources, _
                      ByRef PersoCible As ClsJeuPerso, _
                      ByRef MaisonCible As ClsJeuBatiment, _
                      ByRef ChateauCible As ClsJeuBatiment, _
                      ByRef Fief As ClsJeuFief, _
                      ByRef NoCommentaire As Integer, _
                      ByRef Commentaires As ClsJeuCommentaires, _
                      ByVal Parametres As ClsJeuParam)
    If Vivant Then
        If Not EnDeplacement Then
        Select Case Action
            'On ne bouge pas.
            Case 0: RessourceCourante = -1
                    'MatierePremiere = 0
                    If Not Chateaux(NumeroFief).Visible Then DansSonChateau = False
                    If Not ChateauCible.Visible Then DansUnChateau = False
                    'If Not Maisons(Numero).Visible Then DansSaMaison = False
                    If Not (MaisonCible.Visible And MaisonCible.Construit) Then DansUneMaison = False
                    'Maison.EnConstruction = False
            'On récolte des ressources à l'extérieur.
            Case 2:
                RessourceCourante = RessourceCible.TypeRessource
                Recolter_Ressources RessourceCourante, Maisons(Numero), NoCommentaire, Parametres
                'Action = Parametres.Ressources_NomCompetence(RessourceCible.TypeRessource)
                DansUneMaison = False
            'On construit une maison.
            Case 6:
                If MaisonCible.Visible Then
                    RessourceCourante = -1
                    MaisonCible.EnConstruction = True
                    Batir_Maison Fief, Chateaux(NumeroFief), MaisonCible, Parametres
                Else
                    Arreter
                End If
            'Se repose.
            Case 100:
                If MaisonCible.Visible Then
                    DansSaMaison = True
                    RessourceCourante = -1
                Else
                    DansSaMaison = False
                End If
            'Travail à la maison.
            Case 101, 102, 600:
                If MaisonCible.Visible Then
                    DansSaMaison = True
                    RessourceCourante = Parametres.Batiment_RessourceFabrique(MaisonCible.TypeBatiment)
                    'MatierePremiere = Parametres.Batiment_TypeMatierePremiere(Maison.TypeBatiment)
                    'Action = Parametres.Batiment_DesignationTravail(MaisonCible.TypeBatiment)
                    If RessourceCourante >= 0 Then
                        'Action = Parametres.Ressources_NomCompetence(RessourceCourante)
                        Action = 102
                        Recolter_Ressources RessourceCourante, MaisonCible, NoCommentaire, Parametres
                    ElseIf MaisonCible.Service Then
                        Preparer_Service MaisonCible, NoCommentaire, Parametres
                    Else
                        Arreter
                    End If
                Else
                    DansSaMaison = False
                    Arreter
                End If
            'Va prendre un trésor.
            Case 204:
                Tresor_Ramasser IndiceTresor, Me
                DansSaMaison = False
                Arreter
                'Action = 0
            Case 500:
                If MaisonCible.Visible Then
                    Fabriquer_Objet Parametres, Maisons(Numero), NoCommentaire
                Else
                    DansSaMaison = False
                    Arreter
                End If
            'Rentre dans un magasin.
            Case 201:
                DansUneMaison = MaisonCible.Visible And MaisonCible.Construit
                Arreter
            'Retourne à son chateau.
            Case 202:
                DansSonChateau = Chateaux(NumeroFief).Visible
                Action = 0
            'Entre dans un chateau allié.
            Case 203:
                DansUnChateau = ChateauCible.Visible
                Action = 0
            'Vend dans un chateau.
            Case 206:
                If ChateauCible.Numero = NumeroFief Then
                    DansSonChateau = Chateaux(NumeroFief).Visible
                Else
                    DansUnChateau = ChateauCible.Visible
                End If
                Fiefs(IndiceChateau).Vendre_ObjetsSelectionnees Me, Parametres
                Action = 0
            'Vend dans un marché.
            Case 207:
                If MaisonCible.Numero = Numero Then
                    DansSaMaison = Chateaux(Numero).Visible
                Else
                    DansUneMaison = MaisonCible.Visible
                End If
                Fiefs(Persos(IndiceMaison).NumeroFief).Vendre_ObjetsSelectionnees Me, Parametres
                Action = 0
            Case Else: RessourceCourante = -1
                       'MatierePremiere = 0
                       'Maison.EnConstruction = False
        End Select
        End If
    ElseIf Not EnDeplacement Then
        'Liste des actions que le joueur peut faire quand il est mort.
        Select Case Action
            Case 0:
                If Not Chateaux(NumeroFief).Visible Then DansSonChateau = False
                If Not ChateauCible.Visible Then DansUnChateau = False
                'If Not Maisons(Numero).Visible Then DansSaMaison = False
                If Not (MaisonCible.Visible And MaisonCible.Construit) Then DansUneMaison = False
            Case 100:
                DansSaMaison = MaisonCible.Visible
                Arreter
            Case 101, 102:
                DansSaMaison = MaisonCible.Visible
                Arreter
            Case 201:
                DansUneMaison = MaisonCible.Visible And MaisonCible.Construit
                Arreter
            'Retourne à son chateau.
            Case 202:
                DansSonChateau = Chateaux(NumeroFief).Visible
                Action = 0
        End Select
    End If
End Sub

Public Sub Orienter_Direction(ByVal OrientationX As Long, OrientationY As Long)
    If OrientationY < PositionY - Longueur_Pas / 2 Then
        'Déplacement en haut.
        If OrientationX < PositionX - Longueur_Pas / 2 Then
            'Vers la gauche.
            If Abs(OrientationY - PositionY) > Abs(OrientationX - PositionX) Then
                'Déplacement en haut à gauche. (+vers le haut)
                Direction = 1
            Else
                'Déplacement en haut à gauche. (+ vers la gauche)
                Direction = 4
            End If
        ElseIf OrientationX > PositionX + Longueur_Pas / 2 Then
            'Vers la droite.
            If Abs(OrientationY - PositionY) > Abs(OrientationX - PositionX) Then
                'Déplacement en haut à droite. (+ vers le haut)
                Direction = 1
            Else
                'Déplacement en haut à droite. (+ vers la droite)
                Direction = 2
            End If
        Else
            Direction = 1
        End If
        EnDeplacement = True
    ElseIf OrientationY > PositionY + Longueur_Pas / 2 Then
        'Déplacement en bas.
        If OrientationX < PositionX - Longueur_Pas / 2 Then
            'Vers la gauche.
            If Abs(OrientationY - PositionY) > Abs(OrientationX - PositionX) Then
                'Déplacement en bas à gauche. (+vers le haut)
                Direction = 3
            Else
                'Déplacement en haut à gauche. (+ vers la gauche)
                Direction = 4
            End If
        ElseIf OrientationX > PositionX + Longueur_Pas / 2 Then
            'Vers la droite.
            If Abs(OrientationY - PositionY) > Abs(OrientationX - PositionX) Then
                'Déplacement en bas à droite. (+ vers le bas)
                Direction = 3
            Else
                'Déplacement en bas à droite. (+ vers la droite)
                Direction = 2
            End If
        Else
            Direction = 3
        End If
        EnDeplacement = True
    ElseIf OrientationX < PositionX - Longueur_Pas / 2 Then
        'Déplacement à gauche.
        Direction = 4
        EnDeplacement = True
    ElseIf OrientationX > PositionX + Longueur_Pas / 2 Then
        'Déplacement à droite.
        Direction = 2
        EnDeplacement = True
    Else
        EnDeplacement = False
    End If
End Sub

Public Function Marcher2() As Boolean
    Dim l As Double
    Dim dx As Long, dy As Long
    'Cette fonction modifie la position du personnage.
    If EnDeplacement Then
        dx = DirectionX - PositionX
        dy = DirectionY - PositionY
        l = Sqr(dx ^ 2 + dy ^ 2)
        If l >= Longueur_Pas And l > 0 Then
            PositionX = Longueur_Pas * dx / l + PositionX
            PositionY = Longueur_Pas * dy / l + PositionY
            Visible = True
        Else
            'Arreter
            PositionX = DirectionX
            PositionY = DirectionY
            EnDeplacement = False
        End If
    End If
    Marcher2 = EnDeplacement
End Function

Private Function Marcher3() As Boolean
    'Le personnage se déplace en suivant un Pathinding décrit dans la variable "Chemin".
    Dim l As Double
    Dim dx As Long, dy As Long
    If EnDeplacement Then
        If CheminDirectionX <> DirectionX Or CheminDirectionY <> DirectionY Or Not CheminComplet Then
            Calculer_Chemin Monde
            CheminDirectionX = DirectionX
            CheminDirectionY = DirectionY
        End If
        dx = CheminX(1) - PositionX
        dy = CheminY(1) - PositionY
        l = Sqr(dx ^ 2 + dy ^ 2)
        If l >= Longueur_Pas And l > 0 Then
            PositionX = Longueur_Pas * dx / l + PositionX
            PositionY = Longueur_Pas * dy / l + PositionY
            Visible = True
        Else
            'Arreter
            If CheminX.Count > 1 Then
                CheminX.Remove 1
                CheminY.Remove 1
                dx = CheminX(1) - PositionX
                dy = CheminY(1) - PositionY
                l = Sqr(dx ^ 2 + dy ^ 2)
                'Change de chemin et d'orientation.
                Orienter_Direction CheminX(1), CheminY(1)
                If l = 0 Then
                    PositionX = DirectionX
                    PositionY = DirectionY
                    EnDeplacement = False
                Else
                    PositionX = Longueur_Pas * dx / l + PositionX
                    PositionY = Longueur_Pas * dy / l + PositionY
                    Visible = True
                End If
            Else
                PositionX = DirectionX
                PositionY = DirectionY
                EnDeplacement = False
            End If
        End If
    End If
    Marcher3 = EnDeplacement
End Function

Private Sub Calculer_Chemin(ByVal Monde As ClsJeuMonde)
    Dim i As Integer
    Dim l1 As Double
    Dim l2 As Double
    Dim dx As Long, dy As Long
    Dim TempX As Long
    Dim TempY As Long
    Dim TempV As Double 'Enregistre la vitesse des cases adjacentes.
    Dim TempR As Double 'Enregistre le meilleurs temps.
    Dim TempRi As Integer 'Enregistre l'indice du dernier temps.
    Dim TempPi As Integer 'Enregistre l'indice du temps précédent.
    Dim NbEtapesMax As Integer

    Dim IntervalX As Integer
    Dim IntervalY As Integer
    Dim Quitter As Boolean
    
    'If Not EnDeplacement Then
    Collection_Vider CheminX
    Collection_Vider CheminY
    IntervalX = Monde.LargeurCase ' * 2 'LongueurPas * 2  '* 2.5
    IntervalY = Monde.LargeurCase ' * 2 'LongueurPas * 2  '* 2.5
    TempCheminX = 0
    TempCheminY = 0
    'Set TempChemin = Nothing
    If Not Definir_Volant And Vivant And Jeu.Regle_Pathfinding > 0 Then
        TempPi = -1
        If Action = 300 And Persos(IndicePerso).EnDeplacement Then
            NbEtapesMax = 2 '8
            CheminComplet = False
        Else
            NbEtapesMax = 1024
            CheminComplet = True
        End If
        While (Abs(TempX - DirectionX) > Monde.LargeurCase Or Abs(TempY - DirectionY) > Monde.LargeurCase) And CheminX.Count < NbEtapesMax And Not Quitter 'Or Cheminx.Count = 0
            'Regarde plusieurs cotés.
            If CheminX.Count = 0 Then
                dx = PositionX - DirectionX
                dy = PositionY - DirectionY
            Else
                dx = CheminX(CheminX.Count) - DirectionX
                dy = CheminY(CheminY.Count) - DirectionY
            End If
            l1 = Sqr(dx ^ 2 + dy ^ 2)
            For i = 0 To 7
                If CheminX.Count = 0 Then
                    TempCheminX = PositionX
                    TempCheminY = PositionY
                Else
                    TempCheminX = CheminX(CheminX.Count)
                    TempCheminY = CheminY(CheminY.Count)
                End If
                Select Case i
                Case 0: 'En haut à gauche.
                        TempCheminX = TempCheminX - IntervalX * 0.7
                        TempCheminY = TempCheminY - IntervalY * 0.7
                Case 1: 'En haut.
                        TempCheminX = TempCheminX
                        TempCheminY = TempCheminY - IntervalY
                Case 2: 'En haut à droite.
                        TempCheminX = TempCheminX + IntervalX * 0.7
                        TempCheminY = TempCheminY - IntervalY * 0.7
                Case 3: 'A gauche.
                        TempCheminX = TempCheminX - IntervalX
                        TempCheminY = TempCheminY
                Case 4: 'A droite.
                        TempCheminX = TempCheminX + IntervalX
                        TempCheminY = TempCheminY
                Case 5: 'En bas à gauche.
                        TempCheminX = TempCheminX - IntervalX * 0.7
                        TempCheminY = TempCheminY + IntervalY * 0.7
                Case 6: 'En bas.
                        TempCheminX = TempCheminX
                        TempCheminY = TempCheminY + IntervalY
                Case 7: 'En bas à droite.
                        TempCheminX = TempCheminX + IntervalX * 0.7
                        TempCheminY = TempCheminY + IntervalY * 0.7
                End Select
                TempV = IntervalX / Monde.CoefVitesseSimple(TempCheminX, TempCheminY + Hauteur - AffTerrain.Hauteur)
                'TempV(i) = IntervalX * (1 - Monde.CoefVitesseSimple(TempX, TempY + Hauteur - AffTerrain.Hauteur)) * 5
                
                dx = TempCheminX - DirectionX
                dy = TempCheminY - DirectionY
                l2 = Sqr(dx ^ 2 + dy ^ 2)
                TempV = TempV + (l2 - l1)
                If i = 0 Or TempV < TempR Then
                    TempR = TempV
                    TempRi = i
                    TempX = TempCheminX
                    TempY = TempCheminY
                End If
            Next i
            If CheminX.Count = 0 Then
                CheminX.Add TempX
                CheminY.Add TempY
            Else
                If TempPi = TempRi Then
                    'Si l'on continue dans la même direction, inutil d'ajouter une étape supplémentaire.
                    CheminX.Remove CheminX.Count
                    CheminY.Remove CheminY.Count
                    CheminX.Add TempX
                    CheminY.Add TempY
                ElseIf CheminX.Count >= 2 Then
                    If TempX = CheminX(CheminX.Count - 1) And TempY = CheminY(CheminY.Count - 1) Then
                        'On arrête si l'on tourne en rond.
                        Quitter = True
                    Else
                        CheminX.Add TempX
                        CheminY.Add TempY
                        TempPi = TempRi
                    End If
                ElseIf TempX = CheminX(CheminX.Count) And TempY = CheminY(CheminY.Count) Then
                    'On arrête si l'on tourne en rond.
                    Quitter = True
                Else
                    CheminX.Add TempX
                    CheminY.Add TempY
                    TempPi = TempRi
                End If
            End If
        Wend
    Else
        CheminComplet = True
    End If
    'Ajoute la dernière étape
    If CheminX.Count = 0 Then
        CheminX.Add DirectionX
        CheminY.Add DirectionY
    Else
        CheminX.Remove CheminX.Count
        CheminY.Remove CheminY.Count
        CheminX.Add DirectionX
        CheminY.Add DirectionY
        If Jeu.Regle_Pathfinding > 1 Then
            Do
            Loop Until Not Chemin_Supprimer_EtapesSuperflues Or CheminX.Count < 3
        End If
    End If
    'Supprime une étape sur 2.
    'For i = CheminX.Count - 1 To 2 Step -2
    '    CheminX.Remove i
    '    CheminY.Remove i
    'Next i
End Sub

Private Function Chemin_Supprimer_EtapesSuperflues() As Boolean
    'Cette fonction supprime les étapes superflues où le terrain ne change pas.
    'Pour cela, on parcours chaque étape et analyse le terrain.
    Dim i As Integer
    Dim dx As Long, dy As Long
    Dim l As Double
    Dim l2 As Integer
    Dim TempT1 As Single 'Variables temporaires mémorisant le type de terrain rencontré.
    Dim TempT2 As Single
    Dim SupprimerEtape As Boolean 'Si cette variable passe à vrai, on supprime l'étape en question.
    l2 = LongueurPas ^ 2
    For i = 2 To CheminX.Count - 1
        SupprimerEtape = True
        TempCheminX = CheminX(i - 1)
        TempCheminY = CheminY(i - 1)
        TempT1 = Monde.CoefVitesseSimple(TempCheminX, TempCheminY + Hauteur - AffTerrain.Hauteur)
        Do
            dx = CheminX(i + 1) - TempCheminX
            dy = CheminY(i + 1) - TempCheminY
            l = Sqr(dx ^ 2 + dy ^ 2)
            If l > 0 Then
                TempCheminX = LongueurPas * dx / l + TempCheminX
                TempCheminY = LongueurPas * dy / l + TempCheminY
                TempT2 = Monde.CoefVitesseSimple(TempCheminX, TempCheminY + Hauteur - AffTerrain.Hauteur)
                'If TempT1 <> TempT2 Then
                If TempT1 <> TempT2 Then
                    SupprimerEtape = False
                End If
            End If
            'Debug.Print "X=" & TempCheminX & ", Y=" & TempCheminY & "   " & (TempCheminX - CheminX(i + 1)) ^ 2 + (TempCheminY - CheminY(i + 1)) ^ 2
        Loop Until (TempCheminX - CheminX(i + 1)) ^ 2 + (TempCheminY - CheminY(i + 1)) ^ 2 < l2 Or Not SupprimerEtape
        If SupprimerEtape Then
            CheminX.Remove i
            CheminY.Remove i
            i = i - 1
            Chemin_Supprimer_EtapesSuperflues = True
        End If
        If i >= CheminX.Count - 1 Or (Not CheminComplet And i >= CheminX.Count - 2) Then
            Exit Function
        End If
        'If i > CheminX.Count - 2 Then
    Next i
End Function

Private Sub Collision_Projectile()
    Dim Temp As Integer
    'Attaque une cible si elle passe dessus.
    Temp = Obstacle(Numero, PositionX, PositionY, Largeur, Hauteur, True)
    If Temp >= 9 And Temp <= 11 Then
        If BonusSurfaceAttaque > 0 Then
            Select Case Temp
            Case 9:
                If ProjectileCible = 0 Then ' Or Not Kamikaze Then
                    Perso_Attaquer_Surface Me, Persos(IndiceObstacle).PositionX, Persos(IndiceObstacle).PositionY, BonusSurfaceAttaque, ControleJoueur
                    If Kamikaze Or Persos(NumeroChef).Attaque = 0 Then Vivant = False
                End If
            Case 10:
                If ProjectileCible = 1 Then ' Or Not Kamikaze Then
                    Perso_Attaquer_Surface Me, Maisons(IndiceObstacle).PositionX, Maisons(IndiceObstacle).PositionY, BonusSurfaceAttaque, ControleJoueur
                    Vivant = False
                End If
            Case 11:
                If ProjectileCible = 2 Then '  Or Not Kamikaze Then
                    Perso_Attaquer_Surface Me, Chateaux(IndiceObstacle).PositionX, Chateaux(IndiceObstacle).PositionY, BonusSurfaceAttaque, ControleJoueur, IndiceObstacle
                    Vivant = False
                End If
            End Select
        Else
            Select Case Temp
            Case 9:
                If ProjectileCible = 0 Then ' Or Not Kamikaze Then
                    Perso_Attaquer_Perso Me, Persos(IndiceObstacle)
                    If Kamikaze Or Persos(NumeroChef).Attaque = 0 Then Vivant = False
                End If
            Case 10:
                If ProjectileCible = 1 Then ' Or Not Kamikaze Then
                    Perso_Attaquer_Batiment Me, Maisons(IndiceObstacle), Persos(IndiceObstacle)
                    Vivant = False
                End If
            Case 11:
                If ProjectileCible = 2 Then ' Or Not Kamikaze Then
                    Perso_Attaquer_Batiment Me, Chateaux(IndiceObstacle), Persos(IndiceObstacle), IndiceObstacle
                    Vivant = False
                End If
            End Select
        End If
        'Vivant = False
    End If
End Sub

Public Sub RecuperationLente(ByVal Maison As ClsJeuBatiment, _
                             ByVal Fief As ClsJeuFief, _
                             ByVal Parametres As ClsJeuParam)
    If Vivant Then
        If EnDeplacement Then
            'Le personnage se fatigue en marchant.
            Definir_Energie = Energie - VitesseFatigueDeplacement
            Definir_Moral = Moral - VitesseDemoralisationDeplacement
            If Pietinement > 0 Then 'Le personnage écrase en marchant.
                Perso_Pietinement Me
            End If
        Else
            Select Case Action
            Case 0:
                'On se repose dehors.
                Definir_Vie = Vie + VitesseRegenerationVieEnAttente * Definir_MaxVie
                Definir_Energie = Energie + VitesseRegenerationEnergieEnAttente * Definir_MaxEnergie
                Definir_Magie = Magie + VitesseRegenerationMagieEnAttente * Definir_MaxMagie
                Definir_Moral = Moral + VitesseRegenerationMoralEnAttente * Definir_MaxMoral
            Case 2:
                'Se fatigue en récoltant des ressources à l'extérieur.
                Definir_Energie = Energie - VitesseFatigueTravail * (NivRessources(RessourceCourante) + Parametres.PersosMaxRessources)
                Definir_Moral = Moral - VitesseDemoralisationTravail * (NivRessources(RessourceCourante) + Parametres.PersosMaxRessources)
            Case 6:
                'Le personnage se fatigue en construisant sa maison.
                Definir_Energie = Energie - VitesseFatigueConstruction * CaracSpeciales(0)
                Definir_Moral = Moral - VitesseDemoralisationConstruction * CaracSpeciales(0)
            Case 100:
                'On se repose à la maison.
                Definir_Vie = Vie + VitesseRegenerationVieEnRepos * Definir_MaxVie
                Definir_Energie = Energie + VitesseRegenerationEnergieEnRepos * Definir_MaxEnergie
                Definir_Magie = Magie + VitesseRegenerationMagieEnRepos * Definir_MaxMagie
                Definir_Moral = Moral + VitesseRegenerationMoralEnRepos * Definir_MaxMoral
            Case 102:
                If RessourceCourante >= 0 Then
                    'Se fatigue en récoltant des ressources à la maison.
                    Definir_Energie = Energie - VitesseFatigueTravail * (NivRessources(RessourceCourante) + Parametres.PersosMaxRessources)
                    Definir_Moral = Moral - VitesseDemoralisationTravail * (NivRessources(RessourceCourante) + Parametres.PersosMaxRessources)
                Else
                    'Se fatigue en Fabriquant des services.
                    Definir_Energie = Energie - VitesseFatigueTravail * CaracServices(ServiceCourant - 1)
                    Definir_Moral = Moral - VitesseDemoralisationTravail * CaracServices(ServiceCourant - 1)
                End If
            Case 500:
                'Le personnage se fatigue en fabriquant un objet.
                Definir_Energie = Energie - VitesseFatigueFabrication * CaracObjets(Parametres.Objet_Competence(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, IndiceObjet)))
                Definir_Moral = Moral - VitesseDemoralisationFabrication * CaracObjets(Parametres.Objet_Competence(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, IndiceObjet)))
            End Select
        End If
        'Le personnage récupère.
        If RecuperationVie > 0 Then
            If RecuperationVie >= VitesseRecuperationVie * Definir_MaxVie Then
                Definir_Vie = Vie + VitesseRecuperationVie * Definir_MaxVie
                RecuperationVie = RecuperationVie - VitesseRecuperationVie * Definir_MaxVie
            Else
                Definir_Vie = Vie + RecuperationVie
                RecuperationVie = 0
            End If
        End If
        If RecuperationEnergie > 0 Then
            If RecuperationEnergie >= VitesseRecuperationEnergie * Definir_MaxEnergie Then
                Definir_Energie = Energie + VitesseRecuperationEnergie * Definir_MaxEnergie
                RecuperationEnergie = RecuperationEnergie - VitesseRecuperationEnergie * Definir_MaxEnergie
            Else
                Definir_Energie = Energie + RecuperationEnergie
                RecuperationEnergie = 0
            End If
        End If
        If RecuperationMagie > 0 Then
            If RecuperationMagie >= VitesseRecuperationMagie * Definir_MaxMagie Then
                Definir_Magie = Magie + VitesseRecuperationMagie * Definir_MaxMagie
                RecuperationMagie = RecuperationMagie - VitesseRecuperationMagie * Definir_MaxMagie
            Else
                Definir_Magie = Magie + RecuperationMagie
                RecuperationMagie = 0
            End If
        End If
        If RecuperationMoral > 0 Then
            If RecuperationMoral >= VitesseRecuperationMoral * Definir_MaxMoral Then
                Definir_Moral = Moral + VitesseRecuperationMoral * Definir_MaxMoral
                RecuperationMoral = RecuperationMoral - VitesseRecuperationMoral * Definir_MaxMoral
            Else
                Definir_Moral = Moral + RecuperationMoral
                RecuperationMoral = 0
            End If
        End If
        If Definir_Attaque < Definir_MaxAttaque Then
            Definir_Energie = Energie - VitesseFatigueCombat * Vitalite
            Definir_Moral = Moral - VitesseDemoralisationCombat * Vitalite
        End If
        If Definir_Defense < Definir_MaxDefense Then
            Definir_Energie = Energie - VitesseFatigueCombat * Vitalite
            Definir_Moral = Moral - VitesseDemoralisationCombat * Vitalite
        End If
        
        Definir_Vie = Vie + BonusVitesseRegenerationVie * Definir_MaxVie 'Régénération proportionnelle.
        Definir_Energie = Energie + BonusVitesseRegenerationEnergie * Definir_MaxEnergie 'Régénération proportionnelle.
        Definir_Magie = Magie + BonusVitesseRegenerationMagie * Definir_MaxMagie  'Régénération proportionnelle.
        'Definir_Magie = Magie + VitesseRegenerationMagie * BonusMaxMagie / MaxMagie 'Régénération linéaire.
        Definir_Moral = Moral + BonusVitesseRegenerationMoral * Definir_MaxMoral 'Régénération proportionnelle.
        
        'Génération de ressources.
        Generer_Ressource
        Gestion_Effets_Temporaires
        Gestion_Effets_Recharge
        
        'Poison.
        If Poison >= 1 Then
            Definir_Vie = Vie - Poison
            AffDegat.Ajouter_Degat PositionX + Largeur / 2, PositionY, Poison
            If Not Vivant Then 'Le personnage est tué par celui qui lui a lancé le poison.
                If PoisonLanceur >= 0 Then
                    Perso_Tuer_Perso Persos(PoisonLanceur), Me
                End If
            End If
        End If
        
    'Phase de ressurection.
    ElseIf DansSonChateau Then
        If Chateaux(NumeroFief).Visible Then
            'If (Fief.NombreCitoyens - Fief.NombreMorts + 1) = 0 Or (Fief.NombreCitoyens / (Fief.NombreCitoyens - Fief.NombreMorts + 1)) = 0 Then MsgBox "Erreur dans la fonction RécupérationLente de la Class JeuPerso" '%%%
            If ResurrectionNombre > 0 Or ResurrectionNombre = -1 Then
                If Fief.NombreCitoyens - Fief.NombreMorts = 0 Then
                    DureeRessurection = DureeRessurection - VitesseResurrection / Fief.NombreCitoyens
                Else
                    DureeRessurection = DureeRessurection - VitesseResurrection / (Fief.NombreCitoyens / (Fief.NombreCitoyens - Fief.NombreMorts))
                End If
                If DureeRessurection <= 0 Then
                    Vivant = True
                    If ResurrectionNombre <> -1 Then
                        ResurrectionNombre = ResurrectionNombre - 1
                    End If
                Else
                    Action = 400
                End If
            Else
                Action = 0
            End If
        Else
            'Le chateau a été détruit.
            'Arreter
            Action = 0
            DansSonChateau = False
        End If
    ElseIf Invocation Then
        If DansSaMaison And Maisons(Numero).Visible Then
            'Les invocs peuvent ressuciter dans leur maison.
            DureeRessurection = DureeRessurection - VitesseResurrection
            If DureeRessurection <= 0 Then Vivant = True
        End If
    ElseIf Jeu.Confrontation Then 'En mode confrontation les persos ressuscitent à un endroit aléatoire de la carte.
        'If Not Invocation Then
        Vivant = True
        PositionX = Rnd * (Monde.Largeur * Monde.LargeurCase - Largeur)
        PositionY = Rnd * (Monde.Hauteur * Monde.HauteurCase - Hauteur)
        'End If
    End If
End Sub

Public Sub RecuperationRapide()
    If Vivant Then
        'Le personnage récupère.
        If MaxAttaque = 0 Then MsgBox "Erreur dans la fonction RécupérationRapide de la Class JeuPerso" '%%%
        If Definir_Attaque < Definir_MaxAttaque Then
            Definir_Attaque = Definir_Attaque + VitesseRecuperationAttaque * Vitalite * BonusMaxAttaque / MaxAttaque
        End If
        If Definir_Defense < Definir_MaxDefense Then
            Definir_Defense = Definir_Defense + VitesseRecuperationDefense * Vitalite * BonusMaxDefense / MaxDefense
        End If
        'Rechargement de l'attaque.
        If DureeRechargeAttaque > 0 Then
            DureeRechargeAttaque = DureeRechargeAttaque - 1
            If DureeRechargeAttaque < 0 Then DureeRechargeAttaque = 0
        End If
    End If
End Sub

Private Sub Generer_Ressource()
    Dim i As Integer
    If GenerationRessources Then
        Definir_Argent = Definir_Argent + GainArgent * VitesseGenerationRessources
        For i = 0 To UBound(GainRessource())
            Definir_ressources(i) = Definir_ressources(i) + GainRessource(i) * VitesseGenerationRessources
        Next i
    End If
End Sub

Public Sub Gestion_Effets_Temporaires()
    Dim i As Integer
    For i = 1 To Nombre_EffetsTemp
        If Definir_EffetTemp_Duree(i) <> -1 Then
            Definir_EffetTemp_Duree(i) = Definir_EffetTemp_Duree(i) - 1
            If Definir_EffetTemp_Duree(i) <= 0 Then
                EffetTemp_Effacer i
                Actualiser_Bonus
                Exit For
        'ElseIf Visible Then
        '    With Parametres
        '    If .Objet_EffetAttaque(Definir_EffetTemp_Numero(i)) >= 0 Then
        '        Activer_Effet .Objet_EffetAttaque(Definir_EffetTemp_Numero(i)), PositionX, PositionY, Largeur, Hauteur, Parametres
        '    End If
        '    End With
            End If
        End If
    Next i
End Sub
Public Sub Gestion_Effets_Recharge()
    Dim i As Integer
    For i = 1 To Nombre_EffetsRecharge
        Definir_EffetsRecharge_Duree(i) = Definir_EffetsRecharge_Duree(i) - 1
        If Definir_EffetsRecharge_Duree(i) <= 0 Then
            EffetRecharge_Effacer i
            Exit For
        End If
    Next i
End Sub

'Public Property Let Definir_Action(ByVal Valeur As Integer)
'    Select Case Action
'        Case "En attente": Definir_Action = 0
'        Case "Se déplacer": Definir_Action = 1
'        Case "Récolter ressources": Definir_Action = 2
'        Case "Aller à": Definir_Action = 3
'        Case "Construire": Definir_Action = 6
'        Case "Aller construire": Definir_Action = 7
'        Case "En repos": Definir_Action = 100
'        Case "Rentrer à la maison": Definir_Action = 101
'        Case "Au travail": Definir_Action = 102
'        Case "Entrer dans": Definir_Action = 201
'        Case "Retourner château": Definir_Action = 202
'        Case "Aller château": Definir_Action = 203
'        Case "Aller trésor": Definir_Action = 204
'        Case "Donner": Definir_Action = 205
'        Case "Attaquer": Definir_Action = 300
'        Case "Détruire maison": Definir_Action = 301
'        Case "Détruire château": Definir_Action = 302
'        Case "Ressusciter": Definir_Action = 400
'        Case "Fabriquer": Definir_Action = 500
'        Case "Préparer Service": Definir_Action = 600
'        Case Else: Definir_Action = 102
'    End Select
'End Property

'Public Property Get Definir_Action() As String
'    Select Case Action
'        Case 0: Definir_Action = "En attente"
'        Case 1: Definir_Action = "Se déplacer"
'        Case 2: Definir_Action = "Récolter ressources"
'        Case 3: Definir_Action = "Aller à"
'        Case 6: Definir_Action = "Construire"
'        Case 7: Definir_Action = "Aller construire"
'        Case 100: Definir_Action = "En repos"
'        Case 101: Definir_Action = "Rentrer à la maison"
'        Case 102: Definir_Action = "Au travail"
'        Case 201: Definir_Action = "Entrer dans"
'        Case 202: Definir_Action = "Retourner château"
'        Case 203: Definir_Action = "Aller château"
'        Case 204: Definir_Action = "Aller trésor"
'        Case 205: Definir_Action = "Donner"
'        Case 300: Definir_Action = "Attaquer"
'        Case 301: Definir_Action = "Détruire maison"
'        Case 302: Definir_Action = "Détruire château"
'        Case 400: Definir_Action = "Ressusciter"
'        Case 500: Definir_Action = "Fabriquer"
'        Case 600: Definir_Action = "Préparer Service"
'        Case Else: Definir_Action = "Au travail"
'    End Select
'End Property

' ************************************************************************************
' *                                Fonctions de travail                              *
' ************************************************************************************

Public Sub Recolter_Ressources(ByVal Index As Long, _
                               ByVal Maison As ClsJeuBatiment, _
                               ByRef NoCommentaire As Integer, _
                               ByVal Parametres As ClsJeuParam)
    Dim i As Integer
    Dim MatierePremiere As Integer
    Dim QuantiteMatierePremiere As Double
    Dim VitesseRecolte As Double
    
    MatierePremiere = Parametres.Ressources_TypeMatierePremiere(Index)
    QuantiteMatierePremiere = Parametres.Ressources_QuantiteMatierePremiere(Index)
    'VitesseRecolte = VitesseAction * Me.Vitalite * MaxRessources(index)
    VitesseRecolte = 1 / VitesseAction * Me.Vitalite * (Definir_Niveau_CompetenceRessources(Index) + Parametres.PersosMaxRessources)
    'Debug.Print Definir_Niveau_CompetenceRessources(index) & "+" & Parametres.PersosMaxRessources
    'Debug.Print VitesseRecolte
    
    If Index >= 0 Then
        If Definir_ressources(Index) < Definir_MaxRessources(Index) Then
            If MatierePremiere >= 0 Then 'Il faut une autre ressource comme matière première.
                If Definir_ressources(MatierePremiere) >= Parametres.Ressources_VitesseExtraction(Index) * VitesseRecolte * QuantiteMatierePremiere Then
                    'On retranche la matière première.
                    Definir_ressources(MatierePremiere) = Me.Definir_ressources(MatierePremiere) - Parametres.Ressources_VitesseExtraction(Index) * VitesseRecolte * QuantiteMatierePremiere
                    'On augmente la ressource.
                    Definir_ressources(Index) = Definir_ressources(Index) + Parametres.Ressources_VitesseExtraction(Index) * VitesseRecolte
                Else
                    Action = 100
                    NoCommentaire = 1
                End If
            Else
                Definir_ressources(Index) = Definir_ressources(Index) + Parametres.Ressources_VitesseExtraction(Index) * VitesseRecolte
            End If
            'Laisse une pause pour que les intelligences artificielles se réorganisent.
            If Definir_ressources(Index) = Definir_MaxRessources(Index) Then
                If IA Then Arreter
            End If
        'Si on a plus de place, on stock dans la maison.
        ElseIf DansSaMaison And _
               Maison.Definir_Stock < Maison.MaxStock Then
            If MatierePremiere >= 0 Then 'Il faut une autre ressource comme matière première.
                If Definir_ressources(MatierePremiere) >= Parametres.Ressources_VitesseExtraction(Index) * VitesseRecolte * QuantiteMatierePremiere Then
                    'On retranche la matière première.
                    Definir_ressources(MatierePremiere) = Definir_ressources(MatierePremiere) - Parametres.Ressources_VitesseExtraction(Index) * VitesseRecolte * QuantiteMatierePremiere
                    'On augmente la ressource.
                    Maison.Definir_Stock = Maison.Definir_Stock + Parametres.Ressources_VitesseExtraction(Index) * VitesseRecolte
                Else
                    Action = 100
                    NoCommentaire = 1
                End If
            Else
                Maison.Definir_Stock = Maison.Definir_Stock + Parametres.Ressources_VitesseExtraction(Index) * VitesseRecolte
            End If
        Else
            Arreter
            NoCommentaire = 0
        End If
    End If
    If Action <> 0 And Action <> 100 Then
        'Il gagne de l'expérience avec les objets qu'il utilise.
        If Bonus_CompetenceRessources(Index) Then
            'Ajuste le gain d'expérience en fonction des bonus éventuels.
            VitesseRecolte = VitesseRecolte * (NivRessources(Index) + Parametres.PersosMaxRessources) / (BonusNivRessources(Index) + Parametres.PersosMaxRessources)
            For i = 0 To UBound(ObjetsEquipes())
                If ObjetsEquipes(i).Param_BonusCompetencesRessources(Index) > 0 Then
                        'Gestion_Niveau CaracObjets(ObjetsEquipes(i).Numero), _
                                       Parametres.Ressources_VitesseExtraction(index) * VitesseRecolte, _
                                       ExpObjets(ObjetsEquipes(i).Numero), NivObjets(ObjetsEquipes(i).Numero), ExpNivObjets(ObjetsEquipes(i).Numero), ExpNivPlusObjets(ObjetsEquipes(i).Numero), 8
                        Gestion_Niveau CaracObjets(ObjetsEquipes(i).Competence), _
                                       VitesseRecolte, _
                                       ExpObjets(ObjetsEquipes(i).Competence), NivObjets(ObjetsEquipes(i).Competence), ExpNivObjets(ObjetsEquipes(i).Competence), ExpNivPlusObjets(ObjetsEquipes(i).Competence), 8
                End If
            Next i
        End If
        'Il gagne de l'expérience dans sa compétence de ressource.
        Gestion_Niveau MaxRessources(Index), _
                       VitesseRecolte, _
                       ExpRessources(Index), NivRessources(Index), ExpNivRessources(Index), ExpNivPlusRessources(Index), 6
    End If
End Sub

Public Sub Preparer_Service(ByVal Maison As ClsJeuBatiment, _
                            ByRef NoCommentaire As Integer, _
                            ByVal Parametres As ClsJeuParam)
    Dim MatierePremiere As Integer
    Dim VitesseCreation As Double
    Dim i As Integer
    
    ServiceCourant = Maison.Service
    MatierePremiere = Parametres.Service_MatierePremiere(ServiceCourant)
    VitesseCreation = 1 / VitesseAction * Me.Vitalite * _
                      Definir_Niveau_CompetenceServices(ServiceCourant - 1)
    
    'On stock le service dans la maison.
    If Maison.Definir_Stock < Maison.MaxStock Then
        If MatierePremiere >= 0 Then 'Il faut une autre ressource comme matière première.
            If Definir_ressources(MatierePremiere) >= Parametres.Service_VitesseCreation(ServiceCourant) * VitesseCreation _
                                                      * Parametres.Service_QuantiteMatierePremiere(ServiceCourant) Then
                'On retranche la matière première.
                Definir_ressources(MatierePremiere) = Me.Definir_ressources(MatierePremiere) - Parametres.Service_VitesseCreation(ServiceCourant) * VitesseCreation _
                                                      * Parametres.Service_QuantiteMatierePremiere(ServiceCourant)
                'On augmente la service.
                Maison.Definir_Stock = Maison.Definir_Stock + Parametres.Service_VitesseCreation(ServiceCourant) * VitesseCreation
            Else
                Action = 100
                NoCommentaire = 1
            End If
        Else
            Maison.Definir_Stock = Maison.Definir_Stock + Parametres.Service_VitesseCreation(ServiceCourant) * VitesseCreation
        End If
        Action = 600
    Else
        Arreter
        NoCommentaire = 0
    End If
    If Action <> 0 And Action <> 100 Then
        'Il gagne de l'expérience dans sa compétence de service.
        Gestion_Niveau CaracServices(ServiceCourant - 1), _
                       Parametres.Service_VitesseCreation(ServiceCourant) * VitesseCreation * Parametres.Service_CoefExp(ServiceCourant), _
                       ExpServices(ServiceCourant - 1), NivServices(ServiceCourant - 1), ExpNivServices(ServiceCourant - 1), ExpNivPlusServices(ServiceCourant - 1), 7
        'Debug.Print CaracServices(ServiceCourant - 1) & ", " & ExpServices(ServiceCourant - 1) & ", " & NivServices(ServiceCourant - 1) & ", " & ExpNivServices(ServiceCourant - 1) & ", " & ExpNivPlusServices(ServiceCourant - 1)
        'Il gagne de l'expérience avec les objets qu'il utilise.
        If Bonus_CompetenceServices(ServiceCourant - 1) Then
            For i = 0 To UBound(ObjetsEquipes())
                If ObjetsEquipes(i).Param_BonusCompetencesServices(ServiceCourant) > 0 Then
                    Gestion_Niveau CaracObjets(ObjetsEquipes(i).Competence), _
                                   Parametres.Service_VitesseCreation(ServiceCourant) * VitesseCreation * Parametres.Service_CoefExp(ServiceCourant), _
                                   ExpObjets(ObjetsEquipes(i).Competence), NivObjets(ObjetsEquipes(i).Competence), ExpNivObjets(ObjetsEquipes(i).Competence), ExpNivPlusObjets(ObjetsEquipes(i).Competence), 7
                End If
            Next i
        End If
    End If
End Sub

Public Sub Manger_Ressources(ByVal Index As Long, _
                             ByVal Parametres As ClsJeuParam)
    'Le personnage mange une unité de ressource et récupère ce que la ressource lui redonne.
    If Ressources(Index) >= 1 Then
        Ressources(Index) = Ressources(Index) - 1
        RecuperationVie = RecuperationVie + Parametres.Ressources_GainVie(Index) * Manger_Ressources_Coef(0)
        RecuperationEnergie = RecuperationEnergie + Parametres.Ressources_GainEnergie(Index) * Manger_Ressources_Coef(1)
        RecuperationMagie = RecuperationMagie + Parametres.Ressources_GainMagie(Index) * Manger_Ressources_Coef(2)
        RecuperationMoral = RecuperationMoral + Parametres.Ressources_GainMoral(Index) * Manger_Ressources_Coef(3)
    End If
End Sub

Public Function Manger_Ressources_Coef(ByVal NoCarac As Integer) As Double
    Select Case NoCarac
    Case 0:
        If MaxVie = 0 Then MsgBox "Erreur dans la fonction Manger_Ressources_Coef de la Class JeuPerso" '%%%
        Manger_Ressources_Coef = BonusMaxVie / MaxVie
    Case 1:
        If MaxEnergie = 0 Then MsgBox "Erreur dans la fonction Manger_Ressources_Coef de la Class JeuPerso" '%%%
        Manger_Ressources_Coef = BonusMaxEnergie / MaxEnergie
    Case 2:
        If MaxMagie = 0 Then MsgBox "Erreur dans la fonction Manger_Ressources_Coef de la Class JeuPerso" '%%%
        Manger_Ressources_Coef = BonusMaxMagie / MaxMagie
    Case 3:
        If MaxMoral = 0 Then MsgBox "Erreur dans la fonction Manger_Ressources_Coef de la Class JeuPerso" '%%%
        Manger_Ressources_Coef = BonusMaxMoral / MaxMoral
    End Select
End Function

Public Sub Batir_Maison(ByVal Fief As ClsJeuFief, _
                        ByRef Chateau As ClsJeuBatiment, _
                        ByRef Maison As ClsJeuBatiment, _
                        ByVal Parametres As ClsJeuParam)
    'Le maison se constuit petit à petit.
    Dim i As Integer
    Dim VitesseBatir As Double
    VitesseBatir = VitesseConstruction * Vitalite _
                * (Definir_Niveau_CompetenceSpeciales(0) + Parametres.PersosNiveauSpecialesDepart) / Parametres.PersosNiveauSpecialesDepart * _
                Parametres.Batiment_CoefVitesseConstruction(Maison.TypeBatiment)
    Maison.Definir_Vie = Maison.Definir_Vie + VitesseBatir
    'Il gagne de l'expérience dans sa compétence contruire.
    Gestion_Niveau CaracSpeciales(0), _
                   VitesseBatir * CoefExpCompConstruction * CoefExpCompSpeciales, _
                   ExpSpeciales(0), NivSpeciales(0), ExpNivSpeciales(0), ExpNivPlusSpeciales(0), 9
    'Il gagne de l'expérience avec les objets qu'il utilise.
    If Bonus_CompetenceSpeciales(0) Then
        For i = 0 To UBound(ObjetsEquipes())
            If ObjetsEquipes(i).Param_BonusCompetencesSpeciales(0) > 0 Then
                Gestion_Niveau CaracObjets(ObjetsEquipes(i).Competence), _
                               VitesseBatir * CoefExpCompConstruction * CoefExpCompSpeciales, _
                               ExpObjets(ObjetsEquipes(i).Competence), NivObjets(ObjetsEquipes(i).Competence), ExpNivObjets(ObjetsEquipes(i).Competence), ExpNivPlusObjets(ObjetsEquipes(i).Competence), 8
            End If
        Next i
    End If
    If Maison.Construit And Maison.Definir_Vie = Maison.MaxVie Then
        If Maison.Numero = Numero Then
            Action = 101
        Else
            DansUneMaison = True
            Arreter
        End If
        'Calculer_Apparence Maisons(Numero), Parametres
        'Si c'est un batiment récent, on change d'époque.
        If Parametres.Batiment_NoEpoque(Maison.TypeBatiment) > Fief.Epoque Then
            'Fief.Epoque = Parametres.Batiment_NoEpoque(Maison.TypeBatiment)
            Fief.Changer_Epoque Chateau, Parametres
        End If
    Else
        Action = 6
    End If
End Sub

Public Function PeutConstruire_Maison(ByVal IndiceMaison As Long, _
                                      ByVal Parametres As ClsJeuParam) As Boolean
    'Renvoie vrai si le personnage a assez d'argent et de ressources pour construire ce bâtiment.
    Dim i As Long
    If IndiceMaison >= 0 Then
        PeutConstruire_Maison = Definir_Argent >= Parametres.Batiment_PrixArgent(IndiceMaison)
        If PeutConstruire_Maison Then
            For i = 0 To Parametres.NombreRessources - 1
                'If IndiceMaison = 998 Then Debug.Print "PeutConstruire_Maison"
                If Definir_ressources(i) < Parametres.Batiment_PrixRessource(IndiceMaison, i) Then
                    PeutConstruire_Maison = False
                    Exit For
                End If
            Next i
            If Parametres.Batiment_Empereur(IndiceMaison) Then
                If Not Definir_Empereur Then
                    'Batiments réservés à l'empereur.
                    PeutConstruire_Maison = False
                End If
            End If
        End If
    End If
End Function

Public Function PeutDevenir_Empereur(ByVal Fief As ClsJeuFief, _
                                     ByVal Parametres As ClsJeuParam) As Boolean
    'Renvoie vrai si le personnage peut devenir l'empereur de son fief.
    PeutDevenir_Empereur = Not Fief.Empereur >= 0 And Argent >= Fief.Empereur_Prix(Parametres) And Vivant
End Function

Public Property Get Definir_Empereur() As Boolean
    Definir_Empereur = Empereur
End Property
Public Property Let Definir_Empereur(ByVal Valeur As Boolean)
    Dim i As Long
    If Valeur Then
        For i = 0 To Parametres.NombreObjets - 1
            If Parametres.Objet_Empereur(i) Then
                Objet_Ajouter_Inventaire_Fast i
                'Objet_Ajouter_Inventaire i, Parametres, Commentaires, True
            End If
        Next i
    Else
        Fiefs(NumeroFief).Empereur_Changer -1
        'Supprime le bâtiment de l'empereur.
        If Parametres.Batiment_Empereur(Maisons(Numero).TypeBatiment) Then
            If Maisons(Numero).Visible Then
                Maisons(Numero).Demolir Me, Parametres, True
                Activer_Effet Parametres.EffetMaisonDetruite, Maisons(Numero).PositionX, Maisons(Numero).PositionY, Maisons(Numero).Largeur, Maisons(Numero).Hauteur, Parametres
            End If
        End If
        'Supprime les objets réservés à l'empereur.
        For i = 0 To Nombre_Objets_Equipes - 1
            If ObjetsEquipes(i).Actif Then
                If Parametres.Objet_Empereur(ObjetsEquipes(i).Numero) Then
                    Objet_Equipes_Detruire i
                End If
            End If
        Next i
        For i = 0 To Nombre_Objets_Inventaire - 1
            If ObjetsInventaires(i).Actif Then
                If Parametres.Objet_Empereur(ObjetsInventaires(i).Numero) Then
                    Objet_Inventaire_Detruire i
                End If
            End If
        Next i
    End If
    Empereur = Valeur
End Property

Public Sub Creer_Objet(ByVal NumeroObjet As Integer, _
                       ByVal Maison As ClsJeuBatiment, _
                       ByVal Parametres As ClsJeuParam)
'    Debug.Print Parametres.Objet_Epoque(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, NumeroObjet)) & " : " & Parametres.Objet_Nom(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, NumeroObjet))
'    If Parametres.Objet_Epoque(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, NumeroObjet)) = 5 Then
'        IndiceObjet = IndiceObjet
'    End If
    If NumeroObjet <> IndiceObjet Then
        IndiceObjet = NumeroObjet
        EtatFabricationObjet = 0
    End If
    Action = 500
End Sub

Private Sub Fabriquer_Objet(ByVal Parametres As ClsJeuParam, _
                            ByRef Maison As ClsJeuBatiment, _
                            ByRef NoCommentaire As Integer)
    Dim i As Integer
    Dim VitesseFabrication As Double
    Dim Objet As Integer
    Dim Competence As Integer
    
    If PeutConstruire_Objet(IndiceObjet, Maison, Parametres) Then
        Objet = Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, IndiceObjet)
        Competence = Parametres.Objet_Competence(Objet)
        VitesseFabrication = VitesseFabricationObjet * Vitalite * Definir_Niveau_CompetenceObjets(Competence)
        
        EtatFabricationObjet = EtatFabricationObjet + VitesseFabrication
        'VitesseFabrication / Parametres.Objet_DureeFabrication(Objet) * CoefExpCompObjets,
        
        'Il gagne de l'expérience avec les objets qu'il utilise et qui lui offrent un bonus.
        'Ajuste l'expérience.
        VitesseFabrication = VitesseFabrication * CoefExpCompObjets * CoefExpCompObjetsFabrication * Parametres.Objet_CoefXPFabrication(Objet) * CaracObjets(Competence) / BonusNivObjets(Competence)
        If Bonus_CompetenceObjets(Competence) Then
            For i = 0 To UBound(ObjetsEquipes())
                If ObjetsEquipes(i).Param_BonusCompetencesObjets(Competence) > 0 Then
                    Gestion_Niveau CaracObjets(ObjetsEquipes(i).Competence), _
                                   VitesseFabrication, _
                                   ExpObjets(ObjetsEquipes(i).Competence), NivObjets(ObjetsEquipes(i).Competence), ExpNivObjets(ObjetsEquipes(i).Competence), ExpNivPlusObjets(ObjetsEquipes(i).Competence), 8
                End If
            Next i
        End If
        'Il gagne de l'expérience dans sa compétence objet.
        Gestion_Niveau CaracObjets(Competence), _
                       VitesseFabrication, _
                       ExpObjets(Competence), NivObjets(Competence), ExpNivObjets(Competence), ExpNivPlusObjets(Competence), 8
        If EtatFabricationObjet >= Parametres.Objet_DureeFabrication(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, IndiceObjet)) Then
            'Terminé.
            EtatFabricationObjet = 0
            Maison.Stock_Objets(IndiceObjet) = Maison.Stock_Objets(IndiceObjet) + 1
            NoCommentaire = 0
            Action = 100
            'Le joueur paie son coût.
            For i = 0 To Parametres.NombreRessources - 1
                Definir_ressources(i) = Definir_ressources(i) _
                                        - Parametres.Objet_PrixRessources(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, IndiceObjet), i)
            Next i
        Else
            Action = 500
        End If
    Else
        Arreter
    End If
End Sub

Public Function PeutConstruire_Objet(ByVal NumeroObjet As Long, _
                                     ByVal Maison As ClsJeuBatiment, _
                                     ByVal Parametres As ClsJeuParam) As Boolean
    Dim i As Long
    'Vérifie qu'il y a de la place pour stocker l'objet.
    If NumeroObjet >= 0 And Maison.Definir_Stock <= Maison.MaxStock - 1 Then
        PeutConstruire_Objet = True
        For i = 0 To Parametres.NombreRessources - 1
            If Definir_ressources(i) < Parametres.Objet_PrixRessources(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, NumeroObjet), i) Then
                PeutConstruire_Objet = False
                Exit For
            End If
        Next i
    End If
End Function

' ************************************************************************************
' *                             Manipulation de l'inventaire                         *
' ************************************************************************************

Public Sub Objet_Acheter_Maison(ByVal NumeroObjet As Integer, _
                                ByVal Maison As ClsJeuBatiment, _
                                ByVal Parametres As ClsJeuParam, _
                                ByRef Commentaires As ClsJeuCommentaires, _
                                Optional ByVal Quantite As Integer = 1)
    Dim i As Integer
    Dim Prix As Integer
    'Transfère un objet d'une maison vers son inventaire.
    For i = 1 To Quantite
        Prix = Prix_Marchandee(Parametres.Objet_PrixVente(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, NumeroObjet)) * Parametres.CoefficientAchat, Parametres)
        If Argent >= Prix Then
            Argent = Argent - Prix
            Maison.Argent = Maison.Argent + Prix
            Objet_Ramasser_Maison NumeroObjet, Maison, Parametres, Commentaires
            'Affiche l'argent dépensé.
            AffDegat.Ajouter_Degat Maison.PositionX + Maison.Largeur / 2, Maison.PositionY, Prix, , 2
            'Compétence marchandage.
            Gestion_Niveau_Comp_Speciale 3, Prix, Parametres
        Else
            Exit For
        End If
    Next i
End Sub
Public Sub Objet_Ramasser_Maison(ByVal NumeroObjet As Integer, _
                                 ByVal Maison As ClsJeuBatiment, _
                                 ByVal Parametres As ClsJeuParam, _
                                 ByRef Commentaires As ClsJeuCommentaires, _
                                 Optional ByVal Quantite As Integer = 1)
    'Transfère un objet d'une maison dans son inventaire.
    Dim i As Integer
    For i = 1 To Quantite
        If Maison.Stock_Objets(NumeroObjet) >= 1 And _
           Not Inventaire_Plein Then
            'If Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, NumeroObjet) = -1 Then Debug.Print "Erreur : Objet_Ramasser_Maison"
            Objet_Ajouter_Inventaire Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, NumeroObjet), Parametres, Commentaires, True
            Maison.Stock_Objets(NumeroObjet) = Maison.Stock_Objets(NumeroObjet) - 1
        Else
            Exit For
        End If
    Next i
End Sub
Public Sub Objet_Deposer_Maison(ByVal NumeroObjet As Integer, _
                                ByVal Maison As ClsJeuBatiment, _
                                ByVal Parametres As ClsJeuParam, _
                                Optional ByVal ObjetInventaire As Boolean = True)
    Dim i As Integer
    Dim ObjetDepose As Boolean
    Dim TempObjet As Integer
    If ObjetInventaire Then
        TempObjet = ObjetsInventaires(NumeroObjet).Numero
    Else
        TempObjet = NumeroObjet
    End If
    'Transfère un objet de son inventaire vers sa maison
    If Maison.Definir_Stock <= Maison.MaxStock - 1 Then
        For i = 0 To Parametres.Batiment_NombreObjetsFabriques(Maison.TypeBatiment) - 1
            If TempObjet = Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, i) Then
                If ObjetInventaire Then
                    If ObjetsInventaires(NumeroObjet).Nombre >= 2 Then
                        ObjetsInventaires(NumeroObjet).Quantite = ObjetsInventaires(NumeroObjet).Quantite - 1
                    Else
                        ObjetsInventaires(NumeroObjet).Decharger
                    End If
                End If
                Maison.Stock_Objets(i) = Maison.Stock_Objets(i) + 1
                ObjetDepose = True
                Exit For
            End If
        Next i
    End If
    If Not ObjetDepose Then
        'Objet_Deposer_Sol NumeroObjet, ObjetInventaire
        Objet_DeposerInventaire_Maison TempObjet, Maison, Parametres
    End If
End Sub
Public Sub Objet_DeposerInventaire_Maison(ByVal NumeroObjet As Integer, _
                                          ByVal Maison As ClsJeuBatiment, _
                                          ByVal Parametres As ClsJeuParam)
    'Dépose l'objet dans l'inventaire de la maison.
    Dim i As Integer
    Dim ObjetDepose As Boolean
    'Transfère un objet de son inventaire vers sa maison
    For i = 0 To Maison.Definir_Nombre_Objet_Inventaire - 1
        If Maison.Definir_ObjetsInventaire(i) = -1 Then
            Maison.Definir_ObjetsInventaire(i) = NumeroObjet
            ObjetDepose = True
            Exit For
        End If
    Next i
    If Not ObjetDepose Then
        Objet_Deposer_Sol NumeroObjet, False
    End If
End Sub
Public Sub Objet_Deposer_Sol(ByVal NumeroObjet As Integer, _
                             Optional ByVal ObjetInventaire As Boolean = True)
    'Se débarasse de l'objet sélectionné.
    Dim IndexTresor As Long
    Dim TempObjet As Integer
    If ObjetInventaire Then
        TempObjet = ObjetsInventaires(NumeroObjet).Numero
    Else
        TempObjet = NumeroObjet
    End If
    'If Not ObjetsInventaires(NumeroObjet).Insepararable Then
    If Not Parametres.Objet_Inseparable(TempObjet) Then
        IndexTresor = Tresor_Creer(PositionX, PositionY)
        Tresors(IndexTresor).Ajouter_Objet TempObjet, Parametres
        If ObjetInventaire Then
            ObjetsInventaires(NumeroObjet).Decharger
        Else
            ObjetSelectionne = -1
        End If
    End If
    'Dépose également les ressources sélectionnées.
    Ressources_Deposer_Sol IndexTresor
End Sub
Public Sub Objet_Vendre_Chateau(ByVal NumeroObjet As Integer, _
                                ByVal Fief As ClsJeuFief, _
                                ByVal Parametres As ClsJeuParam, _
                                Optional ByVal ObjetInventaire As Boolean = True)
    'Vend l'objet sélectionné au chateau.
    Dim i  As Integer
    Dim TempObjet As Integer
    If ObjetInventaire Then
        TempObjet = ObjetsInventaires(NumeroObjet).Numero
    Else
        TempObjet = NumeroObjet
    End If
    With Parametres
    'If Not ObjetsInventaires(NumeroObjet).Insepararable Then
    If Not .Objet_Inseparable(TempObjet) Then
        'Definir_Argent = Definir_Argent + ObjetsInventaires(NumeroObjet).PrixVente
        Definir_Argent = Definir_Argent + .Objet_PrixVente(TempObjet)
        'Le fief recycle l'objet et le reconvertit en ressources.
        For i = 0 To .NombreRessources - 1
            'Fief.Definir_Ressources(i) = Fief.Definir_Ressources(i) + ObjetsInventaires(NumeroObjet).Param_PrixRessources(i)
            Fief.Definir_ressources(i) = Fief.Definir_ressources(i) + .Objet_PrixRessources(TempObjet, i)
        Next i
        If ObjetInventaire Then
            ObjetsInventaires(NumeroObjet).Decharger
        End If
    End If
    End With
End Sub
Public Sub Objet_Equipes_Detruire(ByVal NumeroObjet As Integer)
    'L'objet redonne une apparence normale au personnage.
    If ObjetsEquipes(NumeroObjet).ChangerApparencePerso Then
        ChoisirApparence = False
        Apparence = 0
    End If
    'Détruit de l'objet équipé sélectionné.
    ObjetsEquipes(NumeroObjet).Decharger
    Actualiser_Bonus
End Sub
Public Sub Objet_Equipes_Vider_Tout()
    'Se débarasse de tous les objets qu'il possède.
    Dim i As Integer
    For i = 0 To Nombre_Objets_Equipes - 1
        Objet_Equipes_Detruire i
    Next i
End Sub
Public Sub Objet_Inventaire_Detruire(ByVal NumeroObjet As Integer)
    'Détruit de l'objet dans l'inventaire sélectionné.
    ObjetsInventaires(NumeroObjet).Decharger
End Sub
Public Sub Objet_Inventaire_Vider_Tout()
    'Se débarasse de tous les objets qu'il a dans son inventaire.
    Dim i As Integer
    For i = 0 To Nombre_Objets_Inventaire - 1
        Objet_Inventaire_Detruire i
    Next i
End Sub
Public Sub Objet_Ajouter_Inventaire(ByVal IndiceObjet As Integer, _
                                    ByVal Parametres As ClsJeuParam, _
                                    ByRef Commentaires As ClsJeuCommentaires, _
                                    Optional ByVal NouveauObjet As Boolean)
    'Place un nouvel objet dans son inventaire.
    Dim i As Integer
    If IA_Equiper = 0 Then NouveauObjet = False
    For i = 0 To Nombre_Objets_Inventaire - 1
        If Not ObjetsInventaires(i).Actif Then
            ObjetsInventaires(i).Charger IndiceObjet, Parametres
            'Si c'est un nouveau objet, on s'équipe directement avec à condition que le joueur ne soit pas équipé.
            If NouveauObjet And Equipement_Libre(IndiceObjet, Parametres) And _
               Not ObjetsInventaires(i).Utilisable Then  'Empêche les objets utilisables d'être utilisés directement.
                Objet_Equiper i, Parametres, Commentaires
            End If
            Exit For
        ElseIf ObjetsInventaires(i).Numero = IndiceObjet Then
            'Superpose les objets de même catégorie.
            If ObjetsInventaires(i).Nombre < Parametres.Objet_Superposable(ObjetsInventaires(i).Numero) Then
                'S'il est superposable et inférieur au nombre maximum d'objets superposables.
                ObjetsInventaires(i).Nombre = ObjetsInventaires(i).Nombre + 1
                Exit For
            End If
        End If
    Next i
End Sub
Public Sub Objet_Supprimer(ByVal IndiceObjet As Integer)
    'Supprime l'objet d'indice "IndiceObjet".
    Dim i As Integer
    For i = 0 To Nombre_Objets_Equipes - 1
        If ObjetsEquipes(i).Actif Then
            If ObjetsEquipes(i).Numero = IndiceObjet Then
                Objet_Equipes_Detruire i
                Exit Sub
            End If
        End If
    Next i
    For i = 0 To Nombre_Objets_Inventaire - 1
        If ObjetsInventaires(i).Actif Then
            If ObjetsInventaires(i).Numero = IndiceObjet Then
                Objet_Inventaire_Detruire i
                Exit Sub
            End If
        End If
    Next i
End Sub
Public Sub Objet_Ajouter_Inventaire_Fast(ByVal IndiceObjet As Integer)
    'Même fonction que la précedente, mais optimisée.
    Dim i As Integer
    For i = 0 To Nombre_Objets_Inventaire - 1
        If Not ObjetsInventaires(i).Actif Then
            ObjetsInventaires(i).Charger IndiceObjet, Parametres
            Exit For
        ElseIf ObjetsInventaires(i).Numero = IndiceObjet Then
            'Superpose les objets de même catégorie.
            If ObjetsInventaires(i).Nombre < Parametres.Objet_Superposable(ObjetsInventaires(i).Numero) Then
                'S'il est superposable et inférieur au nombre maximum d'objets superposables.
                ObjetsInventaires(i).Nombre = ObjetsInventaires(i).Nombre + 1
                Exit For
            End If
        End If
    Next i
End Sub
Public Function Inventaire_Plein() As Boolean
    'Renvoie vrai s'il n'y a plus de place dans l'inventaire.
    Dim i As Integer
    Inventaire_Plein = True
    For i = 0 To Nombre_Objets_Inventaire - 1
        If Not ObjetsInventaires(i).Actif Then
            Inventaire_Plein = False
            Exit For
        End If
    Next i
End Function
Public Function Objet_Equiper(ByVal NumeroObjet As Integer, _
                              ByVal Parametres As ClsJeuParam, _
                              ByRef Commentaires As ClsJeuCommentaires, _
                              Optional ByVal ObjetInventaire As Boolean = True, _
                              Optional ByVal CoteDroit As Boolean) As Boolean
    'Renvoie vrai, si l'objet est détruit.
    'CoteDroit : s'il ont place l'arme ou l'anneau à droite.
    Dim TempObjet As Integer, TempObjet1 As Integer, TempObjet2 As Integer
    TempObjet1 = -1
    TempObjet2 = -1
    
    If ObjetInventaire Then
        TempObjet = ObjetsInventaires(NumeroObjet).Numero
    Else
        TempObjet = NumeroObjet
    End If
    
    With Parametres
    'S'équipe avec l'objet de l'inventaire désigné.
    'If Not ObjetsInventaires(NumeroObjet).Utilisable Then
    If Not Parametres.Objet_Monture(TempObjet) Or _
       Parametres.Race_UtilisationMonture(Race) Then 'Ne peut prendre de monture.
        If Not .Objet_Utilisable(TempObjet) Then
            Objet_Equiper = True
            If Not Equipement_Libre(TempObjet, Parametres) Then
                Select Case .Objet_Categorie(TempObjet)
                Case 0: 'Objet à une main.
                    If CoteDroit Then
                        If ObjetsEquipes(0).Actif And ObjetsEquipes(1).Actif Then
                            TempObjet1 = ObjetsEquipes(1).Numero
                            Objet_Equipes_Detruire 1
                        End If
                    ElseIf ObjetsEquipes(1).Actif Or ObjetsEquipes(0).Actif And ObjetsEquipes(0).Categorie = 1 Then
                        TempObjet1 = ObjetsEquipes(0).Numero
                        Objet_Equipes_Detruire 0
                    End If
                Case 1: 'Arme à 2 mains
                    TempObjet1 = ObjetsEquipes(0).Numero
                    Objet_Equipes_Detruire 0
                    If ObjetsEquipes(1).Actif Then
                        TempObjet2 = ObjetsEquipes(1).Numero
                        Objet_Equipes_Detruire 1
                    End If
                Case 6, 7: 'Anneaux.
                    If CoteDroit Then
                        If ObjetsEquipes(6).Actif And ObjetsEquipes(7).Actif Then
                            TempObjet1 = ObjetsEquipes(7).Numero
                            Objet_Equipes_Detruire 7
                        End If
                    ElseIf ObjetsEquipes(7).Actif Or ObjetsEquipes(6).Actif And ObjetsEquipes(6).Categorie = 7 Then
                        TempObjet1 = ObjetsEquipes(6).Numero
                        Objet_Equipes_Detruire 6
                    End If
                Case Else
                    'TempObjet1 = ObjetsEquipes(ObjetsInventaires(NumeroObjet).Categorie).Numero
                    TempObjet1 = ObjetsEquipes(.Objet_Categorie(TempObjet)).Numero
                    'Objet_Equipes_Detruire ObjetsInventaires(NumeroObjet).Categorie
                    Objet_Equipes_Detruire .Objet_Categorie(TempObjet)
                End Select
            End If
            'L'objet métamorphose le personnage.
            'If ObjetsInventaires(NumeroObjet).ChangerApparencePerso Then
            If .Objet_ChangerApparencePerso(TempObjet) Then
                ChoisirApparence = True
                Apparence = ApparenceChoisie
            End If
            'Select Case ObjetsInventaires(NumeroObjet).Categorie
            Select Case .Objet_Categorie(TempObjet)
            Case 0: 'Objet à 1 main.
                If CoteDroit Then
                    If ObjetsEquipes(1).Actif Then
                        If ObjetsEquipes(0).Actif Then
                            ObjetsEquipes(1).Charger TempObjet, Parametres
                        Else
                            ObjetsEquipes(0).Charger ObjetsEquipes(1).Numero, Parametres
                            ObjetsEquipes(1).Charger TempObjet, Parametres
                        End If
                    Else
                        If ObjetsEquipes(0).Actif And ObjetsEquipes(0).Categorie = 1 Then
                            Objet_Ranger 0, Parametres, Commentaires
                        End If
                        ObjetsEquipes(1).Charger TempObjet, Parametres
                    End If
                Else
                    If ObjetsEquipes(0).Actif Then
                    'ObjetsEquipes(1).Charger ObjetsInventaires(NumeroObjet).Numero, Parametres
                        If ObjetsEquipes(1).Actif Then
                            Objet_Ranger 1, Parametres, Commentaires
                            ObjetsEquipes(1).Charger ObjetsEquipes(0).Numero, Parametres
                            ObjetsEquipes(0).Charger TempObjet, Parametres
                        Else
                            If ObjetInventaire Then
                                ObjetsEquipes(1).Charger TempObjet, Parametres
                            Else
                                ObjetsEquipes(1).Charger ObjetsEquipes(0).Numero, Parametres
                                ObjetsEquipes(0).Charger TempObjet, Parametres
                            End If
                        End If
                    Else
                        'ObjetsEquipes(0).Charger ObjetsInventaires(NumeroObjet).Numero, Parametres
                        ObjetsEquipes(0).Charger TempObjet, Parametres
                    End If
                End If
            Case 1: 'Objet a 2 mains.
                'ObjetsEquipes(0).Charger ObjetsInventaires(NumeroObjet).Numero, Parametres
                ObjetsEquipes(0).Charger TempObjet, Parametres
            Case 6, 7: 'Anneaux.
                If CoteDroit Then
                    If ObjetsEquipes(7).Actif Then
                        If ObjetsEquipes(6).Actif Then
                            ObjetsEquipes(7).Charger TempObjet, Parametres
                        Else
                            ObjetsEquipes(6).Charger ObjetsEquipes(7).Numero, Parametres
                            ObjetsEquipes(7).Charger TempObjet, Parametres
                        End If
                    Else
                        ObjetsEquipes(7).Charger TempObjet, Parametres
                    End If
                Else
                    If ObjetsEquipes(6).Actif Then
                        If ObjetsEquipes(7).Actif Then
                            Objet_Ranger 7, Parametres, Commentaires
                            ObjetsEquipes(7).Charger ObjetsEquipes(6).Numero, Parametres
                            ObjetsEquipes(6).Charger TempObjet, Parametres
                        Else
                            If ObjetInventaire Then
                                ObjetsEquipes(7).Charger TempObjet, Parametres
                            Else
                                ObjetsEquipes(7).Charger ObjetsEquipes(6).Numero, Parametres
                                ObjetsEquipes(6).Charger TempObjet, Parametres
                            End If
                        End If
                    Else
                        'ObjetsEquipes(6).Charger ObjetsInventaires(NumeroObjet).Numero, Parametres
                        ObjetsEquipes(6).Charger TempObjet, Parametres
                    End If
                End If
            Case Else
                'ObjetsEquipes(ObjetsInventaires(NumeroObjet).Categorie).Charger ObjetsInventaires(NumeroObjet).Numero, Parametres
                ObjetsEquipes(.Objet_Categorie(TempObjet)).Charger TempObjet, Parametres
            End Select
            If ObjetInventaire Then
                ObjetsInventaires(NumeroObjet).Decharger
            Else
                ObjetSelectionne = -1
            End If
            'On récupère l'équipement qui a été remplacé.
            If TempObjet1 >= 0 Then
                Objet_Ajouter_Inventaire TempObjet1, Parametres, Commentaires
            End If
            If TempObjet2 >= 0 Then
                Objet_Ajouter_Inventaire TempObjet2, Parametres, Commentaires
            End If
            'On supprime les armes en trop.
            If ObjetsEquipes(0).Actif And ObjetsEquipes(0).Categorie = 1 And ObjetsEquipes(1).Actif Then
                ObjetsEquipes(1).Decharger
            End If
            Actualiser_Bonus
            'Calculer_Apparence Maisons(Numero), Parametres
        'ElseIf (Vivant Or ObjetsInventaires(NumeroObjet).UtilisationMort) And NumeroObjetUtilise = -1 And Not (Vivant And ObjetsInventaires(NumeroObjet).SelfResurrection) Then
        ElseIf EffetRecharge_Pret(TempObjet, Parametres) Then
            If (Vivant Or .Objet_UtilisationMort(TempObjet)) And NumeroObjetUtilise = -1 And Not (Vivant And .Objet_SelfResurrection(TempObjet)) Then
                'If Vie > ObjetsInventaires(NumeroObjet).PerteVie Or Not Vivant Then
                If Vie > .Objet_PerteVie(TempObjet) Or Not Vivant Then
                    'If Energie > ObjetsInventaires(NumeroObjet).PerteEnergie Then
                    If Energie > .Objet_PerteEnergie(TempObjet) Then
                        'If Magie > ObjetsInventaires(NumeroObjet).PerteMagie Then
                        If Magie >= .Objet_PerteMagie(TempObjet) Then
                            'If Moral > ObjetsInventaires(NumeroObjet).PerteMoral Then
                            If Moral > .Objet_PerteMoral(TempObjet) Then
                                'If Argent >= ObjetsInventaires(NumeroObjet).PerteArgent Then
                                If Argent >= .Objet_PerteArgent(TempObjet) Then
                                    'If Ressources(PerteRessourceIndice) >= ObjetsInventaires(NumeroObjet).PerteRessourceQuantite Then
                                    If Ressources(.Objet_PerteRessourceIndice(TempObjet)) >= .Objet_PerteRessourceQuantite(TempObjet) Then
                                        'On utilise l'objet.
                                        'NumeroObjetUtilise = ObjetsInventaires(NumeroObjet).Numero
                                        NumeroObjetUtilise = TempObjet
                                        'On gagne de l'expérience.
                                        'Gestion_Niveau CaracObjets(ObjetsInventaires(NumeroObjet).Competence), _
                                                       ObjetsInventaires(NumeroObjet).Experience, _
                                                       ExpObjets(ObjetsInventaires(NumeroObjet).Competence), NivObjets(ObjetsInventaires(NumeroObjet).Competence), ExpNivObjets(ObjetsInventaires(NumeroObjet).Competence), ExpNivPlusObjets(ObjetsInventaires(NumeroObjet).Competence), 8
                                        Gestion_Niveau CaracObjets(.Objet_Competence(TempObjet)), _
                                                       .Objet_Experience(TempObjet), _
                                                       ExpObjets(.Objet_Competence(TempObjet)), NivObjets(.Objet_Competence(TempObjet)), ExpNivObjets(.Objet_Competence(TempObjet)), ExpNivPlusObjets(.Objet_Competence(TempObjet)), 8
                                        'On perd des caractéristiques.
                                        'Definir_Vie = Vie - ObjetsInventaires(NumeroObjet).PerteVie
                                        Definir_Vie = Vie - .Objet_PerteVie(TempObjet)
                                        'Definir_Energie = Energie - ObjetsInventaires(NumeroObjet).PerteEnergie
                                        Definir_Energie = Energie - .Objet_PerteEnergie(TempObjet)
                                        'Definir_Magie = Magie - ObjetsInventaires(NumeroObjet).PerteMagie
                                        Definir_Magie = Magie - .Objet_PerteMagie(TempObjet)
                                        'Definir_Moral = Moral - ObjetsInventaires(NumeroObjet).PerteMoral
                                        Definir_Moral = Moral - .Objet_PerteMoral(TempObjet)
                                        'L'objet disparaît.
                                        'If ObjetsInventaires(NumeroObjet).UtilisationUnique Then
                                        '    ObjetsInventaires(NumeroObjet).Decharger
                                        'End If
                                        If .Objet_UtilisationUnique(TempObjet) Then
                                            If ObjetInventaire Then
                                                ObjetsInventaires(NumeroObjet).Decharger
                                            Else
                                                ObjetSelectionne = -1
                                            End If
                                        End If
                                        'Definir_Argent = Definir_Argent - ObjetsInventaires(NumeroObjet).PerteArgent
                                        Definir_Argent = Definir_Argent - .Objet_PerteArgent(TempObjet)
                                        'If ObjetsInventaires(NumeroObjet).PerteRessourceQuantite > 0 Then
                                        If .Objet_PerteRessourceQuantite(TempObjet) > 0 Then
                                            'Definir_Ressources(ObjetsInventaires(NumeroObjet).PerteRessourceIndice) = Definir_Ressources(ObjetsInventaires(NumeroObjet).PerteRessourceIndice) - ObjetsInventaires(NumeroObjet).PerteRessourceQuantite
                                            Definir_ressources(.Objet_PerteRessourceIndice(TempObjet)) = Definir_ressources(.Objet_PerteRessourceIndice(TempObjet)) - .Objet_PerteRessourceQuantite(TempObjet)
                                        End If
                                        'Compte le temps de rechargement.
                                        If .Objet_DureeRecharge(TempObjet) > 0 Then
                                            EffetRecharge_Ajouter TempObjet, .Objet_DureeRecharge(TempObjet)
                                        End If
                                        Objet_Equiper = True
                                    Else
                                        'N'a plus assez de ressources.
                                        Definir_Commentaires = Commentaires.Message(47, Me)
                                    End If
                                Else
                                    'N'a plus d'argent.
                                    Definir_Commentaires = Commentaires.Message(45, Me)
                                End If
                            Else
                                'N'a plus de point de moral.
                                Definir_Commentaires = Commentaires.Message(20, Me)
                            End If
                        Else
                            'N'a plus de point de magie.
                            Definir_Commentaires = Commentaires.Message(19, Me)
                        End If
                    Else
                        'N'a plus de point d'énergie.
                        Definir_Commentaires = Commentaires.Message(18, Me)
                    End If
                Else
                    'N'a plus de point de vie.
                    Definir_Commentaires = Commentaires.Message(17, Me)
                End If
                'L'objet n'est pas encore rechargé.
                
            End If
        End If
    End If
    End With
End Function
Private Function Equipement_Libre(ByVal IndiceObjet As Integer, _
                                  ByVal Parametres As ClsJeuParam) As Boolean
    'Verifie s'il y a de la place pour s'équiper avec l'objet définie.
    Select Case Parametres.Objet_Categorie(IndiceObjet)
    Case 0: 'Arme à une main.
        If Not (ObjetsEquipes(0).Actif And ObjetsEquipes(1).Actif) And _
           Not (ObjetsEquipes(0).Actif And ObjetsEquipes(0).Categorie = 1) Then
            Equipement_Libre = True
        End If
    Case 1: 'Arme à 2 mains
        If Not (ObjetsEquipes(0).Actif Or ObjetsEquipes(1).Actif) Then
            Equipement_Libre = True
        End If
    Case 6, 7: 'Anneau.
        If Not (ObjetsEquipes(6).Actif And ObjetsEquipes(7).Actif) Then
            Equipement_Libre = True
        End If
    Case Else:
        If Not ObjetsEquipes(Parametres.Objet_Categorie(IndiceObjet)).Actif Then
            Equipement_Libre = True
        End If
    End Select
End Function

Public Function Objet_Posseder(ByVal NumeroObjet As Integer) As Boolean
    'Renvoie si le personnage possède un au moins un objet de ce type.
    Dim i As Integer
    Objet_Posseder = Objet_NombreType(NumeroObjet) > 0
End Function

Public Property Get Objet_EquipeNombreType(ByVal NumeroObjet As Integer) As Integer
    'Renvoie le nombre d'objets que le personnage a d'équipé sur lui.
    Dim i As Integer
    For i = 0 To Nombre_Objets_Equipes - 1
        If Objet_Equipes_Actif(i) And _
           Objet_Equipes_Type(i) = NumeroObjet Then
            Objet_EquipeNombreType = Objet_EquipeNombreType + 1
        End If
    Next i
End Property

Public Function Objet_NombreType(ByVal NumeroObjet As Integer) As Integer
    'Renvoie le nombre d'objet de ce type que possède ce personnage.
    Dim i As Integer
    For i = 0 To UBound(ObjetsEquipes())
        If ObjetsEquipes(i).Actif And _
           ObjetsEquipes(i).Numero = NumeroObjet Then
            Objet_NombreType = Objet_NombreType + 1
        End If
    Next i
    For i = 0 To UBound(ObjetsInventaires())
        If ObjetsInventaires(i).Actif And _
           ObjetsInventaires(i).Numero = NumeroObjet Then
            Objet_NombreType = Objet_NombreType + ObjetsInventaires(i).Nombre
        End If
    Next i
End Function

Public Function Objet_Nombre_Total(Optional ByVal Inventaire As Boolean) As Integer
    'Renvoie le nombre d'objets que possède le personnage.
    Dim i As Integer
    If Not Inventaire Then
        For i = 0 To UBound(ObjetsEquipes())
            If ObjetsEquipes(i).Actif Then
                Objet_Nombre_Total = Objet_Nombre_Total + 1
            End If
        Next i
    End If
    For i = 0 To UBound(ObjetsInventaires())
        If ObjetsInventaires(i).Actif Then
            Objet_Nombre_Total = Objet_Nombre_Total + 1
        End If
    Next i
End Function

Public Function Ressources_TotalPrix() As Long
    'Renvoie la valeur en argent de tous les objets que le perso a.
    Dim i As Integer
    Dim Total As Long
    For i = 0 To UBound(Ressources())
        Total = Total + Ressources(i) * Parametres.Ressources_PrixVente(i)
    Next i
    Ressources_TotalPrix = Total
End Function
Public Function Objets_TotalPrix() As Long
    'Renvoie la valeur en argent de tous les objets que le perso a.
    Dim i As Integer
    Dim Total As Long
    For i = 0 To UBound(ObjetsEquipes())
        If ObjetsEquipes(i).Actif Then
            If Not Parametres.Objet_Inseparable(ObjetsEquipes(i).Numero) Then
                Total = Total + Parametres.Objet_PrixVente(ObjetsEquipes(i).Numero)
            End If
        End If
    Next i
    For i = 0 To UBound(ObjetsInventaires())
        If ObjetsInventaires(i).Actif Then
            If Not Parametres.Objet_Inseparable(ObjetsInventaires(i).Numero) Then
                Total = Total + Parametres.Objet_PrixVente(ObjetsInventaires(i).Numero)
            End If
        End If
    Next i
    Objets_TotalPrix = Total
End Function

Public Sub Objet_Ranger(ByVal NumeroObjet As Integer, _
                        ByVal Parametres As ClsJeuParam, _
                        ByRef Commentaires As ClsJeuCommentaires)
    'On passe l'indice de l'objet équipé.
    'Remet l'objet équipé dans l'inventaire.
    If Not Inventaire_Plein Then
        Objet_Ajouter_Inventaire ObjetsEquipes(NumeroObjet).Numero, Parametres, Commentaires
        Objet_Equipes_Detruire NumeroObjet
        Actualiser_Bonus
        'Calculer_Apparence Maisons(Numero), Parametres
        'Réajuste les caractéristiques de combat.
        If Attaque > Definir_MaxAttaque Then
            Attaque = Definir_MaxAttaque
        End If
        If Defense > Definir_MaxDefense Then
            Defense = Definir_MaxDefense
        End If
    End If
End Sub

Public Sub Objet_Equipe_Charger(ByVal i As Integer, _
                                ByVal IndiceObjet As Integer, _
                                ByVal Parametres As ClsJeuParam)
    'If IndiceObjet = -1 Then Debug.Print "Erreur : Objet_Equipe_Charger"
    ObjetsEquipes(i).Charger IndiceObjet, Parametres
End Sub
Public Sub Objet_Inventaire_Charger(ByVal i As Integer, _
                                    ByVal IndiceObjet As Integer, _
                                    ByVal Parametres As ClsJeuParam)
    'If IndiceObjet = -1 Then Debug.Print "Erreur : Objet_Inventaire_Charger"
    ObjetsInventaires(i).Charger IndiceObjet, Parametres
End Sub

Public Sub Objet_Ranger_Armes(ByRef Commentaires As ClsJeuCommentaires, _
                              ByVal Parametres As ClsJeuParam)
    Arreter
    If ObjetsEquipes(0).Actif Then
        Objet_Ranger 0, Parametres, Commentaires
    End If
    If ObjetsEquipes(1).Actif Then
        Objet_Ranger 1, Parametres, Commentaires
    End If
End Sub

Public Sub Objet_Prendre_Premier()
    'Prend le premier objet et le sélectionne pour le donner.
    Dim i As Integer
'    For i = 0 To UBound(ObjetsEquipes())
'        If ObjetsEquipes(i).Actif Then
'            ObjetSelectionne = ObjetsEquipes(i).Numero
'            ObjetsEquipes(i).Decharger
'        End If
'    Next i
    For i = 0 To UBound(ObjetsInventaires())
        If ObjetsInventaires(i).Actif Then
            ObjetSelectionne = ObjetsInventaires(i).Numero
            ObjetsInventaires(i).Decharger
        End If
    Next i
End Sub

Public Sub Actualiser_Bonus()
    'Utiliser cette fonction chaque fois que l'on s'équipe ou se deséquipe d'un objet ou au changement d'un niveau de compétence d'objet.
    Dim i As Long
    Dim j As Long
    Dim TempBonusDeplacement As Long
    Dim TempBonusMaxVie As Long
    Dim TempBonusMaxEnergie As Long
    Dim TempBonusMaxMagie As Long
    Dim TempBonusMaxMoral As Long
    Dim TempBonusMaxAttaque As Long
    Dim TempBonusMaxDefense As Long
    
    Nombre_ApparenceObjet = 0
    BonusVitalite = 1
    BonusBerserk = Berserk
    
    GenerationRessources = False
    AttaqueAutomatique = False
    Collection_Vider BonusAttaqueCauseEffet
    If AttaqueCauseEffet >= 0 Then
        BonusAttaqueCauseEffet.Add AttaqueCauseEffet
    End If
    GainArgent = 0
    For i = 0 To UBound(BonusNivRessources())
        GainRessource(i) = 0
    Next i
    
    TempBonusDeplacement = BonusDeplacement
    TempBonusMaxVie = BonusMaxVie
    TempBonusMaxEnergie = BonusMaxEnergie
    TempBonusMaxMagie = BonusMaxMagie
    TempBonusMaxMoral = BonusMaxMoral
    TempBonusMaxAttaque = BonusMaxAttaque
    TempBonusMaxDefense = BonusMaxDefense
    
    PerteVie = 0
    PerteEnergie = 0
    PerteMagie = 0
    PerteMoral = 0
    PerteArgent = 0
    PerteRessourceIndice = 0
    PerteRessourceQuantite = 0
    Poison = 0
    PoisonLanceur = -1
    DureeRechargeAttaqueObjet = 0
    NombreProjectiles = 1
    
    Ajuster_Bonus
    
    'On cherche en premier les bonus des objets.
    For i = 0 To UBound(ObjetsEquipes())
        If ObjetsEquipes(i).Actif Then
            For j = 0 To UBound(BonusNivObjets())
                BonusNivObjets(j) = BonusNivObjets(j) + Objet_Equipe_BonusSup(Definir_Niveau_CompetenceObjets(ObjetsEquipes(i).Competence), ObjetsEquipes(i).Param_BonusCompetencesObjets(j))
            Next j
        End If
    Next i
    For i = 1 To Nombre_EffetsTemp
        For j = 0 To UBound(BonusNivObjets())
            BonusNivObjets(j) = BonusNivObjets(j) + Objet_Equipe_BonusSup(Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(EffetsTemp(i).Numero)), Parametres.Objet_BonusCompetencesObjets(EffetsTemp(i).Numero, j))
        Next j
    Next i
    'Ajustements de la race.
    For i = 0 To UBound(BonusNivObjets())
        BonusNivObjets(i) = BonusNivObjets(i) * CoefRaceCompObjets(i)
    Next i
    
    If Not ChoisirApparence Then
        'While ApparenceObjet.Count > 0
        '    ApparenceObjet.Remove 1
        'Wend
        Nombre_ApparenceObjet = 0
    End If
    'Objets sur soi.
    'For i = UBound(ObjetsEquipes()) To 0 Step -1
    For i = 2 To UBound(ObjetsEquipes())
        If ObjetsEquipes(i).Numero > -1 Then
            Ajouter_Bonus_Objet ObjetsEquipes(i).Numero
        End If
    Next i
    'On prend le bonus de l'arme en dernier.
    For i = 0 To 1
        If ObjetsEquipes(i).Numero > -1 Then
            Ajouter_Bonus_Objet ObjetsEquipes(i).Numero
        End If
    Next i
    'Effets temporaires.
    For i = 1 To Nombre_EffetsTemp
        Ajouter_Bonus_Objet Definir_EffetTemp_Numero(i), True, Definir_EffetTemp_Proprietaire(i)
    Next i
    
    If Not ChoisirApparence Then
        ApparenceObjetAfficher = Nombre_ApparenceObjet > 0
    End If
    
    'Ajustements de la race.
    BonusMaxVie = BonusMaxVie * CoefRaceVie
    BonusMaxEnergie = BonusMaxEnergie * CoefRaceEnergie
    BonusMaxMagie = BonusMaxMagie * CoefRaceMagie
    BonusMaxMoral = BonusMaxMoral * CoefRaceMoral
    BonusMaxAttaque = BonusMaxAttaque * CoefRaceAttaque
    BonusMaxDefense = BonusMaxDefense * CoefRaceDefense
    For i = 0 To UBound(BonusNivSpeciales())
        BonusNivSpeciales(i) = BonusNivSpeciales(i) * CoefRaceCompSpeciales(i)
    Next i
    For i = 0 To UBound(BonusNivRessources())
        BonusNivRessources(i) = BonusNivRessources(i) + Parametres.PersosMaxRessources * CoefRaceCompRessources(i) - Parametres.PersosMaxRessources
    Next i
    For i = 0 To UBound(BonusNivServices())
        BonusNivServices(i) = BonusNivServices(i) * CoefRaceCompServices(i)
    Next i
    
    'On évite les scores négatifs.
    If BonusDeplacement <= 1 Then BonusDeplacement = IIf(Immobile, 0, 1)
    If BonusMaxVie <= 0 Then BonusMaxVie = 1
    If BonusMaxEnergie <= 0 Then BonusMaxEnergie = 1
    If BonusMaxMagie <= 0 Then BonusMaxMagie = 1
    If BonusMaxMoral <= 0 Then BonusMaxMoral = 1
    If BonusMaxAttaque <= 0 Then BonusMaxAttaque = 1
    If BonusMaxDefense <= 0 Then BonusMaxDefense = 1
    
    BonusDeplacement = BonusDeplacement + BonusVitesse
    
    'On rééquilibre les bonus.
    If TempBonusMaxVie = 0 Or TempBonusMaxEnergie = 0 Or TempBonusMaxMagie = 0 Or TempBonusMaxMoral = 0 Or TempBonusMaxAttaque = 0 Or TempBonusMaxDefense = 0 Then MsgBox "Erreur dans la fonction ActualiserBonus de la Class JeuPerso" '%%%
    Vie = Vie * BonusMaxVie / TempBonusMaxVie
    Energie = Energie * BonusMaxEnergie / TempBonusMaxEnergie
    Magie = Magie * BonusMaxMagie / TempBonusMaxMagie
    Moral = Moral * BonusMaxMoral / TempBonusMaxMoral
    RecuperationVie = RecuperationVie * BonusMaxVie / TempBonusMaxVie
    RecuperationEnergie = RecuperationEnergie * BonusMaxEnergie / TempBonusMaxEnergie
    RecuperationMagie = RecuperationMagie * BonusMaxMagie / TempBonusMaxMagie
    RecuperationMoral = RecuperationMoral * BonusMaxMoral / TempBonusMaxMoral
    Attaque = Attaque * BonusMaxAttaque / TempBonusMaxAttaque
    Defense = Defense * BonusMaxDefense / TempBonusMaxDefense
    If Attaque > BonusMaxAttaque Then Attaque = BonusMaxAttaque
    If Armure > BonusMaxDefense Then BonusMaxDefense = Armure
    If Defense > BonusMaxDefense Then Defense = BonusMaxDefense
    
    'On élimine le surplus de ressources.
    For i = 0 To UBound(Ressources())
        If Ressources(i) > Definir_MaxRessources(i) Then
            Ressources(i) = Definir_MaxRessources(i)
        End If
    Next i
End Sub

Private Sub Ajouter_Bonus_Objet(ByVal Numero As Integer, _
                                Optional ByVal EffetTemp As Boolean, _
                                Optional ByVal Proprietaire As Long = -1)
    Dim j As Long
    Dim Competence As Integer 'Score de la compétence objet.
    With Parametres
    If Proprietaire = -1 Then
        Competence = Definir_Niveau_CompetenceObjets(.Objet_Competence(Numero))
    Else
        Competence = Persos(Proprietaire).Definir_Niveau_CompetenceObjets(.Objet_Competence(Numero))
    End If

    If Not Definir_Volant Then 'Les volants n'ont pas de bonus de déplacement.
        'On ne garde que le dernier bonus. (Le cheval avant les chaussures !!!)
        If Objet_Equipe_BonusSup(Competence, .Objet_BonusDeplacement(Numero)) >= 1 Then
            BonusDeplacement = LongueurPas + Objet_Equipe_BonusSup(Competence, .Objet_BonusDeplacement(Numero))
        End If
    End If
    'Les malus en déplacement s'appliquent aux volants et aux non volants.
    If Objet_Equipe_BonusSup(Competence, .Objet_BonusDeplacement(Numero)) < 0 Then
        BonusDeplacement = LongueurPas + .Objet_BonusDeplacement(Numero)
    End If
    BonusVitesse = BonusVitesse + Objet_Equipe_BonusSup(Competence, .Objet_BonusVitesse(Numero))
    BonusVitalite = BonusVitalite + .Objet_BonusVitalite(Numero)
    BonusMaxVie = BonusMaxVie + Objet_Equipe_BonusSup(Competence, .Objet_BonusVie(Numero))
    BonusMaxEnergie = BonusMaxEnergie + Objet_Equipe_BonusSup(Competence, .Objet_BonusEnergie(Numero))
    BonusMaxMagie = BonusMaxMagie + Objet_Equipe_BonusSup(Competence, .Objet_BonusMagie(Numero))
    BonusMaxMoral = BonusMaxMoral + Objet_Equipe_BonusSup(Competence, .Objet_BonusMoral(Numero))
    BonusMaxAttaque = BonusMaxAttaque + Objet_Equipe_BonusSup(Competence, .Objet_BonusAttaque(Numero))
    BonusMaxDefense = BonusMaxDefense + Objet_Equipe_BonusSup(Competence, .Objet_BonusDefense(Numero))
    If .Objet_AttaqueResistance(Numero) > 0 And Not EffetTemp Then
        BonusAttaqueResistance = .Objet_AttaqueResistance(Numero)
    End If
    BonusPorteeAttaque = BonusPorteeAttaque + Objet_Equipe_PorteeSup(Competence, .Objet_BonusPorteeAttaque(Numero))
    If .Objet_LanceProjectile(Numero) > 0 Then
        LanceProjectile = .Objet_LanceProjectile(Numero)
    End If
    BonusSurfaceAttaque = BonusSurfaceAttaque + Objet_Equipe_SurfaceSup(Competence, .Objet_SurfaceAttaque(Numero))
    If .Objet_AttaqueCauseEffet(Numero) >= 0 Then
        'Debug.Print Nom
        BonusAttaqueCauseEffet.Add .Objet_AttaqueCauseEffet(Numero)
    End If
    If .Objet_AttaqueAutomatique(Numero) Then AttaqueAutomatique = True
    If .Objet_Berserk(Numero) > 0 Then BonusBerserk = .Objet_Berserk(Numero)
    
    If Not EffetTemp Then
        PerteVie = PerteVie + .Objet_PerteVie(Numero)
        PerteEnergie = PerteEnergie + .Objet_PerteEnergie(Numero)
        PerteMagie = PerteMagie + .Objet_PerteMagie(Numero)
        PerteMoral = PerteMoral + .Objet_PerteMoral(Numero)
        PerteArgent = PerteArgent + .Objet_PerteArgent(Numero)
        If .Objet_PerteRessourceIndice(Numero) > 0 Then
            PerteRessourceIndice = .Objet_PerteRessourceIndice(Numero)
        End If
        PerteRessourceQuantite = PerteRessourceQuantite + .Objet_PerteRessourceQuantite(Numero)
    End If
    'Poison.
    If .Objet_Poison(Numero) > 0 Then
        Poison = Poison + IIf(Objet_Equipe_BonusSup(Competence, .Objet_Poison(Numero)) < 1, 1, Objet_Equipe_BonusSup(Competence, .Objet_Poison(Numero)))
        If Proprietaire >= 0 Then
            PoisonLanceur = Proprietaire
        End If
    End If
    DureeRechargeAttaqueObjet = DureeRechargeAttaqueObjet + .Objet_DureeRechargeAttaque(Numero)
    'Nombre de projectiles.
    NombreProjectiles = NombreProjectiles + Objet_Equipe_BonusSup(Competence, .Objet_ProjectileSupplementaire(Numero))
    
    For j = 0 To UBound(BonusNivSpeciales())
        BonusNivSpeciales(j) = BonusNivSpeciales(j) + Objet_Equipe_BonusSup(Competence, .Objet_BonusCompetencesSpeciales(Numero, j))
    Next j
    For j = 0 To UBound(BonusNivRessources())
        BonusNivRessources(j) = BonusNivRessources(j) + Objet_Equipe_BonusSup(Competence, .Objet_BonusCompetencesRessources(Numero, j))
    Next j
    For j = 0 To UBound(BonusNivServices())
        BonusNivServices(j) = BonusNivServices(j) + Objet_Equipe_BonusSup(Competence, .Objet_BonusCompetencesServices(Numero, j))
    Next j
    
    If .Objet_GainArgent(Numero) > 0 Then
        GainArgent = GainArgent + Objet_Equipe_GainSup(Competence, .Objet_GainArgent(Numero))
        GenerationRessources = True
    End If
    For j = 0 To UBound(GainRessource())
        If .Objet_GainRessource(Numero, j) > 0 Then
            GainRessource(j) = GainRessource(j) + Objet_Equipe_GainSup(Competence, .Objet_GainRessource(Numero, j))
            GenerationRessources = True
        End If
    Next j
    BonusArmure = BonusArmure + Objet_Equipe_ArmureSup(Competence, .Objet_BonusArmure(Numero))
    BonusMaxRessourcesObjet = BonusMaxRessourcesObjet + Objet_Equipe_BonusSup(Competence, .Objet_BonusMaxRessources(Numero))
    BonusVitesseRegenerationVie = BonusVitesseRegenerationVie + .Objet_BonusVitesseRegenerationVie(Numero)
    BonusVitesseRegenerationEnergie = BonusVitesseRegenerationEnergie + .Objet_BonusVitesseRegenerationEnergie(Numero)
    BonusVitesseRegenerationMagie = BonusVitesseRegenerationMagie + .Objet_BonusVitesseRegenerationMagie(Numero)
    BonusVitesseRegenerationMoral = BonusVitesseRegenerationMoral + .Objet_BonusVitesseRegenerationMoral(Numero)
    
    'Bonus de résistances.
    For j = 0 To .NombreResistances - 1
        If .Objet_BonusResistance(Numero, j) <> 0 Then
            BonusResistances(j) = BonusResistances(j) - .Objet_BonusResistance(Numero, j)
            If BonusResistances(j) < .ResistanceMinimum Then
                BonusResistances(j) = .ResistanceMinimum
            ElseIf BonusResistances(j) > .ResistanceMaximum Then
                   BonusResistances(j) = .ResistanceMaximum
            End If
        End If
    Next j
    
    If Not (ChoisirApparence Or ApparenceVerouillee) Or .Objet_Monture(Numero) Then
        If .Objet_ChangerApparence(Numero) Then
            'Nombre_ApparenceObjet = Nombre_ApparenceObjet + 1
            'ApparenceObjet(NombreApparenceObjet - 1) = .Objet_NumeroApparence(Numero, Feminin)
            ApparenceObjet.Add .Objet_NumeroApparence(Numero, Feminin)
            If .Objet_Monture(Numero) Then
                Monture = ApparenceObjet.Count
                ApparenceLargeur = .Objet_ApparenceLargeur(Numero)
                ApparenceHauteur = .Objet_ApparenceHauteur(Numero)
                If ApparenceLargeur = 0 Then
                    ApparenceLargeur = Largeur
                    ApparenceDecalageXTemp = 0
                Else
                    ApparenceDecalageXTemp = -((ApparenceLargeur - Largeur) / 2)
                End If
                If ApparenceHauteur = 0 Then
                    ApparenceHauteur = Hauteur
                    ApparenceDecalageYTemp = 0
                Else
                    ApparenceDecalageYTemp = -((ApparenceHauteur - Hauteur) / 2)
                End If
                ApparenceDecalageX(0) = .Objet_ApparenceDecalageX(Numero, 3)
                ApparenceDecalageY(0) = .Objet_ApparenceDecalageY(Numero, 3)
                ApparenceAvantPlan(0) = .Objet_ApparenceAvantPlan(Numero, 3)
                For j = 1 To 4
                    ApparenceDecalageX(j) = .Objet_ApparenceDecalageX(Numero, j)
                    ApparenceDecalageY(j) = .Objet_ApparenceDecalageY(Numero, j)
                    ApparenceAvantPlan(j) = .Objet_ApparenceAvantPlan(Numero, j)
                Next j
            End If
        End If
    End If
    If .Objet_Volant(Numero) Then BonusVolant = True 'L'objet permet au personnage de voler.
    
    End With
End Sub

Public Sub Ajuster_Bonus()
    Dim j As Long
    BonusVitesseRegenerationVie = VitesseRegenerationVie
    BonusVitesseRegenerationEnergie = VitesseRegenerationEnergie
    BonusVitesseRegenerationMagie = VitesseRegenerationMagie
    BonusVitesseRegenerationMoral = VitesseRegenerationMoral
    
    BonusDeplacement = LongueurPas
    BonusVitesse = 0
    BonusMaxVie = MaxVie
    BonusMaxEnergie = MaxEnergie
    BonusMaxMagie = MaxMagie
    BonusMaxMoral = MaxMoral
    BonusMaxAttaque = MaxAttaque
    BonusMaxDefense = MaxDefense
    BonusPorteeAttaque = PorteeAttaque
    LanceProjectile = -1
    For j = 0 To UBound(BonusNivSpeciales())
        BonusNivSpeciales(j) = CaracSpeciales(j)
    Next j
    For j = 0 To UBound(BonusNivRessources())
        BonusNivRessources(j) = NivRessources(j)
    Next j
    For j = 0 To UBound(BonusNivServices())
        BonusNivServices(j) = CaracServices(j)
    Next j
    For j = 0 To UBound(BonusNivObjets())
        BonusNivObjets(j) = CaracObjets(j)
    Next j
    For j = 0 To UBound(BonusResistances())
        BonusResistances(j) = Resistances(j)
    Next j
    BonusMaxRessourcesObjet = BonusMaxRessourcesRace
    BonusAttaqueResistance = AttaqueResistance
    BonusSurfaceAttaque = SurfaceAttaque
    BonusArmure = Armure
    Monture = 0
    BonusVolant = Volant
End Sub

'Bonus supplémentaire de l'objet grâce à la compétence.
Public Property Get Objet_Equipe_BonusSup(ByVal CompetenceObjet As Integer, _
                                          ByVal ValeurBonus As Single) As Integer
    Objet_Equipe_BonusSup = Int((((ValeurBonus * _
                                CompetenceObjet / _
                                CaracObjetsDepart) - ValeurBonus) * _
                                CoefExpCompObjetsBonus) + ValeurBonus)
End Property
Public Property Get Objet_Equipe_GainSup(ByVal CompetenceObjet As Integer, _
                                         ByVal ValeurBonus As Single) As Single
    Objet_Equipe_GainSup = (((ValeurBonus * _
                              CompetenceObjet / _
                              CaracObjetsDepart) - ValeurBonus) * _
                              CoefExpCompObjetsBonus) + ValeurBonus
End Property
Public Property Get Objet_Equipe_PorteeSup(ByVal CompetenceObjet As Integer, _
                                           ByVal ValeurBonus As Single) As Integer
    Objet_Equipe_PorteeSup = Int((((ValeurBonus * _
                                 CompetenceObjet / _
                                 CaracObjetsDepart) - ValeurBonus) * _
                                 CoefExpCompPorteeBonus) + ValeurBonus)
End Property
Public Property Get Objet_Equipe_SurfaceSup(ByVal CompetenceObjet As Integer, _
                                            ByVal ValeurBonus As Single) As Integer
    Objet_Equipe_SurfaceSup = Int((((ValeurBonus * _
                                 CompetenceObjet / _
                                 CaracObjetsDepart) - ValeurBonus) * _
                                 CoefExpCompSurfaceBonus) + ValeurBonus)
End Property

Public Property Get Objet_Equipe_ArmureSup(ByVal CompetenceObjet As Integer, _
                                           ByVal ValeurBonus As Single) As Integer
    Objet_Equipe_ArmureSup = Int((((ValeurBonus * _
                                 CompetenceObjet / _
                                 CaracObjetsDepart) - ValeurBonus) * _
                                 CoefExpCompArmureBonus) + ValeurBonus)
End Property
Public Property Get Objet_Equipe_InvocationSup(ByVal CompetenceObjet As Integer, _
                                               ByVal ValeurBonus As Single) As Single
    Dim Temp As Single
    Temp = (((ValeurBonus * _
           CompetenceObjet / _
           CaracObjetsDepart) - ValeurBonus) * _
           (Jeu.Regle_Invocations_Multiples / 100)) + ValeurBonus
    If Temp < 1 Then
        Objet_Equipe_InvocationSup = 1
    Else
        Objet_Equipe_InvocationSup = Temp
    End If
End Property

' ************************************************************************************
' *                              Propriétés des caractéristiques                     *
' ************************************************************************************

Public Property Get Definir_MaxVie() As Double
    Definir_MaxVie = BonusMaxVie
End Property
Public Property Let Definir_MaxVie(ByVal Valeur As Double)
    'Dim Temp As Double
    'Temp = Valeur * Vie / MaxVie
    MaxVie = Valeur
    'Vie = Max * Temp
End Property
Public Property Get Definir_BonusVie() As Double
    Definir_BonusVie = Int(BonusMaxVie - MaxVie * CoefRaceVie)
End Property
Public Property Get Definir_MaxEnergie() As Double
    Definir_MaxEnergie = BonusMaxEnergie
End Property
Public Property Let Definir_MaxEnergie(ByVal Valeur As Double)
     MaxEnergie = Valeur
End Property
Public Property Get Definir_BonusEnergie() As Double
    Definir_BonusEnergie = Int(BonusMaxEnergie - MaxEnergie * CoefRaceEnergie)
End Property
Public Property Get Definir_MaxMagie() As Double
    Definir_MaxMagie = BonusMaxMagie
End Property
Public Property Let Definir_MaxMagie(ByVal Valeur As Double)
     MaxMagie = Valeur
End Property
Public Property Get Definir_BonusMagie() As Double
    Definir_BonusMagie = Int(BonusMaxMagie - MaxMagie * CoefRaceMagie)
End Property
Public Property Get Definir_MaxMoral() As Double
    Definir_MaxMoral = BonusMaxMoral
End Property
Public Property Let Definir_MaxMoral(ByVal Valeur As Double)
     MaxMoral = Valeur
End Property
Public Property Get Definir_BonusMoral() As Double
    Definir_BonusMoral = Int(BonusMaxMoral - MaxMoral * CoefRaceMoral)
End Property

Public Property Get Definir_Vie() As Double
    If Vie + RecuperationVie > Definir_MaxVie Then
        Definir_Vie = Definir_MaxVie
    Else
        Definir_Vie = Vie + RecuperationVie
    End If
End Property
Public Property Let Definir_Vie(ByVal Valeur As Double)
    Dim i As Integer
    If Vivant And Valeur <= 0 Then
        Valeur = 0
        Vivant = False
    End If
    If Valeur < Vie Then
        Gestion_Niveau MaxVie, (Vie - Valeur) * MaxVie / BonusMaxVie, ExpVie, NivVie, ExpNivVie, ExpNivPlusVie, 0
        'Il gagne de l'expérience avec les objets qu'il utilise.
        If Definir_BonusVie > 0 Then
            For i = 0 To UBound(ObjetsEquipes())
                If ObjetsEquipes(i).BonusVie > 0 Then
                    Gestion_Niveau CaracObjets(ObjetsEquipes(i).Competence), _
                                   (Vie - Valeur) * MaxVie / BonusMaxVie * CoefExpCompObjetsVie, _
                                   ExpObjets(ObjetsEquipes(i).Competence), NivObjets(ObjetsEquipes(i).Competence), ExpNivObjets(ObjetsEquipes(i).Competence), ExpNivPlusObjets(ObjetsEquipes(i).Competence), 8
                End If
            Next i
        End If
    End If
    Vie = Valeur
    If Vie > Definir_MaxVie Then
        Vie = Definir_MaxVie
    End If
End Property
Public Property Get Definir_Energie() As Double
    If Energie + RecuperationEnergie > Definir_MaxEnergie Then
        Definir_Energie = Definir_MaxEnergie
    Else
        Definir_Energie = Energie + RecuperationEnergie
    End If
End Property
Public Property Let Definir_Energie(ByVal Valeur As Double)
    Dim i As Integer
    If Valeur < Energie Then
        Gestion_Niveau MaxEnergie, Energie - Valeur, ExpEnergie, NivEnergie, ExpNivEnergie, ExpNivPlusEnergie, 1
        If Definir_BonusEnergie > 0 Then
            'Il gagne de l'expérience avec les objets qu'il utilise.
            For i = 0 To UBound(ObjetsEquipes())
                If ObjetsEquipes(i).BonusEnergie > 0 Then
                    Gestion_Niveau CaracObjets(ObjetsEquipes(i).Competence), _
                                   (Energie - Valeur) * CoefExpCompObjetsEnergie, _
                                   ExpObjets(ObjetsEquipes(i).Competence), NivObjets(ObjetsEquipes(i).Competence), ExpNivObjets(ObjetsEquipes(i).Competence), ExpNivPlusObjets(ObjetsEquipes(i).Competence), 8
                End If
            Next i
        End If
    End If
    Energie = Valeur
    If Energie < 0 Then
        Energie = 0
    ElseIf Energie > Definir_MaxEnergie Then
        Energie = Definir_MaxEnergie
    End If
End Property
Public Property Get Definir_Magie() As Double
    If Magie + RecuperationMagie > Definir_MaxMagie Then
        Definir_Magie = Definir_MaxMagie
    Else
        Definir_Magie = Magie + RecuperationMagie
    End If
End Property
Public Property Let Definir_Magie(ByVal Valeur As Double)
    Dim i As Integer
    If Valeur < Magie Then
        Gestion_Niveau MaxMagie, (Magie - Valeur) * MaxMagie / BonusMaxMagie, ExpMagie, NivMagie, ExpNivMagie, ExpNivPlusMagie, 2
        If Definir_BonusMagie > 0 Then
            'Il gagne de l'expérience avec les objets qu'il utilise.
            For i = 0 To UBound(ObjetsEquipes())
                If ObjetsEquipes(i).BonusMagie > 0 Then
                    Gestion_Niveau CaracObjets(ObjetsEquipes(i).Competence), _
                                   (Magie - Valeur) * MaxMagie / BonusMaxMagie * CoefExpCompObjetsMagie, _
                                   ExpObjets(ObjetsEquipes(i).Competence), NivObjets(ObjetsEquipes(i).Competence), ExpNivObjets(ObjetsEquipes(i).Competence), ExpNivPlusObjets(ObjetsEquipes(i).Competence), 8
                End If
            Next i
        End If
    End If
    Magie = Valeur
    If Magie < 0 Then
        Magie = 0
    ElseIf Magie > Definir_MaxMagie Then
        Magie = Definir_MaxMagie
    End If
End Property
Public Property Get Definir_Moral() As Double
    If Moral + RecuperationMoral > Definir_MaxMoral Then
        Definir_Moral = Definir_MaxMoral
    Else
        Definir_Moral = Moral + RecuperationMoral
    End If
End Property
Public Property Let Definir_Moral(ByVal Valeur As Double)
    Dim i As Integer
    If Valeur < Moral Then
        Gestion_Niveau MaxMoral, Moral - Valeur, ExpMoral, NivMoral, ExpNivMoral, ExpNivPlusMoral, 3
        If Definir_BonusMoral > 0 Then
            'Il gagne de l'expérience avec les objets qu'il utilise.
            For i = 0 To UBound(ObjetsEquipes())
                If ObjetsEquipes(i).BonusMoral > 0 Then
                    Gestion_Niveau CaracObjets(ObjetsEquipes(i).Competence), _
                                   (Moral - Valeur) * CoefExpCompObjetsMoral, _
                                   ExpObjets(ObjetsEquipes(i).Competence), NivObjets(ObjetsEquipes(i).Competence), ExpNivObjets(ObjetsEquipes(i).Competence), ExpNivPlusObjets(ObjetsEquipes(i).Competence), 8
                End If
            Next i
        End If
    End If
    Moral = Valeur
    If Moral < 0 Then
        Moral = 0
    ElseIf Moral > Definir_MaxMoral Then
        Moral = Definir_MaxMoral
    End If
End Property
Public Property Get Definir_Attaque() As Double
    Definir_Attaque = Attaque
End Property
Public Property Let Definir_Attaque(ByVal Valeur As Double)
    Dim i As Integer
    If Valeur < 0 Then
        Valeur = 0
    ElseIf Valeur > Definir_MaxAttaque Then
        Valeur = Definir_MaxAttaque
    End If
    If Valeur < Attaque Then
        Gestion_Niveau MaxAttaque, (Attaque - Valeur) * MaxAttaque / BonusMaxAttaque, ExpAttaque, NivAttaque, ExpNivAttaque, ExpNivPlusAttaque, 4
        'Il gagne de l'expérience avec les objets qu'il utilise.
        If Definir_BonusAttaque > 0 Then
            For i = 0 To UBound(ObjetsEquipes())
                If ObjetsEquipes(i).BonusAttaque > 0 Then
                    Gestion_Niveau CaracObjets(ObjetsEquipes(i).Competence), _
                                   (Attaque - Valeur) * MaxAttaque / BonusMaxAttaque * CoefExpCompObjetsAttaque, _
                                   ExpObjets(ObjetsEquipes(i).Competence), NivObjets(ObjetsEquipes(i).Competence), ExpNivObjets(ObjetsEquipes(i).Competence), ExpNivPlusObjets(ObjetsEquipes(i).Competence), 8
                End If
            Next i
        End If
    End If
    Attaque = Valeur
End Property
Public Property Get Definir_PorteeAttaque() As Double
    Definir_PorteeAttaque = BonusPorteeAttaque
End Property
Public Property Get Definir_MaxAttaque() As Double
    Definir_MaxAttaque = BonusMaxAttaque
End Property
Public Property Let Definir_MaxAttaque(ByVal Valeur As Double)
    MaxAttaque = Valeur
End Property
Public Property Get Definir_BonusAttaque() As Double
    Definir_BonusAttaque = Int(BonusMaxAttaque - MaxAttaque * CoefRaceAttaque)
End Property
Public Property Get Definir_Defense() As Double
    Definir_Defense = Defense
End Property
Public Property Let Definir_Defense(ByVal Valeur As Double)
    Dim i As Integer
    If Valeur < Definir_Armure Then
        Valeur = Definir_Armure
    End If
    If Valeur > Definir_MaxDefense Then
        Valeur = Definir_MaxDefense
    End If
    If Valeur < Defense Then
        Gestion_Niveau MaxDefense, (Defense - Valeur) * MaxDefense / BonusMaxDefense, ExpDefense, NivDefense, ExpNivDefense, ExpNivPlusDefense, 5
        If Definir_BonusDefense > 0 Then
            'Il gagne de l'expérience avec les objets qu'il utilise.
            For i = 0 To UBound(ObjetsEquipes())
                If ObjetsEquipes(i).BonusDefense > 0 Then
                    Gestion_Niveau CaracObjets(ObjetsEquipes(i).Competence), _
                                   (Defense - Valeur) * MaxDefense / BonusMaxDefense * CoefExpCompObjetsDefense, _
                                   ExpObjets(ObjetsEquipes(i).Competence), NivObjets(ObjetsEquipes(i).Competence), ExpNivObjets(ObjetsEquipes(i).Competence), ExpNivPlusObjets(ObjetsEquipes(i).Competence), 8
                End If
            Next i
        End If
    End If
    Defense = Valeur
End Property
Public Property Get Definir_MaxDefense() As Double
    Definir_MaxDefense = BonusMaxDefense
End Property
Public Property Let Definir_MaxDefense(ByVal Valeur As Double)
    MaxDefense = Valeur
End Property
Public Property Get Definir_BonusDefense() As Double
    Definir_BonusDefense = Int(BonusMaxDefense - MaxDefense * CoefRaceDefense)
End Property
Public Property Get Definir_BonusDeplacement() As Double
    Definir_BonusDeplacement = BonusDeplacement - LongueurPas
End Property

Public Property Get Definir_Armure() As Integer
    Definir_Armure = BonusArmure
End Property
Public Property Let Definir_Armure(ByVal Valeur As Integer)
    BonusArmure = Valeur
End Property

' ************************************************************************************
' *                              Manipulation des niveaux                            *
' ************************************************************************************

Public Sub Gestion_Tous_Niveaux()
    Dim i As Integer
    'ExpNivPlusVie = CoefExpvie
    'ExpNivPlusEnergie = CoefExpenergie
    'ExpNivPlusMagie = CoefExpmagie
    'ExpNivPlusMoral = CoefExpmoral
    'ExpNivPlusAttaque = CoefExpattaque
    'ExpNivPlusDefense = CoefExpdefense
    
    Gestion_Niveau MaxVie, 0, ExpVie, NivVie, ExpNivVie, ExpNivPlusVie, 0
    Gestion_Niveau MaxEnergie, 0, ExpEnergie, NivEnergie, ExpNivEnergie, ExpNivPlusEnergie, 1
    Gestion_Niveau MaxMagie, 0, ExpMagie, NivMagie, ExpNivMagie, ExpNivPlusMagie, 2
    Gestion_Niveau MaxMoral, 0, ExpMoral, NivMoral, ExpNivMoral, ExpNivPlusMoral, 3
    Gestion_Niveau MaxAttaque, 0, ExpAttaque, NivAttaque, ExpNivAttaque, ExpNivPlusAttaque, 4
    Gestion_Niveau MaxDefense, 0, ExpDefense, NivDefense, ExpNivDefense, ExpNivPlusDefense, 5
    'Compétences.
    For i = 0 To Nombre_CompetencesSpeciales - 1
        Gestion_Niveau CaracSpeciales(i), 0, ExpSpeciales(i), NivSpeciales(i), ExpNivSpeciales(i), ExpNivPlusSpeciales(i), 9
    Next i
    For i = 0 To Nombre_CompetencesRessources - 1
        Gestion_Niveau MaxRessources(i), 0, ExpRessources(i), NivRessources(i), ExpNivRessources(i), ExpNivPlusRessources(i), 6
    Next i
    For i = 0 To Nombre_CompetencesServices - 1
        Gestion_Niveau CaracServices(i), 0, ExpServices(i), NivServices(i), ExpNivServices(i), ExpNivPlusServices(i), 7
    Next i
    For i = 0 To Nombre_CompetencesObjets - 1
        Gestion_Niveau CaracObjets(i), 0, ExpObjets(i), NivObjets(i), ExpNivObjets(i), ExpNivPlusObjets(i), 8
    Next i
    'Vie = MaxVie
    'Energie = MaxEnergie
    'Magie = MaxMagie
    'Moral = MaxMoral
    'Attaque = MaxAttaque
    'Defense = MaxDefense
End Sub

'Public Sub Gestion_Niveau_Desctruction_Batiment(ByVal Score As Double, _
'                                                ByVal Parametres As ClsJeuParam)
'    Dim i As Integer
'    Gestion_Niveau CaracSpeciales(1), _
'                   Score * CoefExpCompDestruction * CoefExpCompSpeciales, _
'                   ExpSpeciales(1), NivSpeciales(1), ExpNivSpeciales(1), ExpNivPlusSpeciales(1), 5
'    'Il gagne de l'expérience avec les objets qu'il utilise et qui lui offrent un bonus.
'    If Bonus_CompetenceSpeciales(1) Then
'        For i = 0 To UBound(ObjetsEquipes())
'            If ObjetsEquipes(i).Param_BonusCompetencesSpeciales(1) > 0 Then
'                Gestion_Niveau CaracObjets(ObjetsEquipes(i).Competence), _
'                               Score * CoefExpCompDestruction * CoefExpCompSpeciales, _
'                               ExpObjets(ObjetsEquipes(i).Competence), NivObjets(ObjetsEquipes(i).Competence), ExpNivObjets(ObjetsEquipes(i).Competence), ExpNivPlusObjets(ObjetsEquipes(i).Competence), 4
'            End If
'        Next i
'    End If
'End Sub

'Public Sub Gestion_Niveau_Coup_Critique(ByVal Score As Double, _
'                                        ByVal Parametres As ClsJeuParam)
'    Dim i As Integer
'    Gestion_Niveau CaracSpeciales(2), _
'                   Score * CoefExpCompCoupCritique * CoefExpCompSpeciales / BonusMaxAttaque, _
'                   ExpSpeciales(2), NivSpeciales(2), ExpNivSpeciales(2), ExpNivPlusSpeciales(2), 5
'    'Il gagne de l'expérience avec les objets qu'il utilise et qui lui offrent un bonus.
'    If Bonus_CompetenceSpeciales(2) Then
'        For i = 0 To UBound(ObjetsEquipes())
'            If ObjetsEquipes(i).Param_BonusCompetencesSpeciales(2) > 0 Then
'                Gestion_Niveau CaracObjets(ObjetsEquipes(i).Competence), _
'                               Score * CoefExpCompCoupCritique * CoefExpCompSpeciales, _
'                               ExpObjets(ObjetsEquipes(i).Competence), NivObjets(ObjetsEquipes(i).Competence), ExpNivObjets(ObjetsEquipes(i).Competence), ExpNivPlusObjets(ObjetsEquipes(i).Competence), 4
'            End If
'        Next i
'    End If
'End Sub

Public Sub Gestion_Niveau_Comp_Speciale(ByVal Indice As Integer, _
                                        ByVal Score As Double, _
                                        ByVal Parametres As ClsJeuParam)
    Dim i As Integer
    Dim CoefExpComp As Double
    Dim CoefExpCompObjet As Double
    Select Case Indice
    Case 0:
        CoefExpComp = CoefExpCompConstruction
        CoefExpCompObjet = CoefExpCompConstructionObjet
    Case 1:
        CoefExpComp = CoefExpCompDestruction
        CoefExpCompObjet = CoefExpCompDestructionObjet
    Case 2:
        If Score > Attaque Then Score = Attaque
        'Debug.Print Score
        CoefExpComp = CoefExpCompCoupCritique
        CoefExpCompObjet = CoefExpCompCoupCritiqueObjet
    Case 3:
        CoefExpComp = CoefExpCompMarchandage
        CoefExpCompObjet = CoefExpCompMarchandageObjet
    End Select
    Gestion_Niveau CaracSpeciales(Indice), _
                   Score * CoefExpComp * CoefExpCompSpeciales, _
                   ExpSpeciales(Indice), NivSpeciales(Indice), ExpNivSpeciales(Indice), ExpNivPlusSpeciales(Indice), 9
    'Il gagne de l'expérience avec les objets qu'il utilise et qui lui offrent un bonus.
    If Bonus_CompetenceSpeciales(Indice) Then
        For i = 0 To UBound(ObjetsEquipes())
            If ObjetsEquipes(i).Param_BonusCompetencesSpeciales(Indice) > 0 Then
                Gestion_Niveau CaracObjets(ObjetsEquipes(i).Competence), _
                               Score * CoefExpComp * CoefExpCompSpeciales * CoefExpCompObjet, _
                               ExpObjets(ObjetsEquipes(i).Competence), NivObjets(ObjetsEquipes(i).Competence), ExpNivObjets(ObjetsEquipes(i).Competence), ExpNivPlusObjets(ObjetsEquipes(i).Competence), 8
            End If
        Next i
    End If
End Sub

Public Function Prix_Marchandee(ByVal Prix As Double, _
                                ByVal Parametres As ClsJeuParam) As Integer
    Prix_Marchandee = Prix / 2 * (1 + 1 * Parametres.PersosNiveauSpecialesDepart / Definir_Niveau_CompetenceSpeciales(3))
End Function

Public Sub Gestion_Niveau_Prime(ByVal Score As Double, _
                                ByVal Parametres As ClsJeuParam)
    Dim i As Integer
    Gestion_Niveau CaracServices(1), _
                   Score * Parametres.Service_CoefExp(1 + 1), _
                   ExpServices(1), NivServices(1), ExpNivServices(1), ExpNivPlusServices(1), 7
    If Bonus_CompetenceServices(1) Then
        For i = 0 To UBound(ObjetsEquipes())
            If ObjetsEquipes(i).Param_BonusCompetencesServices(1) > 0 Then
                Gestion_Niveau CaracObjets(ObjetsEquipes(i).Competence), _
                               Score * Parametres.Service_CoefExp(2), _
                               ExpObjets(ObjetsEquipes(i).Competence), NivObjets(ObjetsEquipes(i).Competence), ExpNivObjets(ObjetsEquipes(i).Competence), ExpNivPlusObjets(ObjetsEquipes(i).Competence), 8
            End If
        Next i
    End If
End Sub

Public Sub Gestion_Niveau_Resistances(ByVal Indice As Integer, _
                                      ByVal Score As Double, _
                                      ByVal Parametres As ClsJeuParam)
    Dim i As Integer
    'Il gagne de l'expérience avec les objets qu'il utilise et qui lui offrent un bonus en résistance.
    If Bonus_Resistance(Indice) Then
        For i = 0 To UBound(ObjetsEquipes())
            If ObjetsEquipes(i).Actif Then
                If Parametres.Objet_BonusResistance(ObjetsEquipes(i).Numero, Indice) > 0 Then
                    Gestion_Niveau CaracObjets(Parametres.Objet_Competence(ObjetsEquipes(i).Numero)), _
                                   Score * MaxVie / BonusMaxVie * CoefExpCompObjetsResistance * Parametres.Objet_BonusResistance(ObjetsEquipes(i).Numero, Indice), _
                                   ExpObjets(ObjetsEquipes(i).Competence), NivObjets(ObjetsEquipes(i).Competence), ExpNivObjets(ObjetsEquipes(i).Competence), ExpNivPlusObjets(ObjetsEquipes(i).Competence), 8
                End If
            End If
        Next i
    End If
End Sub

Private Sub Gestion_Niveau(ByRef Carac As Long, _
                           ByVal Gain As Double, _
                           ByRef Exp As Double, _
                           ByRef Niv As Long, _
                           ByRef ExpNiv As Double, _
                           ByRef ExpNivPlus As Double, _
                           ByVal TypeCarac As Integer)
    If GainExperience Then
        Exp = Exp + Gain
        While Exp >= ExpNivPlus 'On change de niveau.
            Carac = Carac + 1
            Niv = Niv + 1
            ExpNiv = ExpNivPlus
            Select Case TypeCarac
            Case 0:
                ExpNivPlus = ExpNivPlus + ((Niv + 1) * CoefExpVie)
            Case 1:
                ExpNivPlus = ExpNivPlus + ((Niv + 1) * CoefExpEnergie)
            Case 2:
                ExpNivPlus = ExpNivPlus + ((Niv + 1) * CoefExpMagie)
            Case 3:
                ExpNivPlus = ExpNivPlus + ((Niv + 1) * CoefExpMoral)
            Case 4:
                ExpNivPlus = ExpNivPlus + ((Niv + 1) * CoefExpAttaque)
            Case 5:
                ExpNivPlus = ExpNivPlus + ((Niv + 1) * CoefExpDefense)
            Case 6:
                ExpNivPlus = ExpNivPlus + ((Niv + 1) * CoefExpCompRessources)
            Case 7:
                ExpNivPlus = ExpNivPlus + ((Niv + 1) * CoefExpCompServices)
            Case 8:
                ExpNivPlus = ExpNivPlus + ((Niv + 1) * CoefExpCompObjets)
            Case 9:
                ExpNivPlus = ExpNivPlus + ((Niv + 1) * CoefExpCompSpeciales)
            End Select
            'Debug.Print ExpNiv & "/" & ExpNivPlus
            Actualiser_Bonus
        Wend
    End If
End Sub
Public Property Get Progression_Vie() As Double
    Progression_Vie = (ExpVie - ExpNivVie) / (ExpNivPlusVie - ExpNivVie)
End Property
Public Property Get Progression_Energie() As Double
    Progression_Energie = (ExpEnergie - ExpNivEnergie) / (ExpNivPlusEnergie - ExpNivEnergie)
End Property
Public Property Get Progression_Magie() As Double
    Progression_Magie = (ExpMagie - ExpNivMagie) / (ExpNivPlusMagie - ExpNivMagie)
End Property
Public Property Get Progression_Moral() As Double
    Progression_Moral = (ExpMoral - ExpNivMoral) / (ExpNivPlusMoral - ExpNivMoral)
End Property
Public Property Get Progression_Attaque() As Double
    Progression_Attaque = (ExpAttaque - ExpNivAttaque) / (ExpNivPlusAttaque - ExpNivAttaque)
End Property
Public Property Get Progression_Defense() As Double
    Progression_Defense = (ExpDefense - ExpNivDefense) / (ExpNivPlusDefense - ExpNivDefense)
End Property
Public Property Get Progression_CompetenceSpeciale(ByVal Index As Integer) As Double
    Progression_CompetenceSpeciale = (ExpSpeciales(Index) - ExpNivSpeciales(Index)) / (ExpNivPlusSpeciales(Index) - ExpNivSpeciales(Index))
End Property
Public Property Get Progression_CompetenceRessource(ByVal Index As Integer) As Double
    Progression_CompetenceRessource = (ExpRessources(Index) - ExpNivRessources(Index)) / (ExpNivPlusRessources(Index) - ExpNivRessources(Index))
End Property
Public Property Get Progression_CompetenceService(ByVal Index As Integer) As Double
    Progression_CompetenceService = (ExpServices(Index) - ExpNivServices(Index)) / (ExpNivPlusServices(Index) - ExpNivServices(Index))
End Property
Public Property Get Progression_CompetenceObjet(ByVal Index As Integer) As Double
    Progression_CompetenceObjet = (ExpObjets(Index) - ExpNivObjets(Index)) / (ExpNivPlusObjets(Index) - ExpNivObjets(Index))
End Property

'Autres proriétés
Public Property Get Definir_ressources(ByVal Index As Long) As Double
    If Index >= 0 Then
        Definir_ressources = Ressources(Index)
    End If
End Property
Public Property Let Definir_ressources(ByVal Index As Long, ByVal Valeur As Double)
    Ressources(Index) = Valeur
    If Ressources(Index) > Definir_MaxRessources(Index) Then
       Ressources(Index) = Definir_MaxRessources(Index)
    ElseIf Ressources(Index) < 0 Then
        Ressources(Index) = 0
    End If
End Property
Public Property Get Definir_RessourcesTotales() As Double
    'Renvoie le nombre total de ressources.
    Dim i As Integer
    For i = 0 To Parametres.NombreRessources - 1
        Definir_RessourcesTotales = Definir_RessourcesTotales + Ressources(i)
    Next i
End Property

Public Property Get Inactif() As Boolean
    Inactif = Action = 0 Or Action = 100
End Property

Public Property Get IA_Berserk() As Integer
    IA_Berserk = BonusBerserk
End Property
Public Property Let IA_Berserk(ByVal Valeur As Integer)
    Berserk = Valeur
End Property

Private Sub Calculer_Vitalite()
    'Cette fonction renvoie la productivite du personnage de 0 à 1 en fonction de l'énergie et du moral.
    If Vivant Then
        Vitalite = BonusVitalite
        Select Case Vie / Definir_MaxVie
            Case Is < 0.1: Vitalite = Vitalite - 0.45
            Case Is < 0.25: Vitalite = Vitalite - 0.25
            Case Is < 0.5: Vitalite = Vitalite - 0.125
        End Select
        Select Case Energie / Definir_MaxEnergie
            Case Is < 0.1: Vitalite = Vitalite - 0.45
            Case Is < 0.25: Vitalite = Vitalite - 0.25
            Case Is < 0.5: Vitalite = Vitalite - 0.125
        End Select
        Select Case Moral / Definir_MaxMoral
            Case Is < 0.1: Vitalite = Vitalite - 0.45
            Case Is < 0.25: Vitalite = Vitalite - 0.25
            Case Is < 0.5: Vitalite = Vitalite - 0.125
        End Select
        If Vitalite < 0.1 Then Vitalite = 0.1
    Else
        Vitalite = 0
    End If
End Sub

'Propriétés des attaques.
Public Property Get Effet_Attaque() As Integer
    If ObjetsEquipes(0).Actif Then
        Effet_Attaque = ObjetsEquipes(0).EffetAttaque
    Else
        Effet_Attaque = EffetAttaque
    End If
End Property

Public Property Get Definir_NombreAttaqueCauseEffet() As Integer
    Definir_NombreAttaqueCauseEffet = BonusAttaqueCauseEffet.Count
End Property
Public Property Get Definir_AttaqueCauseEffet(ByVal Index As Integer) As Integer
    Definir_AttaqueCauseEffet = BonusAttaqueCauseEffet.Item(Index)
End Property
Public Property Get Definir_AttaqueResistance() As Integer
    Definir_AttaqueResistance = BonusAttaqueResistance
End Property

' ************************************************************************************
' *                                  Compétences                                     *
' ************************************************************************************

'Propriétés des comptétences Speciales.
Public Property Get Nombre_CompetencesSpeciales() As Integer
    Nombre_CompetencesSpeciales = UBound(CaracSpeciales()) + 1
End Property
Public Property Get Definir_ExpSpeciales(ByVal Index As Long) As Double
    If Index >= 0 Then
        Definir_ExpSpeciales = ExpSpeciales(Index)
    End If
End Property
Public Property Let Definir_ExpSpeciales(ByVal Index As Long, ByVal Valeur As Double)
    ExpSpeciales(Index) = Valeur
End Property
Public Property Get Definir_ExpNivSpeciales(ByVal Index As Long) As Double
    Definir_ExpNivSpeciales = ExpNivSpeciales(Index)
End Property
Public Property Let Definir_ExpNivSpeciales(ByVal Index As Long, ByVal Valeur As Double)
    ExpNivSpeciales(Index) = Valeur
End Property
Public Property Get Definir_ExpNivPlusSpeciales(ByVal Index As Long) As Double
    Definir_ExpNivPlusSpeciales = ExpNivPlusSpeciales(Index)
End Property
Public Property Let Definir_ExpNivPlusSpeciales(ByVal Index As Long, ByVal Valeur As Double)
    ExpNivPlusSpeciales(Index) = Valeur
End Property
Public Property Get Definir_Niveau_CompetenceSpeciales(ByVal Index As Long) As Integer
    Definir_Niveau_CompetenceSpeciales = BonusNivSpeciales(Index)
End Property
Public Property Let Definir_Niveau_CompetenceSpeciales(ByVal Index As Long, ByVal Valeur As Integer)
    BonusNivSpeciales(Index) = Valeur
End Property
Public Property Get Niveau_CompetenceSpeciales(ByVal Index As Long) As Integer
    Niveau_CompetenceSpeciales = NivSpeciales(Index)
End Property
Public Property Let Niveau_CompetenceSpeciales(ByVal Index As Long, ByVal Valeur As Integer)
    NivSpeciales(Index) = Valeur
End Property
Public Property Get Definir_Carac_CompetenceSpeciales(ByVal Index As Long) As Integer
    Definir_Carac_CompetenceSpeciales = CaracSpeciales(Index)
End Property
Public Property Let Definir_Carac_CompetenceSpeciales(ByVal Index As Long, ByVal Valeur As Integer)
    CaracSpeciales(Index) = Valeur
End Property
Public Property Get Bonus_CompetenceSpeciales(ByVal Index As Long) As Boolean
    Bonus_CompetenceSpeciales = BonusNivSpeciales(Index) > CaracSpeciales(Index) * CoefRaceCompSpeciales(Index)
End Property

'Propriété des ressources.
Public Property Get Nombre_CompetencesRessources() As Integer
    Nombre_CompetencesRessources = UBound(MaxRessources()) + 1
End Property
Public Property Get Definir_Carac_Ressources(ByVal Index As Long) As Long
    Definir_Carac_Ressources = MaxRessources(Index)
End Property
Public Property Let Definir_Carac_Ressources(ByVal Index As Long, ByVal Valeur As Long)
    MaxRessources(Index) = Valeur
End Property
Public Property Get Definir_MaxRessources(ByVal Index As Long) As Double
    If Index >= 0 Then
        Definir_MaxRessources = MaxRessources(Index) + BonusMaxRessourcesObjet
    End If
End Property
Public Property Let Definir_MaxRessources(ByVal Index As Long, ByVal Valeur As Double)
    MaxRessources(Index) = Valeur
End Property
Public Property Get Definir_ExpRessources(ByVal Index As Long) As Double
    If Index >= 0 Then
        Definir_ExpRessources = ExpRessources(Index)
    End If
End Property
Public Property Let Definir_ExpRessources(ByVal Index As Long, ByVal Valeur As Double)
    ExpRessources(Index) = Valeur
End Property
Public Property Get Definir_ExpNivRessources(ByVal Index As Long) As Double
    Definir_ExpNivRessources = ExpNivRessources(Index)
End Property
Public Property Let Definir_ExpNivRessources(ByVal Index As Long, ByVal Valeur As Double)
    ExpNivRessources(Index) = Valeur
End Property
Public Property Get Definir_ExpNivPlusRessources(ByVal Index As Long) As Double
    Definir_ExpNivPlusRessources = ExpNivPlusRessources(Index)
End Property
Public Property Let Definir_ExpNivPlusRessources(ByVal Index As Long, ByVal Valeur As Double)
    ExpNivPlusRessources(Index) = Valeur
End Property
Public Property Get Definir_Niveau_CompetenceRessources(ByVal Index As Long) As Integer
    Definir_Niveau_CompetenceRessources = BonusNivRessources(Index)
End Property
Public Property Let Definir_Niveau_CompetenceRessources(ByVal Index As Long, ByVal Valeur As Integer)
    BonusNivRessources(Index) = Valeur
End Property
Public Property Get Niveau_CompetenceRessources(ByVal Index As Long) As Integer
    Niveau_CompetenceRessources = NivRessources(Index)
End Property
Public Property Let Niveau_CompetenceRessources(ByVal Index As Long, ByVal Valeur As Integer)
    NivRessources(Index) = Valeur
End Property
Public Property Get Bonus_CompetenceRessources(ByVal Index As Long) As Boolean
    Bonus_CompetenceRessources = BonusNivRessources(Index) + Parametres.PersosMaxRessources > (NivRessources(Index) + Parametres.PersosMaxRessources) * CoefRaceCompRessources(Index)
End Property

'Propriétés des comptétences de services.
Public Property Get Nombre_CompetencesServices() As Integer
    Nombre_CompetencesServices = UBound(CaracServices()) + 1
End Property
Public Property Get Definir_ExpServices(ByVal Index As Long) As Double
    If Index >= 0 Then
        Definir_ExpServices = ExpServices(Index)
    End If
End Property
Public Property Let Definir_ExpServices(ByVal Index As Long, ByVal Valeur As Double)
    ExpServices(Index) = Valeur
End Property
Public Property Get Definir_ExpNivServices(ByVal Index As Long) As Double
    Definir_ExpNivServices = ExpNivServices(Index)
End Property
Public Property Let Definir_ExpNivServices(ByVal Index As Long, ByVal Valeur As Double)
    ExpNivServices(Index) = Valeur
End Property
Public Property Get Definir_ExpNivPlusServices(ByVal Index As Long) As Double
    Definir_ExpNivPlusServices = ExpNivPlusServices(Index)
End Property
Public Property Let Definir_ExpNivPlusServices(ByVal Index As Long, ByVal Valeur As Double)
    ExpNivPlusServices(Index) = Valeur
End Property
Public Property Get Definir_Niveau_CompetenceServices(ByVal Index As Long) As Integer
    Definir_Niveau_CompetenceServices = BonusNivServices(Index)
End Property
Public Property Let Definir_Niveau_CompetenceServices(ByVal Index As Long, ByVal Valeur As Integer)
    BonusNivServices(Index) = Valeur
End Property
Public Property Get Niveau_CompetenceServices(ByVal Index As Long) As Integer
    Niveau_CompetenceServices = NivServices(Index)
End Property
Public Property Let Niveau_CompetenceServices(ByVal Index As Long, ByVal Valeur As Integer)
    NivServices(Index) = Valeur
End Property
Public Property Get Definir_Carac_CompetenceServices(ByVal Index As Long) As Integer
    Definir_Carac_CompetenceServices = CaracServices(Index)
End Property
Public Property Let Definir_Carac_CompetenceServices(ByVal Index As Long, ByVal Valeur As Integer)
    CaracServices(Index) = Valeur
End Property
Public Property Get Bonus_CompetenceServices(ByVal Index As Long) As Boolean
    Bonus_CompetenceServices = BonusNivServices(Index) > CaracServices(Index) * CoefRaceCompServices(Index)
End Property

'Propriétés des comptétences d'objets.
Public Property Get Nombre_CompetencesObjets() As Integer
    Nombre_CompetencesObjets = UBound(CaracObjets()) + 1
End Property
Public Property Get Definir_ExpObjets(ByVal Index As Long) As Double
    If Index >= 0 Then
        Definir_ExpObjets = ExpObjets(Index)
    End If
End Property
Public Property Let Definir_ExpObjets(ByVal Index As Long, ByVal Valeur As Double)
    ExpObjets(Index) = Valeur
End Property
Public Property Get Definir_ExpNivObjets(ByVal Index As Long) As Double
    Definir_ExpNivObjets = ExpNivObjets(Index)
End Property
Public Property Let Definir_ExpNivObjets(ByVal Index As Long, ByVal Valeur As Double)
    ExpNivObjets(Index) = Valeur
End Property
Public Property Get Definir_ExpNivPlusObjets(ByVal Index As Long) As Double
    Definir_ExpNivPlusObjets = ExpNivPlusObjets(Index)
End Property
Public Property Let Definir_ExpNivPlusObjets(ByVal Index As Long, ByVal Valeur As Double)
    ExpNivPlusObjets(Index) = Valeur
End Property
Public Property Get Definir_Niveau_CompetenceObjets(ByVal Index As Long) As Integer
    Definir_Niveau_CompetenceObjets = BonusNivObjets(Index)
End Property
Public Property Let Definir_Niveau_CompetenceObjets(ByVal Index As Long, ByVal Valeur As Integer)
    BonusNivObjets(Index) = Valeur
End Property
Public Property Get Niveau_CompetenceObjets(ByVal Index As Long) As Integer
    Niveau_CompetenceObjets = NivObjets(Index)
End Property
Public Property Let Niveau_CompetenceObjets(ByVal Index As Long, ByVal Valeur As Integer)
    NivObjets(Index) = Valeur
End Property
Public Property Get Definir_Carac_CompetenceObjets(ByVal Index As Long) As Integer
    Definir_Carac_CompetenceObjets = CaracObjets(Index)
End Property
Public Property Let Definir_Carac_CompetenceObjets(ByVal Index As Long, ByVal Valeur As Integer)
    CaracObjets(Index) = Valeur
End Property
Public Property Get Bonus_CompetenceObjets(ByVal Index As Long) As Boolean
    Bonus_CompetenceObjets = BonusNivObjets(Index) > CaracObjets(Index) * CoefRaceCompObjets(Index)
End Property
Public Property Get Definir_Min_Competence(Optional ByVal Categorie As Integer) As Long
    Dim i As Long
    Dim Temp1 As Long, Temp2 As Long
    Select Case Categorie
    Case 0: Temp1 = Definir_Niveau_CompetenceSpeciales(0)
    Case 1: Temp1 = Niveau_CompetenceSpeciales(0) + 1
    Case 2: Temp1 = Definir_ExpSpeciales(0)
    End Select
    For i = 1 To Nombre_CompetencesSpeciales - 1
        Select Case Categorie
        Case 0: Temp2 = Definir_Niveau_CompetenceSpeciales(i)
        Case 1: Temp2 = Niveau_CompetenceSpeciales(i) + 1
        Case 2: Temp2 = Definir_ExpSpeciales(i)
        End Select
        If Temp2 < Temp1 Then
            Temp1 = Temp2
        End If
    Next i
    For i = 0 To Nombre_CompetencesRessources - 1
        Select Case Categorie
        Case 0: Temp2 = Persos(NoPerso).Definir_MaxRessources(i)
        Case 1: Temp2 = Persos(NoPerso).Niveau_CompetenceRessources(i) + 1
        Case 2: Temp2 = Persos(NoPerso).Definir_ExpRessources(i)
        End Select
        If Temp2 < Temp1 Then
            Temp1 = Temp2
        End If
    Next i
    For i = 0 To Nombre_CompetencesServices - 1
        Select Case Categorie
        Case 0: Temp2 = Definir_Niveau_CompetenceServices(i)
        Case 1: Temp2 = Niveau_CompetenceServices(i) + 1
        Case 2: Temp2 = Definir_ExpServices(i)
        End Select
        If Temp2 < Temp1 Then
            Temp1 = Temp2
        End If
    Next i
    For i = 1 To Nombre_CompetencesObjets - 1
        Select Case Categorie
        Case 0: Temp2 = Definir_Niveau_CompetenceObjets(i)
        Case 1: Temp2 = Niveau_CompetenceObjets(i) + 1
        Case 2: Temp2 = Definir_ExpObjets(i)
        End Select
        If Temp2 < Temp1 Then
            Temp1 = Temp2
        End If
    Next i
    Definir_Min_Competence = Temp1
End Property
Public Property Get Definir_Moy_Competence(Optional ByVal Categorie As Integer) As Double
    Dim i As Long
    Dim Temp As Long
    For i = 0 To Nombre_CompetencesSpeciales - 1
        Select Case Categorie
        Case 0: Temp = Temp + Definir_Niveau_CompetenceSpeciales(i)
        Case 1: Temp = Temp + Niveau_CompetenceSpeciales(i) + 1
        Case 2: Temp = Temp + Definir_ExpSpeciales(i)
        End Select
    Next i
    For i = 0 To Nombre_CompetencesRessources - 1
        Select Case Categorie
        Case 0: Temp = Temp + Persos(NoPerso).Definir_MaxRessources(i)
        Case 1: Temp = Temp + Persos(NoPerso).Niveau_CompetenceRessources(i) + 1
        Case 2: Temp = Temp + Persos(NoPerso).Definir_ExpRessources(i)
        End Select
    Next i
    For i = 0 To Nombre_CompetencesServices - 1
        Select Case Categorie
        Case 0: Temp = Temp + Definir_Niveau_CompetenceServices(i)
        Case 1: Temp = Temp + Niveau_CompetenceServices(i) + 1
        Case 2: Temp = Temp + Definir_ExpServices(i)
        End Select
    Next i
    For i = 1 To Nombre_CompetencesObjets - 1
        Select Case Categorie
        Case 0: Temp = Temp + Definir_Niveau_CompetenceObjets(i)
        Case 1: Temp = Temp + Niveau_CompetenceObjets(i) + 1
        Case 2: Temp = Temp + Definir_ExpObjets(i)
        End Select
    Next i
    Definir_Moy_Competence = Temp / (Nombre_CompetencesSpeciales + Nombre_CompetencesRessources + Nombre_CompetencesServices + Nombre_CompetencesObjets - 1)
End Property
Public Property Get Definir_Max_Competence(Optional ByVal Categorie As Integer) As Long
    Dim i As Long
    Dim Temp1 As Long, Temp2 As Long
    For i = 0 To Nombre_CompetencesSpeciales - 1
        Select Case Categorie
        Case 0: Temp2 = Definir_Niveau_CompetenceSpeciales(i)
        Case 1: Temp2 = Niveau_CompetenceSpeciales(i) + 1
        Case 2: Temp2 = Definir_ExpSpeciales(i)
        End Select
        If Temp2 > Temp1 Then
            Temp1 = Temp2
        End If
    Next i
    For i = 0 To Nombre_CompetencesRessources - 1
        Select Case Categorie
        Case 0: Temp2 = Niveau_CompetenceRessources(i) + Parametres.PersosMaxRessources
        Case 1: Temp2 = Niveau_CompetenceRessources(i) + 1
        Case 2: Temp2 = Definir_ExpRessources(i)
        End Select
        If Temp2 > Temp1 Then
            Temp1 = Temp2
        End If
    Next i
    For i = 0 To Nombre_CompetencesServices - 1
        Select Case Categorie
        Case 0: Temp2 = Definir_Niveau_CompetenceServices(i)
        Case 1: Temp2 = Niveau_CompetenceServices(i) + 1
        Case 2: Temp2 = Definir_ExpServices(i)
        End Select
        If Temp2 > Temp1 Then
            Temp1 = Temp2
        End If
    Next i
    For i = 1 To Nombre_CompetencesObjets - 1
        Select Case Categorie
        Case 0: Temp2 = Definir_Niveau_CompetenceObjets(i)
        Case 1: Temp2 = Niveau_CompetenceObjets(i) + 1
        Case 2: Temp2 = Definir_ExpObjets(i)
        End Select
        If Temp2 > Temp1 Then
            Temp1 = Temp2
        End If
    Next i
    Definir_Max_Competence = Temp1
End Property

' ************************************************************************************
' *                               Autres propriétes                                  *
' ************************************************************************************

Public Property Let Definir_Argent(ByVal Valeur As Double)
    Argent = Valeur
    If Argent < 0 Then
        Argent = 0
    ElseIf Argent > 10000000 Then
        Argent = 10000000
    End If
End Property

Public Property Get Definir_Argent() As Double
    Definir_Argent = Argent
End Property

Public Property Get Definir_DureeImmobilisation() As Double
    Definir_DureeImmobilisation = DureeImmobilisation
End Property
Public Property Let Definir_DureeImmobilisation(ByVal Valeur As Double)
    DureeImmobilisation = Valeur
    RessourceCourante = -1
End Property

Public Sub Calculer_Longueur_Pas()
    If Vivant Then
        Longueur_Pas = Int(BonusDeplacement * Vitalite)
        If Not Definir_Volant Then 'Subit le ralentissement du relief du terrain.
            Longueur_Pas = Int(Longueur_Pas * Monde.CoefVitesse(Me, PositionX, PositionY + Hauteur - Monde.HauteurCase)) 'AffTerrain.Hauteur))
        End If
        If Longueur_Pas < 1 And BonusDeplacement <> 0 Then
            Longueur_Pas = 1 'Vitesse minimum sauf si la vitesse de base est déjà de 0.
        ElseIf Longueur_Pas > LongueurPasMax Then
            Longueur_Pas = LongueurPasMax
        End If
        'Calcules d'accélération.
        Longueur_PasTemp = Longueur_PasTemp + Acceleration
        If Longueur_PasTemp > Longueur_Pas Then
            Longueur_PasTemp = Longueur_Pas
        End If
        Longueur_Pas = Longueur_PasTemp
    Else
        Longueur_Pas = VitesseDeplacementFantome 'Tous les fantômes ont la même vitesse.
    End If
End Sub

Public Property Get Definir_VitesseDeplacementFantome() As Integer
    Definir_VitesseDeplacementFantome = VitesseDeplacementFantome
End Property

'Public Sub Calculer_Apparence(ByVal Maison As ClsJeuBatiment, _
'                              ByVal Parametres As ClsJeuParam)
'    If ObjetsEquipes(2).ApparencePerso > 0 Then
'        '###
'    End If
'End Sub


Public Property Get Visible() As Boolean
    Visible = Not (DansSaMaison Or DansUneMaison Or DansSonChateau Or DansUnChateau)
End Property

Public Property Let Visible(ByVal Valeur As Boolean)
    DansSaMaison = Not Valeur
    DansUneMaison = Not Valeur
    DansSonChateau = Not Valeur
    DansUnChateau = Not Valeur
End Property

Public Property Get Attaquable() As Boolean
    Attaquable = Vivant And Visible And Not Inattaquable
End Property

Public Property Get Vivant() As Boolean
    Vivant = Vie > 0
End Property

Public Property Let Vivant(ByVal Valeur As Boolean)
    Dim i As Integer
    If Valeur Then
        'On ressuscite le personnage.
        Arreter
        Vie = Definir_MaxVie
        Definir_Attaque = Definir_MaxAttaque
        Definir_Defense = Definir_MaxDefense
        If Invocation Then
            If Not Projectile Then
                Persos(NumeroChef).NombreSoldats = Persos(NumeroChef).NombreSoldats + 1
            End If
        Else
            Fiefs(NumeroFief).NombreMorts = Fiefs(NumeroFief).NombreMorts - 1
        End If
        If RestaurationComplete Then 'Le personnage récupère tout.
            'Récupère ses ressources.
            For i = 0 To Parametres.NombreRessources - 1
                If Definir_ressources(i) < Parametres.Race_Ressources(Race, i) Then
                    Definir_ressources(i) = Parametres.Race_Ressources(Race, i)
                End If
            Next i
            'Récupère son équipement.
            For i = 0 To Parametres.Race_NombreEquipement(Race) - 1
                If Objet_NombreType(Parametres.Race_Equipement(Race, i)) = 0 Then
                    Objet_Ajouter_Inventaire Parametres.Race_Equipement(Race, i), Parametres, ModMain.Commentaires, IA_Equiper > 0
                End If
            Next i
            'Récupère de la vie
            Energie = Definir_MaxEnergie
            Magie = Definir_MaxMagie
            Moral = Definir_MaxMoral
        ElseIf Jeu.Confrontation Then
            Energie = Definir_MaxEnergie
            Magie = Definir_MaxMagie
            Moral = Definir_MaxMoral
        Else
            If Energie = 0 Then
                Energie = Definir_MaxEnergie 'Restaure les personnages trop fatigués.
            End If
            If Moral = 0 Then
                Moral = Definir_MaxMoral
            End If
        End If
    Else
        'Le personnage est mort... snif...
        Morts = Morts + 1
        Commentaires = ""
        Arreter
        Vie = 0
        'Effet visuel.
        Activer_Effet Parametres.Race_EffetMort(Race), PositionX, PositionY, Largeur, Hauteur, Parametres
        RecuperationVie = 0
        RecuperationEnergie = 0
        RecuperationMagie = 0
        RecuperationMoral = 0
        DureeImmobilisation = 0
        DureeRessurection = 100
        'Efface certains effets temporaires.
        EffetTemp_Effacer_Mort
        'Revient dans son équipe de départ.
        Attaquer_Allies False
        If Invocation Then
            If Projectile Then
                Persos(NumeroChef).ProjectileEnCours = Persos(NumeroChef).ProjectileEnCours - 1
            Else
                Persos(NumeroChef).NombreSoldats = Persos(NumeroChef).NombreSoldats - 1
            End If
            If Maisons(Numero).Visible Then  'Les bâtiments des invocations disparaissent également.
                If Not Maisons(Numero).Construit Then
                    Maisons(Numero).Demolir Me, Parametres
                End If
            End If
        Else
            Fiefs(NumeroFief).NombreMorts = Fiefs(NumeroFief).NombreMorts + 1
        End If
        'L'empereur pert ses objet et son titre.
        If Definir_Empereur Then Definir_Empereur = False
    End If
End Property

Public Property Get Hostile() As Boolean
    Hostile = NumeroEquipe <= 0
End Property

Public Property Get EnCombat() As Boolean
    EnCombat = Action >= 300 And Action < 400
End Property

'Public Property Get EnDeplacement() As Boolean
'    EnDeplacement = DirectionY < PositionY - Me.Longueur_Pas / 2 Or _
                    DirectionY > PositionY + Me.Longueur_Pas / 2 Or _
                    DirectionX < PositionX - Me.Longueur_Pas / 2 Or _
                    DirectionX > PositionX + Me.Longueur_Pas / 2
'End Property

Public Property Get Invocation() As Boolean
    Invocation = NumeroChef >= 0
End Property
Public Property Let Invocation(ByVal Valeur As Boolean)
    If Valeur Then
        NumeroChef = 0
    Else
        NumeroChef = -1
    End If
End Property

Public Property Get Niveau() As Double
    Dim i As Integer
    'Dim Temp As Double
    'Dim CoefCarac As Integer
    
    'CoefCarac = Int((UBound(NivRessources()) + UBound(CaracServices()) + UBound(CaracObjets())) * NiveauCoefCarac)
    'Temp = (NivVie + NivEnergie + NivMagie + NivMoral + NivAttaque + NivDefense) * CoefCarac
    'For i = 0 To UBound(NivSpeciales())
    '    Temp = Temp + NivSpeciales(i)
    'Next i
    'For i = 0 To UBound(NivRessources())
    '    Temp = Temp + NivRessources(i)
    'Next i
    'For i = 0 To UBound(NivServices())
    '    Temp = Temp + NivServices(i)
    'Next i
    'For i = 0 To UBound(NivObjets())
    '    Temp = Temp + NivObjets(i)
    'Next i
    'Niveau = NiveauMultiplicateur * Temp / (6 * CoefCarac + UBound(NivRessources()) + UBound(CaracServices()) + UBound(CaracObjets())) + 1
    Niveau = (NivVie + NivEnergie + NivMagie + NivMoral + NivAttaque + NivDefense) * 0.1 + 1
    For i = 0 To UBound(NivSpeciales())
        Niveau = Niveau + NivSpeciales(i) * 0.01
    Next i
    For i = 0 To UBound(NivRessources())
        Niveau = Niveau + NivRessources(i) * 0.01
    Next i
    For i = 0 To UBound(NivServices())
        Niveau = Niveau + NivServices(i) * 0.01
    Next i
    For i = 0 To UBound(NivObjets())
        Niveau = Niveau + NivObjets(i) * 0.01
    Next i
End Property

Public Property Get Defnir_Score_Tresor() As Long
    Dim i As Integer
    'Defnir_Score_Tresor = (Niveau - 1) * 2 + Definir_Bestiaire_Total / 450
    For i = 0 To UBound(NivSpeciales())
        Defnir_Score_Tresor = Defnir_Score_Tresor + NivSpeciales(i) * 1
    Next i
    For i = 0 To UBound(NivRessources())
        Defnir_Score_Tresor = Defnir_Score_Tresor + NivRessources(i) * 1
    Next i
    For i = 0 To UBound(NivServices())
        Defnir_Score_Tresor = Defnir_Score_Tresor + NivServices(i) * 1
    Next i
    For i = 0 To UBound(NivObjets())
        Defnir_Score_Tresor = Defnir_Score_Tresor + NivObjets(i) * 1
    Next i
    Defnir_Score_Tresor = Int(Defnir_Score_Tresor / 12)
End Property

Public Property Get Metier(ByVal Parametres As ClsJeuParam) As String
    Dim i As Integer, j As Integer
    Dim Max As Double
    Dim Temp As Double
    With Parametres
    For i = 0 To .NombreTypeBatiments - 1
        If .Batiment_Jouable(i) Then
            If .Batiment_Fabrique(i) Then
                Temp = 0
                For j = 0 To .Batiment_NombreObjetsFabriques(i) - 1
                    Temp = Temp + NivObjets(.Objet_Competence(.Batiment_ObjetFabrique(i, j))) + Progression_CompetenceObjet(.Objet_Competence(.Batiment_ObjetFabrique(i, j)))
                Next j
                Temp = Temp / .Batiment_NombreObjetsFabriques(i)
                If Temp <= 0 Then Temp = 0
            ElseIf .Batiment_Service(i) > 0 Then
                Temp = NivServices(.Batiment_Service(i) - 1) + Progression_CompetenceService(.Batiment_Service(i) - 1)
                If Temp <= 0 Then Temp = 0
            ElseIf .Batiment_RessourceFabrique(i) >= 0 Then
                Temp = NivRessources(.Batiment_RessourceFabrique(i)) + Progression_CompetenceRessource(.Batiment_RessourceFabrique(i))
                If Temp <= 0 Then Temp = 0
            Else
                Temp = 0
            End If
            If Temp > Max Then
                'If Temp = Max Then
                '    Metier = Metier & "/" & .Batiment_Metier(i, Feminin)
                'Else
                Metier = .Batiment_Metier(i, Feminin)
                'End If
                Max = Temp
            End If
        End If
    Next i
    End With
End Property

'Propriété des objets.
Public Property Get Nombre_Objets_Inventaire() As Integer
    Nombre_Objets_Inventaire = UBound(ObjetsInventaires()) + 1
End Property
Public Property Get Nombre_Objets_Equipes() As Integer
    Nombre_Objets_Equipes = UBound(ObjetsEquipes()) + 1
End Property
Public Property Get Objet_Inventaire_Apparence(ByVal Indice As Integer) As Integer
    Objet_Inventaire_Apparence = ObjetsInventaires(Indice).Numero
End Property
Public Property Get Objet_Equipes_Apparence(ByVal Indice As Integer) As Integer
    Objet_Equipes_Apparence = ObjetsEquipes(Indice).Numero
End Property
Public Property Get Objet_Inventaire_Actif(ByVal Indice As Integer) As Boolean
    Objet_Inventaire_Actif = ObjetsInventaires(Indice).Actif
End Property
Public Property Let Objet_Inventaire_Actif(ByVal Indice As Integer, ByVal Valeur As Boolean)
    ObjetsInventaires(Indice).Actif = Valeur
End Property
Public Property Get Objet_Inventaire_Categorie(ByVal Indice As Integer) As Integer
    Objet_Inventaire_Categorie = ObjetsInventaires(Indice).Categorie
End Property
Public Property Let Objet_Inventaire_Categorie(ByVal Indice As Integer, ByVal Valeur As Integer)
    ObjetsInventaires(Indice).Categorie = Valeur
End Property
Public Property Get Objet_Inventaire_Nombre(ByVal Indice As Integer) As Integer
    Objet_Inventaire_Nombre = ObjetsInventaires(Indice).Nombre
End Property
Public Property Let Objet_Inventaire_Nombre(ByVal Indice As Integer, ByVal Valeur As Integer)
    ObjetsInventaires(Indice).Nombre = Valeur
End Property
Public Property Get Objet_Inventaire_Type(ByVal NumeroObjet As Integer) As Integer
    Objet_Inventaire_Type = ObjetsInventaires(NumeroObjet).Numero
End Property
Public Property Let Objet_Inventaire_Type(ByVal NumeroObjet As Integer, ByVal Valeur As Integer)
    ObjetsInventaires(NumeroObjet).Numero = Valeur
    If Valeur > -1 Then ObjetsInventaires(NumeroObjet).Actif = True
End Property
Public Property Get Objet_Equipes_Actif(ByVal Indice As Integer) As Boolean
    Objet_Equipes_Actif = ObjetsEquipes(Indice).Actif
End Property
Public Property Let Objet_Equipes_Actif(ByVal Indice As Integer, ByVal Valeur As Boolean)
    ObjetsEquipes(Indice).Actif = Valeur
End Property
Public Property Get Objet_Equipes_Categorie(ByVal Indice As Integer) As Integer
    Objet_Equipes_Categorie = ObjetsEquipes(Indice).Categorie
End Property
Public Property Let Objet_Equipes_Categorie(ByVal Indice As Integer, ByVal Valeur As Integer)
    ObjetsEquipes(Indice).Categorie = Valeur
End Property
Public Property Get Objet_Equipes_Type(ByVal NumeroObjet As Integer) As Integer
    Objet_Equipes_Type = ObjetsEquipes(NumeroObjet).Numero
End Property
Public Property Let Objet_Equipes_Type(ByVal NumeroObjet As Integer, ByVal Valeur As Integer)
    ObjetsEquipes(NumeroObjet).Numero = Valeur
    If Valeur > -1 Then ObjetsEquipes(NumeroObjet).Actif = True
End Property
Public Property Get Objet_Equipes_Competence(ByVal NumeroObjet As Integer) As Integer
    Objet_Equipes_Competence = ObjetsEquipes(NumeroObjet).Competence
End Property

'Propriétés si l'objet en question est indestructible.
Public Property Get Objet_Equipes_Indestructible(ByVal NumeroObjet As Integer) As Boolean
    Objet_Equipes_Indestructible = ObjetsEquipes(NumeroObjet).Indesctructible
End Property
Public Property Get Objet_Inventaire_Indestructible(ByVal NumeroObjet As Integer) As Boolean
    Objet_Inventaire_Indestructible = ObjetsInventaires(NumeroObjet).Indesctructible
End Property

Public Property Get Objet_Equipes_2_Mains() As Boolean
    If ObjetsEquipes(0).Categorie = 1 Then
        Objet_Equipes_2_Mains = True
    Else
        Objet_Equipes_2_Mains = False
    End If
End Property

Public Property Get Definir_Commentaires() As String
    Definir_Commentaires = Commentaires
End Property

Public Property Let Definir_Commentaires(ByVal Valeur As String)
    If (Commentaires = "" Or Valeur = "") And Parole Then 'Seul les humains font des commentaires.
        Commentaires = Valeur
        If Valeur <> "" Then
            TempsDernierCommentaire = Timer
        End If
    End If
End Property

'Propriétés du bestiaire.
Public Property Get Nombre_de_lignes_du_bestiaire() As Integer
    Nombre_de_lignes_du_bestiaire = UBound(Bestiaire()) + 1
End Property
Public Property Get Definir_Bestiaire_Total() As Long
    Dim i As Long
    Dim X As Long
    For i = 0 To Nombre_de_lignes_du_bestiaire - 1
        X = X + Definir_Bestiaire(i)
    Next i
    Definir_Bestiaire_Total = X
End Property
Public Property Get Definir_Bestiaire(ByVal Index As Integer) As Long
    Definir_Bestiaire = Bestiaire(Index)
End Property
Public Property Let Definir_Bestiaire(ByVal Index As Integer, ByVal Valeur As Long)
    Bestiaire(Index) = Valeur
End Property

'Propriétés des objets saisies.
Public Property Get Definir_ObjetSelectionne() As Boolean
    Definir_ObjetSelectionne = ArgentSelectionne > 0 Or ObjetSelectionne >= 0 Or RessourceSelectionne
End Property
Public Property Let Definir_ObjetSelectionne(ByVal Valeur As Boolean)
    Dim i As Integer
    If Not Valeur Then
        ArgentSelectionne = 0
        For i = 0 To UBound(RessourcesSelectionnes())
            RessourcesSelectionnes(i) = 0
        Next i
        RessourceSelectionne = False
        If ObjetSelectionne >= 0 Then
'            For i = 0 To Nombre_Objets_Inventaire - 1
'                If Not ObjetsInventaires(i).Actif Then
'                    ObjetsInventaires(i).Charger ObjetSelectionne, Parametres
'                    i = Nombre_Objets_Inventaire
'                End If
'            Next i
            For i = 0 To Nombre_Objets_Inventaire - 1
                If Not ObjetsInventaires(i).Actif Then
                    ObjetsInventaires(i).Charger ObjetSelectionne, Parametres
                    Exit For
                ElseIf ObjetsInventaires(i).Numero = ObjetSelectionne Then
                    'Superpose les objets de même catégorie.
                    If ObjetsInventaires(i).Nombre < Parametres.Objet_Superposable(ObjetsInventaires(i).Numero) Then
                        'S'il est superposable et inférieur au nombre maximum d'objets superposables.
                        ObjetsInventaires(i).Nombre = ObjetsInventaires(i).Nombre + 1
                        Exit For
                    End If
                End If
            Next i
            'Objet_Ajouter_Inventaire ObjetSelectionne, Parametres, Commentaires, True
            ObjetSelectionne = -1
        End If
    End If
End Property
Public Property Get Definir_RessourcesSelectionnes(ByVal Index As Integer) As Double
    Definir_RessourcesSelectionnes = RessourcesSelectionnes(Index)
End Property
Public Property Let Definir_RessourcesSelectionnes(ByVal Index As Integer, ByVal Valeur As Double)
    RessourcesSelectionnes(Index) = Valeur
End Property
Public Sub Argent_Prendre(ByVal Quantite As Double, Optional ByVal Maximum As Boolean)
    Quantite = Quantite + ArgentSelectionne
    If Maximum Then
        Quantite = CLng(Argent)
    ElseIf Argent < Quantite Then
        Quantite = Argent
    End If
    If Quantite >= 1 Then
        ArgentSelectionne = Quantite
    End If
End Sub
Public Sub Ressources_Prendre(ByVal Index As Integer, ByVal Quantite As Double, Optional ByVal Maximum As Boolean, Optional ByVal Negatif As Boolean)
    If Negatif Then 'On soustrait.
        If Maximum Then
            RessourcesSelectionnes(Index) = 0
        Else
            RessourcesSelectionnes(Index) = RessourcesSelectionnes(Index) - Quantite
            If RessourcesSelectionnes(Index) < 0 Then RessourcesSelectionnes(Index) = 0
        End If
        Ressources_Verifications
    Else
        If Maximum Then
            Quantite = Ressources(Index)
        ElseIf Ressources(Index) < Quantite Then
            Quantite = Ressources(Index)
        End If
        If Quantite >= 1 Then
            RessourcesSelectionnes(Index) = RessourcesSelectionnes(Index) + CLng(Quantite)
            RessourceSelectionne = True
        End If
    End If
End Sub
Public Sub Ressources_Prendre_Tout()
    Dim i  As Integer
    For i = 0 To Parametres.NombreRessources - 1
        Ressources_Prendre i, 0, True
    Next i
End Sub

Public Sub Ressources_Verifications()
    'Vérifie qu'il reste au joueur suffisamment de ressources.
    Dim i As Integer
    If ArgentSelectionne > Argent Then ArgentSelectionne = Argent
    RessourceSelectionne = False
    For i = 0 To UBound(RessourcesSelectionnes())
        If RessourcesSelectionnes(i) > Definir_ressources(i) Then
            RessourcesSelectionnes(i) = Definir_ressources(i)
            If Definir_ressources(i) < 1 Then
                RessourcesSelectionnes(i) = 0
            End If
        End If
        If RessourcesSelectionnes(i) > 0 Then RessourceSelectionne = True
    Next i
End Sub
Private Sub Ressources_Donner(ByRef PersoCible As ClsJeuPerso)
    Dim i As Integer
    Dim QuantiteMax As Double 'Définit la quantité maximale que l'on peut donner.
    
    Ressources_Verifications
    
    'Donne l'argent à la cible.
    PersoCible.Definir_Argent = PersoCible.Definir_Argent + ArgentSelectionne
    Definir_Argent = Definir_Argent - ArgentSelectionne
    
    For i = 0 To UBound(RessourcesSelectionnes())
        QuantiteMax = RessourcesSelectionnes(i)
        If QuantiteMax > PersoCible.Definir_MaxRessources(i) - PersoCible.Definir_ressources(i) Then
            QuantiteMax = PersoCible.Definir_MaxRessources(i) - PersoCible.Definir_ressources(i)
        End If
        'Donne les ressources à la cible.
        PersoCible.Definir_ressources(i) = PersoCible.Definir_ressources(i) + QuantiteMax
        'Retire les ressources au porteur.
        Definir_ressources(i) = Definir_ressources(i) - QuantiteMax
    Next i
    Definir_ObjetSelectionne = False
End Sub
Public Sub Ressources_Deposer_Sol(Optional ByVal IndexTresor As Long = -1)
    'Se débarasse des ressources sélectionnées.
    Dim i As Integer
    
    Ressources_Verifications
    
    If IndexTresor = -1 Then
        IndexTresor = Tresor_Creer(PositionX, PositionY)
    End If
    If ArgentSelectionne > 0 Then
        Tresors(IndexTresor).Definir_Argent = Tresors(IndexTresor).Definir_Argent + ArgentSelectionne
        Definir_Argent = Definir_Argent - ArgentSelectionne
    End If
    For i = 0 To UBound(RessourcesSelectionnes())
        If RessourcesSelectionnes(i) > 0 Then
            Tresors(IndexTresor).Definir_ressources(i) = Tresors(IndexTresor).Definir_ressources(i) + RessourcesSelectionnes(i)
            Definir_ressources(i) = Definir_ressources(i) - RessourcesSelectionnes(i)
        End If
    Next i
    Definir_ObjetSelectionne = False
    Tresors_Verifier IndexTresor
End Sub

'Propriétés de la race.
Public Property Get Definir_CoefRaceCompSpeciales(ByVal Index As Integer) As Single 'Coeficient de la race de la compétences spéciales.
    Definir_CoefRaceCompSpeciales = CoefRaceCompSpeciales(Index)
End Property
Public Property Let Definir_CoefRaceCompSpeciales(ByVal Index As Integer, ByVal Valeur As Single) 'Coeficient de la race de la compétences spéciales.
    CoefRaceCompSpeciales(Index) = Valeur
End Property
Public Property Get Definir_CoefRaceCompRessources(ByVal Index As Integer) As Single 'Coeficient de la race de la compétences des ressources.
    Definir_CoefRaceCompRessources = CoefRaceCompRessources(Index)
End Property
Public Property Let Definir_CoefRaceCompRessources(ByVal Index As Integer, ByVal Valeur As Single) 'Coeficient de la race de la compétences des ressources.
    CoefRaceCompRessources(Index) = Valeur
End Property
Public Property Get Definir_CoefRaceCompServices(ByVal Index As Integer) As Single 'Coeficient de la race de la compétences des Services.
    Definir_CoefRaceCompServices = CoefRaceCompServices(Index)
End Property
Public Property Let Definir_CoefRaceCompServices(ByVal Index As Integer, ByVal Valeur As Single) 'Coeficient de la race de la compétences des Services.
    CoefRaceCompServices(Index) = Valeur
End Property
Public Property Get Definir_CoefRaceCompObjets(ByVal Index As Integer) As Single 'Coeficient de la race de la compétence des objets.
    Definir_CoefRaceCompObjets = CoefRaceCompObjets(Index)
End Property
Public Property Let Definir_CoefRaceCompObjets(ByVal Index As Integer, ByVal Valeur As Single) 'Coeficient de la race de la compétence des objets.
    CoefRaceCompObjets(Index) = Valeur
End Property

'Public Sub Changer_Race(ByVal Index As Integer, _
'                        ByVal Parametres As ClsJeuParam, _
'                        Optional ByVal Grade As Integer = -1, _
'                        Optional ByVal NumeroMonstre As Integer = -1, _
'                        Optional ByVal SansEquipement As Boolean)
'    Dim i, j As Double
'
'    With Parametres
'    If Index <> 0 Then 'On ne gère pas les humains.
'        Changer_Race_Personnage 0, Parametres
'        If NumeroMonstre = -1 Then 'Si le monstre n'est pas entré en paramètre, on le tire au sort.
'            i = RND
'            If Grade < 0 Then
'                Select Case i
'                'Case Is < 0.5: Grade = 0
'                Case Is < 0.75: Grade = 0
'                Case Is < 0.875: Grade = 1
'                Case Is < 0.9375: Grade = 2
'                Case Is < 0.96875: Grade = 3
'                Case Is < 0.984375: Grade = 4
'                Case Else: Grade = 5
'                End Select
'            End If
'            'Debug.Print Grade & " : " & i
'            If Grade > .Peuples_NombreMonstres(Index) - 1 Then
'                'Grade = .Peuples_NombreMonstres(Index) - 1
'                Grade = 0
'            End If
'            For i = 0 To .NombreTotalMonstres - 1
'                If Index = .Monstre_Peuple(i) And _
'                   Grade = .Monstre_Grade(i) Then
'                    NumeroMonstre = i
'                End If
'            Next i
'        End If
'        'On charge les paramètres du monstre.
'        Nom = .Monstre_Nom(NumeroMonstre)
'        Race = .Monstre_Peuple(NumeroMonstre)
'        Largeur = .Monstre_Largeur(NumeroMonstre)
'        Hauteur = .Monstre_Hauteur(NumeroMonstre)
'        IndiceMonstre = NumeroMonstre + 1
'        IA_Equiper = .Monstre_IA_Equiper(NumeroMonstre)
'        IA_Manger = .Monstre_IA_Manger(NumeroMonstre)
'        IA_Pacifique = .Monstre_IA_Pacifique(NumeroMonstre)
'        IA_Temeraire = .Monstre_IA_Temeraire(NumeroMonstre)
'        IA_Berserk = .Monstre_IA_Berserk(NumeroMonstre)
'        If .Monstre_IA_Fichier(NumeroMonstre) <> "" Then
'            ChoisirIA = True
'            FichierIA = .Monstre_IA_Fichier(NumeroMonstre)
'        Else
'            ChoisirIA = False
'            FichierIA = ""
'        End If
'        Parole = .Monstre_Parole(NumeroMonstre)
'        Volant = .Monstre_Volant(NumeroMonstre)
'        Titre = .Monstre_Grade(NumeroMonstre)
'        MaxAttaque = .Monstre_Attaque(NumeroMonstre)
'        Attaque = MaxAttaque
'        PorteeAttaque = .Monstre_PorteeAttaque(NumeroMonstre)
'        EffetAttaque = .Monstre_EffetAttaque(NumeroMonstre)
'        Kamikaze = .Monstre_Kamikaze(NumeroMonstre)
'        MaxDefense = .Monstre_Defense(NumeroMonstre)
'        Defense = MaxDefense
'        Armure = .Monstre_Armure(NumeroMonstre)
'        LongueurPas = .Monstre_Vitesse(NumeroMonstre)
'        MaxVie = .Monstre_Vie(NumeroMonstre)
'        Vie = MaxVie
'        MaxEnergie = .Monstre_Energie(NumeroMonstre)
'        Energie = MaxEnergie
'        MaxMagie = .Monstre_Magie(NumeroMonstre)
'        Magie = MaxMagie
'        MaxMoral = .Monstre_Moral(NumeroMonstre)
'        Moral = MaxMoral
'        Apparence = .Monstre_Apparence(NumeroMonstre)
'        Argent = .Monstre_Argent(NumeroMonstre)
'        Ajuster_Bonus
'        If Not SansEquipement Then
'            'On lui attribue ses ressources.
'            For i = 0 To .NombreRessources - 1
'                Ressources(i) = .Monstre_Ressource(NumeroMonstre, i)
'            Next i
'            'On lui attribue son équipement.
'            Objet_Equipes_Vider_Tout
'            Objet_Inventaire_Vider_Tout
'            For i = 0 To .NombreObjets - 1
'                If .Monstre_Objet(NumeroMonstre, i) > 0 Then
'                    For j = 0 To .Monstre_Objet(NumeroMonstre, i) - 1
'                        Objet_Ajouter_Inventaire i, Parametres, ModMain.Commentaires, IA_Equiper
'                        'Debug.Print Numero & ", " & Parametres.Monstre_Nom(NumeroMonstre) & ", " & Parametres.Objet_Nom(i)
'                    Next j
'                End If
'            Next i
'            Actualiser_Bonus
'        End If
'
'        GainExperience = False
'
'        Calculer_Apparence Maisons(Numero), Parametres
'    End If
'    End With
'End Sub

Public Sub Changer_Race_Personnage(ByVal Index As Integer, _
                                   ByVal Parametres As ClsJeuParam, _
                                   Optional ByVal SansEquipement As Boolean)
    Dim i As Long, j As Long
    'Variables mémorisant le pourcentage de caractéritiques.
    Dim CoefVie As Single
    Dim CoefEnergie As Single
    Dim CoefMagie As Single
    Dim CoefMoral As Single
    Dim CoefAttaque As Single
    Dim CoefDefense As Single
    
    If Definir_MaxVie > 0 Then CoefVie = Vie / Definir_MaxVie
    If Definir_MaxEnergie > 0 Then CoefEnergie = Energie / Definir_MaxEnergie
    If Definir_MaxMagie > 0 Then CoefMagie = Magie / Definir_MaxMagie
    If Definir_MaxMoral > 0 Then CoefMoral = Moral / Definir_MaxMoral
    If Definir_MaxAttaque > 0 Then CoefAttaque = Attaque / Definir_MaxAttaque
    If Definir_MaxDefense > 0 Then CoefDefense = Defense / Definir_MaxDefense
    
    With Parametres
    
    Race = Index
    
    'Tire le sexe.
    If Rnd > 0.5 And _
       Not .Race_RestrictionSexe(Index, True) Then
        Feminin = True
    Else
        Feminin = False
    End If
    If Joueur <= 0 Then 'Les joueurs ont déjà leurs noms.
        If .Race_NomAleatoire(Index) Then
            'Tire un nom au hasard.
            Nom = Noms.Tirer_Nom_Aleatoire(Feminin)
        Else
            'Lui attribut le nom de sa race.
            Nom = .Race_Nom(Index, Feminin)
        End If
    End If
    Largeur = .Race_Largeur(Index)
    Hauteur = .Race_Hauteur(Index)
    Nombre_ApparenceBase = .Race_DefinirNombreCategoriesApparences(Index)
    ChoisirApparence = .Race_DefinirApparence(Index)
    ApparenceVerouillee = .Race_ApparenceVerouillee(Index)
    
    MaxVie = .Race_Vie(Index)
    MaxEnergie = .Race_Energie(Index)
    MaxMagie = .Race_Magie(Index)
    MaxMoral = .Race_Moral(Index)
    MaxAttaque = .Race_Attaque(Index)
    MaxDefense = .Race_Defense(Index)
    
    CoefRaceVie = .Race_CoefVie(Index)
    CoefRaceEnergie = .Race_CoefEnergie(Index)
    CoefRaceMagie = .Race_CoefMagie(Index)
    CoefRaceMoral = .Race_CoefMoral(Index)
    CoefRaceAttaque = .Race_CoefAttaque(Index)
    CoefRaceDefense = .Race_CoefDefense(Index)
    VitesseRegenerationVie = .Race_VitesseRegenerationVie(Index)
    VitesseRegenerationEnergie = .Race_VitesseRegenerationEnergie(Index)
    VitesseRegenerationMagie = .Race_VitesseRegenerationMagie(Index)
    VitesseRegenerationMoral = .Race_VitesseRegenerationMoral(Index)
    RestaurationComplete = .Race_RestaurationComplete(Index)
    ReDim CoefRaceCompSpeciales(.NombreCompetencesSpeciales - 1)
    For i = 0 To .NombreCompetencesSpeciales - 1
        CoefRaceCompSpeciales(i) = .Race_CoefCompSpeciales(Index, i)
    Next i
    ReDim CoefRaceCompRessources(.NombreRessources - 1)
    For i = 0 To .NombreRessources - 1
        CoefRaceCompRessources(i) = .Race_CoefCompRessources(Index, i)
    Next i
    ReDim CoefRaceCompServices(.NombreServices - 1)
    For i = 0 To .NombreServices - 1
        CoefRaceCompServices(i) = .Race_CoefCompServices(Index, i)
    Next i
    ReDim CoefRaceCompObjets(.NombreCompetencesObjets - 1)
    For i = 0 To .NombreCompetencesObjets - 1
        CoefRaceCompObjets(i) = .Race_CoefCompObjets(Index, i)
    Next i
    'Les résistances.
    For i = 0 To .NombreResistances - 1
        Resistances(i) = .Race_Resistances(Index, i)
    Next i
    GainExperience = .Race_Experience(Index)
    Volant = .Race_Volant(Index)
    Parole = .Race_Parole(Index)
    CommentairesSpecifique = .Race_CommentairesSpecifique(Index) > 0
    CommentairesSpecifiqueIndice = .Race_CommentairesSpecifique(Index)
    BonusMaxRessourcesRace = .Race_BonusMaxRessources(Index)
    PorteeAttaque = .Race_PorteeAttaque(Index)
    EffetAttaque = .Race_EffetAttaque(Index)
    AttaqueCauseEffet = .Race_AttaqueCauseEffet(Index)
    AttaqueResistance = .Race_AttaqueResistance(Index)
    Armure = .Race_Armure(Index)
    LongueurPas = .Race_Vitesse(Index)
    Immobile = LongueurPas = 0
    Acceleration = .Race_Acceleration(Index)
    Kamikaze = .Race_Kamikaze(Index)
    Incommandable = .Race_IA_Incommandable(Index)
    If Not SansEquipement Then
        'On lui attribue ses ressources.
        For i = 0 To .NombreRessources - 1
            If Ressources(i) < .Race_Ressources(Index, i) Then
                Ressources(i) = .Race_Ressources(Index, i)
            End If
        Next i
        'On lui attribue son équipement.
        If .Race_NombreEquipement(Index) > 0 Then
            'Objet_Equipes_Vider_Tout
            'Objet_Inventaire_Vider_Tout
            For i = 0 To .Race_NombreEquipement(Index) - 1
                Objet_Ajouter_Inventaire .Race_Equipement(Index, i), Parametres, ModMain.Commentaires, IA_Equiper > 0
            Next i
            Actualiser_Bonus
        End If
    End If
    'Retire l'équipement qu'il n'est pas censé pouvoir porter.
    For i = 0 To Nombre_Objets_Equipes - 1
        If Objet_Equipes_Actif(i) Then
            If .Objet_Monture(Objet_Equipes_Type(i)) And Not .Race_UtilisationMonture(Race) Then
                Objet_Ranger i, Parametres, ModMain.Commentaires
            End If
        End If
    Next i
    
    Ajuster_Bonus
    
    'Projectile = .Race_Projectile(Index)
    Pietinement = .Race_Pietinement(Index)
    Inattaquable = .Race_Inattaquable(Index)
    
    IA_Equiper = .Race_Equiper(Index)
    IA_Manger = .Race_IA_Manger(Index)
    IA_Pacifique = .Race_IA_Pacifique(Index)
    IA_Temeraire = .Race_IA_Temeraire(Index)
    IA_Protecteur = .Race_IA_Protecteur(Index)
    IA_Egoiste = .Race_IA_Egoiste(Index)
    IA_Berserk = .Race_IA_Berserk(Index)
    If .Race_IA_Fichier(Index) <> "" Then
        ChoisirIA = True
        FichierIA = .Race_IA_Fichier(Index)
    Else
        ChoisirIA = False
        FichierIA = ""
    End If
        
    End With
    
    If CoefVie > 0 Then Vie = Definir_MaxVie * CoefVie
    If CoefEnergie > 0 Then Energie = Definir_MaxEnergie * CoefEnergie
    If CoefMagie > 0 Then Magie = Definir_MaxMagie * CoefMagie
    If CoefMoral > 0 Then Moral = Definir_MaxMoral * CoefMoral
    If CoefAttaque > 0 Then Attaque = Definir_MaxAttaque * CoefAttaque
    If CoefDefense > 0 Then Defense = Definir_MaxDefense * CoefDefense
End Sub

Public Sub Changer_Race_Personnage_Equipement(ByVal Parametres As ClsJeuParam, _
                                              ByRef Commentaires As ClsJeuCommentaires)
    Dim i As Integer, j As Integer
    With Parametres
    'On lui attribue son équipement.
    If .Race_NombreObjets(Race) > 0 Then
        For i = 0 To .Race_NombreObjets(Race) - 1
            j = .Race_Objet(Race, i)
            If Parametres.Objet_PrixTalents(j) <= 0 Then
                If .Objet_ChangerApparencePerso(j) Then
                    ChoisirApparence = False
                    Apparence = 0
                End If
                Objet_Ajouter_Inventaire j, Parametres, ModMain.Commentaires, IA_Equiper > 0
            End If
        Next i
    End If
    End With
    Objets_Utilisation_Automatique Parametres, Commentaires
End Sub

Public Sub Objets_Utilisation_Automatique(ByVal Parametres As ClsJeuParam, _
                                           ByRef Commentaires As ClsJeuCommentaires)
    Dim i As Integer
    'Utilise tous les objets automatiques.
    For i = 0 To Nombre_Objets_Inventaire - 1
        If Objet_Inventaire_Actif(i) Then
            If Parametres.Objet_UtilisationAutomatique(Objet_Inventaire_Type(i)) Then
                Objet_Equiper i, Parametres, Commentaires
                Gestion_Utilisation_Speciale
            End If
        End If
    Next i
End Sub

Public Sub Changer_Attitude()
    Attitude = Attitude + 1
    If Attitude >= NombreAttitudes Then
        Attitude = 0
    End If
End Sub

Public Sub Changer_Formation()
    Formation = Formation + 1
    If Formation >= NombreFormations Then
        Formation = 0
    End If
End Sub

Public Property Get Definir_Volant() As Boolean
    Definir_Volant = BonusVolant
End Property
Public Property Let Definir_Volant(ByVal Valeur As Boolean)
    Volant = Valeur
End Property

'Propriétés pour les apparences.
Public Property Let Nombre_ApparenceBase(ByVal Valeur As Integer)
    NombreApparenceBase = Valeur
    If Valeur = 0 Then Valeur = 1
    ReDim Preserve ApparenceBase(Valeur - 1)
    ReDim Preserve ApparenceBaseAfficher(Valeur - 1)
    ReDim Preserve FichierApparenceBase(Valeur - 1)
End Property
Public Property Get Nombre_ApparenceBase() As Integer
    Nombre_ApparenceBase = NombreApparenceBase
End Property
Public Property Let Definir_ApparenceBase(ByVal Index As Integer, _
                                          ByVal Valeur As Integer)
    ApparenceBase(Index) = Valeur
End Property
Public Property Get Definir_ApparenceBaseAfficher(ByVal Index As Integer) As Boolean
    Definir_ApparenceBaseAfficher = ApparenceBaseAfficher(Index)
End Property
Public Property Let Definir_ApparenceBaseAfficher(ByVal Index As Integer, _
                                                  ByVal Valeur As Boolean)
    ApparenceBaseAfficher(Index) = Valeur
End Property
Public Property Get Definir_ApparenceBase(ByVal Index As Integer) As Integer
    Definir_ApparenceBase = ApparenceBase(Index)
End Property
Public Property Let Nombre_ApparenceObjet(ByVal Valeur As Integer)
    While ApparenceObjet.Count > Valeur
        ApparenceObjet.Remove ApparenceObjet.Count
    Wend
End Property
Public Property Get Nombre_ApparenceObjet() As Integer
    Nombre_ApparenceObjet = ApparenceObjet.Count
End Property
Public Property Get Definir_ApparenceObjet(ByVal Index As Integer) As Integer
    Definir_ApparenceObjet = ApparenceObjet(Index)
End Property
Public Property Let Definir_ApparenceObjet(ByVal Index As Integer, _
                                           ByVal Valeur As Integer)
    ApparenceObjet.Add Valeur
End Property
Public Property Let Definir_FichierApparenceBase(ByVal Index As Integer, _
                                                 ByVal Valeur As String)
    FichierApparenceBase(Index) = Valeur
End Property
Public Property Get Definir_FichierApparenceBase(ByVal Index As Integer) As String
    Definir_FichierApparenceBase = FichierApparenceBase(Index)
End Property

Public Property Get Definir_ApparenceLargeur() As Integer
    Definir_ApparenceLargeur = ApparenceLargeur
End Property
Public Property Get Definir_ApparenceHauteur() As Integer
    Definir_ApparenceHauteur = ApparenceHauteur
End Property
Public Property Get Definir_ApparenceDecalageX(ByVal Direction As Integer) As Integer
    Definir_ApparenceDecalageX = ApparenceDecalageX(Direction)
End Property
Public Property Let Definir_ApparenceDecalageX(ByVal Direction As Integer, ByVal Valeur As Integer)
    ApparenceDecalageX(Direction) = Valeur
End Property
Public Property Get Definir_ApparenceDecalageY(ByVal Direction As Integer) As Integer
    Definir_ApparenceDecalageY = ApparenceDecalageY(Direction) - IIf(NumeroPas = 1, 1, 0)  'Ajoute un effet de mouvement.
End Property
Public Property Get Definir_ApparenceDecalageYFixe(ByVal Direction As Integer) As Integer
    Definir_ApparenceDecalageYFixe = ApparenceDecalageY(Direction) 'Sans effet de mouvement.
End Property
Public Property Let Definir_ApparenceDecalageYFixe(ByVal Direction As Integer, ByVal Valeur As Integer)
    ApparenceDecalageY(Direction) = Valeur
End Property
Public Property Get Definir_ApparenceAvantPlan(ByVal Direction As Integer) As Boolean
    Definir_ApparenceAvantPlan = ApparenceAvantPlan(Direction)
End Property
Public Property Let Definir_ApparenceAvantPlan(ByVal Direction As Integer, ByVal Valeur As Boolean)
    ApparenceAvantPlan(Direction) = Valeur
End Property

'Propriétés des résistances.
Public Property Get Definir_Resistance(ByVal Index As Integer) As Single
    If Index >= 0 Then
        Definir_Resistance = BonusResistances(Index)
    Else
        Definir_Resistance = 1 'Valeur par défaut
    End If
End Property
Public Property Get Definir_ResistanceDepart(ByVal Index As Integer) As Single
    Definir_ResistanceDepart = Resistances(Index)
End Property
Public Property Get Bonus_Resistance(ByVal Index As Integer) As Boolean
    Bonus_Resistance = BonusResistances(Index) < Resistances(Index)
End Property

'pripriétés des générations.
Public Property Get Definir_GenerationArgent() As Single
    Definir_GenerationArgent = GainArgent
End Property
Public Property Get Definir_GenerationRessource(ByVal Index As Integer) As Single
    Definir_GenerationRessource = GainRessource(Index)
End Property

'Propriétés des effets temporaires.
Public Property Get Nombre_EffetsTemp() As Integer
    Nombre_EffetsTemp = EffetsTemp.Count
End Property
Public Property Let Nombre_EffetsTemp(ByVal Valeur As Integer)
    Dim i As Integer
    If Valeur <> Nombre_EffetsTemp Then
        EffetTemp_Effacer_Tout
        For i = 1 To Valeur
            EffetTemp_Ajouter 0, 0
        Next i
    End If
End Property
Public Property Get Definir_EffetTemp_Numero(ByVal Index As Integer) As Integer
    Definir_EffetTemp_Numero = EffetsTemp(Index).Numero
End Property
Public Property Let Definir_EffetTemp_Numero(ByVal Index As Integer, ByVal Valeur As Integer)
    EffetsTemp(Index).Numero = Valeur
End Property
Public Property Get Definir_EffetTemp_Duree(ByVal Index As Integer) As Long
    Definir_EffetTemp_Duree = EffetsTemp(Index).Duree
End Property
Public Property Let Definir_EffetTemp_Duree(ByVal Index As Integer, _
                                            ByVal Valeur As Long)
    If Valeur > 1000000 Then Valeur = 1000000  'La limite maximale est d'11 jours et demi.
    EffetsTemp(Index).Duree = Valeur
    If Valeur > EffetsTemp(Index).DureeMax Then EffetsTemp(Index).DureeMax = Valeur
End Property
Public Property Get Definir_EffetTemp_DureeMax(ByVal Index As Integer) As Long
    Definir_EffetTemp_DureeMax = EffetsTemp(Index).DureeMax
End Property
Public Property Let Definir_EffetTemp_DureeMax(ByVal Index As Integer, _
                                               ByVal Valeur As Long)
    EffetsTemp(Index).DureeMax = Valeur
End Property
Public Property Get Definir_EffetTemp_Proprietaire(ByVal Index As Integer) As Long
    Definir_EffetTemp_Proprietaire = EffetsTemp(Index).Proprietaire
End Property
Public Property Let Definir_EffetTemp_Proprietaire(ByVal Index As Integer, _
                                                   ByVal Valeur As Long)
    EffetsTemp(Index).Proprietaire = Valeur
End Property
Public Sub EffetTemp_Ajouter(ByVal No As Integer, _
                             ByVal Duree As Long, _
                             Optional ByVal DureeMax As Long = -1, _
                             Optional Proprietaire As Long = -1)
    Dim i As Integer
    Dim Ajout As Boolean
    Dim EffetTemp As ClsJeuEffetTemp
    Ajout = True
    If Proprietaire = -1 Then Proprietaire = Numero
    For i = 1 To Nombre_EffetsTemp
        If Definir_EffetTemp_Numero(i) = No Then
            If Definir_EffetTemp_Proprietaire(i) = Proprietaire Then
                Definir_EffetTemp_Duree(i) = Definir_EffetTemp_Duree(i) + Duree
                If DureeMax > -1 And Definir_EffetTemp_Duree(i) > DureeMax Then
                    Definir_EffetTemp_Duree(i) = DureeMax
                    Definir_EffetTemp_DureeMax(i) = DureeMax
                End If
                Ajout = False
                Exit For
            End If
        End If
    Next i
    If Ajout Then
        Set EffetTemp = New ClsJeuEffetTemp
        EffetTemp.Numero = No
        EffetTemp.Duree = Duree
        EffetTemp.DureeMax = Duree
        EffetTemp.Proprietaire = Proprietaire
        EffetsTemp.Add EffetTemp
        Actualiser_Bonus
    End If
    If ComReseau.Serveur Then ComMessages.Perso_EffetTempAjouter Numero, No, Duree, DureeMax, Proprietaire, ComReseau.PlayerID, True
End Sub
Public Sub EffetTemp_Effacer(ByVal Index As Integer)
    'Restaure l'objet si possible.
    If Parametres.Objet_EffetAnnuleRestaureObjet(EffetsTemp(Index).Numero) Then
        Objet_Ajouter_Inventaire_Fast EffetsTemp(Index).Numero
    End If
    
    EffetsTemp.Remove Index
    Actualiser_Bonus
    If ComReseau.Serveur Then ComMessages.Perso_EffetTempSupprimer Numero, Index, ComReseau.PlayerID, True
End Sub
Public Sub EffetTemp_Effacer_Tout()
    If Nombre_EffetsTemp > 0 Then
        Do
            EffetTemp_Effacer 1
        Loop Until Nombre_EffetsTemp = 0
        'Actualiser_Bonus
    End If
End Sub
Public Sub EffetTemp_Effacer_Mort()
    'Efface tous les effets temporaires qui se dissipent à la mort.
    Dim i As Long
    Dim j As Long
    Dim Temp As Boolean
    If Nombre_EffetsTemp > 0 Then
        j = 1
        Do
            Temp = True
            For i = j To EffetsTemp.Count
                If Parametres.Objet_EffetAnnulationMort(EffetsTemp(i).Numero) Then
                    EffetTemp_Effacer i
                    j = i
                    Temp = False
                    Exit For
                End If
            Next i
        Loop Until Temp
    End If
End Sub

'Propriétés des rechargements des objets.
Public Property Get Nombre_EffetsRecharge() As Integer
    Nombre_EffetsRecharge = EffetsRecharge.Count
End Property
Public Property Let Nombre_EffetsRecharge(ByVal Valeur As Integer)
    Dim i As Integer
    If Valeur <> Nombre_EffetsRecharge Then
        EffetRecharge_Effacer_Tout
        For i = 1 To Valeur
            EffetRecharge_Ajouter 0, 0
        Next i
    End If
End Property
Public Property Get Definir_EffetsRecharge_Numero(ByVal Index As Integer) As Integer
    Definir_EffetsRecharge_Numero = EffetsRecharge(Index).Numero
End Property
Public Property Let Definir_EffetsRecharge_Numero(ByVal Index As Integer, ByVal Valeur As Integer)
    EffetsRecharge(Index).Numero = Valeur
End Property
Public Property Get Definir_EffetsRecharge_Duree(ByVal Index As Integer) As Long
    Definir_EffetsRecharge_Duree = EffetsRecharge(Index).Duree
End Property
Public Property Let Definir_EffetsRecharge_Duree(ByVal Index As Integer, _
                                                ByVal Valeur As Long)
    EffetsRecharge(Index).Duree = Valeur
End Property
Public Sub EffetRecharge_Ajouter(ByVal No As Integer, _
                                 ByVal Duree As Long)
    Dim EffetTemp As ClsJeuEffetTemp
    Dim i As Integer
    Dim Ajout As Boolean
    'Dim EffetTemp As ClsJeuEffetTemp
    Ajout = True
    For i = 1 To EffetsRecharge.Count
        If EffetsRecharge(i).Numero = No Then
            EffetsRecharge(i).Duree = EffetsRecharge(i).Duree + Duree
            Ajout = False
            Exit For
        End If
    Next i
    If Ajout Then
        Set EffetTemp = New ClsJeuEffetTemp
        EffetTemp.Numero = No
        EffetTemp.Duree = Duree
        EffetsRecharge.Add EffetTemp
    End If
    If ComReseau.Serveur Then ComMessages.Perso_EffetRechargeAjouter Numero, No, Duree, ComReseau.PlayerID, True
End Sub
Public Sub EffetRecharge_Effacer(ByVal Index As Integer)
    EffetsRecharge.Remove Index
End Sub
Public Sub EffetRecharge_Effacer_Tout()
    Collection_Vider EffetsRecharge
End Sub
Public Function EffetRecharge_Pret(ByVal IndiceObjet As Integer, ByVal Parametres As ClsJeuParam) As Boolean
    'Renvoie vrai si l'objet est rechargé.
    Dim i As Integer
    With Parametres
    EffetRecharge_Pret = True
    If .Objet_DureeRecharge(IndiceObjet) > 0 Then
        For i = 1 To EffetsRecharge.Count
            If EffetsRecharge(i).Numero = IndiceObjet Then
                If .Objet_RechargesMultiples(IndiceObjet) = 0 Then
                    EffetRecharge_Pret = False
                'Recharges multiples.
                'ElseIf EffetsRecharge(i).Duree - Objet_Equipe_BonusSup(Definir_Niveau_CompetenceObjets(.Objet_Competence(IndiceObjet)), .Objet_RechargesMultiples(IndiceObjet)) * .Objet_DureeRecharge(IndiceObjet) > .Objet_DureeRecharge(IndiceObjet) Then
                ElseIf Int(Objet_Equipe_BonusSup(Definir_Niveau_CompetenceObjets(.Objet_Competence(IndiceObjet)), .Objet_RechargesMultiples(IndiceObjet)) - EffetRecharge_TempRestant(IndiceObjet) / .Objet_DureeRecharge(IndiceObjet)) <= 0 Then
                    EffetRecharge_Pret = False
                End If
                Exit For
            End If
        Next i
    End If
    End With
End Function
Public Function EffetRecharge_TempRestant(ByVal IndiceObjet As Integer) As Long
    'Renvoie la duree nécessaire à la recharge de l'objet.
    Dim i As Integer
    For i = 1 To EffetsRecharge.Count
        If EffetsRecharge(i).Numero = IndiceObjet Then
            EffetRecharge_TempRestant = EffetsRecharge(i).Duree
            Exit For
        End If
    Next i
End Function
