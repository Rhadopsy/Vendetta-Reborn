VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsIntFenetre"
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
DefLng A-Z

Const NomFichierINI = "Fenetres"
Const NomSectionINI = "Fenetres"

Const Main1 = 0
Const Main2 = 1
Const Armure = 2
Const Tete = 3
Const Pieds = 4
Const Dos = 5
Const Anneau1 = 6
Const Anneau2 = 7
Const Bras = 8
Const Orbe = 9
Const Monture = 10

Const TitrePositionX = 128
Const TitrePositionY = 32

Const BoutonDeplacerPositionX = 256 + 40 - 16 - 24
Const BoutonDeplacerPositionY = 0 + 16
Private BoutonDeplacerEtiquette As String

Const BoutonFermerPositionX = 256 + 40 - 16
Const BoutonFermerPositionY = 0 + 16
Private BoutonFermerEtiquette As String

Const ArgentPersoPositionX = 32
Const ArgentPersoPositionY = 54
Const ArgentMaisonPositionX = 32
Const ArgentMaisonPositionY = 96
Const ArgentBoutonsPrendrePositionX = 32
Const ArgentBoutonsPrendrePositionY = 54
Const ArgentBoutonsPrendreLargeur = 100
Const ArgentBoutonsPrendreHauteur = 24
Private ArgentBoutonsPrendreEtiquette As String

Const EtatBatimentPositionX = 32
Const EtatBatimentPositionY = 54
Const StockBatimentPositionX = 32
Const StockBatimentPositionY = 96 + 42 '- 24

'Boutons de restauration du chateau.
Const NombreBoutonsRestauration = 3
Private BoutonsRestaurationEtiquette1 As String
Private BoutonsRestaurationEtiquette2 As String
Private BoutonsRestaurationEtiquette3 As String
Const BoutonsRestaurationPositionX = 184 + 32
Const BoutonsRestaurationPositionY = 48 '54
Const BoutonsRestaurationIntervalX = 24
Const BoutonsRestaurationIntervalY = 0

'Propriété des informations du bouton "Empereur".
Const BoutonEmpereurPositionX = 160 - 12 - 8
Const BoutonEmpereurPositionY = 24 - 16 + 8
Const BoutonEmpereurLargeur = 24
Const BoutonEmpereurHauteur = 24
Private BoutonEmpereurInfoBullePrefixe As String
Private BoutonEmpereurInfoBulleSuffixe As String

'Propriété des informations des services.
Const ServiceDescriptionPositionX = 32
Const ServiceDescriptionPositionY = 96 + 64
Const ServiceBoutonPositionX = 184 + 32
Const ServiceBoutonPositionY = 96 + 16 + 24
Private ServiceBoutonInfoBullePrefixe As String
Private ServiceBoutonInfoBulleSuffixe As String

'Lignes de fabrication des objets de la maison.
Const NombreLignesObjets = 8 '7
Const FabricationPositionX = 32
Const FabricationPositionY = 96 + 96 - 32 '- 24 '+ 16
Const FabricationIconeZoom = 0.75
Const FabricationIconeDecalageX = 72 + 64
Const FabricationIconeDecalageY = -6 '-8
Const FabricationQuantiteDecalageX = 72 + 32 + 64
Const FabricationQuantiteDecalageY = -0
Const FabricationBoutonDecalageX = 72 + 32 * 2 + 64
Const FabricationBoutonDecalageY = -4
Const FabricationFileAttenteDecalageX = 72 + 32 * 2 + 64 + 8
Const FabricationFileAttenteDecalageY = 0 + 8
Const FabricationProgressionDecalageX = 72 + 32 * 2 + 96
Const FabricationProgressionDecalageY = 0
Const FabricationIntervalX = 0
Const FabricationIntervalY = 24 '32

'Liste des ressources.
Const BoutonsRessourcesTotalPositionX = 216
Const BoutonsRessourcesTotalPositionY = 48 '54
Const BoutonsRessourcesTotalIntervalX = 24
Const BoutonsRessourcesTotalIntervalY = 0
Private BoutonsRessourcesTotalEtiquette1 As String
Private BoutonsRessourcesTotalEtiquette2 As String
Private BoutonsRessourcesTotalEtiquette3 As String
Private BoutonsRessourcesTotalEtiquette4 As String
Private BoutonsRessourcesTotalEtiquette5 As String
Private BoutonsRessourcesTotalEtiquette6 As String

Const NombreLignesRessources = 11
Const RessourcesPositionX = 32
Const RessourcesPositionY = 112 - 24
Const RessourcesImageDecalageX = 72
Const RessourcesImageDecalageY = -10
Const RessourcesScoreDecalageX = 104
Const RessourcesScoreDecalageY = 0
Const RessourcesIntervalX = 0
Const RessourcesIntervalY = 24
Const RessourcesBoutonsPrendreDecalageX = 72 '104
Const RessourcesBoutonsPrendreDecalageY = 0
Const RessourcesBoutonsPrendreLargeur = 24 ' 48
Const RessourcesBoutonsPrendreHauteur = 24
Private RessourcesBoutonsPrendreEtiquette As String

Const SeparateurRessourceDecalageX = 0 'Décalage par rapport à l'étiquette des ressources.
Const SeparateurRessourceDecalageY = 18
Const SeparateurRessourceLongueur = 216 '256

Const BoutonMangerDecalageX = 152 'Décalage par rapport à l'étiquette des ressources.
Const BoutonMangerDecalageY = -6
Const BoutonMangerFormat = "0.0"

Const BoutonCommerceDecalageX = 184 'Décalage par rapport à l'étiquette des ressources.
Const BoutonCommerceDecalageY = -6
Const BoutonCommerceIntervalX = 24
Const BoutonCommerceIntervalY = 0
Const NombreBoutonsCommerce = 3

Const BoutonTransfertDecalageX = BoutonCommerceDecalageX
Const BoutonTransfertDecalageY = BoutonCommerceDecalageY
Const BoutonTransfertIntervalX = BoutonCommerceIntervalX
Const BoutonTransfertIntervalY = BoutonCommerceIntervalY
Const NombreBoutonsTransfert = NombreBoutonsCommerce
Private BoutonFabriquerEtiquette As String
Private BoutonTransferArgentPerso1Etiquette As String
Private BoutonTransferArgentPerso2Etiquette As String
Private BoutonTransferArgentPerso3Etiquette As String
Private BoutonTransferArgentMaison1Etiquette As String
Private BoutonTransferArgentMaison2Etiquette As String
Private BoutonTransferArgentMaison3Etiquette As String
Private BoutonTransferStockPerso1Etiquette As String
Private BoutonTransferStockPerso2Etiquette As String
Private BoutonTransferStockPerso3Etiquette As String
Private BoutonTransferStockMaison1Etiquette As String
Private BoutonTransferStockMaison2Etiquette As String
Private BoutonTransferStockMaison3Etiquette As String
Private BoutonAchatStockMaison1Etiquette As String
Private BoutonAchatStockMaison2Etiquette As String
Private BoutonAchatStockMaison3Etiquette As String

Private EtiquetteObjets  As String
Private EtiquetteStockVide  As String
Private EtiquetteIndisponible  As String

'Fenêtre d'inventaire.
Const InventaireEquipementPositionX = 32
Const InventaireEquipementPositionY = 32 '192
'Const InventaireEquipementEtiquette = "Equipement" '"Equipement de réserve :"
Const InventairePositionX = 32
Const InventairePositionY = 241
Const InventaireDecalageX = 32
Const InventaireDecalageY = 32
Const InventaireNombreColonnes = 8
Const EquipementContourCouleur = &H80FF&
Const EquipementContourEpaisseur = 2
Private InventaireEquipementObjetX() As Integer
Private InventaireEquipementObjetY() As Integer
Const InventaireMaisonX = 288
Const InventaireMaisonY = 208
Private InventaireMaisonMessagePerso As String
Private InventaireMaisonMessageMaison As String
Private InventaireMaisonMessagePersoObjet As String
Private InventaireMaisonMessageMaisonObjet As String
Const InventaireEquipementNombreCases = 32
Private NombreBoutonsInventaireOnglets As Integer
Const InventaireOngletsPositionX = 8
Const InventaireOngletsPositionY = 212
Const InventaireOngletsDecalageX = 0
Const InventaireOngletsDecalageY = 25
Public InventaireOngletsChoix As Integer

'Boutons de défilement.
Const NombreBoutonsDefilement = 3
Const AfficherBoutonsDefilement = True
Const BoutonDefilement1PositionX = 320 - 16 - 4 '(320 - 64) / 2
Const BoutonDefilement1PositionY = 66 '64 + 32 - 8 - 22
Private BoutonDefilement1Etiquette As String
Const BoutonDefilement2PositionX = 320 - 16 - 4 '(320 - 64) / 2
Const BoutonDefilement2PositionY = 368 - 44 - 16 '368 - 32 - 16 + 8
Private BoutonDefilement2Etiquette As String
Const BoutonBarreDefilementPositionX = 320 - 16 - 4
Const BoutonBarreDefilementPositionY = 66 + 44
Private BoutonDefilement3Etiquette As String

'Fenêtre des caratéristiques.
Const CaracNiveauPositionX = 32
Const CaracNiveauPositionY = 70 + 24 * 1 - 12
Const CaracMetierPositionX = 32
Const CaracMetierPositionY = 70 + 24 * 0 - 12
Const CaracViePositionX = 32
Const CaracViePositionY = 70 + 24 * 2 - 12
Const CaracEnergiePositionX = 32
Const CaracEnergiePositionY = 70 + 24 * 3 - 12
Const CaracMagiePositionX = 32
Const CaracMagiePositionY = 70 + 24 * 4 - 12
Const CaracMoralPositionX = 32
Const CaracMoralPositionY = 70 + 24 * 5 - 12
Const CaracVitalitePositionX = 32
Const CaracVitalitePositionY = 70 + 24 * 11 - 24 - 12
Const CaracVitessePositionX = 32
Const CaracVitessePositionY = 70 + 24 * 12 - 24 - 12
Const CaracVitesseUnites = "m/s"
Const CaracAttaquePositionX = 32
Const CaracAttaquePositionY = 70 + 24 * 7 - 12 - 12
Const CaracDefensePositionX = 32
Const CaracDefensePositionY = 70 + 24 * 8 - 12 - 12
Const CaracArmurePositionX = 32
Const CaracArmurePositionY = 70 + 24 * 9 - 12 - 12
Const CaracResistancesPositionX = 32 + 128 + 16
Const CaracResistancesPositionY = 70 + 24 * 9 - 12 - 12
Const CaracResistancesHauteur = 16
Const CaracResistancesLargeur = 72 + 48
Const CaracScoreDecalageX = 72
Const CaracScoreDecalageY = 0
Const CaracBonusDecalageX = 144 + 3
Const CaracBonusDecalageY = 0
Const CaracBonusCouleur = &H80& 'rouge foncé.
Const CaracNivPlusPositionX = 128 + 32 + 64 + 16
Const CaracNivPlusPositionY = 70 - 16 + 16 - 12
Const CaracNivPlusDecalageX = 128 + 64 + 20
Const CaracNivPlusDecalageY = 0
Const CaracNivPlusProgressionDecalageX = 183
Const CaracNivPlusProgressionDecalageY = -4
Const CaracNivPlusProgressionLargeur = 78 + 2
Const CaracNivPlusProgressionHauteur = 22 + 2


'Fenêtre des compétences.
Const CompetencesNombreLignes = 15
Const CompetencesPositionX = 32
Const CompetencesPositionY = 80 + 6 '+ 32 '96 + 16
Const CompetencesScoreDecalageX = 160 + 24 '72
Const CompetencesScoreDecalageY = 0
Const CompetencesIntervalX = 0
Const CompetencesIntervalY = 16 '12 '24
Const CompetencesNivPlusDecalageX = 192 + 16
Const CompetencesNivPlusDecalageY = 0
Const CompetencesNivPlusFormat = "00"
Const CompetenceCouleurActive = &H80& 'rouge foncé.
Const CompetenceBonusCouleur = &H80& 'rouge foncé.
Const CaracNivPlusProgressionDecalageX2 = 183
Const CaracNivPlusProgressionDecalageY2 = 0
Const CaracNivPlusProgressionLargeur2 = 78 + 2 + 1
Const CaracNivPlusProgressionHauteur2 = 22 - 2 - 3
Const CompetenceLargeurTexte = 8
Const BoutonsCompetencesNombre = 4
Const BoutonsCompetences1X = 43
Const BoutonsCompetences1Y = 61
Private BoutonsCompetences1Etiquette As String
Const BoutonsCompetences2X = 75
Const BoutonsCompetences2Y = 61
Private BoutonsCompetences2Etiquette As String
Const BoutonsCompetences3X = 106
Const BoutonsCompetences3Y = 61
Private BoutonsCompetences3Etiquette As String
Const BoutonsCompetences4X = 146
Const BoutonsCompetences4Y = 61
Private BoutonsCompetences4Etiquette As String
Const BoutonsCompetencesLargeur = 28
Const BoutonsCompetencesHauteur = 16

'Fenêtre de bestiaire.
Const BestiaireNombreLignes = 15
Const BestiairePositionX = 32
Const BestiairePositionY = 80 + 6 '+ 32 '96 + 16
Const BestiaireScoreDecalageX = 160 '72
Const BestiaireScoreDecalageY = 0
Const BestiaireIntervalX = 0
Const BestiaireIntervalY = 16 '12 '24

'Fenetre de la carte.
Const CartePositionXDepart = 16
Const CartePositionYDepart = 86 - 32
Const CarteDimension = 304 - 16
Const CarteCouleurFond = &HC0FFC0
Const CarteCouleurPerso = &H8000&     '&HC000&     '&HE0E0E0
Const CarteCouleurAlliesPersos = &HC00000
Const CarteCouleurEnnemisPersos = &HC0&
Const CarteCouleurAttaquePersos = &HC0C0FF
Const CarteCouleurMaison = &H8000& '&HC000& '&HE0E0E0
Const CarteCouleurAlliesMaisons = &HFFC0C0
Const CarteCouleurEnnemisMaisons = &HC0C0FF
Const CarteCouleurChateau = &HE0E0E0
Const CarteCouleurAlliesChateaux = &HE0E0E0
Const CarteCouleurEnnemisChateaux = &H404040
Const CarteTaillePerso = 3
Const CarteTailleMaison = 5
Const CarteTailleChateau = 8
Private CarteLargeur As Integer
Private CarteHauteur As Integer
Private CarteDecalageX As Integer
Private CarteDecalageY As Integer

'Fenêtre des objectifs.
Const ObjectifsPositionX = 32
Const ObjectifsPositionY = 96
Const ObjectifsIntervalX = 0
Const ObjectifsIntervalY = 24
Const ObjectifsFinisDecalageX = 0
Const ObjectifsFinisDecalageY = 8
Const ObjectifsFinisLongueur = 260

'Propriétés de la fenêtre.
Public Deplace As Boolean 'Si vrai, la fenêtre a été déplacée par le joueur.
Public PositionX As Long
Public PositionY As Long
Public Visible As Boolean
Public Bascule As Integer

Private AffBoutonAction As ClsAffBouton
Private BoutonDeplacer As ClsIntBouton
Private BoutonFermer As ClsIntBouton

Public NombreBoutonsManger As Long
Private BoutonsManger() As ClsIntBouton
Private IndiceRessourcesMangeable() As Long 'Mémorise le numéro de la ressource correspondant au bouton manger.
Private BoutonsRessourcesPrendre() As ClsIntBouton
Private BoutonArgentPrendre As ClsIntBouton

Private BoutonInventaireMaison As ClsIntBouton
Private BoutonInventaireOnglets() As ClsIntBouton

Private BoutonsRestaurer() As ClsIntBouton

Private BoutonService As ClsIntBouton

Private BoutonsFabriquer() As ClsIntBouton
Private BoutonsObjets() As ClsIntBouton

Private AffBoutonsCommerce As ClsAffBouton
Private BoutonsCommerce() As ClsIntBouton

Private AffBoutonsTransfert As ClsAffBouton
Private BoutonsTransfertArgentPerso() As ClsIntBouton
Private BoutonsTransfertArgentMaison() As ClsIntBouton
Private BoutonsTransfertStockPerso() As ClsIntBouton
Private BoutonsTransfertStockMaison() As ClsIntBouton

Public DefilementNombreLigne As Integer
Public DefilementDecalageX As Single
Public DefilementObjetDecalageX As Single
Private AffBoutonsDefilement As ClsAffBouton
Private AffBoutonsBarreDefilement As ClsAffBouton
Private AffBoutonsCarreDefilement As ClsAffBouton
Private BoutonsDefilement() As ClsIntBouton

Private AffBoutonsCompetences As ClsAffBouton
Private BoutonsCompetences() As ClsIntBouton

Private CarteChargee As Boolean
Private CarteExiste As Boolean
Private AffCarte As ClsAffSprite

Private Sub Class_Initialize()
    Set AffBoutonAction = New ClsAffBouton
    Set BoutonDeplacer = New ClsIntBouton
    Set BoutonFermer = New ClsIntBouton
    Set BoutonService = New ClsIntBouton
    Set BoutonInventaireMaison = New ClsIntBouton
    Set AffBoutonsCommerce = New ClsAffBouton
    Set AffBoutonsTransfert = New ClsAffBouton
    Set AffBoutonsDefilement = New ClsAffBouton
    Set AffBoutonsBarreDefilement = New ClsAffBouton
    Set AffBoutonsCarreDefilement = New ClsAffBouton
    Set AffBoutonsCompetences = New ClsAffBouton
    Set AffCarte = New ClsAffSprite
    
    ReDim BoutonsFabriquer(0)
    ReDim BoutonsObjets(0)
    Set BoutonsFabriquer(0) = New ClsIntBouton
    Set BoutonsObjets(0) = New ClsIntBouton
    
    ReDim InventaireEquipementObjetX(Monture)
    ReDim InventaireEquipementObjetY(Monture)
    InventaireEquipementObjetX(Main1) = 38
    InventaireEquipementObjetY(Main1) = 124
    InventaireEquipementObjetX(Main2) = 238
    InventaireEquipementObjetY(Main2) = 124
    InventaireEquipementObjetX(Armure) = 137
    InventaireEquipementObjetY(Armure) = 88
    InventaireEquipementObjetX(Tete) = 137
    InventaireEquipementObjetY(Tete) = 88 - 49
    InventaireEquipementObjetX(Pieds) = 137
    InventaireEquipementObjetY(Pieds) = 166 - 4 '88 + 49 + 26
    InventaireEquipementObjetX(Dos) = 137 + 50
    InventaireEquipementObjetY(Dos) = 88
    InventaireEquipementObjetX(Anneau1) = 38
    InventaireEquipementObjetY(Anneau1) = 162
    InventaireEquipementObjetX(Anneau2) = 238
    InventaireEquipementObjetY(Anneau2) = 162
    InventaireEquipementObjetX(Bras) = 88
    InventaireEquipementObjetY(Bras) = 104
    InventaireEquipementObjetX(Orbe) = 38
    InventaireEquipementObjetY(Orbe) = 72
    InventaireEquipementObjetX(Monture) = 187
    InventaireEquipementObjetY(Monture) = 162
    
    'Chargement des étiquettes.
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & NomFichierINI
    FicIni.Section = NomSectionINI
    
    BoutonDeplacerEtiquette = FicIni.Parametre("BoutonDeplacerEtiquette")
    BoutonFermerEtiquette = FicIni.Parametre("BoutonFermerEtiquette")
    ArgentBoutonsPrendreEtiquette = FicIni.Parametre("ArgentBoutonsPrendreEtiquette")
    BoutonsRestaurationEtiquette1 = FicIni.Parametre("BoutonsRestaurationEtiquette1")
    BoutonsRestaurationEtiquette2 = FicIni.Parametre("BoutonsRestaurationEtiquette2")
    BoutonsRestaurationEtiquette3 = FicIni.Parametre("BoutonsRestaurationEtiquette3")
    BoutonEmpereurInfoBullePrefixe = FicIni.Parametre("BoutonDevenirEmpereurPrefixe")
    BoutonEmpereurInfoBulleSuffixe = FicIni.Parametre("BoutonDevenirEmpereurSuffixe")
    ServiceBoutonInfoBullePrefixe = FicIni.Parametre("ServiceBoutonInfoBullePrefixe")
    ServiceBoutonInfoBulleSuffixe = FicIni.Parametre("ServiceBoutonInfoBulleSuffixe")
    BoutonsRessourcesTotalEtiquette1 = FicIni.Parametre("BoutonsRessourcesTotalEtiquette1")
    BoutonsRessourcesTotalEtiquette2 = FicIni.Parametre("BoutonsRessourcesTotalEtiquette2")
    BoutonsRessourcesTotalEtiquette3 = FicIni.Parametre("BoutonsRessourcesTotalEtiquette3")
    BoutonsRessourcesTotalEtiquette4 = FicIni.Parametre("BoutonsRessourcesTotalEtiquette4")
    BoutonsRessourcesTotalEtiquette5 = FicIni.Parametre("BoutonsRessourcesTotalEtiquette5")
    BoutonsRessourcesTotalEtiquette6 = FicIni.Parametre("BoutonsRessourcesTotalEtiquette6")
    RessourcesBoutonsPrendreEtiquette = FicIni.Parametre("RessourcesBoutonsPrendreEtiquette")
    BoutonFabriquerEtiquette = FicIni.Parametre("BoutonFabriquerEtiquette")
    BoutonTransferArgentPerso1Etiquette = FicIni.Parametre("BoutonTransferArgentPerso1Etiquette")
    BoutonTransferArgentPerso2Etiquette = FicIni.Parametre("BoutonTransferArgentPerso2Etiquette")
    BoutonTransferArgentPerso3Etiquette = FicIni.Parametre("BoutonTransferArgentPerso3Etiquette")
    BoutonTransferArgentMaison1Etiquette = FicIni.Parametre("BoutonTransferArgentMaison1Etiquette")
    BoutonTransferArgentMaison2Etiquette = FicIni.Parametre("BoutonTransferArgentMaison2Etiquette")
    BoutonTransferArgentMaison3Etiquette = FicIni.Parametre("BoutonTransferArgentMaison3Etiquette")
    BoutonTransferStockPerso1Etiquette = FicIni.Parametre("BoutonTransferStockPerso1Etiquette")
    BoutonTransferStockPerso2Etiquette = FicIni.Parametre("BoutonTransferStockPerso2Etiquette")
    BoutonTransferStockPerso3Etiquette = FicIni.Parametre("BoutonTransferStockPerso3Etiquette")
    BoutonTransferStockMaison1Etiquette = FicIni.Parametre("BoutonTransferStockMaison1Etiquette")
    BoutonTransferStockMaison2Etiquette = FicIni.Parametre("BoutonTransferStockMaison2Etiquette")
    BoutonTransferStockMaison3Etiquette = FicIni.Parametre("BoutonTransferStockMaison3Etiquette")
    BoutonAchatStockMaison1Etiquette = FicIni.Parametre("BoutonAchatStockMaison1Etiquette")
    BoutonAchatStockMaison2Etiquette = FicIni.Parametre("BoutonAchatStockMaison2Etiquette")
    BoutonAchatStockMaison3Etiquette = FicIni.Parametre("BoutonAchatStockMaison3Etiquette")
    EtiquetteObjets = FicIni.Parametre("EtiquetteObjets")
    EtiquetteStockVide = FicIni.Parametre("EtiquetteStockVide")
    EtiquetteIndisponible = FicIni.Parametre("EtiquetteIndisponible")
    InventaireMaisonMessagePerso = FicIni.Parametre("InventaireMaisonMessagePerso")
    InventaireMaisonMessageMaison = FicIni.Parametre("InventaireMaisonMessageMaison")
    InventaireMaisonMessagePersoObjet = FicIni.Parametre("InventaireMaisonMessagePersoObjet")
    InventaireMaisonMessageMaisonObjet = FicIni.Parametre("InventaireMaisonMessageMaisonObjet")
    BoutonDefilement1Etiquette = FicIni.Parametre("BoutonDefilement1Etiquette")
    BoutonDefilement2Etiquette = FicIni.Parametre("BoutonDefilement2Etiquette")
    BoutonDefilement3Etiquette = FicIni.Parametre("BoutonDefilement3Etiquette")
    BoutonsCompetences1Etiquette = FicIni.Parametre("BoutonCompetence1Etiquette")
    BoutonsCompetences2Etiquette = FicIni.Parametre("BoutonCompetence2Etiquette")
    BoutonsCompetences3Etiquette = FicIni.Parametre("BoutonCompetence3Etiquette")
    BoutonsCompetences4Etiquette = FicIni.Parametre("BoutonCompetence4Etiquette")
End Sub

Public Sub Init(ByRef dd As DirectDraw7, ByRef Parametres As ClsJeuParam)
    Dim i, j As Long
    Dim X, Y As Integer
    
    BoutonDeplacer.PositionX = BoutonDeplacerPositionX
    BoutonDeplacer.PositionY = BoutonDeplacerPositionY
    BoutonDeplacer.InfoBulle = BoutonDeplacerEtiquette
    
    BoutonFermer.PositionX = BoutonFermerPositionX
    BoutonFermer.PositionY = BoutonFermerPositionY
    BoutonFermer.InfoBulle = BoutonFermerEtiquette
    
    ReDim BoutonsRestaurer(NombreBoutonsRestauration - 1)
    For i = 0 To NombreBoutonsRestauration - 1
        Set BoutonsRestaurer(i) = New ClsIntBouton
    Next i
    BoutonsRestaurer(0).InfoBulle = BoutonsRestaurationEtiquette1
    BoutonsRestaurer(1).InfoBulle = BoutonsRestaurationEtiquette2
    BoutonsRestaurer(2).InfoBulle = BoutonsRestaurationEtiquette3
    
    'Boutons manger.
    ReDim IndiceRessourcesMangeable(0)
    ReDim BoutonsManger(0)
    For i = 0 To Parametres.NombreRessources - 1
        If Parametres.Ressources_Mangeable(i) Then
            NombreBoutonsManger = NombreBoutonsManger + 1
            ReDim Preserve BoutonsManger(NombreBoutonsManger - 1)
            Set BoutonsManger(NombreBoutonsManger - 1) = New ClsIntBouton
            ReDim Preserve IndiceRessourcesMangeable(NombreBoutonsManger - 1)
            IndiceRessourcesMangeable(NombreBoutonsManger - 1) = i
        End If
    Next i
    
    ReDim BoutonsCommerce((Parametres.NombreRessources + 1) * NombreBoutonsCommerce - 1)
    AffBoutonsCommerce.Charger dd, 4
    For i = 0 To Parametres.NombreRessources - 1
        For j = 0 To NombreBoutonsCommerce - 1
            Set BoutonsCommerce(i * NombreBoutonsCommerce + j) = New ClsIntBouton
            'BoutonsCommerce(i * NombreBoutonsCommerce + j).InfoBulle = Parametres.EtiquettePrix & " : " & Parametres.Ressources_PrixVente(i)
            BoutonsCommerce((i + DefilementDecalageX) * (NombreBoutonsCommerce) + j).PositionX = RessourcesPositionX + RessourcesIntervalX * i + BoutonCommerceDecalageX + BoutonCommerceIntervalX * j
            BoutonsCommerce((i + DefilementDecalageX) * (NombreBoutonsCommerce) + j).PositionY = RessourcesPositionY + RessourcesIntervalY * i + BoutonCommerceDecalageY + BoutonCommerceIntervalY * j
        Next j
    Next i
    
    For i = 0 To NombreBoutonsCommerce - 1
        Set BoutonsCommerce(UBound(BoutonsCommerce()) - i) = New ClsIntBouton
        With BoutonsCommerce(UBound(BoutonsCommerce()) - i)
        .PositionX = BoutonsRessourcesTotalPositionX + BoutonsRessourcesTotalIntervalX * i
        .PositionY = BoutonsRessourcesTotalPositionY + BoutonsRessourcesTotalIntervalY * i
        .Visible = True
        'Select Case i
        'Case 0: .InfoBulle = BoutonsRessourcesTotalEtiquette1
        'Case 1: .InfoBulle = BoutonsRessourcesTotalEtiquette2
        'Case 2: .InfoBulle = BoutonsRessourcesTotalEtiquette3
        'End Select
        End With
    Next i
    'Bouton pour prendre l'argent.
    Set BoutonArgentPrendre = New ClsIntBouton
    BoutonArgentPrendre.PositionX = ArgentBoutonsPrendrePositionX
    BoutonArgentPrendre.PositionY = ArgentBoutonsPrendrePositionY
    BoutonArgentPrendre.Largeur = ArgentBoutonsPrendreLargeur
    BoutonArgentPrendre.Hauteur = ArgentBoutonsPrendreHauteur
    BoutonArgentPrendre.InfoBulle = ArgentBoutonsPrendreEtiquette
    'Boutons pour prendre les ressources.
    ReDim BoutonsRessourcesPrendre(Parametres.NombreRessources - 1)
    For i = 0 To Parametres.NombreRessources - 1
        Set BoutonsRessourcesPrendre(i) = New ClsIntBouton
        BoutonsRessourcesPrendre(i).InfoBulle = RessourcesBoutonsPrendreEtiquette & " " & Parametres.Ressources_Nom(i) & "  "
        BoutonsRessourcesPrendre(i).Largeur = RessourcesBoutonsPrendreLargeur
        BoutonsRessourcesPrendre(i).Hauteur = RessourcesBoutonsPrendreHauteur
    Next i
    
    ReDim BoutonsTransfertArgentPerso(NombreBoutonsTransfert)
    For i = 0 To NombreBoutonsTransfert - 1
        Set BoutonsTransfertArgentPerso(i) = New ClsIntBouton
        BoutonsTransfertArgentPerso(i).Visible = True
    Next i
    BoutonsTransfertArgentPerso(0).InfoBulle = BoutonTransferArgentPerso1Etiquette
    BoutonsTransfertArgentPerso(1).InfoBulle = BoutonTransferArgentPerso2Etiquette
    BoutonsTransfertArgentPerso(2).InfoBulle = BoutonTransferArgentPerso3Etiquette
    ReDim BoutonsTransfertArgentMaison(NombreBoutonsTransfert)
    For i = 0 To NombreBoutonsTransfert - 1
        Set BoutonsTransfertArgentMaison(i) = New ClsIntBouton
    Next i
    BoutonsTransfertArgentMaison(0).InfoBulle = BoutonTransferArgentMaison1Etiquette
    BoutonsTransfertArgentMaison(1).InfoBulle = BoutonTransferArgentMaison2Etiquette
    BoutonsTransfertArgentMaison(2).InfoBulle = BoutonTransferArgentMaison3Etiquette
    ReDim BoutonsTransfertStockPerso(NombreBoutonsTransfert)
    For i = 0 To NombreBoutonsTransfert - 1
        Set BoutonsTransfertStockPerso(i) = New ClsIntBouton
    Next i
    BoutonsTransfertStockPerso(0).InfoBulle = BoutonTransferStockPerso1Etiquette
    BoutonsTransfertStockPerso(1).InfoBulle = BoutonTransferStockPerso2Etiquette
    BoutonsTransfertStockPerso(2).InfoBulle = BoutonTransferStockPerso3Etiquette
    ReDim BoutonsTransfertStockMaison(NombreBoutonsTransfert)
    For i = 0 To NombreBoutonsTransfert - 1
        Set BoutonsTransfertStockMaison(i) = New ClsIntBouton
    Next i
    BoutonsTransfertStockMaison(0).InfoBulle = BoutonTransferStockMaison1Etiquette
    BoutonsTransfertStockMaison(1).InfoBulle = BoutonTransferStockMaison2Etiquette
    BoutonsTransfertStockMaison(2).InfoBulle = BoutonTransferStockMaison3Etiquette
    AffBoutonsTransfert.Charger dd, 5
    
    BoutonInventaireMaison.PositionX = InventaireMaisonX
    BoutonInventaireMaison.PositionY = InventaireMaisonY
    
    NombreBoutonsInventaireOnglets = Parametres.PersosNombreObjetsInventaire / InventaireEquipementNombreCases
    ReDim BoutonInventaireOnglets(NombreBoutonsInventaireOnglets - 1)
    For i = 0 To NombreBoutonsInventaireOnglets - 1
        Set BoutonInventaireOnglets(i) = New ClsIntBouton
        BoutonInventaireOnglets(i).PositionX = InventaireOngletsPositionX + InventaireOngletsDecalageX * i
        BoutonInventaireOnglets(i).PositionY = InventaireOngletsPositionY + InventaireOngletsDecalageY * i
    Next i
    
    BoutonService.PositionX = ServiceBoutonPositionX
    BoutonService.PositionY = ServiceBoutonPositionY
    
    ReDim BoutonsObjets(Parametres.PersosNombreObjetsInventaire - 1 + Parametres.PersosNombreObjetsEquipes)
    Y = -1
    For i = 0 To UBound(BoutonsObjets())
        If i Mod InventaireEquipementNombreCases = 0 Then
            Y = -1
        End If
        If i Mod InventaireNombreColonnes = 0 And i > 1 Then
            Y = Y + 1
            X = 0
        End If
        Set BoutonsObjets(i) = New ClsIntBouton
        With BoutonsObjets(i)
        .PositionX = InventairePositionX + X * InventaireDecalageX
        .PositionY = InventairePositionY + Y * InventaireDecalageY
        .Enabled = True
        End With
        X = X + 1
    Next i
    For i = Parametres.PersosNombreObjetsInventaire To UBound(BoutonsObjets())
        Set BoutonsObjets(i) = New ClsIntBouton
        With BoutonsObjets(i)
        .PositionX = InventaireEquipementObjetX(i - Parametres.PersosNombreObjetsInventaire)
        .PositionY = InventaireEquipementObjetY(i - Parametres.PersosNombreObjetsInventaire)
        .Enabled = True
        End With
    Next i
    
    ReDim BoutonsDefilement(NombreBoutonsDefilement - 1)
    For i = 0 To NombreBoutonsDefilement - 1
        Set BoutonsDefilement(i) = New ClsIntBouton
    Next i
    BoutonsDefilement(0).InfoBulle = BoutonDefilement1Etiquette
    BoutonsDefilement(1).InfoBulle = BoutonDefilement2Etiquette
    BoutonsDefilement(2).InfoBulle = BoutonDefilement3Etiquette
    
    BoutonsDefilement(0).PositionX = BoutonDefilement1PositionX
    BoutonsDefilement(0).PositionY = BoutonDefilement1PositionY
    BoutonsDefilement(1).PositionX = BoutonDefilement2PositionX
    BoutonsDefilement(1).PositionY = BoutonDefilement2PositionY
    BoutonsDefilement(2).PositionX = BoutonBarreDefilementPositionX
    BoutonsDefilement(2).PositionY = BoutonBarreDefilementPositionY
    
    AffBoutonsDefilement.Charger dd, 6
    AffBoutonsBarreDefilement.Charger dd, 8
    AffBoutonsCarreDefilement.Charger dd, 9
    
    AffBoutonAction.Charger dd, 2
    
    ReDim BoutonsCompetences(BoutonsCompetencesNombre - 1)
    For i = 0 To BoutonsCompetencesNombre - 1
        Set BoutonsCompetences(i) = New ClsIntBouton
    Next i
    BoutonsCompetences(0).PositionX = BoutonsCompetences1X
    BoutonsCompetences(0).PositionY = BoutonsCompetences1Y
    BoutonsCompetences(0).InfoBulle = BoutonsCompetences1Etiquette
    BoutonsCompetences(1).PositionX = BoutonsCompetences2X
    BoutonsCompetences(1).PositionY = BoutonsCompetences2Y
    BoutonsCompetences(1).InfoBulle = BoutonsCompetences2Etiquette
    BoutonsCompetences(2).PositionX = BoutonsCompetences3X
    BoutonsCompetences(2).PositionY = BoutonsCompetences3Y
    BoutonsCompetences(2).InfoBulle = BoutonsCompetences3Etiquette
    BoutonsCompetences(3).PositionX = BoutonsCompetences4X
    BoutonsCompetences(3).PositionY = BoutonsCompetences4Y
    BoutonsCompetences(3).InfoBulle = BoutonsCompetences4Etiquette
    AffBoutonsCompetences.Charger dd, 11
End Sub

Public Sub Afficher_InfosMaison(ByVal Titre As String, _
                                ByVal FileFabricationObjets As Collection, _
                                ByVal Interface As ClsInterface, _
                                ByRef Perso As ClsJeuPerso, _
                                ByRef Maison As ClsJeuBatiment, _
                                ByRef Parametres As ClsJeuParam, _
                                ByRef Fenetre As ClsAffFenetre, _
                                ByRef AffObjets As ClsAffObjets, _
                                ByRef Message As ClsIntMessage, _
                                ByRef SurfBack As DirectDrawSurface7)
    Afficher_Titre Titre, Message, SurfBack
    Afficher_BoutonFermer SurfBack
    Afficher_EtatBatiment Perso, Maison, Parametres, Message, SurfBack
    If Perso.DansSaMaison Then
        Afficher_Argent_Batiment Perso, Maison, Message, SurfBack
    End If
    Afficher_StockBatiment Interface, Perso, Maison, Parametres, Message, SurfBack
    Afficher_Objets_Batiment FileFabricationObjets, Perso, Maison, Parametres, Message, AffObjets, SurfBack
End Sub

Public Sub Afficher_InfosChateau(ByVal Titre As String, _
                                 ByRef Perso As ClsJeuPerso, _
                                 ByRef Maison As ClsJeuBatiment, _
                                 ByRef Chateau As ClsJeuBatiment, _
                                 ByRef Fief As ClsJeuFief, _
                                 ByRef Parametres As ClsJeuParam, _
                                 ByRef Ressources As ClsAffObjets, _
                                 ByRef Fenetre As ClsAffFenetre, _
                                 ByRef Message As ClsIntMessage, _
                                 ByRef SurfBack As DirectDrawSurface7)
    Afficher_Titre Titre, Message, SurfBack
    Afficher_BoutonFermer SurfBack
    'Afficher_BoutonsBatiment SurfBack
    If Perso.DansSaMaison Or Perso.DansUneMaison Then
        Afficher_EtatBatiment Perso, Maison, Parametres, Message, SurfBack
    Else
        Afficher_EtatBatiment Perso, Chateau, Parametres, Message, SurfBack, True
    End If
    Afficher_ListeRessourcesFief Perso, Fief, Parametres, Ressources, Fenetre, Message, SurfBack
End Sub


Public Sub Afficher_InfosRessourcesPerso(ByVal Titre As String, _
                                         ByRef Perso As ClsJeuPerso, _
                                         ByRef Maison As ClsJeuBatiment, _
                                         ByRef Parametres As ClsJeuParam, _
                                         ByRef Ressources As ClsAffObjets, _
                                         ByRef Fenetre As ClsAffFenetre, _
                                         ByRef Message As ClsIntMessage, _
                                         ByRef SurfBack As DirectDrawSurface7)
    Afficher_Titre Titre, Message, SurfBack
    Afficher_BoutonFermer SurfBack
    Afficher_ListeRessourcesPerso Perso, Maison, Parametres, Ressources, Fenetre, Message, SurfBack
    Afficher_Argent_Perso Perso, Message, SurfBack
End Sub

Public Sub Afficher_InfosCaracPerso1(ByVal Titre As String, _
                                    ByRef Perso As ClsJeuPerso, _
                                    ByRef Maison As ClsJeuBatiment, _
                                    ByRef Parametres As ClsJeuParam, _
                                    ByRef Fenetre As ClsAffFenetre, _
                                    ByRef Message As ClsIntMessage, _
                                    ByRef SurfBack As DirectDrawSurface7)
    'Affichage des barres de progression sur le changement de niveau.
    Dim Niveau As Double
    Niveau = Perso.Niveau
    SurfBack.SetFillColor &HFFFFFF
    Interface(0).Afficher_Barre_Statut Niveau - Int(Niveau), 0, 1, CaracNiveauPositionX + CaracNivPlusProgressionDecalageX + PositionX, CaracNiveauPositionY + CaracNivPlusProgressionDecalageY + PositionY, SurfBack, CaracNivPlusProgressionLargeur, CaracNivPlusProgressionHauteur, &HFFFFFF
    '.Afficher CaracNiveauPositionX + CaracNivPlusDecalageX + PositionX, CaracNiveauPositionY + CaracNivPlusDecalageY + PositionY, _
                     Int((Format(Niveau, "0.00") - Int(Niveau)) * 100) & " %", SurfBack
    Interface(0).Afficher_Barre_Statut Perso.Progression_Vie, 0, 1, CaracViePositionX + CaracNivPlusProgressionDecalageX + PositionX, CaracViePositionY + CaracNivPlusProgressionDecalageY + PositionY, SurfBack, CaracNivPlusProgressionLargeur, CaracNivPlusProgressionHauteur, &HFFFFFF
    '.Afficher CaracViePositionX + CaracNivPlusDecalageX + PositionX, CaracViePositionY + CaracNivPlusDecalageY + PositionY, _
                     Int(Perso.Progression_Vie * 100) & " %", SurfBack
    Interface(0).Afficher_Barre_Statut Perso.Progression_Energie, 0, 1, CaracEnergiePositionX + CaracNivPlusProgressionDecalageX + PositionX, CaracEnergiePositionY + CaracNivPlusProgressionDecalageY + PositionY, SurfBack, CaracNivPlusProgressionLargeur, CaracNivPlusProgressionHauteur, &HFFFFFF
    '.Afficher CaracEnergiePositionX + CaracNivPlusDecalageX + PositionX, CaracEnergiePositionY + CaracNivPlusDecalageY + PositionY, _
                     Int(Perso.Progression_energie * 100) & "%", SurfBack
    Interface(0).Afficher_Barre_Statut Perso.Progression_Magie, 0, 1, CaracMagiePositionX + CaracNivPlusProgressionDecalageX + PositionX, CaracMagiePositionY + CaracNivPlusProgressionDecalageY + PositionY, SurfBack, CaracNivPlusProgressionLargeur, CaracNivPlusProgressionHauteur, &HFFFFFF
    '.Afficher CaracMagiePositionX + CaracNivPlusDecalageX + PositionX, CaracMagiePositionY + CaracNivPlusDecalageY + PositionY, _
                     Int(Perso.Progression_Magie * 100) & "%", SurfBack
    Interface(0).Afficher_Barre_Statut Perso.Progression_Moral, 0, 1, CaracMoralPositionX + CaracNivPlusProgressionDecalageX + PositionX, CaracMoralPositionY + CaracNivPlusProgressionDecalageY + PositionY, SurfBack, CaracNivPlusProgressionLargeur, CaracNivPlusProgressionHauteur, &HFFFFFF
    '.Afficher CaracMoralPositionX + CaracNivPlusDecalageX + PositionX, CaracMoralPositionY + CaracNivPlusDecalageY + PositionY, _
                     Int(Perso.Progression_Moral * 100) & "%", SurfBack
    Interface(0).Afficher_Barre_Statut Perso.Progression_Attaque, 0, 1, CaracAttaquePositionX + CaracNivPlusProgressionDecalageX + PositionX, CaracAttaquePositionY + CaracNivPlusProgressionDecalageY + PositionY, SurfBack, CaracNivPlusProgressionLargeur, CaracNivPlusProgressionHauteur, &HFFFFFF
    '.Afficher CaracAttaquePositionX + CaracNivPlusDecalageX + PositionX, CaracAttaquePositionY + CaracNivPlusDecalageY + PositionY, _
                     Int(Perso.Progression_Attaque * 100) & "%", SurfBack
    Interface(0).Afficher_Barre_Statut Perso.Progression_Defense, 0, 1, CaracDefensePositionX + CaracNivPlusProgressionDecalageX + PositionX, CaracDefensePositionY + CaracNivPlusProgressionDecalageY + PositionY, SurfBack, CaracNivPlusProgressionLargeur, CaracNivPlusProgressionHauteur, &HFFFFFF
    '.Afficher CaracDefensePositionX + CaracNivPlusDecalageX + PositionX, CaracDefensePositionY + CaracNivPlusDecalageY + PositionY, _
                     Int(Perso.Progression_Defense * 100) & "%", SurfBack
End Sub

Public Sub Afficher_InfosCaracPerso2(ByVal Titre As String, _
                                    ByRef Perso As ClsJeuPerso, _
                                    ByRef Maison As ClsJeuBatiment, _
                                    ByRef Parametres As ClsJeuParam, _
                                    ByRef Fenetre As ClsAffFenetre, _
                                    ByRef Message As ClsIntMessage, _
                                    ByRef SurfBack As DirectDrawSurface7)
    Dim Niveau As Double
    'Afficher_Titre Titre, Message, SurfBack
    Afficher_BoutonFermer SurfBack
    'Affichage des caratéristiques.
    'Affiche le nom.
    
    SurfBack.SetForeColor 0
    
    With Message
    
    .Afficher TitrePositionX + PositionX, TitrePositionY + PositionY, _
                     Perso.Nom, SurfBack
    
    'Affiche le niveau.
    Niveau = Perso.Niveau
    .Afficher CaracNiveauPositionX + PositionX, CaracNiveauPositionY + PositionY, _
                     Parametres.EtiquetteNiveau & " :", SurfBack
    .Afficher CaracNiveauPositionX + CaracScoreDecalageX + PositionX, CaracNiveauPositionY + CaracScoreDecalageY + PositionY, _
                     Int(Niveau), SurfBack
    
    'Affiche le métier.
    If Perso.Empereur Then
        'Est l'empereur.
        .Afficher CaracMetierPositionX + CaracScoreDecalageX + PositionX, CaracMetierPositionY + CaracScoreDecalageY + PositionY, _
                         Parametres.Epoque_EmpereurNom(Fiefs(Perso.NumeroFief).Epoque, Perso.Feminin), SurfBack
    ElseIf Maison.Construit Then
        .Afficher CaracMetierPositionX + PositionX, CaracMetierPositionY + PositionY, _
                         Parametres.EtiquetteMetier & " :", SurfBack
            .Afficher CaracMetierPositionX + CaracScoreDecalageX + PositionX, CaracMetierPositionY + CaracScoreDecalageY + PositionY, _
                             Parametres.Batiment_Metier(Maison.TypeBatiment, Perso.Feminin), SurfBack
    End If
    
    'Affiche la vie.
    .Afficher CaracViePositionX + PositionX, CaracViePositionY + PositionY, _
                     Parametres.EtiquetteVie & " :", SurfBack
    .Afficher CaracViePositionX + CaracScoreDecalageX + PositionX, CaracViePositionY + CaracScoreDecalageY + PositionY, _
                     Int(Perso.Vie) & " / " & Perso.Definir_MaxVie, SurfBack
    'Affiche l'énergie.
    .Afficher CaracEnergiePositionX + PositionX, CaracEnergiePositionY + PositionY, _
                     Parametres.EtiquetteEnergie & " :", SurfBack
    .Afficher CaracEnergiePositionX + CaracScoreDecalageX + PositionX, CaracEnergiePositionY + CaracScoreDecalageY + PositionY, _
                     Int(Perso.Energie) & " / " & Perso.Definir_MaxEnergie, SurfBack
                     'Format(Perso.Energie, "0.000000") & " / " & Perso.MaxEnergie, SurfBack
    'Affiche la magie.
    .Afficher CaracMagiePositionX + PositionX, CaracMagiePositionY + PositionY, _
                     Parametres.EtiquetteMagie & " :", SurfBack
    .Afficher CaracMagiePositionX + CaracScoreDecalageX + PositionX, CaracMagiePositionY + CaracScoreDecalageY + PositionY, _
                     Int(Perso.Magie) & " / " & Perso.Definir_MaxMagie, SurfBack
    'Affiche le moral.
    .Afficher CaracMoralPositionX + PositionX, CaracMoralPositionY + PositionY, _
                     Parametres.EtiquetteMoral & " :", SurfBack
    .Afficher CaracMoralPositionX + CaracScoreDecalageX + PositionX, CaracMoralPositionY + CaracScoreDecalageY + PositionY, _
                     Int(Perso.Moral) & " / " & Perso.Definir_MaxMoral, SurfBack
    
    'Affiche La vitalité.
    .Afficher CaracVitalitePositionX + PositionX, CaracVitalitePositionY + PositionY, _
                     Parametres.EtiquetteVitalite & " :", SurfBack
    .Afficher CaracVitalitePositionX + CaracScoreDecalageX + PositionX, CaracVitalitePositionY + CaracScoreDecalageY + PositionY, _
                     Int(Perso.Vitalite * 100) & " %", SurfBack
    'Affiche la vitesse.
    .Afficher CaracVitessePositionX + PositionX, CaracVitessePositionY + PositionY, _
                     Parametres.EtiquetteVitesse & " :", SurfBack
    If Perso.Longueur_PasTemp = Perso.LongueurPasMax Then
        .Afficher CaracVitessePositionX + CaracScoreDecalageX + PositionX, CaracVitessePositionY + CaracScoreDecalageY + PositionY, _
                         UCase(Parametres.EtiquetteMaximumAbr), SurfBack
    Else
        .Afficher CaracVitessePositionX + CaracScoreDecalageX + PositionX, CaracVitessePositionY + CaracScoreDecalageY + PositionY, _
                         IIf(Perso.Vivant, Perso.Longueur_PasTemp & " / " & Int(Perso.BonusDeplacement * Perso.Vitalite), Perso.Definir_VitesseDeplacementFantome), SurfBack
    End If
    
    'Affiche l'attaque.
    .Afficher CaracAttaquePositionX + PositionX, CaracAttaquePositionY + PositionY, _
                     Parametres.EtiquetteAttaque & " :", SurfBack
    .Afficher CaracAttaquePositionX + CaracScoreDecalageX + PositionX, CaracAttaquePositionY + CaracScoreDecalageY + PositionY, _
                     Int(Perso.Definir_Attaque) & " / " & Perso.Definir_MaxAttaque, SurfBack
    'Affiche la défense.
    .Afficher CaracDefensePositionX + PositionX, CaracDefensePositionY + PositionY, _
                     Parametres.EtiquetteDefense & " :", SurfBack
    .Afficher CaracDefensePositionX + CaracScoreDecalageX + PositionX, CaracDefensePositionY + CaracScoreDecalageY + PositionY, _
                     Int(Perso.Definir_Defense) & " / " & Perso.Definir_MaxDefense, SurfBack
    'Affiche l'armure.
    .Afficher CaracArmurePositionX + PositionX, CaracArmurePositionY + PositionY, _
                     Parametres.EtiquetteArmure & " :", SurfBack
    .Afficher CaracArmurePositionX + CaracScoreDecalageX + PositionX, CaracArmurePositionY + CaracScoreDecalageY + PositionY, _
                     Perso.Definir_Armure, SurfBack
    
    'Affiche les résistances.
    .Afficher CaracResistancesPositionX + PositionX, CaracResistancesPositionY + PositionY, _
                     Parametres.EtiquetteResistances, SurfBack
    
    'Affichage des informations sur le changement de niveau.
    .Afficher CaracNivPlusPositionX + PositionX, CaracNivPlusPositionY + PositionY, _
                     Parametres.EtiquetteNivPlus & " :", SurfBack
    .Afficher CaracNiveauPositionX + CaracNivPlusDecalageX + PositionX, CaracNiveauPositionY + CaracNivPlusDecalageY + PositionY, _
                     Int((Format(Niveau, "0.00") - Int(Niveau)) * 100) & " %", SurfBack
    .Afficher CaracViePositionX + CaracNivPlusDecalageX + PositionX, CaracViePositionY + CaracNivPlusDecalageY + PositionY, _
                     Int(Perso.Progression_Vie * 100) & " %", SurfBack
    .Afficher CaracEnergiePositionX + CaracNivPlusDecalageX + PositionX, CaracEnergiePositionY + CaracNivPlusDecalageY + PositionY, _
                     Int(Perso.Progression_Energie * 100) & "%", SurfBack
    .Afficher CaracMagiePositionX + CaracNivPlusDecalageX + PositionX, CaracMagiePositionY + CaracNivPlusDecalageY + PositionY, _
                     Int(Perso.Progression_Magie * 100) & "%", SurfBack
    .Afficher CaracMoralPositionX + CaracNivPlusDecalageX + PositionX, CaracMoralPositionY + CaracNivPlusDecalageY + PositionY, _
                     Int(Perso.Progression_Moral * 100) & "%", SurfBack
    .Afficher CaracAttaquePositionX + CaracNivPlusDecalageX + PositionX, CaracAttaquePositionY + CaracNivPlusDecalageY + PositionY, _
                     Int(Perso.Progression_Attaque * 100) & "%", SurfBack
    .Afficher CaracDefensePositionX + CaracNivPlusDecalageX + PositionX, CaracDefensePositionY + CaracNivPlusDecalageY + PositionY, _
                     Int(Perso.Progression_Defense * 100) & "%", SurfBack
                     
    'Affiche les bonus
    SurfBack.SetForeColor CaracBonusCouleur
    If Perso.Definir_BonusVie >= 1 Or Perso.Definir_BonusVie <= -1 Then
        .Afficher CaracViePositionX + CaracBonusDecalageX + PositionX, CaracViePositionY + CaracBonusDecalageY + PositionY, _
                         IIf(Perso.Definir_BonusVie >= 1, "+", "") & Int(Perso.Definir_BonusVie), SurfBack
    End If
    If Perso.Definir_BonusEnergie >= 1 Or Perso.Definir_BonusEnergie <= -1 Then
        .Afficher CaracEnergiePositionX + CaracBonusDecalageX + PositionX, CaracEnergiePositionY + CaracBonusDecalageY + PositionY, _
                          IIf(Perso.Definir_BonusEnergie >= 1, "+", "") & Int(Perso.Definir_BonusEnergie), SurfBack
    End If
    If Perso.Definir_BonusMagie >= 1 Or Perso.Definir_BonusMagie <= -1 Then
        .Afficher CaracMagiePositionX + CaracBonusDecalageX + PositionX, CaracMagiePositionY + CaracBonusDecalageY + PositionY, _
                          IIf(Perso.Definir_BonusMagie >= 1, "+", "") & Int(Perso.Definir_BonusMagie), SurfBack
    End If
    If Perso.Definir_BonusMoral >= 1 Or Perso.Definir_BonusMoral <= -1 Then
        .Afficher CaracMoralPositionX + CaracBonusDecalageX + PositionX, CaracMoralPositionY + CaracBonusDecalageY + PositionY, _
                          IIf(Perso.Definir_BonusMoral >= 1, "+", "") & Int(Perso.Definir_BonusMoral), SurfBack
    End If
    If Perso.Definir_BonusAttaque >= 1 Or Perso.Definir_BonusAttaque <= -1 Then
        .Afficher CaracAttaquePositionX + CaracBonusDecalageX + PositionX, CaracAttaquePositionY + CaracBonusDecalageY + PositionY, _
                         IIf(Perso.Definir_BonusAttaque >= 1, "+", "") & Int(Perso.Definir_BonusAttaque), SurfBack
    End If
    If Perso.Definir_BonusDefense >= 1 Or Perso.Definir_BonusDefense <= -1 Then
        .Afficher CaracDefensePositionX + CaracBonusDecalageX + PositionX, CaracDefensePositionY + CaracBonusDecalageY + PositionY, _
                         IIf(Perso.Definir_BonusDefense >= 1, "+", "") & Int(Perso.Definir_BonusDefense), SurfBack
    End If
    If Perso.BonusVitalite <> 1 And Perso.Vivant Then
        .Afficher CaracVitalitePositionX + CaracBonusDecalageX + PositionX, CaracVitalitePositionY + CaracBonusDecalageY + PositionY, _
                          IIf(Perso.BonusVitalite > 1, "+", "-") & Int((Perso.BonusVitalite - 1) * 100) & Parametres.EtiquettePourcent, SurfBack
    End If
    If (Perso.Definir_BonusDeplacement >= 1 Or Perso.Definir_BonusDeplacement <= -1) And Perso.Vivant Then
        .Afficher CaracVitessePositionX + CaracBonusDecalageX + PositionX, CaracVitessePositionY + CaracBonusDecalageY + PositionY, _
                          IIf(Perso.Definir_BonusDeplacement >= 1, "+", "") & Perso.Definir_BonusDeplacement, SurfBack
    End If
    
    End With

End Sub

Public Sub Afficher_InfosCompetences1(ByVal TabComp As Collection, _
                                      ByVal Titre As String, _
                                      ByRef Perso As ClsJeuPerso, _
                                      ByRef Maison As ClsJeuBatiment, _
                                      ByRef Parametres As ClsJeuParam, _
                                      ByRef Fenetre As ClsAffFenetre, _
                                      ByRef Message As ClsIntMessage, _
                                      ByRef SurfBack As DirectDrawSurface7)
    'Affichage des barres de progression pour les compétences.
    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    SurfBack.SetFillColor &HFFFFFF
    X = -DefilementDecalageX
'    For i = 0 To Perso.Nombre_CompetencesSpeciales - 1
'        If X >= 0 And X <= CompetencesNombreLignes Then
'            Interface(0).Afficher_Barre_Statut Perso.Progression_CompetenceSpeciale(i), 0, 1, CompetencesPositionX + CompetencesIntervalX * X + CaracNivPlusProgressionDecalageX2 + PositionX, CompetencesPositionY + CompetencesIntervalY * X + CaracNivPlusProgressionDecalageY2 + PositionY, SurfBack, CaracNivPlusProgressionLargeur2, CaracNivPlusProgressionHauteur2, &HFFFFFF
'        End If
'        X = X + 1
'    Next i
'    For i = 0 To Perso.Nombre_CompetencesRessources - 1
'        If X >= 0 And X <= CompetencesNombreLignes Then
'            Interface(0).Afficher_Barre_Statut Perso.Progression_CompetenceRessource(i), 0, 1, CompetencesPositionX + CompetencesIntervalX * X + CaracNivPlusProgressionDecalageX2 + PositionX, CompetencesPositionY + CompetencesIntervalY * X + CaracNivPlusProgressionDecalageY2 + PositionY, SurfBack, CaracNivPlusProgressionLargeur2, CaracNivPlusProgressionHauteur2, &HFFFFFF
'        End If
'        X = X + 1
'    Next i
'    For i = 0 To Perso.Nombre_CompetencesServices - 1
'        If X >= 0 And X <= CompetencesNombreLignes Then
'            Interface(0).Afficher_Barre_Statut Perso.Progression_CompetenceService(i), 0, 1, CompetencesPositionX + CompetencesIntervalX * X + CaracNivPlusProgressionDecalageX2 + PositionX, CompetencesPositionY + CompetencesIntervalY * X + CaracNivPlusProgressionDecalageY2 + PositionY, SurfBack, CaracNivPlusProgressionLargeur2, CaracNivPlusProgressionHauteur2, &HFFFFFF
'        End If
'        X = X + 1
'    Next i
'    For i = 0 To Perso.Nombre_CompetencesObjets - 1
'        If X >= 0 And X <= CompetencesNombreLignes Then
'            Interface(0).Afficher_Barre_Statut Perso.Progression_CompetenceObjet(i), 0, 1, CompetencesPositionX + CompetencesIntervalX * X + CaracNivPlusProgressionDecalageX2 + PositionX, CompetencesPositionY + CompetencesIntervalY * X + CaracNivPlusProgressionDecalageY2 + PositionY, SurfBack, CaracNivPlusProgressionLargeur2, CaracNivPlusProgressionHauteur2, &HFFFFFF
'        End If
'        X = X + 1
'    Next i
    For i = DefilementDecalageX + 1 To TabComp.Count
        If Y > CompetencesNombreLignes Then Exit For
        Interface(0).Afficher_Barre_Statut TabComp(i).NivPlus, 0, 1, CompetencesPositionX + CompetencesIntervalX * Y + CaracNivPlusProgressionDecalageX2 + PositionX, CompetencesPositionY + CompetencesIntervalY * Y + CaracNivPlusProgressionDecalageY2 + PositionY, SurfBack, CaracNivPlusProgressionLargeur2, CaracNivPlusProgressionHauteur2, &HFFFFFF
        X = X + 1: Y = Y + 1
    Next i
End Sub

Public Sub Afficher_InfosCompetences2(ByVal TabComp As Collection, _
                                      ByVal Titre As String, _
                                      ByVal CompetencesTri As Integer, _
                                      ByVal CompetencesAfficherToutes As Boolean, _
                                      ByRef Fenetre As ClsAffFenetre, _
                                      ByRef Message As ClsIntMessage, _
                                      ByRef SurfBack As DirectDrawSurface7)
    'Dim i, j, k As Integer
    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    Afficher_Titre Titre, Message, SurfBack
    Afficher_BoutonFermer SurfBack
    
    With Message
    
    .Afficher CaracNivPlusPositionX + PositionX, _
                     CaracNivPlusPositionY + PositionY, _
                     Parametres.EtiquetteNivPlus, SurfBack
    If TabComp.Count > 0 Then
        X = -DefilementDecalageX
        'y = 0
        'Liste des compétences.
        For i = DefilementDecalageX + 1 To TabComp.Count
            If Y > CompetencesNombreLignes Then Exit For
            'Colorie la compétence active.
            If TabComp(i).Active Then
                SurfBack.SetForeColor CompetenceCouleurActive
            Else
                SurfBack.SetForeColor 0
            End If
            .Afficher CompetencesPositionX + CompetencesIntervalX * Y + PositionX, _
                             CompetencesPositionY + CompetencesIntervalY * Y + PositionY, _
                             TabComp(i).Nom, _
                             SurfBack
            If TabComp(i).Bonus Then
                SurfBack.SetForeColor CompetenceBonusCouleur
            Else
                SurfBack.SetForeColor 0
            End If
            .Afficher CompetencesPositionX + CompetencesIntervalX * Y + CompetencesScoreDecalageX + PositionX - CompetenceLargeurTexte * Len(Str(TabComp(i).Valeur)), _
                      CompetencesPositionY + CompetencesIntervalY * Y + CompetencesScoreDecalageY + PositionY, _
                      TabComp(i).Valeur, _
                      SurfBack
            SurfBack.SetForeColor 0
            .Afficher CompetencesPositionX + CompetencesIntervalX * Y + CompetencesNivPlusDecalageX + PositionX, _
                             CompetencesPositionY + CompetencesIntervalY * Y + CompetencesNivPlusDecalageY + PositionY, _
                             Format(Int(TabComp(i).NivPlus * 100), CompetencesNivPlusFormat) & "%", _
                             SurfBack
            X = X + 1: Y = Y + 1
        Next i
        If TabComp.Count > CompetencesNombreLignes + 1 Then
            Afficher_BoutonsDefilement TabComp.Count - 1 - CompetencesNombreLignes, SurfBack
            BoutonsDefilement(0).Visible = True
        Else
            BoutonsDefilement(0).Visible = False
        End If
    Else
        BoutonsDefilement(0).Visible = False
    End If
    
    'On affiche les boutons de tri des compétences.
    For i = 0 To BoutonsCompetencesNombre - 1
        AffBoutonsCompetences.Afficher BoutonsCompetences(i).PositionX + PositionX, _
                                       BoutonsCompetences(i).PositionY + PositionY, _
                                       i, IIf(i = 3, -CDbl(CompetencesAfficherToutes), -CDbl(CompetencesTri = i)), SurfBack, True
    Next i
    
'    'Compétences Spéciales.
'    For i = 0 To Perso.Nombre_CompetencesSpeciales - 1
'        If X >= 0 And X <= CompetencesNombreLignes Then
'            'Colorie la compétence active.
'            If Perso.Action = 6 And i = 0 Or _
'               (Perso.Action = 301 Or Perso.Action = 302) And i = 1 Then
'                SurfBack.SetForeColor CompetenceCouleurActive
'            Else
'                SurfBack.SetForeColor 0
'            End If
'            .Afficher CompetencesPositionX + CompetencesIntervalX * X + PositionX, _
'                             CompetencesPositionY + CompetencesIntervalY * X + PositionY, _
'                             Parametres.Speciales_NomCompetences(i) & " :", _
'                             SurfBack
'            If Perso.Bonus_CompetenceSpeciales(i) Then
'                SurfBack.SetForeColor CompetenceBonusCouleur
'            Else
'                SurfBack.SetForeColor 0
'            End If
'            .Afficher CompetencesPositionX + CompetencesIntervalX * X + CompetencesScoreDecalageX + PositionX - CompetenceLargeurTexte * Len(Str(Perso.Definir_Niveau_CompetenceSpeciales(i))), _
'                      CompetencesPositionY + CompetencesIntervalY * X + CompetencesScoreDecalageY + PositionY, _
'                      Perso.Definir_Niveau_CompetenceSpeciales(i), _
'                      SurfBack
'            SurfBack.SetForeColor 0
'            .Afficher CompetencesPositionX + CompetencesIntervalX * X + CompetencesNivPlusDecalageX + PositionX, _
'                             CompetencesPositionY + CompetencesIntervalY * X + CompetencesNivPlusDecalageY + PositionY, _
'                             Format(Int(Perso.Progression_CompetenceSpeciale(i) * 100), CompetencesNivPlusFormat) & "%", _
'                             SurfBack
'        End If
'        X = X + 1
'    Next i
'
'    'Compétences de Ressources.
'    For i = 0 To Perso.Nombre_CompetencesRessources - 1
'        If X >= 0 And X <= CompetencesNombreLignes Then
'            'Colorie la compétence active.
'            If i = Perso.RessourceCourante Then
'                SurfBack.SetForeColor CompetenceCouleurActive
'            Else
'                SurfBack.SetForeColor 0
'            End If
'            .Afficher CompetencesPositionX + CompetencesIntervalX * X + PositionX, _
'                             CompetencesPositionY + CompetencesIntervalY * X + PositionY, _
'                             Parametres.Ressources_NomCompetence(i) & " :", _
'                             SurfBack
'
'            If Perso.Bonus_CompetenceRessources(i) Then
'                SurfBack.SetForeColor CompetenceBonusCouleur
'            Else
'                SurfBack.SetForeColor 0
'            End If
'            .Afficher CompetencesPositionX + CompetencesIntervalX * X + CompetencesScoreDecalageX + PositionX - CompetenceLargeurTexte * Len(Str(Perso.Definir_Niveau_CompetenceRessources(i) + Parametres.PersosMaxRessources)), _
'                      CompetencesPositionY + CompetencesIntervalY * X + CompetencesScoreDecalageY + PositionY, _
'                      Perso.Definir_Niveau_CompetenceRessources(i) + Parametres.PersosMaxRessources, _
'                      SurfBack
'            SurfBack.SetForeColor 0
'            .Afficher CompetencesPositionX + CompetencesIntervalX * X + CompetencesNivPlusDecalageX + PositionX, _
'                             CompetencesPositionY + CompetencesIntervalY * X + CompetencesNivPlusDecalageY + PositionY, _
'                             Format(Int(Perso.Progression_CompetenceRessource(i) * 100), CompetencesNivPlusFormat) & "%", _
'                             SurfBack
'        End If
'        X = X + 1
'    Next i
'
'    'Compétences de services.
'    For i = 0 To Perso.Nombre_CompetencesServices - 1
'        If X >= 0 And X <= CompetencesNombreLignes Then
'            'Colorie la compétence active.
'            If Perso.Action = 600 And Perso.ServiceCourant = i + 1 Then
'                SurfBack.SetForeColor CompetenceCouleurActive
'            Else
'                SurfBack.SetForeColor 0
'            End If
'            .Afficher CompetencesPositionX + CompetencesIntervalX * X + PositionX, _
'                             CompetencesPositionY + CompetencesIntervalY * X + PositionY, _
'                             Parametres.Service_NomCompetence(i + 1) & " :", _
'                             SurfBack
'            If Perso.Bonus_CompetenceServices(i) Then
'                SurfBack.SetForeColor CompetenceBonusCouleur
'            Else
'                SurfBack.SetForeColor 0
'            End If
'            .Afficher CompetencesPositionX + CompetencesIntervalX * X + CompetencesScoreDecalageX + PositionX - CompetenceLargeurTexte * Len(Str(Perso.Definir_Niveau_CompetenceServices(i))), _
'                             CompetencesPositionY + CompetencesIntervalY * X + CompetencesScoreDecalageY + PositionY, _
'                             Perso.Definir_Niveau_CompetenceServices(i), _
'                             SurfBack
'            SurfBack.SetForeColor 0
'            .Afficher CompetencesPositionX + CompetencesIntervalX * X + CompetencesNivPlusDecalageX + PositionX, _
'                             CompetencesPositionY + CompetencesIntervalY * X + CompetencesNivPlusDecalageY + PositionY, _
'                             Format(Int(Perso.Progression_CompetenceService(i) * 100), CompetencesNivPlusFormat) & "%", _
'                             SurfBack
'        End If
'        X = X + 1
'    Next i
'
'    'Affiche les compétences d'objets.
'    For i = 1 To Perso.Nombre_CompetencesObjets - 1
'        If X >= 0 And X <= CompetencesNombreLignes Then
'            k = Parametres.CompetenceObjet_NoListe(i)
'            If Perso.Action = 500 Then
'                If Parametres.Objet_Competence(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, Perso.IndiceObjet)) = k Then
'                    SurfBack.SetForeColor CompetenceCouleurActive
'                Else
'                    SurfBack.SetForeColor 0
'                End If
'            Else
'                SurfBack.SetForeColor 0
'            End If
'            If Perso.RessourceCourante >= 0 Then
'                For j = 0 To Perso.Nombre_Objets_Equipes - 1
'                    If Perso.Objet_Equipes_Actif(j) Then
'                        If Parametres.Objet_BonusCompetencesRessources(Perso.Objet_Equipes_Type(j), Perso.RessourceCourante) > 0 And _
'                           Perso.Objet_Equipes_Competence(j) = k Then
'                            SurfBack.SetForeColor CompetenceCouleurActive
'                            j = Perso.Nombre_Objets_Equipes
'                        End If
'                    End If
'                Next j
'            End If
'
'            .Afficher CompetencesPositionX + CompetencesIntervalX * X + PositionX, _
'                             CompetencesPositionY + CompetencesIntervalY * X + PositionY, _
'                             Parametres.CompetenceObjet_Nom(k) & " :", _
'                             SurfBack
'            If Perso.Bonus_CompetenceObjets(k) Then
'                SurfBack.SetForeColor CompetenceBonusCouleur
'            Else
'                SurfBack.SetForeColor 0
'            End If
'            .Afficher CompetencesPositionX + CompetencesIntervalX * X + CompetencesScoreDecalageX + PositionX - CompetenceLargeurTexte * Len(Str(Perso.Definir_Niveau_CompetenceObjets(k))), _
'                             CompetencesPositionY + CompetencesIntervalY * X + CompetencesScoreDecalageY + PositionY, _
'                             Perso.Definir_Niveau_CompetenceObjets(k), _
'                             SurfBack
'            SurfBack.SetForeColor 0
'            .Afficher CompetencesPositionX + CompetencesIntervalX * X + CompetencesNivPlusDecalageX + PositionX, _
'                             CompetencesPositionY + CompetencesIntervalY * X + CompetencesNivPlusDecalageY + PositionY, _
'                             Format(Int(Perso.Progression_CompetenceObjet(k) * 100), CompetencesNivPlusFormat) & "%", _
'                             SurfBack
'        End If
'        X = X + 1
'    Next i

    End With
    
    'Afficher_BoutonsDefilement X - 1 + DefilementDecalageX - CompetencesNombreLignes, SurfBack
End Sub

Public Sub Afficher_InfosInventaire(ByVal Titre As String, _
                                    ByRef Perso As ClsJeuPerso, _
                                    ByRef Maison As ClsJeuBatiment, _
                                    ByRef Parametres As ClsJeuParam, _
                                    ByRef AffObjets As ClsAffObjets, _
                                    ByRef Fenetre As ClsAffFenetre, _
                                    ByRef Message As ClsIntMessage, _
                                    ByRef SurfBack As DirectDrawSurface7)
    Dim i As Integer
    'Dim X As Integer
    'Dim Y As Integer
    Afficher_BoutonFermer SurfBack
    'Afficher_Titre Titre, Message, SurfBack
    SurfBack.SetForeColor 0
    Afficher_Argent_Perso Perso, Message, SurfBack
    Message.Afficher InventaireEquipementPositionX + PositionX, _
                     InventaireEquipementPositionY + PositionY, _
                     Titre, _
                     SurfBack
    'Objets équipés sur soi.
    For i = Perso.Nombre_Objets_Inventaire To UBound(BoutonsObjets())
        If Perso.Objet_Equipes_Type(i - Perso.Nombre_Objets_Inventaire) > -1 Then
            'Pour les armes à 2 mains, affiche une deuxième fois l'arme dans la main gauche.
            If i = Perso.Nombre_Objets_Inventaire And Perso.Objet_Equipes_2_Mains Then
                AffObjets.Afficher InventaireEquipementObjetX(Main2) + PositionX, _
                                   InventaireEquipementObjetY(Main2) + PositionY, _
                                   Perso.Objet_Equipes_Apparence(i - Perso.Nombre_Objets_Inventaire), _
                                   SurfBack, _
                                   True
            End If
            AffObjets.Afficher BoutonsObjets(i).PositionX + PositionX, _
                               BoutonsObjets(i).PositionY + PositionY, _
                               Perso.Objet_Equipes_Apparence(i - Perso.Nombre_Objets_Inventaire), _
                               SurfBack
        End If
    Next i
    If Bascule = 0 Or Not Perso.DansSaMaison Or Not Perso.Vivant Then
        'Objets de l'inventaire de réserve.
        For i = InventaireOngletsChoix * InventaireEquipementNombreCases To (InventaireOngletsChoix + 1) * InventaireEquipementNombreCases - 1 ' Perso.Nombre_Objets_Inventaire - 1
            If Perso.Objet_Inventaire_Apparence(i) > -1 Then
                With BoutonsObjets(i)
                AffObjets.Afficher .PositionX + PositionX, _
                                   .PositionY + PositionY, _
                                   Perso.Objet_Inventaire_Apparence(i), _
                                   SurfBack
                Interface(0).Afficher_Duree_Recharge .PositionX + PositionX, .PositionY + PositionY, Perso.Objet_Inventaire_Apparence(i), Perso, SurfBack
                End With
            End If
        Next i
        'Boutons d'onglets de l'inventaire.
        If NombreBoutonsInventaireOnglets > 1 Then
        For i = 0 To NombreBoutonsInventaireOnglets - 1
            AffBoutonAction.Afficher BoutonInventaireOnglets(i).PositionX + PositionX, _
                                     BoutonInventaireOnglets(i).PositionY + PositionY, _
                                     9, _
                                     -CDbl(InventaireOngletsChoix <> i), _
                                     SurfBack, True
        Next i
        End If
    Else
        'Objets de l'inventaire de la maison.
        For i = 0 To Maison.Definir_Nombre_Objet_Inventaire - 1
            If Maison.Definir_ObjetsInventaire(i) >= 0 Then
                With BoutonsObjets(i)
                AffObjets.Afficher .PositionX + PositionX, _
                                   .PositionY + PositionY, _
                                   Maison.Definir_ObjetsInventaire(i), _
                                   SurfBack
                End With
            End If
        Next i
    End If
    'Bouton de transfert vers la maison.
    If Perso.DansSaMaison And Perso.Vivant Then
        AffBoutonAction.Afficher BoutonInventaireMaison.PositionX + PositionX, _
                                 BoutonInventaireMaison.PositionY + PositionY, _
                                 7, _
                                 Bascule, _
                                 SurfBack, True
    End If
End Sub

Public Sub Afficher_InfosBestiaire(ByVal Titre As String, _
                                   ByRef Perso As ClsJeuPerso, _
                                   ByRef Parametres As ClsJeuParam, _
                                   ByRef Message As ClsIntMessage, _
                                   ByRef SurfBack As DirectDrawSurface7)
    Dim i As Integer
    Dim X As Integer
    Afficher_BoutonFermer SurfBack
    Afficher_Titre Titre, Message, SurfBack
    
    With Message
    
    For i = 0 To Perso.Nombre_de_lignes_du_bestiaire - 1
        If Perso.Definir_Bestiaire(Parametres.Bestiaire_Ordre_Alpha(i)) > 0 And _
           i >= DefilementDecalageX Then
            .Afficher BestiairePositionX + BestiaireIntervalX * X + PositionX, _
                      BestiairePositionY + BestiaireIntervalY * X + PositionY, _
                      Parametres.Race_Nom(Parametres.Bestiaire_Ordre_Alpha(i), False, Perso.Definir_Bestiaire(Parametres.Bestiaire_Ordre_Alpha(i)) >= 2) & " :", _
                      SurfBack
            .Afficher BestiairePositionX + BestiaireIntervalX * X + BestiaireScoreDecalageX + PositionX, _
                      BestiairePositionY + BestiaireIntervalY * X + BestiaireScoreDecalageY + PositionY, _
                      Perso.Definir_Bestiaire(Parametres.Bestiaire_Ordre_Alpha(i)), _
                      SurfBack
'            'Affiche une ligne noire de séparation.
'            Afficher_Separateur_Ressources BestiairePositionX + BestiaireIntervalX * X + PositionX, _
'                                           BestiairePositionY + BestiaireIntervalY * X + PositionY, SurfBack
'            'Affiche le monstre.
'            AffPerso.Afficher BestiairePositionX + BestiaireIntervalX * X + BestiaireScoreDecalageX + PositionX, _
'                              BestiairePositionY + BestiaireIntervalY * X + BestiaireScoreDecalageY + PositionY, _
'                              1,0,
            X = X + 1
            If X > BestiaireNombreLignes Then
                i = Perso.Nombre_de_lignes_du_bestiaire
            End If
        End If
    Next i
    
    End With
    'Afficher_BoutonsDefilement i - 1 + DefilementDecalageX - BestiaireNombreLignes, SurfBack
    Afficher_BoutonsDefilement Perso.Nombre_de_lignes_du_bestiaire - 1 - BestiaireNombreLignes, SurfBack
End Sub

Public Sub Afficher_Carte(ByVal Titre As String, _
                          ByRef Perso As ClsJeuPerso, _
                          ByRef Message As ClsIntMessage, _
                          ByRef SurfBack As DirectDrawSurface7, _
                          ByVal X As Long, ByVal Y As Long, _
                          ByVal Largeur As Integer, ByVal Hauteur As Integer)
    Dim Temp1, Temp2 As Long
    'Dim CarteDecalageX As Integer
    'Dim CarteDecalageY As Integer
    Afficher_BoutonFermer SurfBack
    Afficher_Titre Titre, Message, SurfBack
    Dim i As Long
    If Monde.Largeur > Monde.Hauteur Then
        CarteLargeur = CarteDimension
        CarteHauteur = CarteDimension * Monde.Hauteur / Monde.Largeur
        CarteDecalageX = PositionX
        CarteDecalageY = PositionY + (CarteDimension - (CarteDimension * Monde.Hauteur / Monde.Largeur)) / 2
    Else
        CarteLargeur = CarteDimension * Monde.Largeur / Monde.Hauteur
        CarteHauteur = CarteDimension
        CarteDecalageX = PositionX + (CarteDimension - (CarteDimension * Monde.Largeur / Monde.Hauteur)) / 2
        CarteDecalageY = PositionY
    End If
    
    With SurfBack
    Temp1 = .GetForeColor
    Temp2 = .GetFillColor
    .SetForeColor CarteCouleurFond
    .SetFillColor CarteCouleurFond
    .SetFillStyle 6
    'Affiche la carte
    If Cartes.Chargee Then
        If CarteChargee Then  'Charge la carte, si ce n'est dejà fait.
            If CarteExiste Then 'Affiche la carte que si le fichier existe.
                AffCarte.Afficher CartePositionXDepart + CarteDecalageX, CartePositionYDepart + CarteDecalageY, SurfBack, False
            Else
                .DrawBox CartePositionXDepart + CarteDecalageX, CartePositionYDepart + CarteDecalageY, CartePositionXDepart + CarteLargeur + CarteDecalageX, CartePositionYDepart + CarteHauteur + CarteDecalageY
            End If
        Else
            CarteExiste = Cartes.Definir_ImageCarte_Existe
            If CarteExiste Then
                AffCarte.Charger dd, Cartes.Definir_ImageCarte_Chemin, True
            End If
            CarteChargee = True
        End If
    Else
        .DrawBox CartePositionXDepart + CarteDecalageX, CartePositionYDepart + CarteDecalageY, CartePositionXDepart + CarteLargeur + CarteDecalageX, CartePositionYDepart + CarteHauteur + CarteDecalageY
    End If
    .SetFillStyle 0
    
    For i = 0 To UBound(Chateaux())
        If Chateaux(i).Visible Then
            If Fiefs(i).NumeroEquipe = Perso.NumeroEquipe Then
                If i = Perso.NumeroFief Then
                    .SetForeColor CarteCouleurChateau
                    .SetFillColor CarteCouleurChateau
                Else
                    .SetForeColor CarteCouleurAlliesChateaux
                    .SetFillColor CarteCouleurAlliesChateaux
                End If
            Else
                .SetForeColor CarteCouleurEnnemisChateaux
                .SetFillColor CarteCouleurEnnemisChateaux
            End If
            .DrawBox CartePositionXDepart + (Chateaux(i).PositionX + Chateaux(i).Largeur / 4) / (Monde.Largeur * Monde.LargeurCase) * CarteLargeur + CarteDecalageX, _
                     CartePositionYDepart + (Chateaux(i).PositionY + Chateaux(i).Hauteur / 4) / (Monde.Hauteur * Monde.HauteurCase) * CarteHauteur + CarteDecalageY, _
                     CartePositionXDepart + (Chateaux(i).PositionX + Chateaux(i).Largeur / 4) / (Monde.Largeur * Monde.LargeurCase) * CarteLargeur + CarteTailleChateau + CarteDecalageX, _
                     CartePositionYDepart + (Chateaux(i).PositionY + Chateaux(i).Hauteur / 4) / (Monde.Hauteur * Monde.HauteurCase) * CarteHauteur + CarteTailleChateau + CarteDecalageY
        End If
    Next i
    
    For i = UBound(Persos()) To 0 Step -1
        If Maisons(i).Visible Then
            If Persos(i).NumeroEquipe = Perso.NumeroEquipe Then
                If Persos(i).Numero = Perso.Numero Then
                    .SetForeColor CarteCouleurMaison
                    .SetFillColor CarteCouleurMaison
                Else
                    .SetForeColor CarteCouleurAlliesMaisons
                    .SetFillColor CarteCouleurAlliesMaisons
                End If
            Else
                .SetForeColor CarteCouleurEnnemisMaisons
                .SetFillColor CarteCouleurEnnemisMaisons
            End If
            .DrawBox CartePositionXDepart + 1 + (Maisons(i).PositionX) / (Monde.Largeur * Monde.LargeurCase) * CarteLargeur + CarteDecalageX, _
                     CartePositionYDepart + 2 + (Maisons(i).PositionY) / (Monde.Hauteur * Monde.HauteurCase) * CarteHauteur + CarteDecalageY, _
                     CartePositionXDepart + 1 + (Maisons(i).PositionX) / (Monde.Largeur * Monde.LargeurCase) * CarteLargeur + CarteTailleMaison + CarteDecalageX, _
                     CartePositionYDepart + 2 + (Maisons(i).PositionY) / (Monde.Hauteur * Monde.HauteurCase) * CarteHauteur + CarteTailleMaison + CarteDecalageY
        End If
    Next i
    For i = UBound(Persos()) To 0 Step -1
        If Persos(i).Attaquable Then
            If Persos(i).NumeroEquipe = Perso.NumeroEquipe Then
                If Persos(i).Numero = Perso.Numero Or (Persos(i).Invocation And Persos(i).NumeroChef = Perso.Numero) Then
                    .SetForeColor CarteCouleurPerso
                    .SetFillColor CarteCouleurPerso
                Else
                    .SetForeColor CarteCouleurAlliesPersos
                    .SetFillColor CarteCouleurAlliesPersos
                End If
            Else
                If Persos(i).EnCombat Then
                    If Rnd < 0.5 And Persos(i).Definir_Attaque < Persos(i).Definir_MaxAttaque Then
                        If Rnd < 0.5 Then
                            .SetForeColor CarteCouleurAttaquePersos
                            .SetFillColor CarteCouleurAttaquePersos
                        Else
                            .SetForeColor CarteCouleurEnnemisPersos
                            .SetFillColor CarteCouleurEnnemisPersos
                        End If
                    Else
                        .SetForeColor CarteCouleurEnnemisPersos
                        .SetFillColor CarteCouleurEnnemisPersos
                    End If
                Else
                    .SetForeColor CarteCouleurEnnemisPersos
                    .SetFillColor CarteCouleurEnnemisPersos
                End If
            End If
            .DrawCircle CartePositionXDepart + Persos(i).PositionX / (Monde.Largeur * Monde.LargeurCase) * CarteLargeur + CarteDecalageX + 2, _
                        CartePositionYDepart + Persos(i).PositionY / (Monde.Hauteur * Monde.HauteurCase) * CarteHauteur + CarteDecalageY + 2, _
                        CarteTaillePerso
        End If
    Next i
    'Trace deux droites indiquant l'emplacement du joueur.
    .SetForeColor CarteCouleurPerso
    'Trace une ligne horizontale.
    .DrawLine CartePositionXDepart + CarteDecalageX, _
              CartePositionYDepart + Perso.PositionY / (Monde.Hauteur * Monde.HauteurCase) * CarteHauteur + CarteDecalageY + 2, _
              CartePositionXDepart + CarteLargeur + CarteDecalageX, _
              CartePositionYDepart + Perso.PositionY / (Monde.Hauteur * Monde.HauteurCase) * CarteHauteur + CarteDecalageY + 2
    'Trace une ligne verticale.
    .DrawLine CartePositionXDepart + Perso.PositionX / (Monde.Largeur * Monde.LargeurCase) * CarteLargeur + CarteDecalageX + 2, _
              CartePositionYDepart + CarteDecalageY, _
              CartePositionXDepart + Perso.PositionX / (Monde.Largeur * Monde.LargeurCase) * CarteLargeur + CarteDecalageX + 2, _
              CartePositionYDepart + CarteHauteur + CarteDecalageY
    'Trace le rectangle de l'écran.
    .SetFillStyle 1
    .DrawBox CartePositionXDepart + X / (Monde.Largeur * Monde.LargeurCase) * CarteLargeur + CarteDecalageX, _
             CartePositionYDepart + Y / (Monde.Hauteur * Monde.HauteurCase) * CarteHauteur + CarteDecalageY, _
             CartePositionXDepart + (X + Largeur) / (Monde.Largeur * Monde.LargeurCase) * CarteLargeur + CarteDecalageX, _
             CartePositionYDepart + (Y + Hauteur) / (Monde.Hauteur * Monde.HauteurCase) * CarteHauteur + CarteDecalageY
    
    End With
End Sub

Public Sub Afficher_InfosObjectifs(ByVal Titre As String, _
                                   ByVal Objectifs As Collection, _
                                   ByRef Fenetre As ClsAffFenetre, _
                                   ByRef Message As ClsIntMessage, _
                                   ByRef SurfBack As DirectDrawSurface7)
    Dim i As Integer
    Dim X As Integer
    Afficher_BoutonFermer SurfBack
    Afficher_Titre Titre, Message, SurfBack
    
    With Message
    
    If Objectifs.Count > 0 Then
    For i = 1 To Objectifs.Count
        .Afficher ObjectifsPositionX + ObjectifsIntervalX * i + PositionX, _
                  ObjectifsPositionY + ObjectifsIntervalY * i + PositionY, _
                  Objectifs(i), SurfBack
        If i < Objectifs.Count Then
        SurfBack.DrawLine ObjectifsPositionX + ObjectifsIntervalX * i + PositionX + ObjectifsFinisDecalageX, _
                          ObjectifsPositionY + ObjectifsIntervalY * i + PositionY + ObjectifsFinisDecalageY, _
                          ObjectifsPositionX + ObjectifsIntervalX * i + PositionX + ObjectifsFinisDecalageX + ObjectifsFinisLongueur, _
                          ObjectifsPositionY + ObjectifsIntervalY * i + PositionY + ObjectifsFinisDecalageY
        End If
    Next i
    End If
    
    End With
End Sub

Private Sub Afficher_Titre(ByVal Titre As String, _
                           ByVal Message As ClsIntMessage, _
                           ByRef SurfBack As DirectDrawSurface7)
    SurfBack.SetForeColor 0
    Message.Afficher TitrePositionX + PositionX, _
                     TitrePositionY + PositionY, _
                     Titre, _
                     SurfBack
End Sub

Private Sub Afficher_BoutonFermer(ByRef SurfBack As DirectDrawSurface7)
    AffBoutonAction.Afficher BoutonFermer.PositionX + PositionX, _
                             BoutonFermer.PositionY + PositionY, _
                             6, _
                             0, _
                             SurfBack
    AffBoutonAction.Afficher BoutonDeplacer.PositionX + PositionX, _
                             BoutonDeplacer.PositionY + PositionY, _
                             8, _
                             0, _
                             SurfBack
End Sub

Private Sub Afficher_Argent_Perso(ByVal Perso As ClsJeuPerso, _
                                  ByRef Message As ClsIntMessage, _
                                  ByRef SurfBack As DirectDrawSurface7)
    Dim i As Long
    Message.Afficher ArgentPersoPositionX + PositionX, _
                     ArgentPersoPositionY + PositionY, _
                     UCase(Left(Parametres.EtiquetteArgent, 1)) & Right(Parametres.EtiquetteArgent, Len(Parametres.EtiquetteArgent) - 1) & " : " & Int(Perso.Definir_Argent), _
                     SurfBack
    'Affichage des boutons de transfert.
    
    For i = 0 To NombreBoutonsTransfert - 1
        If Perso.DansSaMaison Then
            BoutonsTransfertArgentPerso(i).PositionX = ArgentPersoPositionX + BoutonTransfertDecalageX + BoutonTransfertIntervalX * i
            BoutonsTransfertArgentPerso(i).PositionY = ArgentPersoPositionY + BoutonTransfertDecalageY + BoutonTransfertIntervalY * i
            'BoutonsTransfertArgentPerso(i).Visible = True
            AffBoutonsTransfert.Afficher BoutonsTransfertArgentPerso(i).PositionX + PositionX, _
                                         BoutonsTransfertArgentPerso(i).PositionY + PositionY, _
                                         i + NombreBoutonsTransfert, 0, SurfBack
        'Else
            'BoutonsTransfertArgentPerso(i).Visible = True
        End If
    Next i
End Sub

Private Sub Afficher_EtatBatiment(ByRef Perso As ClsJeuPerso, _
                                  ByRef Maison As ClsJeuBatiment, _
                                  ByRef Parametres As ClsJeuParam, _
                                  ByRef Message As ClsIntMessage, _
                                  ByRef SurfBack As DirectDrawSurface7, _
                                  Optional ByVal Chateau As Boolean)
    Dim i As Long
    If Chateau Then
        Message.Afficher EtatBatimentPositionX + PositionX, _
                         EtatBatimentPositionY + PositionY, _
                         Parametres.EtiquetteVie & " : " & Int(Maison.Vie) & _
                         IIf(Int(Maison.Vie) <> Int(Maison.Definir_Vie), " (" & Int(Maison.Definir_Vie) & ")", "") & _
                         " / " & Maison.MaxVie, _
                         SurfBack
        'Affiche le bouton "Empereur".
        AffBoutonAction.Afficher BoutonEmpereurPositionX + PositionX, _
                                 BoutonEmpereurPositionY + PositionY, _
                                 10, _
                                 IIf(Perso.PeutDevenir_Empereur(Fiefs(Perso.NumeroFief), Parametres), 0, 1), SurfBack, True
        'Affiche les boutons de restauration du chateau.
        For i = 0 To NombreBoutonsRestauration - 1
            BoutonsRestaurer(i).PositionX = BoutonsRestaurationPositionX + i * BoutonsRestaurationIntervalX
            BoutonsRestaurer(i).PositionY = BoutonsRestaurationPositionY + i * BoutonsRestaurationIntervalY
            AffBoutonAction.Afficher BoutonsRestaurer(i).PositionX + PositionX, _
                                     BoutonsRestaurer(i).PositionY + PositionY, _
                                     2 + i, _
                                     IIf(Perso.Vivant, 0, 1), _
                                     SurfBack
        Next i
    Else
        Message.Afficher EtatBatimentPositionX + PositionX, _
                         EtatBatimentPositionY + PositionY, _
                         Parametres.EtiquetteVie & " : " & Int(Maison.Definir_Vie) & _
                         " / " & Maison.MaxVie, _
                         SurfBack
    End If
End Sub

Private Sub Afficher_Argent_Batiment(ByRef Perso As ClsJeuPerso, _
                                     ByRef Maison As ClsJeuBatiment, _
                                     ByRef Message As ClsIntMessage, _
                                     ByRef SurfBack As DirectDrawSurface7)
    Dim i As Long
    Message.Afficher ArgentMaisonPositionX + PositionX, _
                     ArgentMaisonPositionY + PositionY, _
                     UCase(Left(Parametres.EtiquetteArgent, 1)) & Right(Parametres.EtiquetteArgent, Len(Parametres.EtiquetteArgent) - 1) & " : " & Int(Maison.Argent), _
                     SurfBack
    'Affichage des boutons de transfert.
    For i = 0 To NombreBoutonsTransfert - 1
        If Perso.DansSaMaison And Maison.Construit Then
            BoutonsTransfertArgentMaison(i).PositionX = ArgentMaisonPositionX + BoutonTransfertDecalageX + BoutonTransfertIntervalX * i
            BoutonsTransfertArgentMaison(i).PositionY = ArgentMaisonPositionY + BoutonTransfertDecalageY + BoutonTransfertIntervalY * i
            BoutonsTransfertArgentMaison(i).Visible = True
            AffBoutonsTransfert.Afficher BoutonsTransfertArgentMaison(i).PositionX + PositionX, _
                                         BoutonsTransfertArgentMaison(i).PositionY + PositionY, _
                                         i, 0, SurfBack
        Else
            BoutonsTransfertArgentMaison(i).Visible = False
        End If
    Next i
End Sub

Private Sub Afficher_StockBatiment(ByVal Interface As ClsInterface, _
                                   ByRef Perso As ClsJeuPerso, _
                                   ByRef Maison As ClsJeuBatiment, _
                                   ByRef Parametres As ClsJeuParam, _
                                   ByRef Message As ClsIntMessage, _
                                   ByRef SurfBack As DirectDrawSurface7)
    Dim i As Long
    Dim Temp As Integer
    'If Maison.MaxStock > 0 Then
        If Maison.Fabrique Then
            'Fabrique d'objets.
            If Int(Maison.Definir_Stock) < 1 Then 'Le stock est vide.
                Message.Afficher StockBatimentPositionX + PositionX, _
                                 StockBatimentPositionY + PositionY, _
                                 EtiquetteObjets & _
                                 EtiquetteStockVide, _
                                 SurfBack
            Else
                Message.Afficher StockBatimentPositionX + PositionX, _
                                 StockBatimentPositionY + PositionY, _
                                 EtiquetteObjets & Int(Maison.Definir_Stock) & _
                                 " / " & Maison.MaxStock, _
                                 SurfBack
            End If
        ElseIf Maison.Magasin Then
            'Batiment de ressources.
            If Int(Maison.Definir_Stock) < 1 Then 'Le stock est vide.
                Message.Afficher StockBatimentPositionX + PositionX, _
                                 StockBatimentPositionY + PositionY, _
                                 Parametres.Ressources_Nom(Parametres.Batiment_RessourceFabrique(Maison.TypeBatiment)) & " : " & _
                                 EtiquetteStockVide, _
                                 SurfBack
            Else
                Message.Afficher StockBatimentPositionX + PositionX, _
                                 StockBatimentPositionY + PositionY, _
                                 Parametres.Ressources_Nom(Parametres.Batiment_RessourceFabrique(Maison.TypeBatiment)) & " : " & Int(Maison.Definir_Stock) & _
                                 " / " & Maison.MaxStock, _
                                 SurfBack
            End If
            'Affichage des boutons de transfert.
            For i = 0 To NombreBoutonsTransfert - 1
                If Maison.Construit Then
                    BoutonsTransfertStockMaison(i).PositionX = StockBatimentPositionX + BoutonTransfertDecalageX + BoutonTransfertIntervalX * i
                    BoutonsTransfertStockMaison(i).PositionY = StockBatimentPositionY + BoutonTransfertDecalageY + BoutonTransfertIntervalY * i
                    BoutonsTransfertStockMaison(i).Visible = True
                    If Perso.DansSaMaison Then
                        AffBoutonsTransfert.Afficher BoutonsTransfertStockMaison(i).PositionX + PositionX, _
                                                     BoutonsTransfertStockMaison(i).PositionY + PositionY, _
                                                     i, 0, SurfBack
                    Else
                        AffBoutonsCommerce.Afficher BoutonsTransfertStockMaison(i).PositionX + PositionX, _
                                                    BoutonsTransfertStockMaison(i).PositionY + PositionY, _
                                                    i, 0, SurfBack
                        Temp = Perso.Prix_Marchandee(Parametres.Ressources_PrixVente((Parametres.Batiment_RessourceFabrique(Maison.TypeBatiment))) * Parametres.CoefficientAchat, Parametres)
                        Select Case i
                        Case 0:
                            BoutonsCommerce(Parametres.Batiment_RessourceFabrique(Maison.TypeBatiment) * NombreBoutonsCommerce + i).InfoBulle = BoutonAchatStockMaison1Etiquette
                        Case 1:
                            BoutonsCommerce(Parametres.Batiment_RessourceFabrique(Maison.TypeBatiment) * NombreBoutonsCommerce + i).InfoBulle = BoutonAchatStockMaison2Etiquette
                            If Maison.Definir_Stock < 5 Then
                                Temp = Temp * Maison.Definir_Stock
                            Else
                                Temp = Temp * 5
                            End If
                        Case 2:
                            BoutonsCommerce(Parametres.Batiment_RessourceFabrique(Maison.TypeBatiment) * NombreBoutonsCommerce + i).InfoBulle = BoutonAchatStockMaison3Etiquette
                            Temp = Temp * Maison.Definir_Stock
                        End Select
                        BoutonsCommerce(Parametres.Batiment_RessourceFabrique(Maison.TypeBatiment) * NombreBoutonsCommerce + i).InfoBulle = BoutonsCommerce(Parametres.Batiment_RessourceFabrique(Maison.TypeBatiment) * NombreBoutonsCommerce + i).InfoBulle & "#" & Parametres.EtiquettePrix & " : " & Temp
                    End If
                Else
                    BoutonsTransfertStockMaison(i).Visible = False
                End If
            Next i
        ElseIf Maison.Service Then
            If Maison.MaxStock > 0 Then
            If Int(Maison.Definir_Stock) < 1 Then 'Le stock est vide.
                Message.Afficher StockBatimentPositionX + PositionX, _
                                 StockBatimentPositionY + PositionY, _
                                 Parametres.Service_Nom(Maison.Service) & " : " & _
                                 EtiquetteIndisponible, _
                                 SurfBack
            Else
                Message.Afficher StockBatimentPositionX + PositionX, _
                                 StockBatimentPositionY + PositionY, _
                                 Parametres.Service_Nom(Maison.Service) & " : " & Int(Maison.Definir_Stock) & _
                                 " / " & Maison.MaxStock, _
                                 SurfBack
            End If
            If Perso.DansSaMaison Then
                BoutonService.InfoBulle = ServiceBoutonInfoBullePrefixe & Parametres.EtiquettePrix & " : " & Parametres.Service_Prix(Maison.Service) & ServiceBoutonInfoBulleSuffixe
            Else
                BoutonService.InfoBulle = ServiceBoutonInfoBullePrefixe & Parametres.EtiquettePrix & " : " & Perso.Prix_Marchandee(Parametres.Service_Prix(Maison.Service), Parametres) & ServiceBoutonInfoBulleSuffixe
            End If
            AffBoutonAction.Afficher BoutonService.PositionX + PositionX, _
                                     BoutonService.PositionY + PositionY, _
                                     5, _
                                     IIf(Maison.Service_Utilisable(Perso, Parametres), 0, 1), SurfBack
            If Interface.AutoBatimentService Then  'Bouton en fonction automatique.
                SurfBack.SetForeColor Interface.BarresStatutContourCouleur2
                Interface.Afficher_Contour_Barre_Statut BoutonService.PositionX + PositionX, BoutonService.PositionY + PositionY, AffBoutonAction.Largeur, AffBoutonAction.Hauteur, 1, SurfBack
                SurfBack.SetForeColor 0
            End If
            End If
            If Maison.MaxStock > 0 Or Perso.DansSaMaison Then
            Message.Afficher ServiceDescriptionPositionX + PositionX, _
                         ServiceDescriptionPositionY + PositionY, _
                         Parametres.Service_Description(Maison.Service), _
                         SurfBack
            End If
        End If
    'End If
End Sub

Private Sub Afficher_Objets_Batiment(ByVal FileFabricationObjets As Collection, _
                                     ByRef Perso As ClsJeuPerso, _
                                     ByRef Maison As ClsJeuBatiment, _
                                     ByRef Parametres As ClsJeuParam, _
                                     ByRef Message As ClsIntMessage, _
                                     ByRef AffObjets As ClsAffObjets, _
                                     ByRef SurfBack As DirectDrawSurface7)
    'Affiche les objets dans le batiment et les boutons pour les fabriquer.
    Dim i As Integer
    Dim j As Integer
    Dim X As Integer
    Dim NombreObjetsFabriques As Integer
    Dim ListeAttente() As Integer
    NombreObjetsFabriques = Parametres.Batiment_NombreObjetsFabriques(Maison.TypeBatiment)
    
    If Parametres.Batiment_Fabrique(Maison.TypeBatiment) Then
        'On affiche les boutons de fabrication seulement si la maison est une fabrique.
        If UBound(BoutonsFabriquer()) <> NombreObjetsFabriques - 1 Then
            ReDim BoutonsFabriquer(NombreObjetsFabriques - 1)
            ReDim BoutonsObjets(NombreObjetsFabriques - 1)
            For i = 0 To NombreObjetsFabriques - 1
                Set BoutonsFabriquer(i) = New ClsIntBouton
                Set BoutonsObjets(i) = New ClsIntBouton
            Next i
        End If
        If DefilementObjetDecalageX > NombreObjetsFabriques - NombreLignesObjets Then
            DefilementObjetDecalageX = NombreObjetsFabriques - NombreLignesObjets
        End If
        If DefilementObjetDecalageX < 0 Then
            DefilementObjetDecalageX = 0
        End If
        
        'Calcul les chiffres de la file d'attente de la maison.
        If Perso.DansSaMaison Then
            ReDim ListeAttente(NombreObjetsFabriques - 1)
            If FileFabricationObjets.Count > 0 Then
                For i = 1 To FileFabricationObjets.Count
                    ListeAttente(FileFabricationObjets(i)) = ListeAttente(FileFabricationObjets(i)) + 1
                Next i
            End If
        End If
        
        For i = DefilementObjetDecalageX To NombreObjetsFabriques - 1
            'Nom de l'objet.
            Message.Afficher FabricationPositionX + FabricationIntervalX * X + PositionX, _
                             FabricationPositionY + FabricationIntervalY * X + PositionY, _
                             Parametres.Objet_NomPluriel(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, i)), _
                             SurfBack
            'Trace une ligne entre chaque objet.
            Afficher_Separateur_Ressources FabricationPositionX + PositionX + FabricationIntervalX * X, _
                                           FabricationPositionY + PositionY + FabricationIntervalY * X, _
                                           SurfBack
            'Icone de l'objet.
            BoutonsObjets(i).PositionX = FabricationPositionX + FabricationIntervalX * X + FabricationIconeDecalageX
            BoutonsObjets(i).PositionY = FabricationPositionY + FabricationIntervalY * X + FabricationIconeDecalageY
            AffObjets.Afficher2 BoutonsObjets(i).PositionX + PositionX, _
                                BoutonsObjets(i).PositionY + PositionY, _
                                Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, i), _
                                SurfBack, _
                                , , FabricationIconeZoom 'Int(Maison.Stock_Objets(i) = 0)
            'BoutonsObjets(i).InfoBulle = Definir_InfoBulle_Objet(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, i), Perso, Parametres)
            'Quantité de l'objet.
            Message.Afficher FabricationPositionX + FabricationIntervalX * X + FabricationQuantiteDecalageX + PositionX, _
                             FabricationPositionY + FabricationIntervalY * X + FabricationQuantiteDecalageY + PositionY, _
                             "x" & Int(Maison.Stock_Objets(i)), _
                             SurfBack
            'Bouton fabriquer.
            BoutonsFabriquer(i).PositionX = FabricationPositionX + FabricationIntervalX * X + FabricationBoutonDecalageX
            BoutonsFabriquer(i).PositionY = FabricationPositionY + FabricationIntervalY * X + FabricationBoutonDecalageY
                
            If Perso.DansSaMaison Then
                'BoutonsFabriquer(i).Visible = True
                If Parametres.Objet_Epoque(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, i)) - 1 > Fiefs(Perso.NumeroFief).Epoque Then
                    AffBoutonAction.Afficher BoutonsFabriquer(i).PositionX + PositionX, _
                                             BoutonsFabriquer(i).PositionY + PositionY, _
                                             1, 2, SurfBack, True
                    BoutonsFabriquer(i).Enabled = False
                ElseIf Perso.PeutConstruire_Objet(i, Maison, Parametres) Then
                    AffBoutonAction.Afficher BoutonsFabriquer(i).PositionX + PositionX, _
                                             BoutonsFabriquer(i).PositionY + PositionY, _
                                             1, 0, SurfBack, True
                    BoutonsFabriquer(i).Enabled = True
                Else
                    AffBoutonAction.Afficher BoutonsFabriquer(i).PositionX + PositionX, _
                                             BoutonsFabriquer(i).PositionY + PositionY, _
                                             1, 1, SurfBack, True
                    BoutonsFabriquer(i).Enabled = False
                End If
                BoutonsFabriquer(i).InfoBulle = BoutonFabriquerEtiquette & _
                                                Parametres.Objet_Nom(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, i))
                'Affiche la quantité d'objets à fabriquer en file d'attente.
                If ListeAttente(i) > 0 Then
                    SurfBack.SetForeColor &HFFFFFF
                    Message.Afficher FabricationPositionX + FabricationIntervalX * X + FabricationFileAttenteDecalageX - 8 + PositionX, _
                                     FabricationPositionY + FabricationIntervalY * X + FabricationFileAttenteDecalageY - 8 + PositionY, _
                                     "+" & ListeAttente(i), _
                                     SurfBack
                    SurfBack.SetForeColor 0
                End If
                'Pourcentage de fabrication.
                If Perso.IndiceObjet = i And _
                   Perso.EtatFabricationObjet > 0 Then
                    Message.Afficher FabricationPositionX + FabricationIntervalX * X + FabricationProgressionDecalageX + PositionX, _
                                     FabricationPositionY + FabricationIntervalY * X + FabricationProgressionDecalageY + PositionY, _
                                     Int(Perso.EtatFabricationObjet / Parametres.Objet_DureeFabrication(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, Perso.IndiceObjet)) * 100) & _
                                     "%", _
                                     SurfBack
                End If
            Else
                If Parametres.Objet_Epoque(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, i)) - 1 > Fiefs(Perso.NumeroFief).Epoque Then
                    AffBoutonAction.Afficher BoutonsFabriquer(i).PositionX + PositionX, _
                                             BoutonsFabriquer(i).PositionY + PositionY, _
                                             1, 2, SurfBack, True
                End If
            End If
            
            'Affiche l'époque nécessaire.
            If Parametres.Objet_Epoque(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, i)) - 1 > Fiefs(Perso.NumeroFief).Epoque Then
                BoutonsFabriquer(i).InfoBulle = BoutonsFabriquer(i).InfoBulle & Chr(13) & "&(" & _
                                                Parametres.Epoque_Nom(Parametres.Objet_Epoque(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, i)) - 1) & ")"
            End If
            For j = 0 To Parametres.NombreRessources - 1
                If Parametres.Objet_PrixRessources(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, i), j) > 0 Then
                    BoutonsFabriquer(i).InfoBulle = BoutonsFabriquer(i).InfoBulle & Chr(13)
                    If Parametres.Objet_PrixRessources(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, i), j) > _
                       Perso.Definir_ressources(j) Then
                        BoutonsFabriquer(i).InfoBulle = BoutonsFabriquer(i).InfoBulle & "&"
                    End If
                    
                    BoutonsFabriquer(i).InfoBulle = BoutonsFabriquer(i).InfoBulle & "(" & _
                                                    IIf(Parametres.Objet_PrixRessources(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, i), j) > Perso.Definir_ressources(j), _
                                                    Int(Perso.Definir_ressources(j)) & " / ", _
                                                    "") & _
                                                    Parametres.Objet_PrixRessources(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, i), j) & _
                                                    " " & _
                                                    LCase(Parametres.Ressources_Nom(j)) & ")"
                End If
            Next j
            
            'Evite que des boutons sortent de la fenêtre.
            If X >= NombreLignesObjets - 1 Then
                i = NombreObjetsFabriques
            Else
                X = X + 1
            End If
        Next i
        If NombreObjetsFabriques > NombreLignesObjets Then
            'Afficher_BoutonsDefilementObjet NombreObjetsFabriques - NombreLignesObjets, SurfBack
            Afficher_BoutonsDefilement NombreObjetsFabriques - NombreLignesObjets, SurfBack, True
        Else
            BoutonsDefilement(0).Visible = False
            BoutonsDefilement(1).Visible = False
            BoutonsDefilement(2).Visible = False
        End If
    End If
End Sub

Private Function Definir_InfoBulle_Objet(ByRef NumeroObjet As Integer, _
                                         ByRef Perso As ClsJeuPerso, _
                                         ByRef Parametres As ClsJeuParam, _
                                         ByVal Inventaire As Boolean, _
                                         ByVal Equipe As Boolean, _
                                         Optional ByVal Duree As Long = 0, _
                                         Optional ByVal NumeroEffet As Integer = 0, _
                                         Optional ByVal Nombre As Long = 0) As String
    Dim j, k As Long
    Dim Competence As Integer
    Dim NiveauCompetence As Integer
    Dim Effet As Boolean
    If NumeroObjet >= 0 Then
        Competence = Parametres.Objet_Competence(NumeroObjet)
        Effet = Duree <> 0
        'Affichage des informations de l'objet.
        If Effet Then
            Definir_InfoBulle_Objet = Parametres.Objet_NomEffet(NumeroObjet, Perso.Feminin)
            If Perso.Definir_EffetTemp_Proprietaire(NumeroEffet) <> Perso.Numero Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & " (" & Persos(Perso.Definir_EffetTemp_Proprietaire(NumeroEffet)).Nom & ")"
            End If
            NiveauCompetence = Persos(Perso.Definir_EffetTemp_Proprietaire(NumeroEffet)).Definir_Niveau_CompetenceObjets(Competence)
        Else
            If Nombre >= 2 Then
                'On a plusieurs fois le même objet sur cette case.
                Definir_InfoBulle_Objet = Nombre & " " & Parametres.Objet_NomPluriel(NumeroObjet)
            Else
                Definir_InfoBulle_Objet = Parametres.Objet_Nom(NumeroObjet)
            End If
            NiveauCompetence = Perso.Definir_Niveau_CompetenceObjets(Competence)
        End If
        
        If Competence > 0 And Not Effet Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & " (" & _
                                      Parametres.CompetenceObjet_Nom(Competence) & " : " & _
                                      Perso.Definir_Niveau_CompetenceObjets(Competence) & ")"
        End If
        
        If Not Effet Then
            'Affiche le temps nécessaire à la recharge.
            If Parametres.Objet_DureeRecharge(NumeroObjet) > 0 Then
                If Perso.EffetRecharge_TempRestant(NumeroObjet) > 0 Then
                    Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#&" & _
                                          Parametres.EtiquetteDureeRecharge & " : " & _
                                          Perso.EffetRecharge_TempRestant(NumeroObjet) & " " & Parametres.EtiquetteUniteDuree
                Else
                    Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & Parametres.EtiquetteDureeRecharge & " : " & Parametres.Objet_DureeRecharge(NumeroObjet) & " " & Parametres.EtiquetteUniteDuree
                End If
                If Int(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_RechargesMultiples(NumeroObjet))) > 1 Then
                    Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
                                              Int(Perso.Objet_Equipe_BonusSup(Perso.Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(NumeroObjet)), Parametres.Objet_RechargesMultiples(NumeroObjet)) - Perso.EffetRecharge_TempRestant(NumeroObjet) / Parametres.Objet_DureeRecharge(NumeroObjet)) & "/" & _
                                              Int(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_RechargesMultiples(NumeroObjet))) & _
                                              " " & _
                                              IIf(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_RechargesMultiples(NumeroObjet)) >= 2, Parametres.EtiquetteRechargesMultiples, Parametres.EtiquetteRechargeMultiple)
                End If
            End If
        End If
        
        If Parametres.Objet_BonusPorteeAttaque(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & Parametres.EtiquettePortee & " : " & _
            Perso.Objet_Equipe_PorteeSup(NiveauCompetence, Parametres.Objet_BonusPorteeAttaque(NumeroObjet)) & _
            " " & Parametres.EtiquetteUniteDistance
        ElseIf Parametres.Objet_PorteeUtilisation(NumeroObjet) > 0 And Not Effet Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & Parametres.EtiquettePortee & " : " & _
            Perso.Objet_Equipe_PorteeSup(NiveauCompetence, Parametres.Objet_PorteeUtilisation(NumeroObjet)) & _
            " " & Parametres.EtiquetteUniteDistance
        End If
        If Parametres.Objet_BonusAttaque(NumeroObjet) <> 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & IIf(Parametres.Objet_BonusAttaque(NumeroObjet) > 0, "#+", "#&") & _
            Int(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_BonusAttaque(NumeroObjet)) * Perso.CoefRaceAttaque) & _
            " " & Parametres.EtiquetteAttaque
            'Type d'attaque.
            If Parametres.Objet_BonusAttaque(NumeroObjet) > 0 Then
                If Parametres.Objet_Categorie(NumeroObjet) = 0 Or Parametres.Objet_Categorie(NumeroObjet) = 1 Then
                    If Parametres.Objet_AttaqueResistance(NumeroObjet) > -1 Then
                        Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & _
                                                  " (" & Parametres.Resistance_Attaque(Parametres.Objet_AttaqueResistance(NumeroObjet)) & ")"
                    End If
                End If
                If Parametres.Objet_AttaqueCauseEffet(NumeroObjet) > 0 Then
                    Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#$" & Parametres.Objet_Nom(Parametres.Objet_AttaqueCauseEffet(NumeroObjet)) & IIf(Parametres.Objet_Duree(Parametres.Objet_AttaqueCauseEffet(NumeroObjet)) = -1, "", " : " & CLng(Perso.Objet_Equipe_GainSup(Perso.Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(Parametres.Objet_AttaqueCauseEffet(NumeroObjet))), Parametres.Objet_Duree(Parametres.Objet_AttaqueCauseEffet(NumeroObjet)))) & " " & Parametres.EtiquetteUniteDuree)
                End If
            End If
        End If
        'Surface de dégât de l'attaque de l'objet.
        If Parametres.Objet_SurfaceAttaque(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
            Parametres.EtiquetteSurfaceAttaque & " : " & _
            Perso.Objet_Equipe_SurfaceSup(NiveauCompetence, Parametres.Objet_SurfaceAttaque(NumeroObjet))
        End If
        'Duree de recharge de l'attaque.
        If Parametres.Objet_DureeRechargeAttaque(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
                                      Parametres.EtiquetteDureeRechargeAttaque & " : " & _
                                      Parametres.Objet_DureeRechargeAttaque(NumeroObjet)
        End If
        If Parametres.Objet_BonusDefense(NumeroObjet) <> 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & IIf(Parametres.Objet_BonusDefense(NumeroObjet) > 0, "#+", "#&") & _
            Int(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_BonusDefense(NumeroObjet)) * Perso.CoefRaceDefense) & _
            " " & Parametres.EtiquetteDefense
            '" (" & Parametres.Objet_BonusDefense(NumeroObjet) & ") " &
        End If
        If Parametres.Objet_Categorie(NumeroObjet) = 1 And Not Parametres.Objet_Utilisable(NumeroObjet) Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
            Parametres.EtiquetteArme2mains
        End If
        If Parametres.Objet_BonusVie(NumeroObjet) <> 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & IIf(Parametres.Objet_BonusVie(NumeroObjet) > 0, "#+", "#&") & _
            Int(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_BonusVie(NumeroObjet)) * Perso.CoefRaceVie) & _
            " " & Parametres.EtiquetteVie
            '" (" & Parametres.Objet_BonusVie(NumeroObjet) & ") " &
        End If
        If Parametres.Objet_BonusEnergie(NumeroObjet) <> 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & IIf(Parametres.Objet_BonusEnergie(NumeroObjet) > 0, "#+", "#&") & _
            Int(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_BonusEnergie(NumeroObjet)) * Perso.CoefRaceEnergie) & _
            " " & Parametres.EtiquetteEnergie
            '" (" & Parametres.Objet_BonusEnergie(NumeroObjet) & ") " &
        End If
        If Parametres.Objet_BonusMagie(NumeroObjet) <> 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & IIf(Parametres.Objet_BonusMagie(NumeroObjet) > 0, "#+", "#&") & _
            Int(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_BonusMagie(NumeroObjet)) * Perso.CoefRaceMagie) & _
            " " & Parametres.EtiquetteMagie
            '" (" & Parametres.Objet_BonusMagie(NumeroObjet) & ") " &
        End If
        If Parametres.Objet_BonusMoral(NumeroObjet) <> 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & IIf(Parametres.Objet_BonusMoral(NumeroObjet) > 0, "#+", "#&") & _
            Int(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_BonusMoral(NumeroObjet)) * Perso.CoefRaceMoral) & _
            " " & Parametres.EtiquetteMoral
            '" (" & Parametres.Objet_BonusMoral(NumeroObjet) & ") " &
        End If
        'Bonus de déplacement positifs.
        If Parametres.Objet_BonusDeplacement(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#+" & _
            Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_BonusDeplacement(NumeroObjet)) & _
            " " & Parametres.EtiquetteDeplacement
            '" (" & Parametres.Objet_BonusDeplacement(NumeroObjet) & ") " &
        End If
        'Bonus de déplacement négatifs.
        If Parametres.Objet_BonusDeplacement(NumeroObjet) < 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#&" & _
            Parametres.Objet_BonusDeplacement(NumeroObjet) & _
            " " & Parametres.EtiquetteDeplacement
            '" (" & Parametres.Objet_BonusDeplacement(NumeroObjet) & ") " &
        End If
        If Parametres.Objet_BonusVitesse(NumeroObjet) <> 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & IIf(Parametres.Objet_BonusVitesse(NumeroObjet) > 0, "#+", "#&") & _
            Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_BonusVitesse(NumeroObjet)) & _
            " " & Parametres.EtiquetteVitesse
            '" (" & Parametres.Objet_BonusDeplacement(NumeroObjet) & ") " &
        End If
        If Parametres.Objet_BonusVitalite(NumeroObjet) <> 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & IIf(Parametres.Objet_BonusVitalite(NumeroObjet) > 0, "#+", "#&") & _
            Parametres.Objet_BonusVitalite(NumeroObjet) * 100 & _
            Parametres.EtiquettePourcent & " " & Parametres.EtiquetteVitalite
            '" (" & Parametres.Objet_BonusDeplacement(NumeroObjet) & ") " &
        End If
        'Bonus d'armure.
        If Parametres.Objet_BonusArmure(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#+" & _
            Perso.Objet_Equipe_ArmureSup(NiveauCompetence, Parametres.Objet_BonusArmure(NumeroObjet)) & _
            " " & Parametres.EtiquetteArmure
        End If
        For j = 0 To Parametres.NombreCompetencesSpeciales - 1
            If Parametres.Objet_BonusCompetencesSpeciales(NumeroObjet, j) > 0 Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#+" & _
                Int(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_BonusCompetencesSpeciales(NumeroObjet, j)) * Perso.Definir_CoefRaceCompSpeciales(j)) & _
                " " & Parametres.Speciales_NomCompetences(j)
            End If
        Next j
        For j = 0 To Parametres.NombreRessources - 1
            If Parametres.Objet_BonusCompetencesRessources(NumeroObjet, j) > 0 Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#+" & _
                Int(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_BonusCompetencesRessources(NumeroObjet, j)) * Perso.Definir_CoefRaceCompRessources(j)) & _
                " " & Parametres.Ressources_NomCompetence(j)
            End If
        Next j
        For j = 0 To Perso.Nombre_CompetencesServices - 1
            If Parametres.Objet_BonusCompetencesServices(NumeroObjet, j) > 0 Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#+" & _
                Int(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_BonusCompetencesServices(NumeroObjet, j)) * Perso.Definir_CoefRaceCompServices(j)) & _
                " " & Parametres.Service_NomCompetence(j + 1)
            End If
        Next j
        For k = 1 To Parametres.NombreCompetencesObjets - 1
            j = Parametres.CompetenceObjet_NoListe(k)
            If Parametres.Objet_BonusCompetencesObjets(NumeroObjet, j) > 0 Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#+" & _
                Int(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_BonusCompetencesObjets(NumeroObjet, j)) * Perso.Definir_CoefRaceCompObjets(j)) & _
                " " & Parametres.CompetenceObjet_Nom(j)
            End If
        Next k
        'Bonus de transport de ressources.
        If Parametres.Objet_BonusMaxRessources(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & Parametres.EtiquetteTranport & " " & _
            Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_BonusMaxRessources(NumeroObjet)) & _
            " " & Parametres.EtiquetteRessourcesSupplementaires
        End If
        
        'Bonus de régénération.
        If Parametres.Objet_BonusVitesseRegenerationVie(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#$" & Parametres.EtiquetteRegenere & " " & _
            Format(Parametres.Objet_BonusVitesseRegenerationVie(NumeroObjet), "0.000000000") * 1000 & _
            Parametres.EtiquettePourcent & " " & Parametres.EtiquetteVie
        End If
        If Parametres.Objet_BonusVitesseRegenerationEnergie(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#$" & Parametres.EtiquetteRegenere & " " & _
            Format(Parametres.Objet_BonusVitesseRegenerationEnergie(NumeroObjet), "0.000000000") * 1000 & _
            Parametres.EtiquettePourcent & " " & Parametres.EtiquetteEnergie
        End If
        If Parametres.Objet_BonusVitesseRegenerationMagie(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#$" & Parametres.EtiquetteRegenere & " " & _
            Format(Parametres.Objet_BonusVitesseRegenerationMagie(NumeroObjet), "0.000000000") * 1000 & _
            Parametres.EtiquettePourcent & " " & Parametres.EtiquetteMagie
        End If
        If Parametres.Objet_BonusVitesseRegenerationMoral(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#$" & Parametres.EtiquetteRegenere & " " & _
            Format(Parametres.Objet_BonusVitesseRegenerationMoral(NumeroObjet), "0.000000000") * 1000 & _
            Parametres.EtiquettePourcent & " " & Parametres.EtiquetteMoral
        End If
        
        'Affiche les infos d'invocation.
        If Parametres.Objet_Invocation(NumeroObjet) >= 0 Then
            'Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
                                       "$" & Parametres.EtiquetteInvoque & " : " & _
                                       IIf(IIf(Jeu.Regle_Invocations_Multiples, Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_Quantite(NumeroObjet)), Parametres.Objet_Quantite(NumeroObjet)) < 1, "1", IIf(Jeu.Regle_Invocations_Multiples, Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_Quantite(NumeroObjet)), Parametres.Objet_Quantite(NumeroObjet))) & _
                                       " " & Parametres.Race_Nom(Parametres.Objet_Invocation(NumeroObjet), , IIf(Jeu.Regle_Invocations_Multiples, Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_Quantite(NumeroObjet)), Parametres.Objet_Quantite(NumeroObjet)) > 1)
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
                                      "$" & Parametres.EtiquetteInvoque & " : " & _
                                      Int(Perso.Objet_Equipe_InvocationSup(NiveauCompetence, Parametres.Objet_Quantite(NumeroObjet))) & _
                                      " " & Parametres.Race_Nom(Parametres.Objet_Invocation(NumeroObjet), , Perso.Objet_Equipe_InvocationSup(NiveauCompetence, Parametres.Objet_Quantite(NumeroObjet)) >= 2)
            'Chances d'invoquer 1 créature supplémentaire.
            If Int(Perso.Objet_Equipe_InvocationSup(NiveauCompetence, Parametres.Objet_Quantite(NumeroObjet))) <> Perso.Objet_Equipe_InvocationSup(NiveauCompetence, Parametres.Objet_Quantite(NumeroObjet)) Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
                                          "$" & _
                                          Int((Perso.Objet_Equipe_InvocationSup(NiveauCompetence, Parametres.Objet_Quantite(NumeroObjet)) - Int(Perso.Objet_Equipe_InvocationSup(NiveauCompetence, Parametres.Objet_Quantite(NumeroObjet)))) * 100) & _
                                          Parametres.EtiquetteInvoqueChance1 & _
                                          Parametres.Race_Nom(Parametres.Objet_Invocation(NumeroObjet)) & _
                                          Parametres.EtiquetteInvoqueChance2
            End If
            'Limite d'invocations.
            If Parametres.Objet_QuantiteMax(NumeroObjet) > 0 Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#&" & _
                                          IIf(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_QuantiteMax(NumeroObjet)) >= 2, Parametres.EtiquetteInvoqueMax1Pluriel, Parametres.EtiquetteInvoqueMax1) & " " & _
                                          IIf(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_QuantiteMax(NumeroObjet)) = 0, "1", Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_QuantiteMax(NumeroObjet))) & " " & LCase(Parametres.Race_Nom(Parametres.Objet_Invocation(NumeroObjet), , Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_QuantiteMax(NumeroObjet)) >= 2)) & _
                                          IIf(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_QuantiteMax(NumeroObjet)) >= 2, Parametres.EtiquetteInvoqueMax2Pluriel, Parametres.EtiquetteInvoqueMax2)
            End If
        End If
        
        'Affiche les infos de resurrection.
        If Parametres.Objet_Resurrection(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#$" & Parametres.EtiquetteRessuscite & " " & _
            Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_Quantite(NumeroObjet)) & _
            " " & IIf(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_Quantite(NumeroObjet)) >= 2, Parametres.EtiquettePersonnes, Parametres.EtiquettePersonne)
        End If
        If Parametres.Objet_SelfResurrection(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & _
                "#$" & Parametres.EtiquetteRessusciteSoi
        End If
        
        'Berserk.
        If Parametres.Objet_Berserk(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & _
                "#$" & Parametres.EtiquetteBerserk
        End If
        
        'Affiche les gains de caractéristiques.
        If Parametres.Objet_GainVie(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#$" & Parametres.EtiquetteRedonne & " : " & _
            Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_GainVie(NumeroObjet)) & _
            " " & Parametres.EtiquetteVie
        End If
        If Parametres.Objet_GainEnergie(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#$" & Parametres.EtiquetteRedonne & " : " & _
            Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_GainEnergie(NumeroObjet)) & _
            " " & Parametres.EtiquetteEnergie
        End If
        If Parametres.Objet_GainMagie(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#$" & Parametres.EtiquetteRedonne & " : " & _
            Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_GainMagie(NumeroObjet)) & _
            " " & Parametres.EtiquetteMagie
        End If
        If Parametres.Objet_GainMoral(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#$" & Parametres.EtiquetteRedonne & " : " & _
            Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_GainMoral(NumeroObjet)) & _
            " " & Parametres.EtiquetteMoral
        End If
        'Poison.
        If Parametres.Objet_Poison(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#&" & Parametres.EtiquettePoison & " : " & _
                                      IIf(Int(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_Poison(NumeroObjet))) < 1, "1", Int(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_Poison(NumeroObjet))))

        End If
        If Not Effet Then
            If Parametres.Objet_GainVie(NumeroObjet) + _
               Parametres.Objet_GainEnergie(NumeroObjet) + _
               Parametres.Objet_GainMagie(NumeroObjet) + _
               Parametres.Objet_GainMoral(NumeroObjet) > 0 And _
               Parametres.Objet_Resurrection(NumeroObjet) = 0 And _
               Parametres.Objet_Quantite(NumeroObjet) > 0 Or _
               Parametres.Objet_CreerObjetAllie(NumeroObjet) > 0 Or _
               Parametres.Objet_CreerObjetEnnemi(NumeroObjet) > 0 Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & Parametres.EtiquetteAffecte & " " & _
                Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_Quantite(NumeroObjet)) & _
                " " & IIf(Perso.Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_Quantite(NumeroObjet)) >= 2, Parametres.EtiquettePersonnes, Parametres.EtiquettePersonne)
            End If
        End If
        'Affiche les gains de ressources.
        If Parametres.Objet_GainArgent(NumeroObjet) > 0 Then
            Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & _
                                      "#$" & Parametres.EtiquetteGenere & " " & _
                                      Perso.Objet_Equipe_GainSup(NiveauCompetence, Parametres.Objet_GainArgent(NumeroObjet)) & _
                                      " " & Parametres.EtiquetteArgent
        End If
        For j = 0 To Parametres.NombreRessources - 1
            If Parametres.Objet_GainRessource(NumeroObjet, j) > 0 Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & _
                                          "#$" & Parametres.EtiquetteGenere & " " & _
                                          Perso.Objet_Equipe_GainSup(NiveauCompetence, Parametres.Objet_GainRessource(NumeroObjet, j)) & _
                                          " " & LCase(Parametres.Ressources_Nom(j))
            End If
        Next j
        'Résistances.
        For j = 0 To Parametres.NombreResistances - 1
            If Parametres.Objet_BonusResistance(NumeroObjet, j) <> 0 Then
                If Parametres.Objet_BonusResistance(NumeroObjet, j) > 0 Then
                    Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#$"
                Else
                    Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#&"
                End If
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & _
                                          Parametres.Resistance_NomComplet(j) & " : " & _
                                          IIf(Parametres.Objet_BonusResistance(NumeroObjet, j) > 0, "+", "") & _
                                          Parametres.Objet_BonusResistance(NumeroObjet, j) * 100 & Parametres.EtiquettePourcent
            End If
        Next j
        If Not Effet Then
            'Affiche la durée.
            If Parametres.Objet_Duree(NumeroObjet) > 0 Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & Parametres.EtiquetteDuree & " : " & _
                                          Int(Perso.Objet_Equipe_GainSup(NiveauCompetence, Parametres.Objet_Duree(NumeroObjet))) & " " & Parametres.EtiquetteUniteDuree
            End If
            'Affiche les pertes de caractéristiques.
            If Parametres.Objet_PerteVie(NumeroObjet) > 0 Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#(- " & _
                Parametres.Objet_PerteVie(NumeroObjet) & " " & _
                Parametres.EtiquetteVie & " " & Parametres.EtiquetteParUtilisation & ")"
            End If
            If Parametres.Objet_PerteEnergie(NumeroObjet) > 0 Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#(- " & _
                Parametres.Objet_PerteEnergie(NumeroObjet) & " " & _
                Parametres.EtiquetteEnergie & " " & Parametres.EtiquetteParUtilisation & ")"
            End If
            If Parametres.Objet_PerteMagie(NumeroObjet) > 0 Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#(- " & _
                Parametres.Objet_PerteMagie(NumeroObjet) & " " & _
                Parametres.EtiquetteMagie & " " & Parametres.EtiquetteParUtilisation & ")"
            End If
            If Parametres.Objet_PerteMoral(NumeroObjet) > 0 Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#(- " & _
                Parametres.Objet_PerteMoral(NumeroObjet) & " " & _
                Parametres.EtiquetteMoral & " " & Parametres.EtiquetteParUtilisation & ")"
            End If
            If Parametres.Objet_PerteArgent(NumeroObjet) > 0 Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & IIf(Parametres.Objet_PerteArgent(NumeroObjet) > Perso.Definir_Argent, "&", "") & "(- " & _
                Parametres.Objet_PerteArgent(NumeroObjet) & " " & _
                Parametres.EtiquettePieces & " " & Parametres.EtiquetteParUtilisation & ")"
            End If
            If Parametres.Objet_PerteRessourceQuantite(NumeroObjet) > 0 Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#(" & Parametres.EtiquetteUtilisation & " : " & _
                Parametres.Objet_PerteRessourceQuantite(NumeroObjet) & " " & _
                Parametres.Ressources_Nom(Parametres.Objet_PerteRessourceIndice(NumeroObjet)) & ")"
                '* (Perso.NombreProjectiles + IIf(Equipe, 0, Int(Perso.Objet_Equipe_GainSup(NiveauCompetence, Parametres.Objet_ProjectileSupplementaire(NumeroObjet)))))
                If Parametres.Objet_Utilisable(NumeroObjet) Then
                    Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & _
                    "#" & IIf(Perso.Definir_ressources(Parametres.Objet_PerteRessourceIndice(NumeroObjet)) < Parametres.Objet_PerteRessourceQuantite(NumeroObjet), "&" & Parametres.EtiquetteUtilisationsRessourcesInsuffisantes, _
                    "(" & Int(Perso.Definir_ressources(Parametres.Objet_PerteRessourceIndice(NumeroObjet)) / Parametres.Objet_PerteRessourceQuantite(NumeroObjet)) & IIf(Perso.Definir_ressources(Parametres.Objet_PerteRessourceIndice(NumeroObjet)) / Parametres.Objet_PerteRessourceQuantite(NumeroObjet) >= 2, " " & Parametres.EtiquetteUtilisationsPossibles & ")", " " & Parametres.EtiquetteUtilisationPossible & ")"))
                Else
                    'Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & _
                    "#" & IIf(Perso.Definir_ressources(Parametres.Objet_PerteRessourceIndice(NumeroObjet)) < Parametres.Objet_PerteRessourceQuantite(NumeroObjet) * (Perso.NombreProjectiles + IIf(Equipe, 0, Int(Perso.Objet_Equipe_GainSup(NiveauCompetence, Parametres.Objet_ProjectileSupplementaire(NumeroObjet))))), "&" & Parametres.EtiquetteUtilisationsMunitionsInsuffisantes, _
                    "(" & Int(Perso.Definir_ressources(Parametres.Objet_PerteRessourceIndice(NumeroObjet)) / Parametres.Objet_PerteRessourceQuantite(NumeroObjet) / (Perso.NombreProjectiles + IIf(Equipe, 0, Int(Perso.Objet_Equipe_GainSup(NiveauCompetence, Parametres.Objet_ProjectileSupplementaire(NumeroObjet)))))) & IIf(Perso.Definir_ressources(Parametres.Objet_PerteRessourceIndice(NumeroObjet)) / Parametres.Objet_PerteRessourceQuantite(NumeroObjet) / (Perso.NombreProjectiles + IIf(Equipe, 0, Parametres.Objet_ProjectileSupplementaire(NumeroObjet))) >= 2, " " & Parametres.EtiquetteUtilisationMunitions & ")", " " & Parametres.EtiquetteUtilisationMunition & ")"))
                    Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & _
                    "#" & IIf(Perso.Definir_ressources(Parametres.Objet_PerteRessourceIndice(NumeroObjet)) < Parametres.Objet_PerteRessourceQuantite(NumeroObjet), "&" & Parametres.EtiquetteUtilisationsMunitionsInsuffisantes, _
                    "(" & Int(Perso.Definir_ressources(Parametres.Objet_PerteRessourceIndice(NumeroObjet)) / Parametres.Objet_PerteRessourceQuantite(NumeroObjet)) & IIf(Perso.Definir_ressources(Parametres.Objet_PerteRessourceIndice(NumeroObjet)) / Parametres.Objet_PerteRessourceQuantite(NumeroObjet) >= 2, " " & Parametres.EtiquetteUtilisationMunitions & ")", " " & Parametres.EtiquetteUtilisationMunition & ")"))
                End If
                'Parametres.Ressources_Nom(Parametres.Objet_PerteRessourceIndice(NumeroObjet)) & " " & "restant : " &
            End If
            'Nombre de projectiles à la fois.
            If Perso.Objet_Equipe_GainSup(NiveauCompetence, Parametres.Objet_ProjectileSupplementaire(NumeroObjet)) >= 1 Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#$" & _
                                          Int(Perso.Objet_Equipe_GainSup(NiveauCompetence, Parametres.Objet_ProjectileSupplementaire(NumeroObjet))) + 1 & _
                                          " " & Parametres.EtiquetteProjectilesSupplementaires
            End If
            'Affiche s'il est indestructible.
            If Parametres.Objet_Indestructible(NumeroObjet) Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
                                          "$" & Parametres.EtiquetteIndestructible
            End If
            'Affiche le nombre d'utilisations.
            If Parametres.Objet_UtilisationUnique(NumeroObjet) Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
                                          Parametres.EtiquetteUtilisationUnique
            End If
            'Affiche si l'on peut l'utiliser quand on est mort.
            If Parametres.Objet_UtilisationMort(NumeroObjet) Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
                                          "$" & Parametres.EtiquetteUtilisationMort
            End If
            'Affiche le prix.
            If Not Parametres.Objet_Inseparable(NumeroObjet) And Parametres.Objet_PrixVente(NumeroObjet) > 0 Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
                    Parametres.EtiquettePrix & " : "
                If Inventaire Or Perso.DansSaMaison Then  'Les objets dans l'inventaire son moins cher à vendre.
                    Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & _
                                              Parametres.Objet_PrixVente(NumeroObjet)
                Else
                    Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & _
                                              Perso.Prix_Marchandee(Parametres.Objet_PrixVente(NumeroObjet) * Parametres.CoefficientAchat, Parametres)
                End If
            End If
            'Affiche la description.
            If Parametres.Objet_Description(NumeroObjet) <> "" Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
                    Parametres.Objet_Description(NumeroObjet)
            End If
        End If
        If Inventaire Then
            'If Equipe Then
                'Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
                                          "(Clique pour ranger)"
            'Else
                'Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
                                          "(Clique gauche pour utiliser)"
                If Perso.DansSonChateau Or Perso.DansUnChateau Or (Perso.DansUneMaison And Maisons(Perso.IndiceMaison).Marche) Then
                    If Parametres.Objet_Inseparable(NumeroObjet) Then
                        Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
                                                  "&" & Parametres.EtiquetteObjetInvendable
                    Else
                        Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
                                                  Parametres.EtiquetteObjetVendre
                    End If
                Else
                    If Parametres.Objet_Inseparable(NumeroObjet) Then
                        Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
                                                  "&" & Parametres.EtiquetteObjetIndeposable
                    End If
                    'If Not Parametres.Objet_Inseparable(NumeroObjet) Then
                    '    Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
                                                  "(Clique droit pour se débarasser de l'objet)"
                    'End If
                End If
                If Parametres.Objet_Monture(NumeroObjet) And Not Parametres.Race_UtilisationMonture(Perso.Race) Then
                    Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
                                                  "&" & Parametres.EtiquetteUtilisationMontureImpossible2
                End If
            'End If
        End If
        
        'Durée pour les effets temporaires.
        If Effet Then
            If Duree <> -1 Then
                Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & _
                                         Parametres.EtiquetteDureeRestant & " : " & Int(Duree) & " " & Parametres.EtiquetteUniteDuree
            End If
            If Parametres.Objet_EffetAnnulationImpossible(NumeroObjet) Then
                If Not Parametres.Objet_UtilisationAutomatique(NumeroObjet) Then
                    Definir_InfoBulle_Objet = Definir_InfoBulle_Objet & "#" & Parametres.EtiquetteEffetAnnulationImpossible
                End If
            End If
        End If
    End If
End Function

Private Sub Afficher_ListeRessourcesPerso(ByRef Perso As ClsJeuPerso, _
                                          ByRef Maison As ClsJeuBatiment, _
                                          ByRef Parametres As ClsJeuParam, _
                                          ByRef Ressources As ClsAffObjets, _
                                          ByRef Fenetre As ClsAffFenetre, _
                                          ByRef Message As ClsIntMessage, _
                                          ByRef SurfBack As DirectDrawSurface7)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim DecalageX As Integer
    DecalageX = Int(DefilementDecalageX)
    
    'Affiche les boutons de transaction global.
    For i = 0 To NombreBoutonsCommerce - 1
        With BoutonsCommerce(UBound(BoutonsCommerce()) - i)
        '.Visible = True
        AffBoutonsCommerce.Afficher .PositionX + PositionX, _
                                    .PositionY + PositionY, _
                                    i + NombreBoutonsCommerce * 2, Abs(Not Perso.Vivant), SurfBack
        'If Perso.DansUnChateau Or Perso.DansSonChateau Then
        'Else
        '    .Visible = False
        'End If
        End With
    Next i
    
    If DecalageX > 0 Then
        For i = 0 To DecalageX - 1
            If Parametres.Ressources_Mangeable(i) Then
                BoutonsManger(k).Visible = False
                k = k + 1
            End If
        Next i
    End If
    For i = 0 To NombreLignesRessources - 1
        If i + DecalageX < Parametres.NombreRessources Then
            Message.Afficher RessourcesPositionX + PositionX + RessourcesIntervalX * i, _
                             RessourcesPositionY + PositionY + RessourcesIntervalY * i, _
                             IIf(Perso.DansUnChateau Or Perso.DansSonChateau, Parametres.EtiquettePrix & " : " & Parametres.Ressources_PrixVente(i + DecalageX), Parametres.Ressources_Nom(i + DecalageX) & " :"), _
                             SurfBack
            Message.Afficher RessourcesPositionX + RessourcesScoreDecalageX + PositionX + RessourcesIntervalX * i, _
                             RessourcesPositionY + RessourcesScoreDecalageY + PositionY + RessourcesIntervalY * i, _
                             Int(Perso.Definir_ressources(i + DecalageX)) & IIf(Perso.Definir_MaxRessources(i + DecalageX) < 100, " / ", "/") & _
                             Int(Perso.Definir_MaxRessources(i + DecalageX)), _
                             SurfBack
            Afficher_Separateur_Ressources RessourcesPositionX + PositionX + RessourcesIntervalX * i, _
                                           RessourcesPositionY + PositionY + RessourcesIntervalY * i, _
                                           SurfBack
            Ressources.Afficher RessourcesPositionX + PositionX + RessourcesIntervalX * i + RessourcesImageDecalageX, _
                                RessourcesPositionY + PositionY + RessourcesIntervalY * i + RessourcesImageDecalageY, _
                                i + DecalageX, _
                                SurfBack
            'On affiche les boutons de transaction de ressources.
            For j = 0 To NombreBoutonsCommerce - 1
                'If Perso.DansUnChateau Or Perso.DansSonChateau Then
                BoutonsCommerce((i + DecalageX) * (NombreBoutonsCommerce) + j).PositionX = RessourcesPositionX + RessourcesIntervalX * i + BoutonCommerceDecalageX + BoutonCommerceIntervalX * j
                BoutonsCommerce((i + DecalageX) * (NombreBoutonsCommerce) + j).PositionY = RessourcesPositionY + RessourcesIntervalY * i + BoutonCommerceDecalageY + BoutonCommerceIntervalY * j
                '    BoutonsCommerce((i + DecalageX) * (NombreBoutonsCommerce) + j).Visible = True
                AffBoutonsCommerce.Afficher BoutonsCommerce(i * NombreBoutonsCommerce + j).PositionX + PositionX, _
                                            BoutonsCommerce((i + DecalageX) * (NombreBoutonsCommerce) + j).PositionY + PositionY, _
                                            j + IIf(Perso.DansSonChateau Or Perso.DansUnChateau, NombreBoutonsCommerce, 0), Abs(Not Perso.Vivant), SurfBack
                'Else
                '    BoutonsCommerce((i + DecalageX) * (NombreBoutonsCommerce) + j).Visible = False
                'End If
            Next j
            'Affiche les boutons Manger.
            If Parametres.Ressources_Mangeable(i + DecalageX) Then
                BoutonsManger(k).PositionX = RessourcesPositionX + RessourcesIntervalX * i + BoutonMangerDecalageX
                BoutonsManger(k).PositionY = RessourcesPositionY + RessourcesIntervalY * i + BoutonMangerDecalageY
                BoutonsManger(k).Visible = True
                If Perso.Vivant Then
                    AffBoutonAction.Afficher BoutonsManger(k).PositionX + PositionX, _
                                              BoutonsManger(k).PositionY + PositionY, _
                                              0, 0, SurfBack
                Else
                    AffBoutonAction.Afficher BoutonsManger(k).PositionX + PositionX, _
                                              BoutonsManger(k).PositionY + PositionY, _
                                              0, 1, SurfBack
                End If
                k = k + 1
            Else
                
            End If
            'Afficher les boutons de transfert de ressources.
            If Perso.DansSaMaison And _
               Parametres.Batiment_RessourceFabrique(Maison.TypeBatiment) = i + DecalageX And _
               Maison.MaxStock > 0 Then
                For j = 0 To NombreBoutonsTransfert - 1
                    BoutonsTransfertStockPerso(j).PositionX = RessourcesPositionX + RessourcesIntervalX * i + BoutonTransfertDecalageX + BoutonTransfertIntervalX * j
                    BoutonsTransfertStockPerso(j).PositionY = RessourcesPositionY + RessourcesIntervalY * (i) + BoutonTransfertDecalageY + BoutonTransfertIntervalY * j
                    BoutonsTransfertStockPerso(j).Visible = True
                    AffBoutonsTransfert.Afficher BoutonsTransfertStockPerso(j).PositionX + PositionX, _
                                                 BoutonsTransfertStockPerso(j).PositionY + PositionY, _
                                                 j + NombreBoutonsTransfert, 0, SurfBack
                Next j
            End If
            'Affiche les boutons pour prendre des ressources.
            BoutonsRessourcesPrendre(i + DecalageX).PositionX = RessourcesPositionX + RessourcesIntervalX * i + RessourcesBoutonsPrendreDecalageX
            BoutonsRessourcesPrendre(i + DecalageX).PositionY = RessourcesPositionY + RessourcesIntervalY * i + RessourcesBoutonsPrendreDecalageY
        Else
            i = NombreLignesRessources
        End If
    Next i
    Afficher_BoutonsDefilement Parametres.NombreRessources - NombreLignesRessources, SurfBack
End Sub

Private Sub Afficher_ListeRessourcesFief(ByRef Perso As ClsJeuPerso, _
                                         ByRef Fief As ClsJeuFief, _
                                         ByRef Parametres As ClsJeuParam, _
                                         ByRef Ressources As ClsAffObjets, _
                                         ByRef Fenetre As ClsAffFenetre, _
                                         ByRef Message As ClsIntMessage, _
                                         ByRef SurfBack As DirectDrawSurface7)
    Dim i As Long
    Dim j As Long
    Dim DecalageX As Integer
    DecalageX = Int(DefilementDecalageX)
    For i = 0 To NombreLignesRessources - 1
        If i + DecalageX < Parametres.NombreRessources Then
            Message.Afficher RessourcesPositionX + PositionX + RessourcesIntervalX * i, _
                             RessourcesPositionY + PositionY + RessourcesIntervalY * i, _
                             IIf(Perso.DansUnChateau Or Perso.DansSonChateau, Parametres.EtiquettePrix & " : " & Parametres.Ressources_PrixVente(i + DecalageX) * Parametres.CoefficientAchat, Parametres.Ressources_Nom(i + DecalageX) & " :"), _
                             SurfBack
            Message.Afficher RessourcesPositionX + RessourcesScoreDecalageX + PositionX + RessourcesIntervalX * i, _
                             RessourcesPositionY + RessourcesScoreDecalageY + PositionY + RessourcesIntervalY * i, _
                             Int(Fief.Definir_ressources(i + DecalageX)), _
                             SurfBack
            Afficher_Separateur_Ressources RessourcesPositionX + PositionX + RessourcesIntervalX * i, _
                                           RessourcesPositionY + PositionY + RessourcesIntervalY * i, _
                                           SurfBack
            Ressources.Afficher RessourcesPositionX + PositionX + RessourcesIntervalX * i + RessourcesImageDecalageX, _
                                RessourcesPositionY + PositionY + RessourcesIntervalY * i + RessourcesImageDecalageY, _
                                i + DecalageX, _
                                SurfBack
            'On affiche les boutons de transaction de ressources.
            For j = 0 To NombreBoutonsCommerce - 1
                BoutonsCommerce((i + DecalageX) * (NombreBoutonsCommerce) + j).PositionX = RessourcesPositionX + RessourcesIntervalX * i + BoutonCommerceDecalageX + BoutonCommerceIntervalX * j
                BoutonsCommerce((i + DecalageX) * (NombreBoutonsCommerce) + j).PositionY = RessourcesPositionY + RessourcesIntervalY * i + BoutonCommerceDecalageY + BoutonCommerceIntervalY * j
                AffBoutonsCommerce.Afficher BoutonsCommerce((i + DecalageX) * (NombreBoutonsCommerce) + j).PositionX + PositionX, _
                                            BoutonsCommerce((i + DecalageX) * (NombreBoutonsCommerce) + j).PositionY + PositionY, _
                                            j, IIf(Perso.Vivant, 0, 1), SurfBack
                BoutonsCommerce((i + DecalageX) * NombreBoutonsCommerce + j).InfoBulle = _
                    Parametres.EtiquettePrix & " : " & Parametres.Ressources_PrixVente(i + DecalageX) * Parametres.CoefficientAchat
            Next j
        Else
            i = NombreLignesRessources
        End If
    Next i
    Afficher_BoutonsDefilement Parametres.NombreRessources - NombreLignesRessources, SurfBack
End Sub

Private Sub Afficher_Separateur_Ressources(X As Integer, Y As Integer, ByRef SurfBack As DirectDrawSurface7)
    SurfBack.DrawLine X + SeparateurRessourceDecalageX, _
                      Y + SeparateurRessourceDecalageY, _
                      X + SeparateurRessourceDecalageX + SeparateurRessourceLongueur, _
                      Y + SeparateurRessourceDecalageY
End Sub

Private Sub Afficher_BoutonsDefilement(ByVal NombreLignes As Integer, _
                                       ByRef SurfBack As DirectDrawSurface7, _
                                       Optional ByVal Objet As Boolean)
    DefilementNombreLigne = NombreLignes
    If AfficherBoutonsDefilement Then
        'If DefilementDecalageX > 0 Then
        BoutonsDefilement(0).Visible = True
        AffBoutonsDefilement.Afficher BoutonsDefilement(0).PositionX + PositionX, _
                                      BoutonsDefilement(0).PositionY + PositionY, _
                                      0, 0, SurfBack, True
        'Else
        '    BoutonsDefilement(0).Visible = False
        'End If
        'Debug.Print DefilementDecalageX & " < " & DefilementNombreLigne
        'If DefilementDecalageX < DefilementNombreLigne Then
        AffBoutonsDefilement.Afficher BoutonsDefilement(1).PositionX + PositionX, _
                                      BoutonsDefilement(1).PositionY + PositionY, _
                                      1, 0, SurfBack, True
        BoutonsDefilement(1).Visible = True
        'Else
        '    BoutonsDefilement(1).Visible = False
        'End If
        
        'Affiche la barre de défilement.
        AffBoutonsBarreDefilement.Afficher BoutonsDefilement(2).PositionX + PositionX, _
                                           BoutonsDefilement(2).PositionY + PositionY, _
                                           2, 0, SurfBack, True
        BoutonsDefilement(2).Visible = True
        'Affiche le carre dans la barre de défilement.
        AffBoutonsCarreDefilement.Afficher BoutonsDefilement(2).PositionX + PositionX, _
                                           BoutonsDefilement(2).PositionY + (AffBoutonsBarreDefilement.Hauteur - AffBoutonsCarreDefilement.Hauteur) * IIf(Objet, DefilementObjetDecalageX, DefilementDecalageX) / DefilementNombreLigne + PositionY, _
                                           3, 0, SurfBack, True
    End If
End Sub

'Private Sub Afficher_BoutonsDefilementObjet(ByVal NombreLignes As Integer, _
                                            ByRef SurfBack As DirectDrawSurface7)
'    DefilementNombreLigne = NombreLignes
'    If AfficherBoutonsDefilement Then
'        BoutonsDefilement(0).PositionX = BoutonDefilement1PositionX
'        BoutonsDefilement(0).PositionY = BoutonDefilement1PositionY
'        BoutonsDefilement(1).PositionX = BoutonDefilement2PositionX
'        BoutonsDefilement(1).PositionY = BoutonDefilement2PositionY
'        If DefilementObjetDecalageX > 0 Then
'            BoutonsDefilement(0).Visible = True
'            AffBoutonsDefilement.Afficher BoutonsDefilement(0).PositionX + PositionX, _
'                                          BoutonsDefilement(0).PositionY + PositionY, _
'                                          0, 0, SurfBack, True
'        Else
'            BoutonsDefilement(0).Visible = False
'        End If
'        'Debug.Print DefilementDecalageX & " < " & DefilementNombreLigne
'        If DefilementObjetDecalageX < DefilementNombreLigne Then
'            AffBoutonsDefilement.Afficher BoutonsDefilement(1).PositionX + PositionX, _
'                                          BoutonsDefilement(1).PositionY + PositionY, _
'                                          1, 0, SurfBack, True
'            BoutonsDefilement(1).Visible = True
'        Else
'            BoutonsDefilement(1).Visible = False
'        End If
'    End If
'End Sub

'Propiétés du bouton Deplacer.
Public Property Get BoutonDeplacer_PositionX() As Integer
     BoutonDeplacer_PositionX = BoutonDeplacer.PositionX + PositionX
End Property
Public Property Get BoutonDeplacer_PositionY() As Integer
    BoutonDeplacer_PositionY = BoutonDeplacer.PositionY + PositionY
End Property
Public Property Get BoutonDeplacer_Largeur() As Integer
     BoutonDeplacer_Largeur = AffBoutonAction.Largeur
End Property
Public Property Get BoutonDeplacer_Hauteur() As Integer
     BoutonDeplacer_Hauteur = AffBoutonAction.Hauteur
End Property
Public Sub BoutonDeplacer_Afficher_InfoBulle(ByVal X As Long, _
                                           ByVal Y As Long, _
                                           ByRef Message As ClsIntMessage, _
                                           ByRef SurfBack As DirectDrawSurface7)
     BoutonDeplacer.Afficher_Info_Bulle X, Y, Message, SurfBack
End Sub

'Propiétés du bouton fermer.
Public Property Get BoutonFermer_PositionX() As Integer
     BoutonFermer_PositionX = BoutonFermer.PositionX + PositionX
End Property
Public Property Get BoutonFermer_PositionY() As Integer
    BoutonFermer_PositionY = BoutonFermer.PositionY + PositionY
End Property
Public Property Get BoutonFermer_Largeur() As Integer
     BoutonFermer_Largeur = AffBoutonAction.Largeur
End Property
Public Property Get BoutonFermer_Hauteur() As Integer
     BoutonFermer_Hauteur = AffBoutonAction.Hauteur
End Property
Public Property Get BoutonFermer_Enabled() As Boolean
     BoutonFermer_Enabled = BoutonFermer.Enabled
End Property
Public Sub BoutonFermer_Afficher_InfoBulle(ByVal X As Long, _
                                           ByVal Y As Long, _
                                           ByRef Message As ClsIntMessage, _
                                           ByRef SurfBack As DirectDrawSurface7)
     BoutonFermer.Afficher_Info_Bulle X, Y, Message, SurfBack
End Sub

'Propriété des boutons manger.
Public Property Get BoutonsMangerIndiceRessources(ByVal Index As Long) As Long
     BoutonsMangerIndiceRessources = IndiceRessourcesMangeable(Index)
End Property
Public Property Get BoutonMangerVisible(ByVal Index As Long) As Boolean
     BoutonMangerVisible = BoutonsManger(Index).Visible
End Property
Public Property Get BoutonMangerPositionX(ByVal Index As Long) As Long
     BoutonMangerPositionX = BoutonsManger(Index).PositionX + PositionX
End Property
Public Property Get BoutonMangerPositionY(ByVal Index As Long) As Long
    BoutonMangerPositionY = BoutonsManger(Index).PositionY + PositionY
End Property
Public Property Get BoutonMangerLargeur(ByVal Index As Long) As Long
     BoutonMangerLargeur = AffBoutonAction.Largeur
End Property
Public Property Get BoutonMangerHauteur(ByVal Index As Long) As Long
     BoutonMangerHauteur = AffBoutonAction.Hauteur
End Property
Public Sub BoutonsManger_Afficher_InfoBulle(ByVal Index As Long, _
                                            ByVal X As Long, _
                                            ByVal Y As Long, _
                                            ByVal Perso As ClsJeuPerso, _
                                            ByVal Parametres As ClsJeuParam, _
                                            ByRef Message As ClsIntMessage, _
                                            ByRef SurfBack As DirectDrawSurface7)
    'Affiche l'info bulle.
    Dim i As Integer
    Dim InfoBulle As String
    i = IndiceRessourcesMangeable(Index)
    With Parametres
    InfoBulle = .EtiquetteManger & " : " & LCase(.Ressources_Nom(i)) & "  "
    If .Ressources_GainVie(i) > 0 Then
        InfoBulle = InfoBulle & Chr(13) & _
            "+" & Format(.Ressources_GainVie(i) * Perso.Manger_Ressources_Coef(0), BoutonMangerFormat) & " " & .EtiquetteVie
    End If
    If .Ressources_GainEnergie(i) > 0 Then
        InfoBulle = InfoBulle & Chr(13) & _
            "+" & Format(.Ressources_GainEnergie(i) * Perso.Manger_Ressources_Coef(1), BoutonMangerFormat) & " " & .EtiquetteEnergie
    End If
    If .Ressources_GainMagie(i) > 0 Then
        InfoBulle = InfoBulle & Chr(13) & _
            "+" & Format(.Ressources_GainMagie(i) * Perso.Manger_Ressources_Coef(2), BoutonMangerFormat) & " " & .EtiquetteMagie
    End If
    If .Ressources_GainMoral(i) > 0 Then
        InfoBulle = InfoBulle & Chr(13) & _
            "+" & Format(.Ressources_GainMoral(i) * Perso.Manger_Ressources_Coef(3), BoutonMangerFormat) & " " & .EtiquetteMoral
    End If
    End With
    With BoutonsManger(Index)
        .InfoBulle = InfoBulle
        .Afficher_Info_Bulle X, Y, Message, SurfBack
    End With
End Sub

'Propriétés du bouton "Empereur".
Public Property Get BoutonEmpereur_PositionX() As Long
     BoutonEmpereur_PositionX = BoutonEmpereurPositionX + PositionX
End Property
Public Property Get BoutonEmpereur_PositionY() As Long
    BoutonEmpereur_PositionY = BoutonEmpereurPositionY + PositionY
End Property
Public Property Get BoutonEmpereur_Largeur() As Long
    BoutonEmpereur_Largeur = BoutonEmpereurLargeur
End Property
Public Property Get BoutonEmpereur_Hauteur() As Long
    BoutonEmpereur_Hauteur = BoutonEmpereurHauteur
End Property
Public Property Get BoutonEmpereur_InfoBulle(ByVal Index As Integer) As String
    If Index = 0 Then
        BoutonEmpereur_InfoBulle = BoutonEmpereurInfoBullePrefixe
    Else
        BoutonEmpereur_InfoBulle = BoutonEmpereurInfoBulleSuffixe
    End If
End Property
'Propriétés des boutons de restauration.
Public Property Get Nombre_BoutonsRestauration() As Long
     Nombre_BoutonsRestauration = NombreBoutonsRestauration
End Property
Public Property Get BoutonRestaurationPositionX(ByVal Index As Long) As Long
     BoutonRestaurationPositionX = BoutonsRestaurer(Index).PositionX + PositionX
End Property
Public Property Get BoutonRestaurationPositionY(ByVal Index As Long) As Long
    BoutonRestaurationPositionY = BoutonsRestaurer(Index).PositionY + PositionY
End Property
Public Property Get BoutonRestaurationLargeur(ByVal Index As Long) As Long
     BoutonRestaurationLargeur = AffBoutonAction.Largeur
End Property
Public Property Get BoutonRestaurationHauteur(ByVal Index As Long) As Long
     BoutonRestaurationHauteur = AffBoutonAction.Hauteur
End Property
Public Sub BoutonRestauration_Afficher_InfoBulle(ByVal Index As Long, _
                                                 ByVal X As Long, _
                                                 ByVal Y As Long, _
                                                 ByRef Message As ClsIntMessage, _
                                                 ByRef SurfBack As DirectDrawSurface7)
     BoutonsRestaurer(Index).Afficher_Info_Bulle X, Y, Message, SurfBack
End Sub

'Propriété des boutons de commerce.
Public Property Get NombreColonnesBoutonsCommerce() As Long
     NombreColonnesBoutonsCommerce = NombreBoutonsCommerce
End Property
Public Property Get BoutonCommercePositionX(ByVal Index As Long) As Long
     BoutonCommercePositionX = BoutonsCommerce(Index).PositionX + PositionX
End Property
Public Property Get BoutonCommercePositionY(ByVal Index As Long) As Long
    BoutonCommercePositionY = BoutonsCommerce(Index).PositionY + PositionY
End Property
Public Property Get BoutonCommerceLargeur(ByVal Index As Long) As Long
     BoutonCommerceLargeur = AffBoutonsCommerce.Largeur
End Property
Public Property Get BoutonCommerceHauteur(ByVal Index As Long) As Long
     BoutonCommerceHauteur = AffBoutonsCommerce.Hauteur
End Property
Public Sub BoutonCommerce_Afficher_InfoBulle(ByVal Index As Long, _
                                             ByVal Perso As ClsJeuPerso, _
                                             ByVal Maison As ClsJeuBatiment, _
                                             ByVal Fief As ClsJeuFief, _
                                             ByVal X As Long, _
                                             ByVal Y As Long, _
                                             ByRef Message As ClsIntMessage, _
                                             ByRef SurfBack As DirectDrawSurface7, _
                                             ByVal Achat As Boolean)
    'Dim Temp1 As String
    Dim Temp2 As String
    Dim IndiceRessource As Integer
    Dim PrixRessource As Long
    Dim QuantiteRessource As Integer
    With BoutonsCommerce(Index)
    Select Case Index - ((Parametres.NombreRessources - 1) * NombreColonnesBoutonsCommerce + NombreColonnesBoutonsCommerce)
    Case 2:
        If Perso.DansUnChateau Or Perso.DansSonChateau Or Parametres.Batiment_Marche(Maison.TypeBatiment) Then
            .InfoBulle = BoutonsRessourcesTotalEtiquette1
        Else
            .InfoBulle = BoutonsRessourcesTotalEtiquette4
        End If
    Case 1:
        If Perso.DansUnChateau Or Perso.DansSonChateau Or Parametres.Batiment_Marche(Maison.TypeBatiment) Then
            .InfoBulle = BoutonsRessourcesTotalEtiquette2
        Else
            .InfoBulle = BoutonsRessourcesTotalEtiquette5
        End If
    Case 0:
        If Perso.DansUnChateau Or Perso.DansSonChateau Or Parametres.Batiment_Marche(Maison.TypeBatiment) Then
            .InfoBulle = BoutonsRessourcesTotalEtiquette3
        Else
            .InfoBulle = BoutonsRessourcesTotalEtiquette6
        End If
    Case Else:
        IndiceRessource = Int(Index / NombreColonnesBoutonsCommerce)
        Select Case Index Mod NombreBoutonsCommerce
        Case 0: 'Temp1 = "1"
                QuantiteRessource = 1
        Case 1: 'Temp1 = "5"
                QuantiteRessource = 5
        Case 2: 'Temp1 = Parametres.EtiquetteMaximumAbr
                If Achat Then
                    QuantiteRessource = Perso.Definir_MaxRessources(IndiceRessource) - Perso.Definir_ressources(IndiceRessource)
                Else
                    QuantiteRessource = Perso.Definir_ressources(IndiceRessource)
                End If
        End Select
        If Achat Then
            If Perso.DansSonChateau Or Perso.DansUnChateau Or Parametres.Batiment_Marche(Maison.TypeBatiment) Then
                If QuantiteRessource > Int(Fief.Definir_ressources(IndiceRessource)) Then QuantiteRessource = Int(Fief.Definir_ressources(IndiceRessource))
            Else
                If QuantiteRessource > Int(Maison.Definir_Stock) Then QuantiteRessource = Int(Maison.Definir_Stock)
            End If
        Else
            If QuantiteRessource > Int(Perso.Definir_ressources(IndiceRessource)) Then QuantiteRessource = Int(Perso.Definir_ressources(IndiceRessource))
        End If
        PrixRessource = QuantiteRessource * Parametres.Ressources_PrixVente(IndiceRessource) * IIf(Achat, Parametres.CoefficientAchat, 1)
        Temp2 = Parametres.Ressources_Nom(IndiceRessource)
        If Perso.DansUnChateau Or Perso.DansSonChateau Or Perso.DansUneMaison Then
            If Achat Then
                .InfoBulle = Parametres.EtiquetteObjetAcheter
            Else
                .InfoBulle = Parametres.EtiquetteObjetVente
            End If
        Else
            .InfoBulle = RessourcesBoutonsPrendreEtiquette & " "
        End If
        .InfoBulle = .InfoBulle & QuantiteRessource & " " & Temp2 & "  #" & IIf(Achat And PrixRessource > Perso.Definir_Argent, "&", "") & Parametres.EtiquettePrix & " : " & PrixRessource
    End Select
    .Afficher_Info_Bulle X, Y, Message, SurfBack
    End With
End Sub

'Propriété du bouton pour prendre de l'argent.
Public Property Get BoutonArgentBoutonsPrendre_PositionX() As Integer
     BoutonArgentBoutonsPrendre_PositionX = ArgentBoutonsPrendrePositionX + PositionX
End Property
Public Property Get BoutonArgentBoutonsPrendre_PositionY() As Integer
    BoutonArgentBoutonsPrendre_PositionY = ArgentBoutonsPrendrePositionY + PositionY
End Property
Public Property Get BoutonArgentBoutonsPrendre_Largeur() As Integer
     BoutonArgentBoutonsPrendre_Largeur = ArgentBoutonsPrendreLargeur
End Property
Public Property Get BoutonArgentBoutonsPrendre_Hauteur() As Integer
     BoutonArgentBoutonsPrendre_Hauteur = ArgentBoutonsPrendreHauteur
End Property
Public Sub BoutonArgentBoutonsPrendreAfficher_Info_Bulle(ByVal X As Long, _
                                                         ByVal Y As Long, _
                                                         ByRef Message As ClsIntMessage, _
                                                         ByRef SurfBack As DirectDrawSurface7)
     BoutonArgentPrendre.Afficher_Info_Bulle X, Y, Message, SurfBack
End Sub
'Propriétés des boutons pour prendre des objets ressources.
Public Property Get BoutonsRessourcesPrendrePositionX(ByVal Index As Long) As Integer
     BoutonsRessourcesPrendrePositionX = BoutonsRessourcesPrendre(Index).PositionX + PositionX
End Property
Public Property Get BoutonsRessourcesPrendrePositionY(ByVal Index As Long) As Integer
    BoutonsRessourcesPrendrePositionY = BoutonsRessourcesPrendre(Index).PositionY + PositionY
End Property
Public Property Get BoutonsRessourcesPrendreLargeur(ByVal Index As Long) As Integer
     BoutonsRessourcesPrendreLargeur = RessourcesBoutonsPrendreLargeur
End Property
Public Property Get BoutonsRessourcesPrendreHauteur(ByVal Index As Long) As Integer
     BoutonsRessourcesPrendreHauteur = RessourcesBoutonsPrendreHauteur
End Property
Public Sub BoutonsRessourcesPrendre_Afficher_InfoBulle(ByVal Index As Long, _
                                                       ByVal X As Long, _
                                                       ByVal Y As Long, _
                                                       ByRef Message As ClsIntMessage, _
                                                       ByRef SurfBack As DirectDrawSurface7)
     BoutonsRessourcesPrendre(Index).Afficher_Info_Bulle X, Y, Message, SurfBack
End Sub

'Propriété des boutons de transfert.
Public Property Get NombreColonnesBoutonsTransfert() As Long
     NombreColonnesBoutonsTransfert = NombreBoutonsTransfert
End Property
Public Property Get BoutonTransfertArgentPersoPositionX(ByVal Index As Long) As Long
     BoutonTransfertArgentPersoPositionX = BoutonsTransfertArgentPerso(Index).PositionX + PositionX
End Property
Public Property Get BoutonTransfertArgentPersoPositionY(ByVal Index As Long) As Long
    BoutonTransfertArgentPersoPositionY = BoutonsTransfertArgentPerso(Index).PositionY + PositionY
End Property
Public Sub BoutonTransfertArgentPerso_Afficher_InfoBulle(ByVal Index As Long, _
                                                         ByVal X As Long, _
                                                         ByVal Y As Long, _
                                                         ByRef Message As ClsIntMessage, _
                                                         ByRef SurfBack As DirectDrawSurface7)
     BoutonsTransfertArgentPerso(Index).Afficher_Info_Bulle X, Y, Message, SurfBack
End Sub
Public Property Get BoutonTransfertArgentMaisonPositionX(ByVal Index As Long) As Long
     BoutonTransfertArgentMaisonPositionX = BoutonsTransfertArgentMaison(Index).PositionX + PositionX
End Property
Public Property Get BoutonTransfertArgentMaisonPositionY(ByVal Index As Long) As Long
    BoutonTransfertArgentMaisonPositionY = BoutonsTransfertArgentMaison(Index).PositionY + PositionY
End Property
Public Sub BoutonTransfertArgentMaison_Afficher_InfoBulle(ByVal Index As Long, _
                                                          ByVal X As Long, _
                                                          ByVal Y As Long, _
                                                          ByRef Message As ClsIntMessage, _
                                                          ByRef SurfBack As DirectDrawSurface7)
     BoutonsTransfertArgentMaison(Index).Afficher_Info_Bulle X, Y, Message, SurfBack
End Sub

Public Property Get BoutonTransfertStockPersoPositionX(ByVal Index As Long) As Long
     BoutonTransfertStockPersoPositionX = BoutonsTransfertStockPerso(Index).PositionX + PositionX
End Property
Public Property Get BoutonTransfertStockPersoPositionY(ByVal Index As Long) As Long
    BoutonTransfertStockPersoPositionY = BoutonsTransfertStockPerso(Index).PositionY + PositionY
End Property
Public Sub BoutonTransfertStockPerso_Afficher_InfoBulle(ByVal Index As Long, _
                                                        ByVal X As Long, _
                                                        ByVal Y As Long, _
                                                        ByRef Message As ClsIntMessage, _
                                                        ByRef SurfBack As DirectDrawSurface7)
     BoutonsTransfertStockPerso(Index).Afficher_Info_Bulle X, Y, Message, SurfBack
End Sub

Public Property Get BoutonTransfertStockMaisonPositionX(ByVal Index As Long) As Long
     BoutonTransfertStockMaisonPositionX = BoutonsTransfertStockMaison(Index).PositionX + PositionX
End Property
Public Property Get BoutonTransfertStockMaisonPositionY(ByVal Index As Long) As Long
    BoutonTransfertStockMaisonPositionY = BoutonsTransfertStockMaison(Index).PositionY + PositionY
End Property
Public Sub BoutonTransfertStockMaison_Afficher_InfoBulle(ByVal Index As Long, _
                                                         ByVal X As Long, _
                                                         ByVal Y As Long, _
                                                         ByRef Message As ClsIntMessage, _
                                                         ByRef SurfBack As DirectDrawSurface7)
     BoutonsTransfertStockMaison(Index).Afficher_Info_Bulle X, Y, Message, SurfBack
End Sub

Public Property Get BoutonTransfertLargeur() As Long
     BoutonTransfertLargeur = AffBoutonsTransfert.Largeur
End Property
Public Property Get BoutonTransfertHauteur() As Long
     BoutonTransfertHauteur = AffBoutonsTransfert.Hauteur
End Property

'Bouton de service.
Public Property Get BoutonServicePositionX() As Integer
    BoutonServicePositionX = ServiceBoutonPositionX + PositionX
End Property
Public Property Get BoutonServicePositionY() As Integer
    BoutonServicePositionY = ServiceBoutonPositionY + PositionY
End Property
Public Property Get BoutonServiceLargeur() As Long
    BoutonServiceLargeur = AffBoutonAction.Largeur
End Property
Public Property Get BoutonServiceHauteur() As Long
    BoutonServiceHauteur = AffBoutonAction.Hauteur
End Property
Public Sub BoutonService_Afficher_InfoBulle(ByVal Index As Long, _
                                            ByVal X As Long, _
                                            ByVal Y As Long, _
                                            ByRef Message As ClsIntMessage, _
                                            ByRef SurfBack As DirectDrawSurface7)
     BoutonService.Afficher_Info_Bulle X, Y, Message, SurfBack
End Sub

'Boutons de fabrication.
Public Property Get Nombre_Boutons_Fabriquer() As Integer
    Nombre_Boutons_Fabriquer = UBound(BoutonsFabriquer()) + 1
End Property
Public Property Get BoutonsFabriquerActif(ByVal Index As Integer) As Boolean
    BoutonsFabriquerActif = BoutonsFabriquer(Index).Enabled
End Property
Public Property Get BoutonsFabriquerPositionX(ByVal Index As Integer) As Long
    BoutonsFabriquerPositionX = BoutonsFabriquer(Index).PositionX + PositionX
End Property
Public Property Get BoutonsFabriquerPositionY(ByVal Index As Integer) As Long
    BoutonsFabriquerPositionY = BoutonsFabriquer(Index).PositionY + PositionY
End Property
Public Property Get BoutonsFabriquerLargeur() As Long
    BoutonsFabriquerLargeur = AffBoutonAction.Largeur
End Property
Public Property Get BoutonsFabriquerHauteur() As Long
    BoutonsFabriquerHauteur = AffBoutonAction.Hauteur
End Property
Public Sub BoutonFabriquer_Afficher_InfoBulle(ByVal Index As Long, _
                                              ByVal X As Long, _
                                              ByVal Y As Long, _
                                              ByRef Message As ClsIntMessage, _
                                              ByRef SurfBack As DirectDrawSurface7)
     BoutonsFabriquer(Index).Afficher_Info_Bulle X, Y, Message, SurfBack
End Sub

'Boutons des objets.
Public Property Get Nombre_Boutons_Objets() As Integer
    Nombre_Boutons_Objets = UBound(BoutonsObjets()) + 1
End Property
Public Property Get BoutonsObjetsActif(ByVal Index As Integer) As Boolean
    BoutonsObjetsActif = BoutonsObjets(Index).Enabled
End Property
Public Property Get BoutonsObjetsPositionX(ByVal Index As Integer) As Long
    BoutonsObjetsPositionX = BoutonsObjets(Index).PositionX + PositionX
End Property
Public Property Get BoutonsObjetsPositionY(ByVal Index As Integer) As Long
    BoutonsObjetsPositionY = BoutonsObjets(Index).PositionY + PositionY
End Property
Public Property Get BoutonsObjetsFabriquerLargeur(ByRef AffObjets As ClsAffObjets) As Long
    BoutonsObjetsFabriquerLargeur = AffObjets.Largeur * FabricationIconeZoom
End Property
Public Property Get BoutonsObjetsFabriquerHauteur(ByRef AffObjets As ClsAffObjets) As Long
    BoutonsObjetsFabriquerHauteur = AffObjets.Hauteur * FabricationIconeZoom
End Property
Public Property Get BoutonsObjetsLargeur(ByRef AffObjets As ClsAffObjets) As Long
    BoutonsObjetsLargeur = AffObjets.Largeur
End Property
Public Property Get BoutonsObjetsHauteur(ByRef AffObjets As ClsAffObjets) As Long
    BoutonsObjetsHauteur = AffObjets.Hauteur
End Property
Public Sub BoutonObjets_Afficher_InfoBulle(ByVal Index As Long, _
                                           ByVal Perso As ClsJeuPerso, _
                                           ByVal Maison As ClsJeuBatiment, _
                                           ByVal Parametres As ClsJeuParam, _
                                           ByVal X As Long, _
                                           ByVal Y As Long, _
                                           ByRef Message As ClsIntMessage, _
                                           ByRef AffObjets As ClsAffObjets, _
                                           ByRef SurfBack As DirectDrawSurface7, _
                                           ByVal Inventaire As Boolean, _
                                           Optional ByVal CategorieObjet As Integer = -1, _
                                           Optional ByVal Duree As Long = 0, _
                                           Optional ByVal NumeroEffet As Integer = 0)
    'Dim CategorieObjet As Integer
    If Duree <> 0 Then
        BoutonsObjets(Index).InfoBulle = Definir_InfoBulle_Objet(CategorieObjet, Perso, Parametres, False, False, Duree, NumeroEffet)
    Else
        If Inventaire Then
            If CategorieObjet = -1 Then
                If Index >= Perso.Nombre_Objets_Inventaire Then
                    BoutonsObjets(Index).InfoBulle = Definir_InfoBulle_Objet(Perso.Objet_Equipes_Apparence(Index - Perso.Nombre_Objets_Inventaire), Perso, Parametres, True, True)
                    CategorieObjet = Perso.Objet_Equipes_Categorie(Index - Perso.Nombre_Objets_Inventaire)
                Else
                    BoutonsObjets(Index).InfoBulle = Definir_InfoBulle_Objet(Perso.Objet_Inventaire_Apparence(Index), Perso, Parametres, True, False, , , Perso.Objet_Inventaire_Nombre(Index))
                    CategorieObjet = Perso.Objet_Inventaire_Categorie(Index)
                    'Encadrer_Objet Index, AffObjets, SurfBack, True
                End If
            End If
            'Encadre l'emplacement ou l'objet peut être placé.
            Select Case CategorieObjet
            Case Main1, Main2:
                Encadrer_Objet Main1, AffObjets, SurfBack
                Encadrer_Objet Main2, AffObjets, SurfBack
            Case Anneau1, Anneau2:
                Encadrer_Objet Anneau1, AffObjets, SurfBack
                Encadrer_Objet Anneau2, AffObjets, SurfBack
            Case Else
                Encadrer_Objet CategorieObjet, AffObjets, SurfBack
            End Select
        Else
            BoutonsObjets(Index).InfoBulle = Definir_InfoBulle_Objet(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, Index), Perso, Parametres, False, False)
        End If
    End If
    BoutonsObjets(Index).Afficher_Info_Bulle X, Y, Message, SurfBack
End Sub

Public Sub Encadrer_Objet(ByVal CategorieObjet As Integer, _
                          ByRef AffObjets As ClsAffObjets, _
                          ByRef SurfBack As DirectDrawSurface7, _
                          Optional ByVal Inventaire As Boolean)
    SurfBack.SetForeColor EquipementContourCouleur
    SurfBack.SetFillStyle 1
    If Inventaire Then
        SurfBack.DrawBox BoutonsObjets(CategorieObjet).PositionX + PositionX, _
                         BoutonsObjets(CategorieObjet).PositionY + PositionY, _
                         BoutonsObjets(CategorieObjet).PositionX + AffObjets.Largeur + PositionX, _
                         BoutonsObjets(CategorieObjet).PositionY + AffObjets.Hauteur + PositionY
    Else
        SurfBack.DrawBox InventaireEquipementObjetX(CategorieObjet) - EquipementContourEpaisseur + PositionX, _
                         InventaireEquipementObjetY(CategorieObjet) - EquipementContourEpaisseur + PositionY, _
                         InventaireEquipementObjetX(CategorieObjet) + AffObjets.Largeur + EquipementContourEpaisseur + PositionX, _
                         InventaireEquipementObjetY(CategorieObjet) + AffObjets.Hauteur + EquipementContourEpaisseur + PositionY
    End If
End Sub

'Boutons de défilement.
Public Property Get Nombre_Lignes_Ressources() As Integer
    Nombre_Lignes_Ressources = NombreLignesRessources
End Property
Public Property Get Nombre_Lignes_Objets() As Integer
    Nombre_Lignes_Objets = NombreLignesObjets
End Property
Public Property Get Nombre_Boutons_Defilement() As Integer
    Nombre_Boutons_Defilement = NombreBoutonsDefilement
End Property
Public Property Get BoutonsDefilementPositionX(ByVal Index As Integer) As Long
    BoutonsDefilementPositionX = BoutonsDefilement(Index).PositionX + PositionX
End Property
Public Property Get BoutonsDefilementPositionY(ByVal Index As Integer) As Long
    BoutonsDefilementPositionY = BoutonsDefilement(Index).PositionY + PositionY
End Property
Public Property Get BoutonsDefilementLargeur(ByVal Index As Integer) As Long
    If Index = 2 Then
        BoutonsDefilementLargeur = AffBoutonsBarreDefilement.Largeur
    Else
        BoutonsDefilementLargeur = AffBoutonsDefilement.Largeur
    End If
End Property
Public Property Get BoutonsDefilementHauteur(ByVal Index As Integer) As Long
    If Index = 2 Then
        BoutonsDefilementHauteur = AffBoutonsBarreDefilement.Hauteur
    Else
        BoutonsDefilementHauteur = AffBoutonsDefilement.Hauteur
    End If
End Property
Public Property Get BoutonsDefilementVisible(ByVal Index As Long) As Boolean
    BoutonsDefilementVisible = BoutonsDefilement(Index).Visible
End Property
Public Sub BoutonDefilement_Afficher_InfoBulle(ByVal Index As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               ByRef Message As ClsIntMessage, _
                                               ByRef SurfBack As DirectDrawSurface7)
     BoutonsDefilement(Index).Afficher_Info_Bulle X, Y, Message, SurfBack
End Sub

Public Property Get BoutonInventaire_NombreCases() As Integer
    BoutonInventaire_NombreCases = InventaireEquipementNombreCases
End Property
Public Property Get BoutonInventaireMaison_PositionX() As Integer
    BoutonInventaireMaison_PositionX = BoutonInventaireMaison.PositionX + PositionX
End Property
Public Property Get BoutonInventaireMaison_PositionY() As Integer
    BoutonInventaireMaison_PositionY = BoutonInventaireMaison.PositionY + PositionY
End Property
Public Property Get BoutonInventaireMaison_Largeur() As Integer
    BoutonInventaireMaison_Largeur = AffBoutonAction.Largeur
End Property
Public Property Get BoutonInventaireMaison_Hauteur() As Integer
    BoutonInventaireMaison_Hauteur = AffBoutonAction.Hauteur
End Property
Public Property Get Nombre_Boutons_InventaireOnglets() As Integer
    Nombre_Boutons_InventaireOnglets = NombreBoutonsInventaireOnglets
End Property
Public Property Get BoutonInventaireOnglet_PositionX(ByVal Index As Integer) As Integer
    BoutonInventaireOnglet_PositionX = BoutonInventaireOnglets(Index).PositionX + PositionX
End Property
Public Property Get BoutonInventaireOnglet_PositionY(ByVal Index As Integer) As Integer
    BoutonInventaireOnglet_PositionY = BoutonInventaireOnglets(Index).PositionY + PositionY
End Property
Public Property Get BoutonInventaireOnglet_Largeur() As Integer
    BoutonInventaireOnglet_Largeur = AffBoutonAction.Largeur
End Property
Public Property Get BoutonInventaireOnglet_Hauteur() As Integer
    BoutonInventaireOnglet_Hauteur = AffBoutonAction.Hauteur
End Property
Public Sub BoutonInventaireMaison_Afficher_InfoBulle(ByVal Objet As Boolean, _
                                                     ByVal Perso As ClsJeuPerso, _
                                                     ByVal X As Long, _
                                                     ByVal Y As Long, _
                                                     ByRef Message As ClsIntMessage, _
                                                     ByRef SurfBack As DirectDrawSurface7)
    If Bascule = 0 Or Not Perso.DansSaMaison Then
        If Objet Then
            BoutonInventaireMaison.InfoBulle = InventaireMaisonMessagePersoObjet
        Else
            BoutonInventaireMaison.InfoBulle = InventaireMaisonMessagePerso
        End If
    Else
        If Objet Then
            BoutonInventaireMaison.InfoBulle = InventaireMaisonMessageMaisonObjet
        Else
            BoutonInventaireMaison.InfoBulle = InventaireMaisonMessageMaison
        End If
    End If
    BoutonInventaireMaison.Afficher_Info_Bulle X, Y, Message, SurfBack
End Sub

'Propriétés de la carte.
Public Property Get Definir_Carte_PositionX() As Integer
     Definir_Carte_PositionX = CartePositionXDepart + CarteDecalageX
End Property
Public Property Get Definir_Carte_PositionY() As Integer
    Definir_Carte_PositionY = CartePositionYDepart + CarteDecalageY
End Property
Public Property Get Definir_Carte_Largeur() As Integer
    Definir_Carte_Largeur = CarteLargeur
End Property
Public Property Get Definir_Carte_Hauteur() As Integer
    Definir_Carte_Hauteur = CarteHauteur
End Property

'Propriétés de la fenetre personnage.
Public Property Get BoutonsCaracNombre() As Integer
    BoutonsCaracNombre = 1
End Property
Public Property Get BoutonsCaracPositionX(ByVal Index As Long) As Integer
    Select Case Index
        Case 0: BoutonsCaracPositionX = CaracResistancesPositionX + PositionX
    End Select
End Property
Public Property Get BoutonsCaracPositionY(ByVal Index As Long) As Integer
    Select Case Index
        Case 0: BoutonsCaracPositionY = CaracResistancesPositionY + PositionY
    End Select
End Property
Public Property Get BoutonsCaracLargeur(ByVal Index As Long) As Integer
     Select Case Index
        Case 0: BoutonsCaracLargeur = CaracResistancesLargeur
     End Select
End Property
Public Property Get BoutonsCaracHauteur(ByVal Index As Long) As Integer
     Select Case Index
        Case 0: BoutonsCaracHauteur = CaracResistancesHauteur
     End Select
End Property
Public Function BoutonsCaracInfoBulle(ByVal Index As Integer, _
                                      ByVal Perso As ClsJeuPerso, _
                                      ByVal Parametres As ClsJeuParam) As String
    Dim Temp As String
    Dim i As Integer
    With Parametres
    Select Case Index
    Case 0:
        Temp = .Resistance_Attaque(Perso.Definir_AttaqueResistance)
        For i = 0 To .NombreResistances - 1
            If Perso.Definir_Resistance(i) <> Parametres.ResistanceDefaut Then
                Temp = Temp & "#"
                If Perso.Definir_Resistance(i) > Parametres.ResistanceDefaut Then Temp = Temp & "&"
                Temp = Temp & .Resistance_NomComplet(i) & " : "
                If Perso.Definir_Resistance(i) = Parametres.ResistanceMinimum Then
                    Temp = Temp & UCase(Parametres.EtiquetteMaximumAbr)
                ElseIf Perso.Definir_Resistance(i) = Parametres.ResistanceMaximum Then
                    Temp = Temp & UCase(Parametres.EtiquetteMinimumAbr)
                Else
                    Temp = Temp & IIf(Perso.Definir_Resistance(i) < .ResistanceDefaut, "+", "") & _
                                  Format(-(Perso.Definir_Resistance(i) - 1) * 100, "0") & .EtiquettePourcent
                End If
                'Temp = Temp & IIf(Perso.Definir_ResistanceDepart(i) <> Perso.Definir_Resistance(i), " (" & Perso.Definir_Resistance(i) - Perso.Definir_ResistanceDepart(i) & ")", "")
            End If
        Next i
    End Select
    End With
    BoutonsCaracInfoBulle = Temp
End Function

'Propriété des boutons de compétences.
Public Property Get BoutonsCompetences_Nombre() As Integer
     BoutonsCompetences_Nombre = BoutonsCompetencesNombre
End Property
Public Property Get BoutonsCompetences_PositionX(ByVal Index As Long) As Long
     BoutonsCompetences_PositionX = BoutonsCompetences(Index).PositionX + PositionX
End Property
Public Property Get BoutonsCompetences_PositionY(ByVal Index As Long) As Long
    BoutonsCompetences_PositionY = BoutonsCompetences(Index).PositionY + PositionY
End Property
Public Property Get BoutonsCompetences_Largeur(ByVal Index As Long) As Long
     BoutonsCompetences_Largeur = AffBoutonsCompetences.Largeur
End Property
Public Property Get BoutonsCompetences_Hauteur(ByVal Index As Long) As Long
     BoutonsCompetences_Hauteur = AffBoutonsCompetences.Hauteur
End Property
Public Sub BoutonsCompetences_Afficher_InfoBulle(ByVal Index As Long, _
                                                 ByVal X As Long, _
                                                 ByVal Y As Long, _
                                                 ByRef Message As ClsIntMessage, _
                                                 ByRef SurfBack As DirectDrawSurface7)
    BoutonsCompetences(Index).Afficher_Info_Bulle X, Y, Message, SurfBack
End Sub
