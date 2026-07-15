VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsInterface"
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
DefLng A-Z
Option Explicit

Const NomFichierINI = "Interface"
Const NomSectionINI = "Interface"

Const NomFichierINIFenetres = "Fenetres"
Const NomSectionINIFenetres = "Fenetres"
Const NomSectionINIFenetre = "Fenetre_"

Private FichierINI As String
Private SectionINI As String

Private TitreFenetreInventaire As String
Private TitreFenetreAchat As String
Private TitreFenetreVente As String
Private TitreFenetreRessources As String
Private TitreFenetreCaracteristiques As String
Private TitreFenetreCompetences As String
Private TitreFenetreBestiaire As String
Private TitreFenetreCarte As String
Private TitreFenetreObjectifs As String

Public RectangleSelectionMinX As Integer
Public RectangleSelectionMinY As Integer

Public AfficherCadre As Boolean
Private BordureGauche As Integer
Private BordureDroite As Integer
Private BordureHaut As Integer
Private BordureBas As Integer

Public NombreFenetres As Integer
Private LargeurFenetres As Integer
Private HauteurFenetres As Integer
Private FenetrePositionX() As Integer
Private FenetrePositionY() As Integer
Private FenetrePositionXInit() As Integer
Private FenetrePositionYInit() As Integer

Private AfficherEpoques As Boolean
Private EpoqueCouleur As Long
Private EpoquePositionX As Integer
Private EpoquePositionY As Integer

'Frames par secondes.
Private AfficherFPS As Boolean
Private FPSCouleur As Long
Private FPSMessage As String
Private FPSPositionX As Integer
Private FPSPositionY As Integer

'Barres de statuts.
Private AfficherAction As Boolean
Private MessageActionPositionX As Integer
Private MessageActionPositionY As Integer
Private MessageActionCouleur As Long
Private ViePositionX As Integer
Private ViePositionY As Integer
Private EnergiePositionX As Integer
Private EnergiePositionY As Integer
Private MagiePositionX As Integer
Private MagiePositionY As Integer
Private MoralPositionX As Integer
Private MoralPositionY As Integer
Private MessagesActions() As String

Public AfficherBarresStatut As Boolean
Private AfficherBarresStatutScore As Boolean
Private BarresStatutScoreCouleur As Long
Private BarresStatutScoreDecalageX As Integer
Private BarresStatutScoreDecalageY As Integer
Private BarresStatutScoreLargeurTexte As Single
Public NombreBoutonsBarresStatut As Integer
Private BarresStatutDecalageX As Integer
Private BarresStatutDecalageY As Integer
Private BarresStatutLargeur As Integer
Private BarresStatutHauteur As Integer
Private BarresStatutContourEpaisseur  As Integer
Public BarresStatutContourCouleur1 As Long 'Gris.
Public BarresStatutContourCouleur2 As Long 'Orange.
Private BarresStatutFormatCarac As String
Private BarresStatutSeuilMessage1 As String
Private BarresStatutSeuilMessage2 As String

Public AfficherBoutonsConstruire As Boolean
Private BoutonDemolirEtiquette As String
Private BoutonConstruirePositionX As Integer
Private BoutonConstruirePositionY As Integer
Private BoutonConstruireIntervalX As Integer
Private BoutonConstruireIntervalY As Integer
Private BoutonConstruireNombreColonnes As Integer
Private BoutonConstruireColonneIntervalX As Integer
Private BoutonConstruireColonneIntervalY As Integer
'Paramètres pour masquer automatiquement les boutons de construction.
Private MasquerAutoBoutonsConstruire  As Boolean
Private MasquerAutoBoutonsConstruireDecalageX  As Integer
Private MasquerAutoBoutonsConstruireDecalageY  As Integer
Private MasquerAutoBoutonsConstruireDuree  As Integer
'Variables servant à compter un certain temps avant d'afficher les informations supplémentaires des infos-bulles.
Const BoutonConstruireCompteurDuree = 2
Const BoutonConstruireCompteurIntervales = 0.3
Private BoutonConstruireCompteurTemps As Double
Private BoutonConstruireCompteurindice1 As Integer
Private BoutonConstruireCompteurindice2 As Integer

Public AfficherBoutonsInfos As Boolean
Private NombreBoutonsInfos As Integer
Private BoutonInfosTransparents As Boolean
Private BoutonInfosPositionX As Integer
Private BoutonInfosPositionY As Integer
Private BoutonInfosIntervalX As Integer
Private BoutonInfosIntervalY As Integer
Private BoutonInfosInfoBulle() As String

Public AfficherObjectifs As Boolean
Private NouveauxObjectifs As Boolean
Private Objectifs As Collection

'Boutons radiaux.
Public ActiverBoutonsRadiaux As Boolean
Const AfficherBoutonsRadiaux = True
Const NombrePanneauxBoutonsRadiaux = 2
Const BoutonsRadiauxInfoRayon = 40 '40 '32
Const NombreBoutonsRadiauxInfo = 7
Const BoutonsRadiauxConstruireRayon = 40 '32
Const NombreBoutonsRadiauxConstruire = 7 '6
Const BoutonsRadiauxConstruireRayonSup = 40 '32
Const NombreBoutonsRadiauxConstruireSup = 7 '6
Private BoutonsRadiauxCategorie As Integer
Private BoutonsRadiauxAngle() As Single
Const BoutonsRadiauxAngleDefaut = -1.57
Private BoutonsRadiauxCentreX As Integer
Private BoutonsRadiauxCentreY As Integer

Private AfficherBoutonsMenu As Boolean
Private NombreBoutonsMenu As Integer
Private BoutonMenuTransparents As Boolean
Private BoutonMenuPositionX As Integer
Private BoutonMenuPositionY As Integer
Private BoutonMenuIntervalX As Integer
Private BoutonMenuIntervalY As Integer
Private BoutonMenuInfoBulle() As String

Public AfficherBoutonsDeplacement As Boolean
Private NombreBoutonsDeplacement As Integer
Private BoutonDeplacementPositionX As Integer
Private BoutonDeplacementPositionY As Integer
Private BoutonDeplacementIntervalX As Integer
Private BoutonDeplacementIntervalY As Integer
Private BoutonDeplacementNombreColonnes As Integer
Private BoutonDeplacementColonneIntervalX As Integer
Private BoutonDeplacementColonneIntervalY As Integer
Private BoutonDeplacementInfoBulle() As String
'Paramètres pour masquer automatiquement les boutons de déplacement.
Private MasquerAutoBoutonsDeplacement As Boolean
Private MasquerAutoBoutonsDeplacementDecalageX As Integer
Private MasquerAutoBoutonsDeplacementDecalageY As Integer
Private MasquerAutoBoutonsDeplacementDuree As Integer

Private NombreBoutonsOrdre As Integer
Private BoutonOrdrePositionX As Integer
Private BoutonOrdrePositionY As Integer
Private BoutonOrdreIntervalX As Integer
Private BoutonOrdreIntervalY As Integer

'Touches de raccourci objets.
Private AfficherRaccourciObjets As Boolean
Public ActualiserRaccourciObjets As Boolean 'Si vrai, on met à jour immédiatement la barre de raccourci objets.
Private NombreRaccourciObjets As Integer
Private RaccourciObjetsTransparent As Boolean
Private BoutonRaccourciObjetEtiquette As String
Private BoutonRaccourciObjetDecalageY As Integer

'Effets temporaires.
Private AfficherEffetsTemporaires As Boolean
Private EffetTemporatireTransparent As Boolean
Private EffetTemporatireDureeEffacement As Integer
Private EffetTemporatirePositionX As Integer
Private EffetTemporatirePositionY As Integer
Private EffetTemporatireIntervalX As Integer
Private EffetTemporatireIntervalY As Integer

'Message du mode auto.
Private AfficherMessageModeAuto As Boolean
Private MessageModeAutoCouleur As Long
Private MessageModeAutoMessage As String
Private MessageModeAutoPositionX As Integer
Private MessageModeAutoPositionY As Integer

'Infos-bulles sur le terrain.
Private AfficherInfoBulleTerrain As Boolean

'Stastistiques.
Private AfficherStastistiques As Boolean
Private CouleurStastistiques As Long
Private StastistiquesPositionX As Integer
Private StastistiquesPositionY As Integer
'Variables temporaires pour les stastistiques.
Private StastistiquesPositionYTemp As Integer
Private StastistiquesMessageTemp As String

Private AffObjets As ClsAffObjets
Private AffRessources As ClsAffObjets
Private AffCadre As ClsAffFenetre
Private AffFenetre() As ClsAffFenetre
Private Fenetres() As ClsIntFenetre
Public FenetreConfirmation As ClsIntConfirmation
Private AffBoutonsConstruire() As ClsAffBouton
Private BoutonsConstruire() As ClsIntBouton
Private BoutonsConstruire_IndiceBatiment() As Integer
Private BoutonsRadiauxConstruire() As ClsIntBouton
Private AffBoutonsInfos As ClsAffBouton
Private AffBoutonsMenu As ClsAffBouton
Private BoutonsInfos() As ClsIntBouton
Private BoutonsRadiauxInfos() As ClsIntBouton
Private BoutonsMenu() As ClsIntBouton
Private AffBoutonsDeplacement As ClsAffBouton
Private BoutonsDeplacement() As ClsIntBouton
Private AffBoutonsRaccourciObjets As ClsAffBouton
Private BoutonsRaccourciObjets() As ClsIntBouton
Private BoutonsRaccourciObjetsMessages() As String
Private BoutonsRaccourciObjetsImages() As Collection 'Indice des images des objets affichés.
Private BoutonsEffetsTemporaires As Collection
Private BoutonsCaracteristiques() As ClsIntBouton
Private BoutonsCaracteristiquesSeuil() As ClsIntBouton

Private AutoRecuperationCarac() As Boolean
Private AutoRecuperationCaracPcent() As Single
Public AutoBatimentService As Boolean
Private FileFabricationObjets As Collection

Public DeplacerFenetre As Integer
'Public AttaquerAmi As Boolean 'Si vrai, le joueur peut attaquer ses alliés.

Public CurseurSurBarreStatut As Boolean 'Vrai, le curseur est sur une barre de statut.
Public CurseurSurObjet As Boolean 'Vrai : le curseur est sur un objet.
Private CurseurSurBoutonConstruire As Boolean  'Vrai : le curseur est sur l'un des boutons construire.
Public ForcerCurseurSurBoutonConstruire As Boolean ' Si Vrai, CurseurSurBoutonConstruire est toujours vrai.
Private CurseurSurBoutonDeplacement As Boolean  'Vrai : le curseur est sur l'un des boutons construire.
Public ForcerCurseurSurBoutonDeplacement As Boolean ' Si Vrai, CurseurSurBoutonConstruire est toujours vrai.
Private Etat() As Integer 'Mémorise l'état des boutons de construction
Private NombreBoutonsConstruire As Integer
Private TouchesRaccourciObjets() As Integer

Private Compteur As Single

'Variables temporaires utilisées pour l'affichage plein écran.
Private AffichagePleinEcran As Boolean
Private TempAfficherCadre As Boolean
Private TempAfficherEpoques As Boolean
Private TempAfficherFPS As Boolean
Private TempAfficherAction As Boolean
Private TempAfficherBarresStatut As Boolean
Private TempAfficherBoutonsConstruire As Boolean
Private TempAfficherBoutonsInfos As Boolean
Private TempAfficherBoutonsDeplacement As Boolean
Private TempAfficherStastistiques As Boolean
Private TempAfficherRaccourciObjets As Boolean
Private TempAfficherEffetsTemp As Boolean

'Effets visuels.
Private Effets As Collection

'Fenêtre de dialogue.
Public AfficherDialogue As Boolean
Public DialogueTexte1 As String
Public DialogueTexte2 As String
Public DialogueCurseurPosition As Integer
Public DialogueCurseurVisible As Boolean
Private MessagesPrecedents As Collection
Private MessagesPrecedentsPosition As Integer
Private MessagesPrecedentsAffiche As Boolean

'Fenêtre de confirmation.
Public MessageConfirmationQuitter As String
Public MessageConfirmationAttaquer As String
Public MessageConfirmationOui As String
Public MessageConfirmationNon As String

'Variables créées pour la liste des compétences.
Public CompetencesTri As Integer
Public CompetencesAfficherToutes As Boolean
Private Competence As ClsJeuCompetences
Private TabComp As Collection

Private Sub Class_Initialize()
    'Initialisation des classes.
    Set AffObjets = New ClsAffObjets
    Set AffRessources = New ClsAffObjets
    Set AffCadre = New ClsAffFenetre
    'Set AffFenetre = New ClsAffFenetre
    'Set AffBoutonsConstruire = New ClsAffBouton
    Set AffBoutonsInfos = New ClsAffBouton
    Set AffBoutonsMenu = New ClsAffBouton
    Set AffBoutonsDeplacement = New ClsAffBouton
    Set AffBoutonsRaccourciObjets = New ClsAffBouton
    Set BoutonsEffetsTemporaires = New Collection
    
    Set Objectifs = New Collection
    Set Effets = New Collection
    Set MessagesPrecedents = New Collection
    
    Set Competence = New ClsJeuCompetences
    Set TabComp = New Collection
    CompetencesTri = 1
    CompetencesAfficherToutes = False
End Sub

Public Function ModeDuoPossible(ByVal Ecran As ClsPerEcran) As Boolean
    On Error GoTo Erreur
    Open App.Path & FicIni.Chemin & NomFichierINI & Ecran.Width & "_J1.ini" For Input As #1
    Close #1
    ModeDuoPossible = True
    Exit Function
Erreur:
    ModeDuoPossible = False
End Function

Public Sub Init1(ByVal NumeroJoueur As Integer, _
                 ByVal NombreJoueurs As Integer)
    Dim i As Long: Dim j As Long
    Dim NombreMessagesActions As Integer
    
    'Chargement des paramètres des fenetres.
    FicIni.Fichier = FicIni.Chemin & NomFichierINIFenetres
    FicIni.Section = NomSectionINIFenetres
    
    NombreFenetres = Val(FicIni.Parametre("NombreFenetres"))
    ReDim AffFenetre(NombreFenetres - 1)
    For i = 0 To NombreFenetres - 1
        FicIni.Section = NomSectionINIFenetre & i + 1
        Set AffFenetre(i) = New ClsAffFenetre
        With AffFenetre(i)
            .Largeur = Val(FicIni.Parametre("Largeur"))
            .Hauteur = Val(FicIni.Parametre("Hauteur"))
            .Definir_Nombre_Sprites = Val(FicIni.Parametre("NombreImages"))
            For j = 0 To .Definir_Nombre_Sprites - 1
                .Definir_Image(j) = FicIni.Parametre("Image" & j + 1)
            Next j
        End With
    Next i
    
    If NombreJoueurs = 1 Then
        FichierINI = NomFichierINI & Ecran.Width
    Else
        FichierINI = NomFichierINI & Ecran.Width & "_J" & NumeroJoueur + 1
    End If
    SectionINI = NomSectionINI
    
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    'Chargement des paramêtres d'un fichier ini.
    'Souris.
    RectangleSelectionMinX = Val(FicIni.Parametre("RectangleSelectionMinX"))
    RectangleSelectionMinY = Val(FicIni.Parametre("RectangleSelectionMinY"))
    
    'Cadre.
    AfficherCadre = CBool(FicIni.Parametre("AfficherCadre"))
    BordureGauche = Val(FicIni.Parametre("BordureGauche"))
    BordureDroite = Val(FicIni.Parametre("BordureDroite"))
    BordureHaut = Val(FicIni.Parametre("BordureHaut"))
    BordureBas = Val(FicIni.Parametre("BordureBas"))
    
    'Fenêtres.
    ReDim FenetrePositionX(NombreFenetres - 1)
    ReDim FenetrePositionY(NombreFenetres - 1)
    ReDim FenetrePositionXInit(NombreFenetres - 1)
    ReDim FenetrePositionYInit(NombreFenetres - 1)
    For i = 0 To NombreFenetres - 1
        FenetrePositionX(i) = Val(FicIni.Parametre("Fenetre" & i + 1 & "PositionX"))
        FenetrePositionY(i) = Val(FicIni.Parametre("Fenetre" & i + 1 & "PositionY"))
        FenetrePositionXInit(i) = FenetrePositionX(i)
        FenetrePositionYInit(i) = FenetrePositionY(i)
    Next i
    
    'Epoques.
    AfficherEpoques = CBool(FicIni.Parametre("AfficherEpoques"))
    EpoqueCouleur = Val(FicIni.Parametre("EpoqueCouleur"))
    EpoquePositionX = Val(FicIni.Parametre("EpoquePositionX"))
    EpoquePositionY = Val(FicIni.Parametre("EpoquePositionY"))
    
    'Frames par secondes.
    AfficherFPS = CBool(FicIni.Parametre("AfficherFPS"))
    FPSCouleur = Val(FicIni.Parametre("FPSCouleur"))
    'FPSMessage = FicIni.Parametre("FPSMessage")
    FPSPositionX = Val(FicIni.Parametre("FPSPositionX"))
    FPSPositionY = Val(FicIni.Parametre("FPSPositionY"))
    
    'Message des actions.
    AfficherAction = CBool(FicIni.Parametre("AfficherAction"))
    MessageActionPositionX = Val(FicIni.Parametre("MessageActionPositionX"))
    MessageActionPositionY = Val(FicIni.Parametre("MessageActionPositionY"))
    MessageActionCouleur = Val(FicIni.Parametre("MessageActionCouleur"))
    ViePositionX = Val(FicIni.Parametre("ViePositionX"))
    ViePositionY = Val(FicIni.Parametre("ViePositionY"))
    EnergiePositionX = Val(FicIni.Parametre("EnergiePositionX"))
    EnergiePositionY = Val(FicIni.Parametre("EnergiePositionY"))
    MagiePositionX = Val(FicIni.Parametre("MagiePositionX"))
    MagiePositionY = Val(FicIni.Parametre("MagiePositionY"))
    MoralPositionX = Val(FicIni.Parametre("MoralPositionX"))
    MoralPositionY = Val(FicIni.Parametre("MoralPositionY"))
    
    'Barres de statut.
    AfficherBarresStatut = CBool(FicIni.Parametre("AfficherBarresStatut"))
    AfficherBarresStatutScore = CBool(FicIni.Parametre("AfficherBarresStatutScore"))
    BarresStatutScoreCouleur = Val(FicIni.Parametre("BarresStatutScoreCouleur"))
    BarresStatutScoreDecalageX = Val(FicIni.Parametre("BarresStatutScoreDecalageX"))
    BarresStatutScoreDecalageY = Val(FicIni.Parametre("BarresStatutScoreDecalageY"))
    BarresStatutScoreLargeurTexte = Val(FicIni.Parametre("BarresStatutScoreLargeurTexte"))
    NombreBoutonsBarresStatut = Val(FicIni.Parametre("NombreBoutonsBarresStatut"))
    BarresStatutDecalageX = Val(FicIni.Parametre("BarresStatutDecalageX"))
    BarresStatutDecalageY = Val(FicIni.Parametre("BarresStatutDecalageY"))
    BarresStatutLargeur = Val(FicIni.Parametre("BarresStatutLargeur"))
    BarresStatutHauteur = Val(FicIni.Parametre("BarresStatutHauteur"))
    BarresStatutContourEpaisseur = Val(FicIni.Parametre("BarresStatutContourEpaisseur"))
    BarresStatutContourCouleur1 = Val(FicIni.Parametre("BarresStatutContourCouleur1"))
    BarresStatutContourCouleur2 = Val(FicIni.Parametre("BarresStatutContourCouleur2"))
    BarresStatutFormatCarac = FicIni.Parametre("BarresStatutFormatCarac")
    
    'Boutons construire.
    AfficherBoutonsConstruire = CBool(FicIni.Parametre("AfficherBoutonsConstruire"))
    'BoutonDemolirEtiquette = FicIni.Parametre("BoutonDemolirEtiquette")
    BoutonConstruirePositionX = Val(FicIni.Parametre("BoutonConstruirePositionX"))
    BoutonConstruirePositionY = Val(FicIni.Parametre("BoutonConstruirePositionY"))
    BoutonConstruireIntervalX = Val(FicIni.Parametre("BoutonConstruireIntervalX"))
    BoutonConstruireIntervalY = Val(FicIni.Parametre("BoutonConstruireIntervalY"))
    BoutonConstruireNombreColonnes = Val(FicIni.Parametre("BoutonConstruireNombreColonnes"))
    BoutonConstruireColonneIntervalX = Val(FicIni.Parametre("BoutonConstruireColonneIntervalX"))
    BoutonConstruireColonneIntervalY = Val(FicIni.Parametre("BoutonConstruireColonneIntervalY"))
    MasquerAutoBoutonsConstruire = CBool(FicIni.Parametre("MasquerAutoBoutonsConstruire"))
    MasquerAutoBoutonsConstruireDecalageX = Val(FicIni.Parametre("MasquerAutoBoutonsConstruireDecalageX"))
    MasquerAutoBoutonsConstruireDecalageY = Val(FicIni.Parametre("MasquerAutoBoutonsConstruireDecalageY"))
    MasquerAutoBoutonsConstruireDuree = Val(FicIni.Parametre("MasquerAutoBoutonsConstruireDuree"))
    
    'Boutons infos.
    AfficherBoutonsInfos = CBool(FicIni.Parametre("AfficherBoutonsInfos"))
    NombreBoutonsInfos = Val(FicIni.Parametre("NombreBoutonsInfos"))
    BoutonInfosTransparents = CBool(FicIni.Parametre("BoutonInfosTransparents"))
    BoutonInfosPositionX = Val(FicIni.Parametre("BoutonInfosPositionX"))
    BoutonInfosPositionY = Val(FicIni.Parametre("BoutonInfosPositionY"))
    BoutonInfosIntervalX = Val(FicIni.Parametre("BoutonInfosIntervalX"))
    BoutonInfosIntervalY = Val(FicIni.Parametre("BoutonInfosIntervalY"))
    ReDim BoutonInfosInfoBulle(NombreBoutonsInfos - 1)
    'For i = 0 To UBound(BoutonInfosInfoBulle())
    '    BoutonInfosInfoBulle(i) = FicIni.Parametre("BoutonInfosInfoBulle" & i + 1)
    'Next i
    
    'Boutons de menu.
    AfficherBoutonsMenu = CBool(FicIni.Parametre("AfficherBoutonsMenu"))
    NombreBoutonsMenu = Val(FicIni.Parametre("NombreBoutonsMenu"))
    BoutonMenuTransparents = CBool(FicIni.Parametre("BoutonMenuTransparents"))
    BoutonMenuPositionX = Val(FicIni.Parametre("BoutonMenuPositionX"))
    BoutonMenuPositionY = Val(FicIni.Parametre("BoutonMenuPositionY"))
    BoutonMenuIntervalX = Val(FicIni.Parametre("BoutonMenuIntervalX"))
    BoutonMenuIntervalY = Val(FicIni.Parametre("BoutonMenuIntervalY"))
    ReDim BoutonMenuInfoBulle(NombreBoutonsMenu - 1)
    'For i = 0 To UBound(BoutonMenuInfoBulle())
    '    BoutonMenuInfoBulle(i) = FicIni.Parametre("BoutonMenuInfoBulle" & i + 1)
    'Next i
    
    'Boutons déplacement.
    AfficherBoutonsDeplacement = CBool(FicIni.Parametre("AfficherBoutonsDeplacement"))
    NombreBoutonsDeplacement = Val(FicIni.Parametre("NombreBoutonsDeplacement"))
    BoutonDeplacementPositionX = Val(FicIni.Parametre("BoutonDeplacementPositionX"))
    BoutonDeplacementPositionY = Val(FicIni.Parametre("BoutonDeplacementPositionY"))
    BoutonDeplacementIntervalX = Val(FicIni.Parametre("BoutonDeplacementIntervalX"))
    BoutonDeplacementIntervalY = Val(FicIni.Parametre("BoutonDeplacementIntervalY"))
    BoutonDeplacementNombreColonnes = Val(FicIni.Parametre("BoutonDeplacementNombreColonnes"))
    BoutonDeplacementColonneIntervalX = Val(FicIni.Parametre("BoutonDeplacementColonneIntervalX"))
    BoutonDeplacementColonneIntervalY = Val(FicIni.Parametre("BoutonDeplacementColonneIntervalY"))
    MasquerAutoBoutonsDeplacement = CBool(FicIni.Parametre("MasquerAutoBoutonsDeplacement"))
    MasquerAutoBoutonsDeplacementDecalageX = Val(FicIni.Parametre("MasquerAutoBoutonsDeplacementDecalageX"))
    MasquerAutoBoutonsDeplacementDecalageY = Val(FicIni.Parametre("MasquerAutoBoutonsDeplacementDecalageY"))
    MasquerAutoBoutonsDeplacementDuree = Val(FicIni.Parametre("MasquerAutoBoutonsDeplacementDuree"))
    NombreBoutonsOrdre = Val(FicIni.Parametre("NombreBoutonsOrdre"))
    BoutonOrdrePositionX = Val(FicIni.Parametre("BoutonOrdrePositionX"))
    BoutonOrdrePositionY = Val(FicIni.Parametre("BoutonOrdrePositionY"))
    BoutonOrdreIntervalX = Val(FicIni.Parametre("BoutonOrdreIntervalX"))
    BoutonOrdreIntervalY = Val(FicIni.Parametre("BoutonOrdreIntervalY"))
    ReDim BoutonDeplacementInfoBulle(NombreBoutonsDeplacement - 1 + NombreBoutonsOrdre)
    'For i = 0 To UBound(BoutonDeplacementInfoBulle())
    '    BoutonDeplacementInfoBulle(i) = FicIni.Parametre("BoutonDeplacementInfoBulle" & i + 1)
    'Next i
    
    'Touches de raccourci objets.
    AfficherRaccourciObjets = CBool(FicIni.Parametre("AfficherRaccourciObjets"))
    NombreRaccourciObjets = Val(FicIni.Parametre("NombreRaccourciObjets"))
    RaccourciObjetsTransparent = CBool(FicIni.Parametre("RaccourciObjetsTransparent"))
    ReDim BoutonsRaccourciObjets(NombreRaccourciObjets - 1)
    ReDim BoutonsRaccourciObjetsMessages(NombreRaccourciObjets - 1)
    ReDim BoutonsRaccourciObjetsImages(NombreRaccourciObjets - 1)
    For i = 0 To NombreRaccourciObjets - 1
        Set BoutonsRaccourciObjetsImages(i) = New Collection
    Next i
    
    'Effets temporaires.
    AfficherEffetsTemporaires = CBool(FicIni.Parametre("AfficherEffetsTemporaires"))
    EffetTemporatireTransparent = CBool(FicIni.Parametre("EffetTemporatireTransparent"))
    EffetTemporatireDureeEffacement = Val(FicIni.Parametre("EffetTemporatireDureeEffacement"))
    EffetTemporatirePositionX = Val(FicIni.Parametre("EffetTemporatirePositionX"))
    EffetTemporatirePositionY = Val(FicIni.Parametre("EffetTemporatirePositionY"))
    EffetTemporatireIntervalX = Val(FicIni.Parametre("EffetTemporatireIntervalX"))
    EffetTemporatireIntervalY = Val(FicIni.Parametre("EffetTemporatireIntervalY"))
    
    'Message du mode auto.
    AfficherMessageModeAuto = CBool(FicIni.Parametre("AfficherMessageModeAuto"))
    MessageModeAutoCouleur = Val(FicIni.Parametre("MessageModeAutoCouleur"))
    'MessageModeAutoMessage = FicIni.Parametre("MessageModeAutoMessage")
    MessageModeAutoPositionX = Val(FicIni.Parametre("MessageModeAutoPositionX"))
    MessageModeAutoPositionY = Val(FicIni.Parametre("MessageModeAutoPositionY"))
    
    'Infos bulles terrain.
    AfficherInfoBulleTerrain = CBool(FicIni.Parametre("AfficherInfoBulleTerrain"))
    
    AfficherStastistiques = CBool(FicIni.Parametre("AfficherStastistiques"))
    CouleurStastistiques = Val(FicIni.Parametre("CouleurStastistiques"))
    StastistiquesPositionX = Val(FicIni.Parametre("StastistiquesPositionX"))
    StastistiquesPositionY = Val(FicIni.Parametre("StastistiquesPositionY"))

    'Boutons radiaux.
    ReDim BoutonsRadiauxAngle(NombrePanneauxBoutonsRadiaux - 1)
    
    Set FileFabricationObjets = New Collection
    
    'Variables temporaires pour l'affichage plein écran.
    TempAfficherCadre = AfficherCadre
    TempAfficherEpoques = AfficherEpoques
    TempAfficherFPS = AfficherFPS
    TempAfficherAction = AfficherAction
    TempAfficherBarresStatut = AfficherBarresStatut
    TempAfficherBoutonsConstruire = AfficherBoutonsConstruire
    TempAfficherBoutonsInfos = AfficherBoutonsInfos
    TempAfficherBoutonsDeplacement = AfficherBoutonsDeplacement
    TempAfficherStastistiques = AfficherStastistiques
    TempAfficherRaccourciObjets = AfficherRaccourciObjets
    TempAfficherEffetsTemp = AfficherEffetsTemporaires
    
    'Charge la version française ou anglaise.
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & NomFichierINI
    
    TitreFenetreInventaire = FicIni.Parametre("TitreFenetreInventaire")
    TitreFenetreAchat = FicIni.Parametre("TitreFenetreAchat")
    TitreFenetreVente = FicIni.Parametre("TitreFenetreVente")
    TitreFenetreRessources = FicIni.Parametre("TitreFenetreRessources")
    TitreFenetreCaracteristiques = FicIni.Parametre("TitreFenetreCaracteristiques")
    TitreFenetreCompetences = FicIni.Parametre("TitreFenetreCompetences")
    TitreFenetreBestiaire = FicIni.Parametre("TitreFenetreBestiaire")
    TitreFenetreCarte = FicIni.Parametre("TitreFenetreCarte")
    TitreFenetreObjectifs = FicIni.Parametre("TitreFenetreObjectifs")
    
    FPSMessage = FicIni.Parametre("FPSMessage")
    
    BarresStatutSeuilMessage1 = FicIni.Parametre("BarresStatutSeuilMessage1")
    BarresStatutSeuilMessage2 = FicIni.Parametre("BarresStatutSeuilMessage2")
    
    BoutonDemolirEtiquette = FicIni.Parametre("BoutonDemolirEtiquette")
    For i = 0 To UBound(BoutonInfosInfoBulle())
        BoutonInfosInfoBulle(i) = FicIni.Parametre("BoutonInfosInfoBulle" & i + 1)
    Next i
    For i = 0 To UBound(BoutonMenuInfoBulle())
        BoutonMenuInfoBulle(i) = FicIni.Parametre("BoutonMenuInfoBulle" & i + 1)
    Next i
    For i = 0 To UBound(BoutonDeplacementInfoBulle())
        BoutonDeplacementInfoBulle(i) = FicIni.Parametre("BoutonDeplacementInfoBulle" & i + 1)
    Next i
    BoutonRaccourciObjetEtiquette = FicIni.Parametre("BoutonRaccourciObjetEtiquette")
    MessageModeAutoMessage = FicIni.Parametre("MessageModeAutoMessage")
    
    NombreMessagesActions = Val(FicIni.Parametre("NombreMessagesActions"))
    ReDim MessagesActions(NombreMessagesActions - 1)
    For i = 0 To NombreMessagesActions - 1
        MessagesActions(i) = FicIni.Parametre("MessageAction" & i + 1)
    Next i
    
    MessageConfirmationQuitter = FicIni.Parametre("MessageConfirmationQuitter")
    MessageConfirmationAttaquer = FicIni.Parametre("MessageConfirmationAttaquer")
    MessageConfirmationOui = FicIni.Parametre("MessageConfirmationOui")
    MessageConfirmationNon = FicIni.Parametre("MessageConfirmationNon")
End Sub

Public Sub Init2(ByRef dd As IDirectDraw7, _
                ByRef Ecran As ClsPerEcran, _
                ByRef Parametres As ClsJeuParam)
    Dim i As Long, j As Long, k As Long, l As Long
    Dim Temp As Collection
    Set Temp = New Collection
    
    AffObjets.Charger dd, Parametres
    AffRessources.Charger dd, Parametres, True
    'Chargement des boutons.
    
    NombreBoutonsConstruire = Parametres.NombreTypeBatimentsJouables + 1
    ReDim BoutonsConstruire(NombreBoutonsConstruire)
    'Affecte un indice de batiment à chaque bouton.
    ReDim BoutonsConstruire_IndiceBatiment(NombreBoutonsConstruire)
    l = 1
    For i = 0 To Parametres.NombreTypeBatiments - 1
        If Parametres.Batiment_Jouable(Parametres.Batiment_Tri(i)) Then
            BoutonsConstruire_IndiceBatiment(l) = Parametres.Batiment_Tri(i)
            'Debug.Print l & " : " & Parametres.Batiment_NoEpoque(BoutonsConstruire_IndiceBatiment(l)) & " : " & Parametres.Batiment_Nom(BoutonsConstruire_IndiceBatiment(l))
            l = l + 1
        End If
    Next i
    ReDim Etat(NombreBoutonsConstruire)
    Etat(0) = 1
    ReDim AffBoutonsConstruire(NombreBoutonsConstruire)
    For i = 0 To NombreBoutonsConstruire
        Set AffBoutonsConstruire(i) = New ClsAffBouton
        'AffBoutonsConstruire(i).Charger dd, 0, i + 1
        AffBoutonsConstruire(i).Charger dd, 0, IIf(i = 0, AffBoutonsConstruire(i).NomBoutonDemolir, Parametres.Batiment_ImageBouton(BoutonsConstruire_IndiceBatiment(i))), Parametres.Batiment_CheminImageBouton(BoutonsConstruire_IndiceBatiment(i))
    Next i
    
    k = 0: l = 0
    For i = 0 To UBound(BoutonsConstruire())
        Set BoutonsConstruire(i) = New ClsIntBouton
        'If l = Int(UBound(BoutonsConstruire()) / BoutonConstruireNombreColonnes) + 1 Then
        '    k = k + 1
        '    l = 0
        'End If
        If i Mod BoutonConstruireNombreColonnes = 0 Then
            l = l + 1
            k = 0
        End If
        If MasquerAutoBoutonsConstruire Then
            BoutonsConstruire(i).PositionX = BoutonConstruirePositionX + BoutonConstruireIntervalX * l + BoutonConstruireColonneIntervalX * k + MasquerAutoBoutonsConstruireDecalageX
            BoutonsConstruire(i).PositionY = BoutonConstruirePositionY + BoutonConstruireIntervalY * l + BoutonConstruireColonneIntervalY * k + MasquerAutoBoutonsConstruireDecalageY
        Else
            BoutonsConstruire(i).PositionX = BoutonConstruirePositionX + BoutonConstruireIntervalX * l + BoutonConstruireColonneIntervalX * k
            BoutonsConstruire(i).PositionY = BoutonConstruirePositionY + BoutonConstruireIntervalY * l + BoutonConstruireColonneIntervalY * k
        End If
        k = k + 1
        'l = l + 1
    Next i
    BoutonsConstruire(0).Visible = True
    BoutonsConstruire(0).InfoBulle = BoutonDemolirEtiquette
    BoutonConstruirePositionX = BoutonConstruirePositionX + BoutonConstruireIntervalX
    BoutonConstruirePositionY = BoutonConstruirePositionY + BoutonConstruireIntervalY
    
    'Boutons sur les caracteristiques.
    ReDim BoutonsCaracteristiques(NombreBoutonsBarresStatut - 1)
    ReDim BoutonsCaracteristiquesSeuil(NombreBoutonsBarresStatut - 1)
    ReDim AutoRecuperationCarac(NombreBoutonsBarresStatut - 1)
    ReDim AutoRecuperationCaracPcent(NombreBoutonsBarresStatut - 1)
    For i = 0 To NombreBoutonsBarresStatut - 1
        Set BoutonsCaracteristiques(i) = New ClsIntBouton
        Set BoutonsCaracteristiquesSeuil(i) = New ClsIntBouton
        BoutonsCaracteristiques(i).Largeur = BarresStatutLargeur
        BoutonsCaracteristiques(i).Hauteur = BarresStatutHauteur
    Next i
    BoutonsCaracteristiques(0).PositionX = ViePositionX + BarresStatutDecalageX
    BoutonsCaracteristiques(0).PositionY = ViePositionY + BarresStatutDecalageY
    BoutonsCaracteristiques(1).PositionX = EnergiePositionX + BarresStatutDecalageX
    BoutonsCaracteristiques(1).PositionY = EnergiePositionY + BarresStatutDecalageY
    BoutonsCaracteristiques(2).PositionX = MagiePositionX + BarresStatutDecalageX
    BoutonsCaracteristiques(2).PositionY = MagiePositionY + BarresStatutDecalageY
    BoutonsCaracteristiques(3).PositionX = MoralPositionX + BarresStatutDecalageX
    BoutonsCaracteristiques(3).PositionY = MoralPositionY + BarresStatutDecalageY
    'Positions de la barre de seuil des boutons de caractéristiques.
    For i = 0 To NombreBoutonsBarresStatut - 1
        'BoutonsCaracteristiquesSeuil(i).PositionX=BoutonsCaracteristiques(i).PositionX
        BoutonsCaracteristiquesSeuil(i).PositionY = BoutonsCaracteristiques(i).PositionY
        BoutonsCaracteristiquesSeuil(i).Largeur = BarresStatutContourEpaisseur
        BoutonsCaracteristiquesSeuil(i).Hauteur = BarresStatutHauteur
    Next i
    'Boutons d'infos.
    AffBoutonsInfos.Charger dd, 1
    ReDim BoutonsInfos(NombreBoutonsInfos - 1)
    For i = 0 To UBound(BoutonsInfos())
        Set BoutonsInfos(i) = New ClsIntBouton
        BoutonsInfos(i).PositionX = BoutonInfosPositionX + (BoutonInfosIntervalX * i)
        BoutonsInfos(i).PositionY = BoutonInfosPositionY + (BoutonInfosIntervalY * i)
        BoutonsInfos(i).InfoBulle = BoutonInfosInfoBulle(i)
    Next i
    
    Objectifs_Initialiser
    
    'Boutons radiaux.
    ActiverBoutonsRadiaux = False
    For i = 0 To UBound(BoutonsRadiauxAngle())
        BoutonsRadiauxAngle(i) = BoutonsRadiauxAngleDefaut
    Next i
    ReDim BoutonsRadiauxInfos(NombreBoutonsRadiauxInfo - 1)
    For i = 0 To UBound(BoutonsRadiauxInfos())
        Set BoutonsRadiauxInfos(i) = New ClsIntBouton
        BoutonsRadiauxInfos(i).InfoBulle = BoutonInfosInfoBulle(i)
        BoutonsRadiauxInfos(i).Largeur = AffBoutonsInfos.Largeur
        BoutonsRadiauxInfos(i).Hauteur = AffBoutonsInfos.Hauteur
    Next i
    ReDim BoutonsRadiauxConstruire(NombreBoutonsConstruire)
    For i = 0 To UBound(BoutonsRadiauxConstruire())
        Set BoutonsRadiauxConstruire(i) = New ClsIntBouton
        BoutonsRadiauxConstruire(i).Largeur = AffBoutonsConstruire(i).Largeur
        BoutonsRadiauxConstruire(i).Hauteur = AffBoutonsConstruire(i).Hauteur
    Next i
    
    'Boutons de menu.
    AffBoutonsMenu.Charger dd, 7
    ReDim BoutonsMenu(NombreBoutonsMenu - 1)
    For i = 0 To UBound(BoutonsMenu())
        Set BoutonsMenu(i) = New ClsIntBouton
        BoutonsMenu(i).PositionX = BoutonMenuPositionX + (BoutonMenuIntervalX * i)
        BoutonsMenu(i).PositionY = BoutonMenuPositionY + (BoutonMenuIntervalY * i)
        BoutonsMenu(i).InfoBulle = BoutonMenuInfoBulle(i)
    Next i
    
    'Boutons de déplacement.
    AffBoutonsDeplacement.Charger dd, 3
    ReDim BoutonsDeplacement(NombreBoutonsDeplacement - 1 + NombreBoutonsOrdre)
    k = 0
    l = -1
    For i = 0 To UBound(BoutonsDeplacement())
        Set BoutonsDeplacement(i) = New ClsIntBouton
        If i < NombreBoutonsDeplacement Then
            If i Mod BoutonDeplacementNombreColonnes = 0 Then
                l = l + 1
                k = 0
            End If
            
            If MasquerAutoBoutonsDeplacement Then
                BoutonsDeplacement(i).PositionX = BoutonDeplacementPositionX + BoutonDeplacementIntervalX * l + BoutonDeplacementColonneIntervalX * k + MasquerAutoBoutonsDeplacementDecalageX
                BoutonsDeplacement(i).PositionY = BoutonDeplacementPositionY + BoutonDeplacementIntervalY * l + BoutonDeplacementColonneIntervalY * k + MasquerAutoBoutonsDeplacementDecalageY
            Else
                BoutonsDeplacement(i).PositionX = BoutonDeplacementPositionX + BoutonDeplacementIntervalX * l + BoutonDeplacementColonneIntervalX * k
                BoutonsDeplacement(i).PositionY = BoutonDeplacementPositionY + BoutonDeplacementIntervalY * l + BoutonDeplacementColonneIntervalY * k
            End If
            k = k + 1
        Else
            If i = NombreBoutonsDeplacement Then l = 0
            If MasquerAutoBoutonsDeplacement Then
                BoutonsDeplacement(i).PositionX = BoutonOrdrePositionX + BoutonOrdreIntervalX * l + MasquerAutoBoutonsDeplacementDecalageX
                BoutonsDeplacement(i).PositionY = BoutonOrdrePositionY + BoutonOrdreIntervalY * l + MasquerAutoBoutonsDeplacementDecalageY
            Else
                BoutonsDeplacement(i).PositionX = BoutonOrdrePositionX + BoutonOrdreIntervalX * l
                BoutonsDeplacement(i).PositionY = BoutonOrdrePositionY + BoutonOrdreIntervalY * l
            End If
            l = l + 1
        End If
        BoutonsDeplacement(i).InfoBulle = BoutonDeplacementInfoBulle(i)
    Next i
    
    'Boutons  de raccourci objets.
    For i = 0 To UBound(BoutonsRaccourciObjets())
        Set BoutonsRaccourciObjets(i) = New ClsIntBouton
        BoutonsRaccourciObjets(i).PositionX = EffetTemporatirePositionX + EffetTemporatireIntervalX * i
        BoutonsRaccourciObjets(i).PositionY = EffetTemporatirePositionY + EffetTemporatireIntervalY * i
    Next i
    AffBoutonsRaccourciObjets.Charger dd, 10

    'Chargement du cadre d'arrière plan.
    If AfficherCadre Then
        AffCadre.Definir_Nombre_Sprites = 1
        AffCadre.Definir_Image(0) = "Cadre"
        AffCadre.Charger dd
        AffCadre.Largeur = Ecran.Width
        AffCadre.Hauteur = Ecran.Height
        'AffCadre.Transparent = True
    End If
    
    'Chargement des fenêtres.
    ReDim Fenetres(NombreFenetres - 1)
    For i = 0 To NombreFenetres - 1
        AffFenetre(i).Charger dd
        Set Fenetres(i) = New ClsIntFenetre
        With Fenetres(i)
            .Init dd, Parametres
            .PositionX = FenetrePositionX(i)
            .PositionY = FenetrePositionY(i)
        End With
    Next i
    'Fenêtre confirmation.
    Set FenetreConfirmation = New ClsIntConfirmation
    FenetreConfirmation.Init dd
    
    'ReDim TouchesRaccourciObjets(Parametres.PersosNombreObjetsEquipes - 1 + Parametres.PersosNombreObjetsInventaire)
    ReDim TouchesRaccourciObjets(Parametres.NombreObjets - 1)
    
    DeplacerFenetre = -1
    
    'Remise à zéro des paramètres des dialogues.
    'AfficherDialogue = False
    'DialogueTexte1 = ""
    'DialogueTexte2 = ""
    'DialogueCurseurPosition = 0
    'DialogueCurseurVisible = False
    'Collection_Vider MessagesPrecedents
    'MessagesPrecedentsPosition = 0
    'MessagesPrecedentsAffiche = False
End Sub

Public Sub Afficher(ByVal SourisX As Integer, _
                    ByVal SourisY As Integer, _
                    ByRef Perso As ClsJeuPerso, _
                    ByRef Maison As ClsJeuBatiment, _
                    ByRef SaMaison As ClsJeuBatiment, _
                    ByRef Chateau As ClsJeuBatiment, _
                    ByRef Fief As ClsJeuFief, _
                    ByRef ArmeeSelectionnee As Boolean, _
                    ByRef Parametres As ClsJeuParam, _
                    ByRef SurfBack As IDirectDrawSurface7, _
                    ByVal X As Long, ByVal Y As Long, _
                    ByVal Largeur As Integer, ByVal Hauteur As Integer)
    Afficher_Cadre SurfBack
    
    If AfficherBarresStatut Then
        SurfBack.SetForeColor &HFFFFFF
        'SurfBack.SetFont AffDegat.FontDegats
        Afficher_Moral Perso, SurfBack
        Afficher_Magie Perso, SurfBack
        Afficher_Energie Perso, SurfBack
        Afficher_Vie Perso, SurfBack
        Afficher_Barre_Statut_Auto SurfBack
        'SurfBack.SetFont AffDegat.FontNormale
    End If
    
    If AfficherStastistiques Then Afficher_Stastistiques AffMessage, SurfBack
    
    If AfficherEpoques Then Afficher_Epoques Perso, Fief, Parametres, SurfBack
    
    If AfficherRaccourciObjets Then Afficher_Raccourcis_Objets Perso, SurfBack
    If AfficherEffetsTemporaires Then Afficher_Effets_Temporaires Perso, SurfBack
    
    If AfficherBoutonsDeplacement Then Afficher_BoutonsDeplacement Perso, Maison, ArmeeSelectionnee, SurfBack
    If AfficherBoutonsConstruire Then Afficher_BoutonsConstruire SurfBack
    If AfficherMessageModeAuto Then Afficher_ModeAuto Perso, SurfBack
    
    If AfficherAction Then Afficher_Action Perso, Maison, Parametres, SurfBack

    If AfficherBoutonsInfos Then Afficher_BoutonsInfos SurfBack
    If AfficherBoutonsMenu Then Afficher_BoutonsMenu SurfBack
    
    Afficher_Fenetres Perso, Maison, SaMaison, Chateau, Fief, Parametres, SurfBack, X, Y, Largeur, Hauteur
    
    If ActiverBoutonsRadiaux Then Afficher_BoutonsRadiaux SourisX, SourisY, SurfBack
    Afficher_Effets_Visuels SurfBack
    
    If AfficherDialogue Then Afficher_Dialogue X, Y, Largeur, Hauteur, SurfBack
    
    If FenetreConfirmation.Affiche Then
        If FrmMoteur2D.NombreJoueursMO = 1 Then
            AffTemps.TempEnCours = AffTemps.NombreTemps - 1
            AffTemps.Afficher X, Y, Largeur, Hauteur, 0, 0, SurfBack, True
            AffTemps.Actualiser_Temps
        End If
        FenetreConfirmation.Afficher SurfBack
    End If
    'SurfBack.SetForeColor 0
End Sub

Public Sub Afficher_Objet_Curseur(ByVal Perso As ClsJeuPerso, _
                                  ByVal X As Integer, _
                                  ByVal Y As Integer, _
                                  ByRef SurfBack As IDirectDrawSurface7)
    'Affiche l'objet sélectionné sous le curseur de la souris.
    'X et Y sont les coordonnées de l'emplacement de l'objet.
    Dim i As Integer
    Dim Message As String
    Dim Temp As Boolean
    Perso.Ressources_Verifications 'Vérifie qu'il n'a pas sélectionné plus qu'il ne possède.
    If Perso.ArgentSelectionne > 0 Then
        Message = "Argent : " & Int(Perso.ArgentSelectionne)
        Temp = True
    End If
    With Parametres
    For i = 0 To .NombreRessources - 1
        If Perso.Definir_RessourcesSelectionnes(i) >= 1 Then
            If Temp Then
                Message = Message & "#"
            Else
                Temp = True
            End If
            Message = Message & .Ressources_Nom(i) & " : " & Int(Perso.Definir_RessourcesSelectionnes(i))
        End If
    Next i
    End With
    If Message <> "" Then
        Afficher_Info_Bulles_Terrain X, Y, Message, SurfBack
    End If
    If Perso.ObjetSelectionne >= 0 Then
        AffObjets.Afficher X, Y, Perso.ObjetSelectionne, SurfBack
    End If
End Sub

Public Sub Position_Contour_Fenetre(ByRef X As Single, _
                                    ByRef Y As Single, _
                                    ByVal EcranLargeur As Integer, _
                                    ByVal EcranHauteur As Integer)
    'Repositionne le curseur pour éviter que celui-ci sorte de la fenêtre.
    If DeplacerFenetre >= 0 Then
        If X > EcranLargeur - AffFenetre(DeplacerFenetre).Largeur Then X = EcranLargeur - AffFenetre(DeplacerFenetre).Largeur
        If Y > EcranHauteur - AffFenetre(DeplacerFenetre).Hauteur Then Y = EcranHauteur - AffFenetre(DeplacerFenetre).Hauteur
    End If
End Sub

Public Sub Afficher_Contour_Fenetre(ByVal X As Single, _
                                    ByVal Y As Single, _
                                    ByRef SurfBack As IDirectDrawSurface7)
    'Affiche le contour de la fenêtre quand on la déplace.
    SurfBack.SetFillStyle 1
    SurfBack.SetForeColor &HFFFFFF
    SurfBack.DrawBox X, _
                     Y, _
                     X + AffFenetre(DeplacerFenetre).Largeur, _
                     Y + AffFenetre(DeplacerFenetre).Hauteur
End Sub

Private Sub Activer_EffetVisuel_Clique(ByVal X As Integer, _
                                       ByVal Y As Integer)
    Dim EffetTemp As ClsJeuEffet
    If Parametres.EffetClique > 0 Then
        Set EffetTemp = New ClsJeuEffet
        EffetTemp.Activer_Effet Parametres.EffetClique, X, Y, 0, 0, Parametres
        Effets.Add EffetTemp
    End If
End Sub

Public Sub Activer_Effet_Visuel(ByVal TypeEffet As Integer, _
                                ByVal X As Integer, _
                                ByVal Y As Integer)
    Dim EffetTemp As ClsJeuEffet
    Dim Ajout As Boolean
    Dim i As Integer
    Ajout = True
    If Effets.Count > 0 Then
        'N'active pas cet effet visuel si le même effet existe déjà.
        For i = 1 To Effets.Count
            If Effets(i).Indice = TypeEffet Then
                If CLng(Effets(i).PositionX + AffEffet.Definir_Largeur(Effets(i).Indice) / 2) = X And _
                   CLng(Effets(i).PositionY + AffEffet.Definir_Hauteur(Effets(i).Indice) / 2) = Y Then
                    Ajout = False
                    i = Effets.Count + 1
                'Else
                    'Effets.Remove i
                    'Ajout = True
                End If
                
            End If
        Next i
    End If
    If Ajout Then
        Set EffetTemp = New ClsJeuEffet
        EffetTemp.Activer_Effet TypeEffet, X, Y, 0, 0, Parametres
        Effets.Add EffetTemp
    End If
End Sub

Public Sub Actualiser_Effets_Visuels()
    Dim i As Integer
    If Effets.Count > 0 Then
        For i = 1 To Effets.Count
            Effets(i).Etape_Suivante
            If Not Effets(i).Actif Then
                Effets.Remove i
                'i = i - 1
                'If i = Effets.Count Then Exit Sub
                Exit Sub
            End If
        Next i
    End If
End Sub

Private Sub Afficher_Effets_Visuels(ByVal SurfBack As IDirectDrawSurface7)
    Dim i As Integer
    If Effets.Count > 0 Then
        For i = 1 To Effets.Count
            With Effets(i)
            If .Actif Then
                AffEffet.Afficher .Indice, _
                                  .Etape_Effet, _
                                  .PositionX, _
                                  .PositionY, _
                                  SurfBack
                If Son.Definir_Jouer Then
                    If .Etape_Effet = 1 Then
                        If Not Jeu.EnPause Then
                            Son.Lire .Indice
                        End If
                    ElseIf .NombreEtapes = 0 Then
                        Son.Lire .Indice
                        Effets.Remove i
                        Exit For
                    End If
                End If
            End If
            End With
        Next i
    End If
End Sub

Public Sub Positionner_Fenetre(ByVal X As Single, _
                               ByVal Y As Single)
    Fenetres(DeplacerFenetre).PositionX = X
    Fenetres(DeplacerFenetre).PositionY = Y
    DeplacerFenetre = -1
End Sub

Public Sub RePositionner_Fenetres(ByVal EcranPositionX As Integer, _
                                  ByVal EcranPositionY As Integer, _
                                  ByVal EcranLargeur As Integer, _
                                  ByVal EcranHauteur As Integer)
    Dim i  As Integer
    EcranLargeur = EcranLargeur + EcranPositionX
    EcranHauteur = EcranHauteur + EcranPositionY
    For i = 0 To NombreFenetres - 1
        With Fenetres(i)
        If .PositionX < EcranPositionX Then _
            .PositionX = EcranPositionX
        If .PositionX > EcranLargeur - AffFenetre(i).Largeur Then _
            .PositionX = EcranLargeur - AffFenetre(i).Largeur
        If .PositionY < EcranPositionY Then _
            .PositionY = EcranPositionY
        If .PositionY > EcranHauteur - AffFenetre(i).Hauteur Then _
            .PositionY = EcranHauteur - AffFenetre(i).Hauteur
        End With
    Next i
End Sub

Private Sub Afficher_Cadre(ByRef SurfBack As IDirectDrawSurface7)
    If AfficherCadre Then
        AffCadre.Afficher 0, 0, SurfBack
    End If
End Sub

Private Sub Afficher_Epoques(ByVal Perso As ClsJeuPerso, _
                             ByVal Fief As ClsJeuFief, _
                             ByVal Parametres As ClsJeuParam, _
                             ByRef SurfBack As IDirectDrawSurface7)
    Dim EpoqueHauteur As Integer
    Dim EpoqueLargeur As Integer
    Dim EpoqueDecalageX As Integer
    Dim EpoqueDecalageY As Integer
    Dim PositionX As Integer
    Dim PositionY As Integer
    Dim Message As String
    EpoqueHauteur = 22
    EpoqueLargeur = 180
    EpoqueDecalageX = -90
    EpoqueDecalageY = -4
    PositionX = EpoquePositionX
    PositionY = EpoquePositionY
    SurfBack.SetForeColor EpoqueCouleur
    Message = Parametres.Epoque_Nom(Fief.Epoque)
    If Fief.ChangerEpoque Then
        'SurfBack.SetFillColor 0
        Afficher_Barre_Statut Parametres.Epoque_Duree(Fief.Epoque + 1) * Parametres.VitesseEpoque_Coef - Fief.TempsChangementEpoque, 0, Parametres.Epoque_Duree(Fief.Epoque + 1) * Parametres.VitesseEpoque_Coef, _
                              PositionX + EpoqueDecalageX, _
                              PositionY + EpoqueDecalageY, _
                              SurfBack, _
                              EpoqueLargeur, _
                              EpoqueHauteur, &HFFFFFF
        'SurfBack.SetForeColor EpoqueCouleur
        AffMessage.Afficher PositionX - (Len(Message) * (AffMessage.Largeur_Caractere - 0.5) / 2), PositionY, Message, SurfBack
    Else
        'Cadre noir.
        SurfBack.SetFillColor 0
        SurfBack.SetFillStyle 0
        '.DrawBox PositionX - 1, PositionY - 1, PositionX + Largeur + 1, PositionY + Hauteur + 1
        SurfBack.DrawBox PositionX + EpoqueDecalageX, PositionY, PositionX + EpoqueDecalageX + EpoqueLargeur, PositionY + EpoqueDecalageY + EpoqueHauteur
        SurfBack.SetFillStyle 1
        AffMessage.Afficher PositionX - (Len(Message) * (AffMessage.Largeur_Caractere - 0.5) / 2), PositionY, Message, SurfBack
    End If
    'AffMessage.Afficher EpoquePositionX, EpoquePositionY, Parametres.Epoque_Nom(Fief.Epoque) & IIf(Fief.ChangerEpoque, " (suivant : " & Int((1 - (Chateaux(Perso.NumeroFief).Definir_Vie - Chateaux(Perso.NumeroFief).Vie) / Parametres.Epoque_BonusVieChateau(Fief.Epoque + 1)) * 100) & "%)", ""), SurfBack
    
    'Affiche le chef
    If Fief.Empereur >= 0 Then
        Message = Parametres.Epoque_EmpereurNom(Fief.Epoque, Persos(Fief.Empereur).Feminin) & " " & Persos(Fief.Empereur).Nom
        'Cadre noir.
        'SurfBack.SetFillColor 0
        'SurfBack.SetFillStyle 0
        '.DrawBox PositionX - 1, PositionY - 1, PositionX + Largeur + 1, PositionY + Hauteur + 1
        'SurfBack.DrawBox PositionX + EpoqueDecalageX, PositionY + EpoqueHauteur + EpoqueDecalageY - 1, PositionX + EpoqueDecalageX + EpoqueLargeur, PositionY + (EpoqueDecalageY + EpoqueHauteur) * 2 - 1
        Afficher_Barre_Statut IIf(Persos(Fief.Empereur).Vie > Persos(Fief.Empereur).Definir_Vie, Persos(Fief.Empereur).Definir_Vie, IIf(Persos(Fief.Empereur).Vie > 0, Persos(Fief.Empereur).Vie, 0)), Persos(Fief.Empereur).Definir_Vie, Persos(Fief.Empereur).Definir_MaxVie, PositionX + EpoqueDecalageX, PositionY + EpoqueHauteur + EpoqueDecalageY - 1, SurfBack, EpoqueLargeur, EpoqueHauteur - 4
        'SurfBack.SetFillStyle 1
        'SurfBack.SetForeColor &HFFFFFF
        AffMessage.Afficher PositionX - (Len(Message) * (AffMessage.Largeur_Caractere - 0.5) / 2), PositionY + EpoqueHauteur + EpoqueDecalageY - 1, Message, SurfBack
    End If
End Sub

Public Sub Afficher_FPS(ByVal NombreFrames As String, _
                        ByRef SurfBack As IDirectDrawSurface7)
    If AfficherFPS Then
        SurfBack.SetForeColor FPSCouleur
        AffMessage.Afficher FPSPositionX, FPSPositionY, FPSMessage & NombreFrames, SurfBack
    End If
End Sub

Private Sub Afficher_Fenetres(ByVal Perso As ClsJeuPerso, _
                              ByVal Maison As ClsJeuBatiment, _
                              ByVal SaMaison As ClsJeuBatiment, _
                              ByVal Chateau As ClsJeuBatiment, _
                              ByVal Fief As ClsJeuFief, _
                              ByVal Parametres As ClsJeuParam, _
                              ByRef SurfBack As IDirectDrawSurface7, _
                              ByVal X As Long, ByVal Y As Long, _
                              ByVal Largeur As Integer, ByVal Hauteur As Integer)
    'Fenetres(0).Visible = (Perso.Vivant And _
                          (Perso.DansUneMaison Or _
                           Perso.DansSaMaison Or _
                           Perso.DansUnChateau Or _
                           Perso.DansSonChateau) Or _
                          (Perso.DansSaMaison Or Perso.DansUneMaison) And Maison.Service = 1) And _
                           BoutonsInfos(0).Enabled
    Fenetres(0).Visible = (Perso.Vivant And _
                          (Perso.DansUneMaison Or _
                           Perso.DansSaMaison) Or _
                           Perso.DansUnChateau Or _
                           Perso.DansSonChateau Or _
                          (Perso.DansSaMaison Or Perso.DansUneMaison) And Maison.Service = 1) And _
                           BoutonsInfos(0).Enabled
    'Fenetres(1).Visible = (Perso.DansUnChateau Or Perso.DansSonChateau) And _
                          Perso.Vivant
    'Fenetre d'inventaire
    If Fenetres(4).Visible Then
        'If Fenetres(4).Bascule = 0 Then
            AffFenetre(4).Afficher Fenetres(4).PositionX, Fenetres(4).PositionY, SurfBack, IIf(Fenetres(4).Bascule And Perso.DansSaMaison And Perso.Vivant, IIf(Perso.Feminin, 3, 2), IIf(Perso.Feminin, 1, 0))
        'Else
        '    AffFenetre.Afficher Fenetres(4).PositionX, Fenetres(4).PositionY, SurfBack, 3
        'End If
        'Fenetres(4).Afficher_InfosInventaire "Inventaire", Perso, Parametres, AffFenetre, AffMessage, SurfBack
        Fenetres(4).Afficher_InfosInventaire TitreFenetreInventaire, Perso, Maison, Parametres, AffObjets, AffFenetre(4), AffMessage, SurfBack
        Afficher_TouchesObjets Perso, SurfBack
    End If
    'Fenêtre bâtiment.
    If Fenetres(0).Visible Then
        AffFenetre(0).Afficher Fenetres(0).PositionX, Fenetres(0).PositionY, SurfBack
        If (Perso.DansUneMaison Or Perso.DansSaMaison) And Not Parametres.Batiment_Marche(Maison.TypeBatiment) Then  'Or Maison.EnConstruction Then
            'Fenêtre Maison.
            Fenetres(0).Afficher_InfosMaison Parametres.Batiment_Nom(Maison.TypeBatiment), FileFabricationObjets, Me, Perso, Maison, Parametres, AffFenetre(0), AffObjets, AffMessage, SurfBack
        Else
            'Fenêtre Chateau.
            Fenetres(0).Afficher_InfosChateau TitreFenetreAchat, Perso, Maison, Chateau, Fief, Parametres, AffRessources, AffFenetre(0), AffMessage, SurfBack
        End If
    End If
    'Fenêtre ressources.
    If Fenetres(1).Visible Then
        AffFenetre(1).Afficher Fenetres(1).PositionX, Fenetres(1).PositionY, SurfBack
        If Perso.DansUnChateau Or Perso.DansSonChateau Then
            Fenetres(1).Afficher_InfosRessourcesPerso TitreFenetreVente, Perso, Maison, Parametres, AffRessources, AffFenetre(1), AffMessage, SurfBack
        Else 'If BoutonsInfos(1).Enabled Then
            Fenetres(1).Afficher_InfosRessourcesPerso TitreFenetreRessources, Perso, Maison, Parametres, AffRessources, AffFenetre(1), AffMessage, SurfBack
        End If
    End If
    'Fenêtre des caractéristiques.
    If Fenetres(2).Visible Then
        Fenetres(2).Afficher_InfosCaracPerso1 TitreFenetreCaracteristiques, Perso, SaMaison, Parametres, AffFenetre(2), AffMessage, SurfBack
        AffFenetre(2).Afficher Fenetres(2).PositionX, Fenetres(2).PositionY, SurfBack
        Fenetres(2).Afficher_InfosCaracPerso2 TitreFenetreCaracteristiques, Perso, SaMaison, Parametres, AffFenetre(2), AffMessage, SurfBack
    End If
    'Fenetre de compétence.
    If Fenetres(3).Visible Then
        If TabComp.Count > 0 Then Fenetres(3).Afficher_InfosCompetences1 TabComp, TitreFenetreCompetences, Perso, Maison, Parametres, AffFenetre(3), AffMessage, SurfBack
        AffFenetre(3).Afficher Fenetres(3).PositionX, Fenetres(3).PositionY, SurfBack
        Fenetres(3).Afficher_InfosCompetences2 TabComp, TitreFenetreCompetences, CompetencesTri, CompetencesAfficherToutes, AffFenetre(3), AffMessage, SurfBack
    End If
    
    'Fenetre de bestiaire.
    If Fenetres(5).Visible Then
        AffFenetre(5).Afficher Fenetres(5).PositionX, Fenetres(5).PositionY, SurfBack
        Fenetres(5).Afficher_InfosBestiaire TitreFenetreBestiaire, Perso, Parametres, AffMessage, SurfBack
    End If
    
    'Carte.
    If Fenetres(6).Visible Then
        AffFenetre(6).Afficher Fenetres(6).PositionX, Fenetres(6).PositionY, SurfBack
        Fenetres(6).Afficher_Carte TitreFenetreCarte, Perso, AffMessage, SurfBack, X, Y, Largeur, Hauteur
    End If
    
    'Fenêtre d'objectifs.
    If Fenetres(7).Visible Then
        AffFenetre(7).Afficher Fenetres(7).PositionX, Fenetres(7).PositionY, SurfBack
        Fenetres(7).Afficher_InfosObjectifs TitreFenetreObjectifs, Objectifs, AffFenetre(7), AffMessage, SurfBack
    End If
End Sub

Private Sub Afficher_TouchesObjets(ByVal Perso As ClsJeuPerso, _
                                   ByRef SurfBack As IDirectDrawSurface7)
    Dim i As Integer
    Dim Temp As Integer
    SurfBack.SetFont AffDegat.FontDegats
    For i = 0 To Perso.Nombre_Objets_Equipes - 1
        If Perso.Objet_Equipes_Type(i) > -1 Then
            Temp = TouchesRaccourciObjets(Perso.Objet_Equipes_Type(i)) - BoutonRaccourciObjetDecalageY * IntTouches.NombreTouchesObjets
            'If TouchesRaccourciObjets(Perso.Objet_Equipes_Type(i)) > 0 And _
               TouchesRaccourciObjets(Perso.Objet_Equipes_Type(i)) < 10 Then
            If Temp > 0 And _
               Temp <= IntTouches.NombreTouchesObjets Then
                    AffMessage.Afficher Fenetres(4).BoutonsObjetsPositionX(i + Perso.Nombre_Objets_Inventaire), _
                                        Fenetres(4).BoutonsObjetsPositionY(i + Perso.Nombre_Objets_Inventaire), _
                                        "(" & Temp & ")", _
                                        SurfBack
            ElseIf TouchesRaccourciObjets(Perso.Objet_Equipes_Type(i)) > 0 Then
                AffMessage.Afficher Fenetres(4).BoutonsObjetsPositionX(i + Perso.Nombre_Objets_Inventaire), _
                                        Fenetres(4).BoutonsObjetsPositionY(i + Perso.Nombre_Objets_Inventaire), _
                                        "(+)", _
                                        SurfBack
            End If
        End If
    Next i
    If Not (Fenetres(4).Bascule = 1 And Perso.DansSaMaison And Perso.Vivant) Then
        For i = 0 To Perso.Nombre_Objets_Inventaire - 1
            If Perso.Objet_Inventaire_Type(i) > -1 Then
                Temp = TouchesRaccourciObjets(Perso.Objet_Inventaire_Type(i)) - BoutonRaccourciObjetDecalageY * IntTouches.NombreTouchesObjets
                'If TouchesRaccourciObjets(Perso.Objet_Inventaire_Type(i)) > 0 And _
                   TouchesRaccourciObjets(Perso.Objet_Inventaire_Type(i)) < 10 Then
                If Temp > 0 And _
                   Temp <= IntTouches.NombreTouchesObjets Then
                        AffMessage.Afficher Fenetres(4).BoutonsObjetsPositionX(i), _
                                            Fenetres(4).BoutonsObjetsPositionY(i), _
                                            "(" & Temp & ")", _
                                            SurfBack
                ElseIf TouchesRaccourciObjets(Perso.Objet_Inventaire_Type(i)) > 0 Then
                        AffMessage.Afficher Fenetres(4).BoutonsObjetsPositionX(i), _
                                                Fenetres(4).BoutonsObjetsPositionY(i), _
                                                "(+)", _
                                                SurfBack
                End If
                If Perso.Objet_Inventaire_Nombre(i) > 1 Then
                    AffMessage.Afficher Fenetres(4).BoutonsObjetsPositionX(i) + 4, _
                                        Fenetres(4).BoutonsObjetsPositionY(i) + 19, _
                                        "x" & Perso.Objet_Inventaire_Nombre(i), _
                                        SurfBack
                End If
            End If
        Next i
    End If
    SurfBack.SetFont AffDegat.FontNormale
End Sub

Public Sub Actualiser_BoutonsConstruire(ByRef Perso As ClsJeuPerso, _
                                        ByRef Maison As ClsJeuBatiment, _
                                        ByRef Fief As ClsJeuFief, _
                                        ByRef Parametres As ClsJeuParam)
    Dim i As Long
    Dim j As Long
    'Bouton démolir.
    Etat(0) = -CDbl(Maison.Visible Or Persos(Perso.IndiceMaison).NumeroChef = Perso.Numero And Perso.DansUneMaison)
    For i = 1 To NombreBoutonsConstruire - 1
        'If i = 0 Then
'            'Bouton démolir.
'            If Maison.Visible And Perso.Vivant Then
'                Etat(i) = 1
'            Else
'                Etat(i) = 0
'            End If
        'Else
        'BoutonsConstruire(i).Visible = Fiefs(Perso.NumeroFief).Epoque >= Parametres.Batiment_NoEpoque(i - 1) - 1
            Etat(i) = _
                      IIf((Fiefs(Perso.NumeroFief).Epoque >= Parametres.Batiment_NoEpoque(BoutonsConstruire_IndiceBatiment(i)) - 1), _
                      Abs(Perso.Vivant And Perso.PeutConstruire_Maison(BoutonsConstruire_IndiceBatiment(i), Parametres)) And _
                      (Fiefs(Perso.NumeroFief).Epoque >= Parametres.Batiment_NoEpoque(BoutonsConstruire_IndiceBatiment(i)) - 1), _
                      2)
        'End If
    Next i
End Sub

Private Sub Afficher_BoutonsConstruire(ByRef SurfBack As IDirectDrawSurface7)
    Dim i As Long
    For i = 0 To NombreBoutonsConstruire - 1
        AffBoutonsConstruire(i).Afficher BoutonsConstruire(i).PositionX, _
                                         BoutonsConstruire(i).PositionY, _
                                         0, _
                                         Etat(i), _
                                         SurfBack, True
    Next i
End Sub

Private Sub Afficher_BoutonsDeplacement(ByVal Perso As ClsJeuPerso, _
                                        ByVal Maison As ClsJeuBatiment, _
                                        ByVal ArmeeSelectionnee As Boolean, _
                                        ByRef SurfBack As IDirectDrawSurface7)
    Dim i As Long
    Dim Etat As Integer
    For i = 0 To UBound(BoutonsDeplacement())
        With BoutonsDeplacement(i)
        If ((Not i = 0) Or Chateaux(Persos(NoPerso).NumeroFief).Visible) And _
           ((Not i = 1) Or Maisons(NoPerso).Visible) And _
           ((Not i > 16) Or Perso.NombreSoldats > 0) Then 'S'il na pas de maison, on affiche pas le bouton retour Maison.
           'Affiche ou non les boutons d'ordes pour les armées.
            If i < NombreBoutonsDeplacement Or ArmeeSelectionnee Then
                Select Case i
               'Case 0: Etat = -CDbl(Not Perso.Vivant)
                'Case 0 To 13: Etat = -CDbl(BoutonsDeplacement(i).Enabled)
                Case 14: Etat = -CDbl(Perso.IA)
                'Case 15 To 16: Etat = -CDbl(BoutonsDeplacement(i).Enabled)
                'Case 18: Etat = Perso.Attitude
                Case 17: Etat = Perso.Formation
                Case Else: Etat = -CDbl(BoutonsDeplacement(i).Enabled)
                End Select
                AffBoutonsDeplacement.Afficher .PositionX, _
                                               .PositionY, _
                                               i, _
                                               Etat, _
                                               SurfBack
                .Visible = True
            Else
                .Visible = False
            End If
        Else
            .Visible = False
        End If
        End With
    Next i
    If Perso.NombreSoldats > 0 Then
        SurfBack.SetForeColor MessageModeAutoCouleur
        AffMessage.Afficher BoutonsDeplacement(NombreBoutonsDeplacement).PositionX, BoutonsDeplacement(NombreBoutonsDeplacement).PositionY - BoutonDeplacementIntervalY, _
                            "Nb:#" & Perso.NombreSoldats, _
                            SurfBack
    End If
End Sub

Private Sub Actualiser_Competences(ByVal Perso As ClsJeuPerso, ByVal Maison As ClsJeuBatiment, ByVal Parametres As ClsJeuParam)
    'liste des compétences.
    Dim i As Integer, j As Integer, k As Integer
    Dim Temp As Boolean
    Set Competence = Nothing
    Set Competence = New ClsJeuCompetences
    Collection_Vider TabComp
    For i = 0 To Perso.Nombre_CompetencesSpeciales - 1
        If CompetencesAfficherToutes Or Perso.Niveau_CompetenceSpeciales(i) > 0 Or Perso.Progression_CompetenceSpeciale(i) > 0 Then
            Competence.Nom = Parametres.Speciales_NomCompetences(i) & " :"
            Competence.Valeur = Perso.Definir_Niveau_CompetenceSpeciales(i)
            'Colorie la compétence active.
            If Perso.Action = 6 And i = 0 Or _
               (Perso.Action = 301 Or Perso.Action = 302) And i = 1 Then
                Competence.Active = True
            Else
                Competence.Active = False
            End If
            'Bonus à la compétence.
            Competence.Bonus = Perso.Bonus_CompetenceSpeciales(i)
            Competence.NivPlus = Perso.Progression_CompetenceSpeciale(i)
            TabComp.Add Competence
            Set Competence = Nothing
            Set Competence = New ClsJeuCompetences
        End If
    Next i
    For i = 0 To Perso.Nombre_CompetencesRessources - 1
        If CompetencesAfficherToutes Or Perso.Niveau_CompetenceRessources(i) > 0 Or Perso.Progression_CompetenceRessource(i) > 0 Then
            Competence.Nom = Parametres.Ressources_NomCompetence(i) & " :"
            Competence.Valeur = Perso.Definir_Niveau_CompetenceRessources(i) + Parametres.PersosMaxRessources 'perso.Niveau_CompetenceRessources(i) + Parametres.PersosMaxRessources
            'Colorie la compétence active.
            Competence.Active = i = Perso.RessourceCourante
            'Bonus à la compétence.
            Competence.Bonus = Perso.Bonus_CompetenceRessources(i)
            Competence.NivPlus = Perso.Progression_CompetenceRessource(i)
            TabComp.Add Competence
            Set Competence = Nothing
            Set Competence = New ClsJeuCompetences
        End If
    Next i
    For i = 0 To Perso.Nombre_CompetencesServices - 1
        If CompetencesAfficherToutes Or Perso.Niveau_CompetenceServices(i) > 0 Or Perso.Progression_CompetenceService(i) > 0 Then
            Competence.Nom = Parametres.Service_NomCompetence(i + 1) & " :"
            Competence.Valeur = Perso.Definir_Niveau_CompetenceServices(i) 'perso.Definir_Carac_CompetenceServices(i)
            'Colorie la compétence active.
            If Perso.Action = 600 And Perso.ServiceCourant = i + 1 Then
                Competence.Active = True
            Else
                Competence.Active = False
            End If
            'Bonus à la compétence.
            Competence.Bonus = Perso.Bonus_CompetenceServices(i)
            Competence.NivPlus = Perso.Progression_CompetenceService(i)
            TabComp.Add Competence
            Set Competence = Nothing
            Set Competence = New ClsJeuCompetences
        End If
    Next i
    For i = 1 To Perso.Nombre_CompetencesObjets - 1
        j = Parametres.CompetenceObjet_NoListe(i)
        If CompetencesAfficherToutes Or Perso.Niveau_CompetenceObjets(j) > 0 Or Perso.Progression_CompetenceObjet(j) > 0 Then
            Competence.Nom = Parametres.CompetenceObjet_Nom(j) & " :"
            Competence.Valeur = Perso.Definir_Niveau_CompetenceObjets(j) 'perso.Definir_Carac_CompetenceObjets(i)
            If Perso.Action = 500 Then
                If Parametres.Objet_Competence(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, Perso.IndiceObjet)) = j Then
                    Competence.Active = True
                Else
                    Competence.Active = False
                End If
            'Else
            '    Competence.Active = False
            ElseIf Perso.RessourceCourante >= 0 Then
                For k = 0 To Perso.Nombre_Objets_Equipes - 1
                    If Perso.Objet_Equipes_Actif(k) Then
                        If Parametres.Objet_BonusCompetencesRessources(Perso.Objet_Equipes_Type(k), Perso.RessourceCourante) > 0 And _
                           Perso.Objet_Equipes_Competence(k) = j Then
                            Competence.Active = True
                            Exit For
                        End If
                    End If
                Next k
            End If
            If Not Perso.Inactif And Perso.Vivant Then
                'La compétence est active si le perso possède au moins un objet de cette compétence offrant un bonus d'énergie ou de moral.
                For k = 0 To Perso.Nombre_Objets_Equipes - 1
                    If Perso.Objet_Equipes_Actif(k) Then
                        If (Parametres.Objet_BonusEnergie(Perso.Objet_Equipes_Type(k)) > 0 Or Parametres.Objet_BonusMoral(Perso.Objet_Equipes_Type(k))) And Perso.Objet_Equipes_Competence(k) = j Then
                            Competence.Active = True
                            Exit For
                        End If
                    End If
                Next k
            End If
            Competence.Bonus = Perso.Bonus_CompetenceObjets(j)
            Competence.NivPlus = Perso.Progression_CompetenceObjet(j)
            TabComp.Add Competence
            Set Competence = Nothing
            Set Competence = New ClsJeuCompetences
        End If
    Next i
    Select Case CompetencesTri
    Case 1:
        'On trie les compétences selon leur nom.
        Do
            Temp = True 'Si temp passe à faux, on fait un tri supplémentaire.
            For i = 2 To TabComp.Count
                If TabComp(i - 1).Nom > TabComp(i).Nom Then
                    Competence.Nom = TabComp(i).Nom
                    Competence.Valeur = TabComp(i).Valeur
                    Competence.NivPlus = TabComp(i).NivPlus
                    Competence.Active = TabComp(i).Active
                    Competence.Bonus = TabComp(i).Bonus
                    TabComp(i).Nom = TabComp(i - 1).Nom
                    TabComp(i).Valeur = TabComp(i - 1).Valeur
                    TabComp(i).NivPlus = TabComp(i - 1).NivPlus
                    TabComp(i).Active = TabComp(i - 1).Active
                    TabComp(i).Bonus = TabComp(i - 1).Bonus
                    TabComp(i - 1).Nom = Competence.Nom
                    TabComp(i - 1).Valeur = Competence.Valeur
                    TabComp(i - 1).NivPlus = Competence.NivPlus
                    TabComp(i - 1).Active = Competence.Active
                    TabComp(i - 1).Bonus = Competence.Bonus
                    Temp = False
                End If
            Next i
        Loop Until Temp
    Case 2:
        'On trie les compétences selon leur score.
        Do
            Temp = True 'Si temp passe à faux, on fait un tri supplémentaire.
            For i = 2 To TabComp.Count
                If TabComp(i - 1).Valeur + TabComp(i - 1).NivPlus < TabComp(i).Valeur + TabComp(i).NivPlus Then
                    Competence.Nom = TabComp(i).Nom
                    Competence.Valeur = TabComp(i).Valeur
                    Competence.NivPlus = TabComp(i).NivPlus
                    Competence.Active = TabComp(i).Active
                    Competence.Bonus = TabComp(i).Bonus
                    TabComp(i).Nom = TabComp(i - 1).Nom
                    TabComp(i).Valeur = TabComp(i - 1).Valeur
                    TabComp(i).NivPlus = TabComp(i - 1).NivPlus
                    TabComp(i).Active = TabComp(i - 1).Active
                    TabComp(i).Bonus = TabComp(i - 1).Bonus
                    TabComp(i - 1).Nom = Competence.Nom
                    TabComp(i - 1).Valeur = Competence.Valeur
                    TabComp(i - 1).NivPlus = Competence.NivPlus
                    TabComp(i - 1).Active = Competence.Active
                    TabComp(i - 1).Bonus = Competence.Bonus
                    Temp = False
                End If
            Next i
        Loop Until Temp
    End Select
End Sub

Private Sub Actualiser_BarreStatut(ByVal Indice As Integer)
    'Reposition les informations de la jauge de la barre de statut.
    BoutonsCaracteristiquesSeuil(Indice).PositionX = BoutonsCaracteristiques(Indice).PositionX + BoutonsCaracteristiques(Indice).Largeur * AutoRecuperation_CaracPcent(Indice)
End Sub

Private Sub Actualiser_Raccourcis_Objets(ByVal Perso As ClsJeuPerso)
    Dim i, j As Integer
    Dim Temp As Integer
    For i = 0 To NombreRaccourciObjets - 1
        BoutonsRaccourciObjets(i).Enabled = False
        BoutonsRaccourciObjetsMessages(i) = ""
        Temp = i + 1 + BoutonRaccourciObjetDecalageY * IntTouches.NombreTouchesObjets
        For j = 0 To UBound(TouchesRaccourciObjets())
            If TouchesRaccourciObjets(j) = Temp And Perso.Objet_EquipeNombreType(j) > 0 Then
                BoutonsRaccourciObjets(i).Enabled = True
            End If
        Next j
        While BoutonsRaccourciObjetsImages(i).Count > 0
            BoutonsRaccourciObjetsImages(i).Remove 1
        Wend
        For j = 0 To UBound(TouchesRaccourciObjets())
            If TouchesRaccourciObjets(j) = Temp Then
                If Perso.Objet_NombreType(j) > 0 Then
                    BoutonsRaccourciObjetsImages(i).Add j
                    
                    If Not BoutonsRaccourciObjets(i).Enabled Then
                        If Perso.Objet_NombreType(j) > 1 And _
                           Parametres.Objet_Utilisable(j) And Parametres.Objet_UtilisationUnique(j) Then
                            BoutonsRaccourciObjetsMessages(i) = "x" & Perso.Objet_NombreType(j)
                        ElseIf Parametres.Objet_PerteRessourceQuantite(j) > 0 Then
                            'BoutonsRaccourciObjetsMessages(i) = IIf(Int(Perso.Definir_ressources(Parametres.Objet_PerteRessourceIndice(j)) / Parametres.Objet_PerteRessourceQuantite(j) / (Perso.NombreProjectiles + Int(Perso.Objet_Equipe_GainSup(Perso.Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(j)), Parametres.Objet_ProjectileSupplementaire(j)))) = 0), "&", "") & "(" & Format(Int(Perso.Definir_ressources(Parametres.Objet_PerteRessourceIndice(j)) / Parametres.Objet_PerteRessourceQuantite(j) / (Perso.NombreProjectiles + Int(Perso.Objet_Equipe_GainSup(Perso.Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(j)), Parametres.Objet_ProjectileSupplementaire(j))))), "00") & ")"
                            BoutonsRaccourciObjetsMessages(i) = IIf(Int(Perso.Definir_ressources(Parametres.Objet_PerteRessourceIndice(j)) / Parametres.Objet_PerteRessourceQuantite(j) = 0), "&", "") & "(" & Format(Int(Perso.Definir_ressources(Parametres.Objet_PerteRessourceIndice(j)) / Parametres.Objet_PerteRessourceQuantite(j)), "00") & ")"
                        End If
                    End If
                End If
            End If
        Next j
        For j = 0 To UBound(TouchesRaccourciObjets())
            If TouchesRaccourciObjets(j) = Temp And Perso.Objet_EquipeNombreType(j) > 0 Then
                'Affiche l'arme sélectionnée en avant plan.
                BoutonsRaccourciObjetsImages(i).Add j
                
                If Parametres.Objet_PerteRessourceQuantite(j) > 0 Then
                    'BoutonsRaccourciObjetsMessages(i) = IIf(Int(Perso.Definir_ressources(Parametres.Objet_PerteRessourceIndice(j)) / Parametres.Objet_PerteRessourceQuantite(j) / Perso.NombreProjectiles) = 0, "&", "") & "(" & Format(Int(Perso.Definir_ressources(Parametres.Objet_PerteRessourceIndice(j)) / Parametres.Objet_PerteRessourceQuantite(j) / Perso.NombreProjectiles), "00") & ")"
                    BoutonsRaccourciObjetsMessages(i) = IIf(Int(Perso.Definir_ressources(Parametres.Objet_PerteRessourceIndice(j)) / Parametres.Objet_PerteRessourceQuantite(j) = 0), "&", "") & "(" & Format(Int(Perso.Definir_ressources(Parametres.Objet_PerteRessourceIndice(j)) / Parametres.Objet_PerteRessourceQuantite(j)), "00") & ")"
                End If
                j = UBound(TouchesRaccourciObjets())
            End If
        Next j
    Next i
    ActualiserRaccourciObjets = False
End Sub

Private Sub Afficher_Raccourcis_Objets(ByVal Perso As ClsJeuPerso, _
                                       ByRef SurfBack As IDirectDrawSurface7)
    Dim i As Integer, j As Integer, k As Integer
    For i = 0 To NombreRaccourciObjets - 1
        AffBoutonsRaccourciObjets.Afficher BoutonsRaccourciObjets(i).PositionX, _
                                           BoutonsRaccourciObjets(i).PositionY, _
                                           i, -Int(BoutonsRaccourciObjets(i).Enabled) + BoutonRaccourciObjetDecalageY * 2, SurfBack, RaccourciObjetsTransparent
        For j = 1 To BoutonsRaccourciObjetsImages(i).Count
            k = BoutonsRaccourciObjetsImages(i).Item(j)
            AffObjets.Afficher BoutonsRaccourciObjets(i).PositionX, _
                               BoutonsRaccourciObjets(i).PositionY, _
                               k, _
                               SurfBack
            If BoutonsRaccourciObjets(i).Enabled Then
                If Parametres.Objet_Categorie(k) = 0 Or Parametres.Objet_Categorie(k) = 1 Then
                    With Perso
                    If .Attaque < .Definir_MaxAttaque Or _
                       .Defense < .Definir_MaxDefense Or .DureeRechargeAttaque > 0 Then
                       SurfBack.SetForeColor 0
                        'Affiche la barre d'attaque.
                        If .DureeRechargeAttaque = 0 Then
                        Afficher_Barre_Statut .Attaque, .Attaque, .Definir_MaxAttaque, _
                                              BoutonsRaccourciObjets(i).PositionX, _
                                              BoutonsRaccourciObjets(i).PositionY + AffBoutonsRaccourciObjets.Hauteur - 10, _
                                              SurfBack, _
                                              AffBoutonsRaccourciObjets.Largeur, _
                                              5, _
                                              &HFF&          'Rouge.
                        Else
                        Afficher_Barre_Statut 0, .DureeRechargeAttaqueMax - .DureeRechargeAttaque, .DureeRechargeAttaqueMax, _
                                              BoutonsRaccourciObjets(i).PositionX, _
                                              BoutonsRaccourciObjets(i).PositionY + AffBoutonsRaccourciObjets.Hauteur - 10, _
                                              SurfBack, _
                                              AffBoutonsRaccourciObjets.Largeur, _
                                              5
                        End If
                       'Affiche la barre de défense.
                        Afficher_Barre_Statut .Defense, .Defense, .Definir_MaxDefense, _
                                              BoutonsRaccourciObjets(i).PositionX, _
                                              BoutonsRaccourciObjets(i).PositionY + AffBoutonsRaccourciObjets.Hauteur - 5, _
                                              SurfBack, _
                                              AffBoutonsRaccourciObjets.Largeur, _
                                              5, _
                                              &HFF0000    'Bleu.
                    End If
                    End With
                End If
            End If
            Afficher_Duree_Recharge BoutonsRaccourciObjets(i).PositionX, BoutonsRaccourciObjets(i).PositionY, k, Perso, SurfBack
        Next j
    
        If BoutonsRaccourciObjetsMessages(i) <> "" Then
            SurfBack.SetFont AffDegat.FontDegats
            SurfBack.SetForeColor &HFFFFFF
            AffMessage.Afficher BoutonsRaccourciObjets(i).PositionX + 3, _
                                BoutonsRaccourciObjets(i).PositionY + 18, _
                                BoutonsRaccourciObjetsMessages(i), _
                                SurfBack
            SurfBack.SetFont AffDegat.FontNormale
        End If
    Next i
    SurfBack.SetForeColor &H0&
End Sub

Public Sub Afficher_Duree_Recharge(ByVal X As Integer, ByVal Y As Integer, _
                                   ByVal NumeroObjet As Integer, ByVal Perso As ClsJeuPerso, _
                                   ByRef SurfBack As IDirectDrawSurface7)
    Dim i As Single
    'Dim TempX As Single, TempY As Single, TempR As Single
    'Dim Temp As Single
    Dim Message As String
    Dim Temp As Long
    Dim Min As Long
    With Parametres
    If .Objet_DureeRecharge(NumeroObjet) > 0 Then
        Temp = Perso.Objet_Equipe_BonusSup(Perso.Definir_Niveau_CompetenceObjets(.Objet_Competence(NumeroObjet)), .Objet_RechargesMultiples(NumeroObjet))
        Min = Int(Temp - Perso.EffetRecharge_TempRestant(NumeroObjet) / .Objet_DureeRecharge(NumeroObjet))
        If Perso.EffetRecharge_TempRestant(NumeroObjet) > 0 Then
            'Temp = (Parametres.Objet_DureeRecharge(NumeroObjet) - Perso.EffetRecharge_TempRestant(NumeroObjet)) / Parametres.Objet_DureeRecharge(NumeroObjet) * 6.28
            'SurfBack.SetForeColor &HFFFFFF
            'SurfBack.SetFillStyle 5
            'TempR = AffBoutonsRaccourciObjets.Largeur / 2
            'TempX = X + TempR
            'TempY = Y + TempR
            'SurfBack.SetFillColor &HE0E0E0
            'SurfBack.DrawCircle TempX, TempY, TempR
            'For i = 0 To Temp Step 0.2
            '    SurfBack.DrawLine TempX, TempY, TempX + Cos(i - 1.57) * TempR, TempY + Sin(i - 1.57) * TempR
            'Next i
            'SurfBack.SetFillStyle 0
            If Min >= Temp Or Min <= 0 Then
                AffEffet.Afficher Parametres.EffetDureeRestante, _
                                  Int((Perso.EffetRecharge_TempRestant(NumeroObjet) Mod .Objet_DureeRecharge(NumeroObjet)) / Parametres.Objet_DureeRecharge(NumeroObjet) * Parametres.Effet_NombreEtapes(Parametres.EffetDureeRestante)), _
                                  X, Y, SurfBack
            Else
                AffEffet.Afficher Parametres.EffetDureeRestante2, _
                                  Int((Perso.EffetRecharge_TempRestant(NumeroObjet) Mod .Objet_DureeRecharge(NumeroObjet)) / Parametres.Objet_DureeRecharge(NumeroObjet) * Parametres.Effet_NombreEtapes(Parametres.EffetDureeRestante2)), _
                                  X, Y, SurfBack
            End If
            
        End If
        'Affiche le nombre de fois que l'on peut le lancer au milieu.
        If .Objet_RechargesMultiples(NumeroObjet) > 0 Then
            If Temp > 1 Then
                Message = Min & "/" & Temp
                SurfBack.SetFont AffDegat.FontDegats
                AffMessage.Afficher X + 15 - Len(Message) * 2, Y + 9, Message, SurfBack
                SurfBack.SetFont AffDegat.FontNormale
            End If
        End If
    End If
    End With
End Sub

Private Sub Afficher_Effets_Temporaires(ByVal Perso As ClsJeuPerso, _
                                        ByRef SurfBack As IDirectDrawSurface7)
    Dim i As Integer
    Actualiser_Effets_Temporaires Perso
    With Perso
    If .Nombre_EffetsTemp > 0 Then
        For i = 1 To .Nombre_EffetsTemp
            If EffetTemporatireTransparent And _
               .Definir_EffetTemp_Duree(i) < EffetTemporatireDureeEffacement And Not .Definir_EffetTemp_Duree(i) = -1 And .Vivant Then
                AffObjets.Afficher BoutonsEffetsTemporaires(i).PositionX, _
                                   BoutonsEffetsTemporaires(i).PositionY, _
                                   .Definir_EffetTemp_Numero(i), _
                                   SurfBack, _
                                   True, _
                                   (Cos((EffetTemporatireDureeEffacement - (.Definir_EffetTemp_Duree(i) + Compteur * 10) + 0.1) / 1.5) + 1) / 2, _
                                   1
            Else
                AffObjets.Afficher BoutonsEffetsTemporaires(i).PositionX, _
                                   BoutonsEffetsTemporaires(i).PositionY, _
                                   .Definir_EffetTemp_Numero(i), _
                                   SurfBack, _
                                   , , 1
            End If
            If .Definir_EffetTemp_Duree(i) <> -1 Then
                Afficher_Barre_Statut .Definir_EffetTemp_Duree(i), .Definir_EffetTemp_Duree(i), .Definir_EffetTemp_DureeMax(i), _
                                      BoutonsEffetsTemporaires(i).PositionX + 2, _
                                      BoutonsEffetsTemporaires(i).PositionY + AffBoutonsRaccourciObjets.Hauteur - 7, _
                                      SurfBack, _
                                      AffBoutonsRaccourciObjets.Largeur - 4, _
                                      5, _
                                      &HFFFFFF
            End If

        Next i
    End If
    End With
End Sub

Private Sub Afficher_BoutonsInfos(ByRef SurfBack As IDirectDrawSurface7)
    Dim i As Integer
    For i = 0 To NombreBoutonsInfos - 1
        If i = 8 Then
            If AfficherObjectifs Then
                AffBoutonsInfos.Afficher BoutonsInfos(i).PositionX, _
                                         BoutonsInfos(i).PositionY, _
                                         i, _
                                         IIf(NouveauxObjectifs And Not BoutonsInfos(i).Enabled, 2, -CDbl(BoutonsInfos(i).Enabled)), _
                                         SurfBack, _
                                         BoutonInfosTransparents
            End If
        Else
            AffBoutonsInfos.Afficher BoutonsInfos(i).PositionX, _
                                     BoutonsInfos(i).PositionY, _
                                     i, _
                                     -CDbl(BoutonsInfos(i).Enabled), _
                                     SurfBack, _
                                     BoutonInfosTransparents
        End If
    Next i
End Sub

Public Sub Activer_BoutonsRadiaux(ByRef X As Single, _
                                  ByRef Y As Single, _
                                  ByVal NoBouton As Integer)
    
    If AfficherBoutonsRadiaux Then
        If ActiverBoutonsRadiaux Then
            BoutonsRadiauxCategorie = BoutonsRadiauxCategorie + 1
            If BoutonsRadiauxCategorie > 1 Then
                BoutonsRadiauxCategorie = 0
                'ActiverBoutonsRadiaux = False
            End If
        Else
            ActiverBoutonsRadiaux = True
        End If
        
        If ActiverBoutonsRadiaux Then
            Positionner_BoutonsRadiaux X, Y
            
            'On mémorise le centre.
            BoutonsRadiauxCentreX = X
            BoutonsRadiauxCentreY = Y
            'On place le curseur sur le bouton sélectionné.
            If BoutonsRadiauxCategorie = 0 Then
                X = BoutonsRadiauxInfos(NoBouton).PositionX + AffBoutonsInfos.Largeur / 2
                Y = BoutonsRadiauxInfos(NoBouton).PositionY + AffBoutonsInfos.Hauteur / 2
            End If
        End If
    End If
End Sub

Private Sub Positionner_BoutonsRadiaux(ByVal X As Single, _
                                       ByVal Y As Single)
    Dim i, j, k As Integer
    Dim Temp As Integer
    Dim Angle As Single
    Select Case BoutonsRadiauxCategorie
    Case 0:
        Angle = 3.14 / NombreBoutonsRadiauxInfo * 2
        Temp = Sqr(BoutonsRadiauxInfoRayon ^ 2 / 2)
        For i = 0 To NombreBoutonsRadiauxInfo - 1
             BoutonsRadiauxInfos(i).PositionX = X + Cos((Angle * i) + BoutonsRadiauxAngle(BoutonsRadiauxCategorie)) * BoutonsRadiauxInfoRayon - AffBoutonsInfos.Largeur / 2
             BoutonsRadiauxInfos(i).PositionY = Y + Sin((Angle * i) + BoutonsRadiauxAngle(BoutonsRadiauxCategorie)) * BoutonsRadiauxInfoRayon - AffBoutonsInfos.Hauteur / 2
        Next i
    Case 1:
        j = -1
        For i = 0 To NombreBoutonsConstruire - 1
            If i = 0 Or k > NombreBoutonsRadiauxConstruire + NombreBoutonsRadiauxConstruireSup * j - 1 Then
                j = j + 1
                k = 0
                Angle = 3.14 / (NombreBoutonsRadiauxConstruire + NombreBoutonsRadiauxConstruireSup * j) * 2
                Temp = Sqr((BoutonsRadiauxConstruireRayon + BoutonsRadiauxConstruireRayonSup * j) ^ 2 / 2)
            End If
             BoutonsRadiauxConstruire(i).PositionX = _
                X + Cos((Angle * k) + BoutonsRadiauxAngle(BoutonsRadiauxCategorie)) * _
                (BoutonsRadiauxConstruireRayon + BoutonsRadiauxConstruireRayonSup * j) - AffBoutonsConstruire(i).Largeur / 2
             BoutonsRadiauxConstruire(i).PositionY = _
                Y + Sin((Angle * k) + BoutonsRadiauxAngle(BoutonsRadiauxCategorie)) * _
                (BoutonsRadiauxConstruireRayon + BoutonsRadiauxConstruireRayonSup * j) - AffBoutonsConstruire(i).Hauteur / 2
            k = k + 1
        Next i
    End Select
End Sub

Public Sub Pivoter_BoutonsRadiaux(ByVal Z As Single)
    BoutonsRadiauxAngle(BoutonsRadiauxCategorie) = BoutonsRadiauxAngle(BoutonsRadiauxCategorie) + 3.14 * Z / 2
    If BoutonsRadiauxAngle(BoutonsRadiauxCategorie) < 0 Then
        BoutonsRadiauxAngle(BoutonsRadiauxCategorie) = BoutonsRadiauxAngle(BoutonsRadiauxCategorie) + 6.28
    ElseIf BoutonsRadiauxAngle(BoutonsRadiauxCategorie) > 6.28 Then
        BoutonsRadiauxAngle(BoutonsRadiauxCategorie) = BoutonsRadiauxAngle(BoutonsRadiauxCategorie) - 6.28
    End If
    Positionner_BoutonsRadiaux BoutonsRadiauxCentreX, BoutonsRadiauxCentreY
End Sub

Private Sub Afficher_BoutonsRadiaux(ByVal SourisX As Integer, _
                                    ByVal SourisY As Integer, _
                                    ByRef SurfBack As IDirectDrawSurface7)
    Dim i  As Integer
    SurfBack.SetForeColor &HFFFFFF
    SurfBack.setDrawWidth 3
    Select Case BoutonsRadiauxCategorie
    Case 0:
        For i = 0 To NombreBoutonsRadiauxInfo - 1
            AffBoutonsInfos.Afficher BoutonsRadiauxInfos(i).PositionX, _
                                     BoutonsRadiauxInfos(i).PositionY, _
                                     i, _
                                     -CDbl(BoutonsInfos(i).Enabled), _
                                     SurfBack, _
                                     BoutonInfosTransparents
        Next i
        For i = 0 To NombreBoutonsRadiauxInfo - 1
            If BoutonsRadiauxInfos(i).MouseOver Then Afficher_BoutonsRadiaux_Entourer BoutonsRadiauxInfos(i), SurfBack
        Next i
    Case 1:
        For i = 0 To NombreBoutonsConstruire - 1
            AffBoutonsConstruire(i).Afficher BoutonsRadiauxConstruire(i).PositionX, _
                                             BoutonsRadiauxConstruire(i).PositionY, _
                                             0, _
                                             Etat(i), _
                                             SurfBack, _
                                             True
        Next i
        For i = 0 To NombreBoutonsConstruire - 1
            If BoutonsRadiauxConstruire(i).MouseOver Then Afficher_BoutonsRadiaux_Entourer BoutonsRadiauxConstruire(i), SurfBack
        Next i
    End Select
    SurfBack.DrawCircle BoutonsRadiauxCentreX, BoutonsRadiauxCentreY, 2
    SurfBack.DrawLine BoutonsRadiauxCentreX, _
                      BoutonsRadiauxCentreY, _
                      SourisX, _
                      SourisY
    SurfBack.setDrawWidth 1
End Sub

Private Sub Afficher_BoutonsRadiaux_Entourer(ByVal Bouton As ClsIntBouton, _
                                             ByRef SurfBack As IDirectDrawSurface7)
    With Bouton
    SurfBack.DrawRoundedBox .PositionX - 1, .PositionY - 1, .PositionX + .Largeur + 1, .PositionY + .Hauteur + 1, 1, 1
    End With
End Sub

Private Sub Afficher_BoutonsMenu(ByRef SurfBack As IDirectDrawSurface7)
    Dim i As Integer
    If BoutonsInfos(7).Enabled Then
        If Jeu.Campagne Then
            For i = 1 To NombreBoutonsMenu - 1
                AffBoutonsMenu.Afficher BoutonsMenu(i).PositionX - BoutonMenuIntervalX, _
                                        BoutonsMenu(i).PositionY - BoutonMenuIntervalY, _
                                        i, _
                                        -CDbl(BoutonsMenu(i).MouseOver), _
                                        SurfBack, _
                                        BoutonMenuTransparents
            Next i
        Else
            For i = 0 To NombreBoutonsMenu - 1
                AffBoutonsMenu.Afficher BoutonsMenu(i).PositionX, _
                                        BoutonsMenu(i).PositionY, _
                                        i, _
                                        -CDbl(BoutonsMenu(i).MouseOver), _
                                        SurfBack, _
                                        BoutonMenuTransparents
            Next i
        End If
    End If
End Sub

Private Sub Afficher_ModeAuto(ByRef Perso As ClsJeuPerso, ByRef SurfBack As IDirectDrawSurface7)
    If Perso.IA Then
        SurfBack.SetForeColor MessageModeAutoCouleur
        AffMessage.Afficher MessageModeAutoPositionX, MessageModeAutoPositionY, MessageModeAutoMessage, SurfBack
    End If
End Sub

Private Sub Afficher_Action(ByRef Perso As ClsJeuPerso, _
                            ByRef Maison As ClsJeuBatiment, _
                            ByRef Parametres As ClsJeuParam, _
                            ByRef SurfBack As IDirectDrawSurface7)
    Dim Message As String
    With Perso
    'Etiquette de construction.
    
    Select Case .Action
        Case 0: Message = MessagesActions(0)
        Case 1: Message = MessagesActions(1)
        Case 2: Message = MessagesActions(2)
        Case 3: Message = MessagesActions(3)
        Case 6: Message = MessagesActions(4)
        Case 7: Message = MessagesActions(5)
        Case 100: Message = MessagesActions(6)
        Case 101: Message = MessagesActions(7)
        Case 102: Message = MessagesActions(8)
        Case 201: Message = MessagesActions(9)
        Case 202: Message = MessagesActions(10)
        Case 203: Message = MessagesActions(11)
        Case 204: Message = MessagesActions(12)
        Case 205: Message = MessagesActions(13)
        Case 206: Message = MessagesActions(14)
        Case 207: Message = MessagesActions(14)
        Case 300: Message = MessagesActions(15)
        Case 301: Message = MessagesActions(16)
        Case 302: Message = MessagesActions(17)
        Case 400: Message = MessagesActions(18)
        Case 500: Message = MessagesActions(19)
        Case 600: Message = MessagesActions(20)
        Case Else: Message = MessagesActions(21)
    End Select
    
    If .Action = 6 And _
       Maison.Definir_Vie < Maison.MaxVie Then
        Message = Message & _
                  " : " & _
                  Int(Maison.Definir_Vie / Maison.MaxVie * 100) & _
                  "%"
    ElseIf (.Action = 2 Or .Action = 102) And .RessourceCourante >= 0 Then
        Message = Parametres.Ressources_NomCompetence(.RessourceCourante) & _
                  " : " & _
                  Int(.Definir_ressources(.RessourceCourante)) & _
                  " / " & _
                  .Definir_MaxRessources(.RessourceCourante)
    Else
        Select Case .Action
        Case 3: 'Va chercher des ressources.
        Message = Message & " : " & Ressources(.IndiceRessource).Nom
        Case 201:
        Message = Message & " : " & Parametres.Batiment_Nom(Maisons(.IndiceMaison).TypeBatiment)
        Case 300:
        Message = Message & " : "
        If Persos(.IndicePerso).Nom = "" Then
            Message = Message & Parametres.Race_Nom(Persos(.IndicePerso).Race, Persos(.IndicePerso).Feminin)
        Else
            Message = Message & Persos(.IndicePerso).Nom
        End If
        Message = Message & " (" & Parametres.EtiquetteVie & ":" & Int(Persos(.IndicePerso).Definir_Vie / Persos(.IndicePerso).Definir_MaxVie * 100) & "%)"
        Case 400: 'Etiquette de ressurection.
        Message = Message & _
                  " : " & _
                  Int(100 - .DureeRessurection) & _
                  "%"

        Case 500: 'Etiquette de fabrication d'objet
        Message = Message & _
                  " " & LCase(Parametres.Objet_Nom(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, .IndiceObjet))) & _
                  " : " & Int(.EtatFabricationObjet / Parametres.Objet_DureeFabrication(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, .IndiceObjet)) * 100) & _
                  "%"
        Case 600: 'Etiquette de création de service.
        Message = Parametres.Service_NomCompetence(Maison.Service) & _
                  " : " & _
                  Int(Maison.Definir_Stock) & _
                  " / " & _
                  Int(Maison.MaxStock)
        Case Else 'Etiquette simple.
            Message = Message
        End Select
    End If
    End With
    'Debug.Print Maisons(NoPerso).EnConstruction
    SurfBack.SetForeColor MessageActionCouleur
    AffMessage.Afficher MessageActionPositionX, MessageActionPositionY, Message, SurfBack
End Sub

Private Sub Afficher_Vie(ByRef Perso As ClsJeuPerso, ByRef SurfBack As IDirectDrawSurface7)
    SurfBack.SetForeColor 0
    Afficher_Barre_Statut IIf(Perso.Vie > Perso.Definir_Vie, Perso.Definir_Vie, IIf(Perso.Vie > 0, Perso.Vie, 0)), Perso.Definir_Vie, Perso.Definir_MaxVie, ViePositionX + BarresStatutDecalageX, ViePositionY + BarresStatutDecalageY, SurfBack, BarresStatutLargeur, BarresStatutHauteur
    SurfBack.SetForeColor &HFFFFFF
    AffMessage.Afficher ViePositionX, ViePositionY, Parametres.EtiquetteVie & " :", SurfBack
    If AfficherBarresStatutScore Then
        AffMessage.Afficher ViePositionX + BarresStatutDecalageX + BarresStatutScoreDecalageX - Len(Int(Perso.Vie) & " / " & Perso.Definir_MaxVie) * BarresStatutScoreLargeurTexte, ViePositionY + BarresStatutDecalageY + BarresStatutScoreDecalageY, Int(Perso.Vie) & " / " & Perso.Definir_MaxVie, SurfBack, BarresStatutScoreCouleur, True
    End If
End Sub
Private Sub Afficher_Energie(ByRef Perso As ClsJeuPerso, ByRef SurfBack As IDirectDrawSurface7)
    SurfBack.SetForeColor 0
    Afficher_Barre_Statut IIf(Perso.Energie > 0, Perso.Energie, 0), Perso.Definir_Energie, Perso.Definir_MaxEnergie, EnergiePositionX + BarresStatutDecalageX, EnergiePositionY + BarresStatutDecalageY, SurfBack, BarresStatutLargeur, BarresStatutHauteur
    SurfBack.SetForeColor &HFFFFFF
    AffMessage.Afficher EnergiePositionX, EnergiePositionY, Parametres.EtiquetteEnergie & " :", SurfBack
    If AfficherBarresStatutScore Then
        AffMessage.Afficher EnergiePositionX + BarresStatutDecalageX + BarresStatutScoreDecalageX - Len(Int(Perso.Energie) & " / " & Perso.Definir_MaxEnergie) * BarresStatutScoreLargeurTexte, EnergiePositionY + BarresStatutDecalageY + BarresStatutScoreDecalageY, Int(Perso.Energie) & " / " & Perso.Definir_MaxEnergie, SurfBack, BarresStatutScoreCouleur, True
    End If
End Sub
Private Sub Afficher_Magie(ByRef Perso As ClsJeuPerso, ByRef SurfBack As IDirectDrawSurface7)
    SurfBack.SetForeColor 0
    Afficher_Barre_Statut IIf(Perso.Magie > 0, Perso.Magie, 0), Perso.Definir_Magie, Perso.Definir_MaxMagie, MagiePositionX + BarresStatutDecalageX, MagiePositionY + BarresStatutDecalageY, SurfBack, BarresStatutLargeur, BarresStatutHauteur
    SurfBack.SetForeColor &HFFFFFF
    AffMessage.Afficher MagiePositionX, MagiePositionY, Parametres.EtiquetteMagie & " :", SurfBack
    If AfficherBarresStatutScore Then
        AffMessage.Afficher MagiePositionX + BarresStatutDecalageX + BarresStatutScoreDecalageX - Len(Int(Perso.Magie) & " / " & Perso.Definir_MaxMagie) * BarresStatutScoreLargeurTexte, MagiePositionY + BarresStatutDecalageY + BarresStatutScoreDecalageY, Int(Perso.Magie) & " / " & Perso.Definir_MaxMagie, SurfBack, BarresStatutScoreCouleur, True
    End If
End Sub
Private Sub Afficher_Moral(ByRef Perso As ClsJeuPerso, ByRef SurfBack As IDirectDrawSurface7)
    SurfBack.SetForeColor 0
    Afficher_Barre_Statut IIf(Perso.Moral > 0, Perso.Moral, 0), Perso.Definir_Moral, Perso.Definir_MaxMoral, MoralPositionX + BarresStatutDecalageX, MoralPositionY + BarresStatutDecalageY, SurfBack, BarresStatutLargeur, BarresStatutHauteur
    SurfBack.SetForeColor &HFFFFFF
    AffMessage.Afficher MoralPositionX, MoralPositionY, Parametres.EtiquetteMoral & " :", SurfBack
    If AfficherBarresStatutScore Then
        AffMessage.Afficher MoralPositionX + BarresStatutDecalageX + BarresStatutScoreDecalageX - Len(Int(Perso.Moral) & " / " & Perso.Definir_MaxMoral) * BarresStatutScoreLargeurTexte, MoralPositionY + BarresStatutDecalageY + BarresStatutScoreDecalageY, Int(Perso.Moral) & " / " & Perso.Definir_MaxMoral, SurfBack, BarresStatutScoreCouleur, True
    End If
End Sub
Public Sub Afficher_Barre_Statut(ByVal Min1 As Double, _
                                 ByVal Min2 As Double, _
                                 ByVal Max As Double, _
                                 ByVal PositionX As Long, _
                                 ByVal PositionY As Long, _
                                 ByRef SurfBack As IDirectDrawSurface7, _
                                 ByVal Largeur As Long, _
                                 ByVal Hauteur As Long, _
                                 Optional ByVal Couleur As Long, _
                                 Optional Vertical As Boolean)
    If PositionX < 0 Then
        Largeur = Largeur + PositionX
        PositionX = 0
    End If
    If PositionX + Largeur > ScreenWidth Then
        Largeur = ScreenWidth - PositionX
    End If
    If PositionY < 0 Then
        Hauteur = Hauteur + PositionY
        PositionY = 0
    End If
    If PositionY + Hauteur > ScreenHeight Then
        Hauteur = ScreenHeight - PositionY
    End If
    If Largeur > 0 And Hauteur > 0 Then
        With SurfBack
        .setDrawStyle 10
        '.setDrawStyle 5
        'AffMessage.Afficher PositionX, PositionY, Int(Min2) & " / " & Int(Max), SurfBack
        
        PositionX = PositionX + ScreenPositionX
        PositionY = PositionY + ScreenPositionY
        
        'Cadre noir.
        .SetFillColor 0
        '.DrawBox PositionX - 1, PositionY - 1, PositionX + Largeur + 1, PositionY + Hauteur + 1
        .DrawBox PositionX, PositionY, PositionX + Largeur, PositionY + Hauteur
        '.DrawBox PositionX - 1, PositionY, PositionX + Largeur + 1, PositionY + Hauteur
        
        .setDrawStyle 5
        'Cadre blanc.
        .SetFillColor &HFFFFFF
        '.SetForeColor &HFFFFFF
        If Min2 > 0 Then
            If Vertical Then
                '.DrawBox PositionX, PositionY + Hauteur, _
                         PositionX + Largeur, PositionY - Hauteur * Min2 / Max + Hauteur
                .DrawBox PositionX, PositionY + Hauteur, _
                         PositionX + Largeur, PositionY - Hauteur * Min2 / Max + Hauteur
            Else
                '.DrawBox PositionX, PositionY, _
                         PositionX + Largeur * Min2 / Max, PositionY + Hauteur
                .DrawBox PositionX + 1, PositionY + 1, _
                         PositionX + Largeur * Min2 / Max, PositionY + Hauteur
            End If
        End If
        If Couleur = 0 Then
            Select Case Min1 / Max
                Case Is < 0.1: .SetFillColor &H80& 'rouge foncé.
                Case Is < 0.25: .SetFillColor &HFF& 'rouge.
                Case Is < 0.5: .SetFillColor &H80FF&  'Orange.
                Case Else: .SetFillColor &HC000&  'Vert.
            End Select
        Else
            .SetFillColor Couleur 'Couleur en paramètre.
        End If
        If Vertical Then
            '.DrawBox PositionX, PositionY + Hauteur, _
                     PositionX + Largeur, PositionY - Hauteur * Min1 / Max + Hauteur
            .DrawBox PositionX + 1, PositionY + Hauteur, _
                     PositionX + Largeur, PositionY - Hauteur * Min1 / Max + Hauteur
        Else
            '.DrawBox PositionX, PositionY, _
                     PositionX + Largeur * Min1 / Max, PositionY + Hauteur
            .DrawBox PositionX + 1, PositionY + 1, _
                     PositionX + Largeur * Min1 / Max, PositionY + Hauteur
        End If
        
        .SetFillStyle 1
        .setDrawStyle 0
        End With
    End If
End Sub

Private Sub Afficher_Barre_Statut_Auto(ByRef SurfBack As IDirectDrawSurface7)
    
    If AutoRecuperationCarac(0) Then
        SurfBack.SetForeColor BarresStatutContourCouleur2
        Afficher_Contour_Barre_Statut ViePositionX + BarresStatutDecalageX, _
                                      ViePositionY + BarresStatutDecalageY, _
                                      BarresStatutLargeur, _
                                      BarresStatutHauteur, _
                                      AutoRecuperationCaracPcent(0), _
                                      SurfBack
    Else
        SurfBack.SetForeColor BarresStatutContourCouleur1
        Afficher_Contour_Barre_Statut ViePositionX + BarresStatutDecalageX, _
                                      ViePositionY + BarresStatutDecalageY, _
                                      BarresStatutLargeur, _
                                      BarresStatutHauteur, _
                                      100, _
                                      SurfBack
    End If
    If AutoRecuperationCarac(1) Then
        SurfBack.SetForeColor BarresStatutContourCouleur2
        Afficher_Contour_Barre_Statut EnergiePositionX + BarresStatutDecalageX, _
                                      EnergiePositionY + BarresStatutDecalageY, _
                                      BarresStatutLargeur, _
                                      BarresStatutHauteur, _
                                      AutoRecuperationCaracPcent(1), _
                                      SurfBack
    Else
        SurfBack.SetForeColor BarresStatutContourCouleur1
        Afficher_Contour_Barre_Statut EnergiePositionX + BarresStatutDecalageX, _
                                      EnergiePositionY + BarresStatutDecalageY, _
                                      BarresStatutLargeur, _
                                      BarresStatutHauteur, _
                                      100, _
                                      SurfBack
    End If
    If AutoRecuperationCarac(2) Then
        SurfBack.SetForeColor BarresStatutContourCouleur2
        Afficher_Contour_Barre_Statut MagiePositionX + BarresStatutDecalageX, _
                                      MagiePositionY + BarresStatutDecalageY, _
                                      BarresStatutLargeur, _
                                      BarresStatutHauteur, _
                                      AutoRecuperationCaracPcent(2), _
                                      SurfBack
    Else
        SurfBack.SetForeColor BarresStatutContourCouleur1
        Afficher_Contour_Barre_Statut MagiePositionX + BarresStatutDecalageX, _
                                      MagiePositionY + BarresStatutDecalageY, _
                                      BarresStatutLargeur, _
                                      BarresStatutHauteur, _
                                      100, _
                                      SurfBack
    End If
    If AutoRecuperationCarac(3) Then
        SurfBack.SetForeColor BarresStatutContourCouleur2
        Afficher_Contour_Barre_Statut MoralPositionX + BarresStatutDecalageX, _
                                      MoralPositionY + BarresStatutDecalageY, _
                                      BarresStatutLargeur, _
                                      BarresStatutHauteur, _
                                      AutoRecuperationCaracPcent(3), _
                                      SurfBack
    Else
        SurfBack.SetForeColor BarresStatutContourCouleur1
        Afficher_Contour_Barre_Statut MoralPositionX + BarresStatutDecalageX, _
                                      MoralPositionY + BarresStatutDecalageY, _
                                      BarresStatutLargeur, _
                                      BarresStatutHauteur, _
                                      100, _
                                      SurfBack
    End If
End Sub
Public Sub Afficher_Contour_Barre_Statut(ByVal PositionX As Long, _
                                         ByVal PositionY As Long, _
                                         ByVal Largeur As Integer, _
                                         ByVal Hauteur As Integer, _
                                         ByVal PCent As Single, _
                                         ByRef SurfBack As IDirectDrawSurface7)
    SurfBack.DrawBox PositionX - BarresStatutContourEpaisseur, _
                     PositionY - BarresStatutContourEpaisseur, _
                     PositionX + Largeur + BarresStatutContourEpaisseur, _
                     PositionY + Hauteur + BarresStatutContourEpaisseur
    'Affiche la jauge qui indique le niveau d'auto restauration.
    SurfBack.DrawBox PositionX + (Largeur * PCent), _
                     PositionY, _
                     PositionX + (Largeur * PCent) + BarresStatutContourEpaisseur, _
                     PositionY + Hauteur
End Sub
                          
Public Sub Afficher_Info_Bulles_Terrain(ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Message As String, _
                                        ByRef SurfBack As IDirectDrawSurface7)
    'On utilise temporairement le bouton Deplacement0 pour afficher l'info bulle.
    BoutonsDeplacement(0).InfoBulle = Message
    BoutonsDeplacement(0).Afficher_Info_Bulle X, Y, AffMessage, SurfBack, True
    BoutonsDeplacement(0).InfoBulle = BoutonDeplacementInfoBulle(0)
End Sub

'Public Function Souris_Sur_Interface(ByVal x As Long, _
                                     ByVal y As Long, _
                                     ByRef Ecran As ClsPerEcran, _
                                     ByRef Perso As ClsJeuPerso, _
                                     ByRef Maison As ClsJeuBatiment, _
                                     ByRef SaMaison As ClsJeuBatiment, _
                                     ByRef Fief As ClsJeuFief, _
                                     ByRef Parametres As ClsJeuParam, _
                                     ByRef SurfBack As IDirectDrawSurface7) As Boolean
    'Dim NoBouton As Long
    'Dim i As Integer
    'Souris_Sur_Interface = False
    'If Interieur_Boutons(x, y, NoBouton, Perso, Maison, SaMaison, Fief, Parametres, SurfBack, False) Then
    '    Souris_Sur_Interface = True
    'Else

    'End If
'End Function

'Public Function Clique_Bouton(ByVal x As Long, _
                              ByVal y As Long, _
                              ByRef NoBouton As Long, _
                              ByRef Perso As ClsJeuPerso, _
                              ByRef Maison As ClsJeuBatiment, _
                              ByRef SaMaison As ClsJeuBatiment, _
                              ByRef Fief As ClsJeuFief, _
                              ByRef Parametres As ClsJeuParam, _
                              ByRef SurfBack As IDirectDrawSurface7, _
                              Optional ByVal CliqueDroit As Boolean, _
                              Optional ByVal ToucheClavier As Integer) As Boolean
'    Clique_Bouton = Interieur_Boutons(x, y, NoBouton, Perso, Maison, SaMaison, Fief, Parametres, SurfBack, True, CliqueDroit, ToucheClavier)
'End Function

Private Sub Afficher_Stastistiques(ByRef Message As ClsIntMessage, _
                                   ByRef SurfBack As IDirectDrawSurface7)
    If Not ComReseau.Connecte Then
    Message.Afficher StastistiquesPositionX, _
                     StastistiquesPositionYTemp, _
                     StastistiquesMessageTemp, _
                     SurfBack, CouleurStastistiques, True
    End If
End Sub

'Interface de la fenêtre dialogue.
Private Sub Afficher_Dialogue(ByVal X As Long, ByVal Y As Long, _
                              ByVal Largeur As Integer, _
                              ByVal Hauteur As Integer, _
                              ByRef SurfBack As IDirectDrawSurface7)
    'Affiche le cadre.
    Dim l As Integer, H  As Integer
    Dim i As Integer
    l = 480
    H = 20
    SurfBack.SetForeColor 0
    'Surfback.SetFillStyle 0
    SurfBack.SetFillColor &HE0E0E0
    SurfBack.DrawBox (Largeur - l) / 2, (Hauteur - H) / 2 + 1, (Largeur - l) / 2 + l, (Hauteur - H) / 2 + H
    AffMessage.Afficher (Largeur - l) / 2 + 4, (Hauteur - H) / 2 + 2, DialogueTexte1 & IIf(DialogueCurseurVisible, "  ", "<") & DialogueTexte2, SurfBack, &HFF0000
    If MessagesPrecedentsAffiche And MessagesPrecedents.Count > 1 Then
        'Affiche les messages précédents.
        For i = MessagesPrecedentsPosition - 1 To 1 Step -1
            AffMessage.Afficher (Largeur - l) / 2 + 4, (Hauteur - H) / 2 + 2 - AffMessage.Hauteur_Caractere * (MessagesPrecedentsPosition - i), MessagesPrecedents(i), SurfBack
        Next i
        'Affiche les messages suivants.
        For i = MessagesPrecedentsPosition + 1 To MessagesPrecedents.Count
            AffMessage.Afficher (Largeur - l) / 2 + 4, (Hauteur - H) / 2 + 2 + AffMessage.Hauteur_Caractere * (i - MessagesPrecedentsPosition), MessagesPrecedents(i), SurfBack
        Next i
    End If
    SurfBack.SetForeColor &HFFFFFF
End Sub
Public Sub Dialogue_Valider(ByVal Perso As ClsJeuPerso)
    Dim Message As String
    Dim i As Long: Dim j As Long
    Message = DialogueTexte1 & DialogueTexte2
    If Message <> "" Then
        If LCase(Left(Message, 7)) = "_heure=" Then
            If IsNumeric(Right(Message, Len(Message) - 7)) Then
                AffTemps.Heures = Val(Right(Message, Len(Message) - 7))
                AffTemps.Actualiser_Temps
            End If
        ElseIf LCase(Message) = "_regner sans partage" Then
            Cmd_Perso_Hostile Perso.Numero
        ElseIf LCase(Message) = "_test" Then
            Perso.Definir_Argent = Perso.Definir_Argent + 10000
            For i = 0 To Parametres.NombreRessources - 1
                Perso.Definir_ressources(i) = Perso.Definir_MaxRessources(i)
            Next i
            For i = 0 To Perso.Nombre_Objets_Inventaire - 1
                j = Rnd * (Parametres.NombreObjets - 1)
                If Not Parametres.Objet_Secret(j) Then
                    Perso.Objet_Ajouter_Inventaire j, Parametres, Commentaires, False
                End If
            Next i
        Else
            Cmd_Perso_EnvoyerCommentaire Perso.Numero, Message & String(Len(Message), " ")
        End If
        MessagesPrecedents.Add Message
        If MessagesPrecedentsAffiche Then MessagesPrecedents.Remove MessagesPrecedentsPosition
        MessagesPrecedentsPosition = MessagesPrecedents.Count + 1
        If MessagesPrecedents.Count > 10 Then MessagesPrecedents.Remove 1
    End If
    Dialogue_Annuler
End Sub
Public Sub Dialogue_Annuler()
    AfficherDialogue = False
    MessagesPrecedentsAffiche = False
    DialogueTexte1 = ""
    DialogueTexte2 = ""
End Sub
Public Sub Dialogue_Message(ByVal Suivant As Boolean)
    If MessagesPrecedents.Count > 0 Then
        If Suivant Then
            If MessagesPrecedentsPosition < MessagesPrecedents.Count Then
                MessagesPrecedentsPosition = MessagesPrecedentsPosition + 1
            Else
                MessagesPrecedentsPosition = MessagesPrecedents.Count
            End If
            DialogueTexte1 = MessagesPrecedents(MessagesPrecedentsPosition)
            DialogueTexte2 = ""
            MessagesPrecedentsAffiche = True
        Else
            If MessagesPrecedentsPosition > 1 Then
                MessagesPrecedentsPosition = MessagesPrecedentsPosition - 1
                If MessagesPrecedentsPosition > MessagesPrecedents.Count Then
                    MessagesPrecedentsPosition = MessagesPrecedents.Count
                End If
                DialogueTexte1 = MessagesPrecedents(MessagesPrecedentsPosition)
                DialogueTexte2 = ""
                MessagesPrecedentsAffiche = True
            End If
        End If
    End If
End Sub
Public Sub Dialogue_Effacer(ByVal Suivant As Boolean)
    If Suivant Then
        If DialogueTexte2 <> "" Then
            DialogueTexte2 = Right(DialogueTexte2, Len(DialogueTexte2) - 1)
        End If
    Else
        If DialogueTexte1 <> "" Then
            DialogueTexte1 = Left(DialogueTexte1, Len(DialogueTexte1) - 1)
        End If
    End If
End Sub
Public Sub Dialogue_Curseur(ByVal Position As Integer)
    Select Case Position
    Case 0: 'Retour au début.
        'DialoguePositionCurseur = 0
        DialogueTexte2 = DialogueTexte1 & DialogueTexte2
        DialogueTexte1 = ""
    Case 1: 'Revient d'une case en arrière.
        'DialoguePositionCurseur = DialoguePositionCurseur = -1
        'If DialoguePositionCurseur < 0 Then DialoguePositionCurseur = 0
        If DialogueTexte1 <> "" Then
            DialogueTexte2 = Right(DialogueTexte1, 1) & DialogueTexte2
            DialogueTexte1 = Left(DialogueTexte1, Len(DialogueTexte1) - 1)
        End If
    Case 2: 'Avance d'une case.
        'DialoguePositionCurseur = DialoguePositionCurseur + 1
        'If DialoguePositionCurseur > Len(DialogueTexte) Then DialoguePositionCurseur = Len(DialogueTexte)
        If DialogueTexte2 <> "" Then
            DialogueTexte1 = DialogueTexte1 & Left(DialogueTexte2, 1)
            DialogueTexte2 = Right(DialogueTexte2, Len(DialogueTexte2) - 1)
        End If
    Case 3: 'Va à la fin.
        'DialoguePositionCurseur = Len(DialogueTexte)
        DialogueTexte1 = DialogueTexte1 & DialogueTexte2
        DialogueTexte2 = ""
    End Select
End Sub
Public Sub Dialogue_AjouterLettre(ByVal Lettre As String)
    'Dim Temp1 As String
    'Dim Temp2 As String
    If Lettre <> "" Then
        'Temp1 = Left(DialogueTexte, DialogueCurseurPosition)
        'Temp2 = Right(DialogueTexte, Len(DialogueTexte) - DialogueCurseurPosition)
        'DialogueTexte = Temp1 & Temp2 & Lettre
        DialogueTexte1 = DialogueTexte1 & Lettre
    End If
End Sub

Public Sub ActualiserRapide(ByVal Perso As ClsJeuPerso, ByVal ArmeeSelectionne As Boolean)
    If MasquerAutoBoutonsConstruire Then Deplacer_BoutonsConstruire
    If MasquerAutoBoutonsDeplacement Then Deplacer_BoutonsDeplacement ArmeeSelectionne
End Sub

Public Sub ActualiserLent(ByVal Perso As ClsJeuPerso, _
                          ByRef Maison As ClsJeuBatiment, _
                          ByRef Fief As ClsJeuFief, _
                          ByVal Parametres As ClsJeuParam, _
                          ByRef Messages As ClsJeuMessages)
    'Actualisation toutes les secondes.
    'Dim Temp As Double
    Compteur = Compteur + 1 / Jeu.Definir_Vitesse
    If Compteur >= 1 Then
        'Temp = Timer
        Compteur = 0
        If Not MasquerAutoBoutonsConstruire Or ForcerCurseurSurBoutonConstruire Or ActiverBoutonsRadiaux And BoutonsRadiauxCategorie = 1 Then
            Actualiser_BoutonsConstruire Perso, Maison, Fief, Parametres
        End If
        If Not ComReseau.Connecte Then Actualiser_Stastistiques Perso
        Actualiser_Raccourcis_Objets Perso
        Messages.Actualiser_Messages Perso, Fief, Parametres
        DialogueCurseurVisible = Not DialogueCurseurVisible
        If Fenetres(3).Visible Then Actualiser_Competences Perso, Maison, Parametres
        'Messages.Ajouter_Message Format(Timer - Temp, "0.00")
    ElseIf ActualiserRaccourciObjets Then
        Actualiser_Raccourcis_Objets Perso
    End If
    'Messages(0).Ajouter_Message Compteur
End Sub

Private Sub Actualiser_Effets_Temporaires(ByVal Perso As ClsJeuPerso)
    Dim i As Long
    Dim Bouton As ClsIntBouton
    'Set Bouton = New ClsIntBouton
    While BoutonsEffetsTemporaires.Count > 0
        BoutonsEffetsTemporaires.Remove 1
    Wend
    If AfficherRaccourciObjets Then
        For i = 0 To Perso.Nombre_EffetsTemp - 1
            Set Bouton = New ClsIntBouton
            Bouton.PositionX = EffetTemporatirePositionX + EffetTemporatireIntervalX * (i + NombreRaccourciObjets)
            Bouton.PositionY = EffetTemporatirePositionY + EffetTemporatireIntervalY * (i + NombreRaccourciObjets)
            BoutonsEffetsTemporaires.Add Bouton
            Set Bouton = Nothing
        Next i
    Else
        For i = 0 To Perso.Nombre_EffetsTemp - 1
            Set Bouton = New ClsIntBouton
            Bouton.PositionX = EffetTemporatirePositionX + EffetTemporatireIntervalX * i
            Bouton.PositionY = EffetTemporatirePositionY + EffetTemporatireIntervalY * i
            BoutonsEffetsTemporaires.Add Bouton
            Set Bouton = Nothing
        Next i
    End If
End Sub

Private Sub Actualiser_Stastistiques(ByVal Perso As ClsJeuPerso)
    Dim i As Integer
    Dim NbFiefsActifs As Integer
    NbFiefsActifs = 0
    StastistiquesMessageTemp = ""
    For i = 0 To UBound(Fiefs())
        If Fiefs(i).NombreMorts < Fiefs(i).NombreCitoyens Or _
           Chateaux(i).Vie > 0 Then
            NbFiefsActifs = NbFiefsActifs + 1
            StastistiquesMessageTemp = StastistiquesMessageTemp & AffMessage.Caractere_RetourLigne & _
                             IIf(Fiefs(i).NombreCitoyens = Fiefs(i).NombreMorts, "&", "") & _
                             "(" & Fiefs(i).NumeroEquipe & ") " & _
                             Fiefs(i).Nom & " " & _
                             Fiefs(i).NombreCitoyens - Fiefs(i).NombreMorts & "/" & Fiefs(i).NombreCitoyens & _
                             IIf(Chateaux(i).Visible, " - " & Format((Chateaux(i).Vie / Chateaux(i).MaxVie) * 100, "00") & "%", "")
        End If
    Next i
    StastistiquesPositionYTemp = StastistiquesPositionY - AffMessage.Hauteur_Caractere * NbFiefsActifs
End Sub

Public Function Actions_Automatiques(ByRef Perso As ClsJeuPerso, _
                                     ByRef Maison As ClsJeuBatiment, _
                                     ByVal Parametres As ClsJeuParam)
    'Le joueur récupère automatiquement ses caractéristiques.
    If Perso.Vivant Then
        If AutoRecuperation_Carac(0) Then
            'Perso.Recuperer_AutoVie Parametres, AutoRecuperationCaracPcent(0)
            If Perso.Definir_Vie < Perso.Definir_MaxVie * AutoRecuperationCaracPcent(0) Then
                'Perso.Recuperer_Vie Parametres
                Cmd_Perso_RecupererCarac Perso.Numero, 0
            End If
        End If
        If AutoRecuperation_Carac(1) Then
            'Perso.Recuperer_AutoEnergie Parametres, AutoRecuperationCaracPcent(1)
            If Perso.Definir_Energie < Perso.Definir_MaxEnergie * AutoRecuperationCaracPcent(1) Then
                'Perso.Recuperer_energie Parametres
                Cmd_Perso_RecupererCarac Perso.Numero, 1
            End If
        End If
        If AutoRecuperation_Carac(2) Then
            'Perso.Recuperer_AutoMagie Parametres, AutoRecuperationCaracPcent(2)
            If Perso.Definir_Magie < Perso.Definir_MaxMagie * AutoRecuperationCaracPcent(2) Then
                'Perso.Recuperer_magie Parametres
                Cmd_Perso_RecupererCarac Perso.Numero, 2
            End If
        End If
        If AutoRecuperation_Carac(3) Then
            'Perso.Recuperer_AutoMoral Parametres, AutoRecuperationCaracPcent(3)
            If Perso.Definir_Moral < Perso.Definir_MaxMoral * AutoRecuperationCaracPcent(3) Then
                'Perso.Recuperer_moral Parametres
                Cmd_Perso_RecupererCarac Perso.Numero, 3
            End If
        End If
    End If
    'Le joueur fabrique automatiquement ses objets en file d'attente.
    If Perso.DansSaMaison Then
        If Perso.Inactif And FileFabricationObjets.Count > 0 Then
            If Perso.PeutConstruire_Objet(FileFabricationObjets.Item(1), Maison, Parametres) Then
                'Perso.Creer_Objet FileFabricationObjets.Item(1), Maison, Parametres
                Cmd_Perso_CreerObjet Perso.Numero, FileFabricationObjets.Item(1)
            End If
            FileFabricationObjets.Remove 1
        End If
    Else
        'Vide la file d'attente.
        Collection_Vider FileFabricationObjets
    End If
    
    'Le joueur utilise automatiquement les services.
    If AutoBatimentService And _
       (Perso.DansUneMaison Or Perso.DansSaMaison) And _
       Maison.Service Then
            If Maison.Service_Utilisable(Perso, Parametres) Then
                'Maison.Service_Utiliser Perso, Parametres
                Cmd_Perso_UtiliserService Perso.Numero, Maison.Numero
            End If
    End If
End Function

Private Sub Deplacer_BoutonsConstruire()
    Dim i As Integer
    If CurseurSurBoutonConstruire Or ForcerCurseurSurBoutonConstruire Then
        If BoutonsConstruire(0).PositionX <> BoutonConstruirePositionX Or _
           BoutonsConstruire(0).PositionY <> BoutonConstruirePositionY Then
            For i = 0 To NombreBoutonsConstruire - 1
                BoutonsConstruire(i).PositionX = BoutonsConstruire(i).PositionX - MasquerAutoBoutonsConstruireDecalageX / MasquerAutoBoutonsConstruireDuree
                BoutonsConstruire(i).PositionY = BoutonsConstruire(i).PositionY - MasquerAutoBoutonsConstruireDecalageY / MasquerAutoBoutonsConstruireDuree
            Next i
        End If
    Else
        If BoutonsConstruire(0).PositionX <> BoutonConstruirePositionX + MasquerAutoBoutonsConstruireDecalageX Or _
           BoutonsConstruire(0).PositionY <> BoutonConstruirePositionY + MasquerAutoBoutonsConstruireDecalageY Then
            For i = 0 To NombreBoutonsConstruire - 1
                BoutonsConstruire(i).PositionX = BoutonsConstruire(i).PositionX + MasquerAutoBoutonsConstruireDecalageX / MasquerAutoBoutonsConstruireDuree
                BoutonsConstruire(i).PositionY = BoutonsConstruire(i).PositionY + MasquerAutoBoutonsConstruireDecalageY / MasquerAutoBoutonsConstruireDuree
            Next i
        End If
    End If
End Sub
Private Sub Deplacer_BoutonsDeplacement(ByVal ArmeeSelectionne As Boolean)
    Dim i As Integer
    If CurseurSurBoutonDeplacement Or ForcerCurseurSurBoutonDeplacement Or ArmeeSelectionne Then
        'Les boutons de déplacement apparaissent.
        If BoutonsDeplacement(0).PositionX <> BoutonDeplacementPositionX Or _
           BoutonsDeplacement(0).PositionY <> BoutonDeplacementPositionY Then
            For i = 0 To NombreBoutonsDeplacement - 1 + NombreBoutonsOrdre
                BoutonsDeplacement(i).PositionX = BoutonsDeplacement(i).PositionX - MasquerAutoBoutonsDeplacementDecalageX / MasquerAutoBoutonsDeplacementDuree
                BoutonsDeplacement(i).PositionY = BoutonsDeplacement(i).PositionY - MasquerAutoBoutonsDeplacementDecalageY / MasquerAutoBoutonsDeplacementDuree
            Next i
        End If
    Else
        'Les boutons de déplacement disparaissent sur le coté.
        If BoutonsDeplacement(0).PositionX <> BoutonDeplacementPositionX + MasquerAutoBoutonsDeplacementDecalageX Or _
           BoutonsDeplacement(0).PositionY <> BoutonDeplacementPositionY + MasquerAutoBoutonsDeplacementDecalageY Then
            For i = 0 To NombreBoutonsDeplacement - 1 + NombreBoutonsOrdre
                BoutonsDeplacement(i).PositionX = BoutonsDeplacement(i).PositionX + MasquerAutoBoutonsDeplacementDecalageX / MasquerAutoBoutonsDeplacementDuree
                BoutonsDeplacement(i).PositionY = BoutonsDeplacement(i).PositionY + MasquerAutoBoutonsDeplacementDecalageY / MasquerAutoBoutonsDeplacementDuree
            Next i
        End If
    End If
End Sub

Public Function Interieur_Boutons(ByVal NoJoueur As Integer, _
                                  ByRef X As Single, ByRef Y As Single, ByRef Z As Single, _
                                  ByRef NoBouton As Long, _
                                  ByRef NoBatiment As Integer, _
                                  ByRef Perso As ClsJeuPerso, _
                                  ByRef Maison As ClsJeuBatiment, _
                                  ByRef SaMaison As ClsJeuBatiment, _
                                  ByRef Fief As ClsJeuFief, _
                                  ByRef Parametres As ClsJeuParam, _
                                  ByRef SurfBack As IDirectDrawSurface7, _
                                  ByRef CliqueGauche As Boolean, _
                                  ByRef CliqueDroit As Boolean, _
                                  ByRef CliqueEnfonceGauche As Boolean, _
                                  ByRef CliqueEnfonceMilieu As Boolean, _
                                  ByRef CliqueEnfonceDroit As Boolean, _
                                  ByRef Touche As Integer, _
                                  ByVal Ctrl As Boolean, _
                                  ByVal Shift As Boolean, _
                                  ByVal Alt As Boolean, _
                                  ByRef ArmeeSelectionnee As Boolean, _
                                  ByRef SoldatsSelectionnes As Collection, _
                                  ByVal TouchesRaccourciEquipement As Boolean, _
                                  ByVal DecalageX As Long, ByVal DecalageY As Long, _
                                  ByVal Largeur As Integer, ByVal Hauteur As Integer) As Boolean
    Dim i As Long, j As Long
    NoBouton = 0
    'NoBatiment = 0
    CurseurSurBarreStatut = False
    CurseurSurObjet = False
    
    'Fenêtre de confirmation.
    If FenetreConfirmation.Affiche Then
        For i = 0 To FenetreConfirmation.Definir_LBoutonsChoix
            If FenetreConfirmation.Interieur_Boutons(i, Perso, X, Y) Then
                If CliqueGauche Or CliqueDroit Then
                    Select Case i
                    Case 0: FenetreConfirmation.Valider Perso
                    Case 1: FenetreConfirmation.Annuler Me, NoJoueur
                    End Select
                Else
                    Select Case i
                    Case 0: Afficher_Info_Bulles_Terrain X, Y, MessageConfirmationOui, SurfBack
                    Case 1: Afficher_Info_Bulles_Terrain X, Y, MessageConfirmationNon, SurfBack
                    End Select
                End If
                NoBouton = 999
                Interieur_Boutons = True
                Exit For
            End If
        Next i
    Else
    'Boutons radiaux.
    If Not Interieur_Boutons And ActiverBoutonsRadiaux Then
        Select Case BoutonsRadiauxCategorie
        Case 0:
            For i = 0 To NombreBoutonsRadiauxInfo - 1
                If Interieur_Bouton(X, Y, _
                                    BoutonsRadiauxInfos(i).PositionX, _
                                    BoutonsRadiauxInfos(i).PositionY, _
                                    AffBoutonsInfos.Largeur, _
                                    AffBoutonsInfos.Hauteur) Then
                    If CliqueGauche Then
                        Actionner_Boutons_Infos i, NoJoueur
                        Activer_EffetVisuel_Clique X, Y
                    Else
                        BoutonsInfos(i).Afficher_Info_Bulle X, Y, AffMessage, SurfBack
                    End If
                    NoBouton = 999
                    Interieur_Boutons = True
                    'i = UBound(BoutonsInfos())
                    BoutonsRadiauxInfos(i).MouseOver = True
                Else
                    BoutonsRadiauxInfos(i).MouseOver = False
                End If
            Next i
        Case 1:
            'Boutons radiaux construire.
            For i = 0 To UBound(BoutonsConstruire()) - 1
                If Interieur_Bouton(X, Y, _
                                    BoutonsRadiauxConstruire(i).PositionX, _
                                    BoutonsRadiauxConstruire(i).PositionY, _
                                    AffBoutonsConstruire(i).Largeur, _
                                    AffBoutonsConstruire(i).Hauteur) Then
                    Actualiser_BoutonsConstruire Perso, SaMaison, Fief, Parametres
                    BoutonsConstruire(i).Enabled = (Etat(i) = 1)
                    If BoutonsConstruire(i).Enabled Then
                        If i = 0 Then
                            NoBouton = 999
                        ElseIf CliqueGauche Then
                            NoBouton = BoutonsConstruire_IndiceBatiment(i) + 1
                            NoBatiment = BoutonsConstruire_IndiceBatiment(i) + 1
                        End If
                    Else
                        NoBouton = 999
                    End If
                    If CliqueGauche And BoutonsConstruire(i).Enabled Then
                        Activer_EffetVisuel_Clique X, Y
                        If i = 0 Then
                            'SaMaison.Demolir Perso, Parametres
                            Cmd_Perso_DetuireMaison Perso.Numero
                            Actualiser_BoutonsConstruire Perso, SaMaison, Fief, Parametres
                        End If
                    Else
                        If BoutonConstruireCompteurindice1 = i Then
                            Definir_InfoBulle_Batiment i, Perso, (Abs(Timer - BoutonConstruireCompteurTemps) - BoutonConstruireCompteurDuree) / BoutonConstruireCompteurIntervales
                        Else
                            BoutonConstruireCompteurTemps = Timer
                            BoutonConstruireCompteurindice1 = i
                            Definir_InfoBulle_Batiment i, Perso
                        End If
                        BoutonsConstruire(i).Afficher_Info_Bulle X, Y, AffMessage, SurfBack
                    End If
                    Interieur_Boutons = True
                    For j = i To UBound(BoutonsConstruire()) - 1
                        BoutonsRadiauxConstruire(j).MouseOver = False
                    Next j
                    BoutonsRadiauxConstruire(i).MouseOver = True
                    Exit For
                Else
                    BoutonsRadiauxConstruire(i).MouseOver = False
                End If
            Next i
        End Select
        If CliqueGauche Or CliqueDroit Then
            If CliqueDroit Then
                Activer_BoutonsRadiaux CSng(X), CSng(Y), 0
            Else
                ActiverBoutonsRadiaux = False
            End If
        End If
        If NoBouton = 0 Then NoBouton = 999
        Interieur_Boutons = True
    End If
    
    'Boutons de menu.
    If Not Interieur_Boutons And AfficherBoutonsMenu And BoutonsInfos(7).Enabled Then
        If Jeu.Campagne Then
            For i = 1 To NombreBoutonsMenu - 1
                If Interieur_Bouton(X, Y, _
                                    BoutonsMenu(i).PositionX - BoutonMenuIntervalX, _
                                    BoutonsMenu(i).PositionY - BoutonMenuIntervalY, _
                                    AffBoutonsMenu.Largeur, _
                                    AffBoutonsMenu.Hauteur) Then
                    If CliqueGauche Or CliqueDroit Then
                        Activer_EffetVisuel_Clique X, Y
                        Select Case i
                        Case 1: FenetreConfirmation.Afficher_Fenetre 0, MessageConfirmationQuitter, DecalageX, DecalageY, Largeur, Hauteur
                                'Quitter = True
                        End Select
                    Else
                        BoutonsMenu(i).Afficher_Info_Bulle X, Y, AffMessage, SurfBack
                        BoutonsMenu(i).MouseOver = True
                    End If
                    NoBouton = 999
                    Interieur_Boutons = True
                    'i = UBound(BoutonsInfos())
                Else
                    BoutonsMenu(i).MouseOver = False
                End If
            Next i
        Else
            For i = 0 To NombreBoutonsMenu - 1
                If Interieur_Bouton(X, Y, _
                                    BoutonsMenu(i).PositionX, _
                                    BoutonsMenu(i).PositionY, _
                                    AffBoutonsMenu.Largeur, _
                                    AffBoutonsMenu.Hauteur) Then
                    If CliqueGauche Or CliqueDroit Then
                        Activer_EffetVisuel_Clique X, Y
                        Select Case i
                        Case 0: Enregistrer_Sauvegarde_Monde
                        Case 1: FenetreConfirmation.Afficher_Fenetre 0, MessageConfirmationQuitter, DecalageX, DecalageY, Largeur, Hauteur
                                'Quitter = True
                        End Select
                    Else
                        BoutonsMenu(i).Afficher_Info_Bulle X, Y, AffMessage, SurfBack
                        BoutonsMenu(i).MouseOver = True
                    End If
                    NoBouton = 999
                    Interieur_Boutons = True
                    'i = UBound(BoutonsInfos())
                Else
                    BoutonsMenu(i).MouseOver = False
                End If
            Next i
        End If
    End If
    
    'Boutons de construction.
    CurseurSurBoutonConstruire = False
    If AfficherBoutonsConstruire Then
        For i = 0 To UBound(BoutonsConstruire())
            'If BoutonsConstruire(i).Visible And
            If Interieur_Bouton(X, Y, _
                                BoutonsConstruire(i).PositionX, _
                                BoutonsConstruire(i).PositionY, _
                                AffBoutonsConstruire(i).Largeur, _
                                AffBoutonsConstruire(i).Hauteur) Then
                CurseurSurBoutonConstruire = True
                If Not Interieur_Boutons Then
                    Actualiser_BoutonsConstruire Perso, SaMaison, Fief, Parametres
                    BoutonsConstruire(i).Enabled = (Etat(i) = 1)
                    If BoutonsConstruire(i).Enabled Then
                        If i = 0 Then
                            NoBouton = 999
                        ElseIf CliqueGauche Then
                            NoBouton = BoutonsConstruire_IndiceBatiment(i) + 1
                            NoBatiment = BoutonsConstruire_IndiceBatiment(i) + 1
                        End If
                    Else
                        NoBouton = 999
                    End If
                    If CliqueGauche And BoutonsConstruire(i).Enabled Then
                        Activer_EffetVisuel_Clique X, Y
                        If i = 0 Then
                            'SaMaison.Demolir Perso, Parametres
                            Cmd_Perso_DetuireMaison Perso.Numero
                            'Actualiser_BoutonsConstruire Perso, SaMaison, Fief, Parametres
                        End If
                    ElseIf CliqueDroit Then
                        Activer_EffetVisuel_Clique X, Y
                        Trouver_Maison Perso, BoutonsConstruire_IndiceBatiment(i), Ctrl
                    Else
                        If BoutonConstruireCompteurindice1 = i Then
                            Definir_InfoBulle_Batiment i, Perso, (Abs(Timer - BoutonConstruireCompteurTemps) - BoutonConstruireCompteurDuree) / BoutonConstruireCompteurIntervales
                        Else
                            BoutonConstruireCompteurTemps = Timer
                            BoutonConstruireCompteurindice1 = i
                            Definir_InfoBulle_Batiment i, Perso
                        End If
                        'Messages(0).Ajouter_Message BoutonConstruireCompteurindice1 & " / " & i & " / " & Abs(Timer - BoutonConstruireCompteurTemps)
                        BoutonsConstruire(i).Afficher_Info_Bulle X, Y, AffMessage, SurfBack
                    End If
                    Interieur_Boutons = True
                End If
                Exit For
            End If
        Next i
        'If Not CurseurSurBoutonConstruire And Not ActiverBoutonsRadiaux Then BoutonConstruireCompteurTemps = Timer
    End If
    
    'Boutons dans la fenêtre carte.
    If Not Interieur_Boutons Then
        If BoutonsInfos(6).Enabled Then
            'La carte
            If Interieur_Bouton(X, Y, _
                                Fenetres(6).Definir_Carte_PositionX, _
                                Fenetres(6).Definir_Carte_PositionY, _
                                Fenetres(6).Definir_Carte_Largeur, _
                                Fenetres(6).Definir_Carte_Hauteur) Then
                'If CliqueEnfonceMilieu Then
                If CliqueEnfonceDroit Then
                    If FrmMoteur2D.NombreJoueursMO = 1 Or FrmMoteur2D.EcranDivise Then
                        'Déplace l'écran vers le lieu désigné.
                        FrmMoteur2D.Definir_DecalageX(NoJoueur) = (X - Fenetres(6).Definir_Carte_PositionX) / Fenetres(6).Definir_Carte_Largeur * Monde.Largeur * Monde.LargeurCase
                        FrmMoteur2D.Definir_DecalageY(NoJoueur) = (Y - Fenetres(6).Definir_Carte_PositionY) / Fenetres(6).Definir_Carte_Hauteur * Monde.Hauteur * Monde.HauteurCase
                        'FrmMoteur2D.Definir_ScrollingAuto(NoJoueur) = Jeu.Definir_Vitesse * 3
                        FrmMoteur2D.Definir_ScrollingAuto(NoJoueur) = False
                    End If
                ElseIf ObstacleEnnemi(Perso, _
                                 (X - Fenetres(6).Definir_Carte_PositionX) / Fenetres(6).Definir_Carte_Largeur * Monde.Largeur * Monde.LargeurCase, _
                                 (Y - Fenetres(6).Definir_Carte_PositionY) / Fenetres(6).Definir_Carte_Hauteur * Monde.Hauteur * Monde.HauteurCase, _
                                 Perso.Largeur, _
                                 Perso.Hauteur) Then
                    If CliqueGauche Then ' Or CliqueDroit Then
                        If Perso.Vivant Then
                            If ArmeeSelectionnee Then
                                If CliqueGauche Then
                                    Ordonner_Attaquer_Perso Perso, _
                                                            SoldatsSelectionnes, _
                                                            IndiceObstacle
                                End If
                            Else
                                'If CliqueGauche Then
                                'Perso.Attaquer_Perso IndiceObstacle
                                Cmd_Perso_AttaquerPerso Perso.Numero, IndiceObstacle
                                'Else
                                '    Ordonner_Chargez Perso.Numero, 0, IndiceObstacle, Ctrl, Shift, Alt
                                'End If
                            End If
                        Else
                            'Déplacement quand le joueur est un fantôme.
                            If CliqueGauche Then
                                'Perso.Aller_A (X - Fenetres(6).Definir_Carte_PositionX) / Fenetres(6).Definir_Carte_Largeur * Monde.Largeur * Monde.LargeurCase, _
                                              (Y - Fenetres(6).Definir_Carte_PositionY) / Fenetres(6).Definir_Carte_Hauteur * Monde.Hauteur * Monde.HauteurCase
                                Cmd_Perso_Deplacer Perso.Numero, (X - Fenetres(6).Definir_Carte_PositionX) / Fenetres(6).Definir_Carte_Largeur * Monde.Largeur * Monde.LargeurCase, _
                                                                 (Y - Fenetres(6).Definir_Carte_PositionY) / Fenetres(6).Definir_Carte_Hauteur * Monde.Hauteur * Monde.HauteurCase
                            End If
                        End If
                    Else
                        Afficher_Info_Bulles_Terrain X, Y, "&" & Persos(IndiceObstacle).Nom & " ", SurfBack
                    End If
                ElseIf CliqueGauche Then
                    If ArmeeSelectionnee Then
                        If CliqueGauche Then
                            Ordonner_Deplacement Perso, _
                                                 SoldatsSelectionnes, _
                                                 (X - Fenetres(6).Definir_Carte_PositionX) / Fenetres(6).Definir_Carte_Largeur * Monde.Largeur * Monde.LargeurCase, _
                                                 (Y - Fenetres(6).Definir_Carte_PositionY) / Fenetres(6).Definir_Carte_Hauteur * Monde.Hauteur * Monde.HauteurCase
                        End If
                    Else
                        'Perso.Aller_A (X - Fenetres(6).Definir_Carte_PositionX) / Fenetres(6).Definir_Carte_Largeur * Monde.Largeur * Monde.LargeurCase, _
                                      (Y - Fenetres(6).Definir_Carte_PositionY) / Fenetres(6).Definir_Carte_Hauteur * Monde.Hauteur * Monde.HauteurCase
                        Cmd_Perso_Deplacer Perso.Numero, (X - Fenetres(6).Definir_Carte_PositionX) / Fenetres(6).Definir_Carte_Largeur * Monde.Largeur * Monde.LargeurCase, _
                                                         (Y - Fenetres(6).Definir_Carte_PositionY) / Fenetres(6).Definir_Carte_Hauteur * Monde.Hauteur * Monde.HauteurCase
                    End If
                End If
                NoBouton = 900
                Interieur_Boutons = True
            End If
        End If
    End If
    
    'Boutons dans la fenêtre personnage.
    If Not Interieur_Boutons Then
        If BoutonsInfos(2).Enabled Then
        For i = 0 To Fenetres(2).BoutonsCaracNombre - 1
            If Interieur_Bouton(X, Y, _
                                Fenetres(2).BoutonsCaracPositionX(i), _
                                Fenetres(2).BoutonsCaracPositionY(i), _
                                Fenetres(2).BoutonsCaracLargeur(i), _
                                Fenetres(2).BoutonsCaracHauteur(i)) Then
                Afficher_Info_Bulles_Terrain X, Y, Fenetres(2).BoutonsCaracInfoBulle(i, Perso, Parametres), SurfBack
                NoBouton = 999
                Interieur_Boutons = True
                Exit For
            End If
        Next i
        End If
    End If
    
    'Boutons dans la fenêtre des ressources.
    If Not Interieur_Boutons Then
        If BoutonsInfos(1).Enabled Then
            'Boutons manger.
            For i = 0 To Fenetres(1).NombreBoutonsManger - 1
                If Fenetres(1).BoutonMangerVisible(i) And _
                   Interieur_Bouton(X, Y, _
                                    Fenetres(1).BoutonMangerPositionX(i), _
                                    Fenetres(1).BoutonMangerPositionY(i), _
                                    Fenetres(1).BoutonMangerLargeur(i), _
                                    Fenetres(1).BoutonMangerHauteur(i)) Then
                        If CliqueGauche And Perso.Vivant Then
                            Activer_EffetVisuel_Clique X, Y
                            'Perso.Manger_Ressources Fenetres(1).BoutonsMangerIndiceRessources(i), Parametres
                            Cmd_Perso_RessourceManger Perso.Numero, Fenetres(1).BoutonsMangerIndiceRessources(i)
                        Else
                            Fenetres(1).BoutonsManger_Afficher_InfoBulle i, X, Y, Perso, Parametres, AffMessage, SurfBack
                        End If
                        NoBouton = 999
                        Interieur_Boutons = True
                        Exit For
                End If
            Next i
            
            'Boutons de défilement.
            If Not Interieur_Boutons Then
                Interieur_Boutons = Interieur_BoutonsDefilement(X, Y, 1, NoBouton, CliqueGauche, CliqueEnfonceGauche, SurfBack)
            End If
            
            'Boutons Transfert de dépot de stock.
            If Not Interieur_Boutons And Perso.DansSaMaison And SaMaison.Magasin Then
                For i = 0 To Fenetres(1).NombreColonnesBoutonsTransfert - 1
                    If Interieur_Bouton(X, Y, _
                                        Fenetres(1).BoutonTransfertStockPersoPositionX(i), _
                                        Fenetres(1).BoutonTransfertStockPersoPositionY(i), _
                                        Fenetres(1).BoutonTransfertLargeur(), _
                                        Fenetres(1).BoutonTransfertHauteur()) Then
                            If CliqueGauche And Perso.Vivant Then
                                Activer_EffetVisuel_Clique X, Y
                                Select Case i
                                    Case 0: 'Maison.Transfert_Stock_Deposer 1, Perso, Parametres
                                            Cmd_Perso_RessourcesDeposerMaison Perso.Numero, 1
                                    Case 1: 'Maison.Transfert_Stock_Deposer 5, Perso, Parametres
                                            Cmd_Perso_RessourcesDeposerMaison Perso.Numero, 5
                                    Case 2: 'Maison.Transfert_Stock_Deposer 0, Perso, Parametres, True
                                            Cmd_Perso_RessourcesDeposerMaison Perso.Numero, 0, True
                                End Select
                            Else
                                Fenetres(1).BoutonTransfertStockPerso_Afficher_InfoBulle i, X, Y, AffMessage, SurfBack
                            End If
                            NoBouton = 999
                            Interieur_Boutons = True
                            Exit For
                    End If
                Next i
            End If
            
            'Boutons de vente.
            If Not Interieur_Boutons Then 'And (Perso.DansUnChateau Or Perso.DansSonChateau) Then
                For i = Fenetres(1).DefilementDecalageX To Fenetres(1).DefilementDecalageX + Fenetres(1).Nombre_Lignes_Ressources - 1
                    If i < Parametres.NombreRessources Then
                        For j = 0 To Fenetres(1).NombreColonnesBoutonsCommerce - 1
                            If Interieur_Bouton(X, Y, _
                                                Fenetres(1).BoutonCommercePositionX(i * (Fenetres(1).NombreColonnesBoutonsCommerce) + j), _
                                                Fenetres(1).BoutonCommercePositionY(i * (Fenetres(1).NombreColonnesBoutonsCommerce) + j), _
                                                Fenetres(1).BoutonCommerceLargeur(i * (Fenetres(1).NombreColonnesBoutonsCommerce) + j), _
                                                Fenetres(1).BoutonCommerceHauteur(i * (Fenetres(1).NombreColonnesBoutonsCommerce) + j)) Then
                                If (CliqueGauche Or CliqueDroit) And Perso.Vivant Then
                                    Activer_EffetVisuel_Clique X, Y
                                    If Perso.DansSonChateau Or Perso.DansUnChateau Or (Perso.DansUneMaison And Parametres.Batiment_Marche(Maison.TypeBatiment)) Then
                                        Select Case j
                                            Case 0: 'Fief.Vendre i, 1, Perso, Parametres
                                                    Cmd_Perso_RessourcesVendreChateau Perso.Numero, i, 1
                                            Case 1: 'Fief.Vendre i, 5, Perso, Parametres
                                                    Cmd_Perso_RessourcesVendreChateau Perso.Numero, i, 5
                                            Case 2: 'Fief.Vendre i, 0, Perso, Parametres, True
                                                    Cmd_Perso_RessourcesVendreChateau Perso.Numero, i, Perso.Definir_ressources(i)
                                        End Select
                                    Else
                                        Select Case j
                                            Case 0:
                                                'Perso.Ressources_Prendre i, 1, False, CliqueDroit
                                                If CliqueDroit Then
                                                    Cmd_Perso_RessourceLaisser Perso.Numero, i, 1
                                                Else
                                                    Cmd_Perso_RessourcePrendre Perso.Numero, i, 1
                                                End If
                                            Case 1:
                                                'Perso.Ressources_Prendre i, 5, False, CliqueDroit
                                                If CliqueDroit Then
                                                    Cmd_Perso_RessourceLaisser Perso.Numero, i, 5
                                                Else
                                                    Cmd_Perso_RessourcePrendre Perso.Numero, i, 5
                                                End If
                                            Case 2:
                                                'Perso.Ressources_Prendre i, 0, True, CliqueDroit
                                                If CliqueDroit Then
                                                    Cmd_Perso_RessourceLaisser Perso.Numero, i, Perso.Definir_ressources(i)
                                                Else
                                                    Cmd_Perso_RessourcePrendre Perso.Numero, i, Perso.Definir_ressources(i)
                                                End If
                                        End Select
                                    End If
                                Else
                                    Fenetres(1).BoutonCommerce_Afficher_InfoBulle i * (Fenetres(1).NombreColonnesBoutonsCommerce) + j, Perso, Maison, Fief, X, Y, AffMessage, SurfBack, False
                                End If
                                NoBouton = 999
                                Interieur_Boutons = True
                                Exit For
                            End If
                        Next j
                    Else
                        'i = Fenetres(1).DefilementDecalageX + Fenetres(1).Nombre_Lignes_Ressources
                        Exit For
                    End If
                    'If Interieur_Boutons Then
                    '    i = Fenetres(1).DefilementDecalageX + Fenetres(1).Nombre_Lignes_Ressources
                    'End If
                Next i
            End If
            
            'Boutons Transfert de dépot d'argent.
            If Not Interieur_Boutons And Perso.DansSaMaison Then
                For i = 0 To Fenetres(1).NombreColonnesBoutonsTransfert - 1
                    If Interieur_Bouton(X, Y, _
                                        Fenetres(1).BoutonTransfertArgentPersoPositionX(i), _
                                        Fenetres(1).BoutonTransfertArgentPersoPositionY(i), _
                                        Fenetres(1).BoutonTransfertLargeur(), _
                                        Fenetres(1).BoutonTransfertHauteur()) Then
                            If CliqueGauche And Perso.Vivant Then
                                Activer_EffetVisuel_Clique X, Y
                                Select Case i
                                    'Case 0: Maison.Transfert_Argent_Deposer 1, Perso
                                    'Case 1: Maison.Transfert_Argent_Deposer 5, Perso
                                    'Case 2: Maison.Transfert_Argent_Deposer 0, Perso, True
                                    Case 0: Cmd_Perso_ArgentMaisonLaisser Perso.Numero, 1
                                    Case 1: Cmd_Perso_ArgentMaisonLaisser Perso.Numero, 5
                                    Case 2: Cmd_Perso_ArgentMaisonLaisser Perso.Numero, Perso.Definir_Argent
                                End Select
                            Else
                                Fenetres(1).BoutonTransfertArgentPerso_Afficher_InfoBulle i, X, Y, AffMessage, SurfBack
                            End If
                            NoBouton = 999
                            Interieur_Boutons = True
                            Exit For
                    End If
                Next i
            End If
            
            If Not Interieur_Boutons Then
                'Boutons de transaction globale.
                If Not Interieur_Boutons Then
                    For i = (Parametres.NombreRessources - 1) * Fenetres(1).NombreColonnesBoutonsCommerce + Fenetres(1).NombreColonnesBoutonsCommerce To _
                            (Parametres.NombreRessources - 1) * Fenetres(1).NombreColonnesBoutonsCommerce + Fenetres(1).NombreColonnesBoutonsCommerce * 2 - 1
                        'Debug.Print i & "; " & x & "; " & y & "; " & Fenetres(1).BoutonCommercePositionX(i) & "; " & Fenetres(1).BoutonCommercePositionY(i) & "; " & Interieur_Bouton(x, y, _
                                            Fenetres(1).BoutonCommercePositionX(i), _
                                            Fenetres(1).BoutonCommercePositionY(i), _
                                            Fenetres(1).BoutonCommerceLargeur(i), _
                                            Fenetres(1).BoutonCommerceHauteur(i))
                        If Interieur_Bouton(X, Y, _
                                            Fenetres(1).BoutonCommercePositionX(i), _
                                            Fenetres(1).BoutonCommercePositionY(i), _
                                            Fenetres(1).BoutonCommerceLargeur(i), _
                                            Fenetres(1).BoutonCommerceHauteur(i)) Then
                            If (CliqueGauche Or CliqueDroit) And Perso.Vivant Then
                                Activer_EffetVisuel_Clique X, Y
                                For j = 0 To Parametres.NombreRessources - 1
                                    Select Case i - ((Parametres.NombreRessources - 1) * Fenetres(1).NombreColonnesBoutonsCommerce + Fenetres(1).NombreColonnesBoutonsCommerce)
                                    Case 2:
                                        If Perso.DansSonChateau Or Perso.DansUnChateau Or (Perso.DansUneMaison And Parametres.Batiment_Marche(Maison.TypeBatiment)) Then
                                            'Fief.Vendre j, 0, Perso, Parametres, True
                                            Cmd_Perso_RessourcesVendreChateau Perso.Numero, j, Perso.Definir_ressources(j)
                                        Else
                                            'Perso.Ressources_Prendre j, 0, True, CliqueDroit
                                            Cmd_Perso_PrendreRessource Perso.Numero, j, Perso.Definir_ressources(j), CliqueDroit
                                        End If
                                    Case 1:
                                        If Not Parametres.Ressources_Mangeable(j) And _
                                               Not Parametres.Ressources_NecessairePourConstruire(j) And _
                                               Not Parametres.Ressources_NecessairePourAttaquer(j) Then
                                            If Perso.DansSonChateau Or Perso.DansUnChateau Or (Perso.DansUneMaison And Parametres.Batiment_Marche(Maison.TypeBatiment)) Then
                                                'Fief.Vendre j, 0, Perso, Parametres, True
                                                Cmd_Perso_RessourcesVendreChateau Perso.Numero, j, Perso.Definir_ressources(j)
                                            Else
                                                'Perso.Ressources_Prendre j, 0, True, CliqueDroit
                                                Cmd_Perso_PrendreRessource Perso.Numero, j, Perso.Definir_ressources(j), CliqueDroit
                                            End If
                                        End If
                                    Case 0:
                                        If Perso.Definir_ressources(j) >= 21 Then
                                            If Perso.DansSonChateau Or Perso.DansUnChateau Or (Perso.DansUneMaison And Parametres.Batiment_Marche(Maison.TypeBatiment)) Then
                                                'Fief.Vendre j, CLng(Perso.Definir_Ressources(j) - 20), Perso, Parametres
                                                Cmd_Perso_RessourcesVendreChateau Perso.Numero, j, Int(Perso.Definir_ressources(j) - 20)
                                            Else
                                                If CliqueDroit Then
                                                    If Perso.Definir_RessourcesSelectionnes(j) > 20 Then
                                                        'Perso.Ressources_Prendre j, Perso.Definir_RessourcesSelectionnes(j) - 20, False, True
                                                        Cmd_Perso_PrendreRessource Perso.Numero, j, Perso.Definir_RessourcesSelectionnes(j) - 20, True
                                                    ElseIf Perso.Definir_RessourcesSelectionnes(j) < 20 Then
                                                        'Perso.Ressources_Prendre j, 20
                                                        Cmd_Perso_PrendreRessource Perso.Numero, j, 20
                                                    End If
                                                Else
                                                    If Perso.Definir_RessourcesSelectionnes(j) > Perso.Definir_ressources(j) - 20 Then
                                                        Perso.Definir_RessourcesSelectionnes(j) = 0
                                                        'Perso.Ressources_Prendre j, CLng(Perso.Definir_Ressources(j) - 20)
                                                        Cmd_Perso_PrendreRessource Perso.Numero, j, Int(Perso.Definir_ressources(j) - 20)
                                                    Else
                                                        'Perso.Ressources_Prendre j, CLng(Perso.Definir_Ressources(j) - 20 - Perso.Definir_RessourcesSelectionnes(j))
                                                        Cmd_Perso_PrendreRessource Perso.Numero, j, Int(Perso.Definir_ressources(j) - 20 - Perso.Definir_RessourcesSelectionnes(j))
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End Select
                                Next j
                            Else
                                Fenetres(1).BoutonCommerce_Afficher_InfoBulle i, Perso, Maison, Fief, X, Y, AffMessage, SurfBack, False
                            End If
                            NoBouton = 999
                            Interieur_Boutons = True
                            Exit For
                        End If
                    Next i
                End If
            End If
            
            'Boutons pour prendre des ressources.
'            If Not Interieur_Boutons Then
'                For i = Fenetres(1).DefilementDecalageX To Fenetres(1).DefilementDecalageX + Fenetres(1).Nombre_Lignes_Ressources - 1
'                    If Interieur_Bouton(X, Y, _
'                                        Fenetres(1).BoutonsRessourcesPrendrePositionX(i), _
'                                        Fenetres(1).BoutonsRessourcesPrendrePositionY(i), _
'                                        Fenetres(1).BoutonsRessourcesPrendreLargeur(i), _
'                                        Fenetres(1).BoutonsRessourcesPrendreHauteur(i)) Then
'                        If (CliqueGauche Or CliqueDroit) And Perso.Vivant Then
'                            Activer_EffetVisuel_Clique X, Y
'                            If CliqueGauche Then
'                                Perso.Ressources_Prendre i, 1
'                            Else
'                                Perso.Ressources_Prendre i, 0, True
'                            End If
'                        Else
'                            Fenetres(1).BoutonsRessourcesPrendre_Afficher_InfoBulle i, X, Y, AffMessage, SurfBack
'                        End If
'                        NoBouton = 999
'                        Interieur_Boutons = True
'                        Exit For
'                    End If
'                Next i
'            End If
        End If
    End If

    'Boutons dans la fenêtre de batiment.
    If Fenetres(0).Visible And Not Interieur_Boutons Then
        
        'Boutons d'achat.
        If Perso.DansSonChateau Or Perso.DansUnChateau Or Parametres.Batiment_Marche(Maison.TypeBatiment) Then
            
            If Perso.DansSonChateau Or Perso.DansUnChateau Then
                'Bouton "Empereur".
                If Interieur_Bouton(X, Y, Fenetres(0).BoutonEmpereur_PositionX, _
                                          Fenetres(0).BoutonEmpereur_PositionY, _
                                          Fenetres(0).BoutonEmpereur_Largeur, _
                                          Fenetres(0).BoutonEmpereur_Hauteur) Then
                    If CliqueGauche Or CliqueDroit Then
                        If Perso.PeutDevenir_Empereur(Fief, Parametres) Then
                            Cmd_Perso_DevenirEmpereur Perso.Numero
                            'Fief.Empereur_Changer Perso.Numero, True
                        End If
                    Else
                        If Fief.Empereur >= 0 Then
                            'Indique le nom de l'empereur actuel.
                            Afficher_Info_Bulles_Terrain X, Y, Parametres.Epoque_EmpereurNom(Fief.Epoque, Persos(Fief.Empereur).Feminin) & " " & Persos(Fief.Empereur).Nom, SurfBack
                        Else
                        Afficher_Info_Bulles_Terrain X, Y, Fenetres(0).BoutonEmpereur_InfoBulle(0) & _
                                                           LCase(Parametres.Epoque_EmpereurNom(Fief.Epoque, Perso.Feminin)) & _
                                                           Fenetres(0).BoutonEmpereur_InfoBulle(1) & _
                                                           AffMessage.Caractere_RetourLigne & _
                                                           IIf(Perso.PeutDevenir_Empereur(Fief, Parametres), "", "&") & _
                                                           Fief.Empereur_Prix(Parametres) & " " & Parametres.EtiquetteArgent, _
                                                           SurfBack
                        End If
                    End If
                    NoBouton = 999
                    Interieur_Boutons = True
                End If
                'Boutons de restauration du chateau.
                If Not Interieur_Boutons Then
                    For i = 0 To Fenetres(0).Nombre_BoutonsRestauration - 1
                        If Interieur_Bouton(X, Y, _
                                            Fenetres(0).BoutonRestaurationPositionX(i), _
                                            Fenetres(0).BoutonRestaurationPositionY(i), _
                                            Fenetres(0).BoutonRestaurationLargeur(i), _
                                            Fenetres(0).BoutonRestaurationHauteur(i)) Then
                                If CliqueGauche And Perso.Vivant Then
                                    Activer_EffetVisuel_Clique X, Y
                                    Select Case i
                                        Case 0: 'Fief.Restaurer_Chateau 5, Perso
                                                Cmd_Perso_RestaurerChateau Perso.Numero, 5
                                        Case 1: 'Fief.Restaurer_Chateau 50, Perso
                                                Cmd_Perso_RestaurerChateau Perso.Numero, 50
                                        Case 2: 'Fief.Restaurer_Chateau 0, Perso, True
                                                Cmd_Perso_RestaurerChateau Perso.Numero, Perso.Definir_Argent
                                    End Select
                                Else
                                    'Affiche les infos-bulles.
                                    Fenetres(0).BoutonRestauration_Afficher_InfoBulle i, X, Y, AffMessage, SurfBack
                                End If
                                NoBouton = 999
                                Interieur_Boutons = True
                                Exit For
                        End If
                    Next i
                End If
            End If
            
            If Not Interieur_Boutons Then
                'For i = Fenetres(0).DefilementDecalageX To Parametres.NombreRessources - 1
                For i = Fenetres(0).DefilementDecalageX To Fenetres(0).DefilementDecalageX + Fenetres(0).Nombre_Lignes_Ressources - 1
                    For j = 0 To Fenetres(0).NombreColonnesBoutonsCommerce - 1
                        If Interieur_Bouton(X, Y, _
                                            Fenetres(0).BoutonCommercePositionX(i * (Fenetres(0).NombreColonnesBoutonsCommerce) + j), _
                                            Fenetres(0).BoutonCommercePositionY(i * (Fenetres(0).NombreColonnesBoutonsCommerce) + j), _
                                            Fenetres(0).BoutonCommerceLargeur(i * (Fenetres(0).NombreColonnesBoutonsCommerce) + j), _
                                            Fenetres(0).BoutonCommerceHauteur(i * (Fenetres(0).NombreColonnesBoutonsCommerce) + j)) Then
                                If CliqueGauche And Perso.Vivant Then
                                    Activer_EffetVisuel_Clique X, Y
                                    Select Case j
                                        Case 0: 'Fief.Acheter i, 1, Perso, Parametres
                                                Cmd_Perso_RessourcesAcheterChateau Perso.Numero, i, 1
                                        Case 1: 'Fief.Acheter i, 5, Perso, Parametres
                                                Cmd_Perso_RessourcesAcheterChateau Perso.Numero, i, 5
                                        Case 2: 'Fief.Acheter i, 0, Perso, Parametres, True
                                                Cmd_Perso_RessourcesAcheterChateau Perso.Numero, i, Perso.Definir_MaxRessources(i)
                                    End Select
                                Else
                                    Fenetres(0).BoutonCommerce_Afficher_InfoBulle i * (Fenetres(0).NombreColonnesBoutonsCommerce) + j, Perso, Maison, Fief, X, Y, AffMessage, SurfBack, True
                                    'Debug.Print i * (Fenetres(0).NombreColonnesBoutonsCommerce) + j
                                End If
                                NoBouton = 999
                                Interieur_Boutons = True
                                Exit For
                        End If
                    Next j
                    'Exit For
                Next i
            End If
            'Boutons de défilement.
            If Not Interieur_Boutons Then
                Interieur_Boutons = Interieur_BoutonsDefilement(X, Y, 0, NoBouton, CliqueGauche, CliqueEnfonceGauche, SurfBack)
            End If
        Else
            If Perso.DansSaMaison Then
                'Boutons Transfert de récupération de l'argent.
                If Not Interieur_Boutons Then
                    For i = 0 To Fenetres(0).NombreColonnesBoutonsTransfert - 1
                        If Interieur_Bouton(X, Y, _
                                            Fenetres(0).BoutonTransfertArgentMaisonPositionX(i), _
                                            Fenetres(0).BoutonTransfertArgentMaisonPositionY(i), _
                                            Fenetres(0).BoutonTransfertLargeur(), _
                                            Fenetres(0).BoutonTransfertHauteur()) Then
                                If CliqueGauche And Perso.Vivant Then
                                    Activer_EffetVisuel_Clique X, Y
                                    Select Case i
                                        Case 0: 'SaMaison.Transfert_Argent_Prendre 1, Perso
                                                Cmd_Perso_ArgentMaisonPrendre Perso.Numero, 1
                                        Case 1: 'SaMaison.Transfert_Argent_Prendre 5, Perso
                                                Cmd_Perso_ArgentMaisonPrendre Perso.Numero, 5
                                        Case 2: 'SaMaison.Transfert_Argent_Prendre 0, Perso, True
                                                Cmd_Perso_ArgentMaisonPrendre Perso.Numero, Maison.Argent
                                    End Select
                                Else
                                    Fenetres(0).BoutonTransfertArgentMaison_Afficher_InfoBulle i, X, Y, AffMessage, SurfBack
                                End If
                                NoBouton = 999
                                Interieur_Boutons = True
                                Exit For
                        End If
                    Next i
                End If
                
                'Boutons Transfert de récupération de stock.
                If Not Interieur_Boutons And _
                   SaMaison.Magasin And Not SaMaison.Fabrique Then
                    For i = 0 To Fenetres(0).NombreColonnesBoutonsTransfert - 1
                        If Interieur_Bouton(X, Y, _
                                            Fenetres(0).BoutonTransfertStockMaisonPositionX(i), _
                                            Fenetres(0).BoutonTransfertStockMaisonPositionY(i), _
                                            Fenetres(0).BoutonTransfertLargeur(), _
                                            Fenetres(0).BoutonTransfertHauteur()) Then
                                If CliqueGauche And Perso.Vivant Then
                                    Activer_EffetVisuel_Clique X, Y
                                    Select Case i
                                        Case 0: 'SaMaison.Transfert_Stock_Prendre 1, Perso, Parametres
                                                Cmd_Perso_RessourceMaisonPrendre Perso.Numero, 1
                                        Case 1: 'SaMaison.Transfert_Stock_Prendre 5, Perso, Parametres
                                                Cmd_Perso_RessourceMaisonPrendre Perso.Numero, 5
                                        Case 2: 'SaMaison.Transfert_Stock_Prendre 0, Perso, Parametres, True
                                                Cmd_Perso_RessourceMaisonPrendre Perso.Numero, SaMaison.Definir_Stock
                                    End Select
                                Else
                                    Fenetres(0).BoutonTransfertStockMaison_Afficher_InfoBulle i, X, Y, AffMessage, SurfBack
                                End If
                                NoBouton = 999
                                Interieur_Boutons = True
                                Exit For
                        End If
                    Next i
                End If
                
                'Boutons de fabrication.
                If Not Interieur_Boutons And _
                   SaMaison.Fabrique Then
                    'For i = Fenetres(0).DefilementDecalageX To Fenetres(0).Nombre_Boutons_Fabriquer - 1
                    For i = Fenetres(0).DefilementObjetDecalageX To Fenetres(0).DefilementObjetDecalageX + Fenetres(0).Nombre_Lignes_Objets - 1
                        If Interieur_Bouton(X, Y, _
                                            Fenetres(0).BoutonsFabriquerPositionX(i), _
                                            Fenetres(0).BoutonsFabriquerPositionY(i), _
                                            Fenetres(0).BoutonsFabriquerLargeur, _
                                            Fenetres(0).BoutonsFabriquerHauteur) Then
                                If CliqueGauche And Perso.Vivant Then
                                    Activer_EffetVisuel_Clique X, Y
                                    If Fenetres(0).BoutonsFabriquerActif(i) Then
                                        'On démarre la fabrication de l'objet en question.
                                        'Perso.Creer_Objet i, Maison, Parametres
'                                        FileFabricationObjets.Add i
                                        FileFabricationObjets_Changer i, Perso, True, Ctrl, Shift
                                    End If
                                ElseIf CliqueDroit Then
                                    'Annule une fabrication en cours.
                                    If FileFabricationObjets.Count > 0 Then
'                                        For j = 1 To FileFabricationObjets.Count
'                                            If FileFabricationObjets.Item(j) = i Then
'                                                FileFabricationObjets.Remove j
'                                                j = FileFabricationObjets.Count + 2
'                                                'If FileFabricationObjets.Count = 0 And i = Perso.IndiceObjet Then
'                                                '    Perso.Arreter
'                                                '    perso.EtatFabricationObjet  = 0
'                                                'End If
'                                            End If
'                                        Next j
'                                        If j <> FileFabricationObjets.Count + 3 Then
'                                            If Collection_NombreElements(i, FileFabricationObjets) = 0 Then
'                                                'Perso.Arreter
'                                                'Perso.EtatFabricationObjet = 0
'                                                Cmd_Perso_ObjetArreterFabrication Perso.Numero
'                                            End If
'                                        End If
                                        FileFabricationObjets_Changer i, Perso, False, Ctrl, Shift
                                    Else
                                        'Perso.Arreter
                                        'Perso.EtatFabricationObjet = 0
                                        Cmd_Perso_ObjetArreterFabrication Perso.Numero
                                    End If
                                Else
                                    Fenetres(0).BoutonFabriquer_Afficher_InfoBulle i, X, Y, AffMessage, SurfBack
                                End If
                                NoBouton = 999
                                Interieur_Boutons = True
                                'i = Fenetres(0).Nombre_Boutons_Fabriquer
                                'i = Fenetres(0).DefilementObjetDecalageX + Fenetres(0).Nombre_Lignes_Objets
                                Exit For
                        End If
                        If i >= Fenetres(0).Nombre_Boutons_Objets - 1 Or i = Parametres.Batiment_NombreObjetsFabriques(SaMaison.TypeBatiment) - 1 Then
                            i = Fenetres(0).DefilementObjetDecalageX + Fenetres(0).Nombre_Lignes_Objets
                        End If
                    Next i
                    'Boutons de défilement.
                    If Not Interieur_Boutons Then
                        Interieur_Boutons = Interieur_BoutonsDefilement(X, Y, 0, NoBouton, CliqueGauche, CliqueEnfonceGauche, SurfBack, True)
                    End If
                End If
            
            'Le perso est dans un magasin.
            ElseIf Perso.DansUneMaison And Not Interieur_Boutons And Maison.Magasin Then
                'Boutons d'achat de produit.
                For i = 0 To Fenetres(0).NombreColonnesBoutonsTransfert - 1
                    If Interieur_Bouton(X, Y, _
                                        Fenetres(0).BoutonTransfertStockMaisonPositionX(i), _
                                        Fenetres(0).BoutonTransfertStockMaisonPositionY(i), _
                                        Fenetres(0).BoutonTransfertLargeur(), _
                                        Fenetres(0).BoutonTransfertHauteur()) Then
                        If CliqueGauche And Perso.Vivant Then
                            Activer_EffetVisuel_Clique X, Y
                            Select Case i
                                Case 0: 'Maison.Acheter 1, Perso, Parametres
                                        Cmd_Perso_RessourcesAcheterMaison Perso.Numero, Maison.Numero, 1
                                Case 1: 'Maison.Acheter 5, Perso, Parametres
                                        Cmd_Perso_RessourcesAcheterMaison Perso.Numero, Maison.Numero, 5
                                Case 2: 'Maison.Acheter 0, Perso, Parametres, True
                                        Cmd_Perso_RessourcesAcheterMaison Perso.Numero, Maison.Numero, Perso.Definir_MaxRessources(Parametres.Batiment_RessourceFabrique(Maison.TypeBatiment))
                            End Select
                        Else
                            Fenetres(0).BoutonCommerce_Afficher_InfoBulle Parametres.Batiment_RessourceFabrique(Maison.TypeBatiment) * Fenetres(0).NombreColonnesBoutonsCommerce + i, Perso, Maison, Fief, X, Y, AffMessage, SurfBack, True
                            'Debug.Print Parametres.Batiment_RessourceFabrique(Maison.TypeBatiment) & "*" & Fenetres(0).NombreColonnesBoutonsCommerce & "+" & i
                        End If
                        NoBouton = 999
                        Interieur_Boutons = True
                        Exit For
                    End If
                Next i
            End If
        End If
        
        If Not Interieur_Boutons And _
          (Perso.DansUneMaison Or Perso.DansSaMaison) And _
           Maison.Service Then
            'Bouton de service.
            If Interieur_Bouton(X, Y, _
                                Fenetres(0).BoutonServicePositionX, _
                                Fenetres(0).BoutonServicePositionY, _
                                Fenetres(0).BoutonServiceLargeur, _
                                Fenetres(0).BoutonServiceHauteur) Then
                If CliqueGauche And Maison.Service_Utilisable(Perso, Parametres) Then
                    Activer_EffetVisuel_Clique X, Y
                    'Maison.Service_Utiliser Perso, Parametres
                    Cmd_Perso_UtiliserService Perso.Numero, Maison.Numero
                ElseIf CliqueDroit Then
                    Activer_EffetVisuel_Clique X, Y
                    AutoBatimentService = Not AutoBatimentService
                Else
                    Fenetres(0).BoutonService_Afficher_InfoBulle Parametres.Batiment_RessourceFabrique(Maison.TypeBatiment) * Fenetres(0).NombreColonnesBoutonsCommerce + i, X, Y, AffMessage, SurfBack
                End If
                NoBouton = 999
                Interieur_Boutons = True
            End If
        End If
        
        'Boutons d'objets dans une fabrique du joueur ou d'un autre magasin.
        If Not Interieur_Boutons Then
            If (Perso.DansSaMaison Or Perso.DansUneMaison) Then
                If Parametres.Batiment_Fabrique(Maison.TypeBatiment) Then
                    'For i = 0 To Fenetres(0).Nombre_Boutons_Objets - 1
                    For i = Fenetres(0).DefilementObjetDecalageX To Fenetres(0).DefilementObjetDecalageX + Fenetres(0).Nombre_Lignes_Objets - 1
                        If i < Fenetres(0).Nombre_Boutons_Objets Then
                            If Interieur_Bouton(X, Y, _
                                                Fenetres(0).BoutonsObjetsPositionX(i), _
                                                Fenetres(0).BoutonsObjetsPositionY(i), _
                                                Fenetres(0).BoutonsObjetsFabriquerLargeur(AffObjets) * 2, _
                                                Fenetres(0).BoutonsObjetsFabriquerHauteur(AffObjets)) Then
                                If CliqueGauche Then
                                    Activer_EffetVisuel_Clique X, Y
                                    If Perso.DansSaMaison Then
                                        If Fenetres(0).BoutonsFabriquerActif(i) And _
                                           Maison.Stock_Objets(i) = 0 Then
                                            'On démarre la fabrication de l'objet en question si le stock est vide.
                                            'Perso.Creer_Objet i, Maison, Parametres
                                            'FileFabricationObjets.Add i
                                            FileFabricationObjets_Changer i, Perso, True, Ctrl, Shift
                                        Else
                                            'Perso.Objet_Ramasser_Maison i, Maison, Parametres, Commentaires
                                            Cmd_Perso_ObjetMaisonPrendre Perso.Numero, i, IIf(Shift, IIf(Ctrl, 20, 10), IIf(Ctrl, 5, 1))
                                        End If
                                    ElseIf Maison.Stock_Objets(i) > 0 Then
                                        'Perso.Objet_Acheter_Maison i, Maison, Parametres, Commentaires
                                        If Parametres.Objet_PrixVente(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, i)) > 0 Then
                                            'On ne peut acheter les objets gratuits.
                                            Cmd_Perso_ObjetMaisonAcheter Perso.Numero, Maison.Numero, i, IIf(Shift, IIf(Ctrl, 20, 10), IIf(Ctrl, 5, 1))
                                        End If
                                    End If
                                ElseIf CliqueDroit And Perso.DansSaMaison Then
                                    If FileFabricationObjets.Count > 0 Then
'                                        For j = 1 To FileFabricationObjets.Count
'                                            If FileFabricationObjets.Item(j) = i Then
'                                                FileFabricationObjets.Remove j
'                                                j = FileFabricationObjets.Count + 2
'                                                'If FileFabricationObjets.Count = 0 And i = Perso.IndiceObjet Then
'                                                '    Perso.Arreter
'                                                '    perso.EtatFabricationObjet  = 0
'                                                'End If
'                                            End If
'                                        Next j
'                                        If j <> FileFabricationObjets.Count + 3 Then
'                                            If Collection_NombreElements(i, FileFabricationObjets) = 0 Then
'                                                'Perso.Arreter
'                                                'Perso.EtatFabricationObjet = 0
'                                                Cmd_Perso_ObjetArreterFabrication Perso.Numero
'                                            End If
'                                        End If
                                        FileFabricationObjets_Changer i, Perso, False, Ctrl, Shift
                                    Else
                                        'Perso.Arreter
                                        'Perso.EtatFabricationObjet = 0
                                        Cmd_Perso_ObjetArreterFabrication Perso.Numero
                                    End If
                                Else
                                    Fenetres(0).BoutonObjets_Afficher_InfoBulle i, Perso, Maison, Parametres, X, Y, AffMessage, AffObjets, SurfBack, False
                                End If
                                NoBouton = 999
                                Interieur_Boutons = True
                                'i = Fenetres(0).Nombre_Boutons_Objets
                                'i = Fenetres(0).DefilementObjetDecalageX + Fenetres(0).Nombre_Lignes_Objets
                                Exit For
                            End If
                        End If
                        If i = Parametres.Batiment_NombreObjetsFabriques(Maison.TypeBatiment) - 1 Then
                            Exit For
                        End If
                    Next i
                    'Boutons de défilement.
                    If Not Interieur_Boutons Then
                        Interieur_Boutons = Interieur_BoutonsDefilement(X, Y, 0, NoBouton, CliqueGauche, CliqueEnfonceGauche, SurfBack, True)
                    End If
                End If
            End If
        End If
    End If
    
    'Boutons dans la fenêtre compétence.
    If Not Interieur_Boutons And Fenetres(3).Visible Then
        'Boutons de tri des compétences.
        For i = 0 To Fenetres(3).BoutonsCompetences_Nombre - 1
            If Interieur_Bouton(X, Y, _
                                Fenetres(3).BoutonsCompetences_PositionX(i), _
                                Fenetres(3).BoutonsCompetences_PositionY(i), _
                                Fenetres(3).BoutonsCompetences_Largeur(i), _
                                Fenetres(3).BoutonsCompetences_Hauteur(i)) Then
                If CliqueGauche Or CliqueDroit Then
                    If i < 3 Then
                        If CompetencesTri <> i Then
                            CompetencesTri = i
                            Actualiser_Competences Perso, SaMaison, Parametres
                            Fenetres(3).DefilementDecalageX = 0
                        End If
                    Else
                        CompetencesAfficherToutes = Not CompetencesAfficherToutes
                        Actualiser_Competences Perso, SaMaison, Parametres
                        Fenetres(3).DefilementDecalageX = 0
                    End If
                Else
                    Fenetres(3).BoutonsCompetences_Afficher_InfoBulle i, X, Y, AffMessage, SurfBack
                End If
                NoBouton = 999
                Interieur_Boutons = True
            End If
        Next i
        'Boutons de défilement.
        If Not Interieur_Boutons Then
            If Fenetres(3).BoutonsDefilementVisible(0) Then
                Interieur_Boutons = Interieur_BoutonsDefilement(X, Y, 3, NoBouton, CliqueGauche, CliqueEnfonceGauche, SurfBack)
            End If
        End If
    End If
    
    'Boutons dans la fenêtre inventaire.
    If Not Interieur_Boutons And Fenetres(4).Visible Then
    
        'Boutons Transfert de dépot d'argent.
        If Not Interieur_Boutons And Perso.DansSaMaison Then
            For i = 0 To Fenetres(4).NombreColonnesBoutonsTransfert - 1
                If Interieur_Bouton(X, Y, _
                                    Fenetres(4).BoutonTransfertArgentPersoPositionX(i), _
                                    Fenetres(4).BoutonTransfertArgentPersoPositionY(i), _
                                    Fenetres(4).BoutonTransfertLargeur(), _
                                    Fenetres(4).BoutonTransfertHauteur()) Then
                        If CliqueGauche And Perso.Vivant Then
                            Activer_EffetVisuel_Clique X, Y
                            Select Case i
                                'Case 0: Maison.Transfert_Argent_Deposer 1, Perso
                                'Case 1: Maison.Transfert_Argent_Deposer 5, Perso
                                'Case 2: Maison.Transfert_Argent_Deposer 0, Perso, True
                                Case 0: Cmd_Perso_ArgentMaisonLaisser Perso.Numero, 1
                                Case 1: Cmd_Perso_ArgentMaisonLaisser Perso.Numero, 5
                                Case 2: Cmd_Perso_ArgentMaisonLaisser Perso.Numero, Perso.Definir_Argent
                            End Select
                        Else
                            Fenetres(4).BoutonTransfertArgentPerso_Afficher_InfoBulle i, X, Y, AffMessage, SurfBack
                        End If
                        NoBouton = 999
                        Interieur_Boutons = True
                        Exit For
                End If
            Next i
        End If
    
        For i = 0 To Fenetres(4).Nombre_Boutons_Objets - 1
            'If Fenetres(4).BoutonsObjetsActif(i) Then
            If Interieur_Bouton(X, Y, _
                                Fenetres(4).BoutonsObjetsPositionX(i), _
                                Fenetres(4).BoutonsObjetsPositionY(i), _
                                Fenetres(4).BoutonsObjetsLargeur(AffObjets), _
                                Fenetres(4).BoutonsObjetsHauteur(AffObjets)) Then
                If CliqueGauche Then 'On prend un objet.
                    If Perso.ObjetSelectionne = -1 Then
                        If i < Perso.Nombre_Objets_Inventaire Then
                            If Ctrl And Perso.Objet_Inventaire_Type(i) > -1 Then
                                'Ctrl sert à se séparer de l'objet plus rapidement.
                                'If Perso.ObjetSelectionne >= 0 Then
                                If Perso.Vivant And Not Parametres.Objet_Inseparable(Perso.Objet_Inventaire_Type(i)) Then
                                    If Perso.DansSaMaison Then
                                        If Fenetres(4).Bascule = 0 Then
                                            'Perso.Objet_Deposer_Maison Perso.ObjetSelectionne, Maison, Parametres, False
                                            Cmd_Perso_ObjetMaisonDeposer Perso.Numero, i, True
                                            Cmd_Perso_ObjetInventaireDetruire Perso.Numero, i
                                        Else
                                            Cmd_Perso_ObjetInventaireAjouter Perso.Numero, Maison.Definir_ObjetsInventaire(i)
                                            'Maison.Definir_ObjetsInventaire(i) = -1
                                            Cmd_Perso_ObjetMaisonDetruire Perso.Numero, i
                                            'Perso.Definir_ObjetSelectionne = False
                                            'Cmd_Perso_ObjetDeSelectionne Perso.Numero
                                        End If
                                    ElseIf Perso.DansSonChateau Or Perso.DansUnChateau Or (Perso.DansUneMaison And Maison.Marche) Then
                                        'Perso.Objet_Vendre_Chateau Perso.ObjetSelectionne, Fiefs(Perso.NumeroFief), Parametres, False
                                        Cmd_Perso_ObjetChateauVendre Perso.Numero, i, True
                                    Else
                                        'Perso.Objet_Deposer_Sol Perso.ObjetSelectionne, False
                                        Cmd_PersoObjetDeposerSol Perso.Numero, i, True
                                    End If
                                    'Perso.ObjetSelectionne = -1
                                    'Cmd_Perso_ObjetSelectionne Perso.Numero, -1
                                End If
                                'End If
                            Else
                                If Fenetres(4).Bascule = 0 Or Not Perso.DansSaMaison Or Not Perso.Vivant Then
                                    'Objet de l'inventaire du joueur.
                                    'Perso.ObjetSelectionne = Perso.Objet_Inventaire_Type(i)
                                    Cmd_Perso_ObjetSelectionne Perso.Numero, Perso.Objet_Inventaire_Type(i)
                                    'Perso.Objet_Inventaire_Detruire i
                                    Cmd_Perso_ObjetInventaireDetruire Perso.Numero, i
                                Else
                                    'Objet de l'inventaire de la maison.
                                    'Perso.ObjetSelectionne = Maison.Definir_ObjetsInventaire(i)
                                    Cmd_Perso_ObjetSelectionne Perso.Numero, Maison.Definir_ObjetsInventaire(i)
                                    'Maison.Definir_ObjetsInventaire(i) = -1
                                    Cmd_Perso_ObjetMaisonDetruire Perso.Numero, i
                                End If
                            End If
                        Else
                            'Objet equipé.
                            If Perso.Objet_Equipes_2_Mains And i - Perso.Nombre_Objets_Inventaire = 1 Then
                                'Perso.ObjetSelectionne = Perso.Objet_Equipes_Type(i - Perso.Nombre_Objets_Inventaire - 1)
                                Cmd_Perso_ObjetSelectionne Perso.Numero, Perso.Objet_Equipes_Type(i - Perso.Nombre_Objets_Inventaire - 1)
                                'Perso.Objet_Equipes_Detruire i - Perso.Nombre_Objets_Inventaire - 1
                                Cmd_Perso_ObjetEquipesDetruire Perso.Numero, i - Perso.Nombre_Objets_Inventaire - 1
                            Else
                                'Perso.ObjetSelectionne = Perso.Objet_Equipes_Type(i - Perso.Nombre_Objets_Inventaire)
                                Cmd_Perso_ObjetSelectionne Perso.Numero, Perso.Objet_Equipes_Type(i - Perso.Nombre_Objets_Inventaire)
                                'Perso.Objet_Equipes_Detruire (i - Perso.Nombre_Objets_Inventaire)
                                Cmd_Perso_ObjetEquipesDetruire Perso.Numero, i - Perso.Nombre_Objets_Inventaire
                            End If
                        End If
                    Else
                        If i < Perso.Nombre_Objets_Inventaire Then
                            If Fenetres(4).Bascule = 0 Or Not Perso.DansSaMaison Or Not Perso.Vivant Then
                                'Inventaire du joueur.
                                If Perso.Objet_Inventaire_Type(i) > -1 Then
                                    'Si l'objet est superposable, on le superpose.
                                    If Perso.Objet_Inventaire_Type(i) = Perso.ObjetSelectionne And Perso.Objet_Inventaire_Nombre(i) < Parametres.Objet_Superposable(Perso.Objet_Inventaire_Type(i)) Then
                                        'S'il est superposable et inférieur au nombre maximum d'objets superposables.
                                        'Perso.Objet_Inventaire_Nombre(i) = Perso.Objet_Inventaire_Nombre(i) + 1
                                        Cmd_Perso_ObjetInventaireNombre Perso.Numero, i, Perso.Objet_Inventaire_Nombre(i) + 1
                                        'Perso.ObjetSelectionne = -1
                                        Cmd_Perso_ObjetSelectionne Perso.Numero, -1
                                    ElseIf Perso.Objet_Inventaire_Nombre(i) < 2 Then
                                        'S'il y a déjà un objet à cet emplacement, on intervertit les 2 objets.
                                        j = Perso.Objet_Inventaire_Type(i)
                                        'Perso.Objet_Ajouter_Inventaire ObjetSelectionne, Parametres, Commentaires, False
                                        'Perso.Objet_Inventaire_Charger i, Perso.ObjetSelectionne, Parametres
                                        Cmd_Perso_ObjetInventaireCharger Perso.Numero, i, Perso.ObjetSelectionne
                                        'Perso.ObjetSelectionne = j
                                        Cmd_Perso_ObjetSelectionne Perso.Numero, j
                                    End If
                                Else
                                    'Perso.Objet_Inventaire_Charger i, Perso.ObjetSelectionne, Parametres
                                    Cmd_Perso_ObjetInventaireCharger Perso.Numero, i, Perso.ObjetSelectionne
                                    'Perso.ObjetSelectionne = -1
                                    Cmd_Perso_ObjetSelectionne Perso.Numero, -1
                                End If
                            Else
                                'Inventaire de la maison.
                                If Not Parametres.Objet_Inseparable(Perso.ObjetSelectionne) Then
                                    If Maison.Definir_ObjetsInventaire(i) > -1 Then
                                        'S'il y a déjà un objet à cet emplacement, on intervertit les 2 objets.
                                        j = Maison.Definir_ObjetsInventaire(i)
                                        'Maison.Definir_ObjetsInventaire(i) = Perso.ObjetSelectionne
                                        Cmd_Perso_ObjetMaisonCharger Perso.Numero, i, Perso.ObjetSelectionne
                                        'Perso.ObjetSelectionne = j
                                        Cmd_Perso_ObjetSelectionne Perso.Numero, j
                                    Else
                                        'Maison.Definir_ObjetsInventaire(i) = Perso.ObjetSelectionne
                                        Cmd_Perso_ObjetMaisonCharger Perso.Numero, i, Perso.ObjetSelectionne
                                        'Perso.ObjetSelectionne = -1
                                        Cmd_Perso_ObjetSelectionne Perso.Numero, -1
                                    End If
                                End If
                            End If
                        Else 'Equipe l'objet saisi sur le curseur.
                            If Parametres.Objet_Utilisable(Perso.ObjetSelectionne) Then
                                '+++
                                'If Perso.Objet_Equiper(Perso.ObjetSelectionne, Parametres, Commentaires, False) And _
                                '   Parametres.Objet_UtilisationUnique(Perso.ObjetSelectionne) Then
                                '    Perso.ObjetSelectionne = -1
                                'End If
                                Cmd_Perso_ObjetEquiper Perso.Numero, Perso.ObjetSelectionne, False
                            Else
                                'Perso.Objet_Equiper Perso.ObjetSelectionne, Parametres, Commentaires, False, IIf(i - Perso.Nombre_Objets_Inventaire = 1 Or i - Perso.Nombre_Objets_Inventaire = 7, True, False)
                                'Perso.ObjetSelectionne = -1
                                If i - Perso.Nombre_Objets_Inventaire = 1 Or i - Perso.Nombre_Objets_Inventaire = 7 Then
                                    Cmd_Perso_ObjetEquiperADroite Perso.Numero, Perso.ObjetSelectionne, False
                                Else
                                    Cmd_Perso_ObjetEquiper Perso.Numero, Perso.ObjetSelectionne, False
                                End If
                            End If
                        End If
                    End If
                        
                ElseIf CliqueDroit Then
                    If i < Perso.Nombre_Objets_Inventaire Then
                        If Fenetres(4).Bascule = 0 Or Not Perso.DansSaMaison Or Not Perso.Vivant Then
                            'Inventaire du personnage.
                            If Perso.Objet_Inventaire_Actif(i) Then
                                'Perso.Objet_Equiper i, Parametres, Commentaires
                                Cmd_Perso_ObjetEquiper Perso.Numero, i
                            End If
                        Else
                            'Inventaire de la maison.
                            If Maison.Definir_ObjetsInventaire(i) > -1 Then
                                'Perso.Objet_Ajouter_Inventaire Maison.Definir_ObjetsInventaire(i), Parametres, Commentaires, True
                                Cmd_Perso_ObjetInventaireAjouter Perso.Numero, Maison.Definir_ObjetsInventaire(i)
                                'Maison.Definir_ObjetsInventaire(i) = -1
                                Cmd_Perso_ObjetMaisonDetruire Perso.Numero, i
                            End If
                        End If
                    ElseIf i - Perso.Nombre_Objets_Inventaire = 1 And Perso.Objet_Equipes_2_Mains Then
                        If Perso.Objet_Equipes_Actif(i - Perso.Nombre_Objets_Inventaire - 1) Then
                            'Perso.Objet_Ranger i - Perso.Nombre_Objets_Inventaire - 1, Parametres, Commentaires
                            Cmd_Perso_ObjetEquipesRanger Perso.Numero, i - Perso.Nombre_Objets_Inventaire - 1
                        End If
                    Else
                        If Perso.Objet_Equipes_Actif(i - Perso.Nombre_Objets_Inventaire) Then
                            'Perso.Objet_Ranger i - Perso.Nombre_Objets_Inventaire, Parametres, Commentaires
                            Cmd_Perso_ObjetEquipesRanger Perso.Numero, i - Perso.Nombre_Objets_Inventaire
                        End If
                    End If
                ElseIf Perso.ObjetSelectionne = -1 Then
                    If i < Perso.Nombre_Objets_Inventaire Then
                        Fenetres(4).Encadrer_Objet i, AffObjets, SurfBack, True
                        If Fenetres(4).Bascule = 0 Or Not Perso.DansSaMaison Then
                            If Perso.Objet_Inventaire_Actif(i) Then
                                Fenetres(4).BoutonObjets_Afficher_InfoBulle i, Perso, Maison, Parametres, X, Y, AffMessage, AffObjets, SurfBack, True
                            End If
                        Else
                            If Maison.Definir_ObjetsInventaire(i) > -1 Then
                                'Fenetres(4).BoutonObjets_Afficher_InfoBulle i, Perso, Maison, Parametres, X, Y, AffMessage, AffObjets, SurfBack, True, Maison.Definir_ObjetsInventaire(i)
                            End If
                        End If
                    Else
                        If Perso.Objet_Equipes_Actif(i - Perso.Nombre_Objets_Inventaire) Then
                            Fenetres(4).BoutonObjets_Afficher_InfoBulle i, Perso, Maison, Parametres, X, Y, AffMessage, AffObjets, SurfBack, True
                        Else
                            Fenetres(4).Encadrer_Objet i, AffObjets, SurfBack, True
                        End If
                    End If
                Else
                    If i < Perso.Nombre_Objets_Inventaire Then
                        Fenetres(4).Encadrer_Objet i, AffObjets, SurfBack, True
                    Else
                        Fenetres(4).Encadrer_Objet i, AffObjets, SurfBack, True
                    End If
                    'Fenetres(4).BoutonObjets_Afficher_InfoBulle i, Perso, Maison, Parametres, X, Y, AffMessage, AffObjets, SurfBack, True
                'Else
                    'Fenetres(4).BoutonObjets_Afficher_InfoBulle i, Perso, Maison, Parametres, X, Y, AffMessage, AffObjets, SurfBack, True
                End If
'                If CliqueGauche Or CliqueDroit Then
'                    If i < Perso.Nombre_Objets_Inventaire Then
'                        If CliqueDroit And Perso.Vivant Then
'                            If Perso.DansSaMaison Then
'                                Perso.Objet_Deposer_Maison i, Maison, Parametres
'                            ElseIf Perso.DansSonChateau Or Perso.DansUnChateau Then
'                                Perso.Objet_Vendre_Chateau i, Fief, Parametres
'                            Else
'                                Perso.Objet_Deposer_Sol i
'                            End If
'                        Else
'                            Perso.Objet_Equiper i, Parametres, Commentaires
'                        End If
'                    Else
'                        Perso.Objet_Ranger i - Perso.Nombre_Objets_Inventaire, Parametres, Commentaires
'                    End If
'                Else
'                    Fenetres(4).BoutonObjets_Afficher_InfoBulle i, Perso, Maison, Parametres, X, Y, AffMessage, SurfBack, True
'                End If
                If TouchesRaccourciEquipement Then
                    If Touche >= IntTouches.ToucheObjets And Touche <= IntTouches.ToucheObjets + IntTouches.NombreTouchesObjets Then
                        If i < Perso.Nombre_Objets_Inventaire Then
                            If Perso.Objet_Inventaire_Actif(i) Then
                                If Touche = IntTouches.ToucheObjets + IntTouches.NombreTouchesObjets Then
                                    TouchesRaccourciObjets(Perso.Objet_Inventaire_Type(i)) = -1
                                Else
                                    TouchesRaccourciObjets(Perso.Objet_Inventaire_Type(i)) = Touche - 1 + (BoutonRaccourciObjetDecalageY * IntTouches.NombreTouchesObjets)
                                End If
                                Touche = 0
                                'Actualiser_Raccourcis_Objets Perso
                                ActualiserRaccourciObjets = True
                            End If
                        Else
                            If Perso.Objet_Equipes_Actif(i - Perso.Nombre_Objets_Inventaire) Then
                                If Touche = IntTouches.ToucheObjets + IntTouches.NombreTouchesObjets Then
                                    TouchesRaccourciObjets(Perso.Objet_Equipes_Type(i - Perso.Nombre_Objets_Inventaire)) = -1
                                Else
                                    TouchesRaccourciObjets(Perso.Objet_Equipes_Type(i - Perso.Nombre_Objets_Inventaire)) = Touche - 1 + (BoutonRaccourciObjetDecalageY * IntTouches.NombreTouchesObjets)
                                End If
                                Touche = 0
                                'Actualiser_Raccourcis_Objets Perso
                                ActualiserRaccourciObjets = True
                            End If
                        End If
                    End If
                End If
                CurseurSurObjet = True
                NoBouton = 999
                Interieur_Boutons = True
                'i = Fenetres(4).Nombre_Boutons_Objets
                Exit For
            'End If
            End If
        Next i
        If Interieur_Bouton(X, Y, _
                            Fenetres(4).BoutonInventaireMaison_PositionX, _
                            Fenetres(4).BoutonInventaireMaison_PositionY, _
                            Fenetres(4).BoutonInventaireMaison_Largeur, _
                            Fenetres(4).BoutonInventaireMaison_Hauteur) Then
            If Perso.DansSaMaison And Perso.Vivant Then
                If CliqueGauche Or CliqueDroit Then
                    If Perso.ObjetSelectionne = -1 Then
                        'Passe à l'inventaire de la maison.
                        If Fenetres(4).Bascule = 0 Then
                            Fenetres(4).Bascule = 1
                        Else
                            Fenetres(4).Bascule = 0
                        End If
                    Else
                        If Fenetres(4).Bascule = 0 Then
                            'Dépose l'objet dans la maison.
                            If Not Parametres.Objet_Inseparable(Perso.ObjetSelectionne) Then
                                'Perso.Objet_DeposerInventaire_Maison Perso.ObjetSelectionne, Maison, Parametres
                                Cmd_Perso_ObjetMaisonAjouter Perso.Numero, Perso.ObjetSelectionne
                                'Perso.ObjetSelectionne = -1
                                Cmd_Perso_ObjetSelectionne Perso.Numero, -1
                            End If
                        Else
                            'Récupère l'objet.
                            'Perso.Objet_Ajouter_Inventaire Perso.ObjetSelectionne, Parametres, Commentaires, True
                            Cmd_Perso_ObjetInventaireAjouter Perso.Numero, Perso.ObjetSelectionne
                            'Perso.ObjetSelectionne = -1
                            Cmd_Perso_ObjetSelectionne Perso.Numero, -1
                        End If
                    End If
                Else
                    'Affiche l'infobulle.
                    Fenetres(4).BoutonInventaireMaison_Afficher_InfoBulle Perso.ObjetSelectionne <> -1, Perso, X, Y, AffMessage, SurfBack
                End If
            End If
            CurseurSurObjet = True
            NoBouton = 999
            Interieur_Boutons = True
        End If
        'Onglets de la fenêtre inventaire.
        If Fenetres(4).Nombre_Boutons_InventaireOnglets > 1 Then
        If Not Interieur_Boutons Then
            For i = 0 To Fenetres(4).Nombre_Boutons_InventaireOnglets - 1
                If Interieur_Bouton(X, Y, _
                                    Fenetres(4).BoutonInventaireOnglet_PositionX(i), _
                                    Fenetres(4).BoutonInventaireOnglet_PositionY(i), _
                                    Fenetres(4).BoutonInventaireOnglet_Largeur, _
                                    Fenetres(4).BoutonInventaireOnglet_Hauteur) Then
                    If CliqueGauche Or CliqueDroit Then
                        Fenetres(4).InventaireOngletsChoix = i
                    End If
                    NoBouton = 999
                    Interieur_Boutons = True
                    Exit For
                End If
            Next i
        End If
        End If
        
        'If Perso.ObjetSelectionne >= 0 Then
        '    Fenetres(4).BoutonObjets_Afficher_InfoBulle 0, Perso, Maison, Parametres, X, Y, AffMessage, AffObjets, SurfBack, True, Parametres.Objet_Categorie(Perso.ObjetSelectionne)
        'End If
    End If
    'Enregistre la configuration de l'équipement avec les touches Ctrl+n°
    If Ctrl And Touche >= 2 And Touche <= 11 Then
        For i = 0 To Perso.Nombre_Objets_Equipes - 1
            If Perso.Objet_Equipes_Actif(i) Then
                TouchesRaccourciObjets(Perso.Objet_Equipes_Type(i)) = Touche - 1 + (BoutonRaccourciObjetDecalageY * IntTouches.NombreTouchesObjets)
            End If
        Next i
        'Actualiser_Raccourcis_Objets Perso
        ActualiserRaccourciObjets = True
    End If
    
    'Boutons dans la fenêtre bestiaire.
    If Not Interieur_Boutons And Fenetres(5).Visible Then
        'Boutons de défilement.
        If Not Interieur_Boutons Then
            Interieur_Boutons = Interieur_BoutonsDefilement(X, Y, 5, NoBouton, CliqueGauche, CliqueEnfonceGauche, SurfBack)
        End If
    End If
    
    'Boutons dans toutes les fenêtres.
    If Not Interieur_Boutons Then
        For i = 0 To UBound(Fenetres())
            If Fenetres(i).Visible Then
                'Bouton fermer.
                If Interieur_Bouton(X, Y, _
                                    Fenetres(i).BoutonFermer_PositionX, _
                                    Fenetres(i).BoutonFermer_PositionY, _
                                    Fenetres(i).BoutonFermer_Largeur, _
                                    Fenetres(i).BoutonFermer_Hauteur) Then
                    If CliqueGauche Then
                        Activer_EffetVisuel_Clique X, Y
                        If i = 7 Then
                            Actionner_Boutons_Infos 8, NoJoueur
                        Else
                            Actionner_Boutons_Infos i, NoJoueur
                        End If
                    Else
                        Fenetres(i).BoutonFermer_Afficher_InfoBulle X, Y, AffMessage, SurfBack
                    End If
                    NoBouton = 999
                    Interieur_Boutons = True
                    'i = UBound(Fenetres())
                    Exit For
                End If
                'Bouton deplacer.
                If Interieur_Bouton(X, Y, _
                                    Fenetres(i).BoutonDeplacer_PositionX, _
                                    Fenetres(i).BoutonDeplacer_PositionY, _
                                    Fenetres(i).BoutonDeplacer_Largeur, _
                                    Fenetres(i).BoutonDeplacer_Hauteur) Then
                    If CliqueGauche Then
                        Activer_EffetVisuel_Clique X, Y
                        Fenetres(i).Deplace = True
                        DeplacerFenetre = i
                    ElseIf CliqueDroit Then
                        Activer_EffetVisuel_Clique X, Y
                        Fenetres(i).Deplace = False
                        Fenetres(i).PositionX = FenetrePositionXInit(i)
                        Fenetres(i).PositionY = FenetrePositionYInit(i)
                    Else
                        Fenetres(i).BoutonDeplacer_Afficher_InfoBulle X, Y, AffMessage, SurfBack
                    End If
                    NoBouton = 999
                    Interieur_Boutons = True
                    'i = UBound(Fenetres())
                    Exit For
                End If
            End If
        Next i
    End If
    'Boutons prendre argent.
    If Not Interieur_Boutons Then
        For i = 0 To UBound(Fenetres())
            If i = 1 Or i = 4 Then
                If Fenetres(i).Visible Then
                    If Interieur_Bouton(X, Y, _
                                        Fenetres(i).BoutonArgentBoutonsPrendre_PositionX, _
                                        Fenetres(i).BoutonArgentBoutonsPrendre_PositionY, _
                                        Fenetres(i).BoutonArgentBoutonsPrendre_Largeur, _
                                        Fenetres(i).BoutonArgentBoutonsPrendre_Hauteur) Then
                        If CliqueGauche Then
                            Activer_EffetVisuel_Clique X, Y
                            'Perso.Argent_Prendre 100
                            Cmd_Perso_ArgentPrendre Perso.Numero, 100
                        ElseIf CliqueDroit Then
                            Activer_EffetVisuel_Clique X, Y
                            'Perso.Argent_Prendre 0, True
                            Cmd_Perso_ArgentPrendre Perso.Numero, Perso.Definir_Argent
                        Else
                            Fenetres(i).BoutonArgentBoutonsPrendreAfficher_Info_Bulle X, Y, AffMessage, SurfBack
                        End If
                        NoBouton = 999
                        Interieur_Boutons = True
                        'i = UBound(Fenetres())
                        Exit For
                    End If
                End If
            End If
        Next i
    End If
    
    'Boutons de déplacement.
    CurseurSurBoutonDeplacement = False
    If Not Interieur_Boutons And AfficherBoutonsDeplacement Then
        For i = 0 To UBound(BoutonsDeplacement())
            If BoutonsDeplacement(i).Visible And _
               Interieur_Bouton(X, Y, _
                                BoutonsDeplacement(i).PositionX, _
                                BoutonsDeplacement(i).PositionY, _
                                AffBoutonsDeplacement.Largeur, _
                                AffBoutonsDeplacement.Hauteur) Then
                    BoutonsDeplacement(i).Enabled = CliqueEnfonceGauche
                    Select Case i
                        Case 0: NoBouton = 200
                        Case 1:
                                If Perso.Vivant Or SaMaison.Service = 1 Then
                                    NoBouton = 101
                                Else
                                    NoBouton = 999
                                End If
                        Case 2 To 9: NoBouton = 999
                                If CliqueGauche And Perso.Vivant Then
                                    'IA_Trouver_Ressource Perso.Numero, i - 2
                                    Cmd_Perso_TrouverRessource Perso.Numero, i - 2
                                End If
                        Case 10: NoBouton = 999
                                If (CliqueGauche Or CliqueDroit) And Perso.Vivant Then
                                    If ArmeeSelectionnee Then
                                        Ordonner_Attaquer_Perso Perso, SoldatsSelectionnes, Trouver_Ennemi(Perso.Numero)
                                    Else
                                        'IA_Trouver_Ennemi Perso.Numero, CliqueDroit Or Ctrl Or Shift Or Alt, Ctrl, Shift, Alt
                                        Cmd_Perso_TrouverPersoEnnemi Perso.Numero, CliqueDroit Or Ctrl Or Shift Or Alt, Ctrl, Shift, Alt
                                    End If
                                End If
                        Case 11: NoBouton = 999
                                If (CliqueGauche Or CliqueDroit) And Perso.Vivant Then
                                    If ArmeeSelectionnee Then
                                        Ordonner_Attaquer_Maison Perso, SoldatsSelectionnes, Trouver_MaisonEnnemi(Perso.Numero)
                                    Else
                                        'IA_Trouver_MaisonEnnemi Perso.Numero, CliqueDroit Or Ctrl Or Shift Or Alt, Ctrl, Shift, Alt
                                        Cmd_Perso_TrouverMaisonEnnemi Perso.Numero, CliqueDroit Or Ctrl Or Shift Or Alt, Ctrl, Shift, Alt
                                    End If
                                End If
                        Case 12: NoBouton = 999
                                If (CliqueGauche Or CliqueDroit) And Perso.Vivant Then
                                    If ArmeeSelectionnee Then
                                        Ordonner_Attaquer_Chateau Perso, SoldatsSelectionnes, Trouver_ChateauEnnemi(Perso.Numero)
                                    Else
                                        'IA_Trouver_ChateauEnnemi Perso.Numero, CliqueDroit Or Ctrl Or Shift Or Alt, Ctrl, Shift, Alt
                                        Cmd_Perso_TrouverChateauEnnemi Perso.Numero, CliqueDroit Or Ctrl Or Shift Or Alt, Ctrl, Shift, Alt
                                    End If
                                End If
                        Case 13: NoBouton = 999
                                If (CliqueGauche Or CliqueDroit) And Perso.Vivant Then
                                    'IA_Trouver_Tresor Perso.Numero
                                    Cmd_Perso_TrouverTresor Perso.Numero
                                End If
                        Case 14: NoBouton = 999
                                If CliqueGauche Or CliqueDroit Then
                                    'Perso.IA = Not Perso.IA
                                    Cmd_Perso_Definir_IA Perso.Numero, Not Perso.IA
                                End If
                        Case 15: NoBouton = 999
                                If (CliqueGauche Or CliqueDroit) And Perso.Vivant Then
                                    'Appeler_De_lAide Perso.Numero
                                    'Cmd_Perso_AppelerDelAide Perso.Numero
                                    'Cmd_Perso_AttaquerAllies Perso.Numero
                                    FenetreConfirmation.Afficher_Fenetre 1, MessageConfirmationAttaquer, DecalageX, DecalageY, Largeur, Hauteur
                                End If
                        Case 16: NoBouton = 999
                                If CliqueGauche Then
                                    If ArmeeSelectionnee And Perso.Vivant Then
                                        Ordonner_Arret Perso.Numero, ArmeeSelectionnee, SoldatsSelectionnes
                                    Else
                                        'Perso.Arreter
                                        Cmd_Perso_Arreter Perso.Numero
                                    End If
                                ElseIf CliqueDroit And _
                                       Perso.Vivant Then
                                       Ordonner_Arret Perso.Numero, False, SoldatsSelectionnes
                                End If
'                        Case 17: NoBouton = 999
'                                 If (CliqueGauche Or CliqueDroit) And Perso.Vivant Then
'                                    Ordonner_Chargez Perso.Numero, , , Ctrl, Shift, Alt
'                                 End If
'                        Case 18: NoBouton = 999
'                                 If CliqueGauche Or CliqueDroit Then
'                                    Perso.Changer_Attitude
'                                 End If
                        Case 17: NoBouton = 999
                                 If CliqueGauche Then
                                    'Ordonner_Formation Perso.Numero, True
                                    'Cmd_Perso_OrdonnerFormation Perso.Numero, True
                                    Cmd_Perso_OrdonnerFormation Perso.Numero, True
                                 'ElseIf CliqueDroit Then
                                 '   Ordonner_Formation Perso.Numero
                                 End If
                        Case Is > 17:
                                If CliqueGauche Then
                                    Bouton_ChangerAttitude i - 18, Perso, ArmeeSelectionnee, SoldatsSelectionnes
                                End If
                    End Select
                    If CliqueGauche Or CliqueDroit Then
                        Activer_EffetVisuel_Clique X, Y
                    Else
                        If i >= 2 And i <= 9 Then
                            BoutonsDeplacement(i).InfoBulle = BoutonDeplacementInfoBulle(i) & AffMessage.Caractere_RetourLigne & _
                                                              Parametres.Ressources_Nom(Parametres.TerrainRessources_TypeRessource(i - 2)) & " : " & _
                                                              Int(Perso.Definir_ressources(Parametres.TerrainRessources_TypeRessource(i - 2))) & " / " & Perso.Definir_MaxRessources(Parametres.TerrainRessources_TypeRessource(i - 2))
                        End If
                        BoutonsDeplacement(i).Afficher_Info_Bulle X, Y, AffMessage, SurfBack
                    End If
                    Interieur_Boutons = True
                    'If i < 18 Then
                    CurseurSurBoutonDeplacement = True
                    'i = UBound(BoutonsDeplacement())
                    Exit For
            End If
        Next i
    End If
    
    'Boutons d'informations.
    If Not Interieur_Boutons And AfficherBoutonsInfos Then
        For i = 0 To NombreBoutonsInfos - 1
            If i <> 8 Or AfficherObjectifs Then
            If Interieur_Bouton(X, Y, _
                                BoutonsInfos(i).PositionX, _
                                BoutonsInfos(i).PositionY, _
                                AffBoutonsInfos.Largeur, _
                                AffBoutonsInfos.Hauteur) Then
                        If CliqueGauche Or CliqueDroit Then
                            Activer_EffetVisuel_Clique X, Y
                            Actionner_Boutons_Infos i, NoJoueur
                        Else
                            BoutonsInfos(i).Afficher_Info_Bulle X, Y, AffMessage, SurfBack
                        End If
                    NoBouton = 999
                    Interieur_Boutons = True
                    'i = UBound(BoutonsInfos())
                    Exit For
            End If
            End If
        Next i
    End If
    
    'Boutons de raccourci objets.
    If Not Interieur_Boutons And AfficherRaccourciObjets Then
        For i = 0 To NombreRaccourciObjets - 1
            If Interieur_Bouton(X, Y, _
                                BoutonsRaccourciObjets(i).PositionX, _
                                BoutonsRaccourciObjets(i).PositionY, _
                                AffBoutonsRaccourciObjets.Largeur, _
                                AffBoutonsRaccourciObjets.Hauteur) Then
                If CliqueDroit Then
                    'Efface tous les raccourcis.
                    Activer_EffetVisuel_Clique X, Y
                    For j = 0 To UBound(TouchesRaccourciObjets())
                        If TouchesRaccourciObjets(j) = i + 1 + BoutonRaccourciObjetDecalageY * IntTouches.NombreTouchesObjets Then
                            TouchesRaccourciObjets(j) = 0
                        End If
                    Next j
                    'Actualiser_Raccourcis_Objets Perso
                    ActualiserRaccourciObjets = True
                ElseIf CliqueGauche Then
                    Activer_EffetVisuel_Clique X, Y
                    If Perso.ObjetSelectionne = -1 Then
                        'S'équipe avec les objets sélectionnés.
                        Raccourcis_Objets_Equiper i + 1 + BoutonRaccourciObjetDecalageY * IntTouches.NombreTouchesObjets, Perso
                    Else
                        'Ajoute l'objet sélectionné avec le curseur.
                        TouchesRaccourciObjets(Perso.ObjetSelectionne) = i + 1 + BoutonRaccourciObjetDecalageY * IntTouches.NombreTouchesObjets
                        'Perso.Definir_ObjetSelectionne = False
                        Cmd_Perso_ObjetDeSelectionne Perso.Numero
                    End If
                    'Actualiser_Raccourcis_Objets Perso
                    ActualiserRaccourciObjets = True
                Else
                    BoutonsRaccourciObjets(i).InfoBulle = BoutonRaccourciObjetEtiquette & "(" & IIf(Alt, "Alt+", IIf(Shift, "Maj+", "")) & i + 1 & ")"
                    For j = 0 To UBound(TouchesRaccourciObjets())
                        If TouchesRaccourciObjets(j) = i + 1 + BoutonRaccourciObjetDecalageY * IntTouches.NombreTouchesObjets Then
                            'If Perso.Objet_NombreType(j) > 0 Then
                            BoutonsRaccourciObjets(i).InfoBulle = BoutonsRaccourciObjets(i).InfoBulle & "#" & IIf(Perso.Objet_NombreType(j) > 0, "", "&") & "- " & _
                                                                  IIf(Perso.Objet_NombreType(j) > 1, Parametres.Objet_NomPluriel(j), Parametres.Objet_Nom(j))
                            'End If
                        End If
                    Next j
                    BoutonsRaccourciObjets(i).Afficher_Info_Bulle X, Y, AffMessage, SurfBack
                End If
                NoBouton = 999
                Interieur_Boutons = True
                'i = NombreRaccourciObjets
                Exit For
            End If
        Next i
    End If
    'Decale les boutons de raccourci selon la touche en question.
    If Alt Then
        If BoutonRaccourciObjetDecalageY <> 1 Then
            BoutonRaccourciObjetDecalageY = 1
            'Actualiser_Raccourcis_Objets Perso
            ActualiserRaccourciObjets = True
        End If
    ElseIf Shift Then
        If BoutonRaccourciObjetDecalageY <> 2 Then
            BoutonRaccourciObjetDecalageY = 2
            'Actualiser_Raccourcis_Objets Perso
            ActualiserRaccourciObjets = True
        End If
    Else
        If BoutonRaccourciObjetDecalageY <> 0 Then
            BoutonRaccourciObjetDecalageY = 0
            'Actualiser_Raccourcis_Objets Perso
            ActualiserRaccourciObjets = True
        End If
    End If
    
    'Boutons d'effets temporaires.
    If Not Interieur_Boutons And AfficherEffetsTemporaires Then
        If Perso.Nombre_EffetsTemp > 0 Then
            For i = 1 To Perso.Nombre_EffetsTemp
                If Interieur_Bouton(X, Y, _
                                    BoutonsEffetsTemporaires(i).PositionX, _
                                    BoutonsEffetsTemporaires(i).PositionY, _
                                    AffObjets.Largeur, _
                                    AffObjets.Hauteur) Then
                    If CliqueDroit Then
                        Activer_EffetVisuel_Clique X, Y
                        'If Perso.Nombre_EffetsTemp >= i Then
                        If Not Parametres.Objet_EffetAnnulationImpossible(Perso.Definir_EffetTemp_Numero(i)) Then
                            'Perso.EffetTemp_Effacer i
                            Cmd_Perso_EffetTempEffacer Perso.Numero, i
                        End If
                        'End If
                    ElseIf CliqueGauche Then 'Relance le même effet.
                            Activer_EffetVisuel_Clique X, Y
                            For j = 0 To Perso.Nombre_Objets_Inventaire - 1
                                If Perso.Objet_Inventaire_Actif(j) Then
                                    If Perso.Objet_Inventaire_Apparence(j) = Perso.Definir_EffetTemp_Numero(i) Then
                                        'Perso.Objet_Equiper j, Parametres, Commentaires, True
                                        Cmd_Perso_ObjetEquiper Perso.Numero, j
                                    End If
                                End If
                            Next j
                    Else
                        Fenetres(4).BoutonObjets_Afficher_InfoBulle 0, Perso, Maison, Parametres, X, Y, AffMessage, AffObjets, SurfBack, True, Perso.Definir_EffetTemp_Numero(i), Perso.Definir_EffetTemp_Duree(i), i
                    End If
                    NoBouton = 999
                    Interieur_Boutons = True
                    Exit For
                    'If Perso.Nombre_EffetsTemp > 0 Then
                    '    i = Perso.Nombre_EffetsTemp - 1
                    'Else
                    '    i = 1
                    'End If
                End If
            Next i
        End If
    End If
    
    'Boutons sur les barres de statut.
    If Not Interieur_Boutons And AfficherBarresStatut Then
        For i = 0 To UBound(BoutonsCaracteristiques())
            If Interieur_Bouton(X, Y, _
                                BoutonsCaracteristiques(i).PositionX, _
                                BoutonsCaracteristiques(i).PositionY, _
                                BoutonsCaracteristiques(i).Largeur, _
                                BoutonsCaracteristiques(i).Hauteur) Then
                        If CliqueGauche Or CliqueDroit Then
                            If CliqueDroit Then
                                Activer_EffetVisuel_Clique X, Y
                                AutoRecuperationCarac(i) = Not AutoRecuperationCarac(i)
                            ElseIf Perso.Vivant Then
                                Activer_EffetVisuel_Clique X, Y
                                Select Case i
                                Case 0:
                                    If Perso.Definir_Vie < Perso.Definir_MaxVie Then
                                        'Perso.Recuperer_Vie Parametres
                                        Cmd_Perso_RecupererCarac Perso.Numero, 0
                                    End If
                                Case 1:
                                    If Perso.Definir_Energie < Perso.Definir_MaxEnergie Then
                                        'Perso.Recuperer_Energie Parametres
                                        Cmd_Perso_RecupererCarac Perso.Numero, 1
                                    End If
                                Case 2:
                                    If Perso.Definir_Magie < Perso.Definir_MaxMagie Then
                                        'Perso.Recuperer_Magie Parametres
                                        Cmd_Perso_RecupererCarac Perso.Numero, 2
                                    End If
                                Case 3:
                                    If Perso.Definir_Moral < Perso.Definir_MaxMoral Then
                                        'Perso.Recuperer_Moral Parametres
                                        Cmd_Perso_RecupererCarac Perso.Numero, 3
                                    End If
                                End Select
                            End If
                        Else
                            If Z <> 0 Then 'Régule le niveau d'auto restauration avec la roulette de la souris.
                                If Z < 0 And Z > -0.025 * 120 Then
                                    Z = -120
                                ElseIf Z > 0 And Z < 0.025 * 120 Then
                                    Z = 120
                                End If
                                AutoRecuperationCarac(i) = True
                                AutoRecuperationCaracPcent(i) = AutoRecuperationCaracPcent(i) + Int(Z / 120) * 0.025
                                AutoRecuperationCaracPcent(i) = CSng(Format(AutoRecuperationCaracPcent(i), "0.00"))
                                If AutoRecuperationCaracPcent(i) < 0 Then
                                    AutoRecuperationCaracPcent(i) = 0
                                ElseIf AutoRecuperation_CaracPcent(i) > 1 Then
                                    AutoRecuperationCaracPcent(i) = 1
                                End If
                                Actualiser_BarreStatut i
                                X = BoutonsCaracteristiquesSeuil(i).PositionX
                            End If
                            If AutoRecuperationCarac(i) Then
                                If Interieur_Bouton(X, Y, _
                                   BoutonsCaracteristiquesSeuil(i).PositionX, _
                                   BoutonsCaracteristiquesSeuil(i).PositionY, _
                                   BoutonsCaracteristiquesSeuil(i).Largeur, _
                                   BoutonsCaracteristiquesSeuil(i).Hauteur) Then
                                    BoutonsCaracteristiquesSeuil(i).InfoBulle = BarresStatutSeuilMessage1 & _
                                       "->" & Int(IIf(i = 0, Perso.Definir_MaxVie, IIf(i = 1, Perso.Definir_MaxEnergie, IIf(i = 2, Perso.Definir_MaxMagie, Perso.Definir_MaxMoral))) * AutoRecuperationCaracPcent(i)) & " / " & IIf(i = 0, Perso.Definir_MaxVie, IIf(i = 1, Perso.Definir_MaxEnergie, IIf(i = 2, Perso.Definir_MaxMagie, Perso.Definir_MaxMoral))) & " (" & Int(AutoRecuperationCaracPcent(i) * 100) & "%)" & _
                                       BarresStatutSeuilMessage2
                                        BoutonsCaracteristiquesSeuil(i).Afficher_Info_Bulle X, Y, AffMessage, SurfBack
                                Else
                                    Select Case i
                                    Case 0:
                                        BoutonsCaracteristiques(0).InfoBulle = Format(Perso.Vie, BarresStatutFormatCarac) & " / " & Perso.Definir_MaxVie
                                    Case 1:
                                        BoutonsCaracteristiques(1).InfoBulle = Format(Perso.Energie, BarresStatutFormatCarac) & " / " & Perso.Definir_MaxEnergie
                                    Case 2:
                                        BoutonsCaracteristiques(2).InfoBulle = Format(Perso.Magie, BarresStatutFormatCarac) & " / " & Perso.Definir_MaxMagie
                                    Case 3:
                                        BoutonsCaracteristiques(3).InfoBulle = Format(Perso.Moral, BarresStatutFormatCarac) & " / " & Perso.Definir_MaxMoral
                                    End Select
                                    BoutonsCaracteristiques(i).Afficher_Info_Bulle X, Y, AffMessage, SurfBack
                                End If
                            Else
                                Select Case i
                                Case 0:
                                    BoutonsCaracteristiques(0).InfoBulle = Format(Perso.Vie, BarresStatutFormatCarac) & " / " & Perso.Definir_MaxVie
                                Case 1:
                                    BoutonsCaracteristiques(1).InfoBulle = Format(Perso.Energie, BarresStatutFormatCarac) & " / " & Perso.Definir_MaxEnergie
                                Case 2:
                                    BoutonsCaracteristiques(2).InfoBulle = Format(Perso.Magie, BarresStatutFormatCarac) & " / " & Perso.Definir_MaxMagie
                                Case 3:
                                    BoutonsCaracteristiques(3).InfoBulle = Format(Perso.Moral, BarresStatutFormatCarac) & " / " & Perso.Definir_MaxMoral
                                End Select
                                BoutonsCaracteristiques(i).Afficher_Info_Bulle X, Y, AffMessage, SurfBack
                            End If
                        End If
                    CurseurSurBarreStatut = True
                    NoBouton = 999
                    Interieur_Boutons = True
                    'i = UBound(BoutonsCaracteristiques())
                    Exit For
            End If
        Next i
    End If
    
    'Cadre extérieur.
    If Not Interieur_Boutons Then
        If X > BordureGauche And X < Ecran.Width - BordureDroite And _
           Y > BordureHaut And Y < Ecran.Height - BordureBas Then
           'For i = 0 To NombreFenetres - 1
           '     If Fenetres(i).Visible And _
                   X >= Fenetres(i).PositionX And _
                   X <= Fenetres(i).PositionX + AffFenetre.Largeur And _
                   Y >= Fenetres(i).PositionY And _
                   Y <= Fenetres(i).PositionY + AffFenetre.Hauteur Then
           '         Interieur_Boutons = True
           '         i = NombreFenetres
           '     End If
           'Next i
        Else
            Interieur_Boutons = True
        End If
    End If
    SurfBack.SetForeColor 0
    End If
End Function

Private Sub FileFabricationObjets_Changer(ByVal Objet As Integer, _
                                          ByVal Perso As ClsJeuPerso, _
                                          Optional ByVal Ajout As Boolean = True, _
                                          Optional ByVal KCtrl As Boolean = False, _
                                          Optional ByVal KShift As Boolean = False)
    'Ajoute ou enlève des objets à la fabrication.
    Dim i As Integer
    Dim j As Integer
    Dim Quantite As Integer
    If KShift Then
        If KCtrl Then
            Quantite = 20
        Else
            Quantite = 10
        End If
    ElseIf KCtrl Then
        Quantite = 5
    Else
        Quantite = 1
    End If

    For i = 0 To Quantite - 1
        If Ajout Then
            FileFabricationObjets.Add Objet
        ElseIf FileFabricationObjets.Count > 0 Then
            'Annule une fabrication en cours.
            For j = 1 To FileFabricationObjets.Count
                If FileFabricationObjets.Item(j) = Objet Then
                    FileFabricationObjets.Remove j
                    j = FileFabricationObjets.Count + 2
                End If
            Next j
            If j <> FileFabricationObjets.Count + 3 Then
                If Collection_NombreElements(Objet, FileFabricationObjets) = 0 Then
                    Cmd_Perso_ObjetArreterFabrication Perso.Numero
                End If
            End If
        Else
            Cmd_Perso_ObjetArreterFabrication Perso.Numero
        End If
    Next i
End Sub

Public Sub Bouton_ChangerAttitude(ByVal Attitude As Integer, _
                                  ByVal Perso As ClsJeuPerso, _
                                  ByRef ArmeeSelectionnee As Boolean, _
                                  ByRef SoldatsSelectionnes As Collection)
    Dim j As Long
    Ordonner_Activer_IA Perso.Numero
    For j = 1 To SoldatsSelectionnes.Count
        Cmd_Perso_DefinirIAAttitude SoldatsSelectionnes(j), Attitude
    Next j
    ArmeeSelectionnee = False
    Collection_Vider SoldatsSelectionnes
End Sub

Private Sub Definir_InfoBulle_Batiment(ByVal i As Integer, _
                                       ByVal Perso As ClsJeuPerso, _
                                       Optional ByVal Details As Double)
    Dim j, k, l, m As Integer
    Dim InfoBulle As String
    Dim InfoBulleObjet As String
    Dim Temp As Boolean
    With Parametres
    If i = 0 Then
        BoutonsConstruire(0).InfoBulle = IIf(BoutonsConstruire(0).Enabled, "", "&") & BoutonDemolirEtiquette
    Else
        m = i
        i = BoutonsConstruire_IndiceBatiment(i) + 1
        InfoBulle = .Batiment_Nom(i - 1)
        If Fiefs(Perso.NumeroFief).Epoque < .Batiment_NoEpoque(i - 1) - 1 Then
            InfoBulle = InfoBulle & Chr$(13) & _
                        "&(" & .Epoque_Nom(.Batiment_NoEpoque(i - 1) - 1) & ")"
        End If
        'Coût en pièces d'agent.
        If .Batiment_PrixArgent(i - 1) > 0 Then
            InfoBulle = InfoBulle & Chr$(13)
            If Perso.Definir_Argent < .Batiment_PrixArgent(i - 1) Then
                InfoBulle = InfoBulle & "&(" & Int(Perso.Definir_Argent) & " / "
            Else
                InfoBulle = InfoBulle & "("
            End If
            InfoBulle = InfoBulle & .Batiment_PrixArgent(i - 1) & " " & Parametres.EtiquettePieces & ")"
        End If
        'Coût en ressources.
        For j = 0 To .NombreRessources - 1
            If .Batiment_PrixRessource(i - 1, j) > 0 Then
                InfoBulle = InfoBulle & Chr$(13)
                If Perso.Definir_ressources(j) < .Batiment_PrixRessource(i - 1, j) Then
                    InfoBulle = InfoBulle & "&(" & Int(Perso.Definir_ressources(j)) & " / "
                Else
                    InfoBulle = InfoBulle & "("
                End If
                InfoBulle = InfoBulle & _
                            .Batiment_PrixRessource(i - 1, j) & " " & _
                            LCase$(.Ressources_Nom(j)) & ")"
            End If
        Next j
        'Réservé à l'empereur.
        If .Batiment_Empereur(i - 1) Then
            If Not Perso.Definir_Empereur Then 'Le personnage n'est pas empereur.
                InfoBulle = InfoBulle & Chr$(13) & "&" & .EtiquetteBatimentEmpereur
            End If
        End If
        'Description.
        InfoBulle = InfoBulle & Chr$(13) & .Batiment_Description(i - 1)
        If Details > 0 And .Batiment_Fabrique(i - 1) Then
            For k = 0 To .Batiment_NombreObjetsFabriques(i - 1) - 1
                InfoBulle = InfoBulle & Chr$(13)
                InfoBulleObjet = "- " & .Objet_NomPluriel(.Batiment_ObjetFabrique(i - 1, k)) '& " ("
                Temp = False
                For l = 0 To Parametres.NombreRessources - 1
                    If .Objet_PrixRessources(.Batiment_ObjetFabrique(i - 1, k), l) > 0 Then
                        If Temp Then
                            InfoBulleObjet = InfoBulleObjet & ", "
                        Else
                            InfoBulleObjet = InfoBulleObjet & " ("
                        End If
                        If .Objet_PrixRessources(.Batiment_ObjetFabrique(i - 1, k), l) > Perso.Definir_ressources(l) Then
                            If Left(InfoBulleObjet, 1) <> "&" Then InfoBulleObjet = "&" & InfoBulleObjet
                            InfoBulleObjet = InfoBulleObjet & Int(Perso.Definir_ressources(l)) & "/" & .Objet_PrixRessources(.Batiment_ObjetFabrique(i - 1, k), l)
                        Else
                            InfoBulleObjet = InfoBulleObjet & .Objet_PrixRessources(.Batiment_ObjetFabrique(i - 1, k), l)
                        End If
                        InfoBulleObjet = InfoBulleObjet & " " & .Ressources_Nom(l)
                        Temp = True
                    End If
                Next l
                If Temp Then InfoBulleObjet = InfoBulleObjet & ")"
                If .Objet_Epoque(.Batiment_ObjetFabrique(i - 1, k)) - 1 > Fiefs(Perso.NumeroFief).Epoque Then
                    InfoBulleObjet = InfoBulleObjet & " (" & .Epoque_Nom(.Objet_Epoque(.Batiment_ObjetFabrique(i - 1, k)) - 1) & ")"
                End If
                If .Objet_Epoque(.Batiment_ObjetFabrique(i - 1, k)) - 1 > Fiefs(Perso.NumeroFief).Epoque Then
                    If Left(InfoBulleObjet, 1) <> "&" Then InfoBulleObjet = "&" & InfoBulleObjet
                End If
                InfoBulle = InfoBulle & InfoBulleObjet
                If k + 1 >= Int(Details) Then k = .Batiment_NombreObjetsFabriques(i - 1)
            Next k
        End If
        BoutonsConstruire(m).InfoBulle = InfoBulle
    End If
    End With
End Sub

Private Function Interieur_BoutonsDefilement(ByVal X As Long, _
                                             ByVal Y As Long, _
                                             ByVal NoFenetre As Integer, _
                                             ByRef NoBouton As Long, _
                                             ByVal Clique As Boolean, _
                                             ByVal CliqueMaintenuEnfonce As Boolean, _
                                             ByRef SurfBack As IDirectDrawSurface7, _
                                             Optional ByVal Fabrique As Boolean) As Boolean
    Dim Temp As Single
    Dim i As Integer
    'With Fenetres(NoFenetre)
    For i = 0 To Fenetres(NoFenetre).Nombre_Boutons_Defilement - 1
        If Interieur_Bouton(X, Y, _
                            Fenetres(NoFenetre).BoutonsDefilementPositionX(i), _
                            Fenetres(NoFenetre).BoutonsDefilementPositionY(i), _
                            Fenetres(NoFenetre).BoutonsDefilementLargeur(i), _
                            Fenetres(NoFenetre).BoutonsDefilementHauteur(i)) Then
            If Clique Then Activer_EffetVisuel_Clique X, Y
            If Clique Or CliqueMaintenuEnfonce Then
                Select Case i
                Case 0:
                    If Clique Then
                        Defilement_Haut 1, NoFenetre, Fabrique
                    End If
                Case 1:
                    If Clique Then
                        Defilement_Bas 1, NoFenetre, Fabrique
                    End If
                Case 2:
                    If CliqueMaintenuEnfonce Then
                        Temp = Int(Fenetres(NoFenetre).DefilementNombreLigne * _
                              (Y - Fenetres(NoFenetre).BoutonsDefilementPositionY(i)) / _
                               Fenetres(NoFenetre).BoutonsDefilementHauteur(i) + 0.5)
                        If Fabrique Then
                            Fenetres(NoFenetre).DefilementObjetDecalageX = Temp
                        Else
                            If NoFenetre = 3 Or NoFenetre = 5 Then
                                Fenetres(NoFenetre).DefilementDecalageX = Temp
                            Else
                                Fenetres(1).DefilementDecalageX = Temp
                                Fenetres(0).DefilementDecalageX = Temp
                            End If
                        End If
                    End If
                End Select
            Else
                If Fenetres(NoFenetre).BoutonsDefilementVisible(i) Then
                    Fenetres(NoFenetre).BoutonDefilement_Afficher_InfoBulle i, X, Y, AffMessage, SurfBack
                End If
            End If
            NoBouton = 999
            Interieur_BoutonsDefilement = True
            i = Fenetres(NoFenetre).Nombre_Boutons_Defilement
        End If
    Next i
    'End With
End Function

Public Sub Defilement_Haut(ByVal Valeur As Single, ByVal NoFenetre As Integer, Optional ByVal Fabrique As Boolean)
    'Défilement vers le haut.
    If Valeur < -1 Then
        Valeur = -1
    ElseIf Valeur > 1 Then
        Valeur = 1
    End If
    If Fabrique Then
        If Fenetres(NoFenetre).DefilementObjetDecalageX > 0 Then
            Fenetres(NoFenetre).DefilementObjetDecalageX = Fenetres(NoFenetre).DefilementObjetDecalageX - Valeur
        End If
    Else
        'If Fenetres(NoFenetre).DefilementDecalageX > 0 Then
        If NoFenetre = 3 Or NoFenetre = 5 Then
            Fenetres(NoFenetre).DefilementDecalageX = Fenetres(NoFenetre).DefilementDecalageX - Valeur
            If Fenetres(NoFenetre).DefilementDecalageX < 0 Then Fenetres(NoFenetre).DefilementDecalageX = 0
        Else
            Fenetres(1).DefilementDecalageX = Fenetres(NoFenetre).DefilementDecalageX - Valeur
            If Fenetres(1).DefilementDecalageX < 0 Then Fenetres(1).DefilementDecalageX = 0
            Fenetres(0).DefilementDecalageX = Fenetres(1).DefilementDecalageX
            If Fenetres(0).DefilementDecalageX < 0 Then Fenetres(0).DefilementDecalageX = 0
        End If
        'End If
    End If
End Sub

Public Sub Defilement_Bas(ByVal Valeur As Single, ByVal NoFenetre As Integer, Optional ByVal Fabrique As Boolean)
    'Défilement vers le bas.
    If Valeur < -1 Then
        Valeur = -1
    ElseIf Valeur > 1 Then
        Valeur = 1
    End If
    If Fabrique Then
        If Fenetres(NoFenetre).DefilementObjetDecalageX < Fenetres(NoFenetre).DefilementNombreLigne Then
            Fenetres(NoFenetre).DefilementObjetDecalageX = Fenetres(NoFenetre).DefilementObjetDecalageX + Valeur
        End If
    Else
        'If Fenetres(NoFenetre).DefilementDecalageX < Fenetres(NoFenetre).DefilementNombreLigne Then
        If NoFenetre = 3 Or NoFenetre = 5 Then
            Fenetres(NoFenetre).DefilementDecalageX = Fenetres(NoFenetre).DefilementDecalageX + Valeur
            If Fenetres(NoFenetre).DefilementDecalageX > Fenetres(NoFenetre).DefilementNombreLigne Then Fenetres(NoFenetre).DefilementDecalageX = Fenetres(NoFenetre).DefilementNombreLigne
        Else
            Fenetres(0).DefilementDecalageX = Fenetres(NoFenetre).DefilementDecalageX + Valeur
            If Fenetres(0).DefilementDecalageX > Fenetres(NoFenetre).DefilementNombreLigne Then Fenetres(0).DefilementDecalageX = Fenetres(NoFenetre).DefilementNombreLigne
            Fenetres(1).DefilementDecalageX = Fenetres(0).DefilementDecalageX
            If Fenetres(1).DefilementDecalageX > Fenetres(NoFenetre).DefilementNombreLigne Then Fenetres(1).DefilementDecalageX = Fenetres(NoFenetre).DefilementNombreLigne
        End If
        'End If
    End If
End Sub

Public Property Get Affichage_PleinEcran() As Boolean
    Affichage_PleinEcran = AffichagePleinEcran
End Property

Public Property Let Affichage_PleinEcran(ByVal Valeur As Boolean)
    AffichagePleinEcran = Valeur
    If Valeur Then
        AfficherCadre = False
        AfficherEpoques = False
        'AfficherFPS = False
        AfficherAction = False
        AfficherBarresStatut = False
        AfficherBoutonsConstruire = False
        AfficherBoutonsInfos = False
        AfficherBoutonsDeplacement = False
        AfficherStastistiques = False
        AfficherRaccourciObjets = False
        AfficherEffetsTemporaires = False
    Else
        AfficherCadre = TempAfficherCadre
        AfficherEpoques = TempAfficherEpoques
        'AfficherFPS = TempAfficherFPS
        AfficherAction = TempAfficherAction
        AfficherBarresStatut = TempAfficherBarresStatut
        AfficherBoutonsConstruire = TempAfficherBoutonsConstruire
        AfficherBoutonsInfos = TempAfficherBoutonsInfos
        AfficherBoutonsDeplacement = TempAfficherBoutonsDeplacement
        AfficherStastistiques = TempAfficherStastistiques
        AfficherRaccourciObjets = TempAfficherRaccourciObjets
        AfficherEffetsTemporaires = TempAfficherEffetsTemp
    End If
End Property

Private Function Interieur_Bouton(ByVal X1 As Long, _
                                  ByVal Y1 As Long, _
                                  ByVal X2 As Long, _
                                  ByVal Y2 As Long, _
                                  ByVal Largeur As Long, _
                                  ByVal Hauteur As Long) As Boolean
    Interieur_Bouton = X1 >= X2 And _
                       X1 <= X2 + Largeur And _
                       Y1 >= Y2 And _
                       Y1 <= Y2 + Hauteur
End Function

Public Sub Raccourcis_Objets_Equiper(ByVal Indice As Integer, _
                                     ByRef Perso As ClsJeuPerso, _
                                     Optional ByVal Ctrl As Boolean, _
                                     Optional ByVal Shift As Boolean, _
                                     Optional ByVal Alt As Boolean)
    Dim i As Integer
    Dim Temp1, Temp2 As Boolean
    If Alt Then
        Indice = Indice + IntTouches.NombreTouchesObjets
    ElseIf Shift Then
        Indice = Indice + IntTouches.NombreTouchesObjets * 2
    End If
    With Perso
    For i = 0 To .Nombre_Objets_Inventaire - 1
        If .Objet_Inventaire_Actif(i) Then
            If Touches_Raccourci_Objets(Perso.Objet_Inventaire_Type(i)) = Indice Then
                If .Objet_Inventaire_Categorie(i) = 0 Then
                    If Temp1 And .Objet_Equipes_Actif(1) Then
                        '.Objet_Ranger 1, Parametres, Commentaires
                        Cmd_Perso_ObjetRanger Perso.Numero, 1
                    Else
                        Temp1 = True
                    End If
                End If
                If .Objet_Inventaire_Categorie(i) = 6 Then
                    If Temp2 And .Objet_Equipes_Actif(6) Then
                        '.Objet_Ranger 6, Parametres, Commentaires
                        Cmd_Perso_ObjetRanger Perso.Numero, 6
                    Else
                        Temp2 = True
                    End If
                End If
                '.Objet_Equiper i, Parametres, Commentaires
                Cmd_Perso_ObjetEquiper Perso.Numero, i
            End If
        End If
    Next i
    End With
End Sub

Private Sub Trouver_Maison(ByRef Perso As ClsJeuPerso, _
                           ByVal TypeBatiment As Integer, _
                           ByVal Reparer As Boolean)
    'Le personnage se rend à la maison alliée du type indiqué en paramètres.
    Dim i As Long, j As Long
    j = Perso.IndiceMaison + 1
    For i = 0 To UBound(Maisons())
        If j > UBound(Maisons()) Then
            j = 0
        End If
        If Perso.NumeroEquipe = Persos(j).NumeroEquipe And _
           Maisons(j).TypeBatiment = TypeBatiment And _
           Maisons(j).Visible Then
            'Perso.Aller_Maison Maisons(j)
            Cmd_Perso_AllerMaison Perso.Numero, j, Reparer
            i = UBound(Maisons())
        Else
            j = j + 1
        End If
    Next i
End Sub

'Propriétés des fenêtres.
Public Property Get Fenetres_Visible(ByVal Index As Integer) As Boolean
    Fenetres_Visible = Fenetres(Index).Visible
End Property
Public Property Get Fenetre_InventaireMaison_Visible(ByVal Index As Integer) As Boolean
    Fenetre_InventaireMaison_Visible = Fenetres(4).Visible And Fenetres(4).Bascule = 1
End Property
Public Property Let Fenetres_Visible(ByVal Index As Integer, ByVal Valeur As Boolean)
    Fenetres(Index).Visible = Valeur
End Property
Public Property Get Fenetre_Deplace(ByVal Index As Integer) As Boolean
    Fenetre_Deplace = Fenetres(Index).Deplace
End Property
Public Property Let Fenetre_Deplace(ByVal Index As Integer, ByVal Valeur As Boolean)
    Fenetres(Index).Deplace = Valeur
End Property
Public Property Get Fenetre_PositionX(ByVal Index As Integer) As Integer
    Fenetre_PositionX = Fenetres(Index).PositionX
End Property
Public Property Let Fenetre_PositionX(ByVal Index As Integer, ByVal Valeur As Integer)
    Fenetres(Index).PositionX = Valeur
End Property
Public Property Get Fenetre_PositionY(ByVal Index As Integer) As Integer
    Fenetre_PositionY = Fenetres(Index).PositionY
End Property
Public Property Let Fenetre_PositionY(ByVal Index As Integer, ByVal Valeur As Integer)
    Fenetres(Index).PositionY = Valeur
End Property

Public Sub Actionner_Boutons_Infos(ByVal Index As Integer, Optional ByVal NoJoueur As Integer)
    If Index <= UBound(Fenetres()) - 1 Then
        Fenetres(Index).Visible = Not BoutonsInfos(Index).Enabled
        If Index = 6 And Not Fenetres(6).Visible Then
            'Réactive le scrolling si l'on ferme la carte.
            FrmMoteur2D.Definir_ScrollingAuto(NoJoueur) = True
        'ElseIf Index = 3 And Not Fenetres(3).Visible Then
        '    Actualiser_Competences Perso, SaMaison, Parametres
        End If
    ElseIf Index = 7 Then
        Jeu.EnPause = Not Jeu.EnPause
    ElseIf Index = 8 Then
        Fenetres(7).Visible = Not BoutonsInfos(8).Enabled
        NouveauxObjectifs = False
    End If
    BoutonsInfos(Index).Enabled = Not BoutonsInfos(Index).Enabled
End Sub

Public Sub Fermer_Toutes_les_fenetres()
    Dim i As Integer
    For i = 0 To UBound(Fenetres())
        Fenetres(i).Visible = False
        BoutonsInfos(i).Enabled = False
    Next i
    FrmMoteur2D.Definir_ScrollingAuto(0) = True
    ActiverBoutonsRadiaux = False
End Sub

Public Property Get AutoRecuperation_Carac(ByVal Index As Integer) As Boolean
    AutoRecuperation_Carac = AutoRecuperationCarac(Index)
End Property
Public Property Let AutoRecuperation_Carac(ByVal Index As Integer, ByVal Valeur As Boolean)
    AutoRecuperationCarac(Index) = Valeur
End Property
Public Property Get AutoRecuperation_CaracPcent(ByVal Index As Integer) As Single
    AutoRecuperation_CaracPcent = AutoRecuperationCaracPcent(Index)
End Property
Public Property Let AutoRecuperation_CaracPcent(ByVal Index As Integer, ByVal Valeur As Single)
    AutoRecuperationCaracPcent(Index) = Valeur
    Actualiser_BarreStatut Index
End Property

'Propriétés d'accés aux fichiers ini.
Public Property Get Definir_AfficherCadre() As Boolean
    Definir_AfficherCadre = AfficherCadre
End Property
Public Property Let Definir_AfficherCadre(ByVal Valeur As Boolean)
    AfficherCadre = Valeur
    TempAfficherCadre = Valeur
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherCadre") = -CDbl(AfficherCadre)
End Property

Public Property Get Definir_AfficherEpoques() As Boolean
    Definir_AfficherEpoques = AfficherEpoques
End Property
Public Property Let Definir_AfficherEpoques(ByVal Valeur As Boolean)
    AfficherEpoques = Valeur
    TempAfficherEpoques = Valeur
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherEpoques") = -CDbl(AfficherEpoques)
End Property

Public Property Get Definir_AfficherFPS() As Boolean
    Definir_AfficherFPS = AfficherFPS
End Property
Public Property Let Definir_AfficherFPS(ByVal Valeur As Boolean)
    AfficherFPS = Valeur
    TempAfficherFPS = Valeur
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherFPS") = -CDbl(AfficherFPS)
End Property

Public Property Get Definir_AfficherAction() As Boolean
    Definir_AfficherAction = AfficherAction
End Property
Public Property Let Definir_AfficherAction(ByVal Valeur As Boolean)
    AfficherAction = Valeur
    TempAfficherAction = Valeur
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherAction") = -CDbl(AfficherAction)
End Property

Public Property Get Definir_AfficherBarresStatut() As Boolean
    Definir_AfficherBarresStatut = AfficherBarresStatut
End Property
Public Property Let Definir_AfficherBarresStatut(ByVal Valeur As Boolean)
    AfficherBarresStatut = Valeur
    TempAfficherBarresStatut = Valeur
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherBarresStatut") = -CDbl(AfficherBarresStatut)
End Property

Public Property Get Definir_AfficherBoutonsInfos() As Boolean
    Definir_AfficherBoutonsInfos = AfficherBoutonsInfos
End Property
Public Property Let Definir_AfficherBoutonsInfos(ByVal Valeur As Boolean)
    AfficherBoutonsInfos = Valeur
    TempAfficherBoutonsInfos = Valeur
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherBoutonsInfos") = -CDbl(AfficherBoutonsInfos)
End Property

Public Property Get Definir_AfficherBoutonsConstruire() As Boolean
    Definir_AfficherBoutonsConstruire = AfficherBoutonsConstruire
End Property
Public Property Let Definir_AfficherBoutonsConstruire(ByVal Valeur As Boolean)
    AfficherBoutonsConstruire = Valeur
    TempAfficherBoutonsConstruire = Valeur
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherBoutonsConstruire") = -CDbl(AfficherBoutonsConstruire)
End Property

Public Property Get Definir_AfficherBoutonsDeplacement() As Boolean
    Definir_AfficherBoutonsDeplacement = AfficherBoutonsDeplacement
End Property
Public Property Let Definir_AfficherBoutonsDeplacement(ByVal Valeur As Boolean)
    AfficherBoutonsDeplacement = Valeur
    TempAfficherBoutonsDeplacement = Valeur
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherBoutonsDeplacement") = -CDbl(AfficherBoutonsDeplacement)
End Property

Public Property Get Definir_MasquerAutoBoutonsConstruire() As Boolean
    Definir_MasquerAutoBoutonsConstruire = MasquerAutoBoutonsConstruire
End Property
Public Property Let Definir_MasquerAutoBoutonsConstruire(ByVal Valeur As Boolean)
    MasquerAutoBoutonsConstruire = Valeur
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("MasquerAutoBoutonsConstruire") = -CDbl(MasquerAutoBoutonsConstruire)
End Property

Public Property Get Definir_MasquerAutoBoutonsDeplacement() As Boolean
    Definir_MasquerAutoBoutonsDeplacement = MasquerAutoBoutonsDeplacement
End Property
Public Property Let Definir_MasquerAutoBoutonsDeplacement(ByVal Valeur As Boolean)
    MasquerAutoBoutonsDeplacement = Valeur
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("MasquerAutoBoutonsDeplacement") = -CDbl(MasquerAutoBoutonsDeplacement)
End Property

Public Property Get Definir_AfficherMessageModeAuto() As Boolean
    Definir_AfficherMessageModeAuto = AfficherMessageModeAuto
End Property
Public Property Let Definir_AfficherMessageModeAuto(ByVal Valeur As Boolean)
    AfficherMessageModeAuto = Valeur
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    If AfficherMessageModeAuto Then
        FicIni.Parametre("AfficherMessageModeAuto") = 1
    Else
        FicIni.Parametre("AfficherMessageModeAuto") = 0
    End If
End Property

Public Property Get Definir_AfficherInfoBulleTerrain() As Boolean
    Definir_AfficherInfoBulleTerrain = AfficherInfoBulleTerrain
End Property
Public Property Let Definir_AfficherInfoBulleTerrain(ByVal Valeur As Boolean)
    AfficherInfoBulleTerrain = Valeur
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    If AfficherInfoBulleTerrain Then
        FicIni.Parametre("AfficherInfoBulleTerrain") = 1
    Else
        FicIni.Parametre("AfficherInfoBulleTerrain") = 0
    End If
End Property

Public Property Get Definir_AfficherStastistiques() As Boolean
    Definir_AfficherStastistiques = AfficherStastistiques
End Property
Public Property Let Definir_AfficherStastistiques(ByVal Valeur As Boolean)
    AfficherStastistiques = Valeur
    TempAfficherStastistiques = Valeur
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherStastistiques") = -CDbl(AfficherStastistiques)
End Property

Public Property Get Definir_AfficherRaccourciObjets() As Boolean
    Definir_AfficherRaccourciObjets = AfficherRaccourciObjets
End Property
Public Property Let Definir_AfficherRaccourciObjets(ByVal Valeur As Boolean)
    AfficherRaccourciObjets = Valeur
    TempAfficherRaccourciObjets = Valeur
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherRaccourciObjets") = -CDbl(AfficherRaccourciObjets)
End Property

Public Property Get Touches_Raccourci_Objets(ByVal Index As Integer) As Integer
    Touches_Raccourci_Objets = TouchesRaccourciObjets(Index)
End Property
Public Property Let Touches_Raccourci_Objets(ByVal Index As Integer, ByVal Valeur As Integer)
    TouchesRaccourciObjets(Index) = Valeur
End Property

Public Property Get Definir_AfficherEffetsTemporaires() As Boolean
    Definir_AfficherEffetsTemporaires = AfficherEffetsTemporaires
End Property
Public Property Let Definir_AfficherEffetsTemporaires(ByVal Valeur As Boolean)
    AfficherEffetsTemporaires = Valeur
    TempAfficherEffetsTemp = Valeur
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherEffetsTemporaires") = -CDbl(AfficherEffetsTemporaires)
End Property

Public Sub Objectifs_Initialiser()
    Collection_Vider Objectifs
    AfficherObjectifs = False
End Sub
Public Sub Objectifs_Ajouter(ByVal Message As String)
    Dim i As Integer
    Dim Ajout As Boolean
    Ajout = True
    If Objectifs.Count > 0 Then
        For i = 1 To Objectifs.Count
            If Objectifs(i) = Message Then
                Ajout = False
                i = Objectifs.Count
            End If
        Next i
    End If
    If Ajout Then
        Objectifs.Add Message
        If Not Fenetres(6).Visible Then NouveauxObjectifs = True
        AfficherObjectifs = True
        If Parametres.EffetNouvelObjectif > -1 Then
            Activer_Effet_Visuel Parametres.EffetNouvelObjectif, BoutonsInfos(8).PositionX + AffBoutonsInfos.Largeur / 2, BoutonsInfos(8).PositionY + AffBoutonsInfos.Hauteur / 2
        End If
    End If
End Sub

'Propriétés d'accès aux boutons.
Public Property Get Definir_BoutonConstruirePositionX(ByVal Indice As Integer) As Integer
    Definir_BoutonConstruirePositionX = BoutonsConstruire(Indice).PositionX
End Property
Public Property Get Definir_BoutonConstruirePositionY(ByVal Indice As Integer) As Integer
    Definir_BoutonConstruirePositionY = BoutonsConstruire(Indice).PositionY
End Property
Public Property Get Definir_BoutonInfoEnabled(ByVal Indice As Integer) As Boolean
    Definir_BoutonInfoEnabled = BoutonsInfos(Indice).Enabled
End Property
Public Property Let Definir_BoutonInfoEnabled(ByVal Indice As Integer, ByVal Valeur As Boolean)
    BoutonsInfos(Indice).Enabled = Valeur
End Property
Public Property Get Definir_BoutonInfoPositionX(ByVal Indice As Integer) As Integer
    Definir_BoutonInfoPositionX = BoutonsInfos(Indice).PositionX
End Property
Public Property Get Definir_BoutonInfoPositionY(ByVal Indice As Integer) As Integer
    Definir_BoutonInfoPositionY = BoutonsInfos(Indice).PositionY
End Property
Public Property Get Definir_BoutonMangerVisible(ByVal Indice As Integer) As Boolean
    Definir_BoutonMangerVisible = Fenetres(1).BoutonMangerVisible(Indice)
End Property
Public Property Get Definir_BoutonMangerPositionX(ByVal Indice As Integer) As Integer
    Definir_BoutonMangerPositionX = Fenetres(1).BoutonMangerPositionX(Indice)
End Property
Public Property Get Definir_BoutonMangerPositionY(ByVal Indice As Integer) As Integer
    Definir_BoutonMangerPositionY = Fenetres(1).BoutonMangerPositionY(Indice)
End Property
Public Property Get Definir_BoutonTransfertStockMaisonPositionX(ByVal Indice As Integer) As Integer
    Definir_BoutonTransfertStockMaisonPositionX = Fenetres(0).BoutonTransfertStockMaisonPositionX(Indice)
End Property
Public Property Get Definir_BoutonTransfertStockMaisonPositionY(ByVal Indice As Integer) As Integer
    Definir_BoutonTransfertStockMaisonPositionY = Fenetres(0).BoutonTransfertStockMaisonPositionY(Indice)
End Property
Public Property Get Definir_DefilementObjetDecalageX() As Integer
    Definir_DefilementObjetDecalageX = Fenetres(0).DefilementObjetDecalageX
End Property
Public Property Get Definir_BoutonFabriquerPositionX(ByVal Indice As Integer) As Integer
    Definir_BoutonFabriquerPositionX = Fenetres(0).BoutonsFabriquerPositionX(Indice)
End Property
Public Property Get Definir_BoutonFabriquerPositionY(ByVal Indice As Integer) As Integer
    Definir_BoutonFabriquerPositionY = Fenetres(0).BoutonsFabriquerPositionY(Indice)
End Property
Public Property Get Definir_BoutonsObjetsPositionX(ByVal Indice As Integer) As Integer
    Definir_BoutonsObjetsPositionX = Fenetres(0).BoutonsObjetsPositionX(Indice)
End Property
Public Property Get Definir_BoutonsObjetsPositionY(ByVal Indice As Integer) As Integer
    Definir_BoutonsObjetsPositionY = Fenetres(0).BoutonsObjetsPositionY(Indice)
End Property
Public Property Get Definir_BoutonsDefilementPositionX(ByVal Fenetre As Integer, ByVal Indice As Integer) As Integer
    Definir_BoutonsDefilementPositionX = Fenetres(Fenetre).BoutonsDefilementPositionX(Indice)
End Property
Public Property Get Definir_BoutonsDefilementPositionY(ByVal Fenetre As Integer, ByVal Indice As Integer) As Integer
    Definir_BoutonsDefilementPositionY = Fenetres(Fenetre).BoutonsDefilementPositionY(Indice)
End Property
Public Property Get Definir_InventaireBoutonsObjetsPositionX(ByVal Indice As Integer)
    Definir_InventaireBoutonsObjetsPositionX = Fenetres(4).BoutonsObjetsPositionX(Indice)
End Property
Public Property Get Definir_InventaireBoutonsObjetsPositionY(ByVal Indice As Integer)
    Definir_InventaireBoutonsObjetsPositionY = Fenetres(4).BoutonsObjetsPositionY(Indice)
End Property

Public Property Get Definir_MessagesActions(ByVal Index As Integer) As String
    Definir_MessagesActions = MessagesActions(Index)
End Property
