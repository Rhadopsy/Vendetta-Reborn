VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmInfoRaces 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Races"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "FrmInfoRaces.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   7095
   Begin VB.ListBox LstCapacitesSpeciales 
      Height          =   1230
      Left            =   3600
      TabIndex        =   30
      Top             =   7680
      Width           =   3375
   End
   Begin VB.Frame FrameApercu 
      Caption         =   "Aperçu :"
      Height          =   4935
      Left            =   120
      TabIndex        =   27
      Top             =   1200
      Width           =   3375
      Begin VB.Label LblApercuIndisponible 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Image ImgPortrait 
         Height          =   4575
         Left            =   120
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.ListBox LstParticularite 
      Height          =   1230
      Left            =   3600
      TabIndex        =   14
      Top             =   6120
      Width           =   3375
   End
   Begin VB.ListBox LstMalus 
      Height          =   1815
      Left            =   3600
      TabIndex        =   12
      Top             =   3960
      Width           =   3375
   End
   Begin VB.ListBox LstBonus 
      Height          =   1815
      Left            =   3600
      TabIndex        =   9
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox TxtDescription 
      Height          =   1335
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "FrmInfoRaces.frx":0742
      Top             =   120
      Width           =   3375
   End
   Begin VB.Frame FrameCaracs 
      Caption         =   "Progression des caractéristiques"
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   6240
      Width           =   3375
      Begin ComctlLib.ProgressBar ProgressBarCarac 
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   16
         Top             =   360
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
         Max             =   200
      End
      Begin ComctlLib.ProgressBar ProgressBarCarac 
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   17
         Top             =   720
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
         Max             =   200
      End
      Begin ComctlLib.ProgressBar ProgressBarCarac 
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   18
         Top             =   1080
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
         Max             =   200
      End
      Begin ComctlLib.ProgressBar ProgressBarCarac 
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   19
         Top             =   1440
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
         Max             =   200
      End
      Begin ComctlLib.ProgressBar ProgressBarCarac 
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   20
         Top             =   1920
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
         Max             =   200
      End
      Begin ComctlLib.ProgressBar ProgressBarCarac 
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   21
         Top             =   2280
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
         Max             =   200
      End
      Begin VB.Label LblCarac 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label10"
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   26
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label LblCarac 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label10"
         Height          =   255
         Index           =   4
         Left            =   2640
         TabIndex        =   25
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label LblCarac 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label10"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   24
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label LblCarac 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label10"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   23
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label LblCarac 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label10"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   22
         Top             =   720
         Width           =   495
      End
      Begin VB.Label LblCarac 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label10"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.Label LblDefense 
         Caption         =   "Défense :"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label LblAttaque 
         Caption         =   "Attaque :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label LblMoral 
         Caption         =   "Moral :"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label LblMagie 
         Caption         =   "Magie :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label LblEnergie 
         Caption         =   "Energie :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label LblVie 
         Caption         =   "Vie :"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.ListBox LstRaces 
      Columns         =   4
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label LblCapacitesSpeciales 
      Caption         =   "Capacités spéciales :"
      Height          =   255
      Left            =   3600
      TabIndex        =   29
      Top             =   7440
      Width           =   3375
   End
   Begin VB.Label LblParticularites 
      Caption         =   "Particularité(s) :"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   5880
      Width           =   3375
   End
   Begin VB.Label LblMalus 
      Caption         =   "Malus :"
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   3720
      Width           =   3375
   End
   Begin VB.Label LblBonus 
      Caption         =   "Bonus :"
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   1560
      Width           =   3375
   End
End
Attribute VB_Name = "FrmInfoRaces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Const FichierINI = "Fenetre_InfoRaces"
Const SectionINI = "Fenetre_InfoRaces"

Private Bonus  As String
Private Malus As String
Private PCent As String
Private MultipRegeneration As Integer
Private FormatRegeneration As String

Private Sub Form_Load()
    Dim i As Integer
    
    'Charge les paramètres.
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    
    Me.Caption = FicIni.Parametre("EtiquetteNomFenetre")
    
    FrameApercu = FicIni.Parametre("CadreApercu")
    LblApercuIndisponible.Caption = FicIni.Parametre("EtiquetteApercuIndisponible")
    
    FrameCaracs = FicIni.Parametre("CadreCaracs")
    LblBonus = FicIni.Parametre("EtiquetteBonus")
    LblMalus = FicIni.Parametre("EtiquetteMalus")
    LblParticularites = FicIni.Parametre("EtiquetteParticularites")
    LblCapacitesSpeciales = FicIni.Parametre("EtiquetteCapacitesSpeciales")
    
    LblVie = Parametres.EtiquetteVie & " :"
    LblEnergie = Parametres.EtiquetteEnergie & " :"
    LblMagie = Parametres.EtiquetteMagie & " :"
    LblMoral = Parametres.EtiquetteMoral & " :"
    LblAttaque = Parametres.EtiquetteAttaque & " :"
    LblDefense = Parametres.EtiquetteDefense & " :"
    
    Bonus = FicIni.Parametre("Bonus")
    Malus = FicIni.Parametre("Malus")
    PCent = FicIni.Parametre("PCent")
    MultipRegeneration = Val(FicIni.Parametre("MultipRegeneration"))
    FormatRegeneration = FicIni.Parametre("FormatRegeneration")
    
    'Centre la fenêtre.
    'Move (Screen.Width - Me.Width) / 2, 0
    Move (MDIFrmMain.Width - Me.Width) / 3, 0
    
    For i = 0 To Parametres.NombreRaces - 1
        If Parametres.JouerToutesLesRaces Or Parametres.Race_Jouable(i) Then
            LstRaces.AddItem Parametres.Race_Nom(i)
        End If
    Next i
    If FrmParam.TabStripPersonnage.Tabs(1).Selected Then
        LstRaces.ListIndex = FrmParam.LstRaces.ListIndex
    End If
    If LstRaces.ListIndex = -1 Then LstRaces.ListIndex = 0
End Sub

Private Sub LstRaces_Click()
    On Error Resume Next
    Dim i As Integer, j As Integer, k As Integer
    Dim Temp As String
    'i = LstRaces.ListIndex
    i = Parametres.Race_Indice(LstRaces.List(LstRaces.ListIndex))
    FrmParam.LstRaces.ListIndex = LstRaces.ListIndex
    With Parametres
    Charger_Apercu i
    
    TxtDescription = .Race_Description(i)
    
    'Caractéristiques.
    If .Race_CoefVie(i) * 100 > ProgressBarCarac(0).Max Then
        ProgressBarCarac(0) = ProgressBarCarac(0).Max
    Else
        ProgressBarCarac(0) = .Race_CoefVie(i) * 100
    End If
    If .Race_CoefEnergie(i) * 100 > ProgressBarCarac(1).Max Then
        ProgressBarCarac(1) = ProgressBarCarac(1).Max
    Else
        ProgressBarCarac(1) = .Race_CoefEnergie(i) * 100
    End If
    If .Race_CoefMagie(i) * 100 > ProgressBarCarac(2).Max Then
        ProgressBarCarac(2) = ProgressBarCarac(2).Max
    Else
        ProgressBarCarac(2) = .Race_CoefMagie(i) * 100
    End If
    If .Race_CoefMoral(i) * 100 > ProgressBarCarac(3).Max Then
        ProgressBarCarac(3) = ProgressBarCarac(3).Max
    Else
        ProgressBarCarac(3) = .Race_CoefMoral(i) * 100
    End If
    If .Race_CoefAttaque(i) * 100 > ProgressBarCarac(4).Max Then
        ProgressBarCarac(4) = ProgressBarCarac(4).Max
    Else
        ProgressBarCarac(4) = .Race_CoefAttaque(i) * 100
    End If
    If .Race_CoefDefense(i) * 100 > ProgressBarCarac(5).Max Then
        ProgressBarCarac(5) = ProgressBarCarac(5).Max
    Else
        ProgressBarCarac(5) = .Race_CoefDefense(i) * 100
    End If
    LblCarac(0) = .Race_CoefVie(i) * 100 & PCent
    LblCarac(1) = .Race_CoefEnergie(i) * 100 & PCent
    LblCarac(2) = .Race_CoefMagie(i) * 100 & PCent
    LblCarac(3) = .Race_CoefMoral(i) * 100 & PCent
    LblCarac(4) = .Race_CoefAttaque(i) * 100 & PCent
    LblCarac(5) = .Race_CoefDefense(i) * 100 & PCent
    
    LstBonus.Clear
    LstMalus.Clear
    LstParticularite.Clear
    LstCapacitesSpeciales.Clear
    'Compétences.
    For j = 0 To .NombreCompetencesSpeciales - 1
        If .Race_CoefCompSpeciales(i, j) > 1 Then
            'Bonus
            LstBonus.AddItem Bonus & (.Race_CoefCompSpeciales(i, j) - 1) * 100 & PCent & .Speciales_NomCompetences(j)
        ElseIf .Race_CoefCompSpeciales(i, j) < 1 Then
            'Malus.
            LstMalus.AddItem Malus & (1 - .Race_CoefCompSpeciales(i, j)) * 100 & PCent & .Speciales_NomCompetences(j)
        End If
    Next j
    For j = 0 To .NombreRessources - 1
        If .Race_CoefCompRessources(i, j) > 1 Then
            'Bonus
            LstBonus.AddItem Bonus & (.Race_CoefCompRessources(i, j) - 1) * 100 & PCent & .Ressources_NomCompetence(j)
        ElseIf .Race_CoefCompRessources(i, j) < 1 Then
            'Malus.
            LstMalus.AddItem Malus & (1 - .Race_CoefCompRessources(i, j)) * 100 & PCent & .Ressources_NomCompetence(j)
        End If
    Next j
    For j = 0 To .NombreServices - 2
        If .Race_CoefCompServices(i, j) > 1 Then
            'Bonus
            LstBonus.AddItem Bonus & (.Race_CoefCompServices(i, j) - 1) * 100 & PCent & .Service_NomCompetence(j + 1)
        ElseIf .Race_CoefCompServices(i, j) < 1 Then
            'Malus.
            LstMalus.AddItem Malus & (1 - .Race_CoefCompServices(i, j)) * 100 & PCent & .Service_NomCompetence(j + 1)
        End If
    Next j
    For k = 0 To .NombreCompetencesObjets - 1
        j = Parametres.CompetenceObjet_NoListe(k)
        If .Race_CoefCompObjets(i, j) > 1 Then
            'Bonus
            LstBonus.AddItem Bonus & (.Race_CoefCompObjets(i, j) - 1) * 100 & PCent & .CompetenceObjet_Nom(j)
        ElseIf .Race_CoefCompObjets(i, j) < 1 Then
            'Malus.
            LstMalus.AddItem Malus & (1 - .Race_CoefCompObjets(i, j)) * 100 & PCent & .CompetenceObjet_Nom(j)
        End If
    Next k
    
    'Particularités
    If Not .Race_UtilisationMonture(i) Then
        LstParticularite.AddItem .EtiquetteUtilisationMontureImpossible
    End If
    If .Race_AttaqueResistance(i) > 0 Then
        LstParticularite.AddItem .Resistance_Attaque(.Race_AttaqueResistance(i))
    End If
    
    If .Race_VitesseRegenerationVie(i) > 0 Then _
        LstParticularite.AddItem .EtiquetteRegenerationVie & " : " & Format(.Race_VitesseRegenerationVie(i) * MultipRegeneration, FormatRegeneration) & PCent
    If .Race_VitesseRegenerationEnergie(i) > 0 And Not .Race_Infatigable(i) Then _
        LstParticularite.AddItem .EtiquetteRegenerationEnergie & " : " & Format(.Race_VitesseRegenerationEnergie(i) * MultipRegeneration, FormatRegeneration) & PCent
    If .Race_VitesseRegenerationMagie(i) > 0 Then _
        LstParticularite.AddItem .EtiquetteRegenerationMagie & " : " & Format(.Race_VitesseRegenerationMagie(i) * MultipRegeneration, FormatRegeneration) & PCent
    If .Race_VitesseRegenerationMoral(i) > 0 And Not .Race_Infatigable(i) Then _
        LstParticularite.AddItem .EtiquetteRegenerationMoral & " : " & Format(.Race_VitesseRegenerationMoral(i) * MultipRegeneration, FormatRegeneration) & PCent
    If .Race_Volant(i) Then _
        LstParticularite.AddItem .EtiquetteVolant
    If Not .Race_Parole(i) Then _
        LstParticularite.AddItem .EtiquetteMuet
    If Not .Race_Experience(i) Then _
        LstParticularite.AddItem .EtiquettePasdExperience
    If .Race_CommentairesSpecifique(i) > 0 Then _
        LstParticularite.AddItem .EtiquetteCommentairesSpecifique
    If .Race_BonusMaxRessources(i) > 0 Then _
        LstParticularite.AddItem .EtiquetteTranport & " " & .Race_BonusMaxRessources(i) & " " & .EtiquetteRessourcesSupplementaires
    If .Race_PorteeAttaque(i) <> .PersosPorteeAttaque Then _
        LstParticularite.AddItem Bonus & .Race_PorteeAttaque(i) & " " & .EtiquettePorteeAttaque
    If .Race_EffetAttaque(i) <> 0 Then _
        LstParticularite.AddItem .EtiquetteEffetAttaque & " : " & .Effet_Nom(.Race_EffetAttaque(i))
    If .Race_Armure(i) <> 0 Then _
        LstParticularite.AddItem Bonus & .Race_Armure(i) & " " & .EtiquetteArmure
    If .Race_Vitesse(i) > .PersosVitesse Then _
        LstParticularite.AddItem Bonus & .Race_Vitesse(i) - .PersosVitesse & " " & .EtiquetteVitesse
    If .Race_Vitesse(i) < .PersosVitesse Then _
        LstParticularite.AddItem .Race_Vitesse(i) - .PersosVitesse & " " & .EtiquetteVitesse
    If .Race_Infatigable(i) Then _
        LstParticularite.AddItem .EtiquetteInfatiguable
    If .Race_RestaurationComplete(i) Then _
        LstParticularite.AddItem .EtiquetteRestaurationComplete
    If .Race_Kamikaze(i) Then _
        LstParticularite.AddItem .EtiquetteKamikaze
    'Ressources.
    For j = 0 To .NombreRessources - 1
        If .Race_Ressources(i, j) > 0 Then
            If Temp <> "" Then
                Temp = Temp & ", "
            End If
            Temp = Temp & .Race_Ressources(i, j) & " " & LCase(.Ressources_Nom(j))
        End If
    Next j
    'Les résistances.
    For j = 0 To .NombreResistances - 1
        If .Race_Resistances(i, j) <> .ResistanceDefaut Then
            If Temp <> "" Then
                LstParticularite.AddItem Temp
            End If
            If .Race_Resistances(i, j) = Parametres.ResistanceMinimum Then
                LstParticularite.AddItem .Resistance_NomComplet(j) & " : " & UCase(Parametres.EtiquetteMaximumAbr)
            ElseIf .Race_Resistances(i, j) = Parametres.ResistanceMaximum Then
                LstParticularite.AddItem .Resistance_NomComplet(j) & " : " & UCase(Parametres.EtiquetteMinimumAbr)
            Else
                LstParticularite.AddItem .Resistance_NomComplet(j) & " : " & IIf(.Race_Resistances(i, j) < .ResistanceDefaut, "+", "") & Format(-(.Race_Resistances(i, j) - 1) * 100, "0") & .EtiquettePourcent
            End If
        End If
    Next j
    
    'Objets spécifiques.
    Temp = ""
    For j = 0 To .Race_NombreEquipement(i) - 1
        If .Objet_PrixTalents(.Race_Equipement(i, j)) = 0 Then
            If Temp = "" Then
                Temp = Parametres.EtiquetteEquipement & " : "
            Else
                Temp = Temp & ", "
            End If
            Temp = Temp & LCase(.Objet_Nom(.Race_Equipement(i, j)))
        End If
    Next j
    If Temp <> "" Then
        LstParticularite.AddItem Temp
    End If
    For j = 1 To .Race_NombreObjetsRace(i)
        'If .Objet_PrixTalents(.Race_Objet(i, j)) = 0 Then
        LstCapacitesSpeciales.AddItem .Objet_Nom(.Race_ObjetRace(i, j))
        'End If
    Next j
    End With
    
End Sub

Private Sub Charger_Apercu(ByVal IndiceRace As Integer)
    On Error GoTo Erreur
    ImgPortrait.Picture = LoadPicture(App.Path & "\" & Parametres.Race_Portrait(IndiceRace))
    ImgPortrait.Visible = True
    LblApercuIndisponible.Visible = False
    ImgPortrait.Left = (FrameApercu.Width - ImgPortrait.Width) / 2
    ImgPortrait.Top = (FrameApercu.Height - ImgPortrait.Height) / 2
    Exit Sub
Erreur:
    ImgPortrait.Visible = False
    LblApercuIndisponible.Visible = True
End Sub
