VERSION 5.00
Begin VB.Form FrmChargement 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8475
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   565
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   628
   Begin VB.OptionButton OptionSexe 
      BackColor       =   &H00000000&
      Caption         =   "&Masculin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   0
      Left            =   3240
      TabIndex        =   22
      Top             =   4440
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.OptionButton OptionSexe 
      BackColor       =   &H00000000&
      Caption         =   "&Féminin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   21
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton CmdEtape 
      Caption         =   "Terminer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   6720
      TabIndex        =   20
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox LstRaces 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   2760
      TabIndex        =   17
      Top             =   4200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TxtNom 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   14
      Text            =   "Entrez un nom"
      Top             =   4560
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton CmdEtape 
      Caption         =   "Suite ->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   6720
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CmdEtape 
      Caption         =   "<-Retour"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   12
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CmdChoix 
      Caption         =   "Quitter"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   2760
      TabIndex        =   11
      Top             =   7320
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton CmdChoix 
      Caption         =   "Options"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   2760
      TabIndex        =   10
      Top             =   6600
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton CmdChoix 
      Caption         =   "Charger un scénario"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2760
      TabIndex        =   8
      Top             =   5160
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton CmdChoix 
      Caption         =   "Charger partie"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   2760
      TabIndex        =   9
      Top             =   5880
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton CmdChoix 
      Caption         =   "Champs de bataille"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   2760
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton CmdChoix 
      Caption         =   "Monde personnalisé"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2760
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton CmdChoix 
      Caption         =   "Initiation"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2760
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton CmdChoix 
      Caption         =   "Créer un personnage"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2760
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label LblChoixNom 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Entrez un nom :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2760
      TabIndex        =   25
      Top             =   3720
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label LblErreurNom 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2640
      TabIndex        =   24
      Top             =   5040
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label LblSexeDescription 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Ce choix est imposé par votre race"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2760
      TabIndex        =   23
      Top             =   5160
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label LblChoixSexe 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Choisissez son sexe :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   19
      Top             =   3720
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label LblDescription 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   4560
      TabIndex        =   18
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      Height          =   3135
      Left            =   4440
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Image ImgPortrait 
      Height          =   3015
      Left            =   4560
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label LblChoixRace 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Choisissez sa race :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      Top             =   3720
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label LblNom 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2760
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label LblFontCaracs 
      Caption         =   "Caracs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LblFontCritique 
      Caption         =   "Critique"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LblFontDegats 
      Caption         =   "Degats"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LblFontNormale 
      Caption         =   "Normale"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image ImgChargement 
      Height          =   8415
      Left            =   0
      Top             =   0
      Width           =   9375
   End
   Begin VB.Menu MenuTest 
      Caption         =   "Test"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "FrmChargement"
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

Const FichierFenetreINI = "Fenetre_Menu"
Const SectionFenetreINI = "Fenetre_Menu"

Private EtapeCreationPersonnage As Integer
Private NomChange As Boolean 'Passe à crai dès que l'on a changé le nom du personnage.

Public Afficher As Boolean

Private Sub CmdChoix_Click(Index As Integer)
    Select Case Index
    Case 0:
        Creer_Personnage True
    'Choisi un mode de jeu.
    Case Is < 6:
        Me.Hide
        FrmParam.TabStripMonde.Tabs(Index).Selected = True
        FrmParam.TabStripMonde2.Tabs(Index).Selected = True
        FrmParam.Show
        FrmChargement.Afficher = False
    'Options.
    Case 6:
        frmOptions.Show
        Me.Hide
    'Quitter.
    Case 7:
        If EtapeCreationPersonnage = 0 Then
            Unload Me
            MDIFrmMain.MenuPartieQuitter_Click
        Else
            Creer_Personnage False
        End If
    End Select
End Sub

Private Sub CmdEtape_Click(Index As Integer)
    Dim i As Integer
    Select Case Index
    Case 0:
        EtapeCreationPersonnage = EtapeCreationPersonnage - 1
    Case 1:
        EtapeCreationPersonnage = EtapeCreationPersonnage + 1
    Case 2:
        EtapeCreationPersonnage = 4
    End Select
    
    CmdChoix(0).Enabled = Not EtapeCreationPersonnage > 0
    For i = 1 To 6
        CmdChoix(i).Visible = Not EtapeCreationPersonnage > 0
    Next i
    
    'Saisie du nom.
    LblChoixNom.Visible = EtapeCreationPersonnage = 1
    TxtNom.Visible = EtapeCreationPersonnage = 1
    LblErreurNom.Visible = EtapeCreationPersonnage = 1
    'Choix de la race.
    LblNom.Visible = EtapeCreationPersonnage = 2 Or EtapeCreationPersonnage = 3
    
    LblChoixRace.Visible = EtapeCreationPersonnage = 2
    LstRaces.Visible = EtapeCreationPersonnage = 2
    ImgPortrait.Visible = EtapeCreationPersonnage = 2
    LblDescription.Visible = EtapeCreationPersonnage = 2
    LblChoixSexe.Visible = EtapeCreationPersonnage = 3
    OptionSexe(0).Visible = EtapeCreationPersonnage = 3
    OptionSexe(1).Visible = EtapeCreationPersonnage = 3
    LblSexeDescription.Visible = False
    
    CmdEtape(0).Visible = EtapeCreationPersonnage > 0
    CmdEtape(1).Visible = EtapeCreationPersonnage > 0 And EtapeCreationPersonnage < 3
    CmdEtape(2).Visible = EtapeCreationPersonnage = 3
    
    'Choix du sexe.
    Select Case EtapeCreationPersonnage
    Case 0: 'On quitte la création du personnage.
        CmdEtape(0).Visible = False
        CmdEtape(1).Visible = False
    Case 1: 'Choix du nom.
        TxtNom.SelStart = 0
        TxtNom.SelLength = Len(TxtNom)
        TxtNom.SetFocus
        If NomChange Then
            TxtNom_Change
        Else
            TxtNom_DblClick
        End If
        'CmdEtape(1).Enabled = False
    Case 2: 'Choix de la race
        LblNom = TxtNom
        If LstRaces.ListCount = 0 Then
            For i = 0 To FrmParam.LstRaces.ListCount - 1
                LstRaces.AddItem FrmParam.LstRaces.List(i)
            Next
            LstRaces.ListIndex = 0
        End If
        LblDescription_Click
    Case 3: 'Choix du sexe.
        OptionSexe(0).Enabled = FrmParam.OptionSexe(0).Enabled
        OptionSexe(0).Value = FrmParam.OptionSexe(0).Value
        OptionSexe(1).Enabled = FrmParam.OptionSexe(1).Enabled
        OptionSexe(1).Value = FrmParam.OptionSexe(1).Value
        LblSexeDescription.Visible = FrmParam.LblRestictionSexe.Visible
    Case 4: 'Le personnage est validé.
        FrmParam.TxtNomPerso = TxtNom
        FrmParam.OptionSexe(0).Value = OptionSexe(0).Value
        FrmParam.OptionSexe(1).Value = OptionSexe(1).Value
        FrmParam.CmdCreerPerso_Click
        EtapeCreationPersonnage = 1
        CmdEtape_Click 0
        Rafraichir_Boutons
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo Erreur
    Dim FicIni2 As ClsFicINI
    Set FicIni2 = New ClsFicINI
    FicIni2.fichier = FicIni2.chemin & Langues.Dossier & FichierFenetreINI
    FicIni2.Section = SectionFenetreINI
    'FrmChargement.Afficher = True
    Afficher = FicIni2.Parametre("Afficher")
    ImgChargement.Picture = LoadPicture(App.Path & "\Images\Chargement.bmp")
    'Me.Width = ImgChargement.Width
    'Me.Height = ImgChargement.Height * 15
    'Centre la feuille.
    Move (MDIFrmMain.Width - Me.Width) / 2, (MDIFrmMain.Height - Me.Height) / 2.5
    'Rafraichir_Boutons
    Exit Sub
Erreur:
    Me.Hide
End Sub

Public Sub Rafraichir_Boutons()
    Dim i As Integer
    Dim Temp As Boolean
    Dim FicIni2 As ClsFicINI
    Set FicIni2 = New ClsFicINI
    FicIni2.fichier = FicIni2.chemin & Langues.Dossier & FichierFenetreINI
    FicIni2.Section = SectionFenetreINI
    
    For i = 0 To 7
        CmdChoix(i).Visible = True
    Next i
    Temp = FrmParam.LstPerso.ListCount > 0
    'Active les boutons pour jouer, si on a un personnage de prêt.
    CmdChoix(0).Enabled = True
    For i = 1 To 5
        CmdChoix(i).Enabled = Temp
    Next i
    CmdChoix(6).Enabled = True
    CmdChoix(7).Enabled = True
    
    CmdChoix(0).Caption = FrmParam.TabStripPersonnage.Tabs(1).Caption
    For i = 1 To 5
        CmdChoix(i).Caption = FrmParam.TabStripMonde.Tabs(i).Caption
    Next i
    CmdChoix(6).Caption = MDIFrmMain.MenuOptions.Caption
    CmdChoix(7).Caption = MDIFrmMain.MenuPartieQuitter.Caption
    
    LblChoixNom.Caption = FicIni2.Parametre("EtiquetteChoixNom")
    LblChoixRace.Caption = FicIni2.Parametre("EtiquetteChoixRace")
    LblChoixSexe.Caption = FicIni2.Parametre("EtiquetteChoixSexe")
    CmdEtape(0).Caption = FicIni2.Parametre("BoutonPrecedent")
    CmdEtape(1).Caption = FicIni2.Parametre("BoutonSuivant")
    CmdEtape(2).Caption = FicIni2.Parametre("BoutonFin")
    'Me.Height = 8175
    'Move (MDIFrmMain.Width - Me.Width) / 2, (MDIFrmMain.Height - Me.Height) / 2.5
    If Not Afficher Then
        Me.Hide
        FrmParam.Show
    End If
End Sub

Private Sub Creer_Personnage(ByVal valeur As Boolean)
    If valeur Then
        CmdEtape_Click 1
    Else
        EtapeCreationPersonnage = 1
        CmdEtape_Click 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FicIni.fichier = FicIni.chemin & Langues.Dossier & FichierFenetreINI
    FicIni.Section = SectionFenetreINI
    FicIni.Parametre("Afficher") = -CDbl(Afficher)
End Sub

Private Sub ImgPortrait_Click()
    LblDescription.Visible = True
    ImgPortrait.Visible = False
End Sub

Private Sub LblDescription_Click()
    LblDescription.Visible = False
    ImgPortrait.Visible = True
End Sub

Private Sub LstRaces_Click()
    Dim i As Integer
    i = Parametres.Race_Indice(LstRaces.List(LstRaces.ListIndex))
    FrmParam.LstRaces.ListIndex = LstRaces.ListIndex
    LblDescription.Caption = Parametres.Race_Description(i)
    On Error GoTo Erreur
    ImgPortrait.Picture = LoadPicture(App.Path & "\" & Parametres.Race_Portrait(i))
    ImgPortrait.Left = (Shape1.Width - ImgPortrait.Width) / 2 + Shape1.Left
    ImgPortrait.Top = (Shape1.Height - ImgPortrait.Height) / 2 + Shape1.Top
    Exit Sub
Erreur:
End Sub

Private Sub TxtNom_Change()
    FrmParam.TxtNomPerso = TxtNom
    'LblErreurNom.Visible = FrmParam.LblErreurNom.Visible
    LblErreurNom = FrmParam.LblErreurNom
    CmdEtape(1).Enabled = LblErreurNom = ""
    NomChange = True
End Sub

Private Sub TxtNom_DblClick()
    TxtNom = Noms.Tirer_Nom_Aleatoire(Rnd < 0.5)
    TxtNom.SelStart = 0
    TxtNom.SelLength = Len(TxtNom)
    TxtNom_Change
End Sub
