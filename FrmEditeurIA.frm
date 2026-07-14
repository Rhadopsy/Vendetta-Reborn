VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmEditeurIA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editeur d'Intelligence Artificielle"
   ClientHeight    =   8895
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11175
   Icon            =   "FrmEditeurIA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameApercu 
      Caption         =   "Aperçu de l'équipement selon les époques :"
      Height          =   1935
      Left            =   2040
      TabIndex        =   36
      Top             =   6840
      Width           =   9015
      Begin VB.Label LblApercu 
         Caption         =   "Label1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   255
      End
      Begin VB.Image ImgAperçu 
         Height          =   495
         Index           =   0
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton CmsSupprimerIA 
      Caption         =   "Supprimer"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton CmdRenommer 
      Caption         =   "Renommer"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Frame FrameEpoque 
      Caption         =   "Epoque"
      Height          =   855
      Left            =   2040
      TabIndex        =   10
      Top             =   1680
      Width           =   9015
      Begin VB.TextBox TxtCommentaire 
         Height          =   495
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Text            =   "FrmEditeurIA.frx":0442
         Top             =   240
         Width           =   5535
      End
      Begin ComctlLib.Slider SliderEpoque 
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin VB.Label LblChanger 
         Caption         =   "Changer :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.Label LblEpoque 
         Caption         =   "---"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame FrameListeActions 
      Caption         =   "Liste des actions de l'époque"
      Height          =   4095
      Left            =   2040
      TabIndex        =   3
      Top             =   2640
      Width           =   2655
      Begin VB.ListBox LstActions 
         Height          =   3180
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton CmdDeplacer 
         Height          =   495
         Index           =   1
         Left            =   2040
         Picture         =   "FrmEditeurIA.frx":0448
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3480
         Width           =   495
      End
      Begin VB.CommandButton CmdSupprimer 
         Caption         =   "Supprimer"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton CmdInserer 
         Caption         =   "Ajouter"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton CmdDeplacer 
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "FrmEditeurIA.frx":088A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3480
         Width           =   495
      End
   End
   Begin VB.Frame FrameEditer 
      Caption         =   "Editer l'action sélectionnée"
      Height          =   4095
      Left            =   4800
      TabIndex        =   2
      Top             =   2640
      Width           =   6255
      Begin VB.Frame FrameQuantite 
         Caption         =   "Quantité"
         Height          =   855
         Left            =   2280
         TabIndex        =   28
         Top             =   3120
         Width           =   3855
         Begin VB.CommandButton CmdModifierQuantite 
            Caption         =   "-"
            Height          =   255
            Index           =   1
            Left            =   3480
            TabIndex        =   48
            Top             =   480
            Width           =   255
         End
         Begin VB.CommandButton CmdModifierQuantite 
            Caption         =   "+"
            Height          =   255
            Index           =   0
            Left            =   3480
            TabIndex        =   47
            Top             =   240
            Width           =   255
         End
         Begin VB.CommandButton CmdQuantite 
            Caption         =   "250"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton CmdQuantite 
            Caption         =   "1000"
            Height          =   255
            Index           =   6
            Left            =   600
            TabIndex        =   44
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton CmdQuantite 
            Caption         =   "2000"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   43
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton CmdQuantite 
            Caption         =   "4000"
            Height          =   255
            Index           =   4
            Left            =   600
            TabIndex        =   42
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton CmdQuantite 
            Caption         =   "20"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   32
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton CmdQuantite 
            Caption         =   "10"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   31
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton CmdQuantite 
            Caption         =   "5"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   30
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton CmdQuantite 
            Caption         =   "1"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   495
         End
         Begin ComctlLib.Slider SliderQuantite 
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   33
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   327682
            Max             =   20
         End
         Begin ComctlLib.Slider SliderQuantite 
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   46
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   327682
            Max             =   4000
         End
         Begin VB.Label LblQuantite 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label4"
            Height          =   255
            Left            =   3000
            TabIndex        =   34
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame FrameDesignation 
         Caption         =   "Désignation"
         Height          =   2655
         Left            =   2280
         TabIndex        =   26
         Top             =   360
         Width           =   3855
         Begin VB.ListBox LstValeur 
            Height          =   2205
            Index           =   1
            ItemData        =   "FrmEditeurIA.frx":0CCC
            Left            =   120
            List            =   "FrmEditeurIA.frx":0CCE
            Sorted          =   -1  'True
            TabIndex        =   40
            Top             =   240
            Width           =   3615
         End
         Begin VB.ListBox LstValeur 
            Height          =   1815
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   3615
         End
      End
      Begin VB.Frame FrameChoixAction 
         Caption         =   "Actions"
         Height          =   3615
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   2055
         Begin VB.OptionButton OptionAction 
            Caption         =   "&Gagner x argent"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   49
            Top             =   600
            Width           =   1815
         End
         Begin VB.OptionButton OptionAction 
            Caption         =   "&Gagner x argent"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton OptionAction 
            Caption         =   "Changer &mode"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   39
            Top             =   3240
            Width           =   1815
         End
         Begin VB.OptionButton OptionAction 
            Caption         =   "Attaquer &donjon"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   22
            Top             =   2880
            Width           =   1815
         End
         Begin VB.OptionButton OptionAction 
            Caption         =   "Attaquer &ennemi"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   21
            Top             =   2640
            Width           =   1815
         End
         Begin VB.OptionButton OptionAction 
            Caption         =   "&Patrouiller"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   20
            Top             =   2400
            Width           =   1815
         End
         Begin VB.OptionButton OptionAction 
            Caption         =   "&Travailler"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   19
            Top             =   2040
            Width           =   1815
         End
         Begin VB.OptionButton OptionAction 
            Caption         =   "&Equiper objet"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   18
            Top             =   1680
            Width           =   1815
         End
         Begin VB.OptionButton OptionAction 
            Caption         =   "Construire &batiment"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   1320
            Width           =   1815
         End
         Begin VB.OptionButton OptionAction 
            Caption         =   "&Collecter ressources"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Value           =   -1  'True
            Width           =   1815
         End
      End
   End
   Begin VB.TextBox TxtDescription 
      Height          =   495
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "FrmEditeurIA.frx":0CD0
      Top             =   1080
      Width           =   9015
   End
   Begin VB.FileListBox FileIAs 
      Height          =   6330
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton CmdDuppliquer 
      Caption         =   "Duppliquer"
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton CmdNouveau 
      Caption         =   "Nouveau"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Label LblAttention 
      Caption         =   $"FrmEditeurIA.frx":0CD6
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   120
      TabIndex        =   38
      Top             =   120
      Width           =   10935
   End
   Begin VB.Label LblCommentaireGeneral 
      Caption         =   "Commentaire général :"
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   840
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11160
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu MenuFichier 
      Caption         =   "&Fichiers IAs"
      Begin VB.Menu MenuFichierNouveau 
         Caption         =   "&Nouveau"
         Shortcut        =   ^N
      End
      Begin VB.Menu MenuFichierDuppliquer 
         Caption         =   "&Duppliquer"
         Shortcut        =   ^U
      End
      Begin VB.Menu MenuFichierRenommer 
         Caption         =   "&Renommer"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MenuFichierEnregistrer 
         Caption         =   "&Enregistrer"
         Shortcut        =   ^S
      End
      Begin VB.Menu MenuFichierSupprimer 
         Caption         =   "&Supprimer"
      End
      Begin VB.Menu MenuFichierSeparateur 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFichierQuitter 
         Caption         =   "&Quitter"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MenuActions 
      Caption         =   "&Actions"
      Begin VB.Menu MenuActionEpoque 
         Caption         =   "&Epoque"
         Begin VB.Menu MenuActionEpoqueChanger 
            Caption         =   "&Précédente"
            Index           =   0
            Shortcut        =   {F3}
         End
         Begin VB.Menu MenuActionEpoqueChanger 
            Caption         =   "&Suivante"
            Index           =   1
            Shortcut        =   {F4}
         End
      End
      Begin VB.Menu MenuActionsSeperateur1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuActionsPrecedent 
         Caption         =   "&Précédent"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MenuActionsSuivant 
         Caption         =   "&Suivant"
         Shortcut        =   {F6}
      End
      Begin VB.Menu MenuActionsSeperateur2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuActionInserer 
         Caption         =   "&Ajouter"
         Shortcut        =   ^A
      End
      Begin VB.Menu MenuActionsSupprimer 
         Caption         =   "S&upprimer"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu MenuAide 
      Caption         =   "&?"
      Begin VB.Menu MenuAideManuel 
         Caption         =   "&Manuel"
         Shortcut        =   ^M
      End
      Begin VB.Menu MenuAideSeparateur 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAideAPropos 
         Caption         =   "&A propos"
         Shortcut        =   ^P
      End
   End
End
Attribute VB_Name = "FrmEditeurIA"
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

Const FichierINI = "Fenetre_EditeurIA"
Const SectionINI = "Fenetre_EditeurIA"

Private NombreMaxImagesParLignes As Integer
Private NombreMaxImagesApercu As Integer

Private CheminImagesObjets As String
Private ExtentionFichiersImages As String

Private CheminFichierInfo As String
Private SuffixeFichierDupplication As String
Private MessageNouveau As String
Private FichierNouveau As String
Private MessageRenommer As String
Private MessageSupprimer1 As String
Private MessageSupprimer2 As String
Private MessageErreur As String
Private MessageErreurChargementImage As String
Private MessageEnregistrementQuitter As String
Private ListeRessources As String
Private ListeBatiments As String
Private ListeObjets As String
Private ListeTravailler As String
Private ListePatrouiller As String
Private ListeAttaquerEnnemi As String
Private ListeAttaquerChateau As String
Const NombreModes = 14
Private ListeModes() As String
Private ListeGagnerArgent As String

Private IndiceIA As Integer
Private Modification As Boolean

Private Sub Form_Load()
    Dim i As Integer
    FrmParam.Hide
    FrmInfoRaces.Hide
    
    'Charge les paramètres.
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    
    NombreMaxImagesParLignes = Val(FicIni.Parametre("NombreMaxImagesParLignes"))
    NombreMaxImagesApercu = Val(FicIni.Parametre("NombreMaxImagesApercu"))
    CheminImagesObjets = FicIni.Parametre("CheminImagesObjets")
    ExtentionFichiersImages = FicIni.Parametre("ExtentionFichiersImages")
    CheminFichierInfo = FicIni.Parametre("CheminFichierInfo")
    SuffixeFichierDupplication = FicIni.Parametre("SuffixeFichierDupplication")
    MessageNouveau = FicIni.Parametre("MessageNouveau")
    FichierNouveau = FicIni.Parametre("FichierNouveau")
    MessageRenommer = FicIni.Parametre("MessageRenommer")
    MessageSupprimer1 = FicIni.Parametre("MessageSupprimer1")
    MessageSupprimer2 = FicIni.Parametre("MessageSupprimer2")
    MessageErreur = FicIni.Parametre("MessageErreur")
    MessageErreurChargementImage = FicIni.Parametre("MessageErreurChargementImage")
    MessageEnregistrementQuitter = FicIni.Parametre("MessageEnregistrementQuitter")
    ListeRessources = FicIni.Parametre("ListeRessources")
    ListeBatiments = FicIni.Parametre("ListeBatiments")
    ListeObjets = FicIni.Parametre("ListeObjets")
    ListeTravailler = FicIni.Parametre("ListeTravailler")
    ListePatrouiller = FicIni.Parametre("ListePatrouiller")
    ListeAttaquerEnnemi = FicIni.Parametre("ListeAttaquerEnnemi")
    ListeAttaquerChateau = FicIni.Parametre("ListeAttaquerChateau")
    ReDim ListeModes(NombreModes - 1)
    For i = 0 To NombreModes - 1
        ListeModes(i) = FicIni.Parametre("ListeMode(" & i + 1 & ")")
    Next i
    ListeGagnerArgent = FicIni.Parametre("ListeGagnerArgent")
    
    Caption = FicIni.Parametre("FenetreTitre")
    MenuFichier.Caption = FicIni.Parametre("MenuFichier")
    MenuFichierNouveau.Caption = FicIni.Parametre("MenuFichierNouveau")
    MenuFichierDuppliquer.Caption = FicIni.Parametre("MenuFichierDuppliquer")
    MenuFichierRenommer.Caption = FicIni.Parametre("MenuFichierRenommer")
    MenuFichierEnregistrer.Caption = FicIni.Parametre("MenuFichierEnregistrer")
    MenuFichierSupprimer.Caption = FicIni.Parametre("MenuFichierSupprimer")
    MenuFichierQuitter.Caption = FicIni.Parametre("MenuFichierQuitter")
    MenuActions.Caption = FicIni.Parametre("MenuAction")
    MenuActionEpoque.Caption = FicIni.Parametre("MenuActionEpoque")
    MenuActionEpoqueChanger(0).Caption = FicIni.Parametre("MenuActionEpoquePrecedent")
    MenuActionEpoqueChanger(1).Caption = FicIni.Parametre("MenuActionEpoqueSuivant")
    MenuActionsPrecedent.Caption = FicIni.Parametre("MenuActionsPrecedent")
    MenuActionsSuivant.Caption = FicIni.Parametre("MenuActionsSuivant")
    MenuActionInserer.Caption = FicIni.Parametre("MenuActionInserer")
    MenuActionsSupprimer.Caption = FicIni.Parametre("MenuActionsSupprimer")
    MenuAide.Caption = FicIni.Parametre("MenuAide")
    MenuAideManuel.Caption = FicIni.Parametre("MenuAideManuel")
    MenuAideAPropos.Caption = FicIni.Parametre("MenuAideAPropos")
    LblAttention = FicIni.Parametre("EtiquetteAttention")
    CmdNouveau.Caption = FicIni.Parametre("BoutonNouveau")
    CmdDuppliquer.Caption = FicIni.Parametre("BoutonDuppliquer")
    CmdRenommer.Caption = FicIni.Parametre("BoutonRenommer")
    CmsSupprimerIA.Caption = FicIni.Parametre("BoutonSupprimer")
    LblCommentaireGeneral = FicIni.Parametre("EtiquetteCommentaireGeneral")
    FrameEpoque = Parametres.EtiquetteEpoque
    LblChanger = FicIni.Parametre("EtiquetteChanger")
    FrameListeActions = FicIni.Parametre("CadreListeActions")
    CmdInserer.Caption = FicIni.Parametre("BoutonInsererAction")
    CmdSupprimer.Caption = FicIni.Parametre("BoutonSupprimerAction")
    FrameEditer = FicIni.Parametre("CadreEditer")
    FrameChoixAction = FicIni.Parametre("CadreChoixAction")
    For i = 0 To 9
        OptionAction(i).Caption = FicIni.Parametre("ChoixAction(" & i + 1 & ")")
    Next i
    FrameDesignation = FicIni.Parametre("CadreDesignation")
    FrameQuantite = FicIni.Parametre("CadreQuantite")
    FrameApercu = FicIni.Parametre("CadreApercu")
    
    'Centre la feuille.
    Move (MDIFrmMain.Width - Me.Width) / 2, _
         (MDIFrmMain.Height - Me.Height) / 2
    
    'Prépare l'aperçu.
    For i = 1 To NombreMaxImagesApercu - 1
        Load ImgAperçu(i)
        ImgAperçu(i).Left = ImgAperçu(i - 1).Left + ImgAperçu(i - 1).Width
        ImgAperçu(i).Top = ImgAperçu(i - 1).Top
        Load LblApercu(i)
        LblApercu(i).Left = LblApercu(i - 1).Left + ImgAperçu(i - 1).Width
        LblApercu(i).Top = LblApercu(i - 1).Top
        If i Mod NombreMaxImagesParLignes = 0 Then 'On passe à la ligne suivante.
            ImgAperçu(i).Left = ImgAperçu(0).Left
            ImgAperçu(i).Top = ImgAperçu(i).Top + ImgAperçu(i).Height
            LblApercu(i).Left = LblApercu(0).Left
            LblApercu(i).Top = LblApercu(i).Top + ImgAperçu(i).Height
        End If
    Next i
    
    FileIAs.Path = App.Path & "\" & CheminIA
    If FileIAs.ListCount > 0 Then
        FileIAs.ListIndex = 0
    End If
    SliderEpoque.Max = Parametres.NombreEpoques - 1
    
    SliderEpoque_Scroll
End Sub

Private Sub CmdDeplacer_Click(Index As Integer)
    Dim Temp As Integer
    Select Case Index
    Case 0:
        IA2s(IndiceIA).Intervertir_Etapes SliderEpoque.Value, LstActions.ListIndex, LstActions.ListIndex - 1
        Temp = LstActions.ListIndex - 1
    Case 1:
        IA2s(IndiceIA).Intervertir_Etapes SliderEpoque.Value, LstActions.ListIndex, LstActions.ListIndex + 1
        Temp = LstActions.ListIndex + 1
    End Select
    Actualiser_Actions
    LstActions.ListIndex = Temp
    Modification = True
End Sub

Private Sub CmdDuppliquer_Click()
    MenuFichierDuppliquer_Click
End Sub

Private Sub CmdInserer_Click()
    MenuActionInserer_Click
End Sub

Private Sub CmdModifierQuantite_Click(Index As Integer)
    Select Case Index
    Case 0:
        SliderQuantite(0).Value = SliderQuantite(0).Value + 1
        SliderQuantite(1).Value = SliderQuantite(1).Value + 1
    Case 1:
        SliderQuantite(0).Value = SliderQuantite(0).Value - 1
        SliderQuantite(1).Value = SliderQuantite(1).Value - 1
    End Select
    If SliderQuantite(0).Visible Then SliderQuantite_Scroll 0
    If SliderQuantite(1).Visible Then SliderQuantite_Scroll 1
End Sub

Private Sub CmdNouveau_Click()
    MenuFichierNouveau_Click
End Sub

Private Sub CmdQuantite_Click(Index As Integer)
    If Index < 4 Then
        SliderQuantite(0).Value = Val(CmdQuantite(Index).Caption)
        SliderQuantite_Scroll 0
    Else
        SliderQuantite(1).Value = Val(CmdQuantite(Index).Caption)
        SliderQuantite_Scroll 1
    End If
End Sub

Private Sub CmdRenommer_Click()
    MenuFichierRenommer_Click
End Sub

Private Sub CmdSupprimer_Click()
    MenuActionsSupprimer_Click
End Sub

Private Sub CmsSupprimerIA_Click()
    MenuFichierSupprimer_Click
End Sub

Private Sub FileIAs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu MenuFichier
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Modification Then
        If MsgBox(MessageEnregistrementQuitter, vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            MenuFichierEnregistrer_Click
        End If
    End If
    Charger_Dossiers_IAs
    FrmParam.Show
End Sub

Private Sub LstActions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu MenuActions
    End If
End Sub

Private Sub LstValeur_Click(Index As Integer)
    Dim i As Integer
    If OptionAction(0).Value Then
        'Recherche l'indice de la ressource sélectionnée.
        For i = 0 To Parametres.NombreRessources - 1
            If LstValeur(Index).List(LstValeur(Index).ListIndex) = Parametres.Ressources_Nom(i) Then
                IA2s(IndiceIA).Definir_Valeur(SliderEpoque.Value, LstActions.ListIndex) = i
                Exit For
            End If
        Next i
    ElseIf OptionAction(2).Value Then
        'On recherche l'indice de l'objet sélectionné.
        For i = 0 To Parametres.NombreObjets - 1
            If LstValeur(Index).List(LstValeur(Index).ListIndex) = Parametres.Objet_Nom(i) Then
                IA2s(IndiceIA).Definir_Valeur(SliderEpoque.Value, LstActions.ListIndex) = i
                Exit For
            End If
        Next i
    ElseIf OptionAction(7).Value Then
        'Choisit un mode
        IA2s(IndiceIA).Definir_Valeur(SliderEpoque.Value, LstActions.ListIndex) = LstValeur(Index).ListIndex
    Else
        'Choisit le batiment sélectionné.
        IA2s(IndiceIA).Definir_Valeur(SliderEpoque.Value, LstActions.ListIndex) = Parametres.Batiment_Tri(LstValeur(Index).ListIndex)
    End If
    Actualiser_Action LstActions.ListIndex
    Modification = True
End Sub

Private Sub MenuActionEpoqueChanger_Click(Index As Integer)
    Select Case Index
    Case 0:
        SliderEpoque.Value = SliderEpoque.Value - 1
    Case 1:
        SliderEpoque.Value = SliderEpoque.Value + 1
    End Select
    SliderEpoque_Scroll
End Sub

Private Sub MenuActionInserer_Click()
    IA2s(IndiceIA).Ajouter_Etape SliderEpoque, IA2s(IndiceIA).Definir_NombreEtapes(SliderEpoque)
    Actualiser_Actions
    LstActions.ListIndex = LstActions.ListCount - 1
    Modification = True
End Sub

Private Sub MenuActionsPrecedent_Click()
    CmdDeplacer_Click 0
End Sub

Private Sub MenuActionsSuivant_Click()
    CmdDeplacer_Click 1
End Sub

Private Sub MenuActionsSupprimer_Click()
    Dim Temp As Integer
    IA2s(IndiceIA).Supprimer_Etape SliderEpoque.Value, LstActions.ListIndex
    Temp = LstActions.ListIndex
    Modification = True
    Actualiser_Actions
    If Temp <= LstActions.ListCount - 1 Then
        LstActions.ListIndex = Temp
    End If
End Sub

Private Sub MenuAideAPropos_Click()
    FrmAPropos2.Show
End Sub

Private Sub MenuAideManuel_Click()
    Editer_Texte App.Path & FicIni.Chemin & Langues.Dossier & CheminFichierInfo
End Sub

Private Sub MenuFichierDuppliquer_Click()
    Duppliquer_IA FileIAs.List(FileIAs.ListIndex), FileIAs.List(FileIAs.ListIndex) & SuffixeFichierDupplication
    Actualiser_IAs
    Modification = True
    'Ce place sur le nouveau fichier.
    Selectionner_IA FileIAs.List(FileIAs.ListIndex) & SuffixeFichierDupplication
End Sub

Private Sub MenuFichierEnregistrer_Click()
    Enregistrer_IAs
    FrmParam.FileIA.Refresh
    FrmParam.LstPerso_Click
    Modification = False
End Sub

Private Sub MenuFichierNouveau_Click()
    Dim i As Integer
    Dim Res As String
    Res = InputBox(MessageNouveau, Me.Caption, FichierNouveau & ExtentionFichierIA)
    Nouvelle_IA Res
    Actualiser_IAs
    For i = 0 To FileIAs.ListCount - 1
        If FileIAs.List(i) = Res Then
            FileIAs.ListIndex = i
            i = FileIAs.ListCount
        End If
    Next i
    Modification = True
    Selectionner_IA Res
End Sub

Private Sub MenuFichierQuitter_Click()
    Unload Me
End Sub

Private Sub MenuFichierRenommer_Click()
    Dim Res As String
    Res = InputBox(MessageRenommer, Me.Caption, FileIAs.List(FileIAs.ListIndex))
    If Res <> FileIAs.List(FileIAs.ListIndex) Then
        Renommer_IA FileIAs.List(FileIAs.ListIndex), Res
        Actualiser_IAs
    End If
    Selectionner_IA Res
End Sub

Private Sub MenuFichierSupprimer_Click()
    If MsgBox(MessageSupprimer1 & FileIAs.List(FileIAs.ListIndex) & MessageSupprimer2, vbYesNo + vbQuestion, Me.Caption) = vbYes Then
        Supprimer_IA FileIAs.List(FileIAs.ListIndex)
        Actualiser_IAs
        Modification = True
    End If
End Sub

Private Sub OptionAction_Click(Index As Integer)
    IA2s(IndiceIA).Definir_Categorie(SliderEpoque.Value, LstActions.ListIndex) = Index
    Modification = True
    Actualiser_Action LstActions.ListIndex
    Actualiser_Definition_Action
    If SliderQuantite(0).Visible Then
        SliderQuantite_Scroll 0
    Else
        SliderQuantite_Scroll 1
    End If
End Sub

Private Sub SliderEpoque_Scroll()
    LblEpoque.Caption = Parametres.Epoque_Nom(SliderEpoque.Value)
    Actualiser_Actions
End Sub
Private Sub SliderQuantite_Scroll(Index As Integer)
    LblQuantite = SliderQuantite(Index).Value
    IA2s(IndiceIA).Definir_Quantite(SliderEpoque.Value, LstActions.ListIndex) = SliderQuantite(Index).Value
    Actualiser_Action LstActions.ListIndex
    Modification = True
End Sub

Private Sub FileIAs_Click()
    Dim i  As Long
    'Recherche l'inde de l'IA concerné.
    For i = 0 To FileIAs.ListCount - 1
        If IA2s(i).Nom = FileIAs.List(FileIAs.ListIndex) Then
            IndiceIA = i
            i = FileIAs.ListCount
        End If
    Next i
    TxtDescription.Text = IA2s(IndiceIA).Commentaire
    Actualiser_Actions
End Sub

Private Sub LstActions_Click()
    Actualiser_Definition_Action
End Sub

Private Sub Actualiser_IAs()
    Dim Temp As Integer
    Temp = FileIAs.ListIndex
    FileIAs.Refresh
    If Temp >= 0 And Temp < FileIAs.ListCount - 1 Then
        FileIAs.ListIndex = Temp
    Else
        FileIAs.ListIndex = 0
    End If
End Sub

Private Sub Actualiser_Actions()
    Dim i As Integer
    With IA2s(IndiceIA)
    
    LstActions.Clear
    TxtCommentaire.Text = .Definir_Commentaire(SliderEpoque.Value)
    For i = 0 To .Definir_NombreEtapes(SliderEpoque.Value) - 1
        LstActions.AddItem ""
        Actualiser_Action i
    Next i
    LstActions.ListIndex = 0
    End With
    CmdSupprimer.Enabled = LstActions.ListCount > 1
    MenuActionsSupprimer.Enabled = CmdSupprimer.Enabled
End Sub

Private Sub Actualiser_Action(ByVal i As Integer)
    With IA2s(IndiceIA)
    Select Case .Definir_Categorie(SliderEpoque.Value, i)
        'Ressources.
        Case 0:
            If .Definir_Valeur(SliderEpoque.Value, i) > Parametres.NombreRessources - 1 Then
                .Definir_Valeur(SliderEpoque.Value, i) = 0
            End If
            LstActions.List(i) = ListeRessources & _
                                 .Definir_Quantite(SliderEpoque.Value, i) & " " & _
                                 Parametres.Ressources_Nom(.Definir_Valeur(SliderEpoque.Value, i))
        'Batiment.
        Case 1: LstActions.List(i) = ListeBatiments & _
                                     Parametres.Batiment_Nom(.Definir_Valeur(SliderEpoque.Value, i))
        'Objet.
        Case 2: LstActions.List(i) = ListeObjets & _
                                     .Definir_Quantite(SliderEpoque.Value, i) & " " & _
                                     Parametres.Objet_Nom(.Definir_Valeur(SliderEpoque.Value, i))
        'Travailler.
        Case 3: LstActions.List(i) = ListeTravailler
        'Patrouiller.
        Case 4: LstActions.List(i) = ListePatrouiller
        'Attaquer ennemi.
        Case 5: LstActions.List(i) = ListeAttaquerEnnemi
        'Travailler.
        Case 6: LstActions.List(i) = ListeAttaquerChateau
        'Changement de mode.
        Case 7:
'            Select Case .Definir_Valeur(SliderEpoque.Value, i)
'            Case 0: LstActions.List(i) = ListeModePacifiste
'            Case 1: LstActions.List(i) = ListeModeCombatif
'            Case 2: LstActions.List(i) = ListeModeDonnerTout
'            End Select
            If .Definir_Valeur(SliderEpoque.Value, i) >= NombreModes Then
                .Definir_Valeur(SliderEpoque.Value, i) = 0
            End If
            LstActions.List(i) = ListeModes(.Definir_Valeur(SliderEpoque.Value, i))
            'End If
        Case 8:
            If .Definir_Valeur(SliderEpoque.Value, i) > Parametres.NombreTypeBatiments - 1 Then
                .Definir_Valeur(SliderEpoque.Value, i) = 0
            End If
            LstActions.List(i) = ListeGagnerArgent & " " & .Definir_Quantite(SliderEpoque.Value, i) & " " & Parametres.EtiquetteArgent & " (" & Parametres.Batiment_Metier(.Definir_Valeur(SliderEpoque.Value, i), False) & ")"
        Case 9: 'Gagner de l'argent en vendant des ressources.
            LstActions.List(i) = ListeGagnerArgent & " " & .Definir_Quantite(SliderEpoque.Value, i) & " " & Parametres.EtiquetteArgent & " (" & Parametres.Ressources_Nom(.Definir_Valeur(SliderEpoque.Value, i)) & ")"
    End Select
    End With
    Actualiser_Apercu
End Sub

Private Sub Actualiser_Definition_Action()
    Dim i, j As Integer
    Dim Ligne As String
    Dim Temp As Boolean
    With IA2s(IndiceIA)
    OptionAction(.Definir_Categorie(SliderEpoque.Value, LstActions.ListIndex)).Value = True
    
    SliderQuantite(0).Value = .Definir_Quantite(SliderEpoque.Value, LstActions.ListIndex)
    SliderQuantite(1).Value = .Definir_Quantite(SliderEpoque.Value, LstActions.ListIndex)
    LblQuantite = SliderQuantite(1).Value
    
    Select Case .Definir_Categorie(SliderEpoque.Value, LstActions.ListIndex)
    Case 0: 'Ressources.
        LstValeur(0).Clear
        LstValeur(0).Visible = True
        LstValeur(1).Visible = False
        For i = 0 To Parametres.NombreRessources - 1
            If Parametres.Ressources_Epoque(i) <= SliderEpoque.Value Then
                LstValeur(0).AddItem Parametres.Ressources_Nom(i)
            End If
        Next i
        If IA2s(IndiceIA).Definir_Valeur(SliderEpoque.Value, LstActions.ListIndex) > Parametres.NombreRessources - 1 Then
            LstValeur(0).ListIndex = 0
        Else
            For i = 0 To LstValeur(0).ListCount - 1
                If Parametres.Ressources_Nom(IA2s(IndiceIA).Definir_Valeur(SliderEpoque.Value, LstActions.ListIndex)) = LstValeur(0).List(i) Then
                    LstValeur(0).ListIndex = i
                    Exit For
                End If
            Next i
        End If
    Case 1: 'Batiment.
        LstValeur(0).Clear
        LstValeur(0).Visible = True
        LstValeur(1).Visible = False
        For i = 0 To Parametres.NombreTypeBatiments - 1
            If Parametres.Batiment_NoEpoque(Parametres.Batiment_Tri(i)) <= SliderEpoque.Value + 1 Then
                Ligne = Parametres.Batiment_Nom(Parametres.Batiment_Tri(i)) & " ("
                'Affiche le prix du batiment.
                If Parametres.Batiment_PrixArgent(Parametres.Batiment_Tri(i)) > 0 Then
                    Ligne = Ligne & Parametres.Batiment_PrixArgent(Parametres.Batiment_Tri(i)) & " " & Parametres.EtiquetteArgent
                    Temp = True
                Else
                    Temp = False
                End If
                For j = 0 To Parametres.NombreRessources - 1
                    If Parametres.Batiment_PrixRessource(Parametres.Batiment_Tri(i), j) > 0 Then
                        If Temp Then
                            Ligne = Ligne & ", "
                        End If
                        Ligne = Ligne & Parametres.Batiment_PrixRessource(Parametres.Batiment_Tri(i), j) & " " & _
                                        Parametres.Ressources_Nom(j)
                        Temp = True
                    End If
                Next j
                LstValeur(0).AddItem Ligne & ")"
            Else
                Exit For
            End If
        Next i
        If LstValeur(0).ListCount > 0 Then
            'If .Definir_Valeur(SliderEpoque.Value, LstActions.ListIndex) < LstValeur(0).ListCount Then
                LstValeur(0).ListIndex = Parametres.Batiment_TriInverse(.Definir_Valeur(SliderEpoque.Value, LstActions.ListIndex))
            'Else
            '    LstValeur(0).ListIndex = 0
            'End If
        End If
    Case 2: 'Objets.
        LstValeur(1).Clear
        LstValeur(1).Visible = True
        LstValeur(0).Visible = False
        For i = 0 To Parametres.NombreObjets - 1
            If Parametres.Objet_Epoque(i) <= SliderEpoque.Value + 1 And _
               Parametres.Objet_Epoque(i) >= 0 Then
                LstValeur(1).AddItem Parametres.Objet_Nom(i)
            End If
        Next i
        For i = 0 To LstValeur(1).ListCount - 1
            If Parametres.Objet_Nom(IA2s(IndiceIA).Definir_Valeur(SliderEpoque.Value, LstActions.ListIndex)) = LstValeur(1).List(i) Then
                LstValeur(1).ListIndex = i
                i = LstValeur(1).ListCount
            End If
        Next i
    Case 7: 'Modes.
        LstValeur(0).Clear
        LstValeur(0).Visible = True
        LstValeur(1).Visible = False
        For i = 0 To NombreModes - 1
            LstValeur(0).AddItem ListeModes(i)
        Next i
        If LstValeur(0).ListCount > 0 Then
            If .Definir_Valeur(SliderEpoque.Value, LstActions.ListIndex) < LstValeur(0).ListCount Then
                LstValeur(0).ListIndex = .Definir_Valeur(SliderEpoque.Value, LstActions.ListIndex)
            Else
                LstValeur(0).ListIndex = 0
            End If
        End If
    Case 8: 'Argent.
        LstValeur(0).Clear
        LstValeur(0).Visible = True
        LstValeur(1).Visible = False
        For i = 0 To Parametres.NombreTypeBatiments - 1
            If Parametres.Batiment_NoEpoque(Parametres.Batiment_Tri(i)) <= SliderEpoque.Value + 1 Then
                LstValeur(0).AddItem Parametres.Batiment_Metier(Parametres.Batiment_Tri(i), False) & " (" & _
                                     Parametres.Batiment_Nom(Parametres.Batiment_Tri(i)) & ")"
            Else
                Exit For
            End If
        Next i
        If LstValeur(0).ListCount > 0 Then
            'If .Definir_Valeur(SliderEpoque.Value, LstActions.ListIndex) < LstValeur(0).ListCount Then
                LstValeur(0).ListIndex = Parametres.Batiment_TriInverse(.Definir_Valeur(SliderEpoque.Value, LstActions.ListIndex))
            'Else
            '    LstValeur(0).ListIndex = 0
            'End If
        End If
    Case 9: 'Gagner de l'argent avec des ressources.
        LstValeur(0).Clear
        LstValeur(0).Visible = True
        LstValeur(1).Visible = False
        For i = 0 To Parametres.NombreRessources - 1
            If Parametres.Ressources_Recoltable(i) Then
                LstValeur(0).AddItem Parametres.Ressources_Nom(i)
            End If
        Next i
        For i = 0 To LstValeur(0).ListCount - 1
            If Parametres.Ressources_Nom(IA2s(IndiceIA).Definir_Valeur(SliderEpoque.Value, LstActions.ListIndex)) = LstValeur(0).List(i) Then
                LstValeur(0).ListIndex = i
                Exit For
            End If
        Next i
    End Select
    Select Case .Definir_Categorie(SliderEpoque.Value, LstActions.ListIndex)
    Case 0, 2: FrameDesignation.Visible = True
               FrameQuantite.Visible = True
               SliderQuantite(0).Visible = True
               SliderQuantite(1).Visible = False
               CmdQuantite(0).Visible = True
               CmdQuantite(1).Visible = True
               CmdQuantite(2).Visible = True
               CmdQuantite(3).Visible = True
               CmdQuantite(4).Visible = False
               CmdQuantite(5).Visible = False
               CmdQuantite(6).Visible = False
               CmdQuantite(7).Visible = False
    Case 1, 7: FrameDesignation.Visible = True
               FrameQuantite.Visible = False
    Case 3 To 6: FrameDesignation.Visible = False
                 FrameQuantite.Visible = False
    Case 8, 9: FrameDesignation.Visible = True
            FrameQuantite.Visible = True
            SliderQuantite(0).Visible = False
            SliderQuantite(1).Visible = True
            CmdQuantite(0).Visible = False
            CmdQuantite(1).Visible = False
            CmdQuantite(2).Visible = False
            CmdQuantite(3).Visible = False
            CmdQuantite(4).Visible = True
            CmdQuantite(5).Visible = True
            CmdQuantite(6).Visible = True
            CmdQuantite(7).Visible = True
    End Select
    End With
    'Active/désactive les boutons.
    CmdDeplacer(0).Enabled = LstActions.ListIndex <> 0
    MenuActionsPrecedent.Enabled = CmdDeplacer(0).Enabled
    CmdDeplacer(1).Enabled = LstActions.ListIndex <> LstActions.ListCount - 1
    MenuActionsSuivant.Enabled = CmdDeplacer(1).Enabled
End Sub

Private Sub Actualiser_Apercu()
    On Error GoTo Erreur
    Dim i, j, k, l As Integer
    'Efface les aperçus précedents
    For i = 0 To NombreMaxImagesApercu - 1
        ImgAperçu(i).Visible = False
        LblApercu(i).Visible = False
    Next i
    'Affiche l'aperçu de l'objet.
    For i = 0 To Parametres.NombreEpoques - 1
        'On met un intermédiaire entre les époques.
        ImgAperçu(l).Visible = False
        LblApercu(l).Visible = True
        LblApercu(l).Caption = i + 1 & " :"
        l = l + 1
        For j = 0 To IA2s(IndiceIA).Definir_NombreEtapes(i) - 1
            If IA2s(IndiceIA).Definir_Categorie(i, j) = 2 Then
                For k = 0 To IA2s(IndiceIA).Definir_Quantite(i, j) - 1
                    If Parametres.Objet_CheminImage(IA2s(IndiceIA).Definir_Valeur(i, j)) = "" Then
                        ImgAperçu(l).Picture = LoadPicture(App.Path & CheminImagesObjets & Parametres.Objet_Image(IA2s(IndiceIA).Definir_Valeur(i, j)) & ExtentionFichiersImages)
                    Else
                        ImgAperçu(l).Picture = LoadPicture(Parametres.Objet_CheminImage(IA2s(IndiceIA).Definir_Valeur(i, j)) & Parametres.Objet_Image(IA2s(IndiceIA).Definir_Valeur(i, j)) & ExtentionFichiersImages)
                    End If
                    ImgAperçu(l).ToolTipText = Parametres.Objet_Nom(IA2s(IndiceIA).Definir_Valeur(i, j)) & _
                                               " (" & Parametres.EtiquettePrix & " : " & _
                                               Parametres.Objet_PrixVente(IA2s(IndiceIA).Definir_Valeur(i, j)) * Parametres.CoefficientAchat & _
                                               ")"
                    ImgAperçu(l).Visible = True
                    LblApercu(l).Visible = False
                    l = l + 1
                    If l >= NombreMaxImagesApercu - 1 Then
                        k = IA2s(IndiceIA).Definir_Quantite(i, j)
                        j = IA2s(IndiceIA).Definir_NombreEtapes(i)
                        i = Parametres.NombreEpoques
                    End If
                Next k
            End If
        Next j
    Next i
    Exit Sub
Erreur:
    MsgBox MessageErreurChargementImage & App.Path & CheminImagesObjets & Parametres.Objet_Nom(IA2s(IndiceIA).Definir_Valeur(i, j)) & ExtentionFichiersImages, vbCritical, MessageErreur
End Sub

Private Sub Selectionner_IA(ByVal Nom As String)
    Dim i As Long
    'Ce place sur le fichier renommé
    For i = 0 To FileIAs.ListCount - 1
        If Nom = FileIAs.List(i) Then
            FileIAs.ListIndex = i
            i = FileIAs.ListCount
        End If
    Next i
End Sub

Private Sub TxtCommentaire_Change()
    IA2s(IndiceIA).Definir_Commentaire(SliderEpoque.Value) = TxtCommentaire
    Modification = True
End Sub

Private Sub TxtDescription_Change()
    IA2s(IndiceIA).Commentaire = TxtDescription
    Modification = True
End Sub

