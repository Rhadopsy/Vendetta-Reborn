VERSION 5.00
Begin VB.Form FrmAPropos2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#À propos de MonApplication"
   ClientHeight    =   3780
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "FrmAPropos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2609.023
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdLisezMoi 
      Caption         =   "#&Lisez moi"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   2880
      Width           =   1260
   End
   Begin VB.ListBox LstArtistes 
      Height          =   1620
      ItemData        =   "FrmAPropos.frx":0442
      Left            =   2520
      List            =   "FrmAPropos.frx":0444
      TabIndex        =   4
      Top             =   1080
      Width           =   3015
   End
   Begin VB.ListBox LstEquipe 
      Height          =   1620
      ItemData        =   "FrmAPropos.frx":0446
      Left            =   240
      List            =   "FrmAPropos.frx":0448
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton CmdOK 
      Cancel          =   -1  'True
      Caption         =   "#&OK"
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   3360
      Width           =   1260
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5408.938
      Y1              =   1905.001
      Y2              =   1905.001
   End
   Begin VB.Label LblRealisation 
      Caption         =   "#Réalisateur :"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label LblMail 
      Caption         =   "varaxor@free.fr"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label LblWebMaster 
      Caption         =   "#Webmaster :"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label LblMail 
      Caption         =   "sharlaut@netcourrier.com"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label LblDecouvrez 
      Caption         =   "#Découvrez les prochaines versions, des suppléments et d'autres bonus sur le site officiel :"
      Height          =   615
      Left            =   2520
      TabIndex        =   7
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label LblSiteWeb 
      Caption         =   "http://franck.leveque2.free.fr/Vendetta"
      Height          =   255
      Left            =   2520
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label LblAvertissement 
      Caption         =   $"FrmAPropos.frx":044A
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   1
      Top             =   2880
      Width           =   3870
   End
End
Attribute VB_Name = "FrmAPropos2"
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

Const FichierINI = "Fenetre_APropos"
Const SectionINI = "Fenetre_APropos"

Private LienEmail As String
Private CheminFichierLisezMoi As String

Private Sub Form_Load()
    Dim Temp As String
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    
    Caption = FicIni.Parametre("EtiquetteNomFenetre") & App.Title
    lblVersion.Caption = FicIni.Parametre("EtiquetteNomVersion") & App.Major & "." & App.Minor
    LblSiteWeb.ForeColor = Val(FicIni.Parametre("LienInternetCouleur"))
    LblMail(0).ForeColor = LblSiteWeb.ForeColor
    LblMail(1).ForeColor = LblSiteWeb.ForeColor
    LblSiteWeb.Font.Underline = Val(FicIni.Parametre("LienInternetSouligne"))
    LblMail(0).Font.Underline = True
    LblMail(1).Font.Underline = True
    Temp = App.Path & "\" & FicIni.Parametre("LienInternetCurseur")
    LblSiteWeb.MouseIcon = LoadPicture(Temp)
    LblMail(0).MouseIcon = LoadPicture(Temp)
    LblMail(1).MouseIcon = LoadPicture(Temp)
    LblSiteWeb.MousePointer = Val(FicIni.Parametre("LienInternetPointeur"))
    LblMail(0).MousePointer = LblSiteWeb.MousePointer
    LblMail(1).MousePointer = LblSiteWeb.MousePointer
    
    LblDecouvrez = FicIni.Parametre("EtiquetteDecouvrez")
    LblRealisation = FicIni.Parametre("EtiquetteRealisation")
    LblWebMaster = FicIni.Parametre("EtiquetteWebMaster")
    LblAvertissement = FicIni.Parametre("EtiquetteAvertissement")
    
    CmdLisezMoi.Caption = FicIni.Parametre("BoutonLisezMoi")
    cmdOK.Caption = FicIni.Parametre("BoutonOK")
    
    LienEmail = FicIni.Parametre("LienEmail")
    CheminFichierLisezMoi = FicIni.Chemin & Langues.Dossier & FicIni.Parametre("CheminFichierLisezMoi")
    
    LstEquipe.AddItem FicIni.Parametre("EtiquetteConception")
    LstEquipe.AddItem "- Sylvain Harlaut"
    LstEquipe.AddItem ""
    LstEquipe.AddItem FicIni.Parametre("EtiquetteProgrammation")
    LstEquipe.AddItem "- Sylvain Harlaut"
    LstEquipe.AddItem "- Leveque Franck"
    LstEquipe.AddItem "- Vincent Fleury"
    LstEquipe.AddItem "- Proger"
    LstEquipe.AddItem ""
    LstEquipe.AddItem FicIni.Parametre("EtiquetteProgrammationSiteInternet")
    LstEquipe.AddItem "- Leveque Franck"
    LstEquipe.AddItem ""
    LstEquipe.AddItem FicIni.Parametre("EtiquetteBetaTesteurs")
    LstEquipe.AddItem "- Mad Martigan"
    LstEquipe.AddItem "- Mastergame"
    LstEquipe.AddItem "- Galadrhim"
    LstEquipe.AddItem "- El Diablo"
    LstEquipe.AddItem "- Thug Seb"
    LstEquipe.AddItem "- Matt 050"
    LstEquipe.AddItem "- Tailss"
    LstEquipe.AddItem "- Poorlittle boy"
    LstEquipe.AddItem "- Ibnul"
    LstEquipe.AddItem "- VictorVVV"
    LstEquipe.AddItem ""
    LstEquipe.AddItem FicIni.Parametre("EtiquetteRemerciements")
    LstEquipe.AddItem "- Fab"
    LstEquipe.AddItem "- Yoha"
    LstEquipe.AddItem "- Robin"
    LstEquipe.AddItem "- Vampilou"
    
    LstArtistes.AddItem FicIni.Parametre("EtiquetteGraphismes")
    LstArtistes.AddItem "- Sylvain Harlaut"
    LstArtistes.AddItem "- Ibnul"
    LstArtistes.AddItem "- rpg maker 2000, 2003 et XP"
    LstArtistes.AddItem "- ASCII Corporation"
    LstArtistes.AddItem "- Yoji Ojima"
    LstArtistes.AddItem "- www.skytowergames.com"
    LstArtistes.AddItem "- http://alex.ombra.net"
    LstArtistes.AddItem "- Don Miguel"
    LstArtistes.AddItem "- Grand Dragon"
    LstArtistes.AddItem "- Lasse Kongo"
    LstArtistes.AddItem "- Arc"
    LstArtistes.AddItem "- Roman"
    LstArtistes.AddItem "- DarmaX"
    LstArtistes.AddItem "- Swordman Richy"
    LstArtistes.AddItem "- Kimi sae ireba"
    LstArtistes.AddItem "- Infamous Joe"
    LstArtistes.AddItem "- Luke Galli"
    LstArtistes.AddItem "- Silicon Hero"
    LstArtistes.AddItem "- S.OTT"
    LstArtistes.AddItem "- Yanbetari"
    LstArtistes.AddItem "- PPAO"
    LstArtistes.AddItem "- Summoner"
    LstArtistes.AddItem "- Cam'ron"
    LstArtistes.AddItem "- FFVII Guru"
    LstArtistes.AddItem "- Nanny"
    LstArtistes.AddItem "- Izlude"
    LstArtistes.AddItem "- Oz"
    LstArtistes.AddItem "- Drive By"
    LstArtistes.AddItem "- Sac"
    LstArtistes.AddItem "- DSK"
    LstArtistes.AddItem "- Estebest"
    LstArtistes.AddItem "- LasseCongo83"
    LstArtistes.AddItem "- Dark Phoenix"
    LstArtistes.AddItem "- Lasse Congo"
    LstArtistes.AddItem "- Guajino"
    LstArtistes.AddItem "- ^Loco^"
    LstArtistes.AddItem "- Bleeding Halo"
    LstArtistes.AddItem "- Weeby"
    LstArtistes.AddItem "- Suertermann"
    LstArtistes.AddItem "- Rowain"
    LstArtistes.AddItem "- Syxtem"
    LstArtistes.AddItem "- Log"
    LstArtistes.AddItem "- 2Beers"
    LstArtistes.AddItem "- MED"
    LstArtistes.AddItem "- Maggedon"
    LstArtistes.AddItem "- Mack & Blue"
    LstArtistes.AddItem "- Nallyamp"
    LstArtistes.AddItem "- Gormash"
    LstArtistes.AddItem "- Heimdall"
    LstArtistes.AddItem "- The Brave"
    LstArtistes.AddItem "- Tantalus"
    LstArtistes.AddItem "- Wayfarer Mage"
    LstArtistes.AddItem "- Dani"
    LstArtistes.AddItem "- Gost Clown"
    LstArtistes.AddItem "- SeglaWolf"
    LstArtistes.AddItem "- DarkMatter50"
    LstArtistes.AddItem "- Drakul"
    LstArtistes.AddItem "- Daiou"
    LstArtistes.AddItem "- Star Ocen Zero"
    LstArtistes.AddItem "- Vyse"
    LstArtistes.AddItem "- Viper"
    LstArtistes.AddItem ""
    LstArtistes.AddItem FicIni.Parametre("EtiquetteMusiquesEtSons")
    LstArtistes.AddItem "- rpg maker 2000, 2003 et XP"
    LstArtistes.AddItem "- ASCII Corporation"

End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub LblMail_Click(Index As Integer)
    Ouvir_PageWeb LienEmail & LblMail(Index).Caption
End Sub

Private Sub LblSiteWeb_Click()
    Ouvir_PageWeb LblSiteWeb.Caption
End Sub

Private Sub CmdLisezMoi_Click()
    Editer_Texte App.Path & CheminFichierLisezMoi
End Sub
