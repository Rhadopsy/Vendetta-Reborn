VERSION 5.00
Begin VB.Form FrmApparence 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#Choisissez une apparence"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3270
   Icon            =   "FrmApparence.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   391
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   218
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox FileApparence 
      Height          =   3210
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.PictureBox PicApparence 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Index           =   0
      Left            =   360
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   72
      TabIndex        =   4
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Timer TimerApercu 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox PicApercu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   2040
      ScaleHeight     =   128
      ScaleMode       =   0  'User
      ScaleWidth      =   72
      TabIndex        =   2
      Top             =   3360
      Width           =   1080
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "#&OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton CmdAleatoire 
      Caption         =   "#Apparence aléatoire"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label LblApercu 
      Alignment       =   2  'Center
      Caption         =   "#Aperçu :"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   1024
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu MenuApparence 
      Caption         =   "#&Apparence"
      Begin VB.Menu MenuApparenceAleatoire 
         Caption         =   "#&Aléatoire"
         Shortcut        =   ^R
      End
      Begin VB.Menu MenuApparenceSeparataeur 
         Caption         =   "-"
      End
      Begin VB.Menu MenuApparenceQuitter 
         Caption         =   "#&Quitter"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "FrmApparence"
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

Const FichierINI = "Fenetre_Apparence"
Const SectionINI = "Fenetre_Apparence"

Private ExtentionFichierImage As String
Private LargeurColonne As Integer

Private NombreApparences As Integer

Private Sub Form_Load()
    On Error GoTo Erreur
    Dim i As Integer
    FrmParam.Hide

    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    
    ExtentionFichierImage = FicIni.Parametre("ExtentionFichierImage")
    LargeurColonne = Val(FicIni.Parametre("LargeurColonne"))
    TimerApercu.Interval = Val(FicIni.Parametre("VitesseApercu"))
    
    Caption = FicIni.Parametre("EtiquetteNomFenetre")
    LblApercu.Caption = FicIni.Parametre("EtiquetteApercu")
    CmdAleatoire.Caption = FicIni.Parametre("EtiquetteBoutonApparenceAleatoire")
    CmdQuitter.Caption = FicIni.Parametre("EtiquetteBoutonQuitter")
    MenuApparence.Caption = FicIni.Parametre("EtiquetteMenuApparence")
    MenuApparenceAleatoire.Caption = FicIni.Parametre("EtiquetteMenuApparenceAleatoire")
    MenuApparenceQuitter.Caption = FicIni.Parametre("EtiquetteMenuApparenceQuitter")
    
    NombreApparences = Parametres.Race_DefinirNombreCategoriesApparences(Persos(NoPerso).Race)
    For i = 0 To NombreApparences - 1
        If i > 0 Then
            Load FileApparence(i)
            FileApparence(i).Top = FileApparence(i - 1).Top
            FileApparence(i).Left = FileApparence(i - 1).Left + LargeurColonne
            FileApparence(i).Visible = True
            Load PicApparence(i)
            PicApparence(i).Top = PicApparence(i - 1).Top
            PicApparence(i).Left = PicApparence(i - 1).Left + LargeurColonne
            PicApparence(i).Visible = True
            LblApercu.Left = LblApercu.Left + LargeurColonne
            PicApercu.Left = PicApercu.Left + LargeurColonne
            CmdQuitter.Left = CmdQuitter.Left + LargeurColonne
            Me.Width = Me.Width + 1780
        End If
        FileApparence(i).Path = AffApparence.Definir_Chemin(i, Persos(NoPerso).Race, Persos(NoPerso).Feminin, Parametres)
        'FileApparence(i).Path = AffApparence.Definir_Chemin(i, Persos(0).Feminin)
        FileApparence(i).ListIndex = 0
        'FrameApparence(i).Visible = True
    Next i
    
    'Centre la feuille.
    Move (MDIFrmMain.Width - Me.Width) / 2, (MDIFrmMain.Height - Me.Height) / 2
    
    Charger_Apparences
    Actualisation
    Exit Sub
Erreur:
    MsgBox "Le chemin des apparences : " & AffApparence.Definir_Chemin(i, Persos(NoPerso).Race, Persos(NoPerso).Feminin, Parametres) & " est introuvable.", vbCritical, "Erreur"
    End
End Sub

Private Sub FileApparence_Click(Index As Integer)
    With FileApparence(Index)
    Set PicApparence(Index).Picture = LoadPicture(.Path & "\" & Left$(.List(.ListIndex), InStr(.List(.ListIndex), ".") - 1) & ExtentionFichierImage)
    End With
    Actualisation
End Sub

Public Sub Actualiser_Apercu()
    Dim i As Integer
    Dim X, Y As Single
    'With FileApparence(0)
    'Set PicApercu.Picture = LoadPicture(.Path & "\" & Left$(.List(.ListIndex), InStr(.List(.ListIndex), ".") - 1) & ExtentionFichierImage)
    'End With
    MousePointer = 11
    For i = 0 To NombreApparences - 1
        'If FrameApparence(i).Visible Then
            For X = 0 To PicApparence(i).Width - 1
                For Y = 0 To PicApparence(i).Height - 1
                    If PicApparence(i).Point(X, Y) > 0 Or i = 0 Then
                        'PicApercu.PaintPicture PicApparence(i).Picture, X, Y, 1, 1, X, Y, 1, 1
                        PicApercu.PSet (X, Y), PicApparence(i).Point(X, Y)
                    End If
                Next Y
            Next X
        'End If
    Next i
    MousePointer = 0
End Sub

Private Sub Charger_Apparences()
    Dim i, j As Integer
    For i = 0 To NombreApparences - 1
        With FileApparence(i)
        For j = 0 To .ListCount - 1
            If .List(j) = Persos(NoPerso).Definir_FichierApparenceBase(i) Then
                .ListIndex = j
                j = .ListCount
            End If
        Next j
        If .ListIndex = -1 Then .ListIndex = 0
        End With
    Next i
End Sub

Private Sub Enregistrer_Apparences()
    Dim i As Integer
    For i = 0 To NombreApparences - 1
        With FileApparence(i)
            Persos(NoPerso).Definir_FichierApparenceBase(i) = .List(.ListIndex)
        End With
    Next i
    SavePicture PicApercu.Image, App.Path & CheminSavPerso & Persos(NoPerso).Nom & ExtentionFichierImage
    Enregistrer_Sauvegarde_Perso Persos(NoPerso)
End Sub

Private Sub MenuApparenceAleatoire_Click()
    CmdAleatoire_Click
End Sub

Public Sub CmdAleatoire_Click()
    Dim i As Integer
    For i = 0 To NombreApparences - 1
        With FileApparence(i)
        .ListIndex = Int(Rnd * .ListCount)
        End With
    Next i
End Sub

Private Sub Actualisation()
    'Actualiser_Apercu
    TimerApercu.Enabled = True
End Sub

Private Sub MenuApparenceQuitter_Click()
    CmdQuitter_Click
End Sub

Public Sub CmdQuitter_Click()
    Unload Me
End Sub

Private Sub TimerApercu_Timer()
    Actualiser_Apercu
    TimerApercu.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Actualiser_Apercu
    Enregistrer_Apparences
    FrmParam.Show
    FrmParam.OptionApparence_Click 0
End Sub
