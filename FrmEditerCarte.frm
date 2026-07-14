VERSION 5.00
Begin VB.Form FrmEditeurCarte 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10440
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15240
   ControlBox      =   0   'False
   Icon            =   "FrmEditerCarte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   15240
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   2640
      TabIndex        =   14
      Top             =   120
      Width           =   6015
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   200
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   200
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "xxx"
         Height          =   255
         Left            =   4560
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "cases"
         Height          =   255
         Left            =   5400
         TabIndex        =   20
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Dimensions :"
         Height          =   255
         Left            =   3480
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Hauteur :"
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Largeur :"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.PictureBox PicTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   6375
      Left            =   2640
      TabIndex        =   8
      Top             =   840
      Width           =   6015
      Begin VB.PictureBox PicSelection 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5640
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   22
         Top             =   6000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox PicCarte 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   5775
         Left            =   5640
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   6000
         Width           =   5535
      End
   End
   Begin VB.CommandButton CmsSupprimerIA 
      Caption         =   "Supprimer"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton CmdRenommer 
      Caption         =   "Renommer"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton CmdDuppliquer 
      Caption         =   "Duppliquer"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nouveau"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cartes"
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
      Begin VB.ListBox LstCartes 
         Height          =   4350
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.FileListBox File1 
         Height          =   4380
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   495
      Begin VB.PictureBox PicTerrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Menu MenuFichier 
      Caption         =   "&Fichiers"
      Visible         =   0   'False
      Begin VB.Menu MenuFichierNouveau 
         Caption         =   "&Nouveau"
      End
      Begin VB.Menu MenuFichierFermer 
         Caption         =   "&Fermer"
      End
      Begin VB.Menu menuFichierSeparateur1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFichierQuitter 
         Caption         =   "&Quitter"
      End
   End
End
Attribute VB_Name = "FrmEditeurCarte"
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

Const NombreTerrainDifferent = 11
Const ImageTerrain = "\Images\terrains\" & "Plaine" & ".bmp"
Const Hauteur = 16
Const Largeur = 16

Private Type Carte
    Largeur As Integer
    Hauteur As Integer
    Cases() As Integer
End Type

Private Sub Form_Load()
    Dim i As Integer
    FrmParam.Hide
    Me.Top = 0
    Me.Left = 0
    Set PicTemp.Picture = LoadPicture(App.Path & ImageTerrain)
    
    'Prépare l'aperçu.
    For i = 0 To NombreTerrainDifferent - 1
        If i > 0 Then
            Load PicTerrain(i)
            PicTerrain(i).Left = PicTerrain(i - 1).Left
            PicTerrain(i).Top = PicTerrain(i - 1).Top + PicTerrain(i - 1).Height
            PicTerrain(i).Visible = True
        End If
        'Set PicTerrain(i).Picture = LoadPicture("")
        PicTerrain(i).PaintPicture PicTemp.Picture, _
                                   0, 0, _
                                   Largeur, Hauteur, _
                                   Largeur, i * Hauteur, _
                                   Largeur, Hauteur
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmParam.Show
End Sub

Private Sub PicTerrain_Click(Index As Integer)
    'PicSelection.PaintPicture PicTerrain(Index).Picture,  0, 0, Largeur, Hauteur, 0, 0, Largeur, Hauteur
End Sub
