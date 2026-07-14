VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAffTresors"
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

Const Chemin = "Decors\"

Const LargeurImage As Long = 26
Const HauteurImage As Long = 26

Private Sprite As ClsAffSprite

Private Sub Class_Initialize()
    Set Sprite = New ClsAffSprite
End Sub

Public Sub Charger(ByRef dd As DirectDraw7)
    Sprite.Charger dd, Chemin & "Tresors"
End Sub

Public Sub Afficher(ByVal X As Long, _
                    ByVal Y As Long, _
                    ByVal TypeTresor As Long, _
                    ByRef SurfBack As DirectDrawSurface7)
    'On charge la bonne image en fonction des paramètres.
    Sprite.Selection 0, _
                     HauteurImage, _
                     TypeTresor * LargeurImage, _
                     (TypeTresor + 1) * LargeurImage
    'On affiche l'objet sur l'écran.
    Sprite.Afficher X, Y, SurfBack, True
End Sub

Public Property Get Largeur() As Long
    Largeur = LargeurImage
End Property

Public Property Get Hauteur() As Long
    Hauteur = HauteurImage
End Property



