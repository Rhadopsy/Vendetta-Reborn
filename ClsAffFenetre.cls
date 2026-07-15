VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAffFenetre"
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

Const Chemin = "Fenetres\"
Const Translucide = False
Const Transparent = True

Public Largeur As Long
Public Hauteur As Long

Private NombreSprites As Integer
Private Sprites() As ClsAffSprite
Private Images() As String

Private Sub Class_Initialize()
    ReDim Sprites(0)
    Set Sprites(0) = New ClsAffSprite
End Sub

Public Sub Charger(ByRef dd As IDirectDraw7)
    Dim i As Integer
    For i = 0 To NombreSprites - 1
        Sprites(i).Charger dd, Chemin & Images(i)
    Next i
End Sub

Public Sub Afficher(ByVal X As Long, _
                    ByVal Y As Long, _
                    ByRef SurfBack As IDirectDrawSurface7, _
                    Optional ByVal TypeFenetre As Integer)
    'On charge la bonne image en fonction des paramètres.
    'Sprite.Selection Hauteur * Etat, _
                     Hauteur * (Etat + 1), _
                     Largeur * TypeFenetre, _
                     Largeur * (TypeFenetre + 1)
    'On affiche le terrain sur l'arrière plan.
    Sprites(TypeFenetre).Afficher X, Y, SurfBack, Transparent, Translucide
End Sub

Public Property Get Definir_Nombre_Sprites() As Integer
    Definir_Nombre_Sprites = NombreSprites
End Property
Public Property Let Definir_Nombre_Sprites(ByVal Valeur As Integer)
    Dim i As Integer
    If Valeur = 0 Then Valeur = 1
    NombreSprites = Valeur
    ReDim Sprites(Valeur - 1)
    ReDim Images(Valeur - 1)
    For i = 0 To NombreSprites - 1
        Set Sprites(i) = New ClsAffSprite
        Sprites(i).Selection 0, _
                             Hauteur, _
                             0, _
                             Largeur
    Next i
End Property

Public Property Get Definir_Image(ByVal Indice As Integer) As String
    Definir_Image = Images(Indice)
End Property
Public Property Let Definir_Image(ByVal Indice As Integer, _
                                  ByVal Valeur As String)
    Images(Indice) = Valeur
End Property

