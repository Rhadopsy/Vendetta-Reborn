VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAffDecors"
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

Private Sprites() As ClsAffSprite
Private Largeur() As Integer
Private Hauteur() As Integer

Private Sub Class_Initialize()
    ReDim Sprites(0)
    Set Sprites(0) = New ClsAffSprite
    
    ReDim Largeur(0)
    ReDim Hauteur(0)
End Sub

Public Sub Charger(ByRef dd As DirectDraw7, ByVal Parametres As ClsJeuParam)
    Dim i As Long
    ReDim Sprites(Parametres.NombreDecors - 1)
    ReDim Largeur(Parametres.NombreDecors - 1)
    ReDim Hauteur(Parametres.NombreDecors - 1)
    For i = 0 To UBound(Sprites())
        Set Sprites(i) = New ClsAffSprite
        If Left(Parametres.Decor_Image(i), "1") = "." Then
            Sprites(i).Charger dd, Right(Parametres.Decor_Image(i), Len(Parametres.Decor_Image(i)) - 1), True
        Else
            Sprites(i).Charger dd, Chemin & Parametres.Decor_Image(i)
        End If
        Largeur(i) = Parametres.Decor_Largeur(i)
        Hauteur(i) = Parametres.Decor_Hauteur(i)
    Next i
End Sub

Public Sub Afficher(ByVal X As Long, _
                    ByVal Y As Long, _
                    ByVal TypeDecor As Long, _
                    ByRef SurfBack As DirectDrawSurface7, _
                    Optional ByVal EtapeAnimation As Integer = 0)
    'On charge la bonne image en fonction des paramètres.
    Sprites(TypeDecor).Selection 0, _
                                 Hauteur(TypeDecor), _
                                 Largeur(TypeDecor) * (EtapeAnimation), _
                                 Largeur(TypeDecor) * (EtapeAnimation + 1)
    'On affiche l'objet sur l'écran.
    Sprites(TypeDecor).Afficher X, Y, SurfBack, True
End Sub
