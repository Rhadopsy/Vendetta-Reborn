VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAffEffet"
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

Const Chemin = "Effets\"
Const Transparent = True

Private Sprites() As ClsAffSprite

Private Largeur() As Integer
Private Hauteur() As Integer
Private NombreEtapes() As Integer
Private DureeEtape() As Double
Private Son() As String

Public Sub Charger(ByRef dd As DirectDraw7, _
                   ByRef Parametres As ClsJeuParam)
    Dim i As Long
    ReDim Sprites(Parametres.NombreEffets - 1)
    ReDim Largeur(UBound(Sprites()))
    ReDim Hauteur(UBound(Sprites()))
    ReDim NombreEtapes(UBound(Sprites()))
    ReDim DureeEtape(UBound(Sprites()))
    ReDim Son(UBound(Sprites()))
    For i = 0 To UBound(Sprites())
        Set Sprites(i) = New ClsAffSprite
        If Parametres.Effet_Fichier(i) <> "" And Parametres.Effet_NombreEtapes(i) > 0 Then
            If Left(Parametres.Effet_Fichier(i), 1) = "." Then
                Sprites(i).Charger dd, Right(Parametres.Effet_Fichier(i), Len(Parametres.Effet_Fichier(i)) - 1), True
            Else
                Sprites(i).Charger dd, Chemin & Parametres.Effet_Fichier(i)
            End If
            NombreEtapes(i) = Parametres.Effet_NombreEtapes(i)
        Else
            NombreEtapes(i) = 0
        End If
        Largeur(i) = Parametres.Effet_Largeur(i)
        Hauteur(i) = Parametres.Effet_Hauteur(i)
        
        DureeEtape(i) = Parametres.Effet_DureeEtape(i)
        Son(i) = Parametres.Effet_Son(i)
    Next i
End Sub

Public Sub Afficher(ByVal TypeEffet As Long, _
                    ByVal Etape As Integer, _
                    ByVal X As Long, _
                    ByVal Y As Long, _
                    ByRef SurfBack As DirectDrawSurface7)
    If NombreEtapes(TypeEffet) > 0 Then
        'On charge la bonne image en fonction des paramètres.
        Sprites(TypeEffet).Selection 0, _
                                     Hauteur(TypeEffet), _
                                     Largeur(TypeEffet) * Etape, _
                                     Largeur(TypeEffet) * (Etape + 1)
        'On affiche le terrain sur l'arrière plan.
        Sprites(TypeEffet).Afficher X, Y, SurfBack, Transparent
    End If
End Sub


Public Property Get Definir_Largeur(ByVal Indice As Integer) As Integer
    Definir_Largeur = Largeur(Indice)
End Property
Public Property Get Definir_Hauteur(ByVal Indice As Integer) As Integer
    Definir_Hauteur = Hauteur(Indice)
End Property
