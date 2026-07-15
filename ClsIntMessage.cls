VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsIntMessage"
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

Const CaractereRetourLigne = "#"
Const CaractereCouleurSpeciale1 = "&"
Const CouleurSpeciale1 = &HFF& '&H80& '&H8080FF   '&HFF&
Const CaractereCouleurSpeciale2 = "$"
Const CouleurSpeciale2 = &HFF0000

Const InterligneX = 0
Const InterligneY = 16

Const LargeurCaractere = 7
Const HauteurCaractere = 16

Public Sub Afficher(ByVal X As Long, _
                    ByVal Y As Long, _
                    ByVal Message As String, _
                    ByRef SurfBack As IDirectDrawSurface7, _
                    Optional ByVal Couleur As Long, _
                    Optional ByVal ChangerCouleur As Boolean)
    If Message <> "" Then
        If Not ChangerCouleur Then
            Couleur = SurfBack.GetForeColor
        End If
        Select Case Left(Message, 1)
        Case CaractereCouleurSpeciale1:
            Message = Right(Message, Len(Message) - 1)
            SurfBack.SetForeColor CouleurSpeciale1
        Case CaractereCouleurSpeciale2:
            Message = Right(Message, Len(Message) - 1)
            SurfBack.SetForeColor CouleurSpeciale2
        Case Else
            SurfBack.SetForeColor Couleur
        End Select
        If InStr(Message, Chr(13)) Then
            'Gère des lignes multiples.
            SurfBack.DrawText X, Y, Left(Message, InStr(Message, Chr(13)) - 1), False
            Me.Afficher X + InterligneX, _
                        Y + InterligneY, _
                        Right(Message, Len(Message) - InStr(Message, Chr(13)) - 0), SurfBack, Couleur, True
        ElseIf InStr(Message, CaractereRetourLigne) Then
            SurfBack.DrawText X, Y, Left(Message, InStr(Message, CaractereRetourLigne) - 1), False
            Me.Afficher X + InterligneX, _
                        Y + InterligneY, _
                        Right(Message, Len(Message) - InStr(Message, CaractereRetourLigne) - 0), SurfBack, Couleur, True
        Else
            SurfBack.DrawText X, Y, Message, False
        End If
        SurfBack.SetForeColor Couleur
    End If
End Sub

Public Function Nombre_Lignes(ByVal Chaine As String) As Integer
    Dim i As Integer
    For i = 1 To Len(Chaine)
        If Mid(Chaine, i, 1) = Caractere_RetourLigne Then
            Nombre_Lignes = Nombre_Lignes + 1
        End If
    Next i
    Nombre_Lignes = Nombre_Lignes + 1
End Function

Public Property Get Largeur_Caractere() As Integer
    Largeur_Caractere = LargeurCaractere
End Property
Public Property Get Hauteur_Caractere() As Integer
    Hauteur_Caractere = HauteurCaractere
End Property
Public Property Get Caractere_RetourLigne() As String
    Caractere_RetourLigne = CaractereRetourLigne
End Property

