VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAffBatiment"
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

Const FichierChateaux = "Chateaux-"

Private Sprites() As ClsAffSprite

Public Sub Charger(ByVal Categorie As Integer, _
                   ByRef dd As DirectDraw7, _
                   ByVal Parametres As ClsJeuParam)
    Dim i As Integer
    With Parametres
    If Categorie = 0 Then
        'Correspond aux châteaux.
        ReDim Sprites(.NombrePeuples - 1)
    Else
        'Correspond aux bâtiments.
        ReDim Sprites(.NombreTypeBatiments - 1)
    End If
    
    For i = 0 To UBound(Sprites())
        Set Sprites(i) = New ClsAffSprite
        If Categorie = 0 Then
            'If Parametres.Peuples_CheminImage(i) <> "" Then
            Sprites(i).Charger dd, .Peuples_CheminImage(i) & .Peuples_Image(i), True
            'Else
            '    Sprites(i).Charger dd, Chemin_Chateau(i, Parametres)
            'End If
        Else
            'Sprites(i).Charger dd, Chemin_Maison(i, Parametres)
            'Sprites(i).Charger dd, .Batiment_CheminImage(i) & .Batiment_Image(i), True
            If Left(.Batiment_CheminImage(i), 1) = "." Then
                Sprites(i).Charger dd, Right(.Batiment_CheminImage(i) & .Batiment_Image(i), Len(.Batiment_CheminImage(i) & .Batiment_Image(i)) - 2), True
            Else
                Sprites(i).Charger dd, .Batiment_CheminImage(i) & .Batiment_Image(i), True
            End If
        End If
    Next i
    End With
End Sub

Public Sub Afficher(ByVal X As Long, _
                    ByVal Y As Long, _
                    ByVal TypeBatiment As Long, _
                    ByVal Etat As Long, _
                    ByRef SurfBack As DirectDrawSurface7, _
                    ByVal Largeur As Long, _
                    ByVal Hauteur As Long, _
                    Optional ByVal EtapeAnimation As Integer = 0)
    'On charge la bonne image en fonction des paramètres.
    Sprites(TypeBatiment).Selection Etat * Hauteur, _
                                   (Etat + 1) * Hauteur, _
                                    Largeur * (EtapeAnimation), _
                                    Largeur * (EtapeAnimation + 1)
    'On affiche le batiment sur l'écran.
    Sprites(TypeBatiment).Afficher X, Y, SurfBack, True
End Sub

'Public Property Get Chemin_Chateau(ByVal Index As Integer, _
'                                   ByVal Parametres As ClsJeuParam) As String
'    Chemin_Chateau = Chemin & Parametres.Peuples_Image(Index)
'End Property
'
'Public Property Get Chemin_Maison(ByVal Index As Integer, _
'                                  ByVal Parametres As ClsJeuParam) As String
'    Chemin_Maison = Chemin & Parametres.Batiment_Image(Index)
'End Property
