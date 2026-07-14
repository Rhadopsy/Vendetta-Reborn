Attribute VB_Name = "ModDecors"
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

Public Decors As Collection

Public Sub Decors_Init()
    Set Decors = New Collection
    While Decors.Count > 0
        Decors.Remove 1
    Wend
End Sub

Public Sub Decor_Creer(ByVal Categorie As Integer, _
                       ByVal X As Long, _
                       ByVal Y As Long, _
                       Optional ByVal Nom As String)
    'Renvoie l'indice du décor créé.
    Dim TempDecor As ClsJeuDecors
    Set TempDecor = New ClsJeuDecors
    'If Nom = "" Then
    'TempDecor.Nom = Parametres.Decor_Nom(Categorie)
    'Else
    TempDecor.Nom = Nom
    'End If
    TempDecor.Categorie = Categorie
    TempDecor.X = X
    TempDecor.Y = Y
    TempDecor.Largeur = Parametres.Decor_Largeur(Categorie)
    TempDecor.Hauteur = Parametres.Decor_Hauteur(Categorie)
    TempDecor.NombreEtapeAnimation = Parametres.Decor_NombreEtapeAnimation(Categorie)
    TempDecor.Hauteur = Parametres.Decor_Hauteur(Categorie)
    TempDecor.Temporaire = Parametres.Decor_Duree(Categorie) > 0
    TempDecor.Duree = Parametres.Decor_Duree(Categorie) * Jeu.Definir_Vitesse
    Decors.Add TempDecor
    If ComReseau.Serveur Then ComMessages.Decor_Ajouter Categorie, X, Y, Nom, ComReseau.PlayerID, True
End Sub

Public Sub Actualiser_Affichage_Decors()
    Dim i As Long
    'Gère l'animation.
    For i = 1 To Decors.Count
        Decors(i).ChangerEtapeAnimation
    Next i
    
    'Supprime les décors temporaires.
    For i = 1 To Decors.Count
        If Decors(i).Temporaire Then
            Decors(i).Duree = Decors(i).Duree - 1
            If Decors(i).Duree <= 0 Then
                Decors.Remove i
                If ComReseau.Serveur Then ComMessages.Decor_Supprimer i, ComReseau.PlayerID, True
                Exit For
            End If
        End If
    Next i
End Sub

Public Property Get Decor_DefinirNombre() As Long
    Decor_DefinirNombre = Decors.Count
End Property
Public Property Let Decor_DefinirNombre(ByVal Valeur As Long)
    Dim TempDecor As ClsJeuDecors
    Set TempDecor = New ClsJeuDecors
    If Decors.Count > Valeur Then
        Do
            Decors.Remove 1
        Loop Until Decors.Count = Valeur
    ElseIf Decors.Count < Valeur Then
        Do
            Decors.Add TempDecor
        Loop Until Decors.Count = Valeur
    End If
End Property
