Attribute VB_Name = "ModTresors"
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

Public Tresors As Collection

Public Sub Tresors_Init()
    Set Tresors = New Collection
    While Tresors.Count > 0
        Tresors.Remove 1
    Wend
End Sub

Public Function Tresor_Creer(ByVal X As Long, ByVal Y As Long) As Integer
    'Renvoie l'indice du trésor créé.
    Dim TempTresor As ClsJeuTresors
    Set TempTresor = New ClsJeuTresors
    TempTresor.Init Parametres
    TempTresor.X = X
    TempTresor.Y = Y
    Tresors.Add TempTresor
    Tresor_Creer = Tresors.Count
    If ComReseau.Serveur Then ComMessages.Tresor_Ajouter X, Y, ComReseau.PlayerID, True
End Function

Public Function Tresors_Verifier(ByVal IndiceTresor As Long)
    'Vérifie si le trésor passé en paramètre se supperpose avec un autre.
    Dim i As Long
    Dim j As Long
    If Tresors(IndiceTresor).Vide Then
        Tresor_Effacer IndiceTresor
    Else
        For i = 1 To Tresors.Count
            If i <> IndiceTresor Then
                If Est_Superpose(Tresors(i).X, Tresors(i).Y, _
                                 Tresors(IndiceTresor).X, Tresors(IndiceTresor).Y, _
                                 AffTresor.Largeur, AffTresor.Hauteur, _
                                 AffTresor.Largeur, AffTresor.Hauteur) Then
                    Tresors(i).Definir_Argent = Tresors(i).Definir_Argent + Tresors(IndiceTresor).Definir_Argent
                    For j = 0 To Parametres.NombreRessources - 1
                        Tresors(i).Definir_ressources(j) = Tresors(i).Definir_ressources(j) + Tresors(IndiceTresor).Definir_ressources(j)
                    Next j
                    For j = 1 To Tresors(IndiceTresor).Nombre_Objets
                        Tresors(i).Ajouter_Objet Tresors(IndiceTresor).Definir_Objet(j), Parametres
                    Next j
                    Tresor_Effacer IndiceTresor
                    If ComReseau.Serveur Then ComMessages.Tresor_Apparence i, ComReseau.PlayerID, True
                    Exit For
                End If
            End If
        Next i
    End If
End Function

'Public Sub Tresor_Deposer_Objet(ByVal Perso As ClsJeuPerso, _
'                                ByVal TypeObjet As Integer)
'    Dim TempTresor As ClsJeuTresors
'    Set TempTresor = New ClsJeuTresors
'    TempTresor.Init Parametres
'    TempTresor.Ajouter_Objet TypeObjet, Parametres
'    TempTresor.X = Perso.PositionX
'    TempTresor.Y = Perso.PositionY
'    Tresors.Add TempTresor
'End Sub

Public Sub Tresor_Ramasser(ByVal IndiceTresor As Integer, _
                           ByRef Perso As ClsJeuPerso)
    
    Tresors(IndiceTresor).Rammasser_Tresor Parametres, Commentaires, Perso
    Tresor_Detruire IndiceTresor
End Sub

Private Sub Tresor_Detruire(ByVal IndiceTresor As Integer)
    Dim i As Long
    Tresor_Effacer IndiceTresor
    For i = 0 To UBound(Persos())
        With Persos(i)
        If .Action = 204 Then
            Select Case .IndiceTresor
            Case IndiceTresor: 'Arrêtent tous les personnages qui se dirigeaiant vers ce trésor.
                .Arreter
            Case Is > IndiceTresor: 'Réactualise les indices des autres personnages.
                .IndiceTresor = .IndiceTresor - 1
            End Select
        End If
        End With
    Next i
End Sub

Public Sub Tresor_Effacer(ByVal IndiceTresor As Integer)
    Tresors.Remove IndiceTresor
    If ComReseau.Serveur Then ComMessages.Tresor_Supprimer IndiceTresor, ComReseau.PlayerID, True
End Sub

Public Sub Tresor_Generer()
    'Place un trésor aléatoirement dans le monde.
    Dim i As Long
    Dim j As Long
    Dim k As Long
    If (Rnd < 1 / 60 Or Tresors.Count = 0) And Tresors.Count <= 10 Then
        With Parametres
        'Ajoute un trésor.
        i = Tresor_Creer(Rnd * Monde.Largeur * Monde.LargeurCase - AffTresor.Largeur, Rnd * Monde.Hauteur * Monde.HauteurCase - AffTresor.Hauteur)
        'Ajoute de l'argent.
        If Rnd < 1 / 3 Then
            Tresors(i).Definir_Argent = Rnd * 1000
        End If
        'Ajoute des ressources.
        For j = 0 To Rnd * 6 + 1
            If Rnd < 1 / 6 Then
                Tresors(i).Definir_ressources(Rnd * (.NombreRessources - 1)) = Rnd * 20 + 1
            End If
        Next j
        'Ajoute des objets.
        For j = 0 To Rnd * 6 + 1
            k = Rnd * (.NombreObjets - 1)
            If Not .Objet_Secret(k) Then
                Tresors(i).Ajouter_Objet k, Parametres
            End If
        Next j
        If Tresors(i).Vide Then Tresors(i).Definir_Argent = Rnd * 10000
        If ComReseau.Serveur Then ComMessages.Tresor_Apparence i, ComReseau.PlayerID, True
        End With
    End If
End Sub

Public Property Get Tresor_DefinirNombre() As Long
    Tresor_DefinirNombre = Tresors.Count
End Property
Public Property Let Tresor_DefinirNombre(ByVal Valeur As Long)
    Dim TempTresor As ClsJeuTresors
    Set TempTresor = New ClsJeuTresors
    TempTresor.Init Parametres
    If Tresors.Count > Valeur Then
        Do
            Tresors.Remove 1
        Loop Until Tresors.Count = Valeur
    ElseIf Tresors.Count < Valeur Then
        Do
            Tresors.Add TempTresor
        Loop Until Tresors.Count = Valeur
    End If
End Property
