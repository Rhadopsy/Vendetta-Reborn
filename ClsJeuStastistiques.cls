VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsJeuStastistiques"
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

Public Afficher As Boolean

Public NombrePersos As Integer
Private PersosStats() As ClsJeuPerso
Private PersosNumero() As Long 'Mémorise les indices des personnages concernés.

Private Sub Class_Initialize()
    Afficher = True
End Sub

Public Sub Init(ByVal Parametres As ClsJeuParam)
    Dim i As Long, j As Long
    If Afficher Then
        NombrePersos = 0
        For i = 0 To UBound(Persos())
            If Persos(i).Joueur > 0 Then
                NombrePersos = NombrePersos + 1
            End If
        Next i
        ReDim PersosStats(NombrePersos - 1)
        ReDim PersosNumero(NombrePersos - 1)
        For i = 0 To UBound(Persos())
            If Persos(i).Joueur > 0 Then
                Set PersosStats(j) = New ClsJeuPerso
                PersosStats(j).Init j, Parametres
                PersosNumero(j) = i
                Charger_Sauvegarde_Perso Persos(i).Nom, PersosStats(j)
                j = j + 1
            End If
        Next i
    End If
End Sub

Public Sub Fin_Partie()
    'A lancer en fin de partie pour calculer les stastistiques.
    Dim i As Integer
    If Afficher Then
        For i = 0 To NombrePersos - 1
            Persos(PersosNumero(i)).EffetTemp_Effacer_Tout
            Persos(PersosNumero(i)).Objet_Inventaire_Vider_Tout
            Persos(PersosNumero(i)).Objet_Equipes_Vider_Tout
            Persos(PersosNumero(i)).Actualiser_Bonus
        Next i
    End If
End Sub

Public Function Nom(ByVal i As Integer) As String
    Nom = Persos(PersosNumero(i)).Nom
End Function

Public Function Frags(ByVal i As Integer, ByVal Categorie As Integer) As Long
    'Frags = Persos(PersosNumero(i)).Definir_Bestiaire_Total - PersosStats(i).Definir_Bestiaire_Total
    Select Case Categorie
    Case 0:
        Frags = Persos(PersosNumero(i)).FragsPersos
    Case 1:
        Frags = Persos(PersosNumero(i)).FragsBatiments
    Case 2:
        Frags = Persos(PersosNumero(i)).FragsChateaux
    End Select
End Function
Public Function Morts(ByVal i As Integer) As Long
    Morts = Persos(PersosNumero(i)).Morts
End Function

Public Function Gain_Niveau(ByVal i As Integer) As Double
    Gain_Niveau = Persos(PersosNumero(i)).Niveau - PersosStats(i).Niveau
End Function
Public Function Gain_Tresor(ByVal i As Integer) As Integer
    Gain_Tresor = Persos(PersosNumero(i)).Defnir_Score_Tresor - PersosStats(i).Defnir_Score_Tresor
End Function

Public Function Gain_Vie(ByVal i As Integer) As Integer
    Gain_Vie = Persos(PersosNumero(i)).Definir_MaxVie - PersosStats(i).Definir_MaxVie
End Function
Public Function Gain_Energie(ByVal i As Integer) As Integer
    Gain_Energie = Persos(PersosNumero(i)).Definir_MaxEnergie - PersosStats(i).Definir_MaxEnergie
End Function
Public Function Gain_Magie(ByVal i As Integer) As Integer
    Gain_Magie = Persos(PersosNumero(i)).Definir_MaxMagie - PersosStats(i).Definir_MaxMagie
End Function
Public Function Gain_Moral(ByVal i As Integer) As Integer
    Gain_Moral = Persos(PersosNumero(i)).Definir_MaxMoral - PersosStats(i).Definir_MaxMoral
End Function
Public Function Gain_Attaque(ByVal i As Integer) As Integer
    Gain_Attaque = Persos(PersosNumero(i)).Definir_MaxAttaque - PersosStats(i).Definir_MaxAttaque
End Function
Public Function Gain_Defense(ByVal i As Integer) As Integer
    Gain_Defense = Persos(PersosNumero(i)).Definir_MaxDefense - PersosStats(i).Definir_MaxDefense
End Function

Public Function Gain_CompetenceSpeciales(ByVal i As Integer, ByVal j As Integer) As Integer
    Gain_CompetenceSpeciales = Persos(PersosNumero(i)).Definir_Niveau_CompetenceSpeciales(j) - PersosStats(i).Definir_Niveau_CompetenceSpeciales(j)
End Function
Public Function Gain_CompetenceRessources(ByVal i As Integer, ByVal j As Integer) As Integer
    Gain_CompetenceRessources = Persos(PersosNumero(i)).Definir_Niveau_CompetenceRessources(j) - PersosStats(i).Definir_Niveau_CompetenceRessources(j)
End Function
Public Function Gain_CompetenceServices(ByVal i As Integer, ByVal j As Integer) As Integer
    Gain_CompetenceServices = Persos(PersosNumero(i)).Definir_Niveau_CompetenceServices(j) - PersosStats(i).Definir_Niveau_CompetenceServices(j)
End Function
Public Function Gain_CompetenceObjets(ByVal i As Integer, ByVal j As Integer) As Integer
    Gain_CompetenceObjets = Persos(PersosNumero(i)).Definir_Niveau_CompetenceObjets(j) - PersosStats(i).Definir_Niveau_CompetenceObjets(j)
End Function

Public Function Gain_Bestiaire(ByVal i As Integer, ByVal j As Integer) As Long
    Gain_Bestiaire = Persos(PersosNumero(i)).Definir_Bestiaire(j) - PersosStats(i).Definir_Bestiaire(j)
End Function
