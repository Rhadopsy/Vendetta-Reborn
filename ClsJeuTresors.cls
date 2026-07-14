VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsJeuTresors"
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

Public X As Long
Public Y As Long

Private Fortune As Double 'Renvoie la valeur du trésor
Public IndiceImage As Integer ' (utilisé pour afficher l'image)

Private Argent As Double
Private Ressources() As Double
Private Objets As Collection 'Contient les indices des objets.

Public Sub Init(ByVal Parametres As ClsJeuParam)
    Argent = 0
    ReDim Ressources(Parametres.NombreRessources - 1)
    Set Objets = New Collection
End Sub

Public Sub Ajouter_Objet(ByVal IndiceObjet As Integer, _
                         ByVal Parametres As ClsJeuParam)
    'Ajoute un objet dans le trésor.
    Objets.Add IndiceObjet
    Definir_Fortune = Definir_Fortune + Parametres.Objet_PrixVente(IndiceObjet)
End Sub

Public Sub Rammasser_Tresor(ByVal Parametres As ClsJeuParam, _
                            ByVal Commentaires As ClsJeuCommentaires, _
                            ByRef Perso As ClsJeuPerso)
    'Un personnage récupère le trésor.
    Dim i As Integer
    Perso.Definir_Argent = Perso.Definir_Argent + Argent
    For i = 0 To UBound(Ressources())
        Perso.Definir_Ressources(i) = Perso.Definir_Ressources(i) + Ressources(i)
    Next i
    For i = 1 To Objets.Count
        Perso.Objet_Ajouter_Inventaire Objets(i), Parametres, Commentaires, True
    Next i
End Sub

Private Sub Calculer_IndiceImage()
    Dim i As Integer
    With Parametres
    For i = .NombreTresors - 1 To 0 Step -1
        If Fortune >= Parametres.Tresor_ValeurMin(i) Then
            IndiceImage = i
            Exit For
        End If
    Next i
    End With
End Sub

Public Property Get Nombre_Objets() As Integer
    Nombre_Objets = Objets.Count
End Property

Public Property Get Nom(ByVal Parametres As ClsJeuParam)
    Nom = Parametres.Tresor_Nom(IndiceImage)
End Property

Public Property Get Vide() As Boolean
    Dim i As Long
    Vide = Argent <= 1 And Objets.Count < 1
    If Vide Then
        For i = 0 To UBound(Ressources())
            If Ressources(i) >= 1 Then
                Vide = False
                Exit For
            End If
        Next i
    End If
End Property

' ************************************************************************************
' *                                       Propriétés                                 *
' ************************************************************************************
Public Property Get Definir_Fortune() As Double
    Definir_Fortune = Fortune
End Property
Public Property Let Definir_Fortune(ByVal Valeur As Double)
    Fortune = Valeur
    Calculer_IndiceImage
End Property

Public Property Get Definir_Argent() As Double
    Definir_Argent = Argent
End Property
Public Property Let Definir_Argent(ByVal Valeur As Double)
    Definir_Fortune = Definir_Fortune + Valeur - Argent
    Argent = Valeur
End Property

Public Property Get Definir_Ressources(ByVal Index As Integer) As Double
    Definir_Ressources = Ressources(Index)
End Property
Public Property Let Definir_Ressources(ByVal Index As Integer, _
                                       ByVal Valeur As Double)
    Definir_Fortune = Definir_Fortune + (Valeur - Ressources(Index)) * Parametres.Ressources_PrixVente(Index)
    Ressources(Index) = Valeur
End Property

Public Property Get Definir_Objet(ByVal Index As Integer) As Integer
    Definir_Objet = Objets(Index)
End Property
