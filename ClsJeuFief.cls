VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsJeuFief"
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

'Paramètres de restauration du chateau.
Const QuantiteVieParArgent = 0.2 '5
Const VitesseTransaction = 0.02 / 4 / 4

Private Numero As Integer
Public Nom As String

Public NumeroEquipe As Long

Private TypePeuple As Integer

Public Epoque As Integer
Public ChangerEpoque As Boolean 'Indique s'il on est en train de changer d'époque ou pas.
Public TempsChangementEpoque As Long 'Durée nécessaire avant de passer à la prochaine époque.

Public NombreCitoyens As Long
Public NombreMorts As Long
Public Empereur As Long
'Public IndicePremierCitoyen As Long
'Public IndiceDernierCitoyen As Long

Public Argent As Double
Private Ressources() As Double
Public Marches As Collection 'Mémorise les indices des marchés de ce fief.

Private Sub Class_Initialize()
    Set Marches = New Collection
End Sub

Public Sub Init(ByVal i As Long, ByVal Parametres As ClsJeuParam)
    Numero = i
    ReDim Ressources(Parametres.NombreRessources - 1)
    Empereur_Changer -1
End Sub

Public Sub Changer_Epoque(ByVal Chateau As ClsJeuBatiment, _
                          ByVal Parametres As ClsJeuParam)
    'Fonction appelée à chaque changement d'époque.
    With Parametres
    If Not ChangerEpoque And Epoque < .NombreEpoques - 1 Then
        ChangerEpoque = True
        Chateau.MaxVie = Chateau.MaxVie + .Epoque_BonusVieChateau(Epoque + 1)
        TempsChangementEpoque = .Epoque_Duree(Epoque + 1) * .VitesseEpoque_Coef
        Chateau.RecuperationVie = Chateau.RecuperationVie + .Epoque_BonusVieChateau(Epoque + 1) + 1
    End If
    End With
End Sub

Public Sub Transaction_Ressources(ByVal Parametres As ClsJeuParam)
    Dim i As Integer
    Dim IndiceMin As Integer
    Dim IndiceMax As Integer
    For i = 0 To UBound(Ressources())
        If Int(Ressources(i)) < Int(Ressources(IndiceMin)) And _
           Parametres.Ressources_Epoque(i) <= Epoque Then
            IndiceMin = i
        End If
    Next i
    For i = 0 To UBound(Ressources())
        If Int(Ressources(i)) > Int(Ressources(IndiceMax)) Then
            IndiceMax = i
        End If
    Next i
    If Abs(Ressources(IndiceMax) - Ressources(IndiceMin)) > 1 Then
        Ressources(IndiceMin) = Ressources(IndiceMin) + VitesseTransaction * NombreCitoyens  '*Parametres.Ressources_PrixVente(
        Ressources(IndiceMax) = Ressources(IndiceMax) - VitesseTransaction * NombreCitoyens
    End If
End Sub

Public Sub Restaurer_Chateau(ByVal Quantite As Double, _
                             ByRef Perso As ClsJeuPerso, _
                             Optional RestaurerMax As Boolean)
    If RestaurerMax Or _
       Quantite > Perso.Definir_Argent Then
        Quantite = Perso.Definir_Argent
    End If
    Argent = Argent + Quantite
    Perso.Definir_Argent = Perso.Definir_Argent - Quantite
End Sub

Public Sub Empereur_Changer(ByVal i As Long, Optional ByVal Payer As Boolean)
    Dim j As Long
    If i >= 0 Then
        For j = 0 To LPerso
            If Persos(j).NumeroFief = Numero Then
                If i = j Then
                    'Devient l'empereur
                    Persos(j).Definir_Empereur = True
                ElseIf Persos(j).Empereur Then
                    'N'est plus l'empereur.
                    Persos(j).Definir_Empereur = False
                End If
            End If
        Next j
        If Payer Then
            Persos(i).Definir_Argent = Persos(i).Definir_Argent - Empereur_Prix(Parametres)
        End If
    End If
    Empereur = i
End Sub

Public Function Empereur_Prix(ByVal Parametres As ClsJeuParam) As Long
    'Renvoie le prix nécessaire pour devenir empereur.
    Empereur_Prix = Parametres.Epoque_EmpereurPrix(Epoque)
End Function

Public Sub Restauration(ByVal Chateau As ClsJeuBatiment)
    Dim Quantite As Double
    If Chateau.Definir_Vie < Chateau.MaxVie Then
        Quantite = Chateau.MaxVie - Chateau.Definir_Vie + 1 'Quantité nécessaire pour restaurer entièrement le chateau.
        If Quantite > Argent * QuantiteVieParArgent Then
            Quantite = Argent * QuantiteVieParArgent
        End If
        Chateau.RecuperationVie = Chateau.RecuperationVie + Quantite
        Argent = Argent - Quantite / QuantiteVieParArgent
    End If
    If Epoque < Parametres.NombreEpoques - 1 Then
        TempsChangementEpoque = TempsChangementEpoque - 1 '* 20
        'If Chateau.RecuperationVie = 0 And Chateau.Visible And ChangerEpoque Then
        If TempsChangementEpoque <= 0 And ChangerEpoque Then
            Epoque = Epoque + 1
            ChangerEpoque = False
        End If
    End If
End Sub

Public Sub Vendre(ByVal IndiceRessource As Long, _
                  ByVal Quantite As Double, _
                  ByRef Perso As ClsJeuPerso, _
                  ByVal Parametres As ClsJeuParam, _
                  Optional VendreMax As Boolean)
    If VendreMax Or _
       Quantite > Perso.Definir_ressources(IndiceRessource) Then
        Quantite = Perso.Definir_ressources(IndiceRessource)
    End If
    Quantite = Int(Quantite)
    Perso.Definir_ressources(IndiceRessource) = Perso.Definir_ressources(IndiceRessource) - Quantite
    Ressources(IndiceRessource) = Ressources(IndiceRessource) + Quantite
    Perso.Definir_Argent = Perso.Definir_Argent + Quantite * Parametres.Ressources_PrixVente(IndiceRessource)
End Sub

Public Sub Vendre_ObjetsSelectionnees(ByRef Perso As ClsJeuPerso, _
                                      ByVal Parametres As ClsJeuParam)
    Dim i As Integer
    If Perso.ArgentSelectionne > 0 Then
        Restaurer_Chateau Perso.ArgentSelectionne, Perso
    End If
    For i = 0 To Parametres.NombreRessources - 1
        If Perso.Definir_RessourcesSelectionnes(i) > 0 Then
            Vendre i, Perso.Definir_RessourcesSelectionnes(i), Perso, Parametres
        End If
    Next i
    If Perso.ObjetSelectionne >= 0 Then
        If Not Parametres.Objet_Inseparable(Perso.ObjetSelectionne) Then
            Perso.Objet_Vendre_Chateau Perso.ObjetSelectionne, Fiefs(Perso.NumeroFief), Parametres, False
            Perso.ObjetSelectionne = -1
        End If
    End If
    Perso.Definir_ObjetSelectionne = False
End Sub

Public Sub Acheter(ByVal IndiceRessource As Long, _
                   ByVal Quantite As Double, _
                   ByRef Perso As ClsJeuPerso, _
                   ByVal Parametres As ClsJeuParam, _
                   Optional AcheterMax As Boolean)
    If AcheterMax Or _
       Perso.Definir_Argent < Quantite * Parametres.Ressources_PrixVente(IndiceRessource) * Parametres.CoefficientAchat Then
        Quantite = Perso.Definir_Argent / (Parametres.Ressources_PrixVente(IndiceRessource) * Parametres.CoefficientAchat)
    End If
    If Quantite > Ressources(IndiceRessource) Then
        Quantite = Ressources(IndiceRessource)
    End If
    If Quantite > Perso.Definir_MaxRessources(IndiceRessource) - Int(Perso.Definir_ressources(IndiceRessource)) Then
        Quantite = Perso.Definir_MaxRessources(IndiceRessource) - Int(Perso.Definir_ressources(IndiceRessource))
    End If
    Quantite = Int(Quantite)
    Perso.Definir_ressources(IndiceRessource) = Perso.Definir_ressources(IndiceRessource) + Quantite
    Me.Definir_ressources(IndiceRessource) = Me.Definir_ressources(IndiceRessource) - Quantite
    Perso.Definir_Argent = Perso.Definir_Argent - Quantite * Parametres.Ressources_PrixVente(IndiceRessource) * Parametres.CoefficientAchat
End Sub

Public Property Get Definir_ressources(ByVal Index As Long) As Double
    If Index >= 0 Then
        Definir_ressources = Ressources(Index)
    Else
        Definir_ressources = 0
    End If
End Property

Public Property Let Definir_ressources(ByVal Index As Long, ByVal Valeur As Double)
    Ressources(Index) = Valeur
    If Ressources(Index) < 0 Then
        Ressources(Index) = 0
    End If
End Property


Public Property Get Definir_TypePeuple() As Integer
    Definir_TypePeuple = TypePeuple
End Property
Public Property Let Definir_TypePeuple(ByVal Valeur As Integer)
    TypePeuple = Valeur
    Nom = Parametres.Peuples_Nom(TypePeuple)
End Property


'Paramètres des varibles temporaires utilisées pour les marchés.
Public Sub Marches_Actualiser()
    Dim i As Long
    For i = 0 To UBound(Maisons())
        If Maisons(i).Marche And Maisons(i).Construit Then
            If Persos(i).NumeroEquipe = NumeroEquipe Then
                Marches_Ajouter i
            End If
        End If
    Next i
End Sub
Public Sub Marches_Ajouter(ByVal Indice As Long)
    Dim i  As Long
    For i = 0 To UBound(Fiefs())
        If Fiefs(i).NumeroEquipe = NumeroEquipe Then
            Marches.Add Indice
        End If
    Next i
End Sub
Public Sub Marches_Supprimer(ByVal Indice As Long)
    Dim i As Long: Dim j As Long
    For i = 0 To UBound(Fiefs())
        If Fiefs(i).NumeroEquipe = NumeroEquipe Then
            For j = 1 To Fiefs(i).Marches.Count
                If Fiefs(i).Marches(j) = Indice Then
                    Fiefs(i).Marches.Remove j
                    Exit For
                End If
            Next j
        End If
    Next i
End Sub
