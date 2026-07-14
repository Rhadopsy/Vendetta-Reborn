VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsJeuIA2"
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

Const MessageEquiper = "Equiper_"
Const CaractereCommentaire = "#"
Const CaractereQuantite = "="
Const MessageAge = "Age_"
Const MessageRessource = "Ressource_"
Const MessageBatiment = "Batiment_"
Const MessageObjet = "Objet_"
Const MessageTravailler = "Travailler"
Const MessagePatrouiller = "Patrouiller"
Const MessageAttaquerEnnemi = "Attaquer Ennemi"
Const MessageAttaquerDonjon = "Attaquer Donjon"
Const MessageChangerMode = "ChangerMode_"
Const MessageGagnerArgent = "GagnerArgent_"
Const MessageGagnerArgentRessource = "GagnerArgentRessource_"
Const MessageFin = "#Fin"

Const MessageDefaut = MessageTravailler

Private Type Etape
    Categorie As Integer '0: ressources, 1 batiment, 2 objet, 3 travailler, 4 patrouiller, 5 attaquer monstre, 6, attaquer donjon, 7, Changer Mode, 8 Gagner argent.
    Valeur As Integer
    Quantite As Integer
End Type

Private Type Epoque
    Numero As String
    Commentaire As String
    Etapes() As Etape
    BatimentDefini As Boolean
End Type

Public Nom As String
Public Joueur As Boolean 'Vrai : Le joueur peut éditer et utiliser cette IA. Se trouvent dans la racine du répertoire des Ias.
Public Equiper As Integer
Public Commentaire As String
Private Epoques() As Epoque

Public Sub Charger_IA(ByVal CheminFichier As String, _
                      ByVal NomFichier As String, _
                      ByVal Parametres As ClsJeuParam, _
                      ByVal EstJoueur As Boolean)
    On Error GoTo Erreur
    Dim Temp As String
    Dim i As Integer, j As Integer
    
    Open CheminFichier & NomFichier For Input As #1
    Nom = NomFichier
    Joueur = EstJoueur
    
    ReDim Epoques(Parametres.NombreEpoques - 1)
        
    'Commentaire général.
    Line Input #1, Temp
    If Left(Temp, Len(CaractereCommentaire)) = CaractereCommentaire Then
        Commentaire = Right(Temp, Len(Temp) - Len(CaractereCommentaire))
        Line Input #1, Temp
    End If
    
    'Informations sur la manière de s'équiper.
    If Left(Temp, Len(MessageEquiper)) = MessageEquiper Then
        Equiper = Right(Temp, Len(Temp) - Len(MessageEquiper))
        Line Input #1, Temp
    End If
    
    While Not EOF(1)
        If Left(Temp, Len(MessageAge)) = MessageAge Then
            'On change d'age.
            i = Right(Temp, Len(Temp) - Len(MessageAge)) - 1
            j = 0
        ElseIf Left(Temp, Len(CaractereCommentaire)) = CaractereCommentaire Then
            'C'est un commentaire.
            Epoques(i).Commentaire = Right(Temp, Len(Temp) - Len(CaractereCommentaire))
        ElseIf Left(Temp, Len(MessageRessource)) = MessageRessource Then
            'Ressources.
            Ajouter_Etape i, j
            Epoques(i).Etapes(j).Categorie = 0
            Epoques(i).Etapes(j).Valeur = Mid(Temp, Len(MessageRessource) + 1, InStr(Temp, CaractereQuantite) - Len(MessageRessource) - 1) - 1
            Epoques(i).Etapes(j).Quantite = Mid(Temp, InStr(Temp, CaractereQuantite) + 1, Len(Temp))
            j = j + 1
        ElseIf Left(Temp, Len(MessageBatiment)) = MessageBatiment Then
            'Batiment.
            Ajouter_Etape i, j
            Epoques(i).Etapes(j).Categorie = 1
            Epoques(i).Etapes(j).Valeur = Parametres.Batiment_Numero(Right(Temp, Len(Temp) - Len(MessageBatiment)), "Erreur dans le fichier d'IA : " & NomFichier & Chr(10) & Chr(13) & "A la ligne : " & Temp & " de l'époque " & i + 1, CheminFichier & NomFichier)
            Epoques(i).BatimentDefini = True
            j = j + 1
        ElseIf Left(Temp, Len(MessageObjet)) = MessageObjet Then
            'Objet.
            Ajouter_Etape i, j
            Epoques(i).Etapes(j).Categorie = 2
            Epoques(i).Etapes(j).Valeur = Mid(Temp, Len(MessageObjet) + 1, InStr(Temp, CaractereQuantite) - Len(MessageObjet) - 1) - 1
            Epoques(i).Etapes(j).Quantite = Mid(Temp, InStr(Temp, CaractereQuantite) + 1, Len(Temp))
            j = j + 1
        ElseIf Left(Temp, Len(MessageChangerMode)) = MessageChangerMode Then
            'Changement de mode.
            Ajouter_Etape i, j
            Epoques(i).Etapes(j).Categorie = 7
            Epoques(i).Etapes(j).Valeur = Right(Temp, Len(Temp) - Len(MessageChangerMode)) - 1
            j = j + 1
        ElseIf Left(Temp, Len(MessageGagnerArgent)) = MessageGagnerArgent Then
            'Gagner de l'argent.
            Ajouter_Etape i, j
            Epoques(i).Etapes(j).Categorie = 8
            Epoques(i).Etapes(j).Valeur = Parametres.Batiment_Numero(Mid(Temp, Len(MessageGagnerArgent) + 1, InStr(Temp, CaractereQuantite) - Len(MessageGagnerArgent) - 1), "Erreur dans le fichier d'IA : " & NomFichier & Chr(10) & Chr(13) & "A la ligne : " & Temp & " de l'époque " & i + 1, CheminFichier & NomFichier)
            'Epoques(i).Etapes(j).Valeur = Mid(Temp, Len(MessageGagnerArgent) + 1, InStr(Temp, CaractereQuantite) - Len(MessageGagnerArgent) - 1) - 1
            Epoques(i).Etapes(j).Quantite = Mid(Temp, InStr(Temp, CaractereQuantite) + 1, Len(Temp))
            j = j + 1
        ElseIf Left(Temp, Len(MessageGagnerArgentRessource)) = MessageGagnerArgentRessource Then
            'Gagner de l'argent.
            Ajouter_Etape i, j
            Epoques(i).Etapes(j).Categorie = 9
            Epoques(i).Etapes(j).Valeur = Mid(Temp, Len(MessageGagnerArgentRessource) + 1, InStr(Temp, CaractereQuantite) - Len(MessageGagnerArgentRessource) - 1) - 1
            Epoques(i).Etapes(j).Quantite = Mid(Temp, InStr(Temp, CaractereQuantite) + 1, Len(Temp))
            j = j + 1
        Else
            'Autres action
            Select Case Temp
            Case MessageTravailler: Ajouter_Etape i, j: Epoques(i).Etapes(j).Categorie = 3: j = j + 1
            Case MessagePatrouiller: Ajouter_Etape i, j: Epoques(i).Etapes(j).Categorie = 4: j = j + 1
            Case MessageAttaquerEnnemi: Ajouter_Etape i, j: Epoques(i).Etapes(j).Categorie = 5: j = j + 1
            Case MessageAttaquerDonjon: Ajouter_Etape i, j: Epoques(i).Etapes(j).Categorie = 6: j = j + 1
            End Select
        End If
        Line Input #1, Temp
    Wend
    Close #1
    Exit Sub
Erreur:
    MsgBox "Le fichier " & CheminFichier & NomFichier & " est invalide.", vbCritical, "Erreur"
    End
End Sub

Public Sub Creer_IA_Vierge(ByVal CheminFichier As String, _
                           ByVal NomFichier As String, _
                           ByVal Parametres As ClsJeuParam)
    Dim i As Integer
    Open CheminFichier & NomFichier For Output As #1
    Print #1, CaractereCommentaire
    Print #1, ""
    For i = 0 To Parametres.NombreEpoques - 1
        Print #1, MessageAge & i + 1
        Print #1, CaractereCommentaire
        Print #1, MessageDefaut
        Print #1, ""
    Next i
    Print #1, MessageFin
    Close #1
End Sub

Public Sub Sauvegarder_IA(ByVal CheminFichier As String, _
                          ByVal NomFichier As String, _
                          ByVal Parametres As ClsJeuParam)
    Dim i As Integer, j As Integer
    Open CheminFichier & NomFichier For Output As #1
    Print #1, CaractereCommentaire & Commentaire
    Print #1, MessageEquiper & Equiper
    Print #1, ""
    For i = 0 To Parametres.NombreEpoques - 1
        j = 0
        Print #1, MessageAge & i + 1
        Print #1, CaractereCommentaire & Epoques(i).Commentaire
        For j = 0 To Definir_NombreEtapes(i) - 1
            Select Case Epoques(i).Etapes(j).Categorie
            Case 0: Print #1, MessageRessource & Epoques(i).Etapes(j).Valeur + 1 & CaractereQuantite & Epoques(i).Etapes(j).Quantite
            Case 1: Print #1, MessageBatiment & Parametres.Batiment_ID(Epoques(i).Etapes(j).Valeur)
            Case 2: Print #1, MessageObjet & Epoques(i).Etapes(j).Valeur + 1 & CaractereQuantite & Epoques(i).Etapes(j).Quantite
            Case 3: Print #1, MessageTravailler
            Case 4: Print #1, MessagePatrouiller
            Case 5: Print #1, MessageAttaquerEnnemi
            Case 6: Print #1, MessageAttaquerDonjon
            Case 7: Print #1, MessageChangerMode & Epoques(i).Etapes(j).Valeur + 1
            'Case 8: Print #1, MessageGagnerArgent & Epoques(i).Etapes(j).Valeur + 1 & CaractereQuantite & Epoques(i).Etapes(j).Quantite
            Case 8: Print #1, MessageGagnerArgent & Parametres.Batiment_ID(Epoques(i).Etapes(j).Valeur) & CaractereQuantite & Epoques(i).Etapes(j).Quantite
            Case 9: Print #1, MessageGagnerArgentRessource & Epoques(i).Etapes(j).Valeur + 1 & CaractereQuantite & Epoques(i).Etapes(j).Quantite
            End Select
        Next j
        Print #1, ""
    Next i
    Print #1, MessageFin
    Close #1
End Sub

Public Sub Ajouter_Etape(ByVal NumeroEpoque As Integer, _
                         ByVal NumeroEtape As Integer)
    ReDim Preserve Epoques(NumeroEpoque).Etapes(NumeroEtape)
End Sub

Public Sub Intervertir_Etapes(ByVal NumeroEpoque As Integer, _
                              ByVal Indice1 As Integer, _
                              ByVal Indice2 As Integer)
    Dim Temp1 As Integer, Temp2 As Integer, Temp3 As Integer
    If Indice1 >= 0 And Indice1 <= Definir_NombreEtapes(NumeroEpoque) - 1 And _
       Indice2 >= 0 And Indice2 <= Definir_NombreEtapes(NumeroEpoque) - 1 Then
        Temp1 = Epoques(NumeroEpoque).Etapes(Indice1).Categorie
        Temp2 = Epoques(NumeroEpoque).Etapes(Indice1).Valeur
        Temp3 = Epoques(NumeroEpoque).Etapes(Indice1).Quantite
        
        Epoques(NumeroEpoque).Etapes(Indice1).Categorie = Epoques(NumeroEpoque).Etapes(Indice2).Categorie
        Epoques(NumeroEpoque).Etapes(Indice1).Valeur = Epoques(NumeroEpoque).Etapes(Indice2).Valeur
        Epoques(NumeroEpoque).Etapes(Indice1).Quantite = Epoques(NumeroEpoque).Etapes(Indice2).Quantite
        
        Epoques(NumeroEpoque).Etapes(Indice2).Categorie = Temp1
        Epoques(NumeroEpoque).Etapes(Indice2).Valeur = Temp2
        Epoques(NumeroEpoque).Etapes(Indice2).Quantite = Temp3
    End If
End Sub

Public Sub Supprimer_Etape(ByVal NumeroEpoque As Integer, _
                           ByVal Indice As Integer)
    Dim i As Integer
    If Definir_NombreEtapes(NumeroEpoque) >= 1 Then
        If Indice <> Definir_NombreEtapes(NumeroEpoque) - 1 Then
            For i = Indice To Definir_NombreEtapes(NumeroEpoque) - 1
                Intervertir_Etapes NumeroEpoque, i, i + 1
            Next i
        End If
        ReDim Preserve Epoques(NumeroEpoque).Etapes(Definir_NombreEtapes(NumeroEpoque) - 2)
    End If
End Sub

'Propriétés de la classe.
Public Property Get Definir_NombreEtapes(ByVal NumeroEpoque As Integer) As Integer
    Definir_NombreEtapes = UBound(Epoques(NumeroEpoque).Etapes()) + 1
End Property
Public Property Get Definir_Commentaire(ByVal NumeroEpoque As Integer) As String
    Definir_Commentaire = Epoques(NumeroEpoque).Commentaire
End Property
Public Property Let Definir_Commentaire(ByVal NumeroEpoque As Integer, ByVal Valeur As String)
    Epoques(NumeroEpoque).Commentaire = Valeur
End Property
Public Property Get Definir_Categorie(ByVal NumeroEpoque As Integer, _
                                      ByVal NumeroEtape As Integer) As Integer
    Definir_Categorie = Epoques(NumeroEpoque).Etapes(NumeroEtape).Categorie
End Property
Public Property Let Definir_Categorie(ByVal NumeroEpoque As Integer, _
                                      ByVal NumeroEtape As Integer, _
                                      ByVal Valeur As Integer)
    Epoques(NumeroEpoque).Etapes(NumeroEtape).Categorie = Valeur
End Property
Public Property Get Definir_Valeur(ByVal NumeroEpoque As Integer, _
                                   ByVal NumeroEtape As Integer) As Integer
    Definir_Valeur = Epoques(NumeroEpoque).Etapes(NumeroEtape).Valeur
End Property
Public Property Let Definir_Valeur(ByVal NumeroEpoque As Integer, _
                                   ByVal NumeroEtape As Integer, _
                                   ByVal Valeur As Integer)
    Epoques(NumeroEpoque).Etapes(NumeroEtape).Valeur = Valeur
End Property
Public Property Get Definir_Quantite(ByVal NumeroEpoque As Integer, _
                                     ByVal NumeroEtape As Integer) As Integer
    Definir_Quantite = Epoques(NumeroEpoque).Etapes(NumeroEtape).Quantite
End Property
Public Property Let Definir_Quantite(ByVal NumeroEpoque As Integer, _
                                     ByVal NumeroEtape As Integer, _
                                     ByVal Valeur As Integer)
    Epoques(NumeroEpoque).Etapes(NumeroEtape).Quantite = Valeur
End Property
Public Property Get Definir_BatimentDefini(ByVal NumeroEpoque As Integer) As Boolean
    Definir_BatimentDefini = Epoques(NumeroEpoque).BatimentDefini
End Property

