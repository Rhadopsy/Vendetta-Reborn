VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsJeuCommentaires"
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

Const FichierINI = "Commentaires"
Const SectionINI = "Commentaires"

Private Afficher As Boolean
Public Duree As Single
Public DureeMin As Single
Public DureeMax As Single
Public DureeIntervale As Single

Public Couleur As Long

Public DecalageX As Integer
Public DecalageY As Integer

Private NombreTypeMessages As Integer
Private NombreMaxMessagesParType As Integer

'Liste des messages en commentaire.
Private NombreMessages() As String
Private Messages() As String


Private Sub Class_Initialize()
    Dim i As Long
    Dim j As Long
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    
    Afficher = CBool(FicIni.Parametre("Afficher"))
    Duree = Val(FicIni.Parametre("Duree")) / 100
    DureeMin = Val(FicIni.Parametre("DureeMin"))
    DureeMax = Val(FicIni.Parametre("DureeMax"))
    DureeIntervale = Val(FicIni.Parametre("DureeIntervale"))
    Couleur = Val(FicIni.Parametre("Couleur"))
    DecalageX = Val(FicIni.Parametre("DecalageX"))
    DecalageY = Val(FicIni.Parametre("DecalageY"))
    
    NombreTypeMessages = Val(FicIni.Parametre("NombreTypeMessages"))
    NombreMaxMessagesParType = Val(FicIni.Parametre("NombreMaxMessagesParType"))

    ReDim Messages(NombreTypeMessages - 1, NombreMaxMessagesParType - 1)
    ReDim NombreMessages(NombreTypeMessages - 1)
    
    For i = 0 To NombreTypeMessages - 1
        FicIni.Section = "Commentaires_" & i + 1
        NombreMessages(i) = Val(FicIni.Parametre("NombreMessages"))
        For j = 0 To NombreMessages(i) - 1
            Messages(i, j) = FicIni.Parametre("Message(" & j + 1 & ")")
        Next j
    Next i
End Sub

Public Function Message(ByVal NoMessage As Integer, ByVal Perso As ClsJeuPerso) As String
    If NoMessage >= 0 Then
        If Perso.CommentairesSpecifique Then
            Message = Messages(Perso.CommentairesSpecifiqueIndice, Int(Rnd * NombreMessages(Perso.CommentairesSpecifiqueIndice)))
        Else
            Message = Messages(NoMessage, Int(Rnd * NombreMessages(NoMessage)))
        End If
    Else
        Message = ""
    End If
End Function

'Propriétés
Public Property Get Definir_Afficher() As Boolean
    Definir_Afficher = Afficher
End Property
Public Property Let Definir_Afficher(ByVal Valeur As Boolean)
    Afficher = Valeur
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    If Afficher Then
        FicIni.Parametre("Afficher") = 1
    Else
        FicIni.Parametre("Afficher") = 0
    End If
End Property

Public Property Get Definir_Duree() As Single
    Definir_Duree = Duree
End Property
Public Property Let Definir_Duree(ByVal Valeur As Single)
    Duree = Valeur
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("Duree") = Valeur * 100
End Property

Public Property Get Definir_DureeMin() As Single
    Definir_DureeMin = DureeMin
End Property
Public Property Get Definir_DureeMax() As Single
    Definir_DureeMax = DureeMax
End Property
Public Property Get Definir_DureeIntervale() As Single
    Definir_DureeIntervale = DureeIntervale
End Property
