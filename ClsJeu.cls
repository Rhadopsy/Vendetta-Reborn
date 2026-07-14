VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsJeu"
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

Const FichierINI = "Jeu"
Const SectionINI = "Jeu"

Private AffichageRapide As Boolean

Private FPS As Integer
Public FPSMin As Integer
Public FPSMax As Integer

Private Vitesse As Integer
Public VitesseMin As Integer
Public VitesseMax As Integer

Public EnPause As Boolean

Public Campagne As Boolean 'Si vrai, on joue à une partie de la campagne.
Public Chrono As Integer 'Chronomètre
Public IntituleChrono As String
Public VictoireStandard As Boolean
Public DefaiteStandard As Boolean
Public Confrontation As Boolean

Private AfficherFantomes As Boolean
Private AfficherFantomesPNJ As Boolean
Public AfficherProjectiles As Boolean

Private Mode2Joueurs As Boolean
Private Mode2JoueursEcranSepares As Boolean

Const NombreRegles = 4
Const RegleInvocationsMultiples = 0
Const RegleTempsRechargementObjets = 1
Const ReglePathfinding = 2
Const ModeConfrontation = 3
Private Regles() As Integer

Private Sub Class_Initialize()
    Dim i As Byte
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    AffichageRapide = Val(FicIni.Parametre("AffichageRapide"))
    FPS = Val(FicIni.Parametre("FPS"))
    FPSMin = Val(FicIni.Parametre("FPSMin"))
    FPSMax = Val(FicIni.Parametre("FPSMax"))
    Vitesse = Val(FicIni.Parametre("Vitesse"))
    VitesseMin = Val(FicIni.Parametre("VitesseMin"))
    VitesseMax = Val(FicIni.Parametre("VitesseMax"))
    
    AfficherFantomes = Val(FicIni.Parametre("AfficherFantomes"))
    AfficherFantomesPNJ = Val(FicIni.Parametre("AfficherFantomesPNJ"))
    AfficherProjectiles = Val(FicIni.Parametre("AfficherProjectiles"))
    
    Mode2Joueurs = Val(FicIni.Parametre("Mode2Joueurs"))
    Mode2JoueursEcranSepares = Val(FicIni.Parametre("Mode2JoueursEcranSepares"))
    
    ReDim Regles(NombreRegles - 1)
    For i = 0 To NombreRegles - 1
        Regles(i) = Val(FicIni.Parametre("Regle" & i + 1))
    Next i
End Sub

Public Sub Init()
    Chrono = 0
    IntituleChrono = ""
    EnPause = False
    VictoireStandard = True
    DefaiteStandard = True
End Sub

'Propriétés
Public Property Get Definir_FPS() As Integer
    Definir_FPS = FPS
End Property
Public Property Let Definir_FPS(ByVal Valeur As Integer)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    FPS = Valeur
    FicIni.Parametre("FPS") = FPS
End Property

Public Property Get Definir_Vitesse() As Integer
    Definir_Vitesse = Vitesse
End Property
Public Property Let Definir_Vitesse(ByVal Valeur As Integer)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    Vitesse = Valeur
    FicIni.Parametre("Vitesse") = Vitesse
End Property

Public Property Get Definir_AffichageRapide() As Boolean
    Definir_AffichageRapide = AffichageRapide
End Property
Public Property Let Definir_AffichageRapide(ByVal Valeur As Boolean)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    AffichageRapide = Valeur
    If AffichageRapide Then
        FicIni.Parametre("AffichageRapide") = 1
    Else
        FicIni.Parametre("AffichageRapide") = 0
    End If
End Property

Public Property Get Definir_AfficherFantomes() As Boolean
    Definir_AfficherFantomes = AfficherFantomes
End Property
Public Property Let Definir_AfficherFantomes(ByVal Valeur As Boolean)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    AfficherFantomes = Valeur
    FicIni.Parametre("AfficherFantomes") = -CDbl(AfficherFantomes)
End Property

Public Property Get Definir_AfficherFantomesPNJ() As Boolean
    Definir_AfficherFantomesPNJ = AfficherFantomesPNJ
End Property
Public Property Let Definir_AfficherFantomesPNJ(ByVal Valeur As Boolean)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    AfficherFantomesPNJ = Valeur
    If AfficherFantomesPNJ Then
        FicIni.Parametre("AfficherFantomesPNJ") = 1
    Else
        FicIni.Parametre("AfficherFantomesPNJ") = 0
    End If
End Property

Public Property Get Definir_AfficherProjectiles() As Boolean
    Definir_AfficherProjectiles = AfficherProjectiles
End Property
Public Property Let Definir_AfficherProjectiles(ByVal Valeur As Boolean)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    AfficherProjectiles = Valeur
    If AfficherProjectiles Then
        FicIni.Parametre("AfficherProjectiles") = 1
    Else
        FicIni.Parametre("AfficherProjectiles") = 0
    End If
End Property

Public Property Get Definir_Mode2Joueurs() As Boolean
    Definir_Mode2Joueurs = Mode2Joueurs
End Property
Public Property Let Definir_Mode2Joueurs(ByVal Valeur As Boolean)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    Mode2Joueurs = Valeur
    If Mode2Joueurs Then
        FicIni.Parametre("Mode2Joueurs") = 1
    Else
        FicIni.Parametre("Mode2Joueurs") = 0
    End If
End Property

Public Property Get Definir_Mode2JoueursEcranSepares() As Boolean
    Definir_Mode2JoueursEcranSepares = Mode2JoueursEcranSepares
End Property
Public Property Let Definir_Mode2JoueursEcranSepares(ByVal Valeur As Boolean)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    Mode2JoueursEcranSepares = Valeur
    If Mode2JoueursEcranSepares Then
        FicIni.Parametre("Mode2JoueursEcranSepares") = 1
    Else
        FicIni.Parametre("Mode2JoueursEcranSepares") = 0
    End If
End Property

Public Property Get Definir_Regle(ByVal Index As Integer) As Integer
    Definir_Regle = Regles(Index)
End Property
Public Property Let Definir_Regle(ByVal Index As Integer, ByVal Valeur As Integer)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    Regles(Index) = Valeur
    FicIni.Parametre("Regle" & Index + 1) = Valeur
End Property

'Propriétés des regles de jeu.
Public Property Get Regle_Invocations_Multiples() As Integer
    Regle_Invocations_Multiples = Regles(RegleInvocationsMultiples)
End Property
Public Property Get Regle_TempsRechargement_Objets() As Integer
    Regle_TempsRechargement_Objets = Regles(RegleTempsRechargementObjets)
End Property
Public Property Get Regle_Pathfinding() As Integer
    Regle_Pathfinding = Regles(ReglePathfinding)
End Property
Public Property Get Regle_ModeConfrontation() As Boolean
    Regle_ModeConfrontation = Regles(ModeConfrontation)
End Property

Public Property Get Definir_Victoire() As Boolean
    Definir_Victoire = Victoire(0)
End Property
Public Property Let Definir_Victoire(ByVal Valeur As Boolean)
    If Not Victoire(0) And Valeur Then
        Musique.Jouer_Musique 5, False
    End If
    Victoire(0) = Valeur
End Property
Public Property Get Definir_Defaite() As Boolean
    Definir_Defaite = Defaite(0)
End Property
Public Property Let Definir_Defaite(ByVal Valeur As Boolean)
    If Not Defaite(0) And Valeur Then
        Musique.Jouer_Musique 6, False
    End If
    Defaite(0) = Valeur
End Property

Public Property Get Duree() As Double
    Duree = FrmMoteur2D.DureePartie / Jeu.Definir_Vitesse
End Property

Public Sub Activer_Effet_Perso(ByVal TypeEffet As Integer, _
                               ByVal Perso As ClsJeuPerso)
    With Perso
    Activer_Effet TypeEffet, .PositionX, .PositionY, .Largeur, .Hauteur, Parametres
    End With
End Sub

Public Property Get Definir_Chrono() As String
    Dim Duree As Integer, Minutes As Integer, Secondes As Integer
    Duree = Int(Chrono / Vitesse)
    If Duree >= 60 Then
        Minutes = Int(Duree / 60)
        Secondes = Duree - (Minutes * 60)
        'Definir_Chrono = Minutes & IIf(Duree Mod 2 > 0, ":", " ") & Format(Secondes, "00")
        Definir_Chrono = Minutes & ":" & Format(Secondes, "00") & "s"
    Else
        Definir_Chrono = Format(Duree, "00") & "s"
    End If
End Property
