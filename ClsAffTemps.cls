VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAffTemps"
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

Const FichierINI = "Temps"
Const SectionINI = "Temps"
Const SectionINI2 = "Temp_"

Private AfficherTemps As Boolean

Private Chemin As String
Private Largeur As Integer
Private Hauteur As Integer
Public NombreTemps As Integer
Private Type Temp
    Debut As Long
    Afficher As Boolean
    Image As String
    Sprite As ClsAffSprite
End Type

Private Temps() As Temp
Public TempEnCours As Integer

Public Heures As Integer
Public Minutes As Single

Private JourneeDebut As Integer
Private JourneeVitesse As Single

Private Sub Class_Initialize()
    Dim i As Integer
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    Chemin = FicIni.Parametre("Chemin")
    
    AfficherTemps = CBool(FicIni.Parametre("Afficher"))
    Largeur = Val(FicIni.Parametre("Largeur"))
    Hauteur = Val(FicIni.Parametre("Hauteur"))
    NombreTemps = Val(FicIni.Parametre("NombreTemps"))
    JourneeDebut = Val(FicIni.Parametre("JourneeDebut"))
    JourneeVitesse = Val(FicIni.Parametre("JourneeVitesse"))
    ReDim Temps(NombreTemps - 1)
    For i = 0 To NombreTemps - 1
        FicIni.Section = SectionINI2 & i + 1
        With Temps(i)
        .Debut = Val(FicIni.Parametre("Debut"))
        .Afficher = CBool(FicIni.Parametre("Afficher"))
        .Image = FicIni.Parametre("Image")
        Set .Sprite = New ClsAffSprite
        End With
    Next i
End Sub

Public Sub Init()
    Heures = JourneeDebut
    Minutes = 0
    Actualiser_Temps
End Sub

Public Sub Charger(ByRef dd As DirectDraw7)
    Dim i As Integer
    'If AfficherTemps Then
        For i = 0 To NombreTemps - 1
            If Temps(i).Afficher Then
                Temps(i).Sprite.Charger dd, Chemin & Temps(i).Image
            End If
        Next i
    'End If
End Sub

Public Sub Actualiser()
    Minutes = Minutes + JourneeVitesse
    If Minutes > 60 Then
        Actualiser_Temps
        Heures = Heures + Int(Minutes / 60)
        Minutes = Minutes Mod 60
        Heures = Heures Mod 24
    End If
End Sub

Public Sub Actualiser_Temps()
    Dim i  As Integer
    TempEnCours = -1
    'If AfficherTemps Then
        For i = 0 To NombreTemps - 2
            If Heures >= Temps(i).Debut And Heures <= Temps(i + 1).Debut Then
                TempEnCours = i
                Exit For
            End If
        Next i
        If TempEnCours = -1 Then
            TempEnCours = NombreTemps - 1
        End If
    'End If
End Sub

Public Sub Afficher(ByVal DecalageX As Long, _
                    ByVal DecalageY As Long, _
                    ByVal LargeurEcran As Integer, _
                    ByVal HauteurEcran As Integer, _
                    ByVal EcranDecalageX As Integer, _
                    ByVal EcranDecalageY As Integer, _
                    ByRef SurfBack As DirectDrawSurface7, _
                    Optional ByVal ForcerAffichage As Boolean)
    Dim i As Integer, j As Integer
    If AfficherTemps Or ForcerAffichage Then
        If Temps(TempEnCours).Afficher Then
            'On affiche le temps sur toute la surface de l'écran.
            'For i = 0 To Monde.Largeur * Monde.HauteurCase Step Largeur
            For i = EcranDecalageX - DecalageX Mod Largeur To LargeurEcran Step Largeur
                'For j = 0 To Monde.Hauteur * Monde.HauteurCase Step Hauteur
                For j = EcranDecalageY - DecalageY Mod Hauteur To HauteurEcran Step Hauteur
                    Temps(TempEnCours).Sprite.Selection 0, Hauteur, 0, Largeur
                    Temps(TempEnCours).Sprite.Afficher i, j, SurfBack, True
                    'Temps(TempEnCours).Sprite.Afficher i - DecalageX, j - DecalageY, SurfBack, True
                    'Exit For
                Next j
                'Exit For
            Next i
        End If
    End If
    'Messages(0).Ajouter_Message Heures & " : " & Format(Minutes, "00")
    'Messages(0).Ajouter_Message i & " : " & j & " : " & DecalageX & " : " & DecalageY
    'If i <> -DecalageX Or j <> -DecalageY Then MsgBox "test"
End Sub

'Propriétés
Public Property Get Definir_Afficher() As Boolean
    Definir_Afficher = AfficherTemps
End Property
Public Property Let Definir_Afficher(ByVal Valeur As Boolean)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    AfficherTemps = Valeur
    FicIni.Parametre("Afficher") = -CBool(AfficherTemps)
End Property


