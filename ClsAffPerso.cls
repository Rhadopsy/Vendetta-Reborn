VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAffPerso"
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

Const Chemin = "Persos\"

Private NombreSpritesHeros As Long
Private DossiersPersos() As String
Private Sprites() As ClsAffSprite
Private SpritesNom() As String

Private Sub Class_Initialize()
    Dim i As Integer
    Dim Temp As String
    ReDim Sprites(0)
    Set Sprites(0) = New ClsAffSprite
    ReDim SpritesNom(0)
    
    'Recherche les dossiers où il y a les images des personnages.
    ReDim DossiersPersos(0)
    Temp = Dir(App.Path & CheminPerso, vbDirectory)
    While Temp <> ""
        If InStr(Temp, ".") = 0 Then
            ReDim Preserve DossiersPersos(UBound(DossiersPersos()) + 1)
            DossiersPersos(UBound(DossiersPersos()) - 1) = Sprites(0).CheminImage & Chemin & Temp
        End If
        Temp = Dir()
    Wend
    ReDim Preserve DossiersPersos(UBound(DossiersPersos()) - 1)
    'Recherche les sous répertoires.
    For i = 0 To UBound(DossiersPersos())
        Temp = Dir(App.Path & DossiersPersos(i) & "\", vbDirectory)
        While Temp <> ""
            If InStr(Temp, ".") = 0 Then
                ReDim Preserve DossiersPersos(UBound(DossiersPersos()) + 1)
                DossiersPersos(UBound(DossiersPersos())) = DossiersPersos(i) & "\" & Temp
            End If
            Temp = Dir()
        Wend
    Next i
    'ReDim Preserve DossiersPersos(UBound(DossiersPersos()) - 1)
End Sub

Public Sub Charger(ByRef dd As IDirectDraw7, _
                   ByRef Parametres As ClsJeuParam)
    Dim i As Long, j As Long
    Dim Temp As String
    
    For j = 0 To UBound(DossiersPersos())
        Temp = Dir(App.Path & DossiersPersos(j) & "\", vbNormal)
        While Temp <> ""
            NombreSpritesHeros = NombreSpritesHeros + 1
            Temp = Dir()
        Wend
    Next j
    
    'ReDim Preserve Sprites(1 + NombreSpritesHeros - 1)
    'ReDim SpritesNom(1 + NombreSpritesHeros - 1)
    ReDim Preserve Sprites(NombreSpritesHeros - 1)
    ReDim SpritesNom(NombreSpritesHeros - 1)
    
    'Chargement des images des héros.
    For j = 0 To UBound(DossiersPersos())
        Temp = Dir(App.Path & DossiersPersos(j) & "\", vbNormal)
        While Temp <> ""
            Set Sprites(i) = New ClsAffSprite
            Sprites(i).Charger dd, DossiersPersos(j) & "\" & _
                                   Left$(Temp, InStr(Temp, ".") - 1), True
            SpritesNom(i) = Temp
            i = i + 1
            Temp = Dir()
        Wend
    Next j
    
    'Chargement des images des humains.
    'Set Sprites(NombreSpritesHeros) = New ClsAffSprite
    'Sprites(NombreSpritesHeros).Charger dd, Chemin & Parametres.Peuples_Nom(0)
    'SpritesNom(NombreSpritesHeros) = Parametres.Peuples_Nom(0)
End Sub

Public Function Charger_Hero(ByVal FichierHero As String, _
                             ByVal Feminin As Boolean) As Integer
    'Dim i, j As Long
    'Dim Temp As String
    Dim i As Integer
    
    If FichierHero <> "" Then
        For i = 0 To UBound(SpritesNom())
            If SpritesNom(i) = FichierHero Then
                Charger_Hero = i
                i = UBound(SpritesNom())
            End If
        Next i
        'For j = 0 To UBound(DossiersPersos())
        '    Temp = Dir(App.Path & CheminPerso & DossiersPersos(j) & "\", vbNormal)
        '    While Temp <> ""
        '        If Temp = FichierHero Then
        '            Charger_Hero = i
        '        End If
        '        i = i + 1
        '        Temp = Dir()
        '    Wend
        'Next j
    End If
End Function

Public Function Tirer_Apparence_Au_Hasard(ByVal DossierPerso As String, _
                                          ByVal Parametres As ClsJeuParam, _
                                          ByVal CheminAbsolu As Boolean) As String
    Dim i As Long, j As Long
    Dim Temp As String
    
    If CheminAbsolu Then
        Temp = Dir(App.Path & DossierPerso & "\", vbNormal)
    Else
        Temp = Dir(App.Path & CheminPerso & DossierPerso & "\", vbNormal)
    End If
    While Temp <> ""
        i = i + 1
        Temp = Dir()
    Wend
    j = Int(Rnd * i)
    i = 0
    If CheminAbsolu Then
        Temp = Dir(App.Path & DossierPerso & "\", vbNormal)
    Else
        Temp = Dir(App.Path & CheminPerso & DossierPerso & "\", vbNormal)
    End If
    While Temp <> ""
        If i = j Then
            Tirer_Apparence_Au_Hasard = Temp
        End If
        i = i + 1
        Temp = Dir()
    Wend
End Function

Public Sub Afficher(ByVal X As Long, _
                    ByVal Y As Long, _
                    ByRef NumeroPas As Integer, _
                    ByVal Direction As Integer, _
                    ByVal LargeurImage As Integer, _
                    ByVal HauteurImage As Integer, _
                    ByVal Perso As ClsJeuPerso, _
                    ByVal ApparenceChoisie As Integer, _
                    ByRef SurfBack As IDirectDrawSurface7, _
                    Optional ByVal ChoisirApparence As Boolean, _
                    Optional ByVal Translucide As Boolean, _
                    Optional ByVal DecalageX As Integer = 0, _
                    Optional ByVal DecalageY As Integer = 0)
    
    Dim i As Integer
    Dim Temp As Boolean
    
    With Perso
    If .Monture > 0 Then
        Temp = True 'Passe à faux si l'on affiche la monture en avant plan.
        'Affiche la monture si elle est en avant plan.
        If .Definir_ApparenceAvantPlan(Direction) Then
            AffApparence.Afficher_Element .Definir_ApparenceObjet(.Monture), X + .ApparenceDecalageXTemp, Y + .ApparenceDecalageYTemp, NumeroPas, Direction, .Definir_ApparenceLargeur, .Definir_ApparenceHauteur, SurfBack, Translucide, DecalageX, DecalageY
            Temp = False
        End If
        If DecalageX = 0 And DecalageY = 0 Then
            Afficher2 X + .Definir_ApparenceDecalageX(Direction), Y + .Definir_ApparenceDecalageY(Direction), 1, Direction, LargeurImage, HauteurImage, ApparenceChoisie, SurfBack, ChoisirApparence, Translucide, DecalageX, DecalageY
        Else
            Afficher2 X + .Definir_ApparenceDecalageX(Direction) + DecalageX, Y + .Definir_ApparenceDecalageY(Direction) + DecalageY, 1, Direction, LargeurImage, HauteurImage, ApparenceChoisie, SurfBack, ChoisirApparence, Translucide, DecalageX + .Definir_ApparenceDecalageX(Direction), DecalageY + .Definir_ApparenceDecalageY(Direction)
        End If
        If Temp Then
            AffApparence.Afficher_Element .Definir_ApparenceObjet(.Monture), X + .ApparenceDecalageXTemp, Y + .ApparenceDecalageYTemp, NumeroPas, Direction, .Definir_ApparenceLargeur, .Definir_ApparenceHauteur, SurfBack, Translucide, DecalageX, DecalageY
        End If
    Else
        Afficher2 X, Y, NumeroPas, Direction, LargeurImage, HauteurImage, ApparenceChoisie, SurfBack, ChoisirApparence, Translucide, DecalageX, DecalageY
    End If
    End With
End Sub

Private Sub Afficher2(ByVal X As Long, _
                      ByVal Y As Long, _
                      ByRef NumeroPas As Integer, _
                      ByVal Direction As Integer, _
                      ByVal LargeurImage As Integer, _
                      ByVal HauteurImage As Integer, _
                      ByVal ApparenceChoisie As Integer, _
                      ByRef SurfBack As IDirectDrawSurface7, _
                      ByVal ChoisirApparence As Boolean, _
                      ByVal Translucide As Boolean, _
                      ByVal DecalageX As Integer, _
                      ByVal DecalageY As Integer)
    'Pour la direction :
    '0 = immobile '1 = haut '2 = droite '3 = bas '4 = gauche.
    With Sprites(ApparenceChoisie)
    
    If Direction > 0 Then
        If NumeroPas = 3 Then
            .Selection (Direction - 1) * HauteurImage, _
                        Direction * HauteurImage, _
                        LargeurImage, _
                        2 * LargeurImage
        Else
            .Selection (Direction - 1) * HauteurImage, _
                        Direction * HauteurImage, _
                        NumeroPas * LargeurImage, _
                        (NumeroPas + 1) * LargeurImage
        End If
    Else
        .Selection HauteurImage * 2, _
                    HauteurImage * 3, _
                    LargeurImage, _
                    2 * LargeurImage
    End If
    
    'On affiche le sprite sur l'arrière plan.
    .Afficher X, Y, SurfBack, True, Translucide, DecalageX, DecalageY
    
    End With
End Sub

Public Property Get CheminPerso() As String
    CheminPerso = Sprites(0).CheminImage & Chemin
End Property

Public Sub Ajouter_Dossier(ByVal Dossier As String)
    ReDim Preserve DossiersPersos(UBound(DossiersPersos()) + 1)
    DossiersPersos(UBound(DossiersPersos())) = Dossier
End Sub
