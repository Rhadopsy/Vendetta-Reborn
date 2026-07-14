VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAffTerrain"
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

Public Largeur As Long
Public Hauteur As Long

Const OptimiserAffichageLargeurMax = 512
Const OptimiserAffichageHauteurMax = 512

Private OptimiserAffichageLargeur As Integer
Private OptimiserAffichageHauteur As Integer

Private Const Chemin = "\Images\terrains\"
Private Const Extension = ".bmp"

Private Sprite As DirectDrawSurface7 'creation surface Primaire.
Private AffSprite As DDSURFACEDESC2  'description surface Primaire.
Private RectSprite As RECT           'rectangle Titre.

Private TerrainSprite() As DirectDrawSurface7 'creation surface Primaire.
Private AffTerrainSprite As DDSURFACEDESC2 'description surface Primaire.

Private EtapeAnimation1 As Integer
Private EtapeAnimation2 As Integer
Private EtapeAnimationRetour As Boolean


Private Sub Class_Initialize()
    AffSprite.lFlags = DDSD_CAPS Or DDSD_CKSRCBLT
    AffSprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    AffTerrainSprite.lFlags = DDSD_CAPS Or DDSD_CKSRCBLT
    AffTerrainSprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
End Sub

Public Sub Charger(ByRef dd As DirectDraw7, ByVal TerrainZone As String)
' Autant le faire directement pour éviter l'appel en cascade d'objets
On Local Error GoTo NomaP
    Set Sprite = dd.CreateSurfaceFromFile(App.Path & Chemin & TerrainZone & Extension, AffSprite)
    Set FrmMoteur2D.picTemp.Picture = LoadPicture(App.Path & Chemin & TerrainZone & Extension)
    Exit Sub
NomaP:
    Err.Clear
    MsgBox "Texture du terrain " & TerrainZone & " introuvable. Abandon.", vbCritical + vbOKOnly
    FrmMoteur2D.FatalEnd
End Sub

Public Sub Afficher(ByVal X As Long, _
                    ByVal Y As Long, _
                    ByVal ApparenceTerrain As Long, _
                    ByVal TypeTerrain As Long, _
                    ByRef SurfBack As DirectDrawSurface7)
    'On charge la bonne image en fonction des paramètres.
    Dim Etat As Integer
    
    RectSprite.Top = Hauteur * TypeTerrain
    RectSprite.Bottom = Hauteur * (TypeTerrain + 1)
    RectSprite.Left = Largeur * ApparenceTerrain
    RectSprite.Right = Largeur * (ApparenceTerrain + 1)
    
    SurfBack.BltFast X, Y, Sprite, RectSprite, DDBLTFAST_SRCCOLORKEY
End Sub

Public Sub Animation_Suivante()
    If Monde.TerrainAnime Then
        EtapeAnimation1 = EtapeAnimation1 + 1
        If EtapeAnimation1 > Jeu.Definir_Vitesse Then
            EtapeAnimation1 = 0
            If EtapeAnimationRetour Then
                EtapeAnimation2 = EtapeAnimation2 - 1
            Else
                EtapeAnimation2 = EtapeAnimation2 + 1
            End If
            If EtapeAnimation2 = Monde.NombreEtapesAnimation - 1 Then
                EtapeAnimationRetour = True
            ElseIf EtapeAnimation2 = 0 Then
                EtapeAnimationRetour = False
            End If
        End If
    End If
End Sub

Public Sub Modeliser_Carte(ByRef dd As DirectDraw7, ByVal Monde As ClsJeuMonde)
    Dim i As Long, j As Long
    Dim k As Long, l As Long
    Dim m As Long, n As Long, o As Long
    Dim TpTerrain As Integer
    Dim CaTerrain As Integer
    
    Dim RectBack As RECT
    Dim AffBack As DDSURFACEDESC2
    Dim AffFore As DDSURFACEDESC2
    Dim RectFore As RECT
    
    Dim NombreEtapesAnimation As Integer
    
    'Calcul la largeur des grandes cases maximales.
    For i = OptimiserAffichageLargeurMax To 0 Step -Monde.LargeurCase
        If (Monde.Largeur * Monde.LargeurCase) Mod i = 0 Then
            OptimiserAffichageLargeur = i
            Exit For
        End If
    Next i
    For i = OptimiserAffichageHauteurMax To 0 Step -Monde.HauteurCase
        If (Monde.Hauteur * Monde.HauteurCase) Mod i = 0 Then
            OptimiserAffichageHauteur = i
            Exit For
        End If
    Next i
    'Debug.Print OptimiserAffichageLargeur & ", " & OptimiserAffichageHauteur
    
    NombreEtapesAnimation = IIf(Monde.TerrainAnime, Monde.NombreEtapesAnimation, 1)
    
    'AffTerrainSprite.lHeight = Monde.Hauteur * Hauteur
    'AffTerrainSprite.lWidth = Monde.Largeur * Largeur
    'FrmMoteur2D.picCarte.Width = Monde.Largeur * Largeur
    'FrmMoteur2D.picCarte.Height = Monde.Hauteur * Hauteur
    FrmMoteur2D.picCarte.Width = OptimiserAffichageLargeur
    FrmMoteur2D.picCarte.Height = OptimiserAffichageHauteur
    'Debug.Print FrmMoteur2D.picCarte.Width & ", " & FrmMoteur2D.picCarte.Height
    ReDim TerrainSprite(((Int(Monde.Largeur / OptimiserAffichageLargeur * Largeur) + 1) * (Int(Monde.Hauteur / OptimiserAffichageHauteur * Hauteur) + 1) * NombreEtapesAnimation) - 1)
    'ReDim TerrainSprite(((Int(Monde.Largeur / OptimiserAffichageLargeur * Largeur)) * (Int(Monde.Hauteur / OptimiserAffichageHauteur * Hauteur)) * NombreEtapesAnimation) - 1)
    'ReDim TerrainSprite(NombreEtapesAnimation - 1)
    For l = 0 To NombreEtapesAnimation - 1
        For m = 0 To Int(Monde.Hauteur / OptimiserAffichageHauteur * Hauteur)
            For n = 0 To Int(Monde.Largeur / OptimiserAffichageLargeur * Largeur)
                'Set TerrainSprite(l) = dd.CreateSurfaceFromFile(App.Path & Chemin & "Plaine" & Extension, AffTerrainSprite)
                'Sprite.GetSurfaceDesc AffFore
                'Sprite.Lock RectFore, AffFore, DDLOCK_WAIT, 0
                'TerrainSprite(l).GetSurfaceDesc AffBack
                'TerrainSprite(l).Lock RectBack, AffBack, DDLOCK_WAIT, 0
                For i = 0 To OptimiserAffichageHauteur / Hauteur
                    For j = 0 To OptimiserAffichageLargeur / Largeur
                        k = (i + (m * OptimiserAffichageHauteur / Hauteur)) * Monde.Largeur + j + (n * OptimiserAffichageLargeur / Largeur)
                        If i + (m * OptimiserAffichageHauteur / Hauteur) <= Monde.Hauteur - 1 And j + (n * OptimiserAffichageLargeur / Largeur) <= Monde.Largeur - 1 Then
                            
                            TpTerrain = Monde.Param_TerrainApparence(k)
                            CaTerrain = Monde.Param_TerrainCategorie(k)
                            'Animation (dans le cas de l'eau).
                            If Monde.Param_TerrainAnimation(k) Then CaTerrain = CaTerrain + l
                            'RectSprite.Left = Largeur * TpTerrain
                            'RectSprite.Right = Largeur * (TpTerrain + 1)
                            'RectSprite.Top = Hauteur * CaTerrain
                            'RectSprite.Bottom = Hauteur * (CaTerrain + 1)
                            
                            'TerrainSprite.BltFast i * Monde.LargeurCase, _
                                                  j * Monde.HauteurCase, Sprite, RectSprite, DDBLTFAST_SRCCOLORKEY
                            FrmMoteur2D.picCarte.PaintPicture FrmMoteur2D.picTemp.Picture, _
                                                              j * Largeur, i * Hauteur, _
                                                              Largeur, Hauteur, _
                                                              Largeur * TpTerrain, Hauteur * CaTerrain, _
                                                              Largeur, Hauteur
                            'For m = 0 To Hauteur - 1
                            '    For n = 0 To Largeur - 1
                                    'Debug.Print Sprite.GetLockedPixel(Largeur * TpTerrain + n + 1, Hauteur * CaTerrain + m + 1)
                            '        TerrainSprite(l).SetLockedPixel j * Largeur + n, _
                                                                    i * Hauteur + m, _
                                                                    Sprite.GetLockedPixel(Largeur * TpTerrain + n, Hauteur * CaTerrain + m)
                            '    Next n
                            'Next m
                            'k = k + 1
                            'FrmMoteur2D.Afficher_Image_Chargement dd, 80, _
                            "Optimisation de l'affichage " & Format(((k + (l * Monde.Hauteur * Monde.Largeur)) / (Monde.Hauteur * Monde.Largeur) / NombreEtapesAnimation) * 100, "00") & " %"
                        End If
                    Next j
                Next i
                
                FrmMoteur2D.Afficher_Image_Chargement dd, 80, _
                                                      Parametres.Etiquette_Chargement(7) & Format(o / UBound(TerrainSprite()) * 100, "00") & " " & Parametres.EtiquettePourcent & Parametres.Etiquette_Chargement(8)
                If Clavier.Quitter Then
                    Monde.TerrainOptimise = False
                    ReDim TerrainSprite(0)
                    'Décharge les images.
                    Set FrmMoteur2D.picCarte.Picture = LoadPicture("")
                    FrmMoteur2D.picCarte.Height = 0: FrmMoteur2D.picCarte.Width = 0
                    Set FrmMoteur2D.picTemp.Picture = LoadPicture("")
                    FrmMoteur2D.picTemp.Height = 0: FrmMoteur2D.picTemp.Width = 0
                    Exit Sub
                End If
                'Sprite.Unlock RectFore
                'TerrainSprite(l).Unlock RectBack
                SavePicture FrmMoteur2D.picCarte.Image, App.Path & Chemin & "Carte" & o + 1 & Extension
                Set TerrainSprite(o) = dd.CreateSurfaceFromFile(App.Path & Chemin & "Carte" & o + 1 & Extension, AffTerrainSprite)
                Kill App.Path & Chemin & "Carte" & o + 1 & Extension
                Set FrmMoteur2D.picCarte.Picture = LoadPicture("")
                
                'k = 0
                o = o + 1
            Next n
        Next m
    Next l
    'Décharge les images.
    Set FrmMoteur2D.picCarte.Picture = LoadPicture("")
    FrmMoteur2D.picCarte.Height = 0: FrmMoteur2D.picCarte.Width = 0
    Set FrmMoteur2D.picTemp.Picture = LoadPicture("")
    FrmMoteur2D.picTemp.Height = 0: FrmMoteur2D.picTemp.Width = 0
    Monde.TerrainOptimise = True
End Sub

Public Sub AfficheVeryFast(ByVal PositionX As Long, ByVal PositionY As Long, _
                           ByVal DecalageX As Long, ByVal DecalageY As Long, _
                           ByVal EcranWidth As Long, ByVal EcranHeight As Long, _
                           ByRef SurfBack As DirectDrawSurface7)
    'RectSprite.Left = DecalageX
    'RectSprite.Right = EcranWidth + DecalageX
    'RectSprite.Top = DecalageY
    'RectSprite.Bottom = EcranHeight + DecalageY
    'SurfBack.BltFast PositionX, PositionY, TerrainSprite(EtapeAnimation2), RectSprite, DDBLTFAST_SRCCOLORKEY
    Dim i As Long, j As Long, k As Long
    Dim cX As Long, cX2 As Long, cY As Long
    Dim eHH As Long: eHH = EcranHeight / OptimiserAffichageHauteur + 1
    Dim MondeLargeur As Long: MondeLargeur = Monde.Largeur
    Dim TpTerrain As Integer
    Dim CaTerrain As Integer
    
    'RectSprite.Bottom = Hauteur
        
    With SurfBack
    
    For i = 0 To EcranWidth / OptimiserAffichageLargeur + 1
        k = Int(DecalageY / OptimiserAffichageHauteur) * Int(Monde.Largeur / OptimiserAffichageLargeur * Largeur + 1) + i + Int(DecalageX / OptimiserAffichageLargeur) + _
                (Int((Monde.Hauteur * Hauteur / OptimiserAffichageHauteur) + 1) * Int((Monde.Largeur * Largeur / OptimiserAffichageLargeur) + 1) * EtapeAnimation2)
        'k = Int(DecalageY / Largeur) * MondeLargeur + i + Int(DecalageX / Hauteur)
        cX = i * OptimiserAffichageLargeur - DecalageX Mod OptimiserAffichageLargeur + PositionX
        cX2 = i * OptimiserAffichageLargeur - DecalageX Mod OptimiserAffichageLargeur
        For j = 0 To eHH
            If k <= UBound(TerrainSprite()) Then
                cY = j * OptimiserAffichageHauteur - DecalageY Mod OptimiserAffichageHauteur + PositionY
                RectSprite.Left = 0
                RectSprite.Right = OptimiserAffichageLargeur
                RectSprite.Top = 0
                RectSprite.Bottom = OptimiserAffichageHauteur
                
                If (cY + OptimiserAffichageHauteur) > ScreenHeight Then
                    'Dépasse en bas de l'écran.
                    RectSprite.Bottom = RectSprite.Bottom - (cY + OptimiserAffichageHauteur - ScreenHeight)
                End If
                If cY < PositionY Then
                    'Dépasse en haut de l'écran.
                    RectSprite.Top = RectSprite.Top - cY
                    cY = PositionY
                End If
    
                If (cX + OptimiserAffichageLargeur) > ScreenWidth Then
                    RectSprite.Right = RectSprite.Right - (cX + RectSprite.Right - RectSprite.Left - ScreenWidth)
                    .BltFast cX, cY, TerrainSprite(k), RectSprite, DDBLTFAST_SRCCOLORKEY
                ElseIf cX >= PositionX Then
                    .BltFast cX, cY, TerrainSprite(k), RectSprite, DDBLTFAST_SRCCOLORKEY
                Else
                    RectSprite.Left = RectSprite.Left - cX2
                    .BltFast PositionX, cY, TerrainSprite(k), RectSprite, DDBLTFAST_SRCCOLORKEY
                End If
            End If
            k = k + MondeLargeur * Largeur / OptimiserAffichageLargeur + 1
        Next j
        
    Next i
    
    End With
End Sub

Public Sub AfficheFast(ByVal PositionX As Long, ByVal PositionY As Long, _
                       ByVal DecalageX As Long, ByVal DecalageY As Long, _
                       ByVal EcranWidth As Long, ByVal EcranHeight As Long, _
                       ByRef SurfBack As DirectDrawSurface7)

    Dim i As Long, j As Long, k As Long
    Dim cX As Long, cX2 As Long, cY As Long
    Dim eHH As Long: eHH = EcranHeight / Hauteur
    Dim MondeLargeur As Long: MondeLargeur = Monde.Largeur
    Dim TpTerrain As Integer
    Dim CaTerrain As Integer
    
    RectSprite.Bottom = Hauteur
        
    With SurfBack
    
    For i = 0 To EcranWidth / Largeur
        k = Int(DecalageY / Largeur) * MondeLargeur + i + Int(DecalageX / Hauteur)
        cX = i * Largeur - DecalageX Mod Largeur + PositionX
        cX2 = i * Largeur - DecalageX Mod Largeur
        For j = 0 To eHH
            cY = j * Hauteur - DecalageY Mod Hauteur + PositionY
            TpTerrain = Monde.Param_TerrainApparence(k)
            CaTerrain = Monde.Param_TerrainCategorie(k)
            'Animation (dans le cas de l'eau).
            If Monde.Param_TerrainAnimation(k) Then CaTerrain = CaTerrain + EtapeAnimation2
            RectSprite.Left = Largeur * TpTerrain
            RectSprite.Right = Largeur * (TpTerrain + 1)
            RectSprite.Top = Hauteur * CaTerrain
            RectSprite.Bottom = Hauteur * (CaTerrain + 1)
            
            If (cY + Hauteur) > ScreenHeight Then
                'Dépasse en bas de l'écran.
                RectSprite.Bottom = RectSprite.Bottom - (cY + Hauteur - ScreenHeight)
            ElseIf cY < PositionY Then
                'Dépasse en haut de l'écran.
                RectSprite.Top = RectSprite.Top - cY
                cY = PositionY
            End If

            If (cX + Largeur) > ScreenWidth Then
                RectSprite.Right = RectSprite.Right - (cX + RectSprite.Right - RectSprite.Left - ScreenWidth)
                .BltFast cX, cY, Sprite, RectSprite, DDBLTFAST_SRCCOLORKEY
            ElseIf cX >= PositionX Then
                .BltFast cX, cY, Sprite, RectSprite, DDBLTFAST_SRCCOLORKEY
            Else
                RectSprite.Left = RectSprite.Left - cX2
                .BltFast PositionX, cY, Sprite, RectSprite, DDBLTFAST_SRCCOLORKEY
            End If
            
            k = k + MondeLargeur
        Next j
        
    Next i
    
    End With
End Sub
                       
                       

Private Sub Class_Terminate()
    Set Sprite = Nothing
End Sub
