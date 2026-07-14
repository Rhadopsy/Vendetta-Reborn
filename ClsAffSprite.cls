VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAffSprite"
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

'Affichage Directx.
Private Sprite As DirectDrawSurface7 'creation surface Primaire.
Private AffSprite As DDSURFACEDESC2  'description surface Primaire.
Private RectSprite As RECT           'rectangle Titre.
Private RectSprite2 As RECT           'rectangle Titre.

Private DDBLTFX As DDBLTFX

Const Chemin = "\Images\"
Const Extension = ".bmp"
Const SngAlpha As Single = 0.5

Dim i As Long
Dim j As Long
Dim IColorS As Long
Dim IColorB As Long
Dim iNewColor As Long
Dim RectBack As RECT
Dim AffBack As DDSURFACEDESC2
Dim AffFore As DDSURFACEDESC2
Dim RectFore As RECT

Private Sub Class_Initialize()
    DDBLTFX.ddckDestColorKey_high = 0
    DDBLTFX.ddckDestColorKey_low = 0
    DDBLTFX.ddckSrcColorKey_high = 0
    DDBLTFX.ddckSrcColorKey_low = 0
End Sub

Public Sub Charger(ByRef dd As DirectDraw7, _
                   ByVal Fichier As String, _
                   Optional ByVal CheminAbsolu As Boolean, _
                   Optional ByVal GestionErreurs As Boolean = True)
    On Local Error GoTo Erreur
    'EcranLargeur = ScreenWidth
    'EcranHauteur = ScreenHeight
    AffSprite.lFlags = DDSD_CAPS Or DDSD_CKSRCBLT
    AffSprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    'AffSprite.lWidth =
    'AffSprite.lHeight =
    If CheminAbsolu Then
        Set Sprite = dd.CreateSurfaceFromFile(App.Path & "\" & Fichier & Extension, AffSprite)
    Else
        Set Sprite = dd.CreateSurfaceFromFile(App.Path & Chemin & Fichier & Extension, AffSprite)
    End If
    Exit Sub
Erreur:
    If GestionErreurs Then
        'Debug.Print "Le fichier image :" & vbCrLf & App.Path & Chemin & Fichier & Extension & vbCrLf & "est manquant.", vbCritical
        If CheminAbsolu Then
            MsgBox "Le fichier image :" & vbCrLf & App.Path & "\" & Fichier & Extension & vbCrLf & "ne peux être chargé." & vbCrLf & Err.Description, vbCritical
        Else
            MsgBox "Le fichier image :" & vbCrLf & App.Path & Chemin & Fichier & Extension & vbCrLf & "ne peux être chargé." & vbCrLf & Err.Description, vbCritical
        End If
        FrmMoteur2D.FatalEnd
    End If
End Sub

Public Sub Afficher(ByVal X As Long, _
                    ByVal Y As Long, _
                    ByRef SurfBack As DirectDrawSurface7, _
                    ByVal Transparent As Boolean, _
                    Optional ByVal Translucide As Boolean, _
                    Optional ByVal DecalageX As Integer = 0, _
                    Optional ByVal DecalageY As Integer = 0, _
                    Optional ByVal ValeurTansparence As Single = SngAlpha)
    With SurfBack
    
    If Translucide Then
        Sprite.GetSurfaceDesc AffFore
        Sprite.Lock RectFore, AffFore, DDLOCK_WAIT, 0
        .GetSurfaceDesc AffBack
        .Lock RectBack, AffBack, DDLOCK_WAIT, 0
        For i = 0 To RectSprite.Right - RectSprite.Left
            For j = 0 To RectSprite.Bottom - RectSprite.Top
                If i + X >= 0 And _
                   i + X <= ScreenWidth And _
                   j + Y >= 0 And _
                   j + Y <= ScreenHeight And _
                   Sprite.GetLockedPixel(i + RectSprite.Left, j + RectSprite.Top) > 0 Then
                    IColorS = Sprite.GetLockedPixel(i + RectSprite.Left, j + RectSprite.Top)
                    IColorB = SurfBack.GetLockedPixel(i + X - ScreenPositionX, j + Y - ScreenPositionY)
                    iNewColor = AffBack.ddpfPixelFormat.lRBitMask And (IColorS And AffFore.ddpfPixelFormat.lRBitMask) * ValeurTansparence + (IColorB And AffBack.ddpfPixelFormat.lRBitMask) * (1 - ValeurTansparence)
                    iNewColor = iNewColor Or _
                                AffBack.ddpfPixelFormat.lGBitMask And (IColorS And AffFore.ddpfPixelFormat.lGBitMask) * ValeurTansparence + (IColorB And AffBack.ddpfPixelFormat.lGBitMask) * (1 - ValeurTansparence)
                    iNewColor = iNewColor Or _
                                AffBack.ddpfPixelFormat.lBBitMask And (IColorS And AffFore.ddpfPixelFormat.lBBitMask) * ValeurTansparence + (IColorB And AffBack.ddpfPixelFormat.lBBitMask) * (1 - ValeurTansparence)
                    .SetLockedPixel i + X - ScreenPositionX, j + Y - ScreenPositionY, iNewColor
                End If
            Next j
        Next i
        Sprite.Unlock RectFore
        .Unlock RectBack
    Else
        If X < 0 Then
            RectSprite.Left = RectSprite.Left - X
            X = 0
        End If
        If DecalageX < 0 Then
            RectSprite.Left = RectSprite.Left - DecalageX
        End If
        If X + RectSprite.Right - RectSprite.Left > ScreenWidth Then
            RectSprite.Right = RectSprite.Right - (X + RectSprite.Right - RectSprite.Left - ScreenWidth)
        End If
        If DecalageX > 0 Then
            RectSprite.Right = RectSprite.Right - DecalageX
        End If
        If Y < 0 Then
            RectSprite.Top = RectSprite.Top - Y
            Y = 0
        End If
        If DecalageY < 0 Then
            RectSprite.Top = RectSprite.Top - DecalageY
        End If
        If Y + RectSprite.Bottom - RectSprite.Top > ScreenHeight Then
            RectSprite.Bottom = RectSprite.Bottom - (Y + RectSprite.Bottom - RectSprite.Top - ScreenHeight)
        End If
        If DecalageY > 0 Then
            RectSprite.Bottom = RectSprite.Bottom - DecalageY
        End If
        .BltFast X + DecalageX + ScreenPositionX, Y + DecalageY + ScreenPositionY, Sprite, RectSprite, IIf(Transparent, DDBLTFAST_SRCCOLORKEY, DDBLTFAST_NOCOLORKEY)
    End If
    
    End With
End Sub

Public Sub Afficher2(ByVal X As Long, _
                    ByVal Y As Long, _
                    ByRef SurfBack As DirectDrawSurface7, _
                    ByVal Transparent As Boolean, _
                    Optional ByVal Zoom As Single)
    With RectSprite
    RectSprite2.Top = RectSprite.Top
    RectSprite2.Bottom = RectSprite.Bottom
    RectSprite2.Left = RectSprite.Left
    RectSprite2.Right = RectSprite.Right
    RectSprite.Top = RectSprite.Top + Y
    RectSprite.Bottom = RectSprite.Bottom + Y - (RectSprite.Bottom - RectSprite.Bottom * Zoom)
    RectSprite.Left = RectSprite.Left + X
    RectSprite.Right = RectSprite.Right + X - (RectSprite.Right - RectSprite.Right * Zoom)
    SurfBack.Blt RectSprite, Sprite, RectSprite2, IIf(Transparent, DDBLT_KEYSRC, 0)
    End With
End Sub

Public Sub Selection(ByVal Haut As Long, _
                     ByVal Bas As Long, _
                     ByVal Gauche As Long, _
                     ByVal Droite As Long)
    RectSprite.Top = Haut
    RectSprite.Bottom = Bas
    RectSprite.Left = Gauche
    RectSprite.Right = Droite
End Sub

Public Property Get CheminImage() As String
    CheminImage = Chemin
End Property

Public Property Get Extension_Image() As String
    Extension_Image = Extension
End Property

Private Sub Class_Terminate()
    Set Sprite = Nothing
End Sub

