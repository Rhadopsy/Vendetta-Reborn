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

'Affichage DirectX.
Private Sprite As IDirectDrawSurface7
Private RectSprite As RECT
Private RectSprite2 As RECT

Const Chemin = "\Images\"
Const Extension = ".bmp"
Const SngAlpha As Single = 0.5

Public Sub Charger(ByRef dd As IDirectDraw7, _
                   ByVal Fichier As String, _
                   Optional ByVal CheminAbsolu As Boolean, _
                   Optional ByVal GestionErreurs As Boolean = True)
    Dim NomFichier As String

    On Local Error GoTo Erreur

    If CheminAbsolu Then
        NomFichier = App.Path & "\" & Fichier & Extension
    Else
        NomFichier = App.Path & Chemin & Fichier & Extension
    End If

    Set Sprite = CompatCreateSurfaceFromBitmap(dd, NomFichier, 0)
    Exit Sub

Erreur:
    If GestionErreurs Then
        MsgBox "Le fichier image :" & vbCrLf & NomFichier & vbCrLf & _
               "ne peux être chargé." & vbCrLf & Err.Description, vbCritical
        FrmMoteur2D.FatalEnd
    End If
End Sub

Public Sub Afficher(ByVal X As Long, _
                    ByVal Y As Long, _
                    ByRef SurfBack As IDirectDrawSurface7, _
                    ByVal Transparent As Boolean, _
                    Optional ByVal Translucide As Boolean, _
                    Optional ByVal DecalageX As Integer = 0, _
                    Optional ByVal DecalageY As Integer = 0, _
                    Optional ByVal ValeurTansparence As Single = SngAlpha)

    If Translucide Then
        CompatSurfaceBlendColorKey SurfBack, Sprite, RectSprite, _
                                   X - ScreenPositionX, _
                                   Y - ScreenPositionY, _
                                   ValeurTansparence, 0
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
        SurfBack.BltFast X + DecalageX + ScreenPositionX, _
                         Y + DecalageY + ScreenPositionY, _
                         Sprite, RectSprite, _
                         IIf(Transparent, DDBLTFAST_SRCCOLORKEY, DDBLTFAST_NOCOLORKEY)
    End If
End Sub

Public Sub Afficher2(ByVal X As Long, _
                     ByVal Y As Long, _
                     ByRef SurfBack As IDirectDrawSurface7, _
                     ByVal Transparent As Boolean, _
                     Optional ByVal Zoom As Single)
    Dim BltFx As DDBLTFX
    Dim BltFlags As DDBLTFLAGS

    With RectSprite
        RectSprite2.Top = RectSprite.Top
        RectSprite2.Bottom = RectSprite.Bottom
        RectSprite2.Left = RectSprite.Left
        RectSprite2.Right = RectSprite.Right
        RectSprite.Top = RectSprite.Top + Y
        RectSprite.Bottom = RectSprite.Bottom + Y - (RectSprite.Bottom - RectSprite.Bottom * Zoom)
        RectSprite.Left = RectSprite.Left + X
        RectSprite.Right = RectSprite.Right + X - (RectSprite.Right - RectSprite.Right * Zoom)

        BltFx.dwSize = LenB(BltFx)
        If Transparent Then
            BltFlags = DDBLT_KEYSRC
        Else
            BltFlags = 0
        End If
        SurfBack.Blt RectSprite, Sprite, RectSprite2, BltFlags, BltFx
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
