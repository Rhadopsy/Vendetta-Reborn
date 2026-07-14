VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPerEcran"
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

Public FichierINI As String
Public SectionINI As String

Private ModeFenetre As Boolean
Public Resolution As Integer
Private ResolutionX() As Integer
Private ResolutionY() As Integer

Private Largeur As Long
Private Hauteur As Long

Private Sub Class_Initialize()
    On Error GoTo Erreur
    Dim i As Integer
    Dim NombreResolutions As Integer
    
    FichierINI = "Ecran"
    SectionINI = "Ecran"
    
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    ModeFenetre = Val(FicIni.Parametre("ModeFenetre"))
    NombreResolutions = Val(FicIni.Parametre("NombreResolution"))
    Resolution = Val(FicIni.Parametre("Resolution")) - 1
    
    ReDim ResolutionX(NombreResolutions - 1)
    ReDim ResolutionY(NombreResolutions - 1)
    For i = 0 To NombreResolutions - 1
        FicIni.Section = "Resolution_" & i + 1
        ResolutionX(i) = FicIni.Parametre("Largeur")
        ResolutionY(i) = FicIni.Parametre("Hauteur")
    Next i
    
    Largeur = ResolutionX(Resolution)
    Hauteur = ResolutionY(Resolution)
    Exit Sub
Erreur:
    MsgBox "Erreur l'hors du chargement de " & FichierINI & ".ini", vbCritical, "Erreur"
    End
End Sub

Public Sub Init(ByRef dx As DirectX7, _
                ByRef dd As DirectDraw7, _
                ByRef di As DirectInput, _
                ByRef AnHWND As Long, _
                ByRef surfPrim As DirectDrawSurface7, _
                ByRef sdprime As DDSURFACEDESC2, _
                ByRef SurfBack As DirectDrawSurface7, _
                ByRef SurfTemp As DirectDrawSurface7, _
                ByRef sdback As DDSURFACEDESC2, _
                ByRef cback As DDSCAPS2, _
                ByRef MonFond As RECT, _
                ByRef ddClipper As DirectDrawClipper)

    'On Local Error Resume Next
    Err.Number = 0
    Set di = dx.DirectInputCreate()
    
    'FrmMoteur2D.Top = 0
    'FrmMoteur2D.Left = 0
    
    'If ModeFenetre Then
    If ModeFenetre Then
        FrmMoteur2D.BorderStyle = 1
        FrmMoteur2D.Caption = App.Title
        dd.SetCooperativeLevel FrmMoteur2D.hwnd, DDSCL_NORMAL
        FrmMoteur2D.Top = 0
        FrmMoteur2D.Left = 0
        FrmMoteur2D.Width = Largeur * 15 + 120
        FrmMoteur2D.Height = Hauteur * 15 + 160
        'Centre la fenêtre manuellement.
        FrmMoteur2D.Left = (Screen.Width - FrmMoteur2D.Width) / 2
        FrmMoteur2D.Top = (Screen.Height - FrmMoteur2D.Height) / 2
        'FrmMoteur2D.Picture1.Width = FrmMoteur2D.ScaleWidth
        'FrmMoteur2D.Picture1.Height = FrmMoteur2D.ScaleHeight
        'FrmMoteur2D.Picture1.ScaleHeight = Hauteur * 16
        'FrmMoteur2D.Picture1.ScaleWidth = Largeur * 16
    Else
        'FrmMoteur2D.MinButton = False
        'FrmMoteur2D.MaxButton = False
        FrmMoteur2D.BorderStyle = 0
        dd.SetCooperativeLevel FrmMoteur2D.hwnd, DDSCL_NORMAL
        FrmMoteur2D.Top = 0
        FrmMoteur2D.Left = 0
        Changer_Resolution dd
        FrmMoteur2D.Width = Largeur * 15
        FrmMoteur2D.Height = Hauteur * 15
        FrmMoteur2D.Picture1.Width = FrmMoteur2D.ScaleWidth
        FrmMoteur2D.Picture1.Height = FrmMoteur2D.ScaleHeight
    End If
    'FrmMoteur2D.Width = Largeur * 15
    'FrmMoteur2D.Height = Hauteur * 15
    'FrmMoteur2D.Width = Screen.Width
    'FrmMoteur2D.Height = Screen.Height
    'FrmMoteur2D.BorderStyle = 1
    'FrmMoteur2D.Picture1.Width = FrmMoteur2D.ScaleWidth
    'FrmMoteur2D.Picture1.Height = FrmMoteur2D.ScaleHeight
    'FrmMoteur2D.Picture1.ScaleHeight = Hauteur * 16
    'FrmMoteur2D.Picture1.ScaleWidth = Largeur * 16
    'Description de la surface Primaire.
    sdprime.lFlags = DDSD_CAPS
    sdprime.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    sdprime.lBackBufferCount = 1
    'Else
    '    dd.SetCooperativeLevel FrmMoteur2D.hwnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
    '    Changer_Resolution dd
    '    FrmMoteur2D.BorderStyle = 0
    '    'Description de la surface Primaire.
    '    sdprime.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    '    sdprime.ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_PRIMARYSURFACE Or DDSCAPS_VIDEOMEMORY
    '    sdprime.lBackBufferCount = 1
    'End If
    
    Set surfPrim = Nothing
    Set surfPrim = dd.CreateSurface(sdprime)
    
    'Backbuffer
    'If ModeFenetre Then
    sdback.lFlags = DDSD_CAPS
    sdback.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Set SurfBack = dd.CreateSurfaceFromFile(App.Path & "\Images\Fond" & Ecran.Width & ".bmp", sdback)
    Set SurfTemp = dd.CreateSurfaceFromFile(App.Path & "\Images\Fond" & Ecran.Width & ".bmp", sdback)
    
    'Set d3dDevice = d3d.CreateDevice("IID_IDirect3DHALDevice", SurfBack)
    
    'Set SurfBack = surfPrim.GetAttachedSurface(cback)
    Set ddClipper = dd.CreateClipper(0)
    ddClipper.SetHWnd FrmMoteur2D.Picture1.hwnd
    surfPrim.SetClipper ddClipper
    'surfPrim.Flip Nothing, DDFLIP_WAIT
    'Else
    '    cback.lCaps = DDSCAPS_BACKBUFFER Or DDSCAPS_LOCALVIDMEM
    '    'Set SurfBack = Nothing
    '    Set SurfBack = surfPrim.GetAttachedSurface(cback)
    '    SurfBack.GetSurfaceDesc sdback
    'End If
    
    MonFond.Left = 0
    MonFond.Top = 0
    MonFond.Right = Largeur
    MonFond.Bottom = Hauteur
    
    If Err.Number <> 0 Then
        Err.Clear
        dd.RestoreDisplayMode
        dd.SetCooperativeLevel AnHWND, DDSCL_NORMAL
        MsgBox "Une erreur a eu lieu durant l'initalisation de DirectDraw", vbCritical + vbOKOnly, App.Title
    End If
    
    If dd Is Nothing Then
        MsgBox "Echec de la création de la surface DirectDraw", vbCritical + vbOKOnly, App.Title
    End If
    
End Sub

Private Sub Changer_Resolution(ByRef dd As DirectDraw7)
    On Error GoTo Erreur
    dd.SetDisplayMode Largeur, Hauteur, 16, 0, DDSDM_DEFAULT       ',DDSDM_STANDARDVGAMODE
    Exit Sub
Erreur:
End Sub

Public Sub Femer_Ecran(ByRef dd As DirectDraw7, ByRef Fenetre As Form)
    On Error GoTo Erreur
    dd.RestoreDisplayMode
    dd.SetCooperativeLevel Fenetre.hwnd, DDSCL_NORMAL
    Exit Sub
Erreur:
End Sub

Public Property Get Width() As Long
    Width = Largeur
End Property

Public Property Get Height() As Long
    Height = Hauteur
End Property

Public Property Let Width(v As Long)
    Largeur = v
End Property
Public Property Let Height(v As Long)
    Hauteur = v
End Property

Public Property Get Nombre_Resolutions() As Integer
    Nombre_Resolutions = UBound(ResolutionX()) + 1
End Property
Public Property Get Definir_ResolutionX(ByVal Valeur As Integer) As Integer
    Definir_ResolutionX = ResolutionX(Valeur)
End Property
Public Property Get Definir_ResolutionY(ByVal Valeur As Integer) As Integer
    Definir_ResolutionY = ResolutionY(Valeur)
End Property
Public Property Let Definir_Resolution(ByVal Valeur As Integer)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    Resolution = Valeur
    
    FicIni.Parametre("Resolution") = Valeur + 1
End Property

Public Property Get Definir_ModeFenetre() As Boolean
    Definir_ModeFenetre = ModeFenetre
End Property
Public Property Let Definir_ModeFenetre(ByVal Valeur As Boolean)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    ModeFenetre = Valeur
    
    FicIni.Parametre("ModeFenetre") = -CDbl(Valeur)
End Property


