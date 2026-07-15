Attribute VB_Name = "ModCompat"
Option Explicit

' Centralise les composants COM facultatifs et les adaptateurs nécessaires au
' port twinBASIC. Les helpers DirectDraw utilisent les interfaces natives de
' WinDevLib et ne dépendent pas de dx7vb.dll.

Private Const COMPAT_IMAGE_BITMAP As Long = 0
Private Const COMPAT_LR_LOADFROMFILE As Long = &H10
Private Const COMPAT_LR_CREATEDIBSECTION As Long = &H2000
Private Const COMPAT_SRCCOPY As Long = &HCC0020
Private Const COMPAT_TRANSPARENT As Long = 1
Private Const COMPAT_CLR_INVALID As Long = -1

Private Type COMPAT_BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As LongPtr
End Type

Private Declare PtrSafe Function CompatLoadImageW Lib "user32" Alias "LoadImageW" ( _
    ByVal hInst As LongPtr, _
    ByVal lpszName As LongPtr, _
    ByVal uType As Long, _
    ByVal cxDesired As Long, _
    ByVal cyDesired As Long, _
    ByVal fuLoad As Long) As LongPtr

Private Declare PtrSafe Function CompatGetObjectW Lib "gdi32" Alias "GetObjectW" ( _
    ByVal hObject As LongPtr, _
    ByVal cbBuffer As Long, _
    ByRef lpvObject As COMPAT_BITMAP) As Long

Private Declare PtrSafe Function CompatCreateCompatibleDC Lib "gdi32" Alias "CreateCompatibleDC" ( _
    ByVal hdc As LongPtr) As LongPtr

Private Declare PtrSafe Function CompatDeleteDC Lib "gdi32" Alias "DeleteDC" ( _
    ByVal hdc As LongPtr) As Long

Private Declare PtrSafe Function CompatDeleteObject Lib "gdi32" Alias "DeleteObject" ( _
    ByVal hObject As LongPtr) As Long

Private Declare PtrSafe Function CompatSelectObject Lib "gdi32" Alias "SelectObject" ( _
    ByVal hdc As LongPtr, _
    ByVal hObject As LongPtr) As LongPtr

Private Declare PtrSafe Function CompatBitBlt Lib "gdi32" Alias "BitBlt" ( _
    ByVal hdcDest As LongPtr, _
    ByVal xDest As Long, _
    ByVal yDest As Long, _
    ByVal width As Long, _
    ByVal height As Long, _
    ByVal hdcSource As LongPtr, _
    ByVal xSource As Long, _
    ByVal ySource As Long, _
    ByVal rasterOperation As Long) As Long

Private Declare PtrSafe Function CompatGetPixel Lib "gdi32" Alias "GetPixel" ( _
    ByVal hdc As LongPtr, _
    ByVal X As Long, _
    ByVal Y As Long) As Long

Private Declare PtrSafe Function CompatSetPixelV Lib "gdi32" Alias "SetPixelV" ( _
    ByVal hdc As LongPtr, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal color As Long) As Long

Private Declare PtrSafe Function CompatSetTextColor Lib "gdi32" Alias "SetTextColor" ( _
    ByVal hdc As LongPtr, _
    ByVal color As Long) As Long

Private Declare PtrSafe Function CompatSetBkMode Lib "gdi32" Alias "SetBkMode" ( _
    ByVal hdc As LongPtr, _
    ByVal mode As Long) As Long

Private Declare PtrSafe Function CompatTextOutW Lib "gdi32" Alias "TextOutW" ( _
    ByVal hdc As LongPtr, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal textPointer As LongPtr, _
    ByVal characterCount As Long) As Long

Public Function CompatCreerMoteurScript() As Object
    On Error Resume Next
    Set CompatCreerMoteurScript = CreateObject("MSScriptControl.ScriptControl")
    If Not CompatCreerMoteurScript Is Nothing Then
        CompatCreerMoteurScript.Language = "VBScript"
        CompatCreerMoteurScript.Reset
        CompatCreerMoteurScript.AllowUI = False
    End If
    On Error GoTo 0
End Function

Public Function CompatCreerDocumentXML() As Object
    On Error Resume Next
    Set CompatCreerDocumentXML = CreateObject("MSXML2.DOMDocument.6.0")
    If CompatCreerDocumentXML Is Nothing Then
        Set CompatCreerDocumentXML = CreateObject("MSXML2.DOMDocument.3.0")
    End If
    On Error GoTo 0
End Function

Private Function CompatCreerRequeteHttp() As Object
    On Error Resume Next
    Set CompatCreerRequeteHttp = CreateObject("MSXML2.XMLHTTP.6.0")
    If CompatCreerRequeteHttp Is Nothing Then
        Set CompatCreerRequeteHttp = CreateObject("MSXML2.XMLHTTP.3.0")
    End If
    On Error GoTo 0
End Function

Public Function CompatHttpGet(ByVal URL As String) As String
    Dim Requete As Object

    On Error GoTo Erreur
    Set Requete = CompatCreerRequeteHttp()
    If Requete Is Nothing Then Exit Function
    Requete.Open "GET", URL, False
    Requete.Send

    If Requete.Status >= 200 And Requete.Status < 300 Then
        CompatHttpGet = Requete.responseText
    End If

    Set Requete = Nothing
    Exit Function

Erreur:
    CompatHttpGet = ""
    Set Requete = Nothing
End Function

Public Function CompatCreateSurfaceFromBitmap( _
    ByRef dd As IDirectDraw7, _
    ByVal filePath As String, _
    Optional ByVal colorKey As Long = 0) As IDirectDrawSurface7

    Dim bitmapHandle As LongPtr
    Dim sourceDC As LongPtr
    Dim destinationDC As LongPtr
    Dim previousBitmap As LongPtr
    Dim bitmapInfo As COMPAT_BITMAP
    Dim surfaceDescription As DDSURFACEDESC2
    Dim surface As IDirectDrawSurface7
    Dim key As DDCOLORKEY
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Erreur

    bitmapHandle = CompatLoadImageW(0, StrPtr(filePath), COMPAT_IMAGE_BITMAP, 0, 0, _
                                    COMPAT_LR_LOADFROMFILE Or COMPAT_LR_CREATEDIBSECTION)
    If bitmapHandle = 0 Then
        Err.Raise 53, "CompatCreateSurfaceFromBitmap", "Bitmap introuvable : " & filePath
    End If

    If CompatGetObjectW(bitmapHandle, LenB(bitmapInfo), bitmapInfo) = 0 Then
        Err.Raise vbObjectError + 7100, "CompatCreateSurfaceFromBitmap", _
                  "Impossible de lire les dimensions du bitmap : " & filePath
    End If

    surfaceDescription.dwSize = LenB(surfaceDescription)
    surfaceDescription.dwFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    surfaceDescription.dwWidth = bitmapInfo.bmWidth
    surfaceDescription.dwHeight = bitmapInfo.bmHeight
    surfaceDescription.ddsCaps.dwCaps = DDSCAPS_OFFSCREENPLAIN

    dd.CreateSurface surfaceDescription, surface, Nothing
    If surface Is Nothing Then
        Err.Raise vbObjectError + 7101, "CompatCreateSurfaceFromBitmap", _
                  "DirectDraw n'a pas créé la surface : " & filePath
    End If

    sourceDC = CompatCreateCompatibleDC(0)
    If sourceDC = 0 Then
        Err.Raise vbObjectError + 7102, "CompatCreateSurfaceFromBitmap", _
                  "Impossible de créer le contexte mémoire GDI."
    End If

    previousBitmap = CompatSelectObject(sourceDC, bitmapHandle)
    surface.GetDC destinationDC

    If CompatBitBlt(destinationDC, 0, 0, bitmapInfo.bmWidth, bitmapInfo.bmHeight, _
                    sourceDC, 0, 0, COMPAT_SRCCOPY) = 0 Then
        Err.Raise vbObjectError + 7103, "CompatCreateSurfaceFromBitmap", _
                  "Impossible de copier le bitmap vers la surface DirectDraw."
    End If

    surface.ReleaseDC destinationDC
    destinationDC = 0

    key.dwColorSpaceLowValue = colorKey
    key.dwColorSpaceHighValue = colorKey
    surface.SetColorKey DDCKEY_SRCBLT, key

    Set CompatCreateSurfaceFromBitmap = surface

Nettoyage:
    On Error Resume Next
    If destinationDC <> 0 Then surface.ReleaseDC destinationDC
    If previousBitmap <> 0 Then CompatSelectObject sourceDC, previousBitmap
    If sourceDC <> 0 Then CompatDeleteDC sourceDC
    If bitmapHandle <> 0 Then CompatDeleteObject bitmapHandle
    On Error GoTo 0

    If errorNumber <> 0 Then
        Err.Raise errorNumber, "CompatCreateSurfaceFromBitmap", errorDescription
    End If
    Exit Function

Erreur:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Resume Nettoyage
End Function

Public Sub CompatSurfaceDrawText( _
    ByRef surface As IDirectDrawSurface7, _
    ByRef fontObject As IFont, _
    ByVal color As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal texte As String)

    Dim deviceContext As LongPtr
    Dim previousFont As LongPtr
    Dim previousColor As Long
    Dim previousBkMode As Long

    On Error GoTo Nettoyage

    surface.GetDC deviceContext
    previousFont = CompatSelectObject(deviceContext, fontObject.hFont)
    previousColor = CompatSetTextColor(deviceContext, color)
    previousBkMode = CompatSetBkMode(deviceContext, COMPAT_TRANSPARENT)
    CompatTextOutW deviceContext, X, Y, StrPtr(texte), Len(texte)

Nettoyage:
    On Error Resume Next
    If deviceContext <> 0 Then
        If previousBkMode <> 0 Then CompatSetBkMode deviceContext, previousBkMode
        If previousColor <> COMPAT_CLR_INVALID Then CompatSetTextColor deviceContext, previousColor
        If previousFont <> 0 Then CompatSelectObject deviceContext, previousFont
        surface.ReleaseDC deviceContext
    End If
    On Error GoTo 0
End Sub

Public Sub CompatSurfaceBlendColorKey( _
    ByRef destination As IDirectDrawSurface7, _
    ByRef source As IDirectDrawSurface7, _
    ByRef sourceRect As RECT, _
    ByVal destinationX As Long, _
    ByVal destinationY As Long, _
    ByVal alpha As Single, _
    Optional ByVal colorKey As Long = 0)

    Dim sourceDC As LongPtr
    Dim destinationDC As LongPtr
    Dim destinationDescription As DDSURFACEDESC2
    Dim X As Long
    Dim Y As Long
    Dim targetX As Long
    Dim targetY As Long
    Dim sourceColor As Long
    Dim destinationColor As Long
    Dim redValue As Long
    Dim greenValue As Long
    Dim blueValue As Long
    Dim blendedColor As Long

    If alpha < 0 Then alpha = 0
    If alpha > 1 Then alpha = 1

    On Error GoTo Nettoyage

    destinationDescription.dwSize = LenB(destinationDescription)
    destination.GetSurfaceDesc destinationDescription
    source.GetDC sourceDC
    destination.GetDC destinationDC

    For Y = 0 To sourceRect.Bottom - sourceRect.Top - 1
        targetY = destinationY + Y
        If targetY >= 0 And targetY < destinationDescription.dwHeight Then
            For X = 0 To sourceRect.Right - sourceRect.Left - 1
                targetX = destinationX + X
                If targetX >= 0 And targetX < destinationDescription.dwWidth Then
                    sourceColor = CompatGetPixel(sourceDC, sourceRect.Left + X, sourceRect.Top + Y)
                    If sourceColor <> COMPAT_CLR_INVALID And sourceColor <> colorKey Then
                        destinationColor = CompatGetPixel(destinationDC, targetX, targetY)
                        If destinationColor <> COMPAT_CLR_INVALID Then
                            redValue = CLng((sourceColor And &HFF&) * alpha + _
                                            (destinationColor And &HFF&) * (1 - alpha))
                            greenValue = CLng(((sourceColor \ &H100&) And &HFF&) * alpha + _
                                              ((destinationColor \ &H100&) And &HFF&) * (1 - alpha))
                            blueValue = CLng(((sourceColor \ &H10000) And &HFF&) * alpha + _
                                             ((destinationColor \ &H10000) And &HFF&) * (1 - alpha))
                            blendedColor = redValue Or (greenValue * &H100&) Or (blueValue * &H10000)
                            CompatSetPixelV destinationDC, targetX, targetY, blendedColor
                        End If
                    End If
                End If
            Next X
        End If
    Next Y

Nettoyage:
    On Error Resume Next
    If destinationDC <> 0 Then destination.ReleaseDC destinationDC
    If sourceDC <> 0 Then source.ReleaseDC sourceDC
    On Error GoTo 0
End Sub
