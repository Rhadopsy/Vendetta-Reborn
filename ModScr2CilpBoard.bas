Attribute VB_Name = "ModScr2CilpBoard"
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
'-----------------------------------------------------------------------------------------------------
' Module Scr2CilpBoard : copie d'écran dans le presse papier
' Créé par LEVEQUE Franck Le 28/10/2003
' Références ayant servi à la création de ce module :
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/WinUI/WindowsUserInterface/DataExchange/Clipboard/UsingtheClipboard.asp
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/gdi/bitmaps_5a5h.asp
'
' Fonction à appeller : GrabScreen()
'
' ----------------------------------------------------------------------------------------------------

' Déclarations API nécessaires à la capture d'écran

'Constantes
Public Const SRCCOPY = &HCC0020     ' (DWORD) dest = source
Public Const HORZRES = 8            '  Horizontal width in pixels
Public Const VERTRES = 10           '  Vertical width in pixels
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32
'Types
Public Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type
'Fonctions
#If TWINBASIC Then
Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Public Declare PtrSafe Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As LongPtr
Public Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
Public Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
Public Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Public Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
#Else
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
#End If

' Déclarations API nécessaires au placement d'information dans le Presse Papier

Public Const CF_BITMAP = 2

#If TWINBASIC Then
Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Public Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
#Else
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
#End If

' Cette fonction place une image BITMAP dans le presse-papier
#If TWINBASIC Then
Public Function putInClipboard(hBitMap As LongPtr)
#Else
Public Function putInClipboard(hBitMap As Long)
#End If
    OpenClipboard FrmMoteur2D.hwnd ' Remplace form1.hwnd par ton pointeur d'écran
    EmptyClipboard
    SetClipboardData CF_BITMAP, hBitMap
    CloseClipboard
End Function

' Cette fonction produit une image bitmap et la place dans le presse papier
Public Sub GrabScreen()
#If TWINBASIC Then
Dim hdcScreen As LongPtr, hdcCompatible As LongPtr, hbmScreen As LongPtr
#Else
Dim hdcScreen As Long, hdcCompatible As Long, hbmScreen As Long
#End If
Dim nullStr As String, nullDEVMODE As DEVMODE
hdcScreen = GetWindowDC(0)
hdcCompatible = CreateCompatibleDC(hdcScreen)
hbmScreen = CreateCompatibleBitmap(hdcScreen, _
                     GetDeviceCaps(hdcScreen, HORZRES), _
                     GetDeviceCaps(hdcScreen, VERTRES))
If hbmScreen = 0 Then
    'MsgBox hbmScreen
End If
 
If SelectObject(hdcCompatible, hbmScreen) = 0 Then
    'MsgBox "Incompatible Bitmap Selection"
End If
If BitBlt(hdcCompatible, 0, 0, GetDeviceCaps(hdcScreen, HORZRES), GetDeviceCaps(hdcScreen, VERTRES), hdcScreen, 0, 0, SRCCOPY) = 0 Then
    'MsgBox "Transfer Failed"
End If
putInClipboard hbmScreen
'MsgBox "Capture terminée"
End Sub


