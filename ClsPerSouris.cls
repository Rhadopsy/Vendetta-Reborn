VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPerSouris"
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

Const NombreBoutons = 3

'Si on bascule hors du programme, cette propriété passe à faux.
Public Actif As Boolean

'Const VitesseClic = 0.17 'En secondes.
Private TempRelachement() As Boolean
Private Relachement() As Boolean 'Mémorise l'état du clique précédent.
Private BoutonRelache() As Boolean

'Driver Souris.
Private diDEVMouse As DirectInputDevice
Private mState As DIMOUSESTATE
'Private DernierClic As Double
Private TempClique() As Boolean

Private Sub Class_Initialize()
    On Error GoTo Erreur
    ReDim TempRelachement(NombreBoutons - 1)
    ReDim Relachement(NombreBoutons - 1)
    ReDim BoutonRelache(NombreBoutons - 1)
    ReDim TempClique(NombreBoutons - 1)
    'Effacer_Relachement
    Exit Sub
Erreur:
    MsgBox "Erreur l'hors de l'initialisation de la classe ClsPerSouris.cls.", vbCritical, "Erreur"
    End
End Sub

Public Sub Init(ByRef di As DirectInput, _
                ByRef dd As DirectDraw7, _
                ByRef Fenetre As Form, _
                ByRef X As Single, _
                ByRef Y As Single, _
                ByRef Ecran As ClsPerEcran)
    'Initialisation de la souris.
    On Error GoTo Erreur
    Set diDEVMouse = di.CreateDevice("GUID_SysMouse")
    diDEVMouse.SetCommonDataFormat DIFORMAT_MOUSE
    If Ecran.Definir_ModeFenetre Then
        diDEVMouse.SetCooperativeLevel Fenetre.hwnd, DISCL_NONEXCLUSIVE Or DISCL_FOREGROUND
    Else
        diDEVMouse.SetCooperativeLevel Fenetre.hwnd, DISCL_EXCLUSIVE Or DISCL_FOREGROUND
    End If
    diDEVMouse.Acquire 'Acquisition de la souris.
    Actif = True
    Exit Sub
Erreur:
    Actif = False
End Sub

Public Property Get Definir_Clique(ByVal NumeroBouton As Integer, ByVal ActiverClique As Boolean) As Boolean
    TempClique(NumeroBouton) = ActiverClique
    ActualiserClique NumeroBouton
    Definir_Clique = TempClique(NumeroBouton)
End Property

Public Function Clique(ByVal NumeroBouton As Integer) As Boolean
    TempClique(NumeroBouton) = mState.buttons(NumeroBouton) <> 0 'And Not mState.buttons(NumeroBouton) <> 128
    ActualiserClique NumeroBouton
    Clique = TempClique(NumeroBouton)
End Function

Public Sub ActualiserClique(ByVal NumeroBouton As Integer)
    BoutonRelache(NumeroBouton) = Not (TempClique(NumeroBouton) Or Relachement(NumeroBouton))
    TempRelachement(NumeroBouton) = Relachement(NumeroBouton)
    Relachement(NumeroBouton) = Not TempClique(NumeroBouton)
    If Not TempRelachement(NumeroBouton) Then
        TempClique(NumeroBouton) = False
    End If
End Sub

Public Function Maintenu_Enfonce(ByVal NumeroBouton As Integer) As Boolean
    If Not TempRelachement(NumeroBouton) And _
       Not Relachement(NumeroBouton) Then
        Maintenu_Enfonce = True
    End If
End Function

Public Function Bouton_Relache(ByVal NumeroBouton As Integer) As Boolean
    'Bouton_Relache = Not TempRelachement(NumeroBouton) And Relachement(NumeroBouton)
    Bouton_Relache = BoutonRelache(NumeroBouton)
    'Debug.Print TempRelachement2(NumeroBouton) & ", " & TempRelachement(NumeroBouton) & ", " & Relachement(NumeroBouton)
End Function

'Public Sub Effacer_Relachement()
'    Dim i As Integer
'    For i = 0 To NombreBoutons - 1
'        Relachement(i) = True
'        'TempRelachement(i) = True
'    Next i
'End Sub

Public Sub Centrer_curseur(ByRef X As Single, _
                           ByRef Y As Single, _
                           ByRef l As Integer, _
                           ByRef H As Integer)
    X = l / 2
    Y = H / 2
End Sub

Public Sub Positions(ByRef di As DirectInput, _
                     ByRef dd As DirectDraw7, _
                     ByRef Fenetre As Form, _
                     ByRef X As Single, _
                     ByRef Y As Single, _
                     ByRef Z As Single, _
                     ByRef MinX As Integer, _
                     ByRef MaxX As Integer, _
                     ByRef MinY As Integer, _
                     ByRef MaxY As Integer)
    On Error GoTo Erreur
    If Ecran.Definir_ModeFenetre Then
        'FrmMoteur2D.Form_MouseMove 0, 0, X, Y
    Else
        diDEVMouse.GetDeviceStateMouse mState
        X = X + mState.X
        Y = Y + mState.Y
        Z = Z + mState.Z
            
        If X < MinX Then X = MinX
        If X > MaxX Then X = MaxX
        If Y < MinY Then Y = MinY
        If Y > MaxY Then Y = MaxY
    End If
    Exit Sub
Erreur:
    Init di, dd, Fenetre, X, Y, Ecran
End Sub

Private Sub Class_Terminate()
    diDEVMouse.Unacquire
    Set diDEVMouse = Nothing
End Sub
