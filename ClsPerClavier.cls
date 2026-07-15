VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPerClavier"
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

Const VitesseCurseurXMin = 40 '4
Const VitesseCurseurXMax = 450 '30
Const VitesseCurseurXAcc = 7.5 '1.5 '0.75 'Accélération par seconde.
Const VitesseCurseurYMin = 40 '4
Const VitesseCurseurYMax = 450 '30
Const VitesseCurseurYAcc = 7.5 '1.5 '0.75 'Accélération par seconde.
Const VitesseCurseurZ = 5
Const NombreBoutons = 3
Private VitesseCurseurX As Double
Private VitesseCurseurY As Double

Private Temps As Double 'Avec le Timer, permet de garder une vitesse de curseur régulière.

Const VitessePression = 0.16 'En secondes.
Private DernierePressionTemps As Double
Private DernierePressionTouche As Integer

Private TempRelachement() As Boolean
Private Relachement() As Boolean 'Mémorise l'état du clique précédent.
Private BoutonRelache() As Boolean

Private didev As IDirectInputDevice7               'Driver Clavier
Private iKeyCounter  As Long
Private diState As DIKEYBOARDSTATE               'Test des touches Clavier

Private Sub Class_Initialize()
    On Error GoTo Erreur
    ReDim TempRelachement(NombreBoutons)
    ReDim Relachement(NombreBoutons)
    ReDim BoutonRelache(NombreBoutons)
    Exit Sub
Erreur:
    MsgBox "Erreur l'hors de l'initialisation de la classe ClsPerClavier.cls.", vbCritical, "Erreur"
    End
End Sub

Public Sub Init(ByRef di As IDirectInput7, _
                ByRef Fenetre As Form)
    'Initialisation du clavier.
    Set didev = di.CreateDevice("GUID_SysKeyboard")
    didev.SetCommonDataFormat DIFORMAT_KEYBOARD
    didev.SetCooperativeLevel Fenetre.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    didev.Acquire
    
    Temps = Timer
End Sub

Private Sub Class_Terminate()
    'Libération des périphériques d'entrée.
    didev.Unacquire
End Sub

Public Function Quitter() As Boolean
    didev.GetDeviceStateKeyboard diState
    Quitter = diState.Key(1) <> 0
End Function

Public Function Pression_Clavier(ByRef Ctrl As Boolean, _
                                 ByRef Shift As Boolean, _
                                 ByRef Alt As Boolean, _
                                 ByRef Haut As Boolean, _
                                 ByRef Bas As Boolean, _
                                 ByRef Gauche As Boolean, _
                                 ByRef Droite As Boolean) As Integer
    'Ctrl : 29.
    'Alt : 56.
    'Haut = False: Bas = False: Gauche = False: Droite = False
    didev.GetDeviceStateKeyboard diState
    For iKeyCounter = 0 To 255
        Select Case iKeyCounter
        Case 29, 157:
            Ctrl = diState.Key(29) <> 0 Or diState.Key(157) <> 0
        Case 42, 54:
            Shift = diState.Key(42) <> 0 Or diState.Key(54) <> 0
        Case 56, 184:
            Alt = diState.Key(56) <> 0 Or diState.Key(184) <> 0
        Case 200:
            Haut = diState.Key(200) <> 0
        Case 208:
            Bas = diState.Key(208) <> 0
        Case 203:
            Gauche = diState.Key(203) <> 0
        Case 205:
            Droite = diState.Key(205) <> 0
        Case Else:
            If diState.Key(iKeyCounter) <> 0 Then
                Pression_Clavier = iKeyCounter
                Exit For
            End If
        End Select
        'If diState.Key(iKeyCounter) <> 0 Then Messages(0).Ajouter_Message iKeyCounter
    Next iKeyCounter
    TempRelachement(NombreBoutons) = Relachement(NombreBoutons)
    Relachement(NombreBoutons) = Not (Pression_Clavier > 0)
    If Not TempRelachement(NombreBoutons) Then
        Pression_Clavier = 0
    End If
    'If Pression_Clavier <> 0 Then Messages(0).Ajouter_Message Pression_Clavier
    'Messages(0).Ajouter_Message Haut & " " & Bas & " " & Gauche & " " & Droite
End Function

Public Function Pression_ClavierDialogue(ByRef Ctrl As Boolean, _
                                         ByRef Shift As Boolean, _
                                         ByRef Alt As Boolean, _
                                         ByRef Haut As Boolean, _
                                         ByRef Bas As Boolean, _
                                         ByRef Gauche As Boolean, _
                                         ByRef Droite As Boolean) As Integer
    
    didev.GetDeviceStateKeyboard diState
    For iKeyCounter = 0 To 255
        Select Case iKeyCounter
        Case 29, 157:
            Ctrl = diState.Key(29) <> 0 Or diState.Key(157) <> 0
        Case 42, 54:
            Shift = diState.Key(42) <> 0 Or diState.Key(54) <> 0
        Case 56, 184:
            Alt = diState.Key(56) <> 0 Or diState.Key(184) <> 0
        Case Else
            If diState.Key(iKeyCounter) <> 0 Then
                Pression_ClavierDialogue = iKeyCounter
                Exit For
            End If
        End Select
    Next iKeyCounter
    If DernierePressionTemps + VitessePression < Timer Or Pression_ClavierDialogue <> DernierePressionTouche Then
        DernierePressionTouche = Pression_ClavierDialogue
        DernierePressionTemps = Timer
    Else
        Pression_ClavierDialogue = 0
    End If
End Function
' ************************************************************************************
' *      Fonctions pour utiliser le pavé numérique comme souris par le 2e joueur     *
' ************************************************************************************

Public Sub Positions(ByRef X As Single, _
                     ByRef Y As Single, _
                     ByRef Z As Single, _
                     ByRef MinX As Integer, _
                     ByRef MaxX As Integer, _
                     ByRef MinY As Integer, _
                     ByRef MaxY As Integer)
    'Fonction utilisée pour déplacer le curseur du deuxième utilisateur.
    Dim TempVitesseCurseur As Long
    Dim Temp As Double
    Temp = Timer - Temps
    If Temp > 1 / Jeu.Definir_Vitesse Then Temp = 1 / Jeu.Definir_Vitesse
    If diState.Key(75) <> 0 Or _
       diState.Key(76) <> 0 Or _
       diState.Key(77) <> 0 Or _
       diState.Key(72) <> 0 Or _
       diState.Key(80) <> 0 Then
        If diState.Key(75) <> 0 Or diState.Key(77) <> 0 Then
            VitesseCurseurX = VitesseCurseurX + VitesseCurseurXAcc
            TempVitesseCurseur = VitesseCurseurX * Temp
            If TempVitesseCurseur > VitesseCurseurXMax * Temp Then
                TempVitesseCurseur = VitesseCurseurXMax * Temp
            End If
            With Ecran
            If diState.Key(75) <> 0 Then
                'Gauche.
                X = X - TempVitesseCurseur
            ElseIf diState.Key(77) <> 0 Then
                'Droite.
                X = X + TempVitesseCurseur
            End If
            End With
        End If
        If diState.Key(72) <> 0 Or diState.Key(76) <> 0 Or diState.Key(80) <> 0 Then
            VitesseCurseurY = VitesseCurseurY + VitesseCurseurYAcc
            TempVitesseCurseur = VitesseCurseurY * Temp
            If TempVitesseCurseur > VitesseCurseurYMax * Temp Then
                TempVitesseCurseur = VitesseCurseurYMax * Temp
            End If
            With Ecran
            If diState.Key(72) <> 0 Then
                'Haut.
                Y = Y - TempVitesseCurseur
            ElseIf diState.Key(76) <> 0 Or diState.Key(80) <> 0 Then
                'Bas.
                Y = Y + TempVitesseCurseur
            End If
            End With
        End If
        'Messages(0).Ajouter_Message TempVitesseCurseur & " (" & Temp & ")"
    Else
        If diState.Key(81) <> 0 Or diState.Key(71) <> 0 Then
            'Scrolling haut.
            Z = Z - VitesseCurseurZ * Temp
        ElseIf diState.Key(73) <> 0 Or diState.Key(79) <> 0 Then
            'Scolling bas.
            Z = Z + VitesseCurseurZ * Temp
        End If
        VitesseCurseurX = VitesseCurseurXMin
        VitesseCurseurY = VitesseCurseurYMin
    End If
    If X < MinX Then
        X = MinX
    ElseIf X > MaxX Then
        X = MaxX
    End If
    If Y < MinY Then
        Y = MinY
    ElseIf Y > MaxY Then
        Y = MaxY
    End If
    Temps = Timer
End Sub

Public Function Clique(ByVal NumeroBouton As Integer) As Boolean
    Select Case NumeroBouton
    Case 0:
        Clique = diState.Key(82) <> 0
    Case 1:
        Clique = diState.Key(156) <> 0
    Case 2:
        Clique = diState.Key(83) <> 0
    End Select
    BoutonRelache(NumeroBouton) = Not (Clique Or Relachement(NumeroBouton))
    TempRelachement(NumeroBouton) = Relachement(NumeroBouton)
    Relachement(NumeroBouton) = Not Clique
    If Not TempRelachement(NumeroBouton) Then
        Clique = False
    End If
End Function

Public Function Maintenu_Enfonce(ByVal NumeroBouton As Integer) As Boolean
    If Not TempRelachement(NumeroBouton) And _
       Not Relachement(NumeroBouton) Then
        Maintenu_Enfonce = True
    End If
End Function

Public Function Bouton_Relache(ByVal NumeroBouton As Integer) As Boolean
    Bouton_Relache = BoutonRelache(NumeroBouton)
End Function
