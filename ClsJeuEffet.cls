VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsJeuEffet"
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

Public Indice As Integer
Public PositionX As Long
Public PositionY As Long
Public Duree As Double
Public NombreEtapes As Integer
Public Temps As Double
Public Actif As Boolean

Public Sub Activer_Effet(ByVal TypeEffet As Integer, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal LargeurCible As Long, _
                         ByVal HauteurCible As Long, _
                         ByVal Parametres As ClsJeuParam, _
                         Optional ByVal ForcerEffet As Boolean)
    If Not Actif Or ForcerEffet Then
        Indice = TypeEffet
        PositionX = X + (LargeurCible - Parametres.Effet_Largeur(Indice)) / 2
        PositionY = Y + (HauteurCible - Parametres.Effet_Hauteur(Indice)) / 2
        Duree = Parametres.Effet_DureeEtape(Indice)
        NombreEtapes = Parametres.Effet_NombreEtapes(Indice)
        Temps = 0 'Timer
        Actif = True
    End If
End Sub

'Public Property Get Actif() As Boolean
    'Actif = Timer < Temps + Duree * (NombreEtapes - 1)
'    Actif = Abs(Timer - Temps) < (NombreEtapes - 1) * Duree
    'Debug.Print "Timer : " & Timer & " | " & _
                "Temps : " & Temps & " | " & _
                "Duree : " & Duree & " | " & _
                "Res : " & Actif
'End Property

Public Property Get Etape_Effet() As Integer
    'Dim Res As Long
    'On Error GoTo Erreur
    'Dim i As Integer
    'For i = 0 To NombreEtapes - 1
    '    If Timer < Temps + Duree * i Then
    '        Etape_Effet = i
    '        i = NombreEtapes
    '    End If
    'Next i
    'Debug.Print "Timer : " & Timer & " | " & _
                "Temps : " & Temps & " | " & _
                "Duree : " & Duree & " | " & _
                "Resul : " & Abs(Timer - Temps) / Duree
'    Res = Abs(Timer - Temps) / Duree
'    If Res < NombreEtapes Then
'        Etape_Effet = Res
'    End If
    Etape_Effet = Temps
    'Exit Property
'Erreur:
    'Debug.Print Err.Description
    'Exit Property
End Property

Public Sub Etape_Suivante()
    If Actif Then
        Temps = Temps + 1 / Duree
        Actif = Temps < NombreEtapes
    End If
End Sub
