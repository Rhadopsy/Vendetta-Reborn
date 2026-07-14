VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsJeuNoms"
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

Const FichierNomM = "NomsM.ini"
Const FichierNomF = "NomsF.ini"

Private LNomsM As Long
Private LNomsF As Long
Dim NomsM() As String
Dim NomsF() As String

Private Sub Class_Initialize()
    'Charge les fichiers de noms.
    On Error GoTo Erreur
    Dim i As Long
    Open App.Path & FicIni.Chemin & FichierNomM For Input As #1
    ReDim NomsM(0)
    Line Input #1, NomsM(0)
    While Not EOF(1)
        i = i + 1
        ReDim Preserve NomsM(i)
        Line Input #1, NomsM(i)
    Wend
    Close #1
    i = 0
    Open App.Path & FicIni.Chemin & FichierNomF For Input As #1
    ReDim NomsF(0)
    Line Input #1, NomsF(0)
    While Not EOF(1)
        i = i + 1
        ReDim Preserve NomsF(i)
        Line Input #1, NomsF(i)
    Wend
    Close #1
    LNomsM = UBound(NomsM())
    LNomsF = UBound(NomsF())
    Exit Sub
Erreur:
    MsgBox "Erreur l'hors du chargement des noms", vbCritical, "Erreur"
    End
End Sub

Public Function Tirer_Nom_Aleatoire(Optional ByVal Feminin As Boolean) As String
    'Tire un nom au hasard.
    If Feminin Then
        Tirer_Nom_Aleatoire = NomsF(Rnd * LNomsF)
    Else
        Tirer_Nom_Aleatoire = NomsM(Rnd * LNomsM)
    End If
End Function
