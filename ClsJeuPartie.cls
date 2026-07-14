VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsJeuPartie"
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

Const FichierINI = "Partie"
Const SectionINI = "Partie"

'Cette classe contient toutes les informations de la partie.
Public IndiceMode As Integer
Public RessourcesDeDepart As Integer
Public EnCours As Integer '0: en attente. 1: Chargement en cours, 2: Partie en cours.
Private Modes_Nom() As String
Private EnCours_Nom() As String
Public FichierSauvegarde As String

Private Sub Class_Initialize()
    Dim Temp As String
    Dim i As Integer
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    
    i = 0
    Temp = FicIni.Parametre("Mode" & i + 1)
    While Temp <> ""
        ReDim Preserve Modes_Nom(i)
        Modes_Nom(i) = Temp
        i = i + 1
        Temp = FicIni.Parametre("Mode" & i + 1)
    Wend
    
    ReDim EnCours_Nom(2)
    EnCours_Nom(0) = FicIni.Parametre("EnCours1")
    EnCours_Nom(1) = FicIni.Parametre("EnCours2")
    EnCours_Nom(2) = FicIni.Parametre("EnCours3")
End Sub

Public Property Get Definir_Info_Partie() As String
    If Jeu.Regle_ModeConfrontation Then
        Definir_Info_Partie = Modes_Nom(5) & vbCrLf & vbCrLf & EnCours_Nom(EnCours)
    Else
        Definir_Info_Partie = Modes_Nom(IndiceMode) & vbCrLf & _
                              IIf(FichierSauvegarde = "", "", " - " & FichierSauvegarde & vbCrLf) & vbCrLf & _
                              EnCours_Nom(EnCours)
    End If
End Property
