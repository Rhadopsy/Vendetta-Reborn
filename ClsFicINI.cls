VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFicINI"
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

Private JournalErreurActif As Boolean
Private JournalErreurFichier As String

Const FichierINI = "Donnees"
Const SectionINI = "Donnees"

'Fonction d'accés aux fichiers ini.
#If TWINBASIC Then
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
#Else
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If

'Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Public ExtensionFichierINI As String
Const ExtensionFichierINIDefaut = ".ini"
Const STRING_SIZE = 255

Public Fichier As String
Public Section As String
Public Chemin As String
Public CheminModsOfficiels As String
Public CheminMods As String

Public DernierChargement As String

Private lLength As Long
Private sReturned As String * STRING_SIZE

Private Sub Class_Initialize()
    'Section = App.Title
    ExtensionFichierINI = ExtensionFichierINIDefaut
    
    Fichier = FichierINI
    Section = SectionINI
    JournalErreurFichier = Parametre("JournalErreurFichier")
    JournalErreurActif = Parametre("JournalErreurActif")
    Chemin = Parametre("Chemin")
    CheminModsOfficiels = Parametre("CheminModsOfficiels")
    CheminMods = Parametre("CheminMods")
    If JournalErreurActif Then
        Kill App.Path & "\" & JournalErreurFichier
    End If
End Sub

Public Property Get Parametre(ByVal Nom As String) As String
    'MDIFrmMain.Caption = App.Title & " - chargement paramètres - " & Section & " - " & Nom
    lLength = GetPrivateProfileString(Section, Nom, "", sReturned, STRING_SIZE, App.Path & "\" & Fichier & ExtensionFichierINI)
    Parametre = Mid$(sReturned, 1, lLength)
    DernierChargement = Nom & "=" & Parametre
    If JournalErreurActif Then
        Open App.Path & "\" & JournalErreurFichier For Append As #1
        Print #1, "[" & Fichier & "] [" & Section & "] : " & DernierChargement
        Close #1
    End If
End Property

Public Property Let Parametre(ByVal Nom As String, ByVal Valeur As String)
    WritePrivateProfileString Section, Nom, Valeur, App.Path & "\" & Fichier & ExtensionFichierINI
End Property
