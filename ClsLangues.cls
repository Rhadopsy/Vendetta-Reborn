VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLangues"
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

Const FichierINI = "Langues"
Const SectionINI = "Langues"
Const SectionsINI = "Langue_"

Private Type Langue
    Nom As String
    Suffixe As String
    Dossier As String
End Type

Public Nombre As Integer
Private Langues() As Langue
Private LangueChoisie As Integer

Private Sub Class_Initialize()
    Dim i As Integer

    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    Nombre = Val(FicIni.Parametre("NombreLangues"))
    ReDim Langues(Nombre - 1)
    LangueChoisie = Val(FicIni.Parametre("LangueChoisie")) - 1
    For i = 0 To UBound(Langues())
        FicIni.Section = SectionsINI & i + 1
        Langues(i).Nom = FicIni.Parametre("Nom")
        Langues(i).Dossier = FicIni.Parametre("Dossier")
        Langues(i).Suffixe = FicIni.Parametre("Suffixe")
    Next i
End Sub

Public Property Get Definir_Nom(ByVal Indice As Integer) As String
    Definir_Nom = Langues(Indice).Nom
End Property
Public Property Get Nom() As String
    Nom = Langues(LangueChoisie).Nom
End Property
Public Property Get Dossier() As String
    Dossier = Langues(LangueChoisie).Dossier
End Property
Public Property Get Suffixe() As String
    Suffixe = Langues(LangueChoisie).Suffixe
End Property

Public Property Get Definir_Langue() As Integer
    Definir_Langue = LangueChoisie
End Property
Public Property Let Definir_Langue(ByVal Valeur As Integer)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    LangueChoisie = Valeur
    FicIni.Parametre("LangueChoisie") = Valeur + 1
End Property
