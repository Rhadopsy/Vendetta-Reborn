Attribute VB_Name = "ModEffets"
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

Private TempEffet As ClsJeuEffet
Public Effets As Collection


Public Sub Effets_Init()
    Set Effets = New Collection
    While Effets.Count > 0
        Effets.Remove 1
    Wend
End Sub

Public Sub Activer_Effet(ByVal TypeEffet As Integer, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal LargeurCible As Long, _
                         ByVal HauteurCible As Long, _
                         ByVal Parametres As ClsJeuParam, _
                         Optional ByVal ForcerEffet As Boolean)
    If TypeEffet >= 0 Then
        Set TempEffet = Nothing
        Set TempEffet = New ClsJeuEffet
        TempEffet.Activer_Effet TypeEffet, X, Y, LargeurCible, HauteurCible, Parametres, ForcerEffet
        Effets.Add TempEffet
    End If
End Sub

