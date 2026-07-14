VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsJeuRessources"
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

'Position actuelle de la ressources.
Public PositionX As Long
Public PositionY As Long

Public Nom As String
Public Apparence As Long
Public TypeRessource As Long
Public Quantite As Double
Public MaxQuantite As Double

Public Property Get Visible() As Boolean
    Visible = Quantite > 0
End Property

Public Property Let Definir_Type(ByVal Valeur As Integer)
    Nom = Parametres.TerrainRessources_Nom(Valeur)
    Apparence = Valeur
    TypeRessource = Parametres.TerrainRessources_TypeRessource(Valeur)
    MaxQuantite = Parametres.TerrainRessources_Quantite(Valeur)
    Quantite = MaxQuantite
End Property

Public Sub Definir_Type_Aleatoire(ByVal Parametres As ClsJeuParam)
    Dim i As Integer
    With Parametres
    i = Rnd * .NombreTerrainRessource
    If i >= .NombreTerrainRessource Then i = 0
    Definir_Type = i
    End With
End Sub

