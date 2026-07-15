VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPerSon"
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

Const FichierINI = "Sons"
Const SectionINI = "Sons"

Private Jouer As Boolean
Public Chemin As String

Private ds As IDirectSound
Private SoundDesc As DSBUFFERDESC
Private WavFormat As WAVEFORMATEX

Private Type Son
    Actif As Boolean
    DsSound As IDirectSoundBuffer
End Type

Private Sons() As Son

Private Sub Class_Initialize()
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    Jouer = FicIni.Parametre("Jouer")
    Chemin = FicIni.Parametre("Chemin")
End Sub

Public Sub Init(ByVal Fenetre As Form, ByVal dx As DirectX7, ByVal Parametres As ClsJeuParam)
    On Error GoTo Erreur
    Dim i As Integer, NbEffets As Integer
    NbEffets = Parametres.NombreEffets
    ReDim Sons(NbEffets - 1) As Son
    If Definir_Jouer Then
        Set ds = dx.DirectSoundCreate("")
        ds.SetCooperativeLevel Fenetre.hwnd, DSSCL_NORMAL
        For i = 0 To NbEffets - 1
            'Set Sons(i).DsSound = New IDirectSoundBuffer
            If Left(Parametres.Effet_Son(i), 1) = "." Then
                'Chemin absolu (fichier mod)
                Sons(i).Actif = Charger_Son(i, App.Path & Right(Parametres.Effet_Son(i), Len(Parametres.Effet_Son(i)) - 1))
            Else
                Sons(i).Actif = Charger_Son(i, App.Path & Chemin & Parametres.Effet_Son(i))
            End If
        Next i
    End If
    Exit Sub
Erreur:
    MsgBox "Impossible d'initialiser les bruitages" & Chr(10) & Chr(13) & "Les bruitages ont été désactivés.", vbCritical, "Erreur"
    Definir_Jouer = False
End Sub

Private Function Charger_Son(ByVal i As Integer, _
                             ByVal Fichier As String) As Boolean
    On Error GoTo Erreur
    Set Sons(i).DsSound = ds.CreateSoundBufferFromFile(Fichier, SoundDesc, WavFormat)
    Charger_Son = True
    Exit Function
Erreur:
    Charger_Son = False
End Function

Public Sub Lire(ByVal i As Integer)
    If Sons(i).Actif Then
        Sons(i).DsSound.Play DSBPLAY_DEFAULT
    End If
End Sub

'Propriétés
Public Property Get Definir_Jouer() As Boolean
    Definir_Jouer = Jouer
End Property
Public Property Let Definir_Jouer(ByVal Valeur As Boolean)
    Jouer = Valeur
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("Jouer") = -CDbl(Jouer)
End Property

