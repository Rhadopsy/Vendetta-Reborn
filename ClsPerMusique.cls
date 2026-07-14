VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPerMusique"
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

Const FichierINI = "Musiques"
Const SectionINI = "Musiques"

Const NombreCategories = 9
Const Menu = 0
Const Debut = 1
Const Ville = 2
Const Combat = 3
Const Mort = 4
Const Victoire = 5
Const Defaite = 6
Const Boss = 7
Const Emperereur = 8

Private ExtensionMusiques1 As String
Private ExtensionMusiques2 As String
Private ExtensionMusiques3 As String

Private Jouer As Boolean
Private UtiliserDirectMusic As Boolean
Public Chemin As String

Private Type Musique
    Jouer As Boolean
    Chemin As String
    Fichiers As Collection
    Nombre As Integer
End Type
Private Musiques() As Musique

Private IntervalMusiques As Long  'Nombre de secondes entre chaque musiques.

Private TempsMusiqueJouee As Double
Private TypeMusiqueJouee As Integer
Private NomMusiqueJouee As String

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'Variables pour Direct Music.
Dim Perf As DirectMusicPerformance
Dim Seg As DirectMusicSegment
Dim Segstate As DirectMusicSegmentState
Dim Loader As DirectMusicLoader
Public GetStartTime As Long

Private Sub Class_Initialize()
    Dim Temp As String
    Dim i As Integer
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    ReDim Musiques(NombreCategories - 1) As Musique
    
    Jouer = FicIni.Parametre("Jouer")
    UtiliserDirectMusic = FicIni.Parametre("UtiliserDirectMusic")
    
    ExtensionMusiques1 = FicIni.Parametre("ExtensionMusiques1")
    ExtensionMusiques2 = FicIni.Parametre("ExtensionMusiques2")
    ExtensionMusiques3 = FicIni.Parametre("ExtensionMusiques3")
    
    Chemin = FicIni.Parametre("Chemin")
    
    Musiques(Menu).Chemin = FicIni.Parametre("CheminMenu")
    Musiques(Debut).Chemin = FicIni.Parametre("CheminDebut")
    Musiques(Ville).Chemin = FicIni.Parametre("CheminVille")
    Musiques(Combat).Chemin = FicIni.Parametre("CheminCombat")
    Musiques(Mort).Chemin = FicIni.Parametre("CheminMort")
    Musiques(Victoire).Chemin = FicIni.Parametre("CheminVictoire")
    Musiques(Defaite).Chemin = FicIni.Parametre("CheminDefaite")
    Musiques(Boss).Chemin = FicIni.Parametre("CheminBoss")
    Musiques(Emperereur).Chemin = FicIni.Parametre("CheminEmpereur")
    
    IntervalMusiques = FicIni.Parametre("IntervalMusiques") 'Nombre de secondes entre chaque musiques.
    
    'Recherche les noms des fichiers de musiques dans les différents répertoires.
    For i = 0 To NombreCategories - 1
        With Musiques(i)
        Set .Fichiers = New Collection
        Temp = Dir(App.Path & Chemin & .Chemin, vbNormal)
        While Temp <> ""
            If Right(LCase(Temp), Len(ExtensionMusiques1)) = ExtensionMusiques1 Or _
               Right(LCase(Temp), Len(ExtensionMusiques2)) = ExtensionMusiques2 Or _
               Right(LCase(Temp), Len(ExtensionMusiques3)) = ExtensionMusiques3 Then
                .Fichiers.Add Temp
            End If
            Temp = Dir()
        Wend
        .Nombre = .Fichiers.Count
        .Jouer = .Nombre > 0
        End With
    Next i
    
    'mciSendString "Close sonido", 0, 0, 0
    'TypeMusiqueJouee = -1
    'Set ds = dx.DirectSoundCreate("")
    'ds.SetCooperativeLevel Fenetre.hwnd, DSSCL_NORMAL
End Sub

'Public Sub Jouer_Musique(ByVal Categorie As Long, _
                         Optional ByVal MusiqueFacultative As Boolean)
'    Dim RetVal As Double
'
'    Open App.Path & Chemin & "Lecture.txt" For Input As #1 '
'    Line Input #1, cstr(Categorie)
'    Line Input #1, MusiqueFacultative
'    Close #1
'
'    RetVal = Shell(App.Path & "LecMusique.exe", 1)
'End Sub

Public Sub DM_Init(ByVal dx As DirectX7)
    On Error GoTo LocalErrors
    
    Set Loader = dx.DirectMusicLoaderCreate()

    Set Perf = dx.DirectMusicPerformanceCreate()
    Perf.Init Nothing, 0
    Perf.SetPort -1, 80
    Perf.SetMasterAutoDownload True
    
Exit Sub
LocalErrors:
    MsgBox Err.Number, Err.Description
End Sub

Private Sub DM_Charger(ByVal Dossier As String, _
                       ByVal Fichier As String)
    'Dim Minutes As Integer
    Dim Res As Integer
    'Dim length As Integer
    'Dim length2 As Integer

    On Error GoTo LocalErrors
   
    If Not Seg Is Nothing And Not Segstate Is Nothing Then ' There is a Segment and a SegmentState
        If Perf.IsPlaying(Seg, Segstate) Then  ' Segment currently playing, so exit
            DM_Stop
        End If
    End If
   
    Set Seg = Loader.LoadSegment(Dossier & Fichier)
    
    Loader.SetSearchDirectory Dossier
    Perf.SetMasterAutoDownload True
    
    Seg.SetStandardMidiFile
    Exit Sub
LocalErrors:
    Res = MsgBox("Le fichier " & Chr(34) & Fichier & Chr(34) & " ne peux être joué." & Chr(13) & "Voulez-vous désactiver la musique ?", vbYesNo + vbCritical, "Erreur")
    If Res = 6 Then Definir_JouerMusique = False
End Sub

Public Sub DM_Jouer(ByVal Categorie As Long, _
                    Optional ByVal MusiqueFacultative As Boolean)
    If Seg Is Nothing Then
        With Musiques(Categorie)
        'Sélectionne une musique au hasard.
        NomMusiqueJouee = .Fichiers(Int(Rnd * (.Nombre)) + 1)
        'Ouvre le fichier.
        DM_Charger App.Path & Chemin & .Chemin, NomMusiqueJouee
        'Joue le fichier.
        Seg.SetStartPoint 0
        Set Segstate = Perf.PlaySegment(Seg, 0, 0)
        GetStartTime = Segstate.GetStartTime()
        
        TempsMusiqueJouee = Timer
        TypeMusiqueJouee = Categorie
        End With
    Else
        With Musiques(Categorie)
        If .Jouer And _
           ((Perf.GetMusicTime() - GetStartTime >= Seg.GetLength) Or _
           Categorie <> TypeMusiqueJouee And Not MusiqueFacultative) Then
            
            'Arret du fichier précédent.
            DM_Stop
            NomMusiqueJouee = .Fichiers(Int(Rnd * (.Nombre)) + 1)
            'Ouvre le fichier.
            DM_Charger App.Path & Chemin & .Chemin, NomMusiqueJouee
            'Joue le fichier.
            Seg.SetStartPoint 0
            Set Segstate = Perf.PlaySegment(Seg, 0, 0)
            GetStartTime = Segstate.GetStartTime()
            
            TempsMusiqueJouee = Timer
            TypeMusiqueJouee = Categorie
        End If
        End With
    End If
End Sub

Public Sub DM_Stop()
    If Not Seg Is Nothing Then
        Perf.Stop Seg, Segstate, 0, 0
    End If
End Sub

Public Sub Jouer_Musique(ByVal Categorie As Long, _
                         Optional ByVal MusiqueFacultative As Boolean)
    If Jouer Then
        If TypeMusiqueJouee = Boss And Categorie = Combat Then 'Le combat, n'annule pas la musique d'un BOSS
            MusiqueFacultative = True
        End If
        If UtiliserDirectMusic Then
            DM_Jouer Categorie, MusiqueFacultative
        Else
            With Musiques(Categorie)
            If .Jouer And _
               (Timer > TempsMusiqueJouee + IntervalMusiques Or _
               Categorie <> TypeMusiqueJouee And Not MusiqueFacultative) Then
                
                'Arret du fichier précédent.
                mciSendString "Close sonido", 0, 0, 0
                NomMusiqueJouee = .Fichiers(Int(Rnd * (.Nombre)) + 1)
                'Ouvre le fichier.
                mciSendString "OPEN " & _
                              Chr(34) & App.Path & Chemin & .Chemin & _
                              NomMusiqueJouee & Chr(34) & _
                              " Alias Sonido", 0, 0, 0
                'Joue le fichier.
                mciSendString "Play sonido", 0, 0, 0
                
                TempsMusiqueJouee = Timer
                TypeMusiqueJouee = Categorie
            End If
            End With
        End If
    End If
End Sub

Public Sub StopSon()
    mciSendString "Stop sonido", 0, 0, 0
    mciSendString "Close sonido", 0, 0, 0
End Sub
Private Sub Class_Terminate()
    StopSon
End Sub

'Propriétés
Public Property Get Definir_JouerMusique() As Boolean
    Definir_JouerMusique = Jouer
End Property
Public Property Let Definir_JouerMusique(ByVal Valeur As Boolean)
    Jouer = Valeur
    If Not Jouer Then
        If UtiliserDirectMusic Then
            DM_Stop
        Else
            StopSon
        End If
    End If
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("Jouer") = -CDbl(Jouer)
    TypeMusiqueJouee = -1
End Property

'Propriétés
Public Property Get Definir_UtiliserDirectMusic() As Boolean
    Definir_UtiliserDirectMusic = UtiliserDirectMusic
End Property
Public Property Let Definir_UtiliserDirectMusic(ByVal Valeur As Boolean)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("UtiliserDirectMusic") = -CDbl(Valeur)
End Property
