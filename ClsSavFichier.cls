VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSAVFichier"
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

Private NV_NOM_FICHIER As String
Private NV_REDEMARRAGE As Boolean
Private Parametres(5000, 2) As String
Private NOMBREPARAMETRES As Integer
Private Init As Boolean

Property Let NomFichier(ByVal UserFichier As String)
On Error Resume Next
Dim Chaine As String
Dim TAMPON As String
Dim Commentaire As Boolean
Dim Egal As Integer
Dim NbEgal As Integer
Dim i As Integer
Dim NoFichier As Integer
Dim SousClef As String
Dim SauvTampon As String

NV_NOM_FICHIER = UserFichier
NOMBREPARAMETRES = 0
NoFichier = FreeFile
Open NV_NOM_FICHIER For Input As NoFichier
If Err <> 0 Then
    Erreur
    Exit Property
End If
While Not EOF(NoFichier)
    Input #NoFichier, Chaine
    
    TAMPON = ""
    If Len(Chaine) > 2 Then SauvTampon = Mid(Chaine, 2, Len(Chaine) - 1)
    '1 - Elimination des espaces
    For i = 1 To Len(Chaine)
        If Mid(Chaine, i, 1) <> " " Then TAMPON = TAMPON & Mid(Chaine, i, 1)
    Next i
    
    '2 - Conversion en majuscule
    'Chaine = UCase(TAMPON)
    
    '3 - Recherche du caractère "[" qui => commentaire et du caractère "="
    Commentaire = False
    NbEgal = 0
    For i = 1 To Len(Chaine)
        TAMPON = Mid(Chaine, i, 1)
        If TAMPON = "[" Or TAMPON = "*" Then Commentaire = True
        If TAMPON = "*" Then SousClef = SauvTampon
        If TAMPON = "=" Then
            Egal = i
            NbEgal = NbEgal + 1
        End If
    Next i
                
    If NbEgal <> 1 And Not Commentaire Then Commentaire = True
        
    '4 - Ajout à la base
    If Not Commentaire Then
        NOMBREPARAMETRES = NOMBREPARAMETRES + 1
        Parametres(NOMBREPARAMETRES, 0) = Mid(Chaine, 1, Egal - 1)
        Parametres(NOMBREPARAMETRES, 1) = Mid(Chaine, Egal + 1)
        Parametres(NOMBREPARAMETRES, 2) = SousClef
    End If
Wend
Init = True
Close NoFichier
NV_REDEMARRAGE = False
End Property

Public Property Get Valeur(ByVal NomArg As String) As String
Dim i As Integer
If Not Init Then
    Exit Property
End If
Valeur = Empty
For i = 1 To NOMBREPARAMETRES
    'If Parametres(i, 0) = UCase(Trim(NomArg)) Then Valeur = Parametres(i, 1)
    If Parametres(i, 0) = Trim(NomArg) Then Valeur = Parametres(i, 1)
Next i
'If Valeur = Empty Then Erreur
End Property

Public Property Let Valeur(ByVal NomArg As String, ByVal Valeur As String)
On Error Resume Next
Dim ListeLigne(500) As String
Dim NombreLigne As Integer
Dim Trouve As Boolean
Dim i As Integer
Dim Egal As Integer
Dim NoFichier As Byte

NoFichier = FreeFile
'1 - Lecture et Stockage des données du fichier
Open NV_NOM_FICHIER For Input As NoFichier
If Err <> 0 Then
    Erreur
    Exit Property
End If

While Not EOF(NoFichier)
    Input #NoFichier, ListeLigne(NombreLigne)
    NombreLigne = NombreLigne + 1
Wend
Close (NoFichier)
'2 - Remplacement de la valeur correspondant à la clef
i = 0
While Not Trouve
    If InStr(1, ListeLigne(i), "[", vbTextCompare) = 0 Then
        Egal = InStr(1, ListeLigne(i), "=", vbTextCompare)
        If (Egal <> 0) And (InStr(1, ListeLigne(i), NomArg, 1) <> 0) Then
            ListeLigne(i) = Left(ListeLigne(i), Egal - 1) & "=" & Trim(Valeur)
            Trouve = True
        End If
    End If
    i = i + 1
Wend
    
'3 - Réecriture dans le fichier
Open NV_NOM_FICHIER For Output As #NoFichier
For i = 0 To NombreLigne - 1
    Print #NoFichier, ListeLigne(i)
Next i
Close (NoFichier)

'4 - Mise à jour des paramètres en mémoire
For i = 1 To NOMBREPARAMETRES
    'If Parametres(i, 0) = UCase(Trim(NomArg)) Then Parametres(i, 1) = Valeur
    If Parametres(i, 0) = Trim(NomArg) Then Parametres(i, 1) = Valeur
Next i
NV_REDEMARRAGE = True
End Property

Public Property Get NomParametre(ByVal NoParametre As Integer) As String
NomParametre = Parametres(NoParametre, 0)
End Property

Public Property Get NombreParametre() As Integer
NombreParametre = NOMBREPARAMETRES
End Property

Public Property Get Redemarrage() As Boolean
Redemarrage = NV_REDEMARRAGE
End Property

Private Sub Erreur()
MsgBox "Fichier " & NV_NOM_FICHIER & " manquant ou détérioré", vbCritical, "Erreur"
End
End Sub


