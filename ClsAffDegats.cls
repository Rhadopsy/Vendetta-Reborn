VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAffDegats"
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

Const FichierINI = "Combat"
Const SectionINI = "Degats"

Private AfficherDegats As Boolean

Private TailleNormale As Integer
Private TailleCritique As Integer
Private CouleurDegat As Long
Private CouleurGueri As Long
Private CouleurVente As Long
Private DureeMax As Integer
Private DecalageX As Integer
Private DecalageY As Integer
Private DeplacementX As Integer
Private DeplacementY As Integer

Public Degats As Collection

Public FontDegats As IFont
Public FontCritique As IFont
Public FontNormale As IFont

Private Sub Class_Initialize()
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    AfficherDegats = CBool(FicIni.Parametre("AfficherDegats"))
    TailleNormale = Val(FicIni.Parametre("TailleNormale"))
    TailleCritique = Val(FicIni.Parametre("TailleCritique"))
    CouleurDegat = Val(FicIni.Parametre("CouleurDegat"))
    CouleurGueri = Val(FicIni.Parametre("CouleurGueri"))
    CouleurVente = Val(FicIni.Parametre("CouleurVente"))
    DureeMax = Val(FicIni.Parametre("DureeMax"))
    DecalageX = Val(FicIni.Parametre("DecalageX"))
    DecalageY = Val(FicIni.Parametre("DecalageY"))
    DeplacementX = Val(FicIni.Parametre("DeplacementX"))
    DeplacementY = Val(FicIni.Parametre("DeplacementY"))

    Set Degats = New Collection
    Set FontDegats = FrmChargement.LblFontDegats.Font
    Set FontCritique = FrmChargement.LblFontCritique.Font
    Set FontNormale = FrmChargement.LblFontNormale.Font
    FontDegats.Size = TailleNormale
    FontCritique.Size = TailleCritique
    FontNormale.Size = 10
End Sub

Public Sub Degats_Init()
    Collection_Vider Degats
End Sub

Public Sub Ajouter_Degat(ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal Score As String, _
                         Optional ByVal Critique As Boolean = False, _
                         Optional ByVal TypeDegat As Integer = 0)
    'TypeDegat 0: perte de vie (rouge) 1: gain de vie Bleu.
    Dim Temp As ClsJeuDegat
    Set Temp = New ClsJeuDegat
    Score = CLng(Score)
    If Score <> 0 Then
        Select Case TypeDegat
        Case 0: Temp.Couleur = CouleurDegat
        Case 1: Temp.Couleur = CouleurGueri
        Case 2: Temp.Couleur = CouleurVente ': Score = "+" & Score
        End Select
        Temp.Score = Score
        Temp.X = X + DecalageX - Len(Score) * IIf(Critique, TailleCritique, TailleNormale) / 2
        Temp.Y = Y + DecalageY
        Temp.Critique = Critique
        Degats.Add Temp
    End If
End Sub

Public Sub Ajouter_Texte(ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal Texte As String)
    Dim Temp As ClsJeuDegat
    Set Temp = New ClsJeuDegat
    Temp.X = X + DecalageX + 4 - Len(Texte) * TailleNormale / 2
    Temp.Y = Y + DecalageY - 10
    Temp.Score = Texte
    Temp.Couleur = CouleurDegat
    Degats.Add Temp
End Sub

Public Sub Afficher_Degats(ByVal DecalageX As Long, _
                           ByVal DecalageY As Long, _
                           ByVal EcranWidth As Long, _
                           ByVal EcranHeight As Long, _
                           ByRef SurfBack As IDirectDrawSurface7)
    Dim i As Long
    Dim X As Integer, Y As Integer
    If AfficherDegats Then
        For i = 1 To Degats.Count
            If Degats(i).Critique Then
                SurfBack.SetFont FontCritique
            Else
                SurfBack.SetFont FontDegats
            End If
            SurfBack.SetForeColor Degats(i).Couleur
            X = Degats(i).X - DecalageX
            Y = Degats(i).Y - DecalageY
            If X > ScreenPositionX And _
               X < EcranWidth + ScreenPositionX And _
               Y > ScreenPositionY And _
               Y < EcranHeight + ScreenPositionY Then
                SurfBack.DrawText X, _
                                  Y, _
                                  Degats(i).Score, False
            End If
        Next i
        SurfBack.SetFont FontNormale
    End If
End Sub

Public Sub Actualiser_Degats()
    Dim i As Long
    i = 1
    While i <= Degats.Count
        If Degats(i).Duree > DureeMax Then
            Degats.Remove i
        Else
            Degats(i).Duree = Degats(i).Duree + 1
            Degats(i).X = Degats(i).X + DeplacementX
            Degats(i).Y = Degats(i).Y + DeplacementY
            i = i + 1
        End If
    Wend
End Sub

'Propriétés
Public Property Get Definir_AfficherDegats() As Boolean
    Definir_AfficherDegats = AfficherDegats
End Property
Public Property Let Definir_AfficherDegats(ByVal Valeur As Boolean)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    AfficherDegats = Valeur
    FicIni.Parametre("AfficherInfoBulle") = -CBool(AfficherDegats)
End Property

