VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsIntBouton"
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
Option Compare Binary

Const FichierINI = "Boutons"
Const SectionINI = "Boutons"

Private AfficherInfoBulle As Boolean
Private InfoBulleDecalageX As Integer
Private InfoBulleDecalageY As Integer
Private CouleurInfoBulle As Long
Private CouleurRectangleFond As Long
Private BordureRectangleFond As Integer

'Caractéristiques du bouton.
Public PositionX As Long
Public PositionY As Long
Public Largeur As Integer
Public Hauteur As Integer
Public Enabled As Boolean
Public Visible As Boolean
Public MouseOver As Boolean

Public InfoBulle As String

Private Sub Class_Initialize()
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    AfficherInfoBulle = Val(FicIni.Parametre("AfficherInfoBulle"))
    InfoBulleDecalageX = Val(FicIni.Parametre("InfoBulleDecalageX"))
    InfoBulleDecalageY = Val(FicIni.Parametre("InfoBulleDecalageY"))
    CouleurInfoBulle = Val(FicIni.Parametre("CouleurInfoBulle"))
    CouleurRectangleFond = Val(FicIni.Parametre("CouleurRectangleFond"))
    BordureRectangleFond = Val(FicIni.Parametre("BordureRectangleFond"))
End Sub

Public Sub Afficher_Info_Bulle(ByVal X As Long, _
                               ByVal Y As Long, _
                               ByVal Message As ClsIntMessage, _
                               ByRef SurfBack As IDirectDrawSurface7, _
                               Optional ByVal ForcerAffichage As Boolean)
    Dim LargeurFond As Integer
    Dim HauteurFond As Integer
    Dim TempMessage As String
    If AfficherInfoBulle And InfoBulle <> "" Or ForcerAffichage Then
        SurfBack.SetForeColor 0
        X = X + InfoBulleDecalageX
        Y = Y + InfoBulleDecalageY
        TempMessage = InfoBulle
        Do
            If Len(Left$(TempMessage, InStr(TempMessage, Message.Caractere_RetourLigne))) > LargeurFond Then
                LargeurFond = Len(Left(TempMessage, InStr(TempMessage, Message.Caractere_RetourLigne)))
            End If
            If Len(Left$(TempMessage, InStr(TempMessage, Chr$(13)))) > LargeurFond Then
                LargeurFond = Len(Left$(TempMessage, InStr(TempMessage, Chr$(13))))
            End If
            HauteurFond = HauteurFond + 1
            If InStr(TempMessage, Message.Caractere_RetourLigne) Then
                TempMessage = Right$(TempMessage, Len(TempMessage) - InStr(TempMessage, Message.Caractere_RetourLigne) - 0)
            ElseIf InStr(TempMessage, Chr$(13)) Then
                TempMessage = Right$(TempMessage, Len(TempMessage) - InStr(TempMessage, Chr$(13)) - 0)
            Else
                If Len(TempMessage) > LargeurFond Then
                    LargeurFond = Len(TempMessage)
                End If
                TempMessage = ""
            End If
        'Loop Until InStr(TempMessage, Message.Caractere_RetourLigne) = 0 And InStr(TempMessage, Chr(13)) = 0
        Loop Until Len(TempMessage) = 0
        LargeurFond = LargeurFond * Message.Largeur_Caractere
        HauteurFond = HauteurFond * Message.Hauteur_Caractere
        
        'Avec ce code, les infos bulles ne sortent pas de l'écran.
        If X < ScreenPositionX Then
            X = ScreenPositionX
        ElseIf X + LargeurFond + BordureRectangleFond > ScreenWidth Then
            X = ScreenWidth - LargeurFond - BordureRectangleFond
        End If
        If Y < ScreenPositionY Then
            Y = ScreenPositionY
        ElseIf Y + HauteurFond + BordureRectangleFond > ScreenHeight Then
            Y = ScreenHeight - HauteurFond - BordureRectangleFond
        End If
        
        SurfBack.SetFillColor CouleurRectangleFond
        SurfBack.DrawBox X - BordureRectangleFond, _
                         Y - BordureRectangleFond, _
                         X + LargeurFond + BordureRectangleFond, _
                         Y + HauteurFond + BordureRectangleFond
        
        SurfBack.SetForeColor CouleurInfoBulle
        Message.Afficher X, _
                         Y, _
                         InfoBulle, SurfBack
    End If
End Sub

'Propriétés
Public Property Get Definir_AfficherInfoBulle() As Boolean
    Definir_AfficherInfoBulle = AfficherInfoBulle
End Property
Public Property Let Definir_AfficherInfoBulle(ByVal Valeur As Boolean)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    AfficherInfoBulle = Valeur
    FicIni.Parametre("AfficherInfoBulle") = -CDbl(AfficherInfoBulle)
End Property
