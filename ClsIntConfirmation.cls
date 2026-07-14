VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsIntConfirmation"
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
'Fenêtre de confirmation.
Option Explicit

Const ConfirmationQuitter = 0
Const ConfirmationAttaquer = 1

Const LBoutonsChoix = 1
Const FichierImageFenetre = "Fenêtre-Confirmation"
Const LargeurFenetre = 320
Const HauteurFenetre = 160
Const Bouton1PositionX = 64
Const Bouton1PositionY = 100
Const Bouton2PositionX = 320 - 32 - 64
Const Bouton2PositionY = 100
Const BoutonsLargeur = 32
Const BoutonsHauteur = 32
Const EtiquetteCouleur = 0 '&HFFFFFF
Const EtiquettePositionX = 16
Const EtiquettePositionY = 16

Public Affiche As Boolean
Public Message As String
Public Categorie As Integer
Public PositionX As Long
Public PositionY As Long
Public Visible As Boolean

Private AffFenetre As ClsAffFenetre

Private AffBouton As ClsAffBouton
Private BoutonsChoix() As ClsIntBouton

Private MessagePositionX As Long
Private MessagePositionY As Long

'Variables temporaires.
Private i As Integer

Private Sub Class_Initialize()
    Set AffFenetre = New ClsAffFenetre
    Set AffBouton = New ClsAffBouton
    ReDim BoutonsChoix(LBoutonsChoix)
    For i = 0 To LBoutonsChoix
        Set BoutonsChoix(i) = New ClsIntBouton
    Next i
    AffFenetre.Definir_Nombre_Sprites = 1
    AffFenetre.Definir_Image(0) = FichierImageFenetre
End Sub

Public Sub Init(ByRef dd As DirectDraw7)
    'Charge les sprites.
    AffFenetre.Charger dd
    AffBouton.Charger dd, 7
End Sub

Public Sub Afficher_Fenetre(ByVal Indice As Integer, _
                            ByVal ValMessage As String, _
                            ByVal DecalageX As Long, ByVal DecalageY As Long, _
                            ByVal LargeurEcran As Long, ByVal HauteurEcran As Long)
    Categorie = Indice
    Message = ValMessage
    Affiche = True
    PositionX = DecalageX + (LargeurEcran - LargeurFenetre) / 2
    PositionY = DecalageY + (HauteurEcran - HauteurFenetre) / 2
    BoutonsChoix(0).PositionX = PositionX + Bouton1PositionX
    BoutonsChoix(0).PositionY = PositionY + Bouton1PositionY
    BoutonsChoix(1).PositionX = PositionX + Bouton2PositionX
    BoutonsChoix(1).PositionY = PositionY + Bouton2PositionY
    MessagePositionX = PositionX + EtiquettePositionX
    MessagePositionY = PositionY + EtiquettePositionY
End Sub

Public Sub Afficher(ByRef SurfBack As DirectDrawSurface7)
    'Affiche la fenêtre.
    AffFenetre.Afficher PositionX, PositionY, SurfBack
    'Affiche les boutons.
    For i = 0 To LBoutonsChoix
        AffBouton.Afficher BoutonsChoix(i).PositionX, _
                           BoutonsChoix(i).PositionY, _
                           i + 2, _
                           -CDbl(BoutonsChoix(i).MouseOver), _
                           SurfBack, True
    Next i
    'Affiche le message.
    AffMessage.Afficher MessagePositionX, MessagePositionY, Message, SurfBack, EtiquetteCouleur, True
End Sub

Public Sub Valider(ByRef Perso As ClsJeuPerso)
    Select Case Categorie
    Case 0: Quitter = True
    Case 1: Cmd_Perso_AttaquerAllies Perso.Numero
    End Select
    Affiche = False
End Sub
Public Sub Annuler(ByRef Interface As ClsInterface, ByVal NoJoueur As Integer)
    If Interface.Definir_BoutonInfoEnabled(7) Then
        Interface.Actionner_Boutons_Infos 7, NoJoueur
    End If
    Affiche = False
End Sub

Public Function Interieur_Boutons(ByVal Numero As Integer, _
                                  ByVal Perso As ClsJeuPerso, _
                                  ByRef X As Single, _
                                  ByRef Y As Single) As Boolean
    With BoutonsChoix(Numero)
    If Interieur_Bouton(X, Y, _
                        .PositionX, .PositionY, _
                        BoutonsLargeur, BoutonsHauteur) Then
        .MouseOver = True
        Interieur_Boutons = True
        'Messages(0).Ajouter_Message IIf(Numero = 0, "Oui", "Non")
    Else
        .MouseOver = False
    End If
    End With
End Function

Private Function Interieur_Bouton(ByVal X1 As Long, _
                                  ByVal Y1 As Long, _
                                  ByVal X2 As Long, _
                                  ByVal Y2 As Long, _
                                  ByVal Largeur As Long, _
                                  ByVal Hauteur As Long) As Boolean
    Interieur_Bouton = X1 >= X2 And _
                       X1 <= X2 + Largeur And _
                       Y1 >= Y2 And _
                       Y1 <= Y2 + Hauteur
End Function

Public Property Get Definir_LBoutonsChoix() As Integer
    Definir_LBoutonsChoix = LBoutonsChoix
End Property

