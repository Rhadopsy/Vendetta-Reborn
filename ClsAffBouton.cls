VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAffBouton"
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

Const FichierINI = "Boutons"
Const SectionINI = "Boutons"

Private CheminBoutonConstruire As String
Public NomBoutonDemolir As String
Private LargeurBoutonConstruire As Integer
Private HauteurBoutonConstruire As Integer

Private NomBoutonInfos As String
Private LargeurBoutonInfos As Integer
Private HauteurBoutonInfos As Integer

Private NomBoutonDeplacement As String
Private LargeurBoutonDeplacement As Integer
Private HauteurBoutonDeplacement As Integer

Private NomBoutonCommerce As String
Private LargeurBoutonCommerce As Integer
Private HauteurBoutonCommerce As Integer

Private NomBoutonAction As String
Private LargeurBoutonAction As Integer
Private HauteurBoutonAction As Integer

Private NomBoutonTransfert As String
Private LargeurBoutonTransfert As Integer
Private HauteurBoutonTransfert As Integer

Private NomBoutonMenu As String
Private LargeurBoutonMenu As Integer
Private HauteurBoutonMenu As Integer

Private NomBoutonDefilement As String
Private LargeurBoutonDefilement As Integer
Private HauteurBoutonDefilement As Integer
Private LargeurBoutonBarreDefilement As Integer
Private HauteurBoutonBarreDefilement As Integer
Private LargeurBoutonCarreDefilement As Integer
Private HauteurBoutonCarreDefilement As Integer

Private NomBoutonObjets As String
Private LargeurBoutonObjets As Integer
Private HauteurBoutonObjets As Integer

Private NomBoutonCompetences As String
Private LargeurBoutonCompetences As Integer
Private HauteurBoutonCompetences As Integer

Const Chemin = "Boutons\"

Private Sprite As ClsAffSprite
Private IndiceBouton As Integer
Public Largeur As Long
Public Hauteur As Long

Private Sub Class_Initialize()
    Set Sprite = New ClsAffSprite
    
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    CheminBoutonConstruire = FicIni.Parametre("CheminBoutonConstruire")
    NomBoutonDemolir = FicIni.Parametre("NomBoutonDemolir")
    LargeurBoutonConstruire = Val(FicIni.Parametre("LargeurBoutonConstruire"))
    HauteurBoutonConstruire = Val(FicIni.Parametre("HauteurBoutonConstruire"))
    
    NomBoutonInfos = FicIni.Parametre("NomBoutonInfos")
    LargeurBoutonInfos = Val(FicIni.Parametre("LargeurBoutonInfos"))
    HauteurBoutonInfos = Val(FicIni.Parametre("HauteurBoutonInfos"))
    
    NomBoutonDeplacement = FicIni.Parametre("NomBoutonDeplacement")
    LargeurBoutonDeplacement = Val(FicIni.Parametre("LargeurBoutonDeplacement"))
    HauteurBoutonDeplacement = Val(FicIni.Parametre("HauteurBoutonDeplacement"))
    
    NomBoutonCommerce = FicIni.Parametre("NomBoutonCommerce")
    LargeurBoutonCommerce = Val(FicIni.Parametre("LargeurBoutonCommerce"))
    HauteurBoutonCommerce = Val(FicIni.Parametre("HauteurBoutonCommerce"))
    
    NomBoutonAction = FicIni.Parametre("NomBoutonAction")
    LargeurBoutonAction = Val(FicIni.Parametre("LargeurBoutonAction"))
    HauteurBoutonAction = Val(FicIni.Parametre("HauteurBoutonAction"))
    
    NomBoutonTransfert = FicIni.Parametre("NomBoutonTransfert")
    LargeurBoutonTransfert = Val(FicIni.Parametre("LargeurBoutonTransfert"))
    HauteurBoutonTransfert = Val(FicIni.Parametre("HauteurBoutonTransfert"))
    
    NomBoutonMenu = FicIni.Parametre("NomBoutonMenu")
    LargeurBoutonMenu = Val(FicIni.Parametre("LargeurBoutonMenu"))
    HauteurBoutonMenu = Val(FicIni.Parametre("HauteurBoutonMenu"))
    
    NomBoutonDefilement = FicIni.Parametre("NomBoutonDefilement")
    LargeurBoutonDefilement = Val(FicIni.Parametre("LargeurBoutonDefilement"))
    HauteurBoutonDefilement = Val(FicIni.Parametre("HauteurBoutonDefilement"))
    LargeurBoutonBarreDefilement = Val(FicIni.Parametre("LargeurBoutonBarreDefilement"))
    HauteurBoutonBarreDefilement = Val(FicIni.Parametre("HauteurBoutonBarreDefilement"))
    LargeurBoutonCarreDefilement = Val(FicIni.Parametre("LargeurBoutonCarreDefilement"))
    HauteurBoutonCarreDefilement = Val(FicIni.Parametre("HauteurBoutonCarreDefilement"))
    
    NomBoutonObjets = FicIni.Parametre("NomBoutonObjets")
    LargeurBoutonObjets = Val(FicIni.Parametre("LargeurBoutonObjets"))
    HauteurBoutonObjets = Val(FicIni.Parametre("HauteurBoutonObjets"))
    
    NomBoutonCompetences = FicIni.Parametre("NomBoutonCompetences")
    LargeurBoutonCompetences = Val(FicIni.Parametre("LargeurBoutonCompetences"))
    HauteurBoutonCompetences = Val(FicIni.Parametre("HauteurBoutonCompetences"))
End Sub

Public Sub Charger(ByVal dd As DirectDraw7, _
                   ByVal TypeBouton As Integer, _
                   Optional ByVal NomFichier As String, _
                   Optional ByVal CheminFichier As String)
    Dim NomBouton As String
    IndiceBouton = TypeBouton
    Select Case IndiceBouton
        Case 0:
                If CheminFichier = "" Then
                    NomBouton = CheminBoutonConstruire & NomFichier
                Else
                    NomBouton = CheminFichier & NomFichier
                End If
                Largeur = LargeurBoutonConstruire
                Hauteur = HauteurBoutonConstruire
        Case 1: NomBouton = NomBoutonInfos
                Largeur = LargeurBoutonInfos
                Hauteur = HauteurBoutonInfos
        Case 2: NomBouton = NomBoutonAction
                Largeur = LargeurBoutonAction
                Hauteur = HauteurBoutonAction
        Case 3: NomBouton = NomBoutonDeplacement
                Largeur = LargeurBoutonDeplacement
                Hauteur = HauteurBoutonDeplacement
        Case 4: NomBouton = NomBoutonCommerce
                Largeur = LargeurBoutonCommerce
                Hauteur = HauteurBoutonCommerce
        Case 5: NomBouton = NomBoutonTransfert
                Largeur = LargeurBoutonTransfert
                Hauteur = HauteurBoutonTransfert
        Case 6: NomBouton = "Defilement"
                Largeur = LargeurBoutonDefilement
                Hauteur = HauteurBoutonDefilement
        Case 7: NomBouton = NomBoutonMenu
                Largeur = LargeurBoutonMenu
                Hauteur = HauteurBoutonMenu
        Case 8: NomBouton = NomBoutonDefilement
                Largeur = LargeurBoutonBarreDefilement
                Hauteur = HauteurBoutonBarreDefilement
        Case 9: NomBouton = NomBoutonDefilement
                Largeur = LargeurBoutonCarreDefilement
                Hauteur = HauteurBoutonCarreDefilement
        Case 9: NomBouton = NomBoutonDefilement
                Largeur = LargeurBoutonCarreDefilement
                Hauteur = HauteurBoutonCarreDefilement
        Case 10: NomBouton = NomBoutonObjets
                 Largeur = LargeurBoutonObjets
                 Hauteur = HauteurBoutonObjets
        Case 11: NomBouton = NomBoutonCompetences
                 Largeur = LargeurBoutonCompetences
                 Hauteur = HauteurBoutonCompetences
    End Select
    If CheminFichier = "" Then
        Sprite.Charger dd, Chemin & NomBouton
    Else
        Sprite.Charger dd, NomBouton, True
    End If
End Sub

Public Sub Afficher(ByVal X As Long, _
                    ByVal Y As Long, _
                    ByVal TypeBouton As Long, _
                    ByVal Etat As Integer, _
                    ByRef SurfBack As DirectDrawSurface7, _
                    Optional ByVal Transparent As Boolean, _
                    Optional ByVal Translucide As Boolean, _
                    Optional ByVal ValeurTransparence As Single)
    'On charge la bonne image en fonction des paramètres.
    Sprite.Selection Hauteur * Etat, _
                     Hauteur * (Etat + 1), _
                     Largeur * TypeBouton, _
                     Largeur * (TypeBouton + 1)
    'On affiche l'objet sur l'arrière plan.
    If ValeurTransparence > 0 Then
        Sprite.Afficher X, Y, SurfBack, Transparent, Translucide, , , ValeurTransparence
    Else
        Sprite.Afficher X, Y, SurfBack, Transparent, Translucide
    End If
End Sub
