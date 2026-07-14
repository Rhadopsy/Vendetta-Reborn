VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsIntTouches"
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

Const FichierINI = "Touches"
Const SectionINI = "Touches"
Const FichierINI2 = "TouchesDialogue"
Const SectionINI2 = "TouchesDialogue"

'Correspondance des touches du clavier.
Public ToucheQuitter As Integer
Public ToucheZoomNormal As Integer
Public ToucheZoomPlus As Integer
Public ToucheZoomMoins As Integer
Public ToucheDialogue As Integer
Public ToucheFenetreBatiment As Integer
Public ToucheFenetrePersonnage As Integer
Public ToucheFenetreRessources As Integer
Public ToucheFenetreCompetences As Integer
Public ToucheFenetreInventaire As Integer
Public ToucheFenetreBestiaire As Integer
Public ToucheFenetreCarte As Integer
Public ToucheFenetreObjectifs As Integer
Public ToucheFenetresFermer As Integer
Public ToucheAction As Integer
Public ToucheRetourChateau As Integer
Public ToucheRetourMaison As Integer
Public ToucheTrouverRessource1 As Integer
Public ToucheTrouverRessource2 As Integer
Public ToucheTrouverRessource3 As Integer
Public ToucheTrouverRessource4 As Integer
Public ToucheTrouverRessource5 As Integer
Public ToucheTrouverRessource6 As Integer
Public ToucheTrouverRessource7 As Integer
Public ToucheTrouverRessource8 As Integer
Public ToucheTrouverAttaquerPerso As Integer
Public ToucheTrouverAttaquerMaison As Integer
Public ToucheTrouverAttaquerChateau As Integer
Public ToucheTrouverTresor As Integer
Public ToucheModeAuto As Integer
Public ToucheAide As Integer
Public ToucheArret As Integer
'Public ToucheChargez As Integer
Public ToucheDefinirAttitude1 As Integer
Public ToucheDefinirAttitude2 As Integer
Public ToucheDefinirAttitude3 As Integer
Public ToucheDefinirAttitude4 As Integer
Public ToucheDefinirAttitude5 As Integer
Public ToucheDefinirFormation As Integer

Public ToucheArgentDeposer As Integer
Public ToucheArgentRecuperer As Integer

Public ToucheDemolirMaison As Integer

Public ToucheAfficherConseil As Integer
Public ToucheMenu As Integer
Public ToucheCaptureEcran As Integer
Public ToucheSauvegarderMonde As Integer
Public ToucheMusique As Integer

Public ToucheInterfacePleinEcran As Integer
Public ToucheChangerPersonnage As Integer
Public ToucheDirectionHaut As Integer
Public ToucheDirectionBas As Integer
Public ToucheDirectionGauche As Integer
Public ToucheDirectionDroite As Integer
Public ToucheObjets As Integer
Public NombreTouchesObjets As Integer

'Touches de raccourci des armées.
Public ToucheSoldatsToutes As Integer
Public ToucheSoldatsMemeType As Integer
Public ToucheSoldatsEnArme As Integer
Public ToucheSoldatsEcran As Integer
Public ToucheSoldatsReposes As Integer
Public ToucheSoldatsEnArret As Integer
Public ToucheSoldatsRapides As Integer
Public ToucheSoldatsNonAffectees As Integer
Public ToucheSoldatsTerrestres As Integer
Public ToucheSoldatsVolants As Integer
Public ToucheSoldatsDefensifs As Integer
Public ToucheSoldatsOffensifs As Integer
Public ToucheSoldatsInverser As Integer
Public ToucheSoldatsSupprimer As Integer

Public ToucheDialogueValider1 As Integer
Public ToucheDialogueValider2 As Integer
Public ToucheDialogueAnnuler1 As Integer
Public ToucheDialogueAnnuler2 As Integer
Public ToucheDialogueEffacerPrecedent As Integer
Public ToucheDialogueEffacerSuivant As Integer
Public ToucheDialogueMessagePrecedent As Integer
Public ToucheDialogueMessageSuivant As Integer
Public ToucheDialogueCurseurDebut As Integer
Public ToucheDialogueCurseurPrecedent As Integer
Public ToucheDialogueCurseurSuivant As Integer
Public ToucheDialogueCurseurFin As Integer
Private ToucheDialogueLettre() As String
Private ToucheDialogueLettreMaj() As String

'Variables temporaires.
Private TempSoldatsSelectionnes As Collection

Private Sub Class_Initialize()
    Dim i As Integer
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    ToucheQuitter = Val(FicIni.Parametre("ToucheQuitter"))
    ToucheZoomNormal = Val(FicIni.Parametre("ToucheZoomNormal"))
    ToucheZoomPlus = Val(FicIni.Parametre("ToucheZoomPlus"))
    ToucheZoomMoins = Val(FicIni.Parametre("ToucheZoomMoins"))
    ToucheDialogue = Val(FicIni.Parametre("ToucheDialogue"))
    ToucheFenetreBatiment = Val(FicIni.Parametre("ToucheFenetreBatiment"))
    ToucheFenetrePersonnage = Val(FicIni.Parametre("ToucheFenetrePersonnage"))
    ToucheFenetreRessources = Val(FicIni.Parametre("ToucheFenetreRessources"))
    ToucheFenetreCompetences = Val(FicIni.Parametre("ToucheFenetreCompetences"))
    ToucheFenetreInventaire = Val(FicIni.Parametre("ToucheFenetreInventaire"))
    ToucheFenetreBestiaire = Val(FicIni.Parametre("ToucheFenetreBestiaire"))
    ToucheFenetreCarte = Val(FicIni.Parametre("ToucheFenetreCarte"))
    ToucheFenetreObjectifs = Val(FicIni.Parametre("ToucheFenetreObjectifs"))
    ToucheFenetresFermer = Val(FicIni.Parametre("ToucheFenetresFermer"))
    ToucheAction = Val(FicIni.Parametre("ToucheAction"))
    ToucheRetourChateau = Val(FicIni.Parametre("ToucheRetourChateau"))
    ToucheRetourMaison = Val(FicIni.Parametre("ToucheRetourMaison"))
    ToucheTrouverRessource1 = Val(FicIni.Parametre("ToucheTrouverRessource1"))
    ToucheTrouverRessource2 = Val(FicIni.Parametre("ToucheTrouverRessource2"))
    ToucheTrouverRessource3 = Val(FicIni.Parametre("ToucheTrouverRessource3"))
    ToucheTrouverRessource4 = Val(FicIni.Parametre("ToucheTrouverRessource4"))
    ToucheTrouverRessource5 = Val(FicIni.Parametre("ToucheTrouverRessource5"))
    ToucheTrouverRessource6 = Val(FicIni.Parametre("ToucheTrouverRessource6"))
    ToucheTrouverRessource7 = Val(FicIni.Parametre("ToucheTrouverRessource7"))
    ToucheTrouverRessource8 = Val(FicIni.Parametre("ToucheTrouverRessource8"))
    ToucheTrouverAttaquerPerso = Val(FicIni.Parametre("ToucheTrouverAttaquerPerso"))
    ToucheTrouverAttaquerMaison = Val(FicIni.Parametre("ToucheTrouverAttaquerMaison"))
    ToucheTrouverAttaquerChateau = Val(FicIni.Parametre("ToucheTrouverAttaquerChateau"))
    ToucheTrouverTresor = Val(FicIni.Parametre("ToucheTrouverTresor"))
    ToucheModeAuto = Val(FicIni.Parametre("ToucheModeAuto"))
    ToucheAide = Val(FicIni.Parametre("ToucheAide"))
    ToucheArret = Val(FicIni.Parametre("ToucheArret"))
    'ToucheChargez = Val(FicIni.Parametre("ToucheChargez"))
    ToucheDefinirAttitude1 = Val(FicIni.Parametre("ToucheDefinirAttitude1"))
    ToucheDefinirAttitude2 = Val(FicIni.Parametre("ToucheDefinirAttitude2"))
    ToucheDefinirAttitude3 = Val(FicIni.Parametre("ToucheDefinirAttitude3"))
    ToucheDefinirAttitude4 = Val(FicIni.Parametre("ToucheDefinirAttitude4"))
    ToucheDefinirAttitude5 = Val(FicIni.Parametre("ToucheDefinirAttitude5"))
    ToucheDefinirFormation = Val(FicIni.Parametre("ToucheDefinirFormation"))
    
    ToucheArgentDeposer = Val(FicIni.Parametre("ToucheArgentDeposer"))
    ToucheArgentRecuperer = Val(FicIni.Parametre("ToucheArgentRecuperer"))
    
    ToucheDemolirMaison = Val(FicIni.Parametre("ToucheDemolirMaison"))
    
    ToucheAfficherConseil = Val(FicIni.Parametre("ToucheAfficherConseil"))
    ToucheMenu = Val(FicIni.Parametre("ToucheMenu"))
    ToucheCaptureEcran = Val(FicIni.Parametre("ToucheCaptureEcran"))
    ToucheSauvegarderMonde = Val(FicIni.Parametre("ToucheSauvegarderMonde"))
    ToucheMusique = Val(FicIni.Parametre("ToucheMusique"))
    
    ToucheInterfacePleinEcran = Val(FicIni.Parametre("ToucheInterfacePleinEcran"))
    ToucheChangerPersonnage = Val(FicIni.Parametre("ToucheChangerPersonnage"))
    ToucheDirectionHaut = Val(FicIni.Parametre("ToucheDirectionHaut"))
    ToucheDirectionBas = Val(FicIni.Parametre("ToucheDirectionBas"))
    ToucheDirectionGauche = Val(FicIni.Parametre("ToucheDirectionGauche"))
    ToucheDirectionDroite = Val(FicIni.Parametre("ToucheDirectionDroite"))
    ToucheObjets = Val(FicIni.Parametre("ToucheObjets"))
    NombreTouchesObjets = Val(FicIni.Parametre("NombreTouchesObjets"))
    
    ToucheSoldatsToutes = Val(FicIni.Parametre("ToucheSoldatsToutes"))
    ToucheSoldatsMemeType = Val(FicIni.Parametre("ToucheSoldatsMemeType"))
    ToucheSoldatsEnArme = Val(FicIni.Parametre("ToucheSoldatsEnArme"))
    ToucheSoldatsEcran = Val(FicIni.Parametre("ToucheSoldatsEcran"))
    ToucheSoldatsReposes = Val(FicIni.Parametre("ToucheSoldatsReposes"))
    ToucheSoldatsEnArret = Val(FicIni.Parametre("ToucheSoldatsEnArret"))
    ToucheSoldatsRapides = Val(FicIni.Parametre("ToucheSoldatsRapides"))
    ToucheSoldatsNonAffectees = Val(FicIni.Parametre("ToucheSoldatsNonAffectees"))
    ToucheSoldatsTerrestres = Val(FicIni.Parametre("ToucheSoldatsTerrestres"))
    ToucheSoldatsVolants = Val(FicIni.Parametre("ToucheSoldatsVolants"))
    ToucheSoldatsDefensifs = Val(FicIni.Parametre("ToucheSoldatsDefensifs"))
    ToucheSoldatsOffensifs = Val(FicIni.Parametre("ToucheSoldatsOffensifs"))
    ToucheSoldatsInverser = Val(FicIni.Parametre("ToucheSoldatsInverser"))
    ToucheSoldatsSupprimer = Val(FicIni.Parametre("ToucheSoldatsSupprimer"))
    
    
    FicIni.Fichier = FicIni.Chemin & FichierINI2
    FicIni.Section = SectionINI2
    ToucheDialogueValider1 = Val(FicIni.Parametre("Valider1"))
    ToucheDialogueValider2 = Val(FicIni.Parametre("Valider2"))
    ToucheDialogueAnnuler1 = Val(FicIni.Parametre("Annuler1"))
    ToucheDialogueAnnuler2 = Val(FicIni.Parametre("Annuler2"))
    ToucheDialogueEffacerPrecedent = Val(FicIni.Parametre("EffacerPrecedent"))
    ToucheDialogueEffacerSuivant = Val(FicIni.Parametre("EffacerSuivant"))
    ToucheDialogueMessagePrecedent = Val(FicIni.Parametre("MessagePrecedent"))
    ToucheDialogueMessageSuivant = Val(FicIni.Parametre("MessageSuivant"))
    ToucheDialogueCurseurDebut = Val(FicIni.Parametre("CurseurDebut"))
    ToucheDialogueCurseurPrecedent = Val(FicIni.Parametre("CurseurPrecedent"))
    ToucheDialogueCurseurSuivant = Val(FicIni.Parametre("CurseurSuivant"))
    ToucheDialogueCurseurFin = Val(FicIni.Parametre("CurseurFin"))
    ReDim ToucheDialogueLettre(255)
    ReDim ToucheDialogueLettreMaj(255)
    For i = 1 To UBound(ToucheDialogueLettre())
        ToucheDialogueLettre(i) = FicIni.Parametre(i)
        ToucheDialogueLettreMaj(i) = FicIni.Parametre(i & "+")
    Next i
    
    Set TempSoldatsSelectionnes = New Collection
End Sub

Public Sub Presser_Touche(ByVal Key As Integer, _
                          ByVal KHaut As Boolean, _
                          ByVal KBAs As Boolean, _
                          ByVal KGauche As Boolean, _
                          ByVal KDroite As Boolean, _
                          ByRef Perso As ClsJeuPerso, _
                          ByRef Maison As ClsJeuBatiment, _
                          ByRef NoFief As Long, _
                          ByRef NoChateau As Long, _
                          ByRef Interface As ClsInterface, _
                          ByRef Message As ClsJeuMessages, _
                          ByRef NoPerso As Long, _
                          ByRef MaxPerso As Long, _
                          ByRef Quitter As Boolean, _
                          ByVal Ctrl As Boolean, _
                          ByVal Shift As Boolean, _
                          ByVal Alt As Boolean, _
                          ByRef ArmeeSelectionnee As Boolean, _
                          ByRef SoldatsSelectionnes As Collection, _
                          ByVal NoJoueur As Integer, _
                          ByVal DecalageX As Long, ByVal DecalageY As Long, _
                          ByVal Largeur As Integer, ByVal Hauteur As Integer)
    Dim i As Long, j As Long, k As Long
    Dim TempX As Long, TempY As Long
    'If Key = 0 Then Exit Sub
    'Alt-F4 : 56, 62. --> Quitte le programme.
    If Key = 62 And Alt Then
        End 'Quitter = True
    ElseIf Key = 15 And Alt Then
        Exit Sub 'Alt Tab : --> Bascule vers un autre programme.
    End If
    
    If Interface.FenetreConfirmation.Affiche Then
        Select Case Key
        Case ToucheDialogueValider1: Interface.FenetreConfirmation.Valider Perso
        Case ToucheDialogueValider2: Interface.FenetreConfirmation.Valider Perso
        Case ToucheDialogueAnnuler1: Interface.FenetreConfirmation.Annuler Interface, NoJoueur
        Case ToucheDialogueAnnuler2: Interface.FenetreConfirmation.Annuler Interface, NoJoueur
        End Select
    Else
    If Interface.AfficherDialogue Then
        'Gère la fenêtre de dialogue.
        Select Case Key
        Case ToucheDialogue: Interface.AfficherDialogue = False
        Case ToucheDialogueValider1: Interface.Dialogue_Valider Perso
        Case ToucheDialogueValider2: Interface.Dialogue_Valider Perso
        Case ToucheDialogueAnnuler1: Interface.Dialogue_Annuler
        Case ToucheDialogueAnnuler2: Interface.Dialogue_Annuler
        Case ToucheDialogueEffacerPrecedent: Interface.Dialogue_Effacer False
        Case ToucheDialogueEffacerSuivant: Interface.Dialogue_Effacer True
        Case ToucheDialogueMessagePrecedent: Interface.Dialogue_Message False
        Case ToucheDialogueMessageSuivant: Interface.Dialogue_Message True
        Case ToucheDialogueCurseurDebut: Interface.Dialogue_Curseur 0
        Case ToucheDialogueCurseurPrecedent: Interface.Dialogue_Curseur IIf(Ctrl, 0, 1)
        Case ToucheDialogueCurseurSuivant: Interface.Dialogue_Curseur IIf(Ctrl, 3, 2)
        Case ToucheDialogueCurseurFin: Interface.Dialogue_Curseur 3
        Case Else
            If Shift Then
                Interface.Dialogue_AjouterLettre ToucheDialogueLettreMaj(Key)
            Else
                Interface.Dialogue_AjouterLettre ToucheDialogueLettre(Key)
            End If
        End Select
    Else
    'Déplacement du personnage.
    If Perso.Action = 303 Then i = 1 'Permet de continuer de marcher en tirant
    If Perso.Vivant Then
        j = Perso.BonusDeplacement + 1
    Else
        j = Perso.Definir_VitesseDeplacementFantome + 1
    End If
    If KHaut Then
        If KGauche Then
            TempX = Perso.PositionX - j: TempY = Perso.PositionY - j
        ElseIf KDroite Then
            TempX = Perso.PositionX + j: TempY = Perso.PositionY - j
        Else
            TempX = Perso.PositionX: TempY = Perso.PositionY - j
        End If
    ElseIf KBAs Then
        If KGauche Then
            TempX = Perso.PositionX - j: TempY = Perso.PositionY + j
        ElseIf KDroite Then
            TempX = Perso.PositionX + j: TempY = Perso.PositionY + j
        Else
            TempX = Perso.PositionX: TempY = Perso.PositionY + j
        End If
        
    Else
        If KGauche Then
            TempX = Perso.PositionX - j: TempY = Perso.PositionY
        ElseIf KDroite Then
            TempX = Perso.PositionX + j: TempY = Perso.PositionY
        End If
    End If
    If KHaut Or KBAs Or KGauche Or KDroite Then
        If TempX < 0 Then
            TempX = 0
        ElseIf TempX > Monde.Largeur * Monde.LargeurCase - Perso.Largeur Then
            TempX = Monde.Largeur * Monde.LargeurCase - Perso.Largeur
        End If
        If TempY < 0 Then
            TempY = 0
        ElseIf TempY > Monde.Hauteur * Monde.HauteurCase - Perso.Hauteur Then
            TempY = Monde.Hauteur * Monde.HauteurCase - Perso.Hauteur
        End If
        Cmd_Perso_DeplacerSansCommentaire Perso.Numero, TempX, TempY
    End If
    If i = 1 Then Perso.Action = 303 'Permet de continuer de marcher en tirant
    
    'Ordres spéciaux pour les armées.
    If Ctrl And (ToucheSoldatsToutes = Key Or _
                 ToucheSoldatsMemeType = Key Or _
                 ToucheSoldatsEnArme = Key Or _
                 ToucheSoldatsEcran = Key Or _
                 ToucheSoldatsReposes = Key Or _
                 ToucheSoldatsEnArret = Key Or _
                 ToucheSoldatsRapides = Key Or _
                 ToucheSoldatsNonAffectees = Key Or _
                 ToucheSoldatsTerrestres = Key Or _
                 ToucheSoldatsVolants = Key Or _
                 ToucheSoldatsDefensifs = Key Or _
                 ToucheSoldatsOffensifs = Key Or _
                 ToucheSoldatsInverser = Key Or _
                 ToucheSoldatsSupprimer = Key) And _
                 Perso.NombreSoldats > 0 Then
    Ordonner_Activer_IA Perso.Numero
    Select Case Key
    Case ToucheSoldatsToutes:
        While SoldatsSelectionnes.Count > 0: SoldatsSelectionnes.Remove 1: Wend
        For i = 0 To UBound(Persos())
        If Persos(i).NumeroChef = NoPerso And Persos(i).Vivant And Not Persos(i).Incommandable And Not Persos(i).ChoisirIA Then
            SoldatsSelectionnes.Add i: If Persos(i).IA = True Then Cmd_Perso_Definir_IA i, False 'Persos(i).IA = False
        End If
        Next i
    Case ToucheSoldatsMemeType:
        If ArmeeSelectionnee Then
            For i = 0 To UBound(Persos())
            If Persos(i).NumeroChef = NoPerso And Persos(i).Vivant And Not Persos(i).Incommandable Then
                For j = 1 To SoldatsSelectionnes.Count
                    If Persos(i).Race = Persos(SoldatsSelectionnes.Item(j)).Race And _
                       i <> SoldatsSelectionnes.Item(j) Then
                        SoldatsSelectionnes.Add i: If Persos(i).IA = True Then Cmd_Perso_Definir_IA i, False 'Persos(i).IA = False
                        j = SoldatsSelectionnes.Count
                    End If
                Next j
            End If
            Next i
        End If
    Case ToucheSoldatsEnArme:
        While SoldatsSelectionnes.Count > 0: SoldatsSelectionnes.Remove 1: Wend
        For i = 0 To UBound(Persos())
        If (Persos(i).Objet_Equipes_Actif(0) Or Persos(i).Objet_Equipes_Actif(1)) And _
           Persos(i).NumeroChef = NoPerso And Persos(i).Vivant And Not Persos(i).Incommandable And Not Persos(i).ChoisirIA Then
            SoldatsSelectionnes.Add i: If Persos(i).IA = True Then Cmd_Perso_Definir_IA i, False 'Persos(i).IA = False
        End If
        Next i
    Case ToucheSoldatsEcran:
        FrmMoteur2D.Selectionner_Soldats Shift, False, True
    Case ToucheSoldatsReposes:
        While SoldatsSelectionnes.Count > 0: SoldatsSelectionnes.Remove 1: Wend
        For i = 0 To UBound(Persos())
        If Persos(i).Vie = Persos(i).Definir_MaxVie And _
           Persos(i).NumeroChef = NoPerso And Persos(i).Vivant And Not Persos(i).Incommandable And Not Persos(i).ChoisirIA Then
            SoldatsSelectionnes.Add i: If Persos(i).IA = True Then Cmd_Perso_Definir_IA i, False 'Persos(i).IA = False
        End If
        Next i
    Case ToucheSoldatsEnArret:
        While SoldatsSelectionnes.Count > 0: SoldatsSelectionnes.Remove 1: Wend
        For i = 0 To UBound(Persos())
        If Persos(i).Inactif And Persos(i).NumeroChef = NoPerso And Persos(i).Vivant And Not Persos(i).Incommandable And Not Persos(i).ChoisirIA Then
            SoldatsSelectionnes.Add i: If Persos(i).IA = True Then Cmd_Perso_Definir_IA i, False 'Persos(i).IA = False
        End If
        Next i
    Case ToucheSoldatsRapides:
        While SoldatsSelectionnes.Count > 0: SoldatsSelectionnes.Remove 1: Wend
        For i = 0 To UBound(Persos())
        If Persos(i).Longueur_Pas >= Perso.Longueur_Pas And Persos(i).NumeroChef = NoPerso And Persos(i).Vivant And Not Persos(i).Incommandable And Not Persos(i).ChoisirIA Then
            SoldatsSelectionnes.Add i: If Persos(i).IA = True Then Cmd_Perso_Definir_IA i, False 'Persos(i).IA = False
        End If
        Next i
    Case ToucheSoldatsNonAffectees:
        While SoldatsSelectionnes.Count > 0: SoldatsSelectionnes.Remove 1: Wend
        For i = 0 To UBound(Persos())
        If Persos(i).Attitude = 0 And Persos(i).IA_Ordre <> 4 And Not Persos(i).ChoisirIA And Persos(i).NumeroChef = NoPerso And Persos(i).Vivant And Not Persos(i).Incommandable And Not Persos(i).ChoisirIA Then
            SoldatsSelectionnes.Add i: If Persos(i).IA = True Then Cmd_Perso_Definir_IA i, False 'Persos(i).IA = False
        End If
        Next i
    Case ToucheSoldatsTerrestres:
        While SoldatsSelectionnes.Count > 0: SoldatsSelectionnes.Remove 1: Wend
        For i = 0 To UBound(Persos())
        If Not Persos(i).Definir_Volant And Persos(i).NumeroChef = NoPerso And Persos(i).Vivant And Not Persos(i).Incommandable And Not Persos(i).ChoisirIA Then
            SoldatsSelectionnes.Add i: If Persos(i).IA = True Then Cmd_Perso_Definir_IA i, False 'Persos(i).IA = False
        End If
        Next i
    Case ToucheSoldatsVolants:
        While SoldatsSelectionnes.Count > 0: SoldatsSelectionnes.Remove 1: Wend
        For i = 0 To UBound(Persos())
        If Persos(i).Definir_Volant And Persos(i).NumeroChef = NoPerso And Persos(i).Vivant And Not Persos(i).Incommandable And Not Persos(i).ChoisirIA Then
            SoldatsSelectionnes.Add i: If Persos(i).IA = True Then Cmd_Perso_Definir_IA i, False 'Persos(i).IA = False
        End If
        Next i
    Case ToucheSoldatsDefensifs:
        While SoldatsSelectionnes.Count > 0: SoldatsSelectionnes.Remove 1: Wend
        For i = 0 To UBound(Persos())
        If Persos(i).Definir_Attaque < Persos(i).Definir_Defense And Persos(i).NumeroChef = NoPerso And Persos(i).Vivant And Not Persos(i).Incommandable And Not Persos(i).ChoisirIA Then
            SoldatsSelectionnes.Add i: If Persos(i).IA Then Cmd_Perso_Definir_IA i, False  'Persos(i).IA = False
        End If
        Next i
    Case ToucheSoldatsOffensifs:
        While SoldatsSelectionnes.Count > 0: SoldatsSelectionnes.Remove 1: Wend
        For i = 0 To UBound(Persos())
        If Persos(i).Definir_Attaque >= Persos(i).Definir_Defense And Persos(i).NumeroChef = NoPerso And Persos(i).Vivant And Not Persos(i).Incommandable And Not Persos(i).ChoisirIA Then
            SoldatsSelectionnes.Add i: If Persos(i).IA Then Cmd_Perso_Definir_IA i, False  'Persos(i).IA = False
        End If
        Next i
    Case ToucheSoldatsInverser:
        For i = 1 To SoldatsSelectionnes.Count
            TempSoldatsSelectionnes.Add SoldatsSelectionnes(i)
        Next i
        While SoldatsSelectionnes.Count > 0: SoldatsSelectionnes.Remove 1: Wend
        For i = 0 To UBound(Persos())
            If Persos(i).NumeroChef = NoPerso And Persos(i).Vivant And Not Persos(i).Incommandable And Not Persos(i).ChoisirIA Then
                SoldatsSelectionnes.Add i: If Persos(i).IA Then Cmd_Perso_Definir_IA i, False
            End If
        Next i
        For i = 1 To TempSoldatsSelectionnes.Count
            k = 1
            For j = 1 To SoldatsSelectionnes.Count
                If TempSoldatsSelectionnes(i) = SoldatsSelectionnes(k) Then
                    If Not Persos(SoldatsSelectionnes(k)).IA Then
                        Cmd_Perso_Definir_IA SoldatsSelectionnes(k), True
                    End If
                    SoldatsSelectionnes.Remove k
                Else
                    k = k + 1
                End If
            Next j
        Next i
        While TempSoldatsSelectionnes.Count > 0: TempSoldatsSelectionnes.Remove 1: Wend
    Case ToucheSoldatsSupprimer:
        If ArmeeSelectionnee Then
            For i = 1 To SoldatsSelectionnes.Count
                If SoldatsSelectionnes.Item(i) <> Perso.Numero Then
                    Cmd_Perso_TuerPerso Perso.Numero, SoldatsSelectionnes.Item(i)
                End If
            Next i
            While SoldatsSelectionnes.Count > 0: SoldatsSelectionnes.Remove 1: Wend
        End If
    End Select
    ArmeeSelectionnee = SoldatsSelectionnes.Count > 0
    
    'Autres touches avec Ctrl.
    ElseIf Ctrl And (ToucheMusique = Key) Then
    Select Case Key
    Case ToucheMusique:
        Musique.Definir_JouerMusique = Not Musique.Definir_JouerMusique
        If Musique.Definir_JouerMusique Then Musique.Jouer_Musique 1, False
    End Select
    
    'Touches simples.
    ElseIf Key > 0 Then
    Select Case Key
    Case ToucheQuitter: 'Quitter = True
                        Interface.FenetreConfirmation.Afficher_Fenetre 0, Interface.MessageConfirmationQuitter, 0, DecalageY, Largeur, Hauteur
    Case ToucheZoomNormal: FrmMoteur2D.Zoom = 1
    Case ToucheZoomPlus: FrmMoteur2D.Zoom = FrmMoteur2D.Zoom + 0.5
    Case ToucheZoomMoins: FrmMoteur2D.Zoom = FrmMoteur2D.Zoom - 0.5: If FrmMoteur2D.Zoom < 1 Then FrmMoteur2D.Zoom = 1
    Case ToucheFenetreBatiment: Interface.Actionner_Boutons_Infos 0, NoJoueur
    Case ToucheDialogue: Interface.AfficherDialogue = True
    Case ToucheFenetreRessources: Interface.Actionner_Boutons_Infos 1, NoJoueur
    Case ToucheFenetrePersonnage: Interface.Actionner_Boutons_Infos 2, NoJoueur
    Case ToucheFenetreCompetences: Interface.Actionner_Boutons_Infos 3, NoJoueur
    Case ToucheFenetreInventaire: Interface.Actionner_Boutons_Infos 4, NoJoueur
    Case ToucheFenetreBestiaire: Interface.Actionner_Boutons_Infos 5, NoJoueur
    Case ToucheFenetreCarte: Interface.Actionner_Boutons_Infos 6, NoJoueur
    Case ToucheFenetreObjectifs: If Interface.AfficherObjectifs Then Interface.Actionner_Boutons_Infos 8, NoJoueur
    Case ToucheFenetresFermer: Interface.Fermer_Toutes_les_fenetres: Message.Effacer
    Case ToucheRetourChateau: Cmd_Perso_Rentrer_Chateau Perso.Numero 'Perso.Rentrer_Chateau Chateau
    Case ToucheAction:
        With Perso
        If .DansSonChateau And .PeutDevenir_Empereur(Fiefs(.NumeroFief), Parametres) Then
            Cmd_Perso_DevenirEmpereur .Numero
        ElseIf (.DansUneMaison Or .DansSaMaison) And _
            Maisons(.IndiceMaison).Service And _
            Maisons(.IndiceMaison).Service_Utilisable(Perso, Parametres) Then
            'Maisons(.IndiceMaison).Service_Utiliser Perso, Parametres
            Cmd_Perso_UtiliserService .Numero, .IndiceMaison
        Else
            Select Case Obstacle(.Numero, .PositionX, .PositionY, .Largeur, .Hauteur, True)
                Case 1: If .Vivant Then Cmd_Perso_Chercher_Ressource .Numero, IndiceObstacle
                Case 4, 5: Cmd_Perso_Rentrer_Maison .Numero, Ctrl And Maisons(.Numero).Construit
                Case 6: Cmd_Perso_Rentrer_Chateau .Numero
                Case 7: If .Vivant Then Cmd_Perso_AllerMaison .Numero, IndiceObstacle, Ctrl '.Aller_Maison Maisons(IndiceObstacle)
                Case 8: If .Vivant Then Cmd_Perso_AllerChateau .Numero, IndiceObstacle '.Aller_Chateau Chateaux(IndiceObstacle), IndiceObstacle
                Case 9: If .Vivant Then Cmd_Perso_AttaquerPerso .Numero, IndiceObstacle '.Attaquer_Perso IndiceObstacle
                Case 10: If .Vivant Then Cmd_Perso_AttaquerMaison .Numero, IndiceObstacle '.Attaquer_Maison IndiceObstacle
                Case 11: If .Vivant Then Cmd_Perso_AttaquerChateau .Numero, IndiceObstacle '.Attaquer_Chateau IndiceObstacle
                Case 12: Cmd_Perso_AllerMaison .Numero, IndiceObstacle, False  '.Aller_Maison Maisons(IndiceObstacle)
                Case 13: If .Vivant Then Cmd_Perso_AllerTresor .Numero, IndiceObstacle   '.Aller_Tresor Tresors(IndiceObstacle), IndiceObstacle, AffTresor.Largeur, AffTresor.Hauteur
            End Select
        End If
        End With
    Case ToucheRetourMaison:
        If Perso.Vivant Or Maison.Service = 1 Then
            'Perso.Rentrer_Maison Maison
            Cmd_Perso_Rentrer_Maison Perso.Numero, Ctrl And Maisons(Perso.Numero).Construit
        End If
    Case ToucheTrouverRessource1: If Perso.Vivant Then Cmd_Perso_TrouverRessource Perso.Numero, 0 'IA_Trouver_Ressource Perso.Numero, 0
    Case ToucheTrouverRessource2: If Perso.Vivant Then Cmd_Perso_TrouverRessource Perso.Numero, 1 'IA_Trouver_Ressource Perso.Numero, 1
    Case ToucheTrouverRessource3: If Perso.Vivant Then Cmd_Perso_TrouverRessource Perso.Numero, 2 'IA_Trouver_Ressource Perso.Numero, 2
    Case ToucheTrouverRessource4: If Perso.Vivant Then Cmd_Perso_TrouverRessource Perso.Numero, 3 'IA_Trouver_Ressource Perso.Numero, 3
    Case ToucheTrouverRessource5: If Perso.Vivant Then Cmd_Perso_TrouverRessource Perso.Numero, 4 'IA_Trouver_Ressource Perso.Numero, 4
    Case ToucheTrouverRessource6: If Perso.Vivant Then Cmd_Perso_TrouverRessource Perso.Numero, 5 'IA_Trouver_Ressource Perso.Numero, 5
    Case ToucheTrouverRessource7: If Perso.Vivant Then Cmd_Perso_TrouverRessource Perso.Numero, 6 'IA_Trouver_Ressource Perso.Numero, 6
    Case ToucheTrouverRessource8: If Perso.Vivant Then Cmd_Perso_TrouverRessource Perso.Numero, 7 'IA_Trouver_Ressource Perso.Numero, 7
    Case ToucheTrouverAttaquerPerso:
        If Perso.Vivant Then
            If ArmeeSelectionnee Then
                Ordonner_Attaquer_Perso Perso, SoldatsSelectionnes, Trouver_Ennemi(Perso.Numero)
            Else
                'IA_Trouver_Ennemi Perso.Numero, Ctrl Or Shift Or Alt, Ctrl, Shift, Alt
                Cmd_Perso_TrouverPersoEnnemi Perso.Numero, Ctrl Or Shift Or Alt, Ctrl, Shift, Alt
            End If
        End If
    Case ToucheTrouverAttaquerMaison:
        If Perso.Vivant Then
            If ArmeeSelectionnee Then
                Ordonner_Attaquer_Maison Perso, SoldatsSelectionnes, Trouver_MaisonEnnemi(Perso.Numero)
            Else
                'IA_Trouver_MaisonEnnemi Perso.Numero, Ctrl Or Shift Or Alt, Ctrl, Shift, Alt
                Cmd_Perso_TrouverMaisonEnnemi Perso.Numero, Ctrl Or Shift Or Alt, Ctrl, Shift, Alt
            End If
        End If
    Case ToucheTrouverAttaquerChateau:
        If Perso.Vivant Then
            If ArmeeSelectionnee Then
                Ordonner_Attaquer_Chateau Perso, SoldatsSelectionnes, Trouver_ChateauEnnemi(Perso.Numero)
            Else
                'IA_Trouver_ChateauEnnemi Perso.Numero, Ctrl Or Shift Or Alt, Ctrl, Shift, Alt
                Cmd_Perso_TrouverChateauEnnemi Perso.Numero, Ctrl Or Shift Or Alt, Ctrl, Shift, Alt
            End If
        End If
    Case ToucheTrouverTresor:
        If Perso.Vivant Then
            'IA_Trouver_Tresor Perso.Numero
            Cmd_Perso_TrouverTresor Perso.Numero
        End If
    Case ToucheModeAuto:
        'Perso.IA = Not Perso.IA
        Cmd_Perso_Definir_IA Perso.Numero, Not Perso.IA
    Case ToucheAide:
        If Perso.Vivant Then
            'Appeler_De_lAide Perso.Numero
            'Cmd_Perso_AppelerDelAide Perso.Numero
            'Cmd_Perso_AttaquerAllies Perso.Numero
            Interface.FenetreConfirmation.Afficher_Fenetre 1, Interface.MessageConfirmationAttaquer, 0, DecalageY, Largeur, Hauteur
        End If
    Case ToucheArret:
        If Ctrl Or Shift Or Alt Then
            If Perso.Vivant Then
                Ordonner_Arret Perso.Numero, False, SoldatsSelectionnes
            End If
        ElseIf ArmeeSelectionnee Then
            If Perso.Vivant Then
                Ordonner_Arret Perso.Numero, ArmeeSelectionnee, SoldatsSelectionnes
            End If
        Else
            'Perso.Arreter
            Cmd_Perso_Arreter Perso.Numero
        End If
            
    'Case ToucheChargez: If Perso.Vivant Then Ordonner_Chargez Perso.Numero, , , Ctrl, Shift, Alt
    Case ToucheDefinirAttitude1:
        Interface.Bouton_ChangerAttitude 0, Perso, ArmeeSelectionnee, SoldatsSelectionnes
    Case ToucheDefinirAttitude2:
        Interface.Bouton_ChangerAttitude 1, Perso, ArmeeSelectionnee, SoldatsSelectionnes
    Case ToucheDefinirAttitude3:
        Interface.Bouton_ChangerAttitude 2, Perso, ArmeeSelectionnee, SoldatsSelectionnes
    Case ToucheDefinirAttitude4:
        Interface.Bouton_ChangerAttitude 3, Perso, ArmeeSelectionnee, SoldatsSelectionnes
    Case ToucheDefinirAttitude5:
        Interface.Bouton_ChangerAttitude 4, Perso, ArmeeSelectionnee, SoldatsSelectionnes
    Case ToucheDefinirFormation:
        'If Ctrl Or Shift Or Alt Then
        '    Ordonner_Formation Perso.Numero
        'Else
            'Ordonner_Formation Perso.Numero, True
            Cmd_Perso_OrdonnerFormation Perso.Numero, True
        'End If

    Case ToucheArgentDeposer:
        If Perso.DansSaMaison Then
            'If Shift Then
            '    Maison.Transfert_Argent_Deposer 5, Perso
            'ElseIf Ctrl Then
            '    Maison.Transfert_Argent_Deposer 1, Perso
            'Else
                'Maison.Transfert_Argent_Deposer 0, Perso, True
                Cmd_Perso_ArgentMaisonLaisser Perso.Numero, Perso.Definir_Argent
            'End If
        End If
    Case ToucheArgentRecuperer:
        If Perso.DansSaMaison Then
            'If Shift Then
            '    Maison.Transfert_Argent_Prendre 5, Perso
            'ElseIf Ctrl Then
            '    Maison.Transfert_Argent_Prendre 1, Perso
            'Else
                Cmd_Perso_ArgentMaisonPrendre Perso.Numero, Maison.Argent
                'Maison.Transfert_Argent_Prendre 0, Perso, True
            'End If
        End If

    Case ToucheDemolirMaison:
        'If Maison.Visible And Perso.Vivant Then
        Cmd_Perso_DetuireMaison Perso.Numero
        'Maison.Demolir Perso, Parametres
        'Interface.Actualiser_BoutonsConstruire Perso, Maison, Fief, Parametres
        'End If
    Case ToucheAfficherConseil: Messages(0).Afficher_Conseil
    Case ToucheMenu:
        Interface.Actionner_Boutons_Infos 7, NoJoueur
        'Jeu.EnPause = Not Jeu.EnPause
        'FrmMoteur2D.FrameMenu.Visible = Jeu.EnPause
    Case ToucheCaptureEcran: FrmMoteur2D.GigaCapture
    Case ToucheSauvegarderMonde: If Not Jeu.Campagne Then Enregistrer_Sauvegarde_Monde
    Case ToucheInterfacePleinEcran:
        'Interface.AfficherBoutonsConstruire = Not Interface.AfficherBoutonsConstruire
        Interface.Affichage_PleinEcran = Not Interface.Affichage_PleinEcran
    Case ToucheChangerPersonnage:
        Enregistrer_Sauvegarde_Interface Persos(Nopersos(0)), Interface
        Persos(Nopersos(0)).ControleJoueur = False
        'i = Persos(Nopersos(0)).NumeroEquipe
        If Shift Then 'Personnage précédent.
            Do
                Nopersos(0) = Nopersos(0) - 1
                If Nopersos(0) < 0 Then Nopersos(0) = MaxPerso - 1
            Loop Until Persos(Nopersos(0)).Joueur = ComReseau.PlayerID 'And _
                 Persos(Nopersos(0)).NumeroEquipe = i
        Else 'Personnage suivant.
            Do
                Nopersos(0) = Nopersos(0) + 1
                If Nopersos(0) > UBound(Persos()) Then Nopersos(0) = 0
            Loop Until Persos(Nopersos(0)).Joueur = ComReseau.PlayerID 'And _
                 Persos(Nopersos(0)).NumeroEquipe = i
        End If
        Persos(Nopersos(0)).ControleJoueur = True
        Charger_Sauvegarde_Interface Persos(Nopersos(0)), Interface
'    Case ToucheDirectionHaut: 'Perso.Aller_A Perso.PositionX, Perso.PositionY - Perso.Longueur_Pas
'                              Cmd_Perso_DeplacerSansCommentaire Perso.Numero, Perso.PositionX, Perso.PositionY - Perso.Longueur_Pas * Monde.CoefVitesse(Perso, Perso.PositionX, Perso.PositionY + Perso.Hauteur - AffTerrain.Hauteur)
'    Case ToucheDirectionBas: 'Perso.Aller_A Perso.PositionX, Perso.PositionY + Perso.Longueur_Pas
'                             Cmd_Perso_DeplacerSansCommentaire Perso.Numero, Perso.PositionX, Perso.PositionY + Perso.Longueur_Pas * Monde.CoefVitesse(Perso, Perso.PositionX, Perso.PositionY + Perso.Hauteur - AffTerrain.Hauteur)
'    Case ToucheDirectionGauche: 'Perso.Aller_A Perso.PositionX - Perso.Longueur_Pas, Perso.PositionY
'                                Cmd_Perso_DeplacerSansCommentaire Perso.Numero, Perso.PositionX - Perso.Longueur_Pas * Monde.CoefVitesse(Perso, Perso.PositionX, Perso.PositionY + Perso.Hauteur - AffTerrain.Hauteur), Perso.PositionY
'    Case ToucheDirectionDroite: 'Perso.Aller_A Perso.PositionX + Perso.Longueur_Pas, Perso.PositionY
'                                Cmd_Perso_DeplacerSansCommentaire Perso.Numero, Perso.PositionX + Perso.Longueur_Pas * Monde.CoefVitesse(Perso, Perso.PositionX, Perso.PositionY + Perso.Hauteur - AffTerrain.Hauteur), Perso.PositionY
    'On s'équipe avec les touches de raccourci.
    Case ToucheObjets To NombreTouchesObjets + ToucheObjets - 1:
        If Not Interface.CurseurSurObjet Then
            Interface.Raccourcis_Objets_Equiper Key - 1, Perso, Ctrl, Shift, Alt
            'Interface.Actualiser_Raccourcis_Objets Perso
            Interface.ActualiserRaccourciObjets = True
        End If
    End Select
    End If
    End If
    End If
End Sub
