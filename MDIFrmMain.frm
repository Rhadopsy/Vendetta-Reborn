VERSION 5.00
Begin VB.MDIForm MDIFrmMain 
   BackColor       =   &H8000000C&
   ClientHeight    =   2910
   ClientLeft      =   165
   ClientTop       =   330
   ClientWidth     =   4980
   Icon            =   "MDIFrmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu MenuPartie 
      Caption         =   "&Partie"
      Enabled         =   0   'False
      Begin VB.Menu MenuPartieMonde 
         Caption         =   "&Campagne"
         Index           =   0
      End
      Begin VB.Menu MenuPartieMonde 
         Caption         =   "&Monde personnalisé"
         Index           =   1
         Shortcut        =   ^P
      End
      Begin VB.Menu MenuPartieMonde 
         Caption         =   "Champ de &bataille"
         Index           =   2
         Shortcut        =   ^B
      End
      Begin VB.Menu MenuPartieMonde 
         Caption         =   "Charger &scénario"
         Index           =   3
         Shortcut        =   ^S
      End
      Begin VB.Menu MenuPartieMonde 
         Caption         =   "Char&ger partie"
         Index           =   4
         Shortcut        =   ^G
      End
      Begin VB.Menu MenuPartieMonde 
         Caption         =   "Mode 2 &joueurs"
         Index           =   5
         Shortcut        =   ^J
      End
      Begin VB.Menu MenuPartieSeparateur1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuPartieCommencer 
         Caption         =   "&Commencer"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MenuPartieSeparateur2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuPartieQuitter 
         Caption         =   "&Quitter"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MenuMonde 
      Caption         =   "&Mondes"
      Enabled         =   0   'False
      Begin VB.Menu MenuMondeRenommer 
         Caption         =   "&Renommer"
      End
      Begin VB.Menu MenuMondeCommentaires 
         Caption         =   "&Commentaires"
      End
      Begin VB.Menu MenuMondeEditer 
         Caption         =   "&Editer"
      End
      Begin VB.Menu MenuMondeSeparateur1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuMondeSupprimer 
         Caption         =   "&Supprimer"
      End
   End
   Begin VB.Menu MenuPersonnage 
      Caption         =   "P&ersonnages"
      Enabled         =   0   'False
      Begin VB.Menu MenuPersonnageNouveau 
         Caption         =   "&Nouveau"
         Shortcut        =   ^N
      End
      Begin VB.Menu MenuPersonnageCharger 
         Caption         =   "&Charger"
      End
      Begin VB.Menu MenuPersonnageExporter 
         Caption         =   "&Exporter"
      End
      Begin VB.Menu MenuPersonnageSeparateur1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuPersonnageSelectionner 
         Caption         =   "Sé&lectionner"
         Begin VB.Menu MenuPersonnageSelectionnerTous 
            Caption         =   "&Tous"
            Shortcut        =   ^A
         End
         Begin VB.Menu MenuPersonnageSeparateur2 
            Caption         =   "-"
         End
         Begin VB.Menu MenuPersonnageSelectionMeilleur 
            Caption         =   "Meilleur &Niveau"
            Index           =   0
            Shortcut        =   ^I
         End
         Begin VB.Menu MenuPersonnageSelectionMeilleur 
            Caption         =   "Meilleur Arti&san"
            Index           =   1
         End
         Begin VB.Menu MenuPersonnageSelectionMeilleur 
            Caption         =   "Meilleur T&ueur"
            Index           =   2
         End
         Begin VB.Menu MenuPersonnageSelectionMeilleur 
            Caption         =   "Meilleur &Guerrier (Max Vie)"
            Index           =   3
         End
         Begin VB.Menu MenuPersonnageSelectionMeilleur 
            Caption         =   "Meilleur &Archer (Max Energie)"
            Index           =   4
         End
         Begin VB.Menu MenuPersonnageSelectionMeilleur 
            Caption         =   "Meilleur &Magicien (Max Magie)"
            Index           =   5
         End
         Begin VB.Menu MenuPersonnageSelectionMeilleur 
            Caption         =   "Meilleur &Seigneur (Max Moral)"
            Index           =   6
         End
      End
      Begin VB.Menu MenuPersonnageSeparateur4 
         Caption         =   "-"
      End
      Begin VB.Menu MenuPersonnageCompetences 
         Caption         =   "Compé&tences"
         Begin VB.Menu MenuPersonnageCompetencesAfficherToutes 
            Caption         =   "A&fficher toutes"
         End
         Begin VB.Menu MenuPersonnageCompetencesSeparateur1 
            Caption         =   "-"
         End
         Begin VB.Menu MenuPersonnageCompetencesTri 
            Caption         =   "&Aucun Tri"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu MenuPersonnageCompetencesTri 
            Caption         =   "Tri par &Nom"
            Index           =   1
         End
         Begin VB.Menu MenuPersonnageCompetencesTri 
            Caption         =   "&Tri par Score"
            Index           =   2
         End
         Begin VB.Menu MenuPersonnageCompetencesSeparateur2 
            Caption         =   "-"
         End
         Begin VB.Menu MenuPersonnageCompetencesAffichage 
            Caption         =   "&Score"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu MenuPersonnageCompetencesAffichage 
            Caption         =   "N&iveau"
            Index           =   1
         End
         Begin VB.Menu MenuPersonnageCompetencesAffichage 
            Caption         =   "&Expérience"
            Index           =   2
         End
      End
      Begin VB.Menu MenuPersonnageSeparateur5 
         Caption         =   "-"
      End
      Begin VB.Menu MenuPersonnageSupprimer 
         Caption         =   "&Supprimer"
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu MenuListes 
      Caption         =   "&Listes"
      Enabled         =   0   'False
      Begin VB.Menu MenuAideBatiments 
         Caption         =   "&Batiments"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MenuAideBestiaire 
         Caption         =   "B&estiaire"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MenuAideObjets 
         Caption         =   "&Objets"
         Shortcut        =   {F3}
      End
      Begin VB.Menu MenuAideRessources 
         Caption         =   "&Ressources"
         Shortcut        =   {F4}
      End
      Begin VB.Menu MenuAideRaces 
         Caption         =   "R&aces"
         Shortcut        =   {F6}
      End
      Begin VB.Menu MenuAidePeuples 
         Caption         =   "&Peuples"
         Shortcut        =   {F7}
      End
      Begin VB.Menu MenuAideSeparateur1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAideStastistiques 
         Caption         =   "&Stastistiques"
         Shortcut        =   {F8}
      End
      Begin VB.Menu MenuDocumentsSeparateur2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuDocumentsArtWorks 
         Caption         =   "&Artworks"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu MenuOptions 
      Caption         =   "&Options"
      Enabled         =   0   'False
      Begin VB.Menu MenuOptionParametres 
         Caption         =   "&Paramètres"
         Shortcut        =   ^O
      End
      Begin VB.Menu MenuOptionMusique 
         Caption         =   "&Musique"
         Shortcut        =   ^M
      End
      Begin VB.Menu MenuOptionsSeparateur 
         Caption         =   "-"
      End
      Begin VB.Menu MenuOptionEditeurCartes 
         Caption         =   "Editer &Cartes"
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu MenuOptionEditeurIA 
         Caption         =   "Editer &IAs"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu MenuAide 
      Caption         =   "&?"
      Enabled         =   0   'False
      Begin VB.Menu MenuAideResume 
         Caption         =   "&Resume"
         Shortcut        =   {F11}
      End
      Begin VB.Menu MenuAideSeparateur2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MenuAideManuel 
         Caption         =   "Aide de &Jeu"
         Shortcut        =   {F12}
      End
      Begin VB.Menu MenuAideSeparateur3 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAideAPropos 
         Caption         =   "A &propos"
      End
   End
End
Attribute VB_Name = "MDIFrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
'Implements DirectPlay8Event

Const FichierINI = "Fenetre_MDIFrmMain"
Const SectionINI = "Fenetre_MDIFrmMain"

Private TitreEffacerMonde As String
Private MessageEffacerMonde As String
Private CheminFichierFAQ As String

Public AffichageCompetences As Integer
Public TriCompetences As Integer

Private Sub MDIForm_Load()
    'On Error GoTo Erreur
    
    Hide
    DoEvents
    
    'Charge les paramètres.
    FicIni.fichier = FicIni.chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    
    TitreEffacerMonde = FicIni.Parametre("TitreEffacerMonde")
    MessageEffacerMonde = FicIni.Parametre("MessageEffacerMonde")
    CheminFichierFAQ = FicIni.Parametre("CheminFichierFAQ")
    
    MenuPartie.Caption = FicIni.Parametre("MenuPartie")
    MenuPartieMonde(0).Caption = FicIni.Parametre("MenuPartieMonde(0)")
    MenuPartieMonde(1).Caption = FicIni.Parametre("MenuPartieMonde(1)")
    MenuPartieMonde(2).Caption = FicIni.Parametre("MenuPartieMonde(2)")
    MenuPartieMonde(3).Caption = FicIni.Parametre("MenuPartieMonde(3)")
    MenuPartieMonde(4).Caption = FicIni.Parametre("MenuPartieMonde(4)")
    MenuPartieMonde(5).Caption = FicIni.Parametre("MenuPartieMonde(5)")
    MenuPartieCommencer.Caption = FicIni.Parametre("MenuPartieCommencer")
    MenuPartieQuitter.Caption = FicIni.Parametre("MenuPartieQuitter")
    
    MenuMonde.Caption = FicIni.Parametre("MenuMonde")
    MenuMondeRenommer.Caption = FicIni.Parametre("MenuMondeRenommer")
    MenuMondeCommentaires.Caption = FicIni.Parametre("MenuMondeCommentaires")
    MenuMondeEditer.Caption = FicIni.Parametre("MenuMondeEditer")
    MenuMondeSupprimer.Caption = FicIni.Parametre("MenuMondeSupprimer")
    
    MenuPersonnage.Caption = FicIni.Parametre("MenuPersonnage")
    MenuPersonnageNouveau.Caption = FicIni.Parametre("MenuPersonnageNouveau")
    MenuPersonnageCharger.Caption = FicIni.Parametre("MenuPersonnageCharger")
    MenuPersonnageExporter.Caption = FicIni.Parametre("MenuPersonnageExporter")
    MenuPersonnageSelectionner.Caption = FicIni.Parametre("MenuPersonnageSelectionner")
    MenuPersonnageSelectionnerTous.Caption = FicIni.Parametre("MenuPersonnageSelectionnerTous")
    MenuPersonnageSelectionMeilleur(0).Caption = FicIni.Parametre("MenuPersonnageSelectionMeilleur(0)")
    MenuPersonnageSelectionMeilleur(1).Caption = FicIni.Parametre("MenuPersonnageSelectionMeilleur(1)")
    MenuPersonnageSelectionMeilleur(2).Caption = FicIni.Parametre("MenuPersonnageSelectionMeilleur(2)")
    MenuPersonnageSelectionMeilleur(3).Caption = FicIni.Parametre("MenuPersonnageSelectionMeilleur(3)")
    MenuPersonnageSelectionMeilleur(4).Caption = FicIni.Parametre("MenuPersonnageSelectionMeilleur(4)")
    MenuPersonnageSelectionMeilleur(5).Caption = FicIni.Parametre("MenuPersonnageSelectionMeilleur(5)")
    MenuPersonnageSelectionMeilleur(6).Caption = FicIni.Parametre("MenuPersonnageSelectionMeilleur(6)")
    MenuPersonnageCompetences.Caption = FicIni.Parametre("MenuPersonnageCompetences")
    MenuPersonnageCompetencesAfficherToutes.Caption = FicIni.Parametre("MenuPersonnageCompetencesAfficherToutes")
    MenuPersonnageCompetencesTri(0).Caption = FicIni.Parametre("MenuPersonnageCompetencesTri(0)")
    MenuPersonnageCompetencesTri(1).Caption = FicIni.Parametre("MenuPersonnageCompetencesTri(1)")
    MenuPersonnageCompetencesTri(2).Caption = FicIni.Parametre("MenuPersonnageCompetencesTri(2)")
    MenuPersonnageCompetencesAffichage(0).Caption = FicIni.Parametre("MenuPersonnageCompetencesAffichage(0)")
    MenuPersonnageCompetencesAffichage(1).Caption = FicIni.Parametre("MenuPersonnageCompetencesAffichage(1)")
    MenuPersonnageCompetencesAffichage(2).Caption = FicIni.Parametre("MenuPersonnageCompetencesAffichage(2)")
    MenuPersonnageSupprimer.Caption = FicIni.Parametre("MenuPersonnageSupprimer")
    
    MenuOptions.Caption = FicIni.Parametre("MenuOptions")
    MenuOptionParametres.Caption = FicIni.Parametre("MenuOptionParametres")
    MenuOptionMusique.Caption = FicIni.Parametre("MenuOptionMusique")
    MenuOptionMusique.Checked = Musique.Definir_JouerMusique
    MenuOptionEditeurCartes.Caption = FicIni.Parametre("MenuOptionEditeurCartes")
    MenuOptionEditeurIA.Caption = FicIni.Parametre("MenuOptionEditeurIA")
    MenuAide.Caption = FicIni.Parametre("MenuAide")
    
    MenuListes.Caption = FicIni.Parametre("MenuListes")
    MenuAideBatiments.Caption = FicIni.Parametre("MenuAideBatiments")
    MenuAideBestiaire.Caption = FicIni.Parametre("MenuAideBestiaire")
    MenuAideObjets.Caption = FicIni.Parametre("MenuAideObjets")
    MenuAideRessources.Caption = FicIni.Parametre("MenuAideRessources")
    MenuAideRaces.Caption = FicIni.Parametre("MenuAideRaces")
    MenuAidePeuples.Caption = FicIni.Parametre("MenuAidePeuples")
    MenuAideStastistiques.Caption = FicIni.Parametre("MenuAideStastistiques")
    MenuAideResume.Caption = FicIni.Parametre("MenuAideResume")
    MenuAideManuel.Caption = FicIni.Parametre("MenuAideManuel")
    MenuAideAPropos.Caption = FicIni.Parametre("MenuAideAPropos")
    
    If Screen.Height > 11500 Then
        'Me.WindowState = 2
        'Me.WindowState = 0
        Width = 15485 - 16
        Height = 11200 + 16 + 120
        'Me.Top = 0
        'Me.Left = 0
    Else
        Me.WindowState = 2
    End If

    'Démarrage automatique.
    'Me.Picture = LoadPicture(App.Path & "\Images\Fenetres\Titre.bmp")
    
    Show
    DoEvents
    
    'Exit Sub
'Erreur:
    'Exit Sub
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    MenuPartieQuitter_Click
End Sub

Private Sub MenuAideManuel_Click()
    Editer_Texte App.Path & FicIni.chemin & Langues.Dossier & CheminFichierFAQ
End Sub

Public Sub MenuAideRaces_Click()
    FrmInfoRaces.Show
    FrmInfoRaces.ZOrder 0
End Sub

Private Sub MenuAideResume_Click()
    Document_Resume
End Sub

Private Sub MenuAideStastistiques_Click()
    Afficher_Stastistiques
End Sub

Private Sub MenuDocumentsArtWorks_Click()
    Document_Artworks
End Sub

Private Sub MenuMondeCommentaires_Click()
    'Set FrmParam.TabStripMonde.SelectedItem = FrmParam.TabStripMonde.Tabs("ChargerMonde")
    If FrmParam.LstMonde.ListCount > 0 Then
        Monde_Changer_Commentaires FrmParam.LstMonde.List(FrmParam.LstMonde.ListIndex), FrmParam.TabStripMonde.Tabs("ChargerScenario").Selected
    End If
End Sub

Private Sub MenuMondeEditer_Click()
    If FrmParam.LstMonde.ListCount > 0 Then
        Scenario_Editer FrmParam.LstMonde.List(FrmParam.LstMonde.ListIndex)
    End If
End Sub

Private Sub MenuMondeRenommer_Click()
    'Set FrmParam.TabStripMonde.SelectedItem = FrmParam.TabStripMonde.Tabs("ChargerMonde")
    If FrmParam.LstMonde.ListCount > 0 Then
        Monde_Changer_Nom FrmParam.LstMonde.List(FrmParam.LstMonde.ListIndex), FrmParam.TabStripMonde.Tabs("ChargerScenario").Selected
    End If
    FrmParam.Actualiser_Liste_Mondes
End Sub

Private Sub MenuMondeSupprimer_Click()
    'Set FrmParam.TabStripMonde.SelectedItem = FrmParam.TabStripMonde.Tabs("ChargerMonde")
    If FrmParam.LstMonde.ListCount > 0 Then
        If MsgBox(MessageEffacerMonde & " " & FrmParam.LstMonde.List(FrmParam.LstMonde.ListIndex) & " ?", vbOKCancel + vbQuestion, TitreEffacerMonde) = vbOK Then
            Monde_Supprimer FrmParam.LstMonde.List(FrmParam.LstMonde.ListIndex), FrmParam.TabStripMonde.Tabs("ChargerScenario").Selected
        End If
    End If
    FrmParam.Actualiser_Liste_Mondes
End Sub

Private Sub MenuOptionEditeurCartes_Click()
    FrmEditeurCarte.Show
End Sub

Private Sub MenuOptionEditeurIA_Click()
    FrmEditeurIA.Show
End Sub

Private Sub MenuOptionMusique_Click()
    Musique.Definir_JouerMusique = Not Musique.Definir_JouerMusique
    MenuOptionMusique.Checked = Musique.Definir_JouerMusique
    If MenuOptionMusique.Checked Then Musique.Jouer_Musique 0, False
End Sub

Private Sub MenuOptionParametres_Click()
    frmOptions.Show
End Sub

Private Sub MenuAideBatiments_click()
    'FrmInfoBatiments.Show
    'FrmInfoBatiments.ZOrder 0
    Document_Batiments
End Sub
Private Sub MenuAideBestiaire_Click()
    Document_Bestiaire
End Sub
Private Sub MenuAideObjets_Click()
    'FrmInfoObjets.Show
    'FrmInfoObjets.ZOrder 0
    Document_Objets
End Sub
Private Sub MenuAideRessources_Click()
    'FrmInfoRessources.Show
    'FrmInfoRessources.ZOrder 0
    Document_Ressources
End Sub
Private Sub MenuAidePeuples_Click()
    Document_Peuples
End Sub

Private Sub MenuAideAPropos_Click()
    FrmAPropos2.Show
End Sub

Private Sub MenuPartieCommencer_Click()
    FrmParam.CmdCommencer_Click
End Sub

Private Sub MenuPartieMonde_Click(Index As Integer)
    With FrmParam
    Set .TabStripMonde.SelectedItem = .TabStripMonde.Tabs(Index + 1)
    End With
End Sub

Public Sub MenuPartieQuitter_Click()
    If ComReseau.Connecte Then ComReseau.Session_Fermer
    Unload FrmParam
    Unload FrmChargement
    Musique.StopSon
    Me.Hide
    End
End Sub

Private Sub MenuPersonnageCharger_Click()
    FrmParam.TabStripPersonnage.Tabs(2).Selected = True
End Sub

Public Sub MenuPersonnageCompetencesAffichage_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2 'UBound(MenuPersonnageCompetencesAffichage())
        MenuPersonnageCompetencesAffichage(i).Checked = False
    Next i
    MenuPersonnageCompetencesAffichage(Index).Checked = True
    AffichageCompetences = Index
    FrmParam.LstPerso_Click
End Sub

Private Sub MenuPersonnageCompetencesAfficherToutes_Click()
    With MenuPersonnageCompetencesAfficherToutes
    .Checked = Not .Checked
    FrmParam.LstPerso_Click
    'Interface(0).CompetencesAfficherToutes = .Checked
    End With
End Sub

Public Sub MenuPersonnageCompetencesTri_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2
        MenuPersonnageCompetencesTri(i).Checked = False
    Next i
    MenuPersonnageCompetencesTri(Index).Checked = True
    TriCompetences = Index
    'Interface(0).CompetencesTri = Index
    FrmParam.LstPerso_Click
End Sub

Private Sub MenuPersonnageNouveau_Click()
    FrmParam.TabStripPersonnage.Tabs(1).Selected = True
End Sub

Private Sub MenuPersonnageExporter_Click()
    Document_Perso Persos(NoPerso)
End Sub

Private Sub MenuPersonnageSelectionMeilleur_Click(Index As Integer)
    Dim i As Long
    Dim Temp1 As Double
    Dim Temp2 As String
    MousePointer = 11
    With FrmParam.LstPerso
    If .ListCount > 0 Then
        'On cherche le meilleur.
        For i = 0 To .ListCount - 1
            Persos(NoPerso).Init NoPerso, Parametres
            Charger_Sauvegarde_Perso .List(i), Persos(NoPerso)
            Select Case Index
            Case 0:
                If Persos(NoPerso).Niveau > Temp1 Then
                    Temp1 = Persos(NoPerso).Niveau
                    Temp2 = .List(i)
                End If
            Case 1:
                If Persos(NoPerso).Definir_Moy_Competence > Temp1 Then
                    Temp1 = Persos(NoPerso).Definir_Moy_Competence
                    Temp2 = .List(i)
                End If
            Case 2:
                If Persos(NoPerso).Definir_Bestiaire_Total > Temp1 Then
                    Temp1 = Persos(NoPerso).Definir_Bestiaire_Total
                    Temp2 = .List(i)
                End If
            Case 3:
                If Persos(NoPerso).Definir_MaxVie > Temp1 Then
                    Temp1 = Persos(NoPerso).Definir_MaxVie
                    Temp2 = .List(i)
                End If
            Case 4:
                If Persos(NoPerso).Definir_MaxEnergie > Temp1 Then
                    Temp1 = Persos(NoPerso).Definir_MaxEnergie
                    Temp2 = .List(i)
                End If
            Case 5:
                If Persos(NoPerso).Definir_MaxMagie > Temp1 Then
                    Temp1 = Persos(NoPerso).Definir_MaxMagie
                    Temp2 = .List(i)
                End If
            Case 6:
                If Persos(NoPerso).Definir_MaxMoral > Temp1 Then
                    Temp1 = Persos(NoPerso).Definir_MaxMoral
                    Temp2 = .List(i)
                End If
            End Select
        Next i
        'On sélectionne le meilleur.
        For i = 0 To .ListCount - 1
            If .List(i) = Temp2 Then
                .Selected(i) = False
                .Selected(i) = True
            Else
                .Selected(i) = False
            End If
        Next i
    End If
    End With
    MousePointer = 0
End Sub

Private Sub MenuPersonnageSelectionnerTous_Click()
    Dim i As Long
    Dim Temp As Long
    'Temp = FrmParam.LblNom
    With FrmParam.LstPerso
    Temp = .ListIndex  'Persos(NoPerso).Nom
    If .ListCount > 0 Then
        FrmParam.ChargerPerso = False
        For i = 0 To .ListCount - 1
            .Selected(i) = True
        Next i
        'For i = 0 To .ListCount - 1
        '    If .List(i) = Temp Then
        '        .Selected(i) = False
        '        .Selected(i) = True
        '    End If
        'Next i
        .ListIndex = Temp
        FrmParam.ChargerPerso = True
    End If
    End With
End Sub

Private Sub MenuPersonnageSupprimer_Click()
    FrmParam.CmdSupprimerPerso_Click
End Sub


'Private Sub DirectPlay8Event_AddRemovePlayerGroup(ByVal lMsgID As Long, ByVal lPlayerID As Long, ByVal lGroupID As Long, fRejectMsg As Boolean)
'    'VB requires that we must implement *every* member of this interface
'End Sub
'
'Private Sub DirectPlay8Event_AppDesc(fRejectMsg As Boolean)
'    'VB requires that we must implement *every* member of this interface
'End Sub
'
'Private Sub DirectPlay8Event_AsyncOpComplete(dpnotify As DxVBLibA.DPNMSG_ASYNC_OP_COMPLETE, fRejectMsg As Boolean)
'    'VB requires that we must implement *every* member of this interface
'End Sub
'
'Private Sub DirectPlay8Event_ConnectComplete(dpnotify As DxVBLibA.DPNMSG_CONNECT_COMPLETE, fRejectMsg As Boolean)
'    'VB requires that we must implement *every* member of this interface
'End Sub
'
'Private Sub DirectPlay8Event_CreateGroup(ByVal lGroupID As Long, ByVal lOwnerID As Long, fRejectMsg As Boolean)
'    'VB requires that we must implement *every* member of this interface
'End Sub
'
'Private Sub DirectPlay8Event_CreatePlayer(ByVal lPlayerID As Long, fRejectMsg As Boolean)
'
'End Sub
'
'Private Sub DirectPlay8Event_DestroyGroup(ByVal lGroupID As Long, ByVal lReason As Long, fRejectMsg As Boolean)
'    'VB requires that we must implement *every* member of this interface
'End Sub
'
'Private Sub DirectPlay8Event_DestroyPlayer(ByVal lPlayerID As Long, ByVal lReason As Long, fRejectMsg As Boolean)
'    'VB requires that we must implement *every* member of this interface
'End Sub
'
'Private Sub DirectPlay8Event_EnumHostsQuery(dpnotify As DxVBLibA.DPNMSG_ENUM_HOSTS_QUERY, fRejectMsg As Boolean)
'    'VB requires that we must implement *every* member of this interface
'End Sub
'
'Private Sub DirectPlay8Event_EnumHostsResponse(dpnotify As DxVBLibA.DPNMSG_ENUM_HOSTS_RESPONSE, fRejectMsg As Boolean)
'    'VB requires that we must implement *every* member of this interface
'End Sub
'
'Private Sub DirectPlay8Event_HostMigrate(ByVal lNewHostID As Long, fRejectMsg As Boolean)
'    'VB requires that we must implement *every* member of this interface
'End Sub
'
'Private Sub DirectPlay8Event_IndicateConnect(dpnotify As DxVBLibA.DPNMSG_INDICATE_CONNECT, fRejectMsg As Boolean)
'    'VB requires that we must implement *every* member of this interface
'End Sub
'
'Private Sub DirectPlay8Event_IndicatedConnectAborted(fRejectMsg As Boolean)
'    'VB requires that we must implement *every* member of this interface
'End Sub
'
'Private Sub DirectPlay8Event_InfoNotify(ByVal lMsgID As Long, ByVal lNotifyID As Long, fRejectMsg As Boolean)
'    'VB requires that we must implement *every* member of this interface
'End Sub
'
'Private Sub DirectPlay8Event_Receive(dpnotify As DxVBLibA.DPNMSG_RECEIVE, fRejectMsg As Boolean)
'    ComReseau.Message_Lire
'    Dim lngOffset As Long
'    Dim bMessageType As Byte
'
'    Call GetDataFromBuffer(dpnotify.ReceivedData, bMessageType, SIZE_BYTE, lngOffset)
'    Select Case bMessageType
'        Case bPacketChat
'            Call ForwardChat(dpnotify, lngOffset)
'        Case bPacketBoard
'            Call ForwardBoard(dpnotify, lngOffset)
'    End Select
'End Sub
'
'Private Sub DirectPlay8Event_SendComplete(dpnotify As DxVBLibA.DPNMSG_SEND_COMPLETE, fRejectMsg As Boolean)
'    'VB requires that we must implement *every* member of this interface
'End Sub
'
'Private Sub DirectPlay8Event_TerminateSession(dpnotify As DxVBLibA.DPNMSG_TERMINATE_SESSION, fRejectMsg As Boolean)
'    'VB requires that we must implement *every* member of this interface
'End Sub
