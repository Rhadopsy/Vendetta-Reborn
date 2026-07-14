VERSION 5.00
Begin VB.Form FrmMoteur2D 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15240
   Icon            =   "FrmMoteur2D.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   741
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11520
      Left            =   0
      ScaleHeight     =   768
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1024
      TabIndex        =   0
      Top             =   0
      Width           =   15360
      Begin VB.PictureBox picTemp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   975
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox picCarte 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   0
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmMoteur2D"
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

Const AfficherChemin = True

Public NombreJoueursMO As Integer 'Nombre de joueurs sur le même ordinateur.

'Variables DirectX
Public di As DirectInput                     'creation objet DirectInput
Dim ddClipper As DirectDrawClipper           'Clipper utilisé en mode fénêtre.

Dim surfPrim As DirectDrawSurface7           'creation surface Primaire
Dim sdprime As DDSURFACEDESC2                'description surface Primaire

Dim SurfBack As DirectDrawSurface7           'creation surface Back-Buffer
Dim sdback As DDSURFACEDESC2                 'description surface Back-Buffer
Dim cback As DDSCAPS2                        '
Dim MonFond As RECT                           'Image qui recouvre l'arrière-plan.

Dim r1 As RECT
Dim r2 As RECT

'Rectangles servant au zoom du terrain.
Dim Rect1 As RECT
Dim Rect2 As RECT

'Gestion de la fréquence d'affichage par chronomètre.
Public DureePartie As Long
Dim TempsDeb As Single
Dim TempsFin As Single
Dim DeltaTemps As Single
Dim DernierFPS As Single
Dim FPS As Long
Public MaxFPS As Long
Private MessageFPS As String
Dim Vitesse As Long
Dim TempVitesse As Integer
Dim MaxVitesse As Long

'Variables de décalages en pixels nécessaires au scrolling.
Dim ScrollingAuto() As Integer
Dim DecalageX() As Long
Dim DecalageY() As Long

Dim IndicePerso As Long

Dim Key As Integer
Dim KCtrl() As Boolean
Dim KShift() As Boolean
Dim KAlt() As Boolean
Dim KHaut As Boolean
Dim KBAs As Boolean
Dim KGauche As Boolean
Dim KDroite As Boolean

Dim CliqueSouris() As Boolean

Dim TypeCurseur() As Integer
Dim IndiceCible() As Long
Dim SourisX() As Single
Dim SourisY() As Single
Dim SourisZ() As Single
Dim SourisTempX() As Single
Dim SourisTempY() As Single
Dim SourisTempZ() As Single
Dim CliqueGauche() As Boolean
Dim CliqueMilieu() As Boolean
Dim CliqueDroit() As Boolean
Dim MaintenuEnfonceGauche() As Boolean
Dim MaintenuEnfonceMilieu() As Boolean
Dim MaintenuEnfonceDroit() As Boolean
Dim BoutonRelacheGauche() As Boolean
Dim BoutonRelacheDroit() As Boolean
Dim CliqueSurTerrain() As Boolean 'Mémorise si le clique précédent a été effectué sur le terrain.
Dim CliqueSurInterface() As Boolean 'Mémorise si le clique précédent a été effectué sur l'interface.

Dim BasculeClavier As Boolean 'Si vrai, on utilise le clavier pour contrôler le curseur. Si non, c'est la souris.

Dim CurseurSurInterface() As Boolean

Public EcranDivise As Boolean
Public Zoom As Single

'Fonctions de selection des soldats à la souris.
Dim SelectionArmee() As Boolean
Dim ArmeeSelectionnee() As Boolean
Dim SoldatsSelectionnes() As Collection

'Classe sprite "image" universel
Dim ChargementImage As Boolean 'Si vrai, on affiche l'image de chargement.
Dim ImageFS As New ClsAffSprite

'Dim IndiceCadre As Integer
'Dim RetourCadre As Integer

Public Sub FatalEnd()
    'sub a appelé en cas de "vautrage" pendant que le jeu fonctionne...
    dd.RestoreDisplayMode
    dd.SetCooperativeLevel Me.hwnd, DDSCL_NORMAL
    Set dd = Nothing
    End
End Sub

Public Sub MoteurJeu()
    Dim NoBouton() As Long
    Dim NoBatiment() As Integer
    'Dim NoBatiment As Integer 'Mémorise l'indice du batiment que l'on veut fabriquer.
    Dim i As Long, j As Long, k As Long
    Dim Temp1, Temp2, Temp3 As Boolean
    
    Dim AffVictoire() As ClsAffSprite
    
    IndicePerso = 0
    DureePartie = 0
    Quitter = False
    
    DeltaTemps = Timer
    
    If FrmParam.CheckJouera2.Value = 1 And _
       FrmParam.CheckJouera2.Enabled And _
       Not Jeu.Campagne Then
        NombreJoueursMO = 2
        'Sélectionne les 2 personnages pris pour le mode 2 joueurs.
        For i = 0 To FrmParam.LstPerso.ListCount - 1
            If FrmParam.LstPerso.List(i) = FrmParam.ComboNomJoueur(0).Text Or _
               FrmParam.LstPerso.List(i) = FrmParam.ComboNomJoueur(1).Text Then
                FrmParam.LstPerso.Selected(i) = True
            End If
        Next i
    Else
        NombreJoueursMO = 1
    End If
    
    ReDim AffVictoire(NombreJoueursMO - 1)
    For i = 0 To NombreJoueursMO - 1
        Set AffVictoire(i) = New ClsAffSprite
    Next i
    
    ReDim Nopersos(NombreJoueursMO - 1)
    For i = 0 To UBound(Nopersos())
        Nopersos(i) = i
    Next i
    
    ReDim NoBouton(NombreJoueursMO - 1)
    ReDim NoBatiment(NombreJoueursMO - 1)
    
    'charge les classe ecran, souris, clavier, interface et DXdraw ...
    ScreenHeight = Ecran.Height
    ScreenWidth = Ecran.Width
    Ecran.Init dx, dd, di, Me.hwnd, surfPrim, sdprime, SurfBack, SurfBack, sdback, cback, MonFond, ddClipper
    If Ecran.Definir_ModeFenetre Then
        surfPrim.SetClipper ddClipper
        r2.Bottom = sdprime.lHeight
        r2.Right = sdprime.lHeight
    End If
    'Changer_Zoom 1
    
    If dd Is Nothing Then Exit Sub
    'Erreur_Inscrire "Ecran : Ok", True
    
    ReDim Messages(NombreJoueursMO - 1)
    ReDim Interface(NombreJoueursMO - 1)
    For i = 0 To NombreJoueursMO - 1
        Set Messages(i) = New ClsJeuMessages
        Messages(i).Init Parametres, i * Ecran.Width / 2
        
        Set Interface(i) = New ClsInterface
        Interface(i).Init1 i, NombreJoueursMO
        Interface(i).Init2 dd, Ecran, Parametres
    Next i
    
    ReDim ScrollingAuto(NombreJoueursMO - 1)
    For i = 0 To NombreJoueursMO - 1
        ScrollingAuto(i) = True
    Next i
    ReDim DecalageX(NombreJoueursMO - 1)
    ReDim DecalageY(NombreJoueursMO - 1)

    'Système d'acquisition
    ReDim KCtrl(NombreJoueursMO - 1)
    ReDim KShift(NombreJoueursMO - 1)
    ReDim KAlt(NombreJoueursMO - 1)
    
    ReDim CliqueSouris(2)
    
    ReDim TypeCurseur(NombreJoueursMO - 1)
    ReDim IndiceCible(NombreJoueursMO - 1)
    ReDim SourisX(NombreJoueursMO - 1)
    ReDim SourisY(NombreJoueursMO - 1)
    ReDim SourisZ(NombreJoueursMO - 1)
    ReDim SourisTempX(NombreJoueursMO - 1)
    ReDim SourisTempY(NombreJoueursMO - 1)
    ReDim SourisTempZ(NombreJoueursMO - 1)
    ReDim CliqueGauche(NombreJoueursMO - 1)
    ReDim CliqueMilieu(NombreJoueursMO - 1)
    ReDim CliqueDroit(NombreJoueursMO - 1)
    ReDim MaintenuEnfonceGauche(NombreJoueursMO - 1)
    ReDim MaintenuEnfonceMilieu(NombreJoueursMO - 1)
    ReDim MaintenuEnfonceDroit(NombreJoueursMO - 1)
    ReDim BoutonRelacheGauche(NombreJoueursMO - 1)
    ReDim BoutonRelacheDroit(NombreJoueursMO - 1)
    ReDim CliqueSurTerrain(NombreJoueursMO - 1)
    ReDim CliqueSurInterface(NombreJoueursMO - 1)

    Souris.Init di, dd, Me, SourisX(0), SourisY(0), Ecran
    Select Case NombreJoueursMO
    Case 1:
        Souris.Centrer_curseur SourisX(0), SourisY(0), Ecran.Width, Ecran.Height
    Case 2:
        Souris.Centrer_curseur SourisX(0), SourisY(0), Ecran.Width * 0.5, Ecran.Height
        Souris.Centrer_curseur SourisX(1), SourisY(1), Ecran.Width * 1.5, Ecran.Height
    End Select
    AffSouris.Charger dd
    ReDim SoldatsSelectionnes(NombreJoueursMO - 1)
    For i = 0 To NombreJoueursMO - 1
        Set SoldatsSelectionnes(i) = New Collection
    Next i
    ReDim SelectionArmee(NombreJoueursMO - 1)
    ReDim ArmeeSelectionnee(NombreJoueursMO - 1)
    
    ReDim CurseurSurInterface(NombreJoueursMO - 1)
    
    'Erreur_Inscrire "Souris : Ok"
    Clavier.Init di, Me
    'Erreur_Inscrire "Clavier : Ok"
    
    'Affiche le "chargement"
    Afficher_Image_Chargement dd
    Afficher_Image_Chargement dd, 15, Parametres.Etiquette_Chargement(0)
    DoEvents
    
    'génération, chargement, création, etc...
    ModMain.Commencer
    'Erreur_Inscrire "Chargement jeu : Ok"
    Afficher_Image_Chargement dd, 60, Parametres.Etiquette_Chargement(5)  ', "graphismes des personnages"
    
    'Chargement des apparences des personnages.
    AffPerso.Charger dd, Parametres
    AffApparence.Charger dd
    AffApparence.Definir_Indices_Generaux Parametres
    
    If Not ComReseau.Client Then
        For i = 0 To UBound(Persos())
            Perso_Finalier i
        Next i
        'Charge les objets uniques.
        For i = 0 To UBound(Persos())
            If Persos(i).Joueur = 0 Then
                Persos(i).Changer_Race_Personnage_Equipement Parametres, Commentaires
            End If
        Next i
    End If
    
    'Erreur_Inscrire "Chargement apparences : Ok"
    For i = 0 To UBound(Persos())
        'Persos(i).Calculer_Apparence Maisons(i), Parametres
        Persos(i).Actualiser_Bonus
    Next i
    For i = 0 To UBound(Fiefs())
        Fiefs(i).Marches_Actualiser
    Next i
    
    Afficher_Image_Chargement dd, 75, Parametres.Etiquette_Chargement(6)  ', "Graphismes des décors 1/2"
    
    'Autres chargement de graphismes.
    Set AffTerrain = New ClsAffTerrain
    AffTerrain.Charger dd, "Plaine": AffTerrain.Largeur = Monde.LargeurCase: AffTerrain.Hauteur = Monde.HauteurCase
    If Monde.TerrainOptimiseDepart Then AffTerrain.Modeliser_Carte dd, Monde
    AffChateau.Charger 0, dd, Parametres
    Afficher_Image_Chargement dd, 85, Parametres.Etiquette_Chargement(9) ', "Graphismes des décors 2/2"
    AffMaison.Charger 1, dd, Parametres
    AffRessources.Charger dd
    AffEffet.Charger dd, Parametres
    AffTresor.Charger dd
    AffDecor.Charger dd, Parametres
    AffTemps.Charger dd
    
    For i = 0 To NombreJoueursMO - 1
        AffVictoire(i).Charger dd, "Fenetres\VictoireDefaite" & Langues.Suffixe
    Next i
    Afficher_Image_Chargement dd, 95, Parametres.Etiquette_Chargement(10) ', "Fichiers de sons"
    
    'Chargement des sons.
    Son.Init Me, dx, Parametres
    
    Afficher_Image_Chargement dd, 100 ', "Chargement terminé"
    
    'Démarre la partie réseau.
    'If ComReseau.Serveur Then
    '    ComReseau.Message_Envoyer Noperso, "Com"
    'End If
    'If ComReseau.Connecte Then
    '    ComReseau.PartieEnCours = True
    'End If
    
    TempsDeb = Timer
    TempsFin = TempsDeb
    
    'Scrolling_Immediat Persos(Noperso).PositionX, Persos(Noperso).PositionY
    Zoom = 1
    
    For i = 0 To NombreJoueursMO - 1
        'Chargement de l'interface.
        Charger_Sauvegarde_Interface Persos(Nopersos(i)), Interface(i)
        'Active le bouton batiment dès le début.
        Interface(i).Actionner_Boutons_Infos 0, i
        CliqueSurInterface(i) = True
        Messages(i).Afficher_Conseil
        
        If ComReseau.Client And ComReseau.Connecte Then
            ComMessages.Infos_Perso_Caracs Nopersos(i), ComReseau.PlayerID
            ComMessages.Infos_Maison Persos(Nopersos(i)).IndiceMaison, ComReseau.PlayerID
            ComMessages.Infos_Chateau Persos(Nopersos(i)).IndiceChateau, ComReseau.PlayerID
            ComMessages.Infos_Perso_Caracs2 Nopersos(i), ComReseau.PlayerID
            ComMessages.Infos_Perso_Competences2 Nopersos(i), ComReseau.PlayerID
            ComMessages.Infos_Perso_InventaireMaison Nopersos(i), ComReseau.PlayerID
            'ComMessages.Infos_Perso_Bestiaire Nopersos(i), ComReseau.PlayerID
        End If
    Next i
    
    IA_Init
    
    Musique.Jouer_Musique 1
    
    ReDim Victoire(NombreJoueursMO - 1)
    ReDim Defaite(NombreJoueursMO - 1)
    For i = 0 To NombreJoueursMO - 1
        Victoire(i) = False
        Defaite(i) = False
    Next i
    
    'UniversalDebugger.Debugger_Init "_debug.bin"
    'UniversalDebugger.Debugger_SetStamp 6
    
    If Not Quitter Then Stastistiques.Init Parametres
    Evenements_Lancer "Commencer"
    
    Partie.EnCours = 2
    
    'Do While DoEvents And Not Quitter
    Do While Not Quitter
        
        'Musique d'arrière plan.
        If Persos(NoPerso).EnCombat Then
            If Persos(NoPerso).Action = 302 Then
                Musique.Jouer_Musique 7, Victoire(0) Or Defaite(0)
            Else
                Musique.Jouer_Musique 3, Victoire(0) Or Defaite(0)
            End If
        ElseIf Not Persos(NoPerso).Vivant Then
            Musique.Jouer_Musique 4, Victoire(0) Or Defaite(0)
        ElseIf Persos(NoPerso).Empereur Then
            Musique.Jouer_Musique 8, True
        Else
            Musique.Jouer_Musique 2, True
        End If
        
        TempsFin = Timer
        DeltaTemps = TempsFin - TempsDeb
        
        'Vérification des ordres du clavier.
        'If Souris.Actif Then 'Si la souris est désactivée, pas besoin d'interroger le clavier.
        If Souris.Actif Then
            If Interface(0).AfficherDialogue Then
                Key = Clavier.Pression_Clavier(KCtrl(0), KShift(0), KAlt(0), KHaut, KBAs, KGauche, KDroite)
                If Key <> IntTouches.ToucheDialogue Then
                    Key = Clavier.Pression_ClavierDialogue(KCtrl(0), KShift(0), KAlt(0), KHaut, KBAs, KGauche, KDroite)
                    If Key = IntTouches.ToucheDialogue Then Key = 0
                End If
            Else
                Key = Clavier.Pression_Clavier(KCtrl(0), KShift(0), KAlt(0), KHaut, KBAs, KGauche, KDroite)
            End If
        Else
            Key = 0
            'Souris.Init di, dd, Me, SourisX(0), SourisY(0), Ecran
        End If
        
        If NombreJoueursMO = 2 Then
            If Jeu.Definir_Mode2JoueursEcranSepares Then
                EcranDivise = True
            Else
                If EcranDivise Then
                    If Abs(Persos(Nopersos(0)).PositionX - Persos(Nopersos(1)).PositionX) < _
                       ScreenWidth * 0.5 And _
                       Abs(Persos(Nopersos(0)).PositionY - Persos(Nopersos(1)).PositionY) < _
                       ScreenHeight * 0.5 Then
                        EcranDivise = False
                    End If
                Else
                    If Abs(Persos(Nopersos(0)).PositionX - Persos(Nopersos(1)).PositionX) > _
                       ScreenWidth * 0.7 Or _
                       Abs(Persos(Nopersos(0)).PositionY - Persos(Nopersos(1)).PositionY) > _
                       ScreenHeight * 0.7 Then
                        EcranDivise = True
                    End If
                End If
            End If
        End If
        
        For i = 0 To NombreJoueursMO - 1
            If i = 0 Or Not EcranDivise Then
                ScreenPositionX = 0
            Else
                ScreenPositionX = Ecran.Width / 2
            End If
            NoPerso = Nopersos(i)
            'Debugger_TimeStamp True, 6
                        
            'Sauvegarde du personnage.
            Enregistrement_Periodique Persos(NoPerso), TempsFin
            
            'Debug.Print Format((Timer - TempsFin) * 1000, "00") & " ms" & " : Début"
            
            'Debugger_TimeStamp True, 1
            
            'If Souris.Bouton_Relache(0) Then _
                Messages.Ajouter_Message "Souris.Bouton_Relache(0)"
            'If Souris.Maintenu_Enfonce(0) Then _
                Messages.Ajouter_Message "Souris.Maintenu_Enfonce(0)"
            'Messages.Ajouter_Message Liste_Soldats_Selectionnes
            'Vérification des ordres de la souris.
            If NombreJoueursMO > 1 Then
                If i = 0 Then
                    If EcranDivise Then
                        Clavier.Positions SourisX(i), SourisY(i), SourisZ(i), 0, Ecran.Width * 0.5, 0, Ecran.Height
                    Else
                        Clavier.Positions SourisX(i), SourisY(i), SourisZ(i), 0, Ecran.Width, 0, Ecran.Height
                    End If
                    CliqueGauche(i) = Clavier.Clique(0)
                    CliqueMilieu(i) = Clavier.Clique(2)
                    CliqueDroit(i) = Clavier.Clique(1)
                    MaintenuEnfonceGauche(i) = Clavier.Maintenu_Enfonce(0)
                    MaintenuEnfonceMilieu(i) = Clavier.Maintenu_Enfonce(2)
                    MaintenuEnfonceDroit(i) = Clavier.Maintenu_Enfonce(1)
                    BoutonRelacheGauche(i) = Clavier.Bouton_Relache(0)
                    BoutonRelacheDroit(i) = Clavier.Bouton_Relache(1)
                Else
                    If Not Ecran.Definir_ModeFenetre Then
                        If EcranDivise Then
                            Souris.Positions di, dd, Me, SourisX(i), SourisY(i), SourisZ(i), Ecran.Width * 0.5, Ecran.Width, 0, Ecran.Height
                        Else
                            Souris.Positions di, dd, Me, SourisX(i), SourisY(i), SourisZ(i), 0, Ecran.Width, 0, Ecran.Height
                        End If
                        CliqueGauche(i) = Souris.Clique(0)
                        CliqueMilieu(i) = Souris.Clique(2)
                        CliqueDroit(i) = Souris.Clique(1)
                    Else
                        CliqueGauche(i) = Souris.Definir_Clique(0, CliqueSouris(0))
                        CliqueMilieu(i) = Souris.Definir_Clique(2, CliqueSouris(2))
                        CliqueDroit(i) = Souris.Definir_Clique(1, CliqueSouris(1))
                    End If
                    MaintenuEnfonceGauche(i) = Souris.Maintenu_Enfonce(0)
                    MaintenuEnfonceMilieu(i) = Souris.Maintenu_Enfonce(2)
                    MaintenuEnfonceDroit(i) = Souris.Maintenu_Enfonce(1)
                    BoutonRelacheGauche(i) = Souris.Bouton_Relache(0)
                    BoutonRelacheDroit(i) = Souris.Bouton_Relache(1)
                End If
                'Messages(i).Ajouter_Message "Clique : " & CliqueGauche(i) & " | " & _
                                            "Maintenu : " & MaintenuEnfonceGauche(i) & " | " & _
                                            "Relaché : " & BoutonRelacheGauche(i)
            Else
                'Commandes clavier.
                Clavier.Positions SourisX(i), SourisY(i), SourisZ(i), 0, Ecran.Width, 0, Ecran.Height
                Temp1 = Clavier.Clique(0)
                Temp2 = Clavier.Clique(2)
                Temp3 = Clavier.Clique(1)
                If Temp1 Or Temp2 Or Temp3 Then BasculeClavier = True
                If BasculeClavier Then
                    CliqueGauche(i) = Temp1
                    CliqueMilieu(i) = Temp2
                    CliqueDroit(i) = Temp3
                    MaintenuEnfonceGauche(i) = Clavier.Maintenu_Enfonce(0)
                    MaintenuEnfonceMilieu(i) = Clavier.Maintenu_Enfonce(2)
                    MaintenuEnfonceDroit(i) = Clavier.Maintenu_Enfonce(1)
                    BoutonRelacheGauche(i) = Clavier.Bouton_Relache(0)
                    BoutonRelacheDroit(i) = Clavier.Bouton_Relache(1)
                End If
            
                If Not Ecran.Definir_ModeFenetre Then
                    Souris.Positions di, dd, Me, SourisX(i), SourisY(i), SourisZ(i), 0, Ecran.Width, 0, Ecran.Height
                    Temp1 = Souris.Clique(0)
                    Temp2 = Souris.Clique(2)
                    Temp3 = Souris.Clique(1)
                    If Temp1 Or Temp2 Or Temp3 Then BasculeClavier = False
                    If Not BasculeClavier Then
                        CliqueGauche(i) = Temp1
                        CliqueMilieu(i) = Temp2
                        CliqueDroit(i) = Temp3
                    End If
                Else
                    Temp1 = Souris.Definir_Clique(0, CliqueSouris(0))
                    Temp2 = Souris.Definir_Clique(2, CliqueSouris(2))
                    Temp3 = Souris.Definir_Clique(1, CliqueSouris(1))
                    If Temp1 Or Temp2 Or Temp3 Then BasculeClavier = False
                    If Not BasculeClavier Then
                        CliqueGauche(i) = Temp1
                        CliqueMilieu(i) = Temp2
                        CliqueDroit(i) = Temp3
                    End If
                End If
                If Not BasculeClavier Then
                    MaintenuEnfonceGauche(i) = Souris.Maintenu_Enfonce(0)
                    MaintenuEnfonceMilieu(i) = Souris.Maintenu_Enfonce(2)
                    MaintenuEnfonceDroit(i) = Souris.Maintenu_Enfonce(1)
                    BoutonRelacheGauche(i) = Souris.Bouton_Relache(0)
                    BoutonRelacheDroit(i) = Souris.Bouton_Relache(1)
                End If
                'Messages(0).Ajouter_Message BasculeClavier & ", " & CliqueGauche(0)
            End If
            
            'If CliqueMilieu(i) Then
                'Le bouton du milieu ferme toutes les fenêtres.
                'Interface(i).Fermer_Toutes_les_fenetres
                'Messages(i).Effacer
            'End If
            
            If Interface(i).DeplacerFenetre >= 0 Then
                If (CliqueGauche(i) Or CliqueDroit(i)) Then
                    'on repositionne une fenêtre.
                    If CliqueGauche(i) Then
                        Interface(i).Positionner_Fenetre SourisX(i), SourisY(i)
                    Else
                        Interface(i).DeplacerFenetre = -1
                    End If
                    CliqueGauche(i) = False
                    CliqueDroit(i) = False
                End If
                'Efface le rectangle de sélection.
                SourisTempX(i) = SourisX(i)
                SourisTempY(i) = SourisY(i)
            Else
                If Persos(NoPerso).Vivant And Persos(NoPerso).NombreSoldats > 0 Then
                    If MaintenuEnfonceGauche(i) Or MaintenuEnfonceDroit(i) Then
                        SelectionArmee(i) = True
                        CliqueSurTerrain(i) = False
                    Else
                        If SelectionArmee(i) Then
                            Selectionner_Soldats i, KShift(i), KCtrl(i)
                        Else
                            If CliqueSurInterface(i) Then
                                CliqueSurInterface(i) = False
                            Else
                                CliqueSurTerrain(i) = True
                            End If
                        End If
                        'Efface le rectangle de sélection.
                        SourisTempX(i) = SourisX(i)
                        SourisTempY(i) = SourisY(i)
                    End If
                Else
                    ArmeeSelectionnee(i) = False
                    Ordonner_Activer_IA NoPerso
                    'Efface le rectangle de sélection.
                    SourisTempX(i) = SourisX(i)
                    SourisTempY(i) = SourisY(i)
                End If
                With Interface(i)
                If .ActiverBoutonsRadiaux And SourisZ(i) <> 0 Then
                    .Pivoter_BoutonsRadiaux SourisZ(i)
                ElseIf Not .CurseurSurBarreStatut Then
                    If SourisZ(i) > 0 Then
                        If .Fenetres_Visible(0) Or .Fenetres_Visible(1) Then
                            If .Fenetres_Visible(0) Then
                                .Defilement_Haut SourisZ(i), 0
                                If Persos(NoPerso).DansSaMaison Or Persos(NoPerso).DansUneMaison And Maisons(Persos(NoPerso).IndiceMaison).Fabrique Then
                                    .Defilement_Haut SourisZ(i), 0, True
                                End If
                            Else
                                .Defilement_Haut SourisZ(i), 1
                            End If
                        End If
                        If .Fenetres_Visible(3) Then
                            .Defilement_Haut SourisZ(i), 3
                        End If
                        If .Fenetres_Visible(5) Then
                            .Defilement_Haut SourisZ(i), 5
                        End If
                    ElseIf SourisZ(i) < 0 Then
                        If .Fenetres_Visible(0) Or .Fenetres_Visible(1) Then
                            If Persos(NoPerso).DansSaMaison Or Persos(NoPerso).DansUneMaison And Maisons(Persos(NoPerso).IndiceMaison).Fabrique Then
                                .Defilement_Bas -SourisZ(i), 0, True
                            End If
                            If .Fenetres_Visible(1) Then
                                .Defilement_Bas -SourisZ(i), 1
                            Else
                                .Defilement_Bas -SourisZ(i), 0
                            End If
                        End If
                        If .Fenetres_Visible(3) Then
                            .Defilement_Bas -SourisZ(i), 3
                        End If
                        If .Fenetres_Visible(5) Then
                            .Defilement_Bas -SourisZ(i), 5
                        End If
                    End If
                End If
                End With
                
                If CliqueMilieu(i) Then
                    'Le bouton du milieu sert à attaquer dans le vide.
                    Cmd_Perso_AttaquerTerrain NoPerso, SourisX(i) - ScreenPositionX - Persos(NoPerso).Largeur / 2 + DecalageX(i), _
                                              SourisY(i) - ScreenPositionY - Persos(NoPerso).Hauteur / 2 + DecalageY(i), IIf(TypeCurseur(i) = 10, 1, IIf(TypeCurseur(i) = 11, 2, 0))
                End If
                If Not CurseurSurInterface(i) And ArmeeSelectionnee(i) And (TypeCurseur(i) = 0 Or TypeCurseur(i) > 3 And TypeCurseur(i) < 12 Or TypeCurseur(i) = 16 Or TypeCurseur(i) = 17) Then
                    'Ordres données aux armées.
                    If CliqueSurTerrain(i) And BoutonRelacheDroit(i) Then
                        If Not Selectionner_Soldats(i, KShift(i), KCtrl(i)) Then
                            ArmeeSelectionnee(i) = False
                            CliqueSurInterface(i) = True
                            CliqueSurTerrain(i) = False
                            Ordonner_Activer_IA NoPerso
                        End If
                    ElseIf CliqueSurTerrain(i) And BoutonRelacheGauche(i) Then
                        Select Case TypeCurseur(i)
                        Case 0, 16:
                            If Persos(NoPerso).Definir_ObjetSelectionne Then
                                If Persos(NoPerso).DansSonChateau Or Persos(NoPerso).DansUnChateau Then
                                    Cmd_Perso_VendreObjetsSelectionnees NoPerso, Persos(NoPerso).IndiceChateau
                                    'Fiefs(Persos(Noperso).IndiceChateau).Vendre_ObjetsSelectionnees Persos(Noperso), Parametres
                                ElseIf Persos(NoPerso).ObjetSelectionne >= 0 Then
                                    If Persos(NoPerso).DansSaMaison Then
                                        Cmd_Perso_ObjetMaisonDeposer NoPerso, Persos(NoPerso).ObjetSelectionne, False
                                        'Persos(Noperso).Objet_Deposer_Maison Persos(Noperso).ObjetSelectionne, Maisons(Noperso), Parametres, False
                                    Else
                                        Cmd_PersoObjetDeposerSol NoPerso, Persos(NoPerso).ObjetSelectionne, False
                                        'Persos(Noperso).Objet_Deposer_Sol Persos(Noperso).ObjetSelectionne, False
                                    End If
                                    Cmd_Perso_DefinirObjetSelectionne NoPerso, -1
                                    'Persos(Noperso).ObjetSelectionne = -1
                                Else
                                    Cmd_PersoRessourcesDeposerSol NoPerso
                                    'Persos(Noperso).Ressources_Deposer_Sol
                                End If
                            Else
                                Ordonner_Deplacement Persos(NoPerso), _
                                                     SoldatsSelectionnes(i), _
                                                     SourisX(i) - ScreenPositionX - Persos(NoPerso).Largeur / 2 + DecalageX(i), _
                                                     SourisY(i) - ScreenPositionY - Persos(NoPerso).Hauteur / 2 + DecalageY(i)
                            End If
                        'Case 4, 5: 'On protège sa maison.
                        '    Ordonner_Garder_Maison Persos(Noperso), SoldatsSelectionnes(i), IndiceCible(i)
                        'Case 6: 'On protège son chateau.
                        '    Ordonner_Garder_Chateau Persos(Noperso), SoldatsSelectionnes(i), Persos(Noperso).NumeroFief
                        Case 4:
                            If IndiceCible(i) = NoPerso Then
                                'On rentre à la maison.
                                Cmd_Perso_Rentrer_Maison NoPerso, False
                            Else
                                'On répare une autre maison.
                                Cmd_Perso_AllerMaison NoPerso, IndiceCible(i), True
                            End If
                        Case 5:
                            'On rentre à la maisons sans réparer.
                            Cmd_Perso_Rentrer_Maison NoPerso, True
                        Case 6:
                            'On rentre au chateau.
                            'Persos(Noperso).Rentrer_Chateau Chateaux(Persos(Noperso).NumeroFief)
                            If Persos(NoPerso).Definir_ObjetSelectionne Then
                                'Persos(Noperso).Aller_Vendre Persos(Noperso).NumeroFief
                                Cmd_Perso_AllerVendre NoPerso, Persos(NoPerso).NumeroFief, False
                            Else
                                Cmd_Perso_Rentrer_Chateau NoPerso
                            End If
                        'Case 7: 'On protège une maison
                        '    Ordonner_Garder_Maison Persos(Noperso), SoldatsSelectionnes(i), IndiceCible(i)
                        Case 7:
                            'Persos(Noperso).Aller_Maison Maisons(IndiceCible(i))
                            Cmd_Perso_AllerMaison NoPerso, IndiceCible(i), False
                        'Case 8: 'On protège un chateau.
                        '    Ordonner_Garder_Chateau Persos(Noperso), SoldatsSelectionnes(i), IndiceCible(i)
                        Case 8:
                            If Persos(NoPerso).Definir_ObjetSelectionne Then
                                'Persos(Noperso).Aller_Vendre IndiceCible(i)
                                Cmd_Perso_AllerVendre NoPerso, IndiceCible(i), False
                            Else
                                'Persos(Noperso).Aller_Chateau Chateaux(IndiceCible(i)), IndiceCible(i)
                                Cmd_Perso_AllerChateau NoPerso, IndiceCible(i)
                            End If
                        Case 9:
                            Ordonner_Attaquer_Perso Persos(NoPerso), _
                                                    SoldatsSelectionnes(i), _
                                                    IndiceCible(i)
                        Case 10:
                            Ordonner_Attaquer_Maison Persos(NoPerso), _
                                                     SoldatsSelectionnes(i), _
                                                     IndiceCible(i)
                        Case 11:
                            Ordonner_Attaquer_Chateau Persos(NoPerso), _
                                                      SoldatsSelectionnes(i), _
                                                      IndiceCible(i)
                        Case 17:
                            'Ordonne à ses soldats de protéger le personnage en question.
                            Ordonner_Garder_Perso Persos(NoPerso), _
                                                  SoldatsSelectionnes(i), _
                                                  IndiceCible(i)
                            ArmeeSelectionnee(i) = False
                            CliqueSurInterface(i) = True
                            CliqueSurTerrain(i) = False
                            Ordonner_Activer_IA NoPerso
                        End Select
                        'If TypeCurseur(i) > 3 And TypeCurseur(i) < 9 Then
                            'Déselectionne la souris après avoir été affecté à la protection d'un bâtiment.
                        '    ArmeeSelectionnee(i) = False
                        '    CliqueSurInterface(i) = True
                        '    CliqueSurTerrain(i) = False
                        '    Ordonner_Activer_IA Noperso
                        'End If
                    End If
                Else
                    If CliqueGauche(i) Or CliqueDroit(i) Or Key Or BoutonRelacheGauche(i) Or BoutonRelacheDroit(i) Or SelectionArmee(i) Then
                    'If CliqueGauche Or CliqueDroit Or Key Then
                        'On clique sur l'interface.
                        If CurseurSurInterface(i) Then
                            If NoBouton(i) > 0 Then
                                If NoBouton(i) < 100 Then
                                    TypeCurseur(i) = 2
                                ElseIf BoutonRelacheGauche(i) Then
                                    Select Case NoBouton(i)
                                    Case 101: 'Persos(Noperso).Rentrer_Maison Maisons(Noperso)
                                        Cmd_Perso_Rentrer_Maison NoPerso, KCtrl(i) And Maisons(NoPerso).Construit
                                    Case 200: 'Persos(Noperso).Rentrer_Chateau Chateaux(Persos(Noperso).NumeroFief)
                                        Cmd_Perso_Rentrer_Chateau NoPerso
                                    'Case Else:
                                        'Persos(NoPerso).Action = NoBouton + 2
                                    End Select
                                    TypeCurseur(i) = 0
                                End If
                            End If
                            CliqueSurInterface(i) = True
                            CliqueSurTerrain(i) = False
                            'Souris.Effacer_Relachement
                        'ElseIf CliqueSurTerrain Then
                        'ElseIf CliqueGauche Then
                        ElseIf BoutonRelacheGauche(i) And _
                               CliqueSurTerrain(i) Then
                            'On clique sur le plateau.
                            'On pose un batiment.
                            If TypeCurseur(i) = 2 Then
                                Cmd_Perso_Creer_Maison NoPerso, _
                                                       SourisX(i) - ScreenPositionX - Parametres.Batiment_Largeur(NoBatiment(i) - 1) / 2 + AffSouris.Largeur / 2 + DecalageX(i), _
                                                       SourisY(i) - ScreenPositionY - Parametres.Batiment_Hauteur(NoBatiment(i) - 1) / 2 + AffSouris.Hauteur / 2 + DecalageY(i), _
                                                       NoBatiment(i) - 1
                                If Not (KCtrl(i) Or KShift(i)) Then TypeCurseur(i) = 4
                                NoBouton(i) = 0
                            'On se déplace.
                            ElseIf TypeCurseur(i) <> 3 Then
                                If TypeCurseur(i) = 1 Then
                                    'Va chercher une ressource.
                                    'Persos(Noperso).Aller_Ressource Ressources(IndiceObstacle), IndiceObstacle, AffRessources.Largeur, AffRessources.Hauteur
                                    Cmd_Perso_Chercher_Ressource NoPerso, IndiceCible(i)
                                Else
                                    Select Case TypeCurseur(i)
                                    Case 0, 16:
                                        'On se déplace.
                                        If Persos(NoPerso).Definir_ObjetSelectionne And Persos(NoPerso).Vivant Then
                                            If Persos(NoPerso).DansSonChateau Or Persos(NoPerso).DansUnChateau Or (Persos(NoPerso).DansUneMaison And Maisons(Persos(NoPerso).IndiceMaison).Marche And Maisons(Persos(NoPerso).IndiceMaison).Construit) Then
                                                Cmd_Perso_VendreObjetsSelectionnees NoPerso, Persos(NoPerso).IndiceChateau
                                                'Fiefs(Persos(Noperso).IndiceChateau).Vendre_ObjetsSelectionnees Persos(Noperso), Parametres
                                            ElseIf Persos(NoPerso).ObjetSelectionne >= 0 Then
                                                If Not Parametres.Objet_Inseparable(Persos(NoPerso).ObjetSelectionne) Then
                                                    If Persos(NoPerso).DansSaMaison Then
                                                        'Persos(Noperso).Objet_Deposer_Maison Persos(Noperso).ObjetSelectionne, Maisons(Noperso), Parametres, False
                                                        Cmd_Perso_ObjetMaisonDeposer NoPerso, Persos(NoPerso).ObjetSelectionne, False
                                                    Else
                                                        'Persos(Noperso).Objet_Deposer_Sol Persos(Noperso).ObjetSelectionne, False
                                                        Cmd_PersoObjetDeposerSol NoPerso, Persos(NoPerso).ObjetSelectionne, False
                                                    End If
                                                    'Persos(Noperso).ObjetSelectionne = -1
                                                    Cmd_Perso_DefinirObjetSelectionne NoPerso, -1
                                                Else
                                                    Cmd_PersoRessourcesDeposerSol NoPerso
                                                End If
                                            Else
                                                'Persos(Noperso).Ressources_Deposer_Sol
                                                Cmd_PersoRessourcesDeposerSol NoPerso
                                            End If
                                        Else
                                            Cmd_Perso_Deplacer NoPerso, SourisX(i) - ScreenPositionX - Persos(NoPerso).Largeur / 2 + DecalageX(i), _
                                                                        SourisY(i) - ScreenPositionY - Persos(NoPerso).Hauteur / 2 + DecalageY(i)
                                        End If
                                    Case 4:
                                        If IndiceCible(i) = NoPerso Then
                                            'On rentre à la maison.
                                            Cmd_Perso_Rentrer_Maison NoPerso, False
                                        Else
                                            'On répare une autre maison.
                                            Cmd_Perso_AllerMaison NoPerso, IndiceCible(i), True
                                        End If
                                    Case 5:
                                        'On rentre à la maisons sans réparer.
                                        Cmd_Perso_Rentrer_Maison NoPerso, True
                                    Case 6:
                                        'On rentre au chateau.
                                        'Persos(Noperso).Rentrer_Chateau Chateaux(Persos(Noperso).NumeroFief)
                                        If Persos(NoPerso).Definir_ObjetSelectionne Then
                                            'Persos(Noperso).Aller_Vendre Persos(Noperso).NumeroFief
                                            If Maisons(IndiceCible(i)).Marche And Maisons(IndiceCible(i)).Construit Then
                                                Cmd_Perso_AllerVendre NoPerso, IndiceCible(i), True
                                            Else
                                                Cmd_Perso_AllerVendre NoPerso, Persos(NoPerso).NumeroFief, False
                                            End If
                                        Else
                                            Cmd_Perso_Rentrer_Chateau NoPerso
                                        End If
                                    Case 7:
                                        'Persos(Noperso).Aller_Maison Maisons(IndiceCible(i))
                                        Cmd_Perso_AllerMaison NoPerso, IndiceCible(i), False
                                    Case 8:
                                        If Persos(NoPerso).Definir_ObjetSelectionne Then
                                            'Persos(Noperso).Aller_Vendre IndiceCible(i)
                                            Cmd_Perso_AllerVendre NoPerso, IndiceCible(i), False
                                        Else
                                            'Persos(Noperso).Aller_Chateau Chateaux(IndiceCible(i)), IndiceCible(i)
                                            Cmd_Perso_AllerChateau NoPerso, IndiceCible(i)
                                        End If
                                    Case 9:
                                        'Persos(Noperso).Attaquer_Perso IndiceCible(i)
                                        Cmd_Perso_AttaquerPerso NoPerso, IndiceCible(i)
                                    Case 10:
                                        'Persos(Noperso).Attaquer_Maison IndiceCible(i)
                                        Cmd_Perso_AttaquerMaison NoPerso, IndiceCible(i)
                                    Case 11:
                                        'Persos(Noperso).Attaquer_Chateau IndiceCible(i)
                                        Cmd_Perso_AttaquerChateau NoPerso, IndiceCible(i)
                                    Case 12: 'Cas d'un temple.
                                        'Persos(Noperso).Aller_Maison Maisons(IndiceCible(i))
                                        Cmd_Perso_AllerMaison NoPerso, IndiceCible(i), False
                                    Case 13:
                                        'Persos(Noperso).Aller_Tresor Tresors(IndiceCible(i)), IndiceCible(i), AffTresor.Largeur, AffTresor.Hauteur
                                        Cmd_Perso_AllerTresor NoPerso, IndiceCible(i)
                                    Case 14:
                                        'Persos(Noperso).Aller_Donner IndiceCible(i)
                                        Cmd_Perso_AllerDonner NoPerso, IndiceCible(i)
                                    End Select
                                End If
                            End If
                            CliqueSurTerrain(i) = False
                        ElseIf CliqueGauche(i) And Not MaintenuEnfonceGauche(i) Then
                            'Souris.Effacer_Relachement
                            CliqueSurTerrain(i) = True
                        'ElseIf CliqueDroit Then
                        ElseIf BoutonRelacheDroit(i) And _
                               CliqueSurTerrain(i) Then
                            'Clique droit.
                            If Persos(NoPerso).Vivant Then
                                Selectionner_Soldats i, KShift(i), KCtrl(i)
                            End If
                            If Not ArmeeSelectionnee(i) Then
                                'If Interface(i).AttaquerAmi Then
                                '    Interface(i).AttaquerAmi = False
                                'Else
                                If Persos(NoPerso).Definir_ObjetSelectionne Then
                                    'Déselectionne les objets sélectionnés.
                                    'Persos(Noperso).Definir_ObjetSelectionne = False
                                    Cmd_Perso_ObjetDeSelectionne NoPerso
                                    If Persos(NoPerso).Action = 206 Or Persos(NoPerso).Action = 207 Then
                                        Cmd_Perso_Arreter NoPerso
                                    End If
                                Else
                                    Select Case TypeCurseur(i)
                                    Case 0:
                                        'Interface(i).Actionner_Boutons_Infos 4,i
                                        Interface(i).Activer_BoutonsRadiaux SourisX(i), SourisY(i), 4
                                        'Interface(i).Activer_BoutonsRadiaux Persos(Noperso).PositionX + Persos(Noperso).Largeur / 2, Persos(Noperso).PositionY + Persos(Noperso).Hauteur / 2, 4
                                    Case 1:
                                        'Interface(i).Actionner_Boutons_Infos 1,i
                                        Interface(i).Activer_BoutonsRadiaux SourisX(i), SourisY(i), 1
                                    Case 2, 3: 'Annule l'opération de construction.
                                        TypeCurseur(i) = 0
                                    Case 4 To 8:
                                        'Interface(i).Actionner_Boutons_Infos 0,i
                                        Interface(i).Activer_BoutonsRadiaux SourisX(i), SourisY(i), 0
                                    Case 9 To 11:
                                        'Ordonner_Chargez Persos(Noperso).Numero, TypeCurseur(i) - 9, IndiceCible(i), KCtrl(i), KShift(i), KAlt(i)
                                        Cmd_PersoOrdonnerChargez Persos(NoPerso).Numero, TypeCurseur(i) - 9, IndiceCible(i), KCtrl(i), KShift(i), KAlt(i)
                                    Case 12:
                                        'Interface(i).Actionner_Boutons_Infos 0,i
                                        Interface(i).Activer_BoutonsRadiaux SourisX(i), SourisY(i), 0
                                    End Select
                                End If
                            End If
                        ElseIf CliqueDroit(i) And Not MaintenuEnfonceDroit(i) Then
                            CliqueSurTerrain(i) = True
                        End If
                    End If
                End If
            End If
            'Debugger_TimeStamp False, 1
            'Debugger_TimeStamp True, 2
            
            'Debugger_TimeStamp False, 2
        
            'debug.Print Format((Timer - TempsFin) * 1000, "00") & " ms" & " : Souris"
            Interface(i).ActualiserRapide Persos(NoPerso), ArmeeSelectionnee(i)
        Next i
        
        'Vérification des ordres du clavier.
        IntTouches.Presser_Touche Key, KHaut, KBAs, KGauche, KDroite, Persos(Nopersos(0)), Maisons(Nopersos(0)), Persos(Nopersos(0)).NumeroFief, Persos(Nopersos(0)).NumeroFief, Interface(0), Messages(0), Nopersos(0), UBound(Persos()) + 1, Quitter, KCtrl(0), KShift(0), KAlt(0), ArmeeSelectionnee(0), SoldatsSelectionnes(0), 0, ScreenPositionX, ScreenPositionY, IIf(NombreJoueursMO = 1, Ecran.Width, Ecran.Width * 0.5), Ecran.Height
                   
        'If (DeltaTemps > 1 / (Jeu.Definir_Vitesse * 2) Or _
        '   DeltaTemps < -(1 / (Jeu.Definir_Vitesse * 2))) Then
        '    For i = 0 To NombreJoueursMO - 1
        '    Next i
        'End If
                   
        If (DeltaTemps > 1 / Jeu.Definir_Vitesse Or _
           DeltaTemps < -(1 / Jeu.Definir_Vitesse)) And _
           Not Jeu.EnPause Then  'Raffraichissement.
            DureePartie = DureePartie + 1
            TempsDeb = TempsFin
            
            'Vérification des ordres réseaux.
            With ComReseau
            If .Connecte Then
                '.Tour_Fin 'Indique qu'il a terminé son tour.
                'Do
                    'Récupère tous les messages
                    .Message_Lire
                '    DoEvents
                'Loop Until ComReseau.Tour_Fini
                '.Tour_Suivant
                '.Tour_Debut 'Entame le tour suivant.
                If .Serveur Then
                    Lire_Commandes 'Lit les instructions données par les joueurs.
                    ComMessages.Actualiser_Memoire
                    ComMessages.Monde_EnCours , True
                Else
                    'Demande les intructions au serveur.
                    For i = 0 To NombreJoueursMO - 1
                        'Informations constantes.
                        ComMessages.Infos_Perso_Caracs Nopersos(i), ComReseau.PlayerID
                        'ComMessages.Infos_Perso_Competences Nopersos(i), ComReseau.PlayerID
                        'ComMessages.Infos_Perso_Ressources Nopersos(i), ComReseau.PlayerID
                        'ComMessages.Infos_Perso_Inventaire Nopersos(i), ComReseau.PlayerID
                        'ComMessages.Infos_Perso_Bestiaire Nopersos(i), ComReseau.PlayerID
                        'Fenetre bâtiment maison.
                        If Interface(i).Fenetres_Visible(0) And (Persos(Nopersos(i)).DansSaMaison Or Persos(Nopersos(i)).DansUneMaison) Then ComMessages.Infos_Maison Persos(Nopersos(i)).IndiceMaison, ComReseau.PlayerID
                        'Fenetre bâtiment chateau.
                        If Interface(i).Fenetres_Visible(0) And (Persos(Nopersos(i)).DansSonChateau Or Persos(Nopersos(i)).DansUnChateau) Then ComMessages.Infos_Chateau Persos(Nopersos(i)).IndiceChateau, ComReseau.PlayerID
                        'Fenetre personnage.
                        If Interface(i).Fenetres_Visible(2) Then ComMessages.Infos_Perso_Caracs2 Nopersos(i), ComReseau.PlayerID
                        'Fenetre compétence.
                        If Interface(i).Fenetres_Visible(3) Then ComMessages.Infos_Perso_Competences2 Nopersos(i), ComReseau.PlayerID
                        'Fenêtre inventaire de la maison.
                        If Interface(i).Fenetre_InventaireMaison_Visible(4) Then ComMessages.Infos_Perso_InventaireMaison Nopersos(i), ComReseau.PlayerID
                        'Fenêtre bestiaire.
                        'If Interface(i).Fenetres_Visible(5) Then ComMessages.Infos_Perso_Bestiaire Nopersos(i), ComReseau.PlayerID
                        'Réactualise les informations automatiquement.
                        Persos(Nopersos(i)).Gestion_Tous_Niveaux
                        Persos(Nopersos(i)).Actualiser_Bonus
                        Persos(Nopersos(i)).Ressources_Verifications
                    Next i
                    ComMessages.Infos_PartieEnCours ComReseau.PlayerID
                    If Partie.EnCours < 2 Then
                        Quitter = True
                    End If
                    ComMessages.Infos_Monde_Encours ComReseau.PlayerID
                End If
            Else
                Lire_Commandes 'Lit les instructions donnée par le joueur.
            End If
            End With
            
            'Debug.Print DeltaTemps
            'If DureePartie > 50 Then
            '    GrabScreen
            '    SavePicture Clipboard.GetData, App.Path & CheminSavMonde & DureePartie & ExtensionFichiersAperçu
            'End If
            'Messages.Ajouter_Message DeltaTemps & " > " & 1 / Jeu.Definir_Vitesse
            'Ordres données aux ia.
            
            'debug.Print Format((Timer - TempsFin) * 1000, "00") & " ms" & " : IA"
            'If ArmeeSelectionnee Then
            '    Scrolling_Immediat Persos(SoldatsSelectionnes.Item(1)).PositionX, _
                                   Persos(SoldatsSelectionnes.Item(1)).PositionY
                'Scrolling_Lent
            'Else
            If NombreJoueursMO = 2 Then
                If EcranDivise Then
                    For i = 0 To NombreJoueursMO - 1
                        Scrolling_Immediat Persos(Nopersos(i)).PositionX, _
                                           Persos(Nopersos(i)).PositionY, _
                                           ScreenWidth / 2, _
                                           ScreenHeight, _
                                           i
                    Next i
                    'DecalageX(1) = DecalageX(1) + ScreenWidth / 2
                Else
                    For i = 0 To NombreJoueursMO - 1
                        Scrolling_Immediat (Persos(Nopersos(0)).PositionX + Persos(Nopersos(1)).PositionX) / 2, _
                                           (Persos(Nopersos(0)).PositionY + Persos(Nopersos(1)).PositionY) / 2, _
                                           ScreenWidth, _
                                           ScreenHeight, _
                                           i
                    Next i
                End If
            Else
                Scrolling_Immediat Persos(NoPerso).PositionX, _
                                   Persos(NoPerso).PositionY, _
                                   ScreenWidth, _
                                   ScreenHeight, _
                                   0
            End If
            
            If ComReseau.Client Then
                'Les clients restent immobiles.
                'Seules gestions effectuées par les clients.
'                If TempVitesse >= 10 Then
'                    TempVitesse = 0
'
'                End If
                Actions_Persos_Client
                TempVitesse = TempVitesse + 1
                If TempVitesse >= 10 Then
                    TempVitesse = 0
                    Actions_Persos_Client_Lent
                End If
            Else
                'Gestion du serveur ou du solo.
                Comportement_IA
                Comportement_IA_Fiefs
                Deplacer_Persos
                TempVitesse = TempVitesse + 1
                If TempVitesse >= 10 Then
                
                    TempVitesse = 0
                    Actions_Persos
                    Actions_Tresors
                End If
            
                Gestion_Utilisation_Speciale
                Evenements_Lancer "Jouer"
                
                If ComReseau.Connecte Then
                    ComReseau.Rafraichir_Joueurs
                    If ComReseau.Serveur Then
                        ComReseau.Joueur_Tous_InfosAJour = False
                    End If
                End If
            End If
            
            'debug.Print Format((Timer - TempsFin) * 1000, "00") & " ms" & " : Personnages"
            
            'Modifie les animations.
            AffTerrain.Animation_Suivante
            Actualiser_Affichage_Batiments
            Actualiser_Affichage_Decors
            Actualiser_Effets_Visuel
            Actualiser_Effets_Interfaces
            AffDegat.Actualiser_Degats
            AffTemps.Actualiser
            
            'Vérifie si victoire.
            For i = 0 To NombreJoueursMO - 1
                'If Not (Victoire(i) Or Defaite(i)) Then
                    If Jeu.VictoireStandard Then
                        Victoire(i) = True
                        For j = 0 To UBound(Fiefs())
                            If Fiefs(j).NumeroEquipe <> Persos(Nopersos(i)).NumeroEquipe Then
                                If Chateaux(j).Visible Or _
                                   Fiefs(j).NombreMorts < Fiefs(j).NombreCitoyens Then
                                        Victoire(i) = False
                                        j = UBound(Fiefs())
                                End If
                            End If
                        Next j
                    End If
                    If Victoire(i) Then
                        'AffVictoire(i).Selection 0, 200 - 1, 0, 320 - 1
                        If i = 0 Then Musique.Jouer_Musique 5, False
                    ElseIf Jeu.DefaiteStandard Then
                        Defaite(i) = True
                        If Chateaux(Persos(Nopersos(i)).NumeroFief).Visible Or _
                           Fiefs(Persos(Nopersos(i)).NumeroFief).NombreMorts < Fiefs(Persos(Nopersos(i)).NumeroFief).NombreCitoyens Then
                                Defaite(i) = False
                                'i = UBound(Fiefs())
                        End If
                        If Defaite(i) Then
                        '    AffVictoire(i).Selection 200, 400 - 1, 0, 320 - 1
                            If i = 0 Then Musique.Jouer_Musique 6, False
                        End If
                    End If
                'End If
                NoPerso = Nopersos(i)
                Interface(i).Actions_Automatiques Persos(NoPerso), Maisons(Persos(NoPerso).IndiceMaison), Parametres
                If Victoire(i) Then
                    AffVictoire(i).Selection 0, 200 - 1, 0, 320 - 1
                ElseIf Defaite(i) Then
                    AffVictoire(i).Selection 200, 400 - 1, 0, 320 - 1
                End If
            
            'Actualise l'interface.
            Interface(i).ActualiserLent Persos(NoPerso), Maisons(NoPerso), Fiefs(Persos(NoPerso).NumeroFief), Parametres, Messages(i)
            Next i
        End If

        If (DeltaTemps > 1 / Jeu.Definir_FPS Or _
           DeltaTemps < -(1 / Jeu.Definir_FPS)) Or Jeu.Definir_AffichageRapide Then  'Raffraichissement.
            'Affiche l'arrière plan avec une couleur uniforme.
            'SurfBack.BltColorFill MonFond, 0& ' 4396468
            
            'Debugger_TimeStamp True, 3
            
            'Affiche les éléments à l'écran.
            'Afficher_Terrain
            
            If NombreJoueursMO = 1 Or Not EcranDivise Then
                If Monde.TerrainOptimise Then
                    AffTerrain.AfficheVeryFast 0, 0, DecalageX(0), DecalageY(0), Ecran.Width, Ecran.Height, SurfBack
                Else
                    AffTerrain.AfficheFast 0, 0, DecalageX(0), DecalageY(0), Ecran.Width, Ecran.Height, SurfBack
                End If
                Afficher_Ressources 0
                Afficher_Batiments 0
                Afficher_Decors 0
                Afficher_Tresors 0
                For i = 0 To NombreJoueursMO - 1
                    If Not ArmeeSelectionnee(i) Then
                        Afficher_CadreSelection Nopersos(i), i, True
                    End If
                    Afficher_CadreSelection_Soldats i
                    Afficher_Etat_Travail Nopersos(i), i
                Next i
                
                If NombreJoueursMO = 2 Then
                    Afficher_Persos 0, 1
                Else
                    Afficher_Persos 0
                End If
                Afficher_Effets_Visuel 0
                AffTemps.Afficher DecalageX(0), DecalageY(0), Ecran.Width, Ecran.Height, 0, 0, SurfBack
                AffDegat.Afficher_Degats DecalageX(0) - ScreenPositionX, DecalageY(0) - ScreenPositionY, Ecran.Width, Ecran.Height, SurfBack
                Afficher_Commentaires 0, Ecran.Width, Ecran.Height
                For i = 0 To NombreJoueursMO - 1
                    Interface(i).Position_Contour_Fenetre SourisX(i), SourisY(i), Ecran.Width, Ecran.Height
                    Interface(i).RePositionner_Fenetres ScreenPositionX, ScreenPositionY, Ecran.Width, Ecran.Height
                    'Affiche les messages d'information.
                    Messages(i).Afficher_Messages SurfBack
                    'Affiche l'interface.
                    If Persos(Nopersos(i)).DansSaMaison Then
                        Interface(i).Afficher SourisX(i), SourisY(i), Persos(Nopersos(i)), Maisons(Nopersos(i)), Maisons(Nopersos(i)), Chateaux(Persos(Nopersos(i)).NumeroFief), Fiefs(Persos(Nopersos(i)).NumeroFief), ArmeeSelectionnee(i), Parametres, SurfBack, DecalageX(i), DecalageY(i), Ecran.Width, Ecran.Height
                    ElseIf Persos(Nopersos(i)).DansSonChateau Then
                        Interface(i).Afficher SourisX(i), SourisY(i), Persos(Nopersos(i)), Maisons(Persos(Nopersos(i)).IndiceMaison), Maisons(Nopersos(i)), Chateaux(Persos(Nopersos(i)).NumeroFief), Fiefs(Persos(Nopersos(i)).NumeroFief), ArmeeSelectionnee(i), Parametres, SurfBack, DecalageX(i), DecalageY(i), Ecran.Width, Ecran.Height
                    ElseIf Persos(Nopersos(i)).DansUnChateau Then
                        Interface(i).Afficher SourisX(i), SourisY(i), Persos(Nopersos(i)), Maisons(Persos(Nopersos(i)).IndiceMaison), Maisons(Nopersos(i)), Chateaux(Persos(Nopersos(i)).IndiceChateau), Fiefs(Persos(Nopersos(i)).IndiceChateau), ArmeeSelectionnee(i), Parametres, SurfBack, DecalageX(i), DecalageY(i), Ecran.Width, Ecran.Height
                    Else
                        Interface(i).Afficher SourisX(i), SourisY(i), Persos(Nopersos(i)), Maisons(Persos(Nopersos(i)).IndiceMaison), Maisons(Nopersos(i)), Chateaux(Persos(Nopersos(i)).NumeroFief), Fiefs(Persos(Nopersos(i)).NumeroFief), ArmeeSelectionnee(i), Parametres, SurfBack, DecalageX(i), DecalageY(i), Ecran.Width, Ecran.Height
                    End If
                    CurseurSurInterface(i) = Interface(i).Interieur_Boutons(i, SourisX(i), SourisY(i), SourisZ(i), NoBouton(i), NoBatiment(i), Persos(Nopersos(i)), Maisons(Persos(Nopersos(i)).IndiceMaison), Maisons(Nopersos(i)), Fiefs(Persos(Nopersos(i)).NumeroFief), Parametres, SurfBack, CliqueGauche(i), CliqueDroit(i), MaintenuEnfonceGauche(i), MaintenuEnfonceMilieu(i), MaintenuEnfonceDroit(i), Key, KCtrl(i), KShift(i), KAlt(i), ArmeeSelectionnee(i), SoldatsSelectionnes(i), i = 0, IIf(i = 1, Ecran.Width * 0.5, 0), 0, IIf(NombreJoueursMO = 1, Ecran.Width, Ecran.Width * 0.5), Ecran.Height)
                Next i
            Else
                If Monde.TerrainOptimise Then
                    AffTerrain.AfficheVeryFast 0, 0, DecalageX(0), DecalageY(0), Ecran.Width * 0.5, Ecran.Height, SurfBack
                    'AffTerrain.AfficheVeryFast Ecran.Width / 2, 0, DecalageX(1), DecalageY(1), Ecran.Width / 2, Ecran.Height, SurfBack
                Else
                    AffTerrain.AfficheFast 0, 0, DecalageX(0), DecalageY(0), Ecran.Width * 0.5, Ecran.Height, SurfBack
                    'AffTerrain.AfficheFast Ecran.Width * 0.5, 0, DecalageX(1), DecalageY(1), Ecran.Width / 2, Ecran.Height, SurfBack
                End If
                
                ScreenPositionX = 0
                ScreenWidth = Ecran.Width / 2
                
                Afficher_Ressources 0
                Afficher_Batiments 0
                Afficher_Decors 0
                Afficher_Tresors 0
                If Not ArmeeSelectionnee(0) Then Afficher_CadreSelection Nopersos(0), 0
                Afficher_CadreSelection_Soldats 0
                Afficher_Etat_Travail Nopersos(0), 0
                
                Afficher_Persos 0
                Afficher_Effets_Visuel 0
                AffTemps.Afficher DecalageX(0), DecalageY(0), Ecran.Width * 0.5, Ecran.Height, DecalageX(0), DecalageY(0), SurfBack
                AffDegat.Afficher_Degats DecalageX(0) - ScreenPositionX, DecalageY(0) - ScreenPositionY, Ecran.Width * 0.5, Ecran.Height, SurfBack
                Afficher_Commentaires 0, Ecran.Width * 0.5, Ecran.Height
                Interface(0).Position_Contour_Fenetre SourisX(0), SourisY(0), Ecran.Width * 0.5, Ecran.Height
                Interface(0).RePositionner_Fenetres ScreenPositionX, ScreenPositionY, Ecran.Width * 0.5, Ecran.Height
                'Affiche les messages d'information.
                Messages(0).Afficher_Messages SurfBack
                'Affiche l'interface.
                If Persos(Nopersos(0)).DansSaMaison Then
                    Interface(0).Afficher SourisX(0), SourisY(0), Persos(Nopersos(0)), Maisons(Nopersos(0)), Maisons(Nopersos(0)), Chateaux(Persos(Nopersos(0)).NumeroFief), Fiefs(Persos(Nopersos(0)).NumeroFief), ArmeeSelectionnee(0), Parametres, SurfBack, DecalageX(0), DecalageY(0), ScreenWidth, Ecran.Height
                ElseIf Persos(Nopersos(0)).DansSonChateau Then
                    Interface(0).Afficher SourisX(0), SourisY(0), Persos(Nopersos(0)), Maisons(Persos(Nopersos(0)).IndiceMaison), Maisons(Nopersos(0)), Chateaux(Persos(Nopersos(0)).NumeroFief), Fiefs(Persos(Nopersos(0)).NumeroFief), ArmeeSelectionnee(0), Parametres, SurfBack, DecalageX(0), DecalageY(0), ScreenWidth, Ecran.Height
                Else
                    Interface(0).Afficher SourisX(0), SourisY(0), Persos(Nopersos(0)), Maisons(Persos(Nopersos(0)).IndiceMaison), Maisons(Nopersos(0)), Chateaux(Persos(Nopersos(0)).IndiceChateau), Fiefs(Persos(Nopersos(0)).IndiceChateau), ArmeeSelectionnee(0), Parametres, SurfBack, DecalageX(0), DecalageY(0), ScreenWidth, Ecran.Height
                End If
                CurseurSurInterface(0) = Interface(0).Interieur_Boutons(0, SourisX(0), SourisY(0), SourisZ(0), NoBouton(0), NoBatiment(0), Persos(Nopersos(0)), Maisons(Persos(Nopersos(0)).IndiceMaison), Maisons(Nopersos(0)), Fiefs(Persos(Nopersos(0)).NumeroFief), Parametres, SurfBack, CliqueGauche(0), CliqueDroit(0), MaintenuEnfonceGauche(0), MaintenuEnfonceMilieu(0), MaintenuEnfonceDroit(0), Key, KCtrl(0), KShift(0), KAlt(0), ArmeeSelectionnee(0), SoldatsSelectionnes(0), True, 0, 0, Ecran.Width * 0.5, Ecran.Height)
                
                ScreenPositionX = Ecran.Width / 2
                
                If Monde.TerrainOptimise Then
                    'AffTerrain.AfficheVeryFast ScreenPositionX, 0, DecalageX(1), DecalageY(1), ScreenWidth, Ecran.Height, SurfBack
                    ScreenWidth = Ecran.Width
                    AffTerrain.AfficheFast Ecran.Width * 0.5, 0, DecalageX(1), DecalageY(1), Ecran.Width * 0.5, Ecran.Height, SurfBack
                    ScreenWidth = Ecran.Width / 2
                Else
                    'AffTerrain.AfficheFast ScreenPositionX, 0, DecalageX(1), DecalageY(1), ScreenWidth, Ecran.Height, SurfBack
                    ScreenWidth = Ecran.Width
                    AffTerrain.AfficheFast Ecran.Width * 0.5, 0, DecalageX(1), DecalageY(1), Ecran.Width * 0.5, Ecran.Height, SurfBack
                    ScreenWidth = Ecran.Width / 2
                End If
                
                Afficher_Ressources 1
                Afficher_Batiments 1
                Afficher_Decors 1
                Afficher_Tresors 1
                
                Afficher_Separateur_Ecran
                
                If Not ArmeeSelectionnee(1) Then Afficher_CadreSelection Nopersos(1), 1
                Afficher_CadreSelection_Soldats 1
                Afficher_Etat_Travail Nopersos(1), 1
                
                Afficher_Persos 1
                Afficher_Effets_Visuel 1
                AffTemps.Afficher DecalageX(1), DecalageY(1), Ecran.Width * 0.5, Ecran.Height, DecalageX(1), DecalageY(1), SurfBack
                AffDegat.Afficher_Degats DecalageX(1) - ScreenPositionX, DecalageY(1) - ScreenPositionY, Ecran.Width * 0.5, Ecran.Height, SurfBack
                Afficher_Commentaires 1, Ecran.Width * 0.5, Ecran.Height
                Interface(1).Position_Contour_Fenetre SourisX(1), SourisY(1), Ecran.Width, Ecran.Height
                Interface(1).RePositionner_Fenetres ScreenPositionX, ScreenPositionY, Ecran.Width * 0.5, Ecran.Height
                'Affiche les messages d'information.
                Messages(1).Afficher_Messages SurfBack
                'Affiche l'interface.
                ScreenPositionX = 0
                ScreenWidth = Ecran.Width
                If Persos(Nopersos(1)).DansSaMaison Then
                    Interface(1).Afficher SourisX(1), SourisY(1), Persos(Nopersos(1)), Maisons(Nopersos(1)), Maisons(Nopersos(1)), Chateaux(Persos(Nopersos(1)).NumeroFief), Fiefs(Persos(Nopersos(1)).NumeroFief), ArmeeSelectionnee(1), Parametres, SurfBack, DecalageX(1), DecalageY(1), Ecran.Width * 0.5, Ecran.Height
                ElseIf Persos(Nopersos(1)).DansSonChateau Then
                    Interface(1).Afficher SourisX(1), SourisY(1), Persos(Nopersos(1)), Maisons(Persos(Nopersos(1)).IndiceMaison), Maisons(Nopersos(1)), Chateaux(Persos(Nopersos(1)).NumeroFief), Fiefs(Persos(Nopersos(1)).NumeroFief), ArmeeSelectionnee(1), Parametres, SurfBack, DecalageX(1), DecalageY(1), Ecran.Width * 0.5, Ecran.Height
                Else
                    Interface(1).Afficher SourisX(1), SourisY(1), Persos(Nopersos(1)), Maisons(Persos(Nopersos(1)).IndiceMaison), Maisons(Nopersos(1)), Chateaux(Persos(Nopersos(1)).IndiceChateau), Fiefs(Persos(Nopersos(1)).IndiceChateau), ArmeeSelectionnee(1), Parametres, SurfBack, DecalageX(1), DecalageY(1), Ecran.Width * 0.5, Ecran.Height
                End If
                CurseurSurInterface(1) = Interface(1).Interieur_Boutons(1, SourisX(1), SourisY(1), SourisZ(1), NoBouton(1), NoBatiment(1), Persos(Nopersos(1)), Maisons(Persos(Nopersos(1)).IndiceMaison), Maisons(Nopersos(1)), Fiefs(Persos(Nopersos(1)).NumeroFief), Parametres, SurfBack, CliqueGauche(1), CliqueDroit(1), MaintenuEnfonceGauche(1), MaintenuEnfonceMilieu(1), MaintenuEnfonceDroit(1), Key, KCtrl(1), KShift(1), KAlt(1), ArmeeSelectionnee(1), SoldatsSelectionnes(1), False, Ecran.Width * 0.5, 0, Ecran.Width * 0.5, Ecran.Height)
            End If
            
            'Changer_Zoom Zoom
            
            'Debugger_TimeStamp False, 3
            'debug.Print Format((Timer - TempsFin) * 1000, "00") & " ms" & " : Affichage Ecran"
            'Debugger_TimeStamp True, 4
        
            For i = 0 To NombreJoueursMO - 1
                NoPerso = Nopersos(i)
                'Affiche les messages d'information.
                'Messages(i).Afficher_Messages Persos(Noperso), Interface(i), Parametres, SurfBack
                
                'If EcranDivise Then
                '    ScreenWidth = Ecran.Width / 2
                '    If i = 0 Then ScreenPositionX = Ecran.Width / 2
                'End If
                'Affiche l'interface.
                'If Persos(Noperso).DansSaMaison Then
                '    Interface(i).Afficher SourisX(i), SourisY(i), Persos(Noperso), Maisons(Noperso), Maisons(Noperso), Chateaux(Persos(Noperso).NumeroFief), Fiefs(Persos(Noperso).NumeroFief), Parametres, SurfBack
                'ElseIf Persos(Noperso).DansSonChateau Then
                '    Interface(i).Afficher SourisX(i), SourisY(i), Persos(Noperso), Maisons(Persos(Noperso).IndiceMaison), Maisons(Noperso), Chateaux(Persos(Noperso).NumeroFief), Fiefs(Persos(Noperso).NumeroFief), Parametres, SurfBack
                'Else
                '    Interface(i).Afficher SourisX(i), SourisY(i), Persos(Noperso), Maisons(Persos(Noperso).IndiceMaison), Maisons(Noperso), Chateaux(Persos(Noperso).IndiceChateau), Fiefs(Persos(Noperso).IndiceChateau), Parametres, SurfBack
                'End If
                'ScreenPositionX = 0
                'ScreenWidth = Ecran.Width
                
                If Interface(i).DeplacerFenetre >= 0 Then
                    Interface(i).Afficher_Contour_Fenetre SourisX(i), SourisY(i), SurfBack
                    AffSouris.Afficher SourisX(i), _
                                       SourisY(i), _
                                       0, _
                                       SurfBack
                Else
                    If NoBouton(i) > 0 And NoBouton(i) < 100 Then
                        TypeCurseur(i) = 2
                    End If
                    If Persos(NoPerso).Definir_ObjetSelectionne Then
                        Interface(i).Afficher_Objet_Curseur Persos(NoPerso), SourisX(i) - AffSouris.Largeur / 2, SourisY(i) - AffSouris.Hauteur / 2, SurfBack
                    End If
                    If Persos(NoPerso).Vivant Then
                        If TypeCurseur(i) = 2 Or TypeCurseur(i) = 3 Then
                            If Persos(NoPerso).PeutConstruire_Maison(NoBatiment(i) - 1, Parametres) Then
                                If Obstacle(NoPerso, _
                                            SourisX(i) - Parametres.Batiment_Largeur(NoBatiment(i) - 1) / 2 + AffSouris.Largeur / 2 + DecalageX(i) - IIf(i = 1 And EcranDivise, Ecran.Width / 2, 0), _
                                            SourisY(i) - Parametres.Batiment_Hauteur(NoBatiment(i) - 1) / 2 + AffSouris.Hauteur / 2 + DecalageY(i), _
                                            Parametres.Batiment_Largeur(NoBatiment(i) - 1), _
                                            Parametres.Batiment_Hauteur(NoBatiment(i) - 1)) Or _
                                            SourisX(i) - Parametres.Batiment_Largeur(NoBatiment(i) - 1) / 2 + AffSouris.Largeur / 2 + DecalageX(i) - IIf(i = 1 And EcranDivise, Ecran.Width / 2, 0) < 0 Or _
                                            SourisY(i) - Parametres.Batiment_Hauteur(NoBatiment(i) - 1) / 2 + AffSouris.Hauteur / 2 + DecalageY(i) < 0 Or _
                                            SourisX(i) + Parametres.Batiment_Largeur(NoBatiment(i) - 1) - AffSouris.Largeur + DecalageX(i) - IIf(i = 1 And EcranDivise, Ecran.Width / 2, 0) > Monde.Largeur * AffTerrain.Largeur Or _
                                            SourisY(i) + Parametres.Batiment_Hauteur(NoBatiment(i) - 1) - AffSouris.Hauteur + DecalageY(i) > Monde.Hauteur * AffTerrain.Hauteur Then
                                    TypeCurseur(i) = 3
                                    Afficher_Surface_A_Batir SourisX(i), SourisY(i), NoBatiment(i) - 1, False
                                Else
                                    TypeCurseur(i) = 2
                                    Afficher_Surface_A_Batir SourisX(i), SourisY(i), NoBatiment(i) - 1, True
                                End If
                            Else
                                TypeCurseur(i) = 0
                            End If
                        Else
                            If CurseurSurInterface(i) Then
                                If NoBouton(i) = 900 Then
                                    TypeCurseur(i) = 15
                                Else
                                    TypeCurseur(i) = 0
                                End If
                            ElseIf Persos(NoPerso).Definir_ObjetSelectionne Then
                                Select Case ObstacleAmi(NoPerso, SourisX(i) + DecalageX(i) - IIf(i = 1 And EcranDivise, Ecran.Width / 2, 0), SourisY(i) + DecalageY(i)) ', Not ArmeeSelectionnee(i))
                                Case 1: 'Donne à ses alliés. (personnage)
                                    If Persos(NoPerso).RessourceSelectionne Or Persos(NoPerso).ArgentSelectionne > 0 Then
                                        'Va donner des ressources ou de l'argent.
                                        TypeCurseur(i) = 14
                                        IndiceCible(i) = IndiceObstacle
                                    ElseIf Not Parametres.Objet_Inseparable(Persos(NoPerso).ObjetSelectionne) Then
                                        'Va donner un objet.
                                        TypeCurseur(i) = 14
                                        IndiceCible(i) = IndiceObstacle
                                    Else
                                        TypeCurseur(i) = 0
                                    End If
                                Case 2: 'Donne à ses alliés. (maison)
                                    If Maisons(IndiceObstacle).Marche And Maisons(IndiceObstacle).Construit Then
                                        TypeCurseur(i) = 6
                                        IndiceCible(i) = IndiceObstacle
                                    Else
                                        If Persos(NoPerso).RessourceSelectionne Or Persos(NoPerso).ArgentSelectionne > 0 Then
                                            'Va donner des ressources ou de l'argent.
                                            TypeCurseur(i) = 14
                                            IndiceCible(i) = IndiceObstacle
                                        ElseIf Not Parametres.Objet_Inseparable(Persos(NoPerso).ObjetSelectionne) Then
                                            'Va donner des ressources ou un objet.
                                            TypeCurseur(i) = 14
                                            IndiceCible(i) = IndiceObstacle
                                        Else
                                            TypeCurseur(i) = 0
                                        End If
                                    End If
                                Case 3: 'Vend au château.
                                    TypeCurseur(i) = 6
                                    IndiceCible(i) = IndiceObstacle
                                Case 4: 'Vend à un château allié.
                                    TypeCurseur(i) = 8
                                Case Else
                                    TypeCurseur(i) = 0
                                End Select
                            Else
                                If i = 1 And EcranDivise Then
                                    TypeCurseur(i) = Obstacle(NoPerso, SourisX(i) + DecalageX(i) - Ecran.Width / 2, SourisY(i) + DecalageY(i), 0, 0, True)
                                Else
                                    TypeCurseur(i) = Obstacle(NoPerso, SourisX(i) + DecalageX(i), SourisY(i) + DecalageY(i), 0, 0, True)
                                End If
                                IndiceCible(i) = IndiceObstacle
                                If (TypeCurseur(i) = 0 Or _
                                    TypeCurseur(i) = 1 Or TypeCurseur(i) = 7 Or _
                                    TypeCurseur(i) = 999) And _
                                    ArmeeSelectionnee(i) Then
                                    If 1 = ObstacleAmi(NoPerso, SourisX(i) + DecalageX(i) - IIf(i = 1 And EcranDivise, Ecran.Width / 2, 0), SourisY(i) + DecalageY(i)) Or _
                                       2 = ObstacleAmi(NoPerso, SourisX(i) + DecalageX(i) - IIf(i = 1 And EcranDivise, Ecran.Width / 2, 0), SourisY(i) + DecalageY(i)) Then ', Not ArmeeSelectionnee(i)) Then
                                        'Affecte certains de ces soldats à la protection d'une personne.
                                        TypeCurseur(i) = 17
                                        IndiceCible(i) = IndiceObstacle
                                    End If
                                ElseIf TypeCurseur(i) = 999 Then
                                    TypeCurseur(i) = 0
                                End If
                            End If
                        End If
                        Select Case TypeCurseur(i)
                        Case 4:
                            If KCtrl(i) And Maisons(IndiceCible(i)).Construit Then
                                TypeCurseur(i) = 5
                            End If
                        Case 7: 'Maisons alliée.
                            If KCtrl(i) And Maisons(IndiceCible(i)).Vie < Maisons(IndiceCible(i)).MaxVie Then
                                TypeCurseur(i) = 4
                            End If
                        End Select
                    Else
                        If CurseurSurInterface(i) Or Persos(NoPerso).Definir_ObjetSelectionne Then
                            If NoBouton(i) = 900 Then
                                TypeCurseur(i) = 15
                            Else
                                TypeCurseur(i) = 0
                            End If
                        Else
                            'Cas où le joueur est un fantôme.
                            'Il ne peut que rentrer au chateau ou dans un temple.
                            TypeCurseur(i) = ObstacleMort(NoPerso, SourisX(i) + DecalageX(i) - IIf(i = 1 And EcranDivise, Ecran.Width / 2, 0), SourisY(i) + DecalageY(i), 0, 0, True)
                            IndiceCible(i) = IndiceObstacle
                        End If
                    End If
                    AffSouris.Afficher SourisX(i), _
                                       SourisY(i), _
                                       TypeCurseur(i), _
                                       SurfBack
                    Afficher_Infos_Bulles_Terrain TypeCurseur(i), Persos(NoPerso), SourisX(i), SourisY(i), i
                End If
                
                'Affichage du nombre de frames par secondes.
                If i = 0 Then FPS = FPS + 1
                If Abs(TempsFin - DernierFPS) > 1 Then
                    MaxFPS = FPS
                    FPS = 0
                    DernierFPS = TempsFin
                    MessageFPS = Format$(MaxFPS, "00")
                    If ComReseau.Connecte Then
                        MessageFPS = MessageFPS & AffMessage.Caractere_RetourLigne & _
                                                "E : " & ComReseau.Definir_NombreMessagesEmis & _
                                                AffMessage.Caractere_RetourLigne & _
                                                "R : " & ComReseau.Definir_NombreMessagesRecus
                        'MessageFPS = MessageFPS & AffMessage.Caractere_RetourLigne & ComReseau.PlayerID
                    End If
                End If
                Interface(i).Afficher_FPS MessageFPS, SurfBack
                
                If Victoire(i) Then
                    AffVictoire(i).Afficher (IIf(NombreJoueursMO = 1, ScreenWidth, ScreenWidth * (0.5 + i)) - 320) / 2, (ScreenHeight - 200) / 2 - 128, SurfBack, True ', True
                ElseIf Defaite(i) Then
                    AffVictoire(i).Afficher (IIf(NombreJoueursMO = 1, ScreenWidth, ScreenWidth * (0.5 + i)) - 320) / 2, (ScreenHeight - 200) / 2 - 128, SurfBack, True ', True
                End If
                SourisZ(i) = 0
            Next i
            'debug.Print Format((Timer - TempsFin) * 1000, "00") & " ms" & " : Affichage Interface"
            
            'Debugger_TimeStamp False, 4
            'Debugger_TimeStamp True, 5
            
            'Affichage.
            Rafraichir_Ecran
            
            'debug.Print Format((Timer - TempsFin) * 1000, "00") & " ms" & " : Fin"
        End If
        
        'Quitter = Victoire
        'Quitter = True
        'Victoire = True
        
        'Debugger_TimeStamp False, 5
        
        'Sleep 1
        DoEvents '<== astuce pour laisser "respirer" le système
        
        'Debugger_TimeStamp False, 6
    Loop
    
    'UniversalDebugger.Debugger_Unload
    'UniversalDebugger.Debugger_WriteText "_debug.log"
    
    'Son.StopSon
    
    Set ImageFS = Nothing
    
    'Sauvegarde des personnages.
    'If ComReseau.Client Then
'    If Jeu.Confrontation Then
'        '###
'    Else
        If Jeu.Duree > 10 Then 'Sauvegarde si que la partie a duré au moins 1 seconde.
            For i = 0 To UBound(Persos())
                If Persos(i).Joueur = ComReseau.PlayerID Then
                    Enregistrer_Sauvegarde_Perso Persos(i)
                End If
            Next i
        End If
'    End If
    'Sauvegarde de l'interface.
    For i = 0 To NombreJoueursMO - 1
        Messages(i).Effacer
        If Persos(Nopersos(i)).Joueur = ComReseau.PlayerID Then
            Enregistrer_Sauvegarde_Interface Persos(Nopersos(i)), Interface(i)
        End If
    Next i
    'Sauvegarde du monde
    If Definir_SavMondeEnQuittant And Not (Victoire(0) Or Defaite(0)) And Not Jeu.Campagne Then
        Enregistrer_Sauvegarde_Monde SavMondeEnQuittantNom
    End If
    
    Partie.EnCours = 0
    
    'Efface le personnage en jeu, si l'on est un client.
    If ComReseau.Client Then
        ComMessages.Infos_Perso_IndiceSupprimer ComReseau.PlayerID
    End If
    
    Stastistiques.Fin_Partie
    Document_Stastistiques
    'Unload Me
    Set AffTerrain = Nothing
End Sub

Public Sub Deplacer_Persos()
    Dim i As Long
    Dim NoMessage As Integer
    For i = 0 To UBound(Persos())
        With Persos(i)
        NoMessage = -1
        If .IA_Berserk > 0 Then
            If .Vivant Then 'And .IA
                IA_ComportementBerserk i
            End If
        End If
        .RecuperationRapide
        .Deplacement1 Ressources(.IndiceRessource), Persos(.IndicePerso), Maisons(.IndiceMaison), Chateaux(.IndiceChateau), Fiefs(.NumeroFief), NoMessage, Commentaires, Parametres
        .Deplacement2 Ressources(.IndiceRessource), Persos(.IndicePerso), Maisons(.IndiceMaison), Chateaux(.IndiceChateau), Fiefs(.NumeroFief), NoMessage, Commentaires, Parametres
        .Travailler Ressources(.IndiceRessource), Persos(.IndicePerso), Maisons(.IndiceMaison), Chateaux(.IndiceChateau), Fiefs(.NumeroFief), NoMessage, Commentaires, Parametres
        '.RecuperationLente Fiefs(.NumeroFief)
        '.Travailler Ressources(.IndiceRessource), Persos(.IndicePerso), Maisons(.IndiceMaison), Chateaux(.IndiceChateau), Fiefs(.NumeroFief), NoMessage, Commentaires, Parametres
        'Affichage des commentaires.
        If NoMessage >= 0 Then
            If Rnd > 0.99 Or i = NoPerso Then 'Persos(i).Joueur > 0 Then
                .Definir_Commentaires = Commentaires.Message(NoMessage, Persos(i))
            End If
        End If
        End With
    Next i
End Sub

Public Sub Actions_Persos()
    Dim j As Long
    For j = 0 To UBound(Persos()) 'NombreActionsParTour - 1
        With Persos(j)
        'NoMessage = -1
        .RecuperationLente Maisons(.Numero), Fiefs(.NumeroFief), Parametres
        'Affichage des commentaires.
        'If NoMessage >= 0 Then
        '    If RND > 0.99 Or IndicePerso = Noperso Then 'Persos(i).Joueur > 0 Then
        '        .Definir_Commentaires = Commentaires.Message(NoMessage, Persos(IndicePerso))
        '    End If
        'End If
        End With
        'IndicePerso = IndicePerso + 1
        'If IndicePerso > UBound(Persos()) Then
        '    IndicePerso = 0
        'End If
    Next j
End Sub

Private Sub Actions_Persos_Client()
    'Gestion des personnages exclusivement aux clients réseau.
    Dim j As Long
    For j = 0 To UBound(Persos())
        With Persos(j)
        .Changer_Pas
        .EnDeplacement = False
        End With
    Next j
End Sub
Private Sub Actions_Persos_Client_Lent()
    'Gestion des personnages exclusivement aux clients réseau.
    Dim j As Long
    For j = 0 To UBound(Persos())
        With Persos(j)
        .Gestion_Effets_Temporaires
        .Gestion_Effets_Recharge
        End With
    Next j
End Sub
Public Sub Actions_Tresors()
    'Fait bouger les trésors.
    If Jeu.Confrontation Then 'Fait apparaître des trésors en mode confrontation.
        Tresor_Generer
    End If
End Sub

Private Sub Scrolling_Immediat(ByVal X As Long, _
                               ByVal Y As Long, _
                               ByVal l As Long, _
                               ByVal H As Long, _
                               ByVal NoJoueur As Integer)
    'On centre l'écran sur le joueur
    If ScrollingAuto(NoJoueur) Then
        DecalageX(NoJoueur) = X - l / 2
        If DecalageX(NoJoueur) > Monde.Largeur * AffTerrain.Largeur - l Then
            DecalageX(NoJoueur) = Monde.Largeur * AffTerrain.Largeur - l
        ElseIf DecalageX(NoJoueur) < 0 Then
            DecalageX(NoJoueur) = 0
        End If
        DecalageY(NoJoueur) = Y - H / 2
        If DecalageY(NoJoueur) > Monde.Hauteur * AffTerrain.Hauteur - H Then
            DecalageY(NoJoueur) = Monde.Hauteur * AffTerrain.Hauteur - H
        ElseIf DecalageY(NoJoueur) < 0 Then
            DecalageY(NoJoueur) = 0
        End If
    'Else
    '    ScrollingAuto(NoJoueur) = ScrollingAuto(NoJoueur) - 1
    End If
End Sub

'Private Sub Scrolling_Lent()
'    'Ajuste le scolling pour le centrer sur le personnage.
'    Dim Decalage As Long
'    With Persos(Noperso)
'    If .Direction > 0 Then
'        Decalage = .Longueur_Pas
'        'Scrolling en bas.
'        If .PositionY > ScreenHeight / 2 + DecalageY + Decalage Then
'            DecalageY = DecalageY + Decalage
'            If DecalageY > Monde.Hauteur * AffTerrain.Hauteur - Ecran.Height Then
'                DecalageY = Monde.Hauteur * AffTerrain.Hauteur - Ecran.Height
'            End If
'        'Scrolling en haut.
'        ElseIf .PositionY < ScreenHeight / 2 + DecalageY - Decalage Then
'            DecalageY = DecalageY - Decalage
'            If DecalageY < 0 Then
'                DecalageY = 0
'            End If
'        End If
'        'Scrolling à droite.
'        If .PositionX > ScreenWidth / 2 + DecalageX + Decalage Then
'            DecalageX = DecalageX + Decalage
'            If DecalageX > Monde.Largeur * AffTerrain.Largeur - Ecran.Width Then
'                DecalageX = Monde.Largeur * AffTerrain.Largeur - Ecran.Width
'            End If
'        'Scrolling à gauche.
'        ElseIf .PositionX < ScreenWidth / 2 + DecalageX - Decalage Then
'            DecalageX = DecalageX - Decalage
'            If DecalageX < 0 Then
'                DecalageX = 0
'            End If
'        End If
'    End If
'    End With
'End Sub

'Private Sub Afficher_Terrain()
'    If NombreJoueursMO = 1 Or Not EcranDivise Then
'        If Monde.TerrainOptimise Then
'            AffTerrain.AfficheVeryFast 0, 0, DecalageX(0), DecalageY(0), Ecran.Width, Ecran.Height, SurfBack
'        Else
'            AffTerrain.AfficheFast 0, 0, DecalageX(0), DecalageY(0), Ecran.Width, Ecran.Height, SurfBack
'        End If
'    Else
'        If Monde.TerrainOptimise Then
'            AffTerrain.AfficheVeryFast 0, 0, DecalageX(0), DecalageY(0), Ecran.Width / 2, Ecran.Height, SurfBack
'            AffTerrain.AfficheVeryFast Ecran.Width / 2, 0, DecalageX(1), DecalageY(1), Ecran.Width / 2, Ecran.Height, SurfBack
'        Else
'            AffTerrain.AfficheFast 0, 0, DecalageX(0), DecalageY(0), Ecran.Width * 0.5, Ecran.Height, SurfBack
'            AffTerrain.AfficheFast Ecran.Width * 0.5, 0, DecalageX(1), DecalageY(1), Ecran.Width / 2, Ecran.Height, SurfBack
'        End If
'    End If

    'Dim i As Long, j As Long, k As Long
    'Dim h As Long, dx As Long, dy As Long

    'With AffTerrain
    'h = Ecran.Height / .Hauteur
    'dx = DecalageX Mod .Largeur
    'dy = DecalageY Mod .Hauteur

    'For i = 0 To Ecran.Width / .Largeur
    '    k = Int(DecalageY / .Largeur) * Monde.Largeur + i + _
    '        Int(DecalageX / .Hauteur)
    '    For j = 0 To h
    '        .Afficher i * .Largeur - dx, _
    '                  j * .Hauteur - dy, _
    '                  Monde.Param_TerrainApparence(k), _
    '                  Monde.Param_TerrainCategorie(k), _
    '                  SurfBack
'
    '        k = k + Monde.Largeur
    '    Next j
    'Next i

    'End With
    
'End Sub

Private Sub Actualiser_Affichage_Batiments()
    Dim i As Long
    For i = 0 To UBound(Maisons())
        If Maisons(i).Visible Then
            Maisons(i).Actualiser_NumeroAnimation
        End If
    Next i
End Sub

Private Sub Afficher_Batiments(ByVal NoJoueur As Integer)
    Dim i As Long
    SurfBack.SetForeColor 0
    For i = 0 To UBound(Maisons())
        With Maisons(i)
        If .Visible Then
            AffMaison.Afficher .PositionX - DecalageX(NoJoueur), _
                               .PositionY - DecalageY(NoJoueur), _
                               .TypeBatiment, _
                               .Etat(Parametres), _
                               SurfBack, _
                               .Largeur, _
                               .Hauteur, _
                               .Definir_NumeroAnimation(Parametres)
        End If
        If .Definir_Vie < .MaxVie And _
           .Visible Then
            Interface(0).Afficher_Barre_Statut .Definir_Vie, .Definir_Vie, .MaxVie, _
                                            .PositionX - DecalageX(NoJoueur), _
                                            .PositionY - DecalageY(NoJoueur) + .Hauteur, _
                                            SurfBack, _
                                            .Largeur, _
                                            5
        End If
        End With
    Next i
    'Affiche les chateaux.
    For i = 0 To UBound(Chateaux())
        With Chateaux(i)
        If .Visible Then
            AffChateau.Afficher .PositionX - DecalageX(NoJoueur), _
                                .PositionY - DecalageY(NoJoueur), _
                                Fiefs(i).Definir_TypePeuple, _
                                IIf(Fiefs(i).Definir_TypePeuple = 0, Fiefs(i).Epoque, 0), _
                                SurfBack, _
                                Chateaux(i).Largeur, _
                                Chateaux(i).Hauteur
        End If
        If .Visible Then
            If .Vie < .MaxVie Then
                Interface(0).Afficher_Barre_Statut .Vie, .Definir_Vie, Chateaux(i).MaxVie, _
                                                   .PositionX - DecalageX(NoJoueur), _
                                                   .PositionY - DecalageY(NoJoueur) + Chateaux(i).Hauteur, _
                                                   SurfBack, _
                                                   Chateaux(i).Largeur, _
                                                   5
            End If
            If Fiefs(i).ChangerEpoque Then
                Interface(0).Afficher_Barre_Statut Parametres.Epoque_Duree(Fiefs(i).Epoque + 1) * Parametres.VitesseEpoque_Coef - Fiefs(i).TempsChangementEpoque, 0, Parametres.Epoque_Duree(Fiefs(i).Epoque + 1) * Parametres.VitesseEpoque_Coef, _
                                                .PositionX + .Largeur - DecalageX(NoJoueur), _
                                                .PositionY - DecalageY(NoJoueur), _
                                                SurfBack, _
                                                5, _
                                                Chateaux(i).Hauteur, &HFFFFFF, True
            End If
        End If
        End With
    Next i
End Sub

Private Sub Afficher_Ressources(ByVal NoJoueur As Integer)
    Dim i As Long
    For i = 0 To UBound(Ressources())
        With Ressources(i)
        AffRessources.Afficher .PositionX - DecalageX(NoJoueur), _
                               .PositionY - DecalageY(NoJoueur), _
                               .Apparence, _
                               SurfBack
        End With
    Next i
End Sub

Private Sub Afficher_Tresors(ByVal NoJoueur As Integer)
    Dim i As Long
    For i = 1 To Tresors.Count
        With Tresors(i)
        AffTresor.Afficher .X - DecalageX(NoJoueur), _
                           .Y - DecalageY(NoJoueur), _
                           .IndiceImage, _
                           SurfBack
        End With
    Next i
End Sub

Private Sub Afficher_Decors(ByVal NoJoueur As Integer)
    Dim i As Long
    For i = 1 To Decors.Count
        With Decors(i)
        AffDecor.Afficher .X - DecalageX(NoJoueur), _
                          .Y - DecalageY(NoJoueur), _
                          .Categorie, _
                          SurfBack, _
                          .EtapeAnimation
        End With
    Next i
End Sub

Private Sub Afficher_Persos(ByVal NoJoueur1 As Integer, _
                            Optional ByVal NoJoueur2 As Integer = -1)
    Dim Temp As Integer
    Dim i As Long
    'Dim CouleurMessage As Long
    If NoJoueur2 = -1 Then NoJoueur2 = NoJoueur1
    
    For i = UBound(Persos()) To 0 Step -1
        With Persos(i)
        If .Visible And _
           ((.Vivant Or (i = Nopersos(NoJoueur1) Or i = Nopersos(NoJoueur2)) And Jeu.Definir_AfficherFantomes) Or _
           Jeu.Definir_AfficherFantomesPNJ And .NumeroFief = Persos(NoPerso).NumeroFief) And _
           (Not .Invocation Or .Vivant) Then 'On n'affiche pas les fantômes des invocations.
            Temp = .NumeroPas
            'Afficher_Chemin Persos(i), DecalageX(NoJoueur1), DecalageY(NoJoueur1)
            'Affiche les personnages présents à l'écran.
            'Affiche les barres d'état.
            'If Persos(i).Action >= 300 And Persos(i).Action < 400 Then
               'i = Persos(NoPerso).IndicePerso And Persos(NoPerso).Action >= 300 And Persos(NoPerso).Action < 400 Then
               'Affiche les barres de statut de l'attaquant.
                'Interface.Afficher_Barre_Statut Persos(i).Vie, Persos(i).Vie, Persos(i).MaxVie, _
                                                Persos(i).PositionX - DecalageX, _
                                                Persos(i).PositionY - DecalageY + persos(i).Hauteur, _
                                                SurfBack, _
                                                persos(i).Largeur, _
                                                5
                'Interface.Afficher_Barre_Statut Persos(i).Attaque, Persos(i).Attaque, Persos(i).MaxAttaque, _
                                                Persos(i).PositionX - DecalageX, _
                                                Persos(i).PositionY - DecalageY + persos(i).Hauteur + 5, _
                                                SurfBack, _
                                                persos(i).Largeur, _
                                                5, _
                                                &HFF&          'Rouge.
                'Interface.Afficher_Barre_Statut Persos(i).Defense, Persos(i).Defense, Persos(i).MaxDefense, _
                                                Persos(i).PositionX - DecalageX, _
                                                Persos(i).PositionY - DecalageY + persos(i).Hauteur + 10, _
                                                SurfBack, _
                                                persos(i).Largeur, _
                                                5, _
                                                &HFF0000    'Bleu.
            'End If
            'Affiche les barres de statut du défenseur.
            'If Persos(i).Vie < Persos(i).MaxVie And Persos(i).Vivant Or _
               Persos(i).Attaque < Persos(i).Definir_MaxAttaque Or _
               Persos(i).Defense < Persos(i).Definir_MaxDefense Then
            If .Vivant Then
                'Affiche la barre de vie.
                If .Vie < .Definir_MaxVie Or _
                   .Attaque < .Definir_MaxAttaque Or _
                   .Defense < .Definir_MaxDefense Or .DureeRechargeAttaque > 0 Then
                    Interface(0).Afficher_Barre_Statut .Vie, .Definir_Vie, .Definir_MaxVie, _
                                                    .PositionX - DecalageX(NoJoueur1), _
                                                    .PositionY - DecalageY(NoJoueur1) + Persos(i).Hauteur, _
                                                    SurfBack, _
                                                    Persos(i).Largeur, _
                                                    5
                End If
                If .Attaque < .Definir_MaxAttaque Or _
                   .Defense < .Definir_MaxDefense Or .DureeRechargeAttaque > 0 Then
                   'Affiche la barre d'attaque.
                   'Durée de recharge de l'attaque.
                    If .DureeRechargeAttaque = 0 Then
                        'Barre d'attaque normale.
                        Interface(0).Afficher_Barre_Statut .Attaque, .Attaque, .Definir_MaxAttaque, _
                                                        .PositionX - DecalageX(NoJoueur1), _
                                                        .PositionY - DecalageY(NoJoueur1) + Persos(i).Hauteur + 5, _
                                                        SurfBack, _
                                                        Persos(i).Largeur, _
                                                        5, _
                                                        &HFF&          'Rouge.
                    Else
                        Interface(0).Afficher_Barre_Statut 0, .DureeRechargeAttaqueMax - .DureeRechargeAttaque, .DureeRechargeAttaqueMax, _
                                                        .PositionX - DecalageX(NoJoueur1), _
                                                        .PositionY - DecalageY(NoJoueur1) + Persos(i).Hauteur + 5, _
                                                        SurfBack, _
                                                        Persos(i).Largeur, _
                                                        5
                    End If
                    'Affiche la barre de défense.
                    Interface(0).Afficher_Barre_Statut .Defense, .Defense, .Definir_MaxDefense, _
                                                    .PositionX - DecalageX(NoJoueur1), _
                                                    .PositionY - DecalageY(NoJoueur1) + Persos(i).Hauteur + 10, _
                                                    SurfBack, _
                                                    Persos(i).Largeur, _
                                                    5, _
                                                    &HFF0000    'Bleu.
                End If
            End If
            If Persos(i).ChoisirApparence Then
                AffPerso.Afficher .PositionX - DecalageX(NoJoueur1), _
                                  .PositionY - DecalageY(NoJoueur1), _
                                  Temp, _
                                  .Direction, _
                                  .Largeur, _
                                  .Hauteur, _
                                  Persos(i), _
                                  .ApparenceChoisie, _
                                  SurfBack, _
                                  .ChoisirApparence, _
                                  Not (.Vivant), _
                                  IIf(.Definir_Volant Or Not (.Vivant), 0, Monde.ImgDecalageX(.PositionX, .PositionY + Persos(i).Hauteur - AffTerrain.Hauteur)), _
                                  IIf(.Definir_Volant Or Not (.Vivant), 0, Monde.ImgDecalageY(.PositionX, .PositionY + Persos(i).Hauteur - AffTerrain.Hauteur))
            Else
                AffApparence.Afficher .PositionX - DecalageX(NoJoueur1), _
                                      .PositionY - DecalageY(NoJoueur1), _
                                      Temp, _
                                      .Direction, _
                                      .Largeur, _
                                      .Hauteur, _
                                      Persos(i), _
                                      SurfBack, _
                                      .ChoisirApparence, _
                                      Not (.Vivant), _
                                      IIf(.Definir_Volant Or Not (.Vivant), 0, Monde.ImgDecalageX(.PositionX, .PositionY + Persos(i).Hauteur - AffTerrain.Hauteur)), _
                                      IIf(.Definir_Volant Or Not (.Vivant), 0, Monde.ImgDecalageY(.PositionX, .PositionY + Persos(i).Hauteur - AffTerrain.Hauteur))
            End If
            .NumeroPas = Temp
        'Else
        '    Persos(i).Deplacement Ressources(Persos(i).IndiceRessource), Persos(Persos(i).IndicePerso), Maisons(Persos(i).IndiceMaison), Chateaux(Persos(i).IndiceChateau), Fiefs(Persos(i).NumeroFief), NoMessage, Parametres
        End If
        End With
    Next i
End Sub

'Private Sub Afficher_Decors()
'    Dim i As Long
'    For i = 0 To UBound(Chateaux())
'        If Chateaux(i).Visible Then
'            AffChateau.Afficher Chateaux(i).PositionX - DecalageX, _
'                                Chateaux(i).PositionY - DecalageY, _
'                                Fiefs(i).TypePeuple, _
'                                1, _
'                                SurfBack
'        End If
'    Next i
'End Sub

Private Sub Afficher_Chemin(ByVal Perso As ClsJeuPerso, ByVal DecalageX As Long, ByVal DecalageY As Long)
'    Dim i As Long
'    If Perso.Visible And Perso.CheminX.Count > 0 And Perso.EnDeplacement And Not Perso.EnCombat Then
'        DecalageX = DecalageX - Perso.Largeur / 2
'        DecalageY = DecalageY - Perso.Hauteur
'        With SurfBack
'        For i = 1 To Perso.CheminX.Count
'            If i = 1 Then
'                .DrawLine Perso.PositionX - DecalageX, Perso.PositionY - DecalageY, Perso.CheminX(i) - DecalageX, Perso.CheminY(i) - DecalageY
'            Else
'                .DrawLine Perso.CheminX(i - 1) - DecalageX, Perso.CheminY(i - 1) - DecalageY, Perso.CheminX(i) - DecalageX, Perso.CheminY(i) - DecalageY
'            End If
'        Next i
'        i = i - 1
'        i = Perso.CheminX.Count
'        .DrawEllipse Perso.CheminX(i) - DecalageX + ScreenPositionX + Perso.Largeur / 12 - Perso.Largeur / 2, Perso.CheminY(i) - DecalageY + ScreenPositionY + Perso.Hauteur - Perso.Largeur / 3 - Perso.Hauteur / 8 - Perso.Hauteur, Perso.CheminX(i) + Perso.Largeur - DecalageX + ScreenPositionX - Perso.Largeur / 12 - Perso.Largeur / 2, Perso.CheminY(i) + Perso.Hauteur - DecalageY + ScreenPositionY + Perso.Largeur / 3 - Perso.Hauteur / 6 - Perso.Hauteur
'        End With
'    End If
End Sub

Private Sub Actualiser_Effets_Visuel()
    'Passe à l'étape suivante des effets visuels.
    Dim i As Long
    For i = 0 To UBound(Maisons())
        EffetsMaisons(i).Etape_Suivante
    Next i
    For i = 0 To UBound(Persos())
        EffetsPersos(i).Etape_Suivante
    Next i
    If Effets.Count > 0 Then
        For i = 1 To Effets.Count
            Effets(i).Etape_Suivante
        Next i
        For i = 1 To Effets.Count
            If Not Effets(i).Actif Then
                Effets.Remove i
                i = i - 1
            End If
            If i >= Effets.Count Then
                Exit Sub
            End If
        Next i
    End If
End Sub

Private Sub Actualiser_Effets_Interfaces()
    Dim i As Integer
    For i = 0 To NombreJoueursMO - 1
        Interface(i).Actualiser_Effets_Visuels
    Next i
End Sub

Private Sub Afficher_Effets_Visuel(ByVal NoJoueur As Integer)
    Dim i As Long
    For i = 0 To UBound(Maisons())
        With Maisons(i)
        If .Construit And .Definir_Vie < .MaxVie * 0.5 Then
            EffetsMaisons(i).Activer_Effet Parametres.EffetMaisonEnFeu, .PositionX, .PositionY, .Largeur, .Hauteur, Parametres
        End If
        End With
        With EffetsMaisons(i)
        If .Actif Then
            AffEffet.Afficher .Indice, _
                              .Etape_Effet, _
                              .PositionX - DecalageX(NoJoueur), _
                              .PositionY - DecalageY(NoJoueur), _
                              SurfBack
            If Son.Definir_Jouer And .Etape_Effet = 1 Then
                If .PositionX - DecalageX(NoJoueur) + Parametres.Effet_Hauteur(.Indice) / 2 >= 0 And _
                   .PositionX - DecalageX(NoJoueur) - Parametres.Effet_Hauteur(.Indice) / 2 <= ScreenWidth And _
                   .PositionY - DecalageY(NoJoueur) + Parametres.Effet_Hauteur(.Indice) / 2 >= 0 And _
                   .PositionY - DecalageY(NoJoueur) - Parametres.Effet_Hauteur(.Indice) / 2 <= ScreenHeight Then
                    If Not Jeu.EnPause Then Son.Lire .Indice
                End If
            End If
        End If
        End With
    Next i
    For i = 0 To UBound(Persos())
        With EffetsPersos(i)
        If .Actif Then
            AffEffet.Afficher .Indice, _
                              .Etape_Effet, _
                              .PositionX - DecalageX(NoJoueur), _
                              .PositionY - DecalageY(NoJoueur), _
                              SurfBack
            If Son.Definir_Jouer And .Etape_Effet = 1 Then
                If .PositionX - DecalageX(NoJoueur) + Parametres.Effet_Hauteur(.Indice) / 2 >= 0 And _
                   .PositionX - DecalageX(NoJoueur) - Parametres.Effet_Hauteur(.Indice) / 2 <= ScreenWidth And _
                   .PositionY - DecalageY(NoJoueur) + Parametres.Effet_Hauteur(.Indice) / 2 >= 0 And _
                   .PositionY - DecalageY(NoJoueur) - Parametres.Effet_Hauteur(.Indice) / 2 <= ScreenHeight Then
                    If Not Jeu.EnPause Then Son.Lire .Indice
                End If
            End If
        End If
        End With
    Next i
    If Effets.Count > 0 Then
        For i = 1 To Effets.Count
            With Effets(i)
            If .Actif Then
                AffEffet.Afficher .Indice, _
                                  .Etape_Effet, _
                                  .PositionX - DecalageX(NoJoueur), _
                                  .PositionY - DecalageY(NoJoueur), _
                                  SurfBack
                If Son.Definir_Jouer And .Etape_Effet = 1 Then
                    If .PositionX - DecalageX(NoJoueur) + Parametres.Effet_Hauteur(.Indice) / 2 >= 0 And _
                       .PositionX - DecalageX(NoJoueur) - Parametres.Effet_Hauteur(.Indice) / 2 <= ScreenWidth And _
                       .PositionY - DecalageY(NoJoueur) + Parametres.Effet_Hauteur(.Indice) / 2 >= 0 And _
                       .PositionY - DecalageY(NoJoueur) - Parametres.Effet_Hauteur(.Indice) / 2 <= ScreenHeight Then
                        If Not Jeu.EnPause Then Son.Lire .Indice
                    End If
                End If
            End If
            End With
        Next i
    End If
End Sub

Private Sub Afficher_CadreSelection(ByVal i As Long, _
                                    ByVal NoJoueur As Integer, _
                                    Optional ByVal CercleVide As Boolean)
    'If RetourCadre Then
    '    IndiceCadre = IndiceCadre - 8
    'Else
    '    IndiceCadre = IndiceCadre + 8
    'End If
    'If IndiceCadre <= 0 Then RetourCadre = False
    'If IndiceCadre >= 128 - 8 Then RetourCadre = True
    'Debug.Print IndiceCadre
    'SurfBack.SetForeColor IndiceCadre * 256 ^ 2 + IndiceCadre * 256 + IndiceCadre
    'Debug.Print SurfBack.GetForeColor
    
    With SurfBack
    
    If Persos(i).EnCombat Then
        .SetForeColor &HFF&
        .SetFillColor &HFF&
    ElseIf Persos(i).Inactif Then
        .SetForeColor &HFFFFFF
        .SetFillColor &HFFFFFF
    Else
        .SetForeColor &HC00000
        .SetFillColor &HC00000
    End If
    .SetFillStyle 1
    
    If Persos(i).DansSaMaison Then
        .DrawRoundedBox Maisons(i).PositionX - DecalageX(NoJoueur) + ScreenPositionX, _
                        Maisons(i).PositionY - DecalageY(NoJoueur) + ScreenPositionY, _
                        Maisons(i).PositionX + Maisons(i).Largeur - DecalageX(NoJoueur) + ScreenPositionX, _
                        Maisons(i).PositionY + Maisons(i).Hauteur - DecalageY(NoJoueur) + ScreenPositionY, 10, 10
    ElseIf Persos(i).DansUneMaison Then
        .DrawRoundedBox Maisons(Persos(i).IndiceMaison).PositionX - DecalageX(NoJoueur) + ScreenPositionX, _
                        Maisons(Persos(i).IndiceMaison).PositionY - DecalageY(NoJoueur) + ScreenPositionY, _
                        Maisons(Persos(i).IndiceMaison).PositionX + Maisons(Persos(i).IndiceMaison).Largeur - DecalageX(NoJoueur) + ScreenPositionX, _
                        Maisons(Persos(i).IndiceMaison).PositionY + Maisons(Persos(i).IndiceMaison).Hauteur - DecalageY(NoJoueur) + ScreenPositionY, 10, 10
    ElseIf Persos(i).DansSonChateau Then
        .DrawRoundedBox Chateaux(Persos(i).NumeroFief).PositionX - DecalageX(NoJoueur) + ScreenPositionX, _
                        Chateaux(Persos(i).NumeroFief).PositionY - DecalageY(NoJoueur) + ScreenPositionY, _
                        Chateaux(Persos(i).NumeroFief).PositionX + Chateaux(Persos(i).NumeroFief).Largeur - DecalageX(NoJoueur) + ScreenPositionX, _
                        Chateaux(Persos(i).NumeroFief).PositionY + Chateaux(Persos(i).NumeroFief).Hauteur - DecalageY(NoJoueur) + ScreenPositionY, 10, 10
    ElseIf Persos(i).DansUnChateau Then
        .DrawRoundedBox Chateaux(Persos(i).IndiceChateau).PositionX - DecalageX(NoJoueur) + ScreenPositionX, _
                        Chateaux(Persos(i).IndiceChateau).PositionY - DecalageY(NoJoueur) + ScreenPositionY, _
                        Chateaux(Persos(i).IndiceChateau).PositionX + Chateaux(Persos(i).IndiceChateau).Largeur - DecalageX(NoJoueur) + ScreenPositionX, _
                        Chateaux(Persos(i).IndiceChateau).PositionY + Chateaux(Persos(i).IndiceChateau).Hauteur - DecalageY(NoJoueur) + ScreenPositionY, 10, 10
    Else
        If CercleVide Then
            'Dessine un cercle.
            '.DrawCircle Persos(i).PositionX + Persos(i).Largeur / 2 - DecalageX(NoJoueur) + ScreenPositionX, _
                        Persos(i).PositionY + Persos(i).Hauteur - DecalageY(NoJoueur) + ScreenPositionY - 4, _
                        Persos(i).Largeur / 3
            If Persos(i).Monture > 0 Then
                .DrawEllipse Persos(i).PositionX - DecalageX(NoJoueur) + ScreenPositionX + Persos(i).Definir_ApparenceLargeur / 12 + Persos(i).ApparenceDecalageXTemp, Persos(i).PositionY - DecalageY(NoJoueur) + ScreenPositionY + Persos(i).Definir_ApparenceHauteur - Persos(i).Definir_ApparenceLargeur / 3 - Persos(i).Definir_ApparenceHauteur / 8 + Persos(i).ApparenceDecalageYTemp, Persos(i).PositionX + Persos(i).Definir_ApparenceLargeur - DecalageX(NoJoueur) + ScreenPositionX - Persos(i).Definir_ApparenceLargeur / 12 + Persos(i).ApparenceDecalageXTemp, Persos(i).PositionY + Persos(i).Definir_ApparenceHauteur - DecalageY(NoJoueur) + ScreenPositionY + Persos(i).Definir_ApparenceLargeur / 3 - Persos(i).Definir_ApparenceHauteur / 6 + Persos(i).ApparenceDecalageYTemp
            Else
                .DrawEllipse Persos(i).PositionX - DecalageX(NoJoueur) + ScreenPositionX + Persos(i).Largeur / 12, Persos(i).PositionY - DecalageY(NoJoueur) + ScreenPositionY + Persos(i).Hauteur - Persos(i).Largeur / 3 - Persos(i).Hauteur / 8, Persos(i).PositionX + Persos(i).Largeur - DecalageX(NoJoueur) + ScreenPositionX - Persos(i).Largeur / 12, Persos(i).PositionY + Persos(i).Hauteur - DecalageY(NoJoueur) + ScreenPositionY + Persos(i).Largeur / 3 - Persos(i).Hauteur / 6
            End If
        Else
            '.SetFillStyle 0
            .DrawEllipse Persos(i).PositionX - DecalageX(NoJoueur) + ScreenPositionX + Persos(i).Largeur / 12, Persos(i).PositionY - DecalageY(NoJoueur) + ScreenPositionY + Persos(i).Hauteur - Persos(i).Largeur / 3 - Persos(i).Hauteur / 8, Persos(i).PositionX + Persos(i).Largeur - DecalageX(NoJoueur) + ScreenPositionX - Persos(i).Largeur / 12, Persos(i).PositionY + Persos(i).Hauteur - DecalageY(NoJoueur) + ScreenPositionY + Persos(i).Largeur / 3 - Persos(i).Hauteur / 6
            .SetFillStyle 1
            'Dessine un rectangle.
            '.DrawRoundedBox Persos(i).PositionX - DecalageX(NoJoueur) + ScreenPositionX, _
                            Persos(i).PositionY - DecalageY(NoJoueur) + ScreenPositionY, _
                            Persos(i).PositionX + Persos(i).Largeur - DecalageX(NoJoueur) + ScreenPositionX, _
                            Persos(i).PositionY + Persos(i).Hauteur - DecalageY(NoJoueur) + ScreenPositionY, 10, 10
        End If
    End If
    If AfficherChemin Then Afficher_Chemin Persos(i), DecalageX(NoJoueur), DecalageY(NoJoueur)
    
    .SetForeColor 0
    End With
End Sub

Private Sub Afficher_CadreSelection_Soldats(ByVal NoJoueur As Integer)
    Dim i As Long
    SurfBack.SetFillStyle 1
    If Persos(Nopersos(NoJoueur)).Vivant Then
        If MaintenuEnfonceGauche(NoJoueur) Then
            SurfBack.SetForeColor &HFFFFFF
            SurfBack.DrawBox SourisTempX(NoJoueur), SourisTempY(NoJoueur), SourisX(NoJoueur), SourisY(NoJoueur)
        ElseIf MaintenuEnfonceDroit(NoJoueur) Then
            SurfBack.SetForeColor &HFF&
            SurfBack.DrawBox SourisTempX(NoJoueur), SourisTempY(NoJoueur), SourisX(NoJoueur), SourisY(NoJoueur)
        End If
        SelectionArmee(NoJoueur) = Abs(SourisTempX(NoJoueur) - SourisX(NoJoueur)) >= Interface(NoJoueur).RectangleSelectionMinX Or Abs(SourisTempY(NoJoueur) - SourisY(NoJoueur)) >= Interface(NoJoueur).RectangleSelectionMinY
    End If
    If ArmeeSelectionnee(NoJoueur) Then
        For i = SoldatsSelectionnes(NoJoueur).Count To 1 Step -1
            If Persos(SoldatsSelectionnes(NoJoueur).Item(i)).Vivant Then
                Afficher_CadreSelection SoldatsSelectionnes(NoJoueur).Item(i), NoJoueur, True
            Else
                SoldatsSelectionnes(NoJoueur).Remove i
            End If
        Next i
    End If
    SurfBack.SetForeColor 0
End Sub

Public Function Selectionner_Soldats(ByVal NoJoueur As Integer, _
                                     Optional ByVal Ajout As Boolean, _
                                     Optional ByVal Suppression As Boolean, _
                                     Optional PleinEcran As Boolean) As Boolean
    'Selectionne tous les soldats situées dans le rectangle.
    'PleinEcran : Si vrai, sélectionnne tout l'écran.
    Dim i As Long, j As Long, k As Long
    Dim TempX1 As Long
    Dim TempY1 As Long
    Dim TempX2 As Long
    Dim TempY2 As Long
    If Ajout And Suppression Then
        Suppression = False
    End If
    If Not (Ajout Or Suppression) Then
        While SoldatsSelectionnes(NoJoueur).Count > 0
            SoldatsSelectionnes(NoJoueur).Remove 1
        Wend
    End If
    If PleinEcran Then
        'Sélectionne toutes les unités à l'écran.
        SelectionArmee(NoJoueur) = True
        TempX1 = 0
        TempY1 = 0
        TempX2 = Ecran.Width
        TempY2 = Ecran.Height
    Else
        TempX1 = SourisX(NoJoueur)
        TempY1 = SourisY(NoJoueur)
        TempX2 = SourisTempX(NoJoueur)
        TempY2 = SourisTempY(NoJoueur)
        If TempX1 > TempX2 Then Intervertir TempX2, TempX1
        If TempY1 > TempY2 Then Intervertir TempY2, TempY1
    End If
    For i = 0 To UBound(Persos())
        With Persos(i)
        If .NumeroChef = Nopersos(NoJoueur) And _
           .Vivant And _
           Not .Incommandable Then 'certaines unités n'acceptent pas les ordres.
            If Est_Superpose(.PositionX, .PositionY, TempX1 + DecalageX(NoJoueur) - ScreenPositionX, TempY1 + DecalageY(NoJoueur) - ScreenPositionY, .Largeur, .Hauteur, TempX2 - TempX1, TempY2 - TempY1) Then
           'If Est_Superpose(.PositionX + persos(i).Largeur / 2, .PositionY + persos(i).Hauteur / 2, TempX1 + DecalageX, TempY1 + DecalageY, , , TempX2 - TempX1, TempY2 - TempY1) Then
                Selectionner_Soldats = True
                If Suppression Then
                    'Efface l'unité sélectionnée.
                    k = 0
                    For j = 1 To SoldatsSelectionnes(NoJoueur).Count
                        k = k + 1
                        If SoldatsSelectionnes(NoJoueur).Item(k) = i Then
                            SoldatsSelectionnes(NoJoueur).Remove k
                            k = k - 1
                            Cmd_Perso_Definir_IA i, True
                            '.IA = True
                        End If
                    Next j
                ElseIf Not (BoutonRelacheDroit(NoJoueur) And .ChoisirIA) Then
                    SoldatsSelectionnes(NoJoueur).Add i
                    Cmd_Perso_Definir_IA i, False
                    '.IA = False
                End If
                If Not SelectionArmee(NoJoueur) Then
                    'Ne sélectionne qu'une unité avec un clique unique.
                    Exit For
                End If
            End If
        End If
        End With
    Next i
    With Persos(Nopersos(NoJoueur))
    If BoutonRelacheGauche(NoJoueur) And _
       Est_Superpose(.PositionX, .PositionY, TempX1 + DecalageX(NoJoueur) - ScreenPositionX, TempY1 + DecalageY(NoJoueur) - ScreenPositionY, .Largeur, .Hauteur, TempX2 - TempX1, TempY2 - TempY1) Then
       'Est_Superpose(.PositionX + persos(noperso).Largeur / 2, .PositionY + persos(noperso).Hauteur / 2, TempX1 + DecalageX, TempY1 + DecalageY, , , TempX2 - TempX1, TempY2 - TempY1) Then
       Selectionner_Soldats = True
       If Suppression Then
            'Efface le joueur de la liste.
            For i = 1 To SoldatsSelectionnes(NoJoueur).Count
                If SoldatsSelectionnes(NoJoueur).Item(i) = Nopersos(NoJoueur) Then
                    SoldatsSelectionnes(NoJoueur).Remove i
                    Exit For
                End If
            Next i
        Else
            SoldatsSelectionnes(NoJoueur).Add NoPerso
        End If
    End If
    End With
    If Ajout And SoldatsSelectionnes(NoJoueur).Count > 1 Then
        'Efface les unités selectionnées en double.
        For i = 1 To SoldatsSelectionnes(NoJoueur).Count
            k = i
            For j = i + 1 To SoldatsSelectionnes(NoJoueur).Count
                k = k + 1
                If SoldatsSelectionnes(NoJoueur).Item(i) = SoldatsSelectionnes(NoJoueur).Item(k) Then
                    SoldatsSelectionnes(NoJoueur).Remove k
                    k = k - 1
                End If
            Next j
        Next i
    End If
    If SoldatsSelectionnes(NoJoueur).Count > 0 Then
        If SoldatsSelectionnes(NoJoueur).Count > 1 Then
            ArmeeSelectionnee(NoJoueur) = True
        Else
            ArmeeSelectionnee(NoJoueur) = SoldatsSelectionnes(NoJoueur).Item(1) <> NoPerso
        End If
    Else
        ArmeeSelectionnee(NoJoueur) = False
    End If
    SelectionArmee(NoJoueur) = False
End Function

'Private Function Liste_Soldats_Selectionnes() As String
'    Dim i As Long
'    Dim Liste As String
'    If SoldatsSelectionnes.Count > 0 Then
'    For i = 1 To SoldatsSelectionnes.Count
'        If i > 1 Then Liste = Liste & ", "
'        Liste = Liste & SoldatsSelectionnes.Item(i)
'    Next i
'    End If
'
'    Liste_Soldats_Selectionnes = ArmeeSelectionnee & Liste
'End Function

Private Sub Intervertir(ByRef a As Long, _
                        ByRef B As Long)
    Dim Temp As Long
    Temp = a
    a = B
    B = Temp
End Sub

Private Sub Afficher_Etat_Travail(ByVal i As Integer, ByVal NoJoueur As Integer)
    'for i=0 to ubound(persos()) 'Permet d'affcher l'état de travail de tous les joueurs en mêmme temps.
        'If Persos(i).Joueur > 0 Then
            With Persos(Nopersos(NoJoueur))
            If .DansSonChateau And Not .Vivant Then
                'Affiche le stade de resurrection.
                If .ResurrectionNombre <> 0 Then 'N'affiche pas si on n'a plus de resurrection.
                Interface(0).Afficher_Barre_Statut 100 - .DureeRessurection, 0, _
                                                100, _
                                                Chateaux(.NumeroFief).PositionX - DecalageX(NoJoueur), _
                                                Chateaux(.NumeroFief).PositionY - DecalageY(NoJoueur) - 5, _
                                                SurfBack, _
                                                Chateaux(.NumeroFief).Largeur, _
                                                5, _
                                                &HFFFFFF
                End If
            Else
                If .RessourceCourante >= 0 Then
                    Select Case .Action
                    Case 2:
                        'Affiche l'état de travail dehors.
                        Interface(0).Afficher_Barre_Statut .Definir_ressources(.RessourceCourante), 0, _
                                                        .Definir_MaxRessources(.RessourceCourante), _
                                                        .PositionX - DecalageX(NoJoueur) + IIf(Persos(i).Monture, Persos(i).Definir_ApparenceDecalageX(1), 0), _
                                                        .PositionY - DecalageY(NoJoueur) + IIf(Persos(i).Monture, Persos(i).Definir_ApparenceDecalageY(1), 0) - 5, _
                                                        SurfBack, _
                                                        Persos(i).Largeur, _
                                                        5, _
                                                        &HFFFFFF
                    Case 102:
                        'Affiche l'état de travail dans la maison.
                        Interface(0).Afficher_Barre_Statut .Definir_ressources(.RessourceCourante), 0, _
                                                        .Definir_MaxRessources(.RessourceCourante), _
                                                        Maisons(i).PositionX - DecalageX(NoJoueur), _
                                                        Maisons(i).PositionY - DecalageY(NoJoueur) - 5, _
                                                        SurfBack, _
                                                        Maisons(i).Largeur, _
                                                        5, _
                                                        &HFFFFFF
                        'Affiche les matières premières restantes.
                        If Parametres.Ressources_TypeMatierePremiere(.RessourceCourante) >= 0 Then
                            Interface(0).Afficher_Barre_Statut .Definir_ressources(Parametres.Ressources_TypeMatierePremiere(.RessourceCourante)), 0, _
                                                            .Definir_MaxRessources(Parametres.Ressources_TypeMatierePremiere(.RessourceCourante)), _
                                                            Maisons(i).PositionX - DecalageX(NoJoueur) - 5, _
                                                            Maisons(i).PositionY - DecalageY(NoJoueur), _
                                                            SurfBack, _
                                                            5, _
                                                            Maisons(i).Hauteur, _
                                                            &HFFFFFF, _
                                                            True
                        End If
                    End Select
                ElseIf .Action = 500 Then
                    'Affiche l'état de travail dans la maison.
                    Interface(0).Afficher_Barre_Statut .EtatFabricationObjet, 0, _
                                                    Parametres.Objet_DureeFabrication(Parametres.Batiment_ObjetFabrique(Maisons(i).TypeBatiment, .IndiceObjet)), _
                                                    Maisons(i).PositionX - DecalageX(NoJoueur), _
                                                    Maisons(i).PositionY - DecalageY(NoJoueur) - 5, _
                                                    SurfBack, _
                                                    Maisons(i).Largeur, _
                                                    5, _
                                                    &HFFFFFF
                End If
                'Affiche le stock de la maison.
                If .DansSaMaison And Maisons(i).MaxStock > 0 Then
                    Interface(0).Afficher_Barre_Statut Maisons(i).Definir_Stock, 0, _
                                                    Maisons(i).MaxStock, _
                                                    Maisons(i).PositionX + Maisons(i).Largeur - DecalageX(NoJoueur), _
                                                    Maisons(i).PositionY - DecalageY(NoJoueur), _
                                                    SurfBack, _
                                                    5, _
                                                    Maisons(i).Hauteur, _
                                                    &HFFFFFF, _
                                                    True
                End If
            End If
            End With
        'End If
    'Next i
End Sub

Private Sub Afficher_Surface_A_Batir(ByVal X As Integer, _
                                     ByVal Y As Integer, _
                                     ByVal TypeMaison As Integer, _
                                     ByVal EmplacementBon As Boolean)
    'Affiche le rectangle de la surface de la maison.
    
    With SurfBack
    
    If EmplacementBon Then
        .SetFillColor &HC000&
    Else
        .SetFillColor &HFF&
    End If
    .SetForeColor SurfBack.GetFillColor
    .SetFillStyle 4
    .DrawRoundedBox X - Parametres.Batiment_Largeur(TypeMaison) / 2 + AffSouris.Largeur / 2, _
                            Y - Parametres.Batiment_Hauteur(TypeMaison) / 2 + AffSouris.Hauteur / 2, _
                            X + Parametres.Batiment_Largeur(TypeMaison) / 2 + AffSouris.Largeur / 2, _
                            Y + Parametres.Batiment_Hauteur(TypeMaison) / 2 + AffSouris.Hauteur / 2, _
                            10, _
                            10
    .SetFillStyle 1
    .SetForeColor 0
    
    End With
    
End Sub

Private Sub Afficher_Commentaires(ByVal NoJoueur As Integer, _
                                  ByVal EcranWidth As Long, _
                                  ByVal EcranHeight As Long)
    Dim i As Long
    Dim X As Integer, Y As Integer
    If Commentaires.Definir_Afficher Then
        'SurfBack.SetForeColor Commentaires.Couleur
        'Affichage des commentaires.
        For i = 0 To UBound(Persos())
            'Persos(i).Commentaires = "ABC" 'Test.
            If Persos(i).Definir_Commentaires <> "" Then
                X = Persos(i).PositionX - DecalageX(NoJoueur) + Commentaires.DecalageX + ScreenPositionX
                Y = Persos(i).PositionY - DecalageY(NoJoueur) + Commentaires.DecalageY + ScreenPositionY
                If Persos(i).Vivant And _
                   X > ScreenPositionX And _
                   X < Ecran.Width And _
                   Y > ScreenPositionY And _
                   Y < Ecran.Height Then
                    AffMessage.Afficher X, _
                                        Y, _
                                        Persos(i).Definir_Commentaires, _
                                        SurfBack, _
                                        Commentaires.Couleur, True
                End If
                'On efface les commentaires affichées depuis trop longtemps.
                If Abs(TempsFin - Persos(i).TempsDernierCommentaire) > Len(Persos(i).Definir_Commentaires) * Commentaires.Duree Then
                    Persos(i).Definir_Commentaires = ""
                End If
                'Debug.Print TempsFin & "-" & Persos(i).TempsDernierCommentaire & ">" & Len(Persos(i).Definir_Commentaires) * Commentaires.Duree
            End If
        Next i
        SurfBack.SetForeColor &H0 'Remet la couleur en noir.
    End If
End Sub

Private Sub Changer_Zoom(ByVal coefzoom As Single)
    'If CoefZoom <> 1 Then
        Rect1.Left = 0
        Rect1.Right = Ecran.Width
        Rect1.Top = 0
        Rect1.Bottom = Ecran.Height
        Rect2.Left = (Ecran.Width - Ecran.Width / coefzoom) / 2
        Rect2.Right = Rect2.Left + Ecran.Width / coefzoom
        Rect2.Top = (Ecran.Height - Ecran.Height / coefzoom) / 2
        Rect2.Bottom = Rect2.Top + Ecran.Height / coefzoom
        
        'Set surfback = dd.DuplicateSurface(SurfBack)
        'Set SurfBack = dd.DuplicateSurface(SurfBack)
        'Set surfback = dd.
        'surfPrim.Blt Rect1, surfback, Rect2, DDBLT_WAIT
        'SurfBack.Blt Rect1, surfback, Rect2, DDBLT_WAIT
        'SurfBack.BltColorFill MonFond, 0&
        'surfPrim.Blt Rect1, SurfBack, Rect2, DDBLT_WAIT
        'surfPrim.Blt r1, SurfBack, r2, DDBLT_WAIT
        
    'End If
    
End Sub

Private Sub Afficher_Infos_Bulles_Terrain(ByVal Curseur As Integer, ByVal Perso As ClsJeuPerso, ByVal SourisX As Integer, ByVal SourisY As Integer, ByVal NoJoueur As Integer)
    'Affiche les infos-bulles de la souris.
    Dim i As Long
    Dim Temp As String
    If Curseur = 4 And IndiceCible(NoJoueur) <> NoPerso Then
        Curseur = 7
    End If
    
    If Interface(NoJoueur).Definir_AfficherInfoBulleTerrain Then
        Select Case Curseur
        Case 0:
            If Not CurseurSurInterface(NoJoueur) And Perso.ObjetSelectionne >= 0 Then
                If Perso.DansSonChateau Or Perso.DansUnChateau Or (Perso.DansUneMaison And Maisons(Perso.IndiceMaison).Marche) Then
                    If Parametres.Objet_Inseparable(Perso.ObjetSelectionne) Then
                        Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, Parametres.EtiquetteObjetVenteImpossible, SurfBack
                    ElseIf Perso.Vivant Then
                        Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, Parametres.EtiquetteObjetVente, SurfBack
                    End If
                Else
                    If Parametres.Objet_Inseparable(Perso.ObjetSelectionne) Then
                        Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, Parametres.EtiquetteObjetDepotImpossible, SurfBack
                    ElseIf Perso.Vivant Then
                        Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, Parametres.EtiquetteObjetDepot, SurfBack
                    End If
                End If
            End If
        Case 1: 'Curseur sur une ressource.
            Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                   Parametres.Ressources_Nom(Ressources(IndiceCible(NoJoueur)).TypeRessource) & _
                                                   " : " & Int(Persos(NoPerso).Definir_ressources(Ressources(IndiceCible(NoJoueur)).TypeRessource)) & _
                                                   " / " & Persos(NoPerso).Definir_MaxRessources(Ressources(IndiceCible(NoJoueur)).TypeRessource), _
                                                   SurfBack
        
        Case 4, 5: 'Curseur sur la maison du joueur.
            With Maisons(NoPerso)
            If .Visible Then
                If Parametres.Batiment_RessourceFabrique(.TypeBatiment) > 0 Then 'And .Magasin Then
                    Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                           Parametres.Batiment_Nom(.TypeBatiment) & _
                                                           " #" & Parametres.Ressources_Nom(Parametres.Batiment_RessourceFabrique(.TypeBatiment)) & _
                                                           " : " & Int(Persos(NoPerso).Definir_ressources(Parametres.Batiment_RessourceFabrique(.TypeBatiment))) & _
                                                           " / " & Persos(NoPerso).Definir_MaxRessources(Parametres.Batiment_RessourceFabrique(.TypeBatiment)) & " ", _
                                                           SurfBack
                ElseIf Parametres.Batiment_Fabrique(.TypeBatiment) Then
                    If Persos(NoPerso).EtatFabricationObjet > 0 Then
                    Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                                     Parametres.Batiment_Nom(.TypeBatiment) & " #" & _
                                                                     Parametres.Objet_Nom(Parametres.Batiment_ObjetFabrique(Maisons(NoPerso).TypeBatiment, Persos(NoPerso).IndiceObjet)) & " : " & _
                                                                     Int(Persos(NoPerso).EtatFabricationObjet / _
                                                                     Parametres.Objet_DureeFabrication(Parametres.Batiment_ObjetFabrique(Maisons(NoPerso).TypeBatiment, Persos(NoPerso).IndiceObjet)) _
                                                                     * 100) & "%", _
                                                                     SurfBack
                    Else
                    Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                           Parametres.Batiment_Nom(.TypeBatiment) & " ", _
                                                           SurfBack
                    End If
                Else
                    Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, Parametres.Batiment_Nom(.TypeBatiment) & "  ", SurfBack
                End If
            End If
            End With
        Case 7:  'Curseur sur une maison alliée.
            With Maisons(IndiceCible(NoJoueur))
            'If Parametres.Batiment_RessourceFabrique(.TypeBatiment) > 0 And .Magasin Then
            '    Interface(0).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                       Parametres.Batiment_Nom(.TypeBatiment) & _
                                                       "#" & Parametres.Ressources_Nom(Parametres.Batiment_RessourceFabrique(.TypeBatiment)) & _
                                                       " : " & Int(Persos(Noperso).Definir_Ressources(Parametres.Batiment_RessourceFabrique(.TypeBatiment))) & _
                                                       " / " & Persos(Noperso).Definir_MaxRessources(Parametres.Batiment_RessourceFabrique(.TypeBatiment)) & " ", _
                                                       SurfBack
            'Else
                Temp = Parametres.Batiment_Nom(.TypeBatiment)
                If Persos(.Numero).Invocation Then
                    Temp = Temp & " (" & Persos(Persos(.Numero).NumeroChef).Nom & ")"
                Else
                    Temp = Temp & " "
                End If
                Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                       Temp & _
                                                       IIf(Not .Construit Or .Definir_Stock < 1 And Not (Parametres.Batiment_Tir(.TypeBatiment) Or Parametres.Batiment_Marche(.TypeBatiment)), "#&" & IIf(Parametres.Batiment_Feminin(.TypeBatiment), Parametres.EtiquetteFermeF, Parametres.EtiquetteFermeM), "") & _
                                                       IIf(Not Persos(IndiceCible(NoJoueur)).Vivant, "#&" & IIf(Parametres.Batiment_Feminin(.TypeBatiment), Parametres.EtiquetteAbandonneF, Parametres.EtiquetteAbandonneM) & " ", ""), _
                                                       SurfBack
            'End If
            End With
        Case 6: 'Curseur le chateau.
            'Indique le pourcentage du changement d'époque.
            If Fiefs(Persos(NoPerso).NumeroFief).ChangerEpoque Then
                Temp = AffMessage.Caractere_RetourLigne & Parametres.Epoque_Nom(Fiefs(Persos(NoPerso).NumeroFief).Epoque + 1) & " : "
                Temp = Temp & Format((Parametres.Epoque_Duree(Fiefs(Persos(NoPerso).NumeroFief).Epoque + 1) * Parametres.VitesseEpoque_Coef - Fiefs(Persos(NoPerso).NumeroFief).TempsChangementEpoque) / (Parametres.Epoque_Duree(Fiefs(Persos(NoPerso).NumeroFief).Epoque + 1) * Parametres.VitesseEpoque_Coef) * 100, "0.00") & " " & Parametres.EtiquettePourcent
            End If
            'Indique le pourcentage de resurrection.
            If Persos(NoPerso).Action = 400 Then
                Temp = Temp & AffMessage.Caractere_RetourLigne & Interface(NoJoueur).Definir_MessagesActions(18) & " : " & Int(100 - Persos(NoPerso).DureeRessurection) & " " & Parametres.EtiquettePourcent
            End If
            If Persos(NoPerso).ResurrectionNombre > -1 And Not Persos(NoPerso).Vivant Then
                If Perso.ResurrectionNombre = 0 Then
                    Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                       Parametres.Peuples_NomChateau(Fiefs(Persos(NoPerso).NumeroFief).Definir_TypePeuple) & "  " & _
                                                       Temp & _
                                                       AffMessage.Caractere_RetourLigne & _
                                                       "&" & Parametres.EtiquetteVieRestante & " : " & Persos(NoPerso).ResurrectionNombre, _
                                                       SurfBack
                ElseIf Perso.ResurrectionNombre = 1 Then
                    Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                       Parametres.Peuples_NomChateau(Fiefs(Persos(NoPerso).NumeroFief).Definir_TypePeuple) & "  " & _
                                                       Temp & _
                                                       AffMessage.Caractere_RetourLigne & _
                                                       Parametres.EtiquetteVieRestante & " : " & Persos(NoPerso).ResurrectionNombre, _
                                                       SurfBack
                Else
                    Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                       Parametres.Peuples_NomChateau(Fiefs(Persos(NoPerso).NumeroFief).Definir_TypePeuple) & "  " & _
                                                       Temp & _
                                                       AffMessage.Caractere_RetourLigne & _
                                                       Parametres.EtiquetteViesRestantes & " : " & Persos(NoPerso).ResurrectionNombre, _
                                                       SurfBack
                    
                End If
            Else
                If Persos(NoPerso).Definir_ObjetSelectionne Then
                    Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                           Parametres.EtiquetteObjetVente, _
                                                           SurfBack
                Else
                    Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                           Parametres.Peuples_NomChateau(Fiefs(Persos(NoPerso).NumeroFief).Definir_TypePeuple) & "  " & Temp, _
                                                           SurfBack
                End If
            End If
        Case 8: 'Curseur sur un chateau allié.
            With Chateaux(IndiceCible(NoJoueur))
            Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                   Parametres.Peuples_NomChateau(Fiefs(IndiceCible(NoJoueur)).Definir_TypePeuple) & "  ", _
                                                   SurfBack
            End With
        Case 9: 'Curseur sur un adversaire.
            With Persos(IndiceCible(NoJoueur))
            'If .Nom = "" Then
            '    Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                       "&" & Parametres.Race_Nom(.Race, .Feminin) & "  ", _
                                                       SurfBack
            'Else
            Temp = "&" & .Nom & "  "
            If .Nombre_EffetsTemp > 0 Then
                For i = 1 To .Nombre_EffetsTemp
                    Temp = Temp & "#&(" & Parametres.Objet_NomEffet(.Definir_EffetTemp_Numero(i), .Feminin) & ")"
                Next i
            End If
            Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                   Temp, _
                                                   SurfBack
            'End If
            End With
        Case 10 'Curseur sur un batiment adverse.
            With Maisons(IndiceCible(NoJoueur))
            Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                   "&" & Parametres.Batiment_Nom(.TypeBatiment) & "  " & IIf(Not Persos(IndiceCible(NoJoueur)).Vivant, "#&" & IIf(Parametres.Batiment_Feminin(.TypeBatiment), Parametres.EtiquetteAbandonneF, Parametres.EtiquetteAbandonneM) & " ", ""), _
                                                   SurfBack
            End With
        Case 11 'Curseur sur un chateau adverse.
            With Chateaux(IndiceCible(NoJoueur))
            Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                   "&" & Parametres.Peuples_NomChateau(Fiefs(IndiceCible(NoJoueur)).Definir_TypePeuple) & "  ", _
                                                   SurfBack
            End With
        Case 12: 'Curseur sur un temple.
            Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                   Parametres.Batiment_Nom(Maisons(IndiceCible(NoJoueur)).TypeBatiment) & " ", _
                                                   SurfBack
        Case 13: 'Curseur sur un trésor.
            Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                   Tresors(IndiceCible(NoJoueur)).Nom(Parametres) & "  ", _
                                                   SurfBack
        Case 14, 17: 'Curseur sur un ami.
            With Persos(IndiceCible(NoJoueur))
            If .Nom = "" Then
                Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                       Parametres.Race_Nom(.Race, .Feminin) & "  ", _
                                                       SurfBack
            Else
                Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                       .Nom & "  ", _
                                                       SurfBack
            End If
            End With
        Case 16: 'Curseur sur une tombe.
            If Decors(IndiceCible(NoJoueur)).Nom <> "" Then
                Interface(NoJoueur).Afficher_Info_Bulles_Terrain SourisX, SourisY, _
                                                           Decors(IndiceCible(NoJoueur)).Nom & "  ", _
                                                           SurfBack
            End If
        End Select
    End If
End Sub

Public Sub Afficher_Image_Chargement(ByRef dd As DirectDraw7, _
                                     Optional ByVal Pourcentage As Integer, _
                                     Optional ByVal Message As String)
    Dim Temp As String
    If Pourcentage = 0 Then
        'Charge une image de chargement au hasard.
        Temp = ImageChargement_Charger
        ChargementImage = (Temp <> "")
        If ChargementImage Then
            ImageFS.Charger dd, Temp
        End If
    End If
    If NombreJoueursMO = 1 Then
        'Actualise la position de la souris.
        Souris.Positions di, dd, Me, SourisX(0), SourisY(0), SourisZ(0), 0, Ecran.Width, 0, Ecran.Height
    End If
    SurfBack.SetFillColor CLng(Pourcentage * 2.5 * 255) '&HFFFFFF
    SurfBack.BltColorFill MonFond, 0&     ' 4396468
    SurfBack.DrawRoundedBox 0, CLng(ScreenHeight - (97 / 768 * ScreenHeight) - ((ScreenHeight - (768 * CSng(Ecran.Width / 1024))) / 2)), CLng(ScreenWidth * Pourcentage / 100), ScreenHeight, 0, 0
    ImageFS.Selection 0, 768, 0, 1024 'Dimensions de ll'image de chargement.
    If ChargementImage Then
        ImageFS.Afficher2 0, CLng((ScreenHeight - (768 * CSng(Ecran.Width / 1024))) / 2), SurfBack, True, CSng(Ecran.Width / 1024)
    End If
    AffSouris.Afficher SourisX(0), SourisY(0), 0, SurfBack
    If Message <> "" Then Interface(0).Afficher_Info_Bulles_Terrain SourisX(0), SourisY(0), Message, SurfBack
    
    Rafraichir_Ecran
    DoEvents
    Exit Sub
End Sub

Private Function ImageChargement_Charger() As String
    'Charge une image de chargement au hasard.
    'Renvoie vrai si cette image est correctement chargée.
    On Error GoTo Erreur
    Dim Temp As String
    Dim Temp2 As Integer
    Dim listeImages As Collection
    Set listeImages = New Collection
    Temp = Dir(App.Path & ImageFS.CheminImage & "Chargement\", vbNormal)
    While Temp <> ""
        listeImages.Add "Chargement\" & Left(Temp, Len(Temp) - 4)
        Temp = Dir()
    Wend
    If listeImages.Count > 0 Then
        Temp2 = Int(Rnd * (listeImages.Count)) + 1 '- 1)) + 1
        ImageChargement_Charger = listeImages(Temp2)
    End If
    Exit Function
Erreur:
    MsgBox "Erreur dans la fonction ImageChargement_Charger (" & Temp2 & "/" & listeImages.Count & ")", vbCritical, "Erreur"
    End
End Function


Private Sub Afficher_Separateur_Ecran()
    SurfBack.DrawRoundedBox Ecran.Width / 2 - 1, 0, Ecran.Width / 2 + 1, Ecran.Height, 0, 0
    'SurfBack.SetFillColor 0
    'SurfBack.DrawBox Ecran.Width / 2, 0, Ecran.Width / 2 + 16, Ecran.Height
End Sub

Private Sub Rafraichir_Ecran()
    If Ecran.Definir_ModeFenetre Then
        'surfPrim.Blt r2, SurfBack, r2, DDBLT_WAIT
        Rect1.Left = Me.Left / 15 + 4
        Rect1.Right = Rect1.Left + Me.Width / 15 - 8 ' Ecran.Width
        Rect1.Top = Me.Top / 15 + 24 + 4
        Rect1.Bottom = Rect1.Top + Me.Height / 15 - 24 - 8 ' Ecran.Height
        Rect2.Left = 0
        Rect2.Right = Rect2.Left + Ecran.Width
        Rect2.Top = 0
        Rect2.Bottom = Rect2.Top + Ecran.Height
        surfPrim.Blt Rect1, SurfBack, Rect2, DDBLT_WAIT
        'surfPrim.BltFast 0, 0, SurfBack, Rect2, DDBLTFAST_SRCCOLORKEY
    Else
        surfPrim.Blt Rect1, SurfBack, Rect2, DDBLT_WAIT
        'surfPrim.Flip Nothing, DDFLIP_NOVSYNC
    End If
End Sub

Public Sub GigaCapture()
    Dim i As Long, j As Long, k As Long
    On Error GoTo Erreur
    For i = 0 To Monde.Largeur * Monde.LargeurCase Step Ecran.Width
        DecalageX(0) = i
        If DecalageX(0) > Monde.Largeur * AffTerrain.Largeur - Ecran.Width Then
            DecalageX(0) = Monde.Largeur * AffTerrain.Largeur - Ecran.Width
        ElseIf DecalageX(0) < 0 Then
            DecalageX(0) = 0
        End If
    
        For j = 0 To Monde.Hauteur * Monde.HauteurCase Step Ecran.Height
            DecalageY(0) = j
            If DecalageY(0) > Monde.Hauteur * AffTerrain.Hauteur - Ecran.Height Then
                DecalageY(0) = Monde.Hauteur * AffTerrain.Hauteur - Ecran.Height
            ElseIf DecalageY(0) < 0 Then
                DecalageY(0) = 0
            End If
            Debug.Print DecalageX(0) & ", " & DecalageY(0)
            If Monde.TerrainOptimise Then
                AffTerrain.AfficheVeryFast 0, 0, DecalageX(0), DecalageY(0), Ecran.Width, Ecran.Height, SurfBack
            Else
                AffTerrain.AfficheFast 0, 0, DecalageX(0), DecalageY(0), Ecran.Width, Ecran.Height, SurfBack
            End If
            Afficher_Ressources 0
            Afficher_Batiments 0
            Afficher_Decors 0
            Afficher_Tresors 0
            Afficher_Persos 0
            Afficher_Effets_Visuel 0
            Afficher_Commentaires 0, Ecran.Width, Ecran.Height
            Rafraichir_Ecran
            GrabScreen
            SavePicture Clipboard.GetData, App.Path & "\" & "Capture" & Format(k, "000") & ".bmp"
            k = k + 1
        Next j
    Next i
    Exit Sub
Erreur:
End Sub

Private Sub Diriger_Curseur(ByVal X As Integer, ByVal Y As Integer, _
                            Optional ByVal Vitesse As Integer = 16)
    Dim dx As Integer, dy As Integer, l As Integer
    dx = X - SourisX(0)
    dy = Y - SourisY(0)
    l = Sqr(dx ^ 2 + dy ^ 2)
    SourisX(0) = Vitesse * X / l + SourisX(0)
    SourisY(0) = Vitesse * Y / l + SourisY(0)
End Sub

Private Sub Form_Load()
    'If Ecran.Definir_ModeFenetre Then Me.BorderStyle = 3
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Quitter = False Then Cancel = 1
    Quitter = True
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Ecran.Definir_ModeFenetre Then
        If Button = 4 Then Button = 3
        CliqueSouris(Button - 1) = True
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Ecran.Definir_ModeFenetre Then
        If Button = 4 Then Button = 3
        CliqueSouris(Button - 1) = False
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Ecran.Definir_ModeFenetre Then
        If NombreJoueursMO > 1 Then
            SourisX(1) = X - 1
            SourisY(1) = (Y - 4) * 1.03
            If EcranDivise Then
                If SourisX(1) < ScreenWidth * 0.5 Then SourisX(1) = ScreenWidth * 0.5
            End If
        Else
            SourisX(0) = X - 1
            SourisY(0) = (Y - 4) * 1.035
        End If
        If SourisX(0) < 0 Then SourisX(0) = 0
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Restaure la résolution d'ecran d'origine
    Ecran.Femer_Ecran dd, Me
    Me.Hide
    DoEvents
    'Vide les surfaces d'affichage
    'Set dd = Nothing
    'Set di = Nothing
    'Purge quelques objets
    'Set Ecran = Nothing
    'Set Souris = Nothing
    'Set Clavier = Nothing
    'Purge la mémoire des listes de données
    PurgeMemoire
    Quitter = True
End Sub

Public Property Get Definir_ScrollingAuto(ByVal NoJoueur As Integer) As Boolean
    Definir_ScrollingAuto = ScrollingAuto(NoJoueur)
End Property
Public Property Let Definir_ScrollingAuto(ByVal NoJoueur As Integer, ByVal Valeur As Boolean)
    ScrollingAuto(NoJoueur) = Valeur
End Property
Public Property Get Definir_DecalageX(ByVal NoJoueur As Integer) As Long
    Definir_DecalageX = DecalageX(NoJoueur)
End Property
Public Property Let Definir_DecalageX(ByVal NoJoueur As Integer, ByVal Valeur As Long)
    DecalageX(NoJoueur) = Valeur - IIf(EcranDivise, Ecran.Width * 0.5, ScreenWidth) * 0.5
    If DecalageX(NoJoueur) > Monde.Largeur * AffTerrain.Largeur - IIf(EcranDivise, Ecran.Width / 2, ScreenWidth) Then
        DecalageX(NoJoueur) = Monde.Largeur * AffTerrain.Largeur - IIf(EcranDivise, Ecran.Width / 2, ScreenWidth)
    ElseIf DecalageX(NoJoueur) < 0 Then
        DecalageX(NoJoueur) = 0
    End If
End Property
Public Property Get Definir_DecalageY(ByVal NoJoueur As Integer) As Long
    Definir_DecalageY = DecalageY(NoJoueur)
End Property
Public Property Let Definir_DecalageY(ByVal NoJoueur As Integer, ByVal Valeur As Long)
    DecalageY(NoJoueur) = Valeur - ScreenHeight * 0.5
    If DecalageY(NoJoueur) > Monde.Hauteur * AffTerrain.Hauteur - ScreenHeight Then
        DecalageY(NoJoueur) = Monde.Hauteur * AffTerrain.Hauteur - ScreenHeight
    ElseIf DecalageY(NoJoueur) < 0 Then
        DecalageY(NoJoueur) = 0
    End If
End Property
