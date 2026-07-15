Attribute VB_Name = "ModMain"
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

#If TWINBASIC Then
Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Const FichierChargement = "Chargement"
Public Const SectionChargement1 = "Chargement_1"
Public Const SectionChargement2 = "Chargement_2"
Public Const SectionChargement3 = "Erreur_Chargement"

Public NavigateurChoisi As Integer
Public CheminNotePad As String
Public CheminExplorer As String
Public LogicielIExplorer As String
Public CheminIExplorer As String
Public LogicielFireFox As String
Public CheminFireFox As String

Public CheminNotePadIntrouvable As String
Public CheminExplorerIntrouvable As String
Public CheminIExplorerIntrouvable As String
Public CheminFireFoxIntrouvable As String

'Variables du jeu utilisables par chaque forme.
Public FicIni As ClsFicINI
Public FicPer As ClsFicINI

Public Langues As ClsLangues

Public Partie As ClsJeuPartie
Public Monde As ClsJeuMonde
Public Cartes As ClsJeuCartes
Public Fiefs() As ClsJeuFief
Public Persos() As ClsJeuPerso
Public Noms As ClsJeuNoms
Public EffetsPersos() As ClsJeuEffet
Public EffetsMaisons() As ClsJeuEffet
Public NoPerso As Long  'Numéro du personnage commandé par l'utilisateur.
Public Nopersos() As Long 'Numéro des personnages choisis par les joueurs.
Public Chateaux() As ClsJeuBatiment
Public Maisons() As ClsJeuBatiment
Public Ressources() As ClsJeuRessources

Public Jeu As ClsJeu
Public Commentaires As ClsJeuCommentaires
Public Messages() As ClsJeuMessages

Public Stastistiques As ClsJeuStastistiques

'Chargement des classes d'affichage.
Public dd As IDirectDraw7                     'creation objet DirectDraw
Public AffPerso As ClsAffPerso
Public AffApparence As ClsAffApparence
Public AffTerrain As ClsAffTerrain
Public AffChateau As ClsAffBatiment
Public AffMaison As ClsAffBatiment
Public AffRessources As ClsAffRessources
Public AffSouris As ClsAffSouris
Public AffEffet As ClsAffEffet
Public AffTresor As ClsAffTresors
Public AffDecor As ClsAffDecors
Public AffMessage As ClsIntMessage
Public AffDegat As ClsAffDegats
Public AffTemps As ClsAffTemps

'Chargement de la classe d'interface.
Public Interface() As ClsInterface
Public IntTouches As ClsIntTouches

'Chargement des classes de périphériques.
Public Ecran As ClsPerEcran
Public Souris As ClsPerSouris
Public Clavier As ClsPerClavier

'Classe Son.
Public Son As ClsPerSon
Public Musique As ClsPerMusique

'Classe Réseau.
Public ComReseau As ClsComReseau
Public ComMessages As ClsComMessages

'Classe paramètres
Public Parametres As ClsJeuParam

'DirectX7 est appelé une fois, et ne peut pas être déchargé!
Public dx As DirectX7                       'creation objet DirectX
'Public d3d As Direct3D7    'Création de l'objet Direct3D.
'Public d3dDevice As Direct3DDevice7
'Public dx8 As New DirectX8                       'creation objet DirectX

'Pour faire des pauses qui libère le CPU
'Declare Sub Sleep Lib "kernel32.dll" (ByVal millisecondes As Long)

'VARIABLEs CONTENANT LA RESOLUTION DE L'ECRAN EN PIXEL : a utilisé pour ini toutes les classes
Public ScreenHeight As Long
Public ScreenWidth As Long
Public ScreenPositionX As Long
Public ScreenPositionY As Long

Public Quitter As Boolean
Public Victoire() As Boolean
Public Defaite() As Boolean

Sub Main()
    Dim EtiquettesChargement() As String
    Dim i As Integer
    Dim Titre As String
    Titre = " Vendetta III - " & App.Title
    Set dx = New DirectX7
    Set FicIni = New ClsFicINI
    Set FicPer = New ClsFicINI
    Set Langues = New ClsLangues
    Set Parametres = New ClsJeuParam
    
    'Charge les etiquettes de chargement.
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierChargement
    FicIni.Section = SectionChargement1
    ReDim EtiquettesChargement(20)
    For i = 0 To UBound(EtiquettesChargement())
        EtiquettesChargement(i) = FicIni.Parametre("EtiquetteChargement" & i + 1)
    Next i
    
    Randomize
    
    DoEvents    'Laisse se charger les DLLs directx
   
    'Chargement du fond sonore
    Set Son = New ClsPerSon
    Set Musique = New ClsPerMusique
    Musique.DM_Init dx
    
    'Chargement de la fenêtre principale AVANT le reste, pour indiquer que le programme est lancé
    MDIFrmMain.Show
    'Load FrmChargement
    FrmChargement.Show
    DoEvents
    
    MDIFrmMain.Caption = Titre & EtiquettesChargement(0)
    
    'FrmAPropos2.Show
    'DoEvents
    MDIFrmMain.MousePointer = 11
   
    'Chargement des classes ...
    Set Partie = New ClsJeuPartie
    Set Monde = New ClsJeuMonde
    Set Cartes = New ClsJeuCartes
    
    MDIFrmMain.Caption = Titre & EtiquettesChargement(1)
    Set Jeu = New ClsJeu
    MDIFrmMain.Caption = Titre & EtiquettesChargement(2)
    Set Commentaires = New ClsJeuCommentaires
    ReDim Messages(0)
    Set Messages(0) = New ClsJeuMessages
    
    Set Stastistiques = New ClsJeuStastistiques
    
    MDIFrmMain.Caption = Titre & EtiquettesChargement(3)
    Set dd = dx.DirectDrawCreate("")
    'Set d3d = dd.GetDirect3D
    
    Set AffPerso = New ClsAffPerso
    Set AffApparence = New ClsAffApparence
    AffApparence.Init1
    'Set AffTerrain = New ClsAffTerrain
    Set AffChateau = New ClsAffBatiment
    Set AffMaison = New ClsAffBatiment
    Set AffRessources = New ClsAffRessources
    Set AffSouris = New ClsAffSouris
    Set AffEffet = New ClsAffEffet
    Set AffTresor = New ClsAffTresors
    Set AffDecor = New ClsAffDecors
    Set AffMessage = New ClsIntMessage
    Set AffDegat = New ClsAffDegats
    Set AffTemps = New ClsAffTemps
    
    MDIFrmMain.Caption = Titre & EtiquettesChargement(4)
    Set Ecran = New ClsPerEcran
    Set Souris = New ClsPerSouris
    Set Clavier = New ClsPerClavier
    
    MDIFrmMain.Caption = Titre & EtiquettesChargement(5)
    ReDim Interface(0)
    Set Interface(0) = New ClsInterface
    Interface(0).Init1 0, 1
    Set IntTouches = New ClsIntTouches
    
    MDIFrmMain.Caption = Titre & EtiquettesChargement(6)
    Init_Documents
    
    MDIFrmMain.Caption = Titre & EtiquettesChargement(7)
    Parametres.Init 0
    Set Noms = New ClsJeuNoms
    MDIFrmMain.Caption = Titre & EtiquettesChargement(8)
    Parametres.Init 1
    MDIFrmMain.Caption = Titre & EtiquettesChargement(9)
    Parametres.Init 2
    MDIFrmMain.Caption = Titre & EtiquettesChargement(10)
    Parametres.Init 3
    MDIFrmMain.Caption = Titre & EtiquettesChargement(11)
    Parametres.Init 4
    MDIFrmMain.Caption = Titre & EtiquettesChargement(12)
    Parametres.Init 5
    MDIFrmMain.Caption = Titre & EtiquettesChargement(13)
    Parametres.Init 6
    MDIFrmMain.Caption = Titre & EtiquettesChargement(14)
    Parametres.Init 7
    
    MDIFrmMain.Caption = Titre & EtiquettesChargement(15)
    Parametres.Init 8
    
    MDIFrmMain.Caption = Titre & EtiquettesChargement(16)
    Init_Combat

    MDIFrmMain.Caption = Titre & EtiquettesChargement(17)
    Init_Sauvegarde_Perso
    Init_Sauvegarde_Monde
    
    MDIFrmMain.Caption = Titre & EtiquettesChargement(18)
    Charger_Dossiers_IAs
    MDIFrmMain.Caption = Titre & EtiquettesChargement(19)
    
    AffApparence.Init2
    
'    Load FrmMoteur2D
'    FrmMoteur2D.Hide
'    dd.SetCooperativeLevel FrmMoteur2D.hwnd, DDSCL_NORMAL
'    AffSouris.Charger dd
'    AffPerso.Charger dd, Parametres
'    AffApparence.Charger dd
'    AffApparence.Definir_Indices_Generaux Parametres
'    'AffTerrain.Charger dd, "Plaine": AffTerrain.Largeur = Monde.LargeurCase: AffTerrain.Hauteur = Monde.HauteurCase
'    AffChateau.Charger 0, dd, Parametres
'    AffMaison.Charger 1, dd, Parametres
'    AffRessources.Charger dd
'    AffEffet.Charger dd, Parametres
'    AffTresor.Charger dd
'    AffDecor.Charger dd, Parametres
    
    Init_Commandes
    Set ComReseau = New ClsComReseau
    Set ComMessages = New ClsComMessages
    
    MDIFrmMain.Caption = Titre & EtiquettesChargement(20)
    
    'Chargement discret (invisible) de la fenêtre de paramètrage
    FrmParam.Hide: DoEvents
    'Affichage
    'FrmParam.Show
    
    MDIFrmMain.MenuPartie.Enabled = True
    MDIFrmMain.MenuMonde.Enabled = True
    MDIFrmMain.MenuPersonnage.Enabled = True
    MDIFrmMain.MenuOptions.Enabled = True
    MDIFrmMain.MenuListes.Enabled = True
    MDIFrmMain.MenuAide.Enabled = True
    
    'Unload FrmAPropos2
    
    MDIFrmMain.Caption = Titre '& " V" & App.Major & "." & App.Minor
    ''Sleep 500
    DoEvents
    'Unload FrmChargement
    
    FrmChargement.Rafraichir_Boutons
    MDIFrmMain.MousePointer = 0
    'MDIFrmMain.Picture = Nothing
    'Tatin...
End Sub

Sub Commencer()
    'chargement, génération, etc...
    Dim i As Long, j As Long, k As Long
    Dim Temp As Long
    Dim Temp2 As Boolean
    Dim Rangees As Long
    Rangees = 20
    
    'If ComReseau.Connecte Then
        'Randomize 12 'Initialise le compteur aléatoire
        'RandomizeD
    'Else
        Randomize
    'End If
    Effets_Init
    AffTemps.Init
    AffDegat.Degats_Init
    Tresors_Init
    Decors_Init
    Evenements_Initialiser
    
    Jeu.Init
    Cartes.Chargee = False
    
    If ComReseau.Connecte And Jeu.Regle_ModeConfrontation Then
        'Version bridée de serveur réseau en mode confrontation.
        Jeu.VictoireStandard = False
        Jeu.DefaiteStandard = False
        Jeu.Confrontation = True
        'Version bridée de serveur réseau en mode confrontation.
        Jeu.Definir_Mode2Joueurs = False 'Désactive le mode 2 joueurs.
        'Charge la première carte par défaut.
        Cartes.DerniereCarte = FrmParam.LstCartes.List(0)
        Cartes.Charger_Monde Monde
        Cartes.Charger_Chateaux Monde, Chateaux()
        Cartes.Charger_Ressources Monde, Ressources()
    Else
        Jeu.VictoireStandard = True
        Jeu.DefaiteStandard = True
        Jeu.Confrontation = False
    End If
    If ComReseau.Serveur And Jeu.Regle_ModeConfrontation Then
        'Version bridée de serveur réseau en mode confrontation.
        'FrmParam.LstCartes.ListIndex = 0
        'FrmParam.CmdChargerCarte_Click
        ReDim Fiefs(FrmParam.LblFiefs - 1)
        ReDim Chateaux(UBound(Fiefs()))
        For i = 0 To UBound(Fiefs())
            Set Fiefs(i) = New ClsJeuFief
            Fiefs(i).Init i, Parametres
            Set Chateaux(i) = New ClsJeuBatiment
        Next i
        Placer_Chateaux False
        Cartes.Chargee = True
        For i = 0 To UBound(Chateaux())
            Chateaux(i).Definir_Vie = 0
            Fiefs(i).Epoque = Parametres.NombreEpoques - 1
        Next i
        i = 0 'Ne met qu'un personnage au début.
        ReDim Preserve Persos(i)
        ReDim EffetsPersos(i)
        ReDim Maisons(i)
        ReDim EffetsMaisons(i)
        For i = 0 To UBound(Persos())
            Set Maisons(i) = New ClsJeuBatiment
            Set EffetsMaisons(i) = New ClsJeuEffet
            Set EffetsPersos(i) = New ClsJeuEffet
            Set Persos(i) = New ClsJeuPerso
            Persos(i).NumeroFief = i
            Persos(i).Init i, Parametres
            Maisons(i).Numero = i
        Next i
        'Charge le personnage du serveur.
        Persos(NoPerso).Joueur = ComReseau.PlayerID
        Persos(NoPerso).Init_Carac Parametres
        Charger_Sauvegarde_Perso FrmParam.LblNom, _
                                 Persos(NoPerso)
        Charger_Sauvegarde_Tresor Persos(NoPerso)
        Persos(NoPerso).NumeroEquipe = 1
        Persos(NoPerso).NumeroFief = 0
        Persos(NoPerso).PositionX = Rnd * (Monde.Largeur * Monde.LargeurCase - Persos(NoPerso).Largeur)
        Persos(NoPerso).PositionY = Rnd * (Monde.Hauteur * Monde.HauteurCase - Persos(NoPerso).Hauteur)
    ElseIf ComReseau.Client Then
        'ComMessages.ChargementPersos = FrmParam.NombrePersonnages
        'Do
            ComMessages.ChargementPersoTermine = False
            Persos(NoPerso).Changer_Race_Personnage_Equipement Parametres, Commentaires
            ComMessages.Infos_Perso_IndiceNouveau Persos(NoPerso).Nom, ComReseau.PlayerID
            Do
                ComReseau.Message_Lire
                'ComReseau.Rafraichir_Joueurs
                DoEvents
            Loop Until ComMessages.ChargementPersoTermine Or Not ComReseau.Connecte
        'Loop Until ComMessages.ChargementPersos > 0
        
        ComMessages.ChargementTermine = False
        ComMessages.Infos_Monde ComReseau.PlayerID
        Temp = Timer
        Do
            ComReseau.Message_Lire
            DoEvents
        Loop Until ComMessages.ChargementTermine = True Or Temp - Timer > 60

    Else
        Monde.Largeur = FrmParam.LblLargeurCarte
        Monde.Hauteur = FrmParam.LblHauteurCarte
        
        ReDim Ressources(FrmParam.LblNombreArbres)
        For i = 0 To UBound(Ressources())
            Set Ressources(i) = New ClsJeuRessources
        Next i
        
        Select Case FrmParam.TabStripMonde.SelectedItem.Key
        Case "Campagnes":
            With FrmParam
            Charger_Scenario .LstCampagnes.List(.LstCampagnes.ListIndex) & "\" & .LstCampagnesScenarios.List(.LstCampagnesScenarios.ListIndex), True
            End With
        Case "NouveauMonde":
            ReDim Fiefs(FrmParam.LblFiefs - 1)
            ReDim Chateaux(UBound(Fiefs()))
            For i = 0 To UBound(Fiefs())
                Set Fiefs(i) = New ClsJeuFief
                Set Chateaux(i) = New ClsJeuBatiment
                Fiefs(i).NumeroEquipe = FrmParam.ComboEquipe(i).ListIndex + 1
                If FrmParam.ComboPeuple(i).ListIndex > Parametres.NombrePeuples - 1 Then
                    'Cas où l'on a choisit un peuple aléatoire.
                    Fiefs(i).Definir_TypePeuple = Int(Rnd * Parametres.NombrePeuples)
                Else
                    Fiefs(i).Definir_TypePeuple = FrmParam.ComboPeuple(i).ListIndex
                End If
                Fiefs(i).NombreCitoyens = FrmParam.LblVillageois(i)
            Next i
            'Erreur_Inscrire "Fiefs : Ok"
            
            k = Val(FrmParam.LblVillageois(0))
            'ReDim Preserve Persos(FrmParam.LblTotalVillageois - 1)
            ReDim Persos(FrmParam.LblTotalVillageois - 1)
            ReDim EffetsPersos(FrmParam.LblTotalVillageois - 1)
            ReDim Maisons(FrmParam.LblTotalVillageois - 1)
            ReDim EffetsMaisons(FrmParam.LblTotalVillageois - 1)
            For i = 0 To UBound(Maisons())
                Set Maisons(i) = New ClsJeuBatiment
                Set EffetsMaisons(i) = New ClsJeuEffet
                Set EffetsPersos(i) = New ClsJeuEffet
                If k = 0 Then
                    j = j + 1
                    k = Val(FrmParam.LblVillageois(j))
                End If
                k = k - 1
                
                Set Persos(i) = New ClsJeuPerso
                With Persos(i)
                .NumeroFief = j
                .Init i, Parametres
                Tirer_Monstre_Aleatoire i, Fiefs(.NumeroFief).Definir_TypePeuple
                End With
                Maisons(i).Numero = i
            Next i
            'Erreur_Inscrire "Personnages : Ok"
        
            For i = 0 To UBound(Fiefs())
                Fiefs(i).Init i, Parametres
                Fiefs(i).Epoque = FrmParam.ComboEpoque.ListIndex
                For j = 0 To Parametres.NombreRessources - 1
                    If Parametres.Ressources_Epoque(j) <= Fiefs(i).Epoque Then
                        'Fiefs(i).Definir_Ressources(j) = Parametres.Epoque_ChateauxRessourcesDepart(Fiefs(i).Epoque)
                        Fiefs(i).Definir_ressources(j) = Parametres.RessourcesDeparts_ChateauxRessourcesDepart(FrmParam.ComboRessourcesDepart.ListIndex)
                    End If
                Next j
            Next i
            Partie.RessourcesDeDepart = FrmParam.ComboRessourcesDepart.ListIndex
            For i = 0 To UBound(Persos())
                Perso_Charger_Ressources_Depart Persos(i)
                Persos(i).NumeroEquipe = Fiefs(Persos(i).NumeroFief).NumeroEquipe
            Next i
        
            'Création du monde.
            If FrmParam.OptionCarte(1).Value Then
                'On charge une carte déjà établie.
                FrmParam.CmdChargerCarte_Click
                Placer_Chateaux False
                Cartes.Chargee = True
            Else
                Monde.Generer_Monde FrmParam.ComboTerrain.ListIndex
                Placer_Chateaux
                Monde.Contruire_Routes2
                Placer_Ressources
            End If
            Placer_Joueurs
            Parametres.VitesseEpoqueSelectionnee = FrmParam.ComboVitesseEpoque.ListIndex
        
        Case "NouvelleBataille":
            ReDim Fiefs(1)
            ReDim Chateaux(1)
            'ReDim Preserve Persos(Val(FrmParam.LblNombreSoldats(0)) + Val(FrmParam.LblNombreSoldats(1)) - 1)
            ReDim Persos(Val(FrmParam.LblNombreSoldats(0)) + Val(FrmParam.LblNombreSoldats(1)) - 1)
            ReDim Maisons(UBound(Persos()))
            
            For i = 0 To UBound(Fiefs())
                Set Fiefs(i) = New ClsJeuFief
                Fiefs(i).Init i, Parametres
                Fiefs(i).Epoque = FrmParam.ComboEpoque.ListIndex
                Fiefs(i).Definir_TypePeuple = 0
                Fiefs(i).NombreMorts = 0
                Fiefs(i).NombreCitoyens = Val(FrmParam.LblNombreSoldats(i))
                Fiefs(i).NumeroEquipe = i + 1
                If FrmParam.ComboArmee(i).ListIndex > Parametres.NombrePeuples - 1 Then
                    'Cas où l'on a choisit un peuple aléatoire.
                    Fiefs(i).Definir_TypePeuple = Int(Rnd * Parametres.NombrePeuples)
                Else
                    Fiefs(i).Definir_TypePeuple = FrmParam.ComboArmee(i).ListIndex
                End If
            Next i
            
            For i = 0 To UBound(Chateaux())
                Set Chateaux(i) = New ClsJeuBatiment
                Chateaux(i).MaxVie = Parametres.Peuples_ChateauMaxVie(Fiefs(i).Definir_TypePeuple)
                Chateaux(i).Definir_Vie = 0
            Next i
            For i = 0 To UBound(Ressources())
                Set Ressources(i) = New ClsJeuRessources
            Next i
            ReDim EffetsPersos(UBound(Persos()))
            ReDim EffetsMaisons(UBound(Maisons()))
            For i = 0 To UBound(Persos())
                Set Maisons(i) = New ClsJeuBatiment
                Set EffetsMaisons(i) = New ClsJeuEffet
            Next i
            For i = 0 To UBound(Persos())
                Set EffetsPersos(i) = New ClsJeuEffet
                'If i <> Noperso Then
                    Set Persos(i) = New ClsJeuPerso
                    Persos(i).Init i, Parametres
                'End If
                Maisons(i).Numero = i
                If i <= Val(FrmParam.LblNombreSoldats(0)) - 1 Then
                    Persos(i).PositionX = k * Persos(i).Largeur + 640 '32
                    Persos(i).PositionY = j * Persos(i).Hauteur + 640 '32
                    Persos(i).NumeroFief = 0
                    Persos(i).IndicePerso = Val(FrmParam.LblNombreSoldats(0)) + i
                    While Persos(i).IndicePerso >= Val(FrmParam.LblNombreSoldats(0)) + Val(FrmParam.LblNombreSoldats(1))
                        Persos(i).IndicePerso = Persos(i).IndicePerso - Val(FrmParam.LblNombreSoldats(1))
                    Wend
                Else
                    Persos(i).NumeroFief = 1
                    Persos(i).PositionX = k * Persos(i).Largeur + 1640
                    Persos(i).PositionY = j * Persos(i).Hauteur + 640 '32
                    'Persos(i).LongueurPas = Parametres.PersosVitesse
                    Persos(i).IndicePerso = i - Val(FrmParam.LblNombreSoldats(0))
                    While Persos(i).IndicePerso >= Val(FrmParam.LblNombreSoldats(0))
                        Persos(i).IndicePerso = Persos(i).IndicePerso - Val(FrmParam.LblNombreSoldats(0))
                    Wend
                End If
                If i <> NoPerso Then
                    Tirer_Monstre_Aleatoire i, Fiefs(Persos(i).NumeroFief).Definir_TypePeuple
                End If
                Persos(i).Action = 300
                Persos(i).DirectionX = Persos(i).PositionX
                Persos(i).DirectionY = Persos(i).PositionY
                'Charge une IA spécifique pour le champs de bataille.
                Persos(i).FichierIA = "Bataille.ia"
                Persos(i).ChoisirIA = True
                If i <> NoPerso Then
                    'Persos(i).IA = True
                    Persos(i).Gestion_Tous_Niveaux
                    
                    Persos(i).Ajuster_Bonus
                End If
                
                Persos(i).Definir_Attaque = Persos(i).Definir_MaxAttaque
                Persos(i).Definir_Defense = Persos(i).Definir_MaxDefense
                j = j + 1
                If j Mod Rangees = 0 Then
                    j = 0
                    k = k + 1
                End If
                Persos(i).NumeroEquipe = Fiefs(Persos(i).NumeroFief).NumeroEquipe
            Next i
            
            Persos(NoPerso).PositionX = Persos(CLng(FrmParam.LblNombreSoldats(0) / 2)).PositionX
            Persos(NoPerso).PositionY = Persos(CLng(FrmParam.LblNombreSoldats(0) / 2)).PositionY
            Temp = Persos(NoPerso).IndicePerso
            Persos(NoPerso).IndicePerso = Persos(CLng(FrmParam.LblNombreSoldats(0) / 2)).IndicePerso
            Persos(CLng(FrmParam.LblNombreSoldats(0) / 2)).PositionX = Persos(NoPerso).DirectionX
            Persos(CLng(FrmParam.LblNombreSoldats(0) / 2)).PositionY = Persos(NoPerso).DirectionY
            Persos(CLng(FrmParam.LblNombreSoldats(0) / 2)).IndicePerso = Temp
            For i = 0 To UBound(Persos())
                If Persos(i).IndicePerso = NoPerso Then
                    Persos(i).IndicePerso = CLng(FrmParam.LblNombreSoldats(0) / 2)
                ElseIf Persos(i).IndicePerso = CLng(FrmParam.LblNombreSoldats(0) / 2) Then
                    Persos(i).IndicePerso = NoPerso
                End If
                'Equipe les différentes armées.
                If FrmParam.CheckEquipementArmee(Persos(i).NumeroFief).Value = 1 Then
                    For j = 0 To FrmParam.LstObjets(Persos(i).NumeroFief).ListCount - 1
                        If FrmParam.LstObjets(Persos(i).NumeroFief).Selected(j) Then
                            For k = 0 To Parametres.NombreObjets - 1
                                If Parametres.Objet_Nom(k) = FrmParam.LstObjets(Persos(i).NumeroFief).List(j) Then
                                    Persos(i).Objet_Ajouter_Inventaire k, Parametres, Commentaires, True
                                    'Charge les munitions nécessaires.
                                    If Parametres.Objet_PerteRessourceQuantite(k) > 0 Then
                                        Persos(i).Definir_ressources(Parametres.Objet_PerteRessourceIndice(k)) = Parametres.PersosMaxRessources
                                    End If
                                End If
                            Next k
                        End If
                    Next j
                End If
            Next i
            
            If FrmParam.OptionCarte(1).Value Then
                'On charge une carte déjà établie.
                FrmParam.CmdChargerCarte_Click
            Else
                Monde.Generer_Monde FrmParam.ComboTerrain.ListIndex
            End If
            Placer_Ressources
            Parametres.VitesseEpoqueSelectionnee = FrmParam.ComboVitesseEpoque.ListIndex
        
        Case "ChargerScenario":
            Charger_Scenario FrmParam.LstMonde.List(FrmParam.LstMonde.ListIndex)
        Case "ChargerMonde":
            Charger_Sauvegarde_Monde FrmParam.FileMonde.List(FrmParam.FileMonde.ListIndex)
        End Select
        
        'Active les IA.
        'Persos(NoPerso).ChoisirIA = True
        j = 0
        For i = 0 To UBound(Persos())
            Persos(i).IA = Not Persos(i).Projectile 'Les projectiles n'ont pas d'IA.
            If Persos(i).Joueur = 0 And _
               Not Persos(i).ChoisirIA And _
               Parametres.Race_IA_Selectionner(Persos(i).Race) Then
                If NiveauIA > Rnd Then
                    Persos(i).ChoisirIA = True
                    'Tire une IA au hasard.
                    Persos(i).FichierIA = FrmParam.FileIA.List(Int(Rnd * FrmParam.FileIA.ListCount))
                End If
            End If
            Persos(i).Actualiser_Bonus
        Next i
        
        'Chargement des personnages.
        'Charger_Sauvegarde_Perso Persos(Noperso).Nom, Persos(Noperso)
        'Temp = 1
        If Jeu.Campagne Then
            'La campagne n'autorise qu'un joueur dans l'équipe 1.
            Persos(NoPerso).Joueur = ComReseau.PlayerID
            Persos(NoPerso).Init_Carac Parametres
            Charger_Sauvegarde_Perso FrmParam.LblNom, _
                                     Persos(NoPerso)
            Charger_Sauvegarde_Tresor Persos(NoPerso)
            Persos(NoPerso).NumeroEquipe = 1
            Persos(NoPerso).NumeroFief = 0
        Else
            For i = 0 To FrmParam.LstPerso.ListCount - 1
                'Charge les compagnons en fonction de leur équipe.
                If FrmParam.LstPerso.Selected(i) Then
                    Temp = Charger_Sauvegarde_Perso_Equipe(FrmParam.LstPerso.List(i)) + 1
                    Temp2 = True 'Reste vrai si le personnage n'a pas de place dans une équipe.
                    For j = 0 To UBound(Persos())
                        If Persos(j).Joueur = 0 And _
                           Persos(j).NumeroEquipe = Temp Then
                            Persos(j).Joueur = ComReseau.PlayerID
                            Persos(j).Init_Carac Parametres
                            Charger_Sauvegarde_Perso FrmParam.LstPerso.List(i), _
                                                     Persos(j)
                            Charger_Sauvegarde_Tresor Persos(j)
                            'Désactive l'IA des personnages qui n'en n'ont pas choisi.
                            'If Not Persos(j).ChoisirIA Then Persos(j).IA = False
                            Persos(j).IA = Persos(j).ChoisirIA
                            If FrmMoteur2D.NombreJoueursMO = 2 Then
                                'If Persos(j).Nom = FrmParam.LblNomJoueur(0) Then
                                If Persos(j).Nom = FrmParam.ComboNomJoueur(0).Text Then
                                    NoPerso = j
                                    Nopersos(0) = j
                                    Persos(j).ControleJoueur = True
                                'ElseIf Persos(j).Nom = FrmParam.LblNomJoueur(1) Then
                                ElseIf Persos(j).Nom = FrmParam.ComboNomJoueur(1).Text Then
                                    Nopersos(1) = j
                                    Persos(j).ControleJoueur = True
                                End If
                            Else
                                If Persos(j).Nom = FrmParam.LblNom Then
                                    NoPerso = j
                                    Nopersos(0) = j
                                    Persos(j).ControleJoueur = True
                                End If
                            End If
                            Temp2 = False
                            j = UBound(Persos())
                        End If
                    Next j
                    If Temp2 Then 'Ce joueur n'a pas trouvé de place dans une équipe.
                        MsgBox "Il n'y a pas assez de place pour tous vos personnages dans le fief " & Temp & ".", vbCritical, "Erreur"
                        Quitter = True
                        i = FrmParam.LstPerso.ListCount
                    End If
                End If
                'Evite que l'on recharge le joueur parmi les compagnons.
                'If FrmParam.LstPerso.Selected(i) And _
                   FrmParam.LstPerso.List(i) <> Persos(Noperso).Nom Then
                    'On ne charge pas plus de compagons qu'il y a de personnage dans le fief.
                    'If Persos(Noperso + Temp).NumeroFief = Persos(Noperso).NumeroFief Then
                    '    Persos(Noperso + Temp).Joueur = 1
                    '    Charger_Sauvegarde_Perso FrmParam.LstPerso.List(i), _
                                                 Persos(Noperso + Temp)
                    '    Temp = Temp + 1
                    'Else
                    '    i = FrmParam.LstPerso.ListCount
                    'End If
                'End if
            Next i
        End If
        
        'Désactive les IAs des joueurs actifs.
        For i = 0 To FrmMoteur2D.NombreJoueursMO - 1
            Persos(Nopersos(i)).IA = False
        Next i
        
        'Ajuste la difficulté.
        If FrmParam.TabStripMonde.SelectedItem.Key = "NouveauMonde" Or _
           FrmParam.TabStripMonde.SelectedItem.Key = "NouvelleBataille" Then
            For i = 0 To UBound(Chateaux())
                Chateaux(i).MaxVie = Chateaux(i).MaxVie * FrmParam.CoefVieChateaux
                If Not FrmParam.TabStripMonde.SelectedItem.Key = "NouvelleBataille" Then
                    Chateaux(i).Definir_Vie = Chateaux(i).MaxVie
                End If
            Next i
            For i = 0 To UBound(Persos())
                With Persos(i)
                If .Joueur = 0 Then
                    Temp = FrmParam.BonusXP
                    Temp2 = .GainExperience
                    .GainExperience = True 'Doit être vrai s'il on veut changer de niveau.
                    .ExpVie = Temp
                    .ExpEnergie = Temp
                    .ExpMagie = Temp
                    .ExpMoral = Temp
                    .ExpAttaque = Temp * 20
                    .ExpDefense = Temp * 20
                    'Compétences spéciales.
                    For j = 0 To Parametres.NombreCompetencesSpeciales - 1
                        .Definir_ExpSpeciales(j) = Temp
                    Next j
                    'Compétences des ressources.
                    For j = 0 To Parametres.NombreRessources - 1
                        .Definir_ExpRessources(j) = Temp
                    Next j
                    'Compétences des Services.
                    For j = 0 To Parametres.NombreServices - 2
                        .Definir_ExpServices(j) = Temp
                    Next j
                    'Compétences des objets.
                    For j = 1 To Parametres.NombreCompetencesObjets - 1
                        .Definir_ExpObjets(j) = Temp
                    Next j
                    .Gestion_Tous_Niveaux
                    .Actualiser_Bonus
                    .GainExperience = Temp2
                End If
                'Définie le nombre de resurrections autorisées.
                .ResurrectionNombre = FrmParam.Definir_Resurrections_Nombre
                .VitesseResurrection = .VitesseResurrection * FrmParam.Definir_Resurrections_Vitesse
                FrmMoteur2D.Afficher_Image_Chargement dd, 40, Parametres.Etiquette_Chargement(4) & Format(i / UBound(Persos()) * 100, "00") & " " & Parametres.EtiquettePourcent
                End With
            Next i
        End If
    End If
    
    'Prépare le réseau.
    'If ComReseau.Connecte Then
    '    Enregistrer_Sauvegarde_Monde "Monde_Reseau"
    '    ComReseau.Fichier_Envoyer App.Path & CheminSavMonde & "Monde_Reseau" & ExtensionFichiersMonde
    '    Monde_Supprimer "Monde_Reseau"
    '    ComReseau.Message_Envoyer "Pret"
    '    Noperso = ComReseau.NumeroJoueur
    'End If
    
    'Erreur_Inscrire "Compagnons : Ok"
End Sub

Sub PurgeMemoire()
    Erase Maisons()
    Erase Chateaux()
    Erase Fiefs()
    Erase Ressources()
End Sub

Public Sub Editer_Texte(ByVal Chemin As String)
    On Error GoTo Erreur
    Dim RetVal As Double
    RetVal = Shell(CheminNotePad & " " & Chemin, 1)
    Exit Sub
Erreur:
    MsgBox CheminNotePadIntrouvable, vbCritical
End Sub

Public Sub Editer_Repertoire(ByVal Chemin As String)
    On Error GoTo Erreur
    Dim RetVal As Double
    RetVal = Shell(CheminExplorer & " " & Chemin, 1)
    Exit Sub
Erreur:
    MsgBox CheminExplorerIntrouvable, vbCritical
End Sub

Public Sub Ouvir_PageWeb(ByVal Chemin As String)
'    On Error GoTo Erreur
'    Dim RetVal As Double
'    Select Case NavigateurChoisi
'    Case 0:
'        RetVal = Shell(CheminIExplorer & " " & Chemin, vbMaximizedFocus)
'    Case 1:
'        RetVal = Shell(CheminFireFox & " " & Chr(34) & Chemin & Chr(34), vbMaximizedFocus)
'    End Select
'    Exit Sub
'Erreur:
'    MsgBox CheminIExplorerIntrouvable, vbCritical
#If TWINBASIC Then
    Dim hwnd As LongPtr
#Else
    Dim hwnd As Long
#End If
    Dim lpOperation As String
    Dim lpFile As String
    Dim lpParameters As String
    Dim lpDirectory As String
    Dim nShowCmd As Long

    ShellExecute hwnd, lpOperation, Chemin, lpParameters, lpDirectory, nShowCmd
End Sub

Public Sub Erreur_Inscrire(ByVal Message As String, Optional ByVal Debut As Boolean)
    If Debut Then
        Open App.Path & "\Erreurs.txt" For Output As #1
    Else
        Open App.Path & "\Erreurs.txt" For Append As #1
    End If
    Print #1, Message
    Close #1
End Sub

Public Sub Collection_Vider(ByRef Tableau As Collection)
    While Tableau.Count > 0
        Tableau.Remove 1
    Wend
End Sub

Public Function Collection_NombreElements(ByVal Element As String, _
                                          ByVal Tableau As Collection) As Integer
    'Renvoie le nombre d'éléments identitiques à celui passé en paramètres.
    Dim i As Integer
    If Tableau.Count > 0 Then
        For i = 1 To Tableau.Count
            If Tableau.Item(i) = Element Then
                Collection_NombreElements = Collection_NombreElements + 1
            End If
        Next i
    End If
End Function
