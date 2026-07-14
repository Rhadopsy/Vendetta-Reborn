VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsComMessages"
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

Const Categorie1MessageChat = 0
Const Categorie1Infos = 1
Const Categorie1Partie = 2
Const Categorie1Monde = 3
Const Categorie1Ressource = 4
Const Categorie1Perso = 5
Const Categorie1Maison = 6
Const Categorie1Chateau = 7
Const Categorie1Fief = 8
Const Categorie1Tresor = 9
Const Categorie1Decor = 10
Const Categorie1Degat = 11
Const Categorie1EffetsVisuels = 12
Const Categorie1Commandes = 13

Const Categorie2InfosPartie = 0
Const Categorie2InfosPartieEnCours = 1
Const Categorie2InfosMonde = 2
Const Categorie2InfosMondeEnCours = 3
Const Categorie2InfosPersoCaracs = 4
Const Categorie2InfosPersoCaracs2 = 5
Const Categorie2InfosPersoListeRessources = 6
Const Categorie2InfosPersoInventaire = 7
Const Categorie2InfosPersoInventaireMaison = 8
Const Categorie2InfosPersoCompetences = 9
Const Categorie2InfosPersoCompetences2 = 16
Const Categorie2InfosPersoListeBestiaire = 10
Const Categorie2InfosMaison = 11
Const Categorie2InfosChateau = 12
Const Categorie2InfosPersoIndiceNouveau = 13
Const Categorie2InfosPersoIndiceSupprimer = 14
Const Categorie2InfosPersoEstTonJoueur = 15
Const Categorie2InfosPersoChargementTermine = 17

Const Categorie2PartieMode = 0
Const Categorie2PartieEnCours = 1

Const Categorie2MondeDimensions = 0
Const Categorie2MondeCarte = 1
Const Categorie2MondeTerrain = 2
Const Categorie2MondeTerrainFin = 3
Const Categorie2MondeNomImageCarte = 4

Const Categorie2RessourceNombre = 0
Const Categorie2RessourceInfos = 1

Const Categorie2PersoNombre = 0
Const Categorie2PersoNom = 1
Const Categorie2PersoFief = 2
Const Categorie2PersoRace = 3
Const Categorie2PersoPositions = 4
Const Categorie2PersoVie = 5
Const Categorie2PersoVie2 = 44
Const Categorie2PersoVie3 = 45
Const Categorie2PersoVieExp = 59
Const Categorie2PersoEnergie = 6
Const Categorie2PersoEnergie2 = 46
Const Categorie2PersoEnergie3 = 47
Const Categorie2PersoEnergieExp = 60
Const Categorie2PersoMagie = 7
Const Categorie2PersoMagie2 = 48
Const Categorie2PersoMagie3 = 49
Const Categorie2PersoMagieExp = 61
Const Categorie2PersoMoral = 8
Const Categorie2PersoMoral2 = 50
Const Categorie2PersoMoral3 = 51
Const Categorie2PersoMoralExp = 62
Const Categorie2PersoAttaque = 9
Const Categorie2PersoAttaque2 = 52
Const Categorie2PersoAttaqueRecharge = 96
Const Categorie2PersoAttaqueExp = 63
Const Categorie2PersoDefense = 10
Const Categorie2PersoDefense2 = 53
Const Categorie2PersoDefenseExp = 64
Const Categorie2PersoJoueur = 11
Const Categorie2PersoEquipe = 12
Const Categorie2PersoIA = 95
'Const Categorie2PersoPas = 13
Const Categorie2PersoDirection = 14
Const Categorie2PersoFeminin = 15
'Const Categorie2PersoVivant = 16
Const Categorie2PersoChoisirApparence = 17
Const Categorie2PersoFichierApparence = 27
Const Categorie2PersoApparenceChoisie = 18
Const Categorie2PersoNombre_ApparenceBase = 19
Const Categorie2PersoDefinir_ApparenceBase = 20
Const Categorie2PersoDefinir_ApparenceDecalage = 93
Const Categorie2PersoDefinir_ApparenceAvantPlan = 94
Const Categorie2PersoDefinir_FichierApparenceBase = 43
Const Categorie2PersoApparenceObjetAfficher = 21
Const Categorie2PersoNombreApparenceObjet = 22
Const Categorie2PersoDefinir_ApparenceObjet = 23
Const Categorie2PersoDansUneMaison = 24
Const Categorie2PersoDansUnChateau = 25
Const Categorie2PersoChangerNombre = 26
Const Categorie2PersoChoisirIA = 28
Const Categorie2PersoChoisirIAScript = 29
Const Categorie2PersoNumeroChef = 80
Const Categorie2PersoNombreSoldats = 81
Const Categorie2PersoAttitudeFormation = 30
Const Categorie2PersoExpVie = 31
Const Categorie2PersoExpEnergie = 32
Const Categorie2PersoExpMagie = 33
Const Categorie2PersoExpMoral = 34
Const Categorie2PersoExpAttaque = 35
Const Categorie2PersoExpDefense = 36
Const Categorie2PersoCompetencesSpeciales = 68
Const Categorie2PersoCompetencesRessources = 69
Const Categorie2PersoCompetencesServices = 70
Const Categorie2PersoCompetencesObjets = 71
Const Categorie2PersoCompetencesSpecialesExp = 37
Const Categorie2PersoCompetencesRessourcesExp = 38
Const Categorie2PersoCompetencesServicesExp = 39
Const Categorie2PersoCompetencesObjetsExp = 40
Const Categorie2PersoBestiaire = 41
Const Categorie2PersoAction = 54
Const Categorie2PersoEtatFabricationObjet = 79
Const Categorie2PersoCibles1 = 55
Const Categorie2PersoCibles2 = 56
Const Categorie2PersoCibles3 = 57
Const Categorie2PersoArgent = 76
Const Categorie2PersoRessources = 58
Const Categorie2PersoRessourcesSelectionnees = 83
Const Categorie2PersoObjetSelectionne = 82
Const Categorie2PersoObjetsEquipes = 77
Const Categorie2PersoObjetsInventaire = 78
Const Categorie2PersoArmure = 65
Const Categorie2PersoVitalite = 66
Const Categorie2PersoVitesse = 67
Const Categorie2PersoEffetTempNombre = 72
Const Categorie2PersoEffetTempNumero = 73
Const Categorie2PersoEffetTempDuree = 90
Const Categorie2PersoEffetTempProprietaire = 74
Const Categorie2PersoEffetTempAjouter = 86
Const Categorie2PersoEffetTempSupprimer = 87
Const Categorie2PersoEffetRecharge = 75
Const Categorie2PersoEffetRechargeAjouter = 91
Const Categorie2PersoEffetRechargeSupprimer = 92
Const Categorie2PersoCommentaires = 84
Const Categorie2PersoDureeResurrection = 85
Const Categorie2PersoActualiser = 88
Const Categorie2PersoActualiserBonus = 89
Const Categorie2PersoEnregistrer = 97

Const Categorie2MaisonType = 0
Const Categorie2MaisonEnConstruction = 1
Const Categorie2MaisonPositions = 2
Const Categorie2MaisonVie = 3
Const Categorie2MaisonStock = 4
Const Categorie2MaisonStockObjet = 5
Const Categorie2MaisonInventaire = 6

Const Categorie2ChateauPositions = 0
Const Categorie2ChateauVie = 1
Const Categorie2ChateauRecuperationVie = 2
Const Categorie2ChateauRessources = 3

Const Categorie2FiefNombre = 0
Const Categorie2FiefPeuple = 1
Const Categorie2FiefEpoque = 2
Const Categorie2FiefEquipe = 3
Const Categorie2FiefTempsChangementEpoque = 4
Const Categorie2FiefNombreCitoyens = 5
Const Categorie2FiefEmpereur = 6

Const Categorie2TresorNombre = 0
Const Categorie2TresorApparence = 1
Const Categorie2TresorPositions = 2
Const Categorie2TresorAjouter = 3
Const Categorie2TresorSupprimer = 4

Const Categorie2DecorNombre = 0
Const Categorie2DecorNom = 1
Const Categorie2DecorApparence = 2
Const Categorie2DecorPositions = 3
Const Categorie2DecorAjouter = 4
Const Categorie2DecorSupprimer = 5

Const Categorie2DegatTexte = 0
Const Categorie2DegatPositions = 1
Const Categorie2DegatVisuel = 2

Const Categorie2EffetApparence = 0
Const Categorie2EffetPositions = 1
'Const Categorie2EffetApparencePerso = 2
'Const Categorie2EffetPositionsPerso = 3
'Const Categorie2EffetApparenceMaison = 4
'Const Categorie2EffetPositionsMaison = 5

Public ChargementTermine As Boolean
Public ChargementPersoTermine As Boolean 'Vrai, le personnage a fini d'être envoyé.
Public ChargementPersos As Integer 'Indique le nombre de personnages qu'il reste encore à charger.
Public ChargementIndicePerso As Long 'Indique l'indice du perso dans lequel on peut charger.

Private NombreComPersos As Long
Private ComPersos() As ClsComPerso
Private NombreComFiefs As Long
Private ComFiefs() As ClsComFief
'Private NombreComTresors As Long
'Private ComTresors As Collection
'Private NombreComDecors As Long
'Private ComDecors As Collection

Private Sub Class_Initialize()
    ReDim ComPersos(0)
    Set ComPersos(0) = New ClsComPerso
    ComPersos(0).Init Parametres
    NombreComPersos = 1
    ReDim ComFiefs(0)
    Set ComFiefs(0) = New ClsComFief
    ComFiefs(0).Init Parametres
    NombreComFiefs = 1
'    Set ComTresors = New Collection
'    NombreComTresors = 0
'    Set ComDecors = New Collection
'    NombreComDecors = 0
End Sub

Public Sub MessageChat(ByVal Message As String, _
                       Optional ByVal NomJoueur As String)
    ComReseau.Message_Envoyer Categorie1MessageChat, 0, 0, 0, Message, NomJoueur, , True
End Sub
Public Sub Commandes(ByVal Commande As ClsCommande)
    With Commande
    ComReseau.Message_Envoyer Categorie1Commandes, .Categorie, .Indice, , .Var1, .Var2
    End With
End Sub

' ************************************************************************************
' *                     Liste des messages demandant des informations                *
' ************************************************************************************
Public Sub Infos_Partie(ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Infos, Categorie2InfosPartie, , IDJoueur
End Sub
Public Sub Infos_PartieEnCours(ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Infos, Categorie2InfosPartieEnCours, , IDJoueur
End Sub
Public Sub Infos_Monde(ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Infos, Categorie2InfosMonde, , IDJoueur
End Sub
Public Sub Infos_Monde_Encours(ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Infos, Categorie2InfosMondeEnCours, , IDJoueur
End Sub
Public Sub Infos_Perso_Caracs(ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Infos, Categorie2InfosPersoCaracs, Indice, IDJoueur
End Sub
Public Sub Infos_Perso_Caracs2(ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Infos, Categorie2InfosPersoCaracs2, Indice, IDJoueur
End Sub
Public Sub Infos_Perso_Ressources(ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Infos, Categorie2InfosPersoListeRessources, Indice, IDJoueur
End Sub
Public Sub Infos_Perso_Inventaire(ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Infos, Categorie2InfosPersoInventaire, Indice, IDJoueur
End Sub
Public Sub Infos_Perso_InventaireMaison(ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Infos, Categorie2InfosPersoInventaireMaison, Indice, IDJoueur
End Sub
Public Sub Infos_Perso_Competences(ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Infos, Categorie2InfosPersoCompetences, Indice, IDJoueur
End Sub
Public Sub Infos_Perso_Competences2(ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Infos, Categorie2InfosPersoCompetences2, Indice, IDJoueur
End Sub
Public Sub Infos_Perso_Bestiaire(ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Infos, Categorie2InfosPersoListeBestiaire, Indice, IDJoueur
End Sub
Public Sub Infos_Maison(ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Infos, Categorie2InfosMaison, Indice, IDJoueur
End Sub
Public Sub Infos_Chateau(ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Infos, Categorie2InfosChateau, Indice, IDJoueur
End Sub
Public Sub Infos_Perso_IndiceNouveau(ByVal NomPerso As String, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Infos, Categorie2InfosPersoIndiceNouveau, IDJoueur, , NomPerso
End Sub
Public Sub Infos_Perso_IndiceSupprimer(ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Infos, Categorie2InfosPersoIndiceSupprimer, IDJoueur
End Sub
Public Sub Infos_Perso_EstTonJoueur(ByVal Indice As Long, ByVal NomPerso As String, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Infos, Categorie2InfosPersoEstTonJoueur, Indice, , NomPerso, , IDJoueur
End Sub
Public Sub Infos_Perso_ChargementTermine(ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Infos, Categorie2InfosPersoChargementTermine, Indice, IDJoueur
End Sub

' ************************************************************************************
' *                       Liste des messages concernant la partie                    *
' ************************************************************************************
Public Sub Partie_(ByVal IDJoueur As Long)
    Partie_Mode IDJoueur
    Partie_EnCours IDJoueur
End Sub
Public Sub Partie_Mode(ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Partie, Categorie2PartieMode, 0, 0, Partie.IndiceMode, , IDJoueur
End Sub
Public Sub Partie_EnCours(ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Partie, Categorie2PartieEnCours, 0, 0, Partie.EnCours, , IDJoueur
End Sub

' ************************************************************************************
' *                       Liste des messages concernant du monde                     *
' ************************************************************************************
Public Sub Monde_Debut(ByVal IDJoueur As Long)
    Dim i As Long
    Dim j As Long
    'Envoie les informations de la carte.
    Monde_Dimensions IDJoueur
    
    'Envoie les informations des ressources.
    Ressource_Nombre IDJoueur
    For i = 0 To UBound(Ressources())
        Ressource_Infos i, IDJoueur
    Next i
    
    'Envoie les informations des personnages.
    Perso_Nombre IDJoueur
    For i = 0 To UBound(Persos())
        With Persos(i)
        Perso_Fief Persos(i), i, IDJoueur
        Perso_Race Persos(i), i, IDJoueur
        Perso_IA Persos(i), i, IDJoueur
        Perso_Joueur Persos(i), i, IDJoueur
        
        Perso_Positions Persos(i), i, IDJoueur
        Perso_Nom Persos(i), i, IDJoueur
        Perso_Equipe Persos(i), i, IDJoueur
        Perso_Vie Persos(i), i, IDJoueur
        Perso_Vie2 Persos(i), i, IDJoueur
        Perso_Attaque Persos(i), i, IDJoueur
        Perso_AttaqueRecharge Persos(i), i, IDJoueur
        Perso_Defense Persos(i), i, IDJoueur
        Perso_Direction Persos(i), i, IDJoueur
        Perso_Feminin Persos(i), i, IDJoueur
        Perso_ChoisirApparence Persos(i), i, IDJoueur
        Perso_ApparenceChoisie Persos(i), i, IDJoueur
        'Perso_Commentaires Persos(i), i, IDJoueur
        Perso_NumeroChef Persos(i), i, IDJoueur
        Perso_Nombre_ApparenceBase Persos(i), i, IDJoueur
        For j = 0 To .Nombre_ApparenceBase - 1
            Perso_Definir_ApparenceBase Persos(i), i, j, IDJoueur
        Next j
        For j = 0 To 4
            Perso_Definir_ApparenceDecalage Persos(i), i, j, IDJoueur
            Perso_Definir_ApparenceAvantPlan Persos(i), i, j, IDJoueur
        Next j
        'Envoie l'affichage des apparences de l'équipement.
        Perso_ApparenceObjetAfficher Persos(i), i, IDJoueur
        Perso_NombreApparenceObjet Persos(i), i, IDJoueur
        For j = 1 To .Nombre_ApparenceObjet
            Perso_Definir_ApparenceObjet Persos(i), i, j, IDJoueur
        Next j
        Perso_DansUneMaison Persos(i), i, IDJoueur
        Perso_DansUnChateau Persos(i), i, IDJoueur
        Maison_Vie i, IDJoueur
        Maison_Type i, IDJoueur
        Maison_EnConstruction i, IDJoueur
        Maison_Stock i, IDJoueur
        
        'Envoie les informations des effets temporaires.
        For j = 1 To .Nombre_EffetsTemp
            Perso_EffetTempAjouter i, .Definir_EffetTemp_Numero(j), .Definir_EffetTemp_Duree(j), .Definir_EffetTemp_DureeMax(j), .Definir_EffetTemp_Proprietaire(j), IDJoueur
        Next j
        For j = 1 To .Nombre_EffetsRecharge
            Perso_EffetRechargeAjouter i, .Definir_EffetsRecharge_Numero(j), .Definir_EffetsRecharge_Duree(j), IDJoueur
        Next j
        End With
    Next i
    
    'Envoie les informations des fiefs.
    Fief_Nombre IDJoueur
    For i = 0 To UBound(Fiefs())
        Chateau_Positions i, IDJoueur
        Fief_Peuple i, IDJoueur
        
        Chateau_Vie i, IDJoueur
        Chateau_RecuperationVie i, IDJoueur
        Fief_NombreCitoyens i, IDJoueur
        Fief_Empereur i, IDJoueur
        Fief_Epoque i, IDJoueur
        Fief_Equipe i, IDJoueur
        Fief_TempsChangementEpoque i, IDJoueur
    Next i
    
    'Envoie les informations des trésors.
    For i = 1 To Tresors.Count
        Tresor_Ajouter Tresors(i).X, Tresors(i).Y, IDJoueur
        Tresor_Apparence i, IDJoueur
    Next i

    'Envoie les informations des décors.
    For i = 1 To Decors.Count
        Decor_Ajouter Decors(i).Categorie, Decors(i).PositionX, Decors(i).PositionY, Decors(i).Nom, IDJoueur
    Next i
    
    'Envoie la carte.
    If Not Jeu.Regle_ModeConfrontation Then
        For i = 0 To Monde.NombreCases - 1
            Monde_Terrain i, IDJoueur
        Next i
        'Monde_Carte IDJoueur
    End If
    
    'Finalise.
    Monde_NomImageCarte IDJoueur
    Monde_TerrainFin IDJoueur
    'Monde_EnCours IDJoueur
End Sub
Public Sub Monde_EnCours(Optional ByVal IDJoueur As Long, _
                         Optional ByVal ToutJoueurs As Boolean)
    Dim i As Long, j As Long, Temp As Boolean
    'Vérifie si tous les joueurs sont prêts à recevoir le message.
    Temp = True
    For i = 1 To ComReseau.lngNumJoueurs - 1
        If Not ComReseau.Joueur_InfosAJour(ComReseau.Joueur_ID(i)) Then
            Temp = False
            Exit For
        End If
    Next i
    
    If Temp Then
        'Envoie les informations des personnages.
        If NombreComPersos <> UBound(Persos()) + 1 Then
            Perso_ChangerNombre IDJoueur
            NombreComPersos = UBound(Persos()) + 1
        End If
        'Actualiser_Memoire
        For i = 0 To UBound(Persos())
            With Persos(i)
            If ComPersos(i).Race <> Persos(i).Race Then
                Perso_Race Persos(i), i, IDJoueur, ToutJoueurs
                ComPersos(i).Race = Persos(i).Race
            End If
            If ComPersos(i).PositionX <> Persos(i).PositionX Or ComPersos(i).PositionY <> Persos(i).PositionY Then
                Perso_Positions Persos(i), i, IDJoueur, ToutJoueurs
                ComPersos(i).PositionX = Persos(i).PositionX: ComPersos(i).PositionY = Persos(i).PositionY
            End If
            If ComPersos(i).Nom <> Persos(i).Nom Then
                Perso_Nom Persos(i), i, IDJoueur, ToutJoueurs
                ComPersos(i).Nom = Persos(i).Nom
            End If
            If ComPersos(i).NumeroEquipe <> Persos(i).NumeroEquipe Then
                Perso_Equipe Persos(i), i, IDJoueur, ToutJoueurs
                ComPersos(i).NumeroEquipe = Persos(i).NumeroEquipe
            End If
            If Persos(i).Visible Then
                If ComPersos(i).Vie <> Persos(i).Vie Or ComPersos(i).BonusMaxVie <> Persos(i).BonusMaxVie Then
                    Perso_Vie Persos(i), i, IDJoueur, ToutJoueurs
                    ComPersos(i).Vie = Persos(i).Vie: ComPersos(i).BonusMaxVie = Persos(i).BonusMaxVie
                End If
                If ComPersos(i).RecuperationVie <> Persos(i).RecuperationVie Then
                    Perso_Vie2 Persos(i), i, IDJoueur, ToutJoueurs
                    ComPersos(i).RecuperationVie = Persos(i).RecuperationVie
                End If
                If ComPersos(i).Attaque <> Persos(i).Attaque Or ComPersos(i).BonusMaxAttaque <> Persos(i).BonusMaxAttaque Then
                    Perso_Attaque Persos(i), i, IDJoueur, ToutJoueurs
                    ComPersos(i).Attaque = Persos(i).Attaque: ComPersos(i).BonusMaxAttaque = Persos(i).BonusMaxAttaque
                End If
                If ComPersos(i).DureeRechargeAttaque <> Persos(i).DureeRechargeAttaque Or ComPersos(i).DureeRechargeAttaqueMax <> Persos(i).DureeRechargeAttaqueMax Then
                    Perso_AttaqueRecharge Persos(i), i, IDJoueur, ToutJoueurs
                    ComPersos(i).DureeRechargeAttaque = Persos(i).DureeRechargeAttaque: ComPersos(i).DureeRechargeAttaqueMax = Persos(i).DureeRechargeAttaqueMax
                End If
                If ComPersos(i).Defense <> Persos(i).Defense Or ComPersos(i).BonusMaxDefense <> Persos(i).BonusMaxDefense Then
                    Perso_Defense Persos(i), i, IDJoueur, ToutJoueurs
                    ComPersos(i).Defense = Persos(i).Defense: ComPersos(i).BonusMaxDefense = Persos(i).BonusMaxDefense
                End If
                If ComPersos(i).Direction <> Persos(i).Direction Then
                    Perso_Direction Persos(i), i, IDJoueur, ToutJoueurs
                    ComPersos(i).Direction = Persos(i).Direction
                End If
                If ComPersos(i).Feminin <> Persos(i).Feminin Then
                    Perso_Feminin Persos(i), i, IDJoueur, ToutJoueurs
                    ComPersos(i).Feminin = Persos(i).Feminin
                End If
                If ComPersos(i).ChoisirApparence <> Persos(i).ChoisirApparence Then
                    Perso_ChoisirApparence Persos(i), i, IDJoueur, ToutJoueurs
                    ComPersos(i).ChoisirApparence = Persos(i).ChoisirApparence
                End If
                If ComPersos(i).ApparenceChoisie <> Persos(i).ApparenceChoisie Then
                    Perso_ApparenceChoisie Persos(i), i, IDJoueur, ToutJoueurs
                    ComPersos(i).ApparenceChoisie = Persos(i).ApparenceChoisie
                End If
                If ComPersos(i).Commentaires <> Persos(i).Definir_Commentaires Then
                    Perso_Commentaires Persos(i), i, IDJoueur, ToutJoueurs
                    ComPersos(i).Commentaires = Persos(i).Definir_Commentaires
                End If
                If ComPersos(i).NumeroChef <> Persos(i).NumeroChef Then
                    Perso_NumeroChef Persos(i), i, IDJoueur, ToutJoueurs
                    ComPersos(i).NumeroChef = Persos(i).NumeroChef
                End If
                If ComPersos(i).Definir_NombreApparenceBase <> Persos(i).Nombre_ApparenceBase Then
                    Perso_Nombre_ApparenceBase Persos(i), i, IDJoueur, ToutJoueurs
                    ComPersos(i).Definir_NombreApparenceBase = Persos(i).Nombre_ApparenceBase
                End If
                For j = 0 To .Nombre_ApparenceBase - 1
                    If ComPersos(i).Definir_ApparenceBase(j) <> Persos(i).Definir_ApparenceBase(j) Then
                        Perso_Definir_ApparenceBase Persos(i), i, j, IDJoueur, ToutJoueurs
                        ComPersos(i).Definir_ApparenceBase(j) = Persos(i).Definir_ApparenceBase(j)
                    End If
                Next j
                For j = 0 To 4
                    If ComPersos(i).Definir_ApparenceDecalageX(j) <> Persos(i).Definir_ApparenceDecalageX(j) Or _
                       ComPersos(i).Definir_ApparenceDecalageY(j) <> Persos(i).Definir_ApparenceDecalageYFixe(j) Then
                        Perso_Definir_ApparenceDecalage Persos(i), i, j, IDJoueur, ToutJoueurs
                        ComPersos(i).Definir_ApparenceDecalageX(j) = Persos(i).Definir_ApparenceDecalageX(j)
                        ComPersos(i).Definir_ApparenceDecalageY(j) = Persos(i).Definir_ApparenceDecalageYFixe(j)
                    End If
                    If ComPersos(i).Definir_ApparenceAvantPlan(j) <> Persos(i).Definir_ApparenceAvantPlan(j) Then
                        Perso_Definir_ApparenceAvantPlan Persos(i), i, j, IDJoueur, ToutJoueurs
                        ComPersos(i).Definir_ApparenceAvantPlan(j) = Persos(i).Definir_ApparenceAvantPlan(j)
                    End If
                Next j
                'Envoie l'affichage des apparences de l'équipement.
                If ComPersos(i).ApparenceObjetAfficher <> Persos(i).ApparenceObjetAfficher Then
                    Perso_ApparenceObjetAfficher Persos(i), i, IDJoueur, ToutJoueurs
                    ComPersos(i).ApparenceObjetAfficher = Persos(i).ApparenceObjetAfficher
                End If
                If ComPersos(i).Definir_Nombre_ApparenceObjet <> Persos(i).Nombre_ApparenceObjet Then
                    Perso_NombreApparenceObjet Persos(i), i, IDJoueur, ToutJoueurs
                    ComPersos(i).Definir_Nombre_ApparenceObjet = Persos(i).Nombre_ApparenceObjet
                End If
                For j = 1 To .Nombre_ApparenceObjet
                    If ComPersos(i).Definir_ApparenceObjet(j) <> Persos(i).Definir_ApparenceObjet(j) Then
                        Perso_Definir_ApparenceObjet Persos(i), i, j, IDJoueur, ToutJoueurs
                        ComPersos(i).Definir_ApparenceObjet(j) = Persos(i).Definir_ApparenceObjet(j)
                    End If
                Next j
            End If
            If ComPersos(i).DansSaMaison <> Persos(i).DansSaMaison Or ComPersos(i).DansUneMaison <> Persos(i).DansUneMaison Then
                Perso_DansUneMaison Persos(i), i, IDJoueur, ToutJoueurs
                ComPersos(i).DansSaMaison = Persos(i).DansSaMaison: ComPersos(i).DansUneMaison = Persos(i).DansUneMaison
            End If
            If ComPersos(i).DansSonChateau <> Persos(i).DansSonChateau Or ComPersos(i).DansUnChateau <> Persos(i).DansUnChateau Then
                Perso_DansUnChateau Persos(i), i, IDJoueur, ToutJoueurs
                ComPersos(i).DansSonChateau = Persos(i).DansSonChateau: ComPersos(i).DansUnChateau = Persos(i).DansUnChateau
            End If
            'Envoie des informations de la maison.
            If ComPersos(i).MaisonVie <> Maisons(i).Vie Or ComPersos(i).MaisonMaxVie <> Maisons(i).MaxVie Then
                Maison_Vie i, IDJoueur, ToutJoueurs
                ComPersos(i).MaisonVie = Maisons(i).Vie: ComPersos(i).MaisonMaxVie = Maisons(i).MaxVie
            End If
            If Maisons(i).Visible Then
                If ComPersos(i).Maison_TypeBatiment <> Maisons(i).TypeBatiment Then
                    Maison_Type i, IDJoueur, ToutJoueurs
                    ComPersos(i).Maison_TypeBatiment = Maisons(i).TypeBatiment
                End If
                If ComPersos(i).MaisonConstruit <> Maisons(i).Construit Or ComPersos(i).MaisonEnConstruction <> Maisons(i).EnConstruction Then
                    Maison_EnConstruction i, IDJoueur, ToutJoueurs
                    ComPersos(i).MaisonConstruit = Maisons(i).Construit: ComPersos(i).MaisonEnConstruction = Maisons(i).EnConstruction
                End If
                If ComPersos(i).MaisonPositionX <> Maisons(i).PositionX Or ComPersos(i).MaisonPositionY <> Maisons(i).PositionY Then
                    Maison_Positions i, IDJoueur, ToutJoueurs
                    ComPersos(i).MaisonPositionX = Maisons(i).PositionX: ComPersos(i).MaisonPositionY = Maisons(i).PositionY
                End If
                If ComPersos(i).MaisonStock <> Maisons(i).Definir_Stock Or ComPersos(i).MaisonMaxStock <> Maisons(i).MaxStock Then
                    Maison_Stock i, IDJoueur, ToutJoueurs
                    ComPersos(i).MaisonStock = Maisons(i).Definir_Stock: ComPersos(i).MaisonMaxStock = Maisons(i).MaxStock
                End If
                'If EffetsMaisons(i).Duree = 1 Then
                '    Effet_ApparenceMaison i, IDJoueur, ToutJoueurs
                '    Effet_PositionsMaison i, IDJoueur, ToutJoueurs
                'End If
            End If
            End With
        Next i
        'Envoie les informations des fiefs.
        If NombreComFiefs <> UBound(Fiefs()) + 1 Then
            Fief_Nombre IDJoueur
            'NombreComFiefs = UBound(Fiefs())
        End If
        For i = 0 To UBound(Fiefs())
            If ComFiefs(i).Vie <> Chateaux(i).Vie Or ComFiefs(i).MaxVie <> Chateaux(i).MaxVie Then
                Chateau_Vie i, IDJoueur, ToutJoueurs
                ComFiefs(i).Vie = Chateaux(i).Vie: ComFiefs(i).MaxVie = Chateaux(i).MaxVie
            End If
            If ComFiefs(i).RecuperationVie <> Chateaux(i).RecuperationVie Then
                Chateau_RecuperationVie i, IDJoueur, ToutJoueurs
                ComFiefs(i).RecuperationVie = Chateaux(i).RecuperationVie
            End If
            If ComFiefs(i).NombreCitoyens <> Fiefs(i).NombreCitoyens Or ComFiefs(i).NombreMorts <> Fiefs(i).NombreMorts Then
                Fief_NombreCitoyens i, IDJoueur, ToutJoueurs
                ComFiefs(i).NombreCitoyens = Fiefs(i).NombreCitoyens: ComFiefs(i).NombreMorts = Fiefs(i).NombreMorts
            End If
            If ComFiefs(i).Empereur <> Fiefs(i).Empereur Then
                Fief_Empereur i, IDJoueur, ToutJoueurs
                ComFiefs(i).Empereur = Fiefs(i).Empereur
            End If
            If ComFiefs(i).Epoque <> Fiefs(i).Epoque Then
                Fief_Epoque i, IDJoueur, ToutJoueurs
                ComFiefs(i).Epoque = Fiefs(i).Epoque
            End If
            If ComFiefs(i).NumeroEquipe <> Fiefs(i).NumeroEquipe Then
                Fief_Equipe i, IDJoueur, ToutJoueurs
                ComFiefs(i).NumeroEquipe = Fiefs(i).NumeroEquipe
            End If
            If ComFiefs(i).TempsChangementEpoque <> Fiefs(i).TempsChangementEpoque Then
                Fief_TempsChangementEpoque i, IDJoueur, ToutJoueurs
                ComFiefs(i).TempsChangementEpoque = Fiefs(i).TempsChangementEpoque
            End If
        Next i
'        Tresor_Nombre IDJoueur, ToutJoueurs
'        For i = 1 To Tresor_DefinirNombre
'            Tresor_Apparence i, IDJoueur, ToutJoueurs
'            Tresor_Positions i, IDJoueur, ToutJoueurs
'        Next i
'        Decor_Nombre IDJoueur, ToutJoueurs
'        For i = 1 To Decor_DefinirNombre
'            Decor_Nom i, IDJoueur, ToutJoueurs
'            Decor_Apparence i, IDJoueur, ToutJoueurs
'            Decor_Positions i, IDJoueur, ToutJoueurs
'        Next i
        For i = 1 To AffDegat.Degats.Count
            If AffDegat.Degats(i).Duree = 1 Then
                Degat_Texte i, IDJoueur, ToutJoueurs
                Degat_Positions i, IDJoueur, ToutJoueurs
                Degat_Visuel i, IDJoueur, ToutJoueurs
            End If
        Next i
        For i = 1 To Effets.Count
            If Effets(i).Temps = 1 Then
                Effet_Apparence i, IDJoueur, ToutJoueurs
                Effet_Positions i, IDJoueur, ToutJoueurs
            End If
        Next i
        'ComReseau.Joueur_InfosAJour(IDJoueur) = True
    End If
End Sub
Public Sub Monde_Dimensions(ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Monde, Categorie2MondeDimensions, Monde.Largeur, Monde.Hauteur, Monde.NombreCases, , IDJoueur
End Sub
Public Sub Monde_Carte(ByVal IDJoueur As Long)
    Dim i As Long
    Dim Temp As String
    With Monde
    For i = 0 To Monde.NombreCases - 1
        Temp = Temp & Chr(Monde.Param_TerrainCategorie(i))
    Next i
    ComReseau.Message_Envoyer Categorie1Monde, Categorie2MondeCarte, Monde.Largeur, Monde.Hauteur, Temp, , IDJoueur
    End With
End Sub
Public Sub Monde_Terrain(ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Monde, Categorie2MondeTerrain, Indice, Monde.Param_TerrainCategorie(Indice), , , IDJoueur
End Sub
Public Sub Monde_TerrainFin(ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Monde, Categorie2MondeTerrainFin, , , , , IDJoueur
End Sub
Public Sub Monde_NomImageCarte(ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Monde, Categorie2MondeNomImageCarte, , , Cartes.NomImageCarte, , IDJoueur
End Sub
' ************************************************************************************
' *                    Liste des messages concernant les ressources                  *
' ************************************************************************************
Public Sub Ressource_Nombre(ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Ressource, Categorie2RessourceNombre, UBound(Ressources()), , , , IDJoueur
End Sub
Public Sub Ressource_Infos(ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Ressource, Categorie2RessourceInfos, Indice, Ressources(Indice).Apparence, Ressources(Indice).PositionX, Ressources(Indice).PositionY, IDJoueur
End Sub

' ************************************************************************************
' *                   Liste des messages concernant les personnages                  *
' ************************************************************************************

Public Sub Perso_Envoyer(ByVal Indice As Long, ByVal NumeroPerso As Long, ByVal IDJoueur As Long)
    'Envoie un personnage en entier.
    Dim i  As Integer
    IDJoueur = -1
    With Persos(Indice)
    Perso_Race Persos(Indice), NumeroPerso, IDJoueur
    'Perso_Joueur Persos(Indice), NumeroPerso, IDJoueur
    Perso_Nom Persos(Indice), NumeroPerso, IDJoueur
    Perso_Feminin Persos(Indice), NumeroPerso, IDJoueur
    Perso_Fief Persos(Indice), NumeroPerso, IDJoueur
    Perso_ChoisirApparence Persos(Indice), NumeroPerso, IDJoueur
    Perso_FichierApparence Persos(Indice), NumeroPerso, IDJoueur
    
    For i = 0 To .Nombre_ApparenceBase - 1
        Perso_Definir_FichierApparenceBase Persos(Indice), NumeroPerso, i, IDJoueur
    Next i
    Perso_ChoisirIA Persos(Indice), NumeroPerso, IDJoueur
    Perso_ChoisirIAScript Persos(Indice), NumeroPerso, IDJoueur
    Perso_AttitudeFormation Persos(Indice), NumeroPerso, IDJoueur
    Perso_ExpVie Persos(Indice), NumeroPerso, IDJoueur
    Perso_ExpEnergie Persos(Indice), NumeroPerso, IDJoueur
    Perso_ExpMagie Persos(Indice), NumeroPerso, IDJoueur
    Perso_ExpMoral Persos(Indice), NumeroPerso, IDJoueur
    Perso_ExpAttaque Persos(Indice), NumeroPerso, IDJoueur
    Perso_ExpDefense Persos(Indice), NumeroPerso, IDJoueur
    For i = 0 To .Nombre_CompetencesSpeciales - 1
        Perso_CompetencesSpecialesExp Persos(Indice), NumeroPerso, i, IDJoueur
    Next i
    For i = 0 To .Nombre_CompetencesRessources - 1
        Perso_CompetencesRessourcesExp Persos(Indice), NumeroPerso, i, IDJoueur
    Next i
    For i = 0 To .Nombre_CompetencesServices - 1
        Perso_CompetencesServicesExp Persos(Indice), NumeroPerso, i, IDJoueur
    Next i
    For i = 0 To .Nombre_CompetencesObjets - 1
        Perso_CompetencesObjetsExp Persos(Indice), NumeroPerso, i, IDJoueur
    Next i
    For i = 0 To .Nombre_de_lignes_du_bestiaire - 1
        Perso_Bestiaire Persos(Indice), NumeroPerso, i, IDJoueur
    Next i
    'Equipement.
    For i = 0 To Persos(Indice).Nombre_Objets_Equipes - 1
        Perso_ObjetsEquipes Persos(Indice), NumeroPerso, i, IDJoueur
    Next i
    For i = 0 To Persos(Indice).Nombre_Objets_Inventaire - 1
        Perso_ObjetsInventaire Persos(Indice), NumeroPerso, i, IDJoueur
    Next i
    'Effets temporaires
'    Perso_EffetTempNombre Persos(Indice), NumeroPerso, IDJoueur
'    For i = 1 To Persos(Indice).Nombre_EffetsTemp
'        Perso_EffetTemp Persos(Indice), NumeroPerso, i, IDJoueur
'        Perso_EffetTempProprietaire Persos(Indice), NumeroPerso, i, IDJoueur
'    Next i
    
    Perso_ActualiserBonus NumeroPerso, IDJoueur
    End With
End Sub

Public Sub Perso_Caracs(ByVal Indice As Long, ByVal IDJoueur As Long)
    Dim i As Integer
    If ComPersos(i).IA <> Persos(i).IA Then
        Perso_IA Persos(i), i, IDJoueur
        ComPersos(i).IA = Persos(i).IA
    End If
    If ComPersos(Indice).Energie <> Persos(Indice).Energie Or ComPersos(Indice).BonusMaxEnergie <> Persos(Indice).BonusMaxEnergie Then
        Perso_Energie Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).Energie = Persos(Indice).Energie: ComPersos(Indice).BonusMaxEnergie = Persos(Indice).BonusMaxEnergie
    End If
    If ComPersos(Indice).RecuperationEnergie <> Persos(Indice).RecuperationEnergie Then
        Perso_Energie2 Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).RecuperationEnergie = Persos(Indice).RecuperationEnergie
    End If
    If ComPersos(Indice).Magie <> Persos(Indice).Magie Or ComPersos(Indice).BonusMaxMagie <> Persos(Indice).BonusMaxMagie Then
        Perso_Magie Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).Magie = Persos(Indice).Magie: ComPersos(Indice).BonusMaxMagie = Persos(Indice).BonusMaxMagie
    End If
    If ComPersos(Indice).RecuperationMagie <> Persos(Indice).RecuperationMagie Then
        Perso_Magie2 Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).RecuperationMagie = Persos(Indice).RecuperationMagie
    End If
    If ComPersos(Indice).Moral <> Persos(Indice).Moral Or ComPersos(Indice).BonusMaxMoral <> Persos(Indice).BonusMaxMoral Then
        Perso_Moral Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).Moral = Persos(Indice).Moral: ComPersos(Indice).BonusMaxMoral = Persos(Indice).BonusMaxMoral
    End If
    If ComPersos(Indice).RecuperationMoral <> Persos(Indice).RecuperationMoral Then
        Perso_Moral2 Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).RecuperationMoral = Persos(Indice).RecuperationMoral
    End If
    
    If ComPersos(Indice).MaxVie <> Persos(Indice).MaxVie Then
        Perso_Vie3 Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).MaxVie = Persos(Indice).MaxVie
    End If
    If ComPersos(Indice).MaxEnergie <> Persos(Indice).MaxEnergie Then
        Perso_Energie3 Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).MaxEnergie = Persos(Indice).MaxEnergie
    End If
    If ComPersos(Indice).MaxMagie <> Persos(Indice).MaxMagie Then
        Perso_Magie3 Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).MaxMagie = Persos(Indice).MaxMagie
    End If
    If ComPersos(Indice).MaxMoral <> Persos(Indice).MaxMoral Then
        Perso_Moral3 Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).MaxMoral = Persos(Indice).MaxMoral
    End If
    If ComPersos(Indice).MaxAttaque <> Persos(Indice).MaxAttaque Then
        Perso_Attaque2 Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).MaxAttaque = Persos(Indice).MaxAttaque
    End If
    If ComPersos(Indice).MaxDefense <> Persos(Indice).MaxDefense Then
        Perso_Defense2 Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).MaxDefense = Persos(Indice).MaxDefense
    End If
    
    If ComPersos(Indice).Action <> Persos(Indice).Action Then
        Perso_Action Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).Action = Persos(Indice).Action
    End If
    If ComPersos(Indice).IndiceObjet <> Persos(Indice).IndiceObjet Or ComPersos(Indice).EtatFabricationObjet <> Persos(Indice).EtatFabricationObjet Then
        Perso_EtatFabricationObjet Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).IndiceObjet = Persos(Indice).IndiceObjet: ComPersos(Indice).EtatFabricationObjet = Persos(Indice).EtatFabricationObjet
    End If
    If ComPersos(Indice).RessourceCourante <> Persos(Indice).RessourceCourante Or ComPersos(Indice).IndiceRessource <> Persos(Indice).IndiceRessource Then
        Perso_Cibles1 Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).RessourceCourante = Persos(Indice).RessourceCourante: ComPersos(Indice).IndiceRessource = Persos(Indice).IndiceRessource
    End If
    If ComPersos(Indice).IndicePerso <> Persos(Indice).IndicePerso Or ComPersos(Indice).IndiceTresor <> Persos(Indice).IndiceTresor Then
        Perso_Cibles2 Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).IndicePerso = Persos(Indice).IndicePerso: ComPersos(Indice).IndiceTresor = Persos(Indice).IndiceTresor
    End If
    If ComPersos(Indice).IndiceMaison <> Persos(Indice).IndiceMaison Or ComPersos(Indice).IndiceChateau <> Persos(Indice).IndiceChateau Then
        Perso_Cibles3 Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).IndiceMaison = Persos(Indice).IndiceMaison: ComPersos(Indice).IndiceChateau = Persos(Indice).IndiceChateau
    End If
    If ComPersos(Indice).NombreSoldats <> Persos(Indice).NombreSoldats Then
        Perso_NombreSoldats Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).NombreSoldats = Persos(Indice).NombreSoldats
    End If
    'Argent.
    If ComPersos(Indice).Argent <> Persos(Indice).Definir_Argent Or ComPersos(Indice).ArgentSelectionne <> Persos(Indice).ArgentSelectionne Then
        Perso_Argent Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).Argent = Persos(Indice).Definir_Argent: ComPersos(Indice).ArgentSelectionne = Persos(Indice).ArgentSelectionne
    End If
    'Resurrection.
    If Persos(Indice).DansSonChateau Then
        If ComPersos(Indice).DureeResurection <> Persos(Indice).DureeRessurection Then
            Perso_DureeResurrection Persos(Indice), Indice, IDJoueur
            ComPersos(Indice).DureeResurection = Persos(Indice).DureeRessurection
        End If
    End If
    
    Perso_Competences Indice, IDJoueur
    Perso_ListeRessources Indice, IDJoueur
    Perso_Inventaire Indice, IDJoueur
    Perso_ListeBestiaire Indice, IDJoueur
    
    'Effets temporaires
'    Perso_EffetTempNombre Persos(Indice), Indice, IDJoueur
'    For i = 1 To Persos(Indice).Nombre_EffetsTemp
'        Perso_EffetTemp Persos(Indice), Indice, i, IDJoueur
'        Perso_EffetTempProprietaire Persos(Indice), Indice, i, IDJoueur
'    Next i
'    For i = 1 To Persos(Indice).Nombre_EffetsRecharge
'        Perso_EffetRecharge Persos(Indice), Indice, i, IDJoueur
'    Next i
End Sub
Public Sub Perso_Caracs2(ByVal Indice As Long, ByVal IDJoueur As Long)
    If ComPersos(Indice).ExpVie <> Persos(Indice).ExpVie Then
        Perso_VieExp Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).ExpVie = Persos(Indice).ExpVie
    End If
    If ComPersos(Indice).ExpEnergie <> Persos(Indice).ExpEnergie Then
        Perso_EnergieExp Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).ExpEnergie = Persos(Indice).ExpEnergie
    End If
    If ComPersos(Indice).ExpMagie <> Persos(Indice).ExpMagie Then
        Perso_MagieExp Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).ExpMagie = Persos(Indice).ExpMagie
    End If
    If ComPersos(Indice).ExpMoral <> Persos(Indice).ExpMoral Then
        Perso_MoralExp Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).ExpMoral = Persos(Indice).ExpMoral
    End If
    If ComPersos(Indice).ExpAttaque <> Persos(Indice).ExpAttaque Then
        Perso_AttaqueExp Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).ExpAttaque = Persos(Indice).ExpAttaque
    End If
    If ComPersos(Indice).ExpDefense <> Persos(Indice).ExpDefense Then
        Perso_DefenseExp Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).ExpDefense = Persos(Indice).ExpDefense
    End If
    If ComPersos(Indice).Armure <> Persos(Indice).Definir_Armure Then
        Perso_Armure Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).Armure = Persos(Indice).Definir_Armure
    End If
    If ComPersos(Indice).Vitalite <> Persos(Indice).Vitalite Then
        Perso_Vitalite Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).Vitalite = Persos(Indice).Vitalite
    End If
    If ComPersos(Indice).Longueur_PasTemp <> Persos(Indice).Longueur_PasTemp Or ComPersos(Indice).BonusDeplacement <> Persos(Indice).BonusDeplacement Then
        Perso_Vitesse Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).Longueur_PasTemp = Persos(Indice).Longueur_PasTemp: ComPersos(Indice).BonusDeplacement = Persos(Indice).BonusDeplacement
    End If
    'Perso_Actualiser Indice, IDJoueur
End Sub

Public Sub Perso_ListeRessources(ByVal Indice As Long, ByVal IDJoueur As Long)
    Dim i As Integer
    For i = 0 To Parametres.NombreRessources - 1
        If ComPersos(Indice).Definir_ressources(i) <> Persos(Indice).Definir_ressources(i) Or ComPersos(Indice).Definir_MaxRessources(i) <> Persos(Indice).Definir_MaxRessources(i) Then
            Perso_Ressources Persos(Indice), Indice, i, IDJoueur
            ComPersos(Indice).Definir_ressources(i) = Persos(Indice).Definir_ressources(i): ComPersos(Indice).Definir_MaxRessources(i) = Persos(Indice).Definir_MaxRessources(i)
        End If
        If ComPersos(Indice).Definir_RessourcesSelectionnes(i) <> Persos(Indice).Definir_RessourcesSelectionnes(i) Then
            Perso_RessourcesSelectionnes Persos(Indice), Indice, i, IDJoueur
            ComPersos(Indice).Definir_RessourcesSelectionnes(i) = Persos(Indice).Definir_RessourcesSelectionnes(i)
        End If
    Next i
End Sub

Public Sub Perso_Inventaire(ByVal Indice As Long, ByVal IDJoueur As Long)
    Dim i As Integer
    If ComPersos(Indice).ObjetSelectionne <> Persos(Indice).ObjetSelectionne Then
        Perso_ObjetSelectionne Persos(Indice), Indice, IDJoueur
        ComPersos(Indice).ObjetSelectionne = Persos(Indice).ObjetSelectionne
    End If
    For i = 0 To Persos(Indice).Nombre_Objets_Equipes - 1
        If ComPersos(Indice).Objet_Equipes_Type(i) <> Persos(Indice).Objet_Equipes_Type(i) Then
            Perso_ObjetsEquipes Persos(Indice), Indice, i, IDJoueur
            ComPersos(Indice).Objet_Equipes_Type(i) = Persos(Indice).Objet_Equipes_Type(i)
        End If
    Next i
    For i = 0 To Persos(Indice).Nombre_Objets_Inventaire - 1
        If ComPersos(Indice).Objet_Inventaire_Type(i) <> Persos(Indice).Objet_Inventaire_Type(i) Or ComPersos(Indice).Objet_Inventaire_Nombre(i) <> Persos(Indice).Objet_Inventaire_Nombre(i) Then
            Perso_ObjetsInventaire Persos(Indice), Indice, i, IDJoueur
            ComPersos(Indice).Objet_Inventaire_Type(i) = Persos(Indice).Objet_Inventaire_Type(i): ComPersos(Indice).Objet_Inventaire_Nombre(i) = Persos(Indice).Objet_Inventaire_Nombre(i)
        End If
    Next i
End Sub

Public Sub Perso_InventaireMaison(ByVal Indice As Long, ByVal IDJoueur As Long)
    Dim i As Integer
    For i = 0 To Parametres.PersosNombreObjetsInventaireMaison - 1
        If ComPersos(Indice).MaisonObjet_Inventaire_Type(i) <> Maisons(Indice).Definir_ObjetsInventaire(i) Then
            Maison_Inventaire Indice, i, IDJoueur
            ComPersos(Indice).MaisonObjet_Inventaire_Type(i) = Maisons(Indice).Definir_ObjetsInventaire(i)
        End If
    Next i
End Sub

Public Sub Perso_Competences(ByVal Indice As Long, ByVal IDJoueur As Long)
    Dim i As Integer
    For i = 0 To Parametres.NombreCompetencesSpeciales - 1
        If ComPersos(Indice).Definir_Carac_CompetenceSpeciales(i) <> Persos(Indice).Definir_Carac_CompetenceSpeciales(i) Or _
           ComPersos(Indice).Definir_Niveau_CompetenceSpeciales(i) <> Persos(Indice).Definir_Niveau_CompetenceSpeciales(i) Then
            Perso_CompetencesSpeciales Persos(Indice), Indice, i, IDJoueur
            ComPersos(Indice).Definir_Carac_CompetenceSpeciales(i) = Persos(Indice).Definir_Carac_CompetenceSpeciales(i)
            ComPersos(Indice).Definir_Niveau_CompetenceSpeciales(i) = Persos(Indice).Definir_Niveau_CompetenceSpeciales(i)
        End If
    Next i
    For i = 0 To Parametres.NombreRessources - 1
        If ComPersos(Indice).Definir_Carac_Ressources(i) <> Persos(Indice).Definir_Carac_Ressources(i) Or _
           ComPersos(Indice).Definir_Niveau_CompetenceRessources(i) <> Persos(Indice).Definir_Niveau_CompetenceRessources(i) Then
            Perso_CompetencesRessources Persos(Indice), Indice, i, IDJoueur
            ComPersos(Indice).Definir_Carac_Ressources(i) = Persos(Indice).Definir_Carac_Ressources(i)
            ComPersos(Indice).Definir_Niveau_CompetenceRessources(i) = Persos(Indice).Definir_Niveau_CompetenceRessources(i)
        End If
    Next i
    For i = 0 To Parametres.NombreServices - 2
        If ComPersos(Indice).Definir_Carac_CompetenceServices(i) <> Persos(Indice).Definir_Carac_CompetenceServices(i) Or _
           ComPersos(Indice).Definir_Niveau_CompetenceServices(i) <> Persos(Indice).Definir_Niveau_CompetenceServices(i) Then
            Perso_CompetencesServices Persos(Indice), Indice, i, IDJoueur
            ComPersos(Indice).Definir_Carac_CompetenceServices(i) = Persos(Indice).Definir_Carac_CompetenceServices(i)
            ComPersos(Indice).Definir_Niveau_CompetenceServices(i) = Persos(Indice).Definir_Niveau_CompetenceServices(i)
        End If
    Next i
    For i = 0 To Parametres.NombreCompetencesObjets - 1
        If ComPersos(Indice).Definir_Carac_CompetenceObjets(i) <> Persos(Indice).Definir_Carac_CompetenceObjets(i) Or _
           ComPersos(Indice).Definir_Niveau_CompetenceObjets(i) <> Persos(Indice).Definir_Niveau_CompetenceObjets(i) Then
            Perso_CompetencesObjets Persos(Indice), Indice, i, IDJoueur
            ComPersos(Indice).Definir_Carac_CompetenceObjets(i) = Persos(Indice).Definir_Carac_CompetenceObjets(i)
            ComPersos(Indice).Definir_Niveau_CompetenceObjets(i) = Persos(Indice).Definir_Niveau_CompetenceObjets(i)
        End If
    Next i
    'Perso_Actualiser Indice, IDJoueur
End Sub

Public Sub Perso_Competences2(ByVal Indice As Long, ByVal IDJoueur As Long)
    Dim i As Integer
    For i = 0 To Parametres.NombreCompetencesSpeciales - 1
        If ComPersos(Indice).Definir_ExpSpeciales(i) <> Persos(Indice).Definir_ExpSpeciales(i) Then
            Perso_CompetencesSpecialesExp Persos(Indice), Indice, i, IDJoueur
            ComPersos(Indice).Definir_ExpSpeciales(i) = Persos(Indice).Definir_ExpSpeciales(i)
        End If
    Next i
    For i = 0 To Parametres.NombreRessources - 1
        If ComPersos(Indice).Definir_ExpRessources(i) <> Persos(Indice).Definir_ExpRessources(i) Then
            Perso_CompetencesRessourcesExp Persos(Indice), Indice, i, IDJoueur
            ComPersos(Indice).Definir_ExpRessources(i) = Persos(Indice).Definir_ExpRessources(i)
        End If
    Next i
    For i = 0 To Parametres.NombreServices - 2
        If ComPersos(Indice).Definir_ExpServices(i) <> Persos(Indice).Definir_ExpServices(i) Then
            Perso_CompetencesServicesExp Persos(Indice), Indice, i, IDJoueur
            ComPersos(Indice).Definir_ExpServices(i) = Persos(Indice).Definir_ExpServices(i)
        End If
    Next i
    For i = 0 To Parametres.NombreCompetencesObjets - 1
        If ComPersos(Indice).Definir_ExpObjets(i) <> Persos(Indice).Definir_ExpObjets(i) Then
            Perso_CompetencesObjetsExp Persos(Indice), Indice, i, IDJoueur
            ComPersos(Indice).Definir_ExpObjets(i) = Persos(Indice).Definir_ExpObjets(i)
        End If
    Next i
End Sub

Public Sub Perso_ListeBestiaire(ByVal Indice As Long, ByVal IDJoueur As Long)
    Dim i As Long
    For i = 0 To Persos(Indice).Nombre_de_lignes_du_bestiaire - 1
        If ComPersos(Indice).Definir_Bestiaire(i) <> Persos(Indice).Definir_Bestiaire(i) Then
            Perso_Bestiaire Persos(Indice), Indice, i, IDJoueur
            ComPersos(Indice).Definir_Bestiaire(i) = Persos(Indice).Definir_Bestiaire(i)
        End If
    Next i
End Sub

Public Sub Maison_ListeStock(ByVal Indice As Long, ByVal IDJoueur As Long)
    Dim i As Integer
    If Parametres.Batiment_Fabrique(Maisons(Indice).TypeBatiment) Then
        For i = 0 To Parametres.Batiment_NombreObjetsFabriques(Maisons(Indice).TypeBatiment) - 1
            If ComPersos(Indice).MaisonStock_Objet(i) <> Maisons(Indice).Stock_Objets(i) Then
                Maison_StockObjet Indice, i, IDJoueur
                ComPersos(Indice).MaisonStock_Objet(i) = Maisons(Indice).Stock_Objets(i)
            End If
        Next i
    End If
End Sub

Public Sub Chateau_ListeStock(ByVal Indice As Long, ByVal IDJoueur As Long)
    Dim i As Integer
    For i = 0 To Parametres.NombreRessources - 1
        If ComFiefs(Indice).Definir_ressources(i) <> Fiefs(Indice).Definir_ressources(i) Then
            Chateau_Ressources Indice, i, IDJoueur
            ComFiefs(Indice).Definir_ressources(i) = Fiefs(Indice).Definir_ressources(i)
        End If
    Next i
End Sub

Public Sub Perso_Nombre(ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoNombre, UBound(Persos()), , , , IDJoueur
End Sub
Public Sub Perso_Nom(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoNom, Indice, , Perso.Nom, , IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_Fief(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoFief, Indice, , Perso.NumeroFief, , IDJoueur
End Sub
Public Sub Perso_Race(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoRace, Indice, , Perso.Race, , IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_Positions(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoPositions, Indice, , Perso.PositionX, Perso.PositionY, IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_Vie(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoVie, Indice, , Perso.Vie, Perso.BonusMaxVie, IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_Vie2(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoVie2, Indice, , Perso.RecuperationVie, , IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_Vie3(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoVie3, Indice, , Perso.MaxVie, , IDJoueur
End Sub
Public Sub Perso_VieExp(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoVieExp, Indice, , Perso.ExpVie, , IDJoueur
End Sub
Public Sub Perso_Energie(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoEnergie, Indice, , Perso.Energie, Perso.BonusMaxEnergie, IDJoueur
End Sub
Public Sub Perso_Energie2(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoEnergie2, Indice, , Perso.RecuperationEnergie, , IDJoueur
End Sub
Public Sub Perso_Energie3(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoEnergie3, Indice, , Perso.MaxEnergie, , IDJoueur
End Sub
Public Sub Perso_EnergieExp(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoEnergieExp, Indice, , Perso.ExpEnergie, , IDJoueur
End Sub
Public Sub Perso_Magie(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoMagie, Indice, , Perso.Magie, Perso.BonusMaxMagie, IDJoueur
End Sub
Public Sub Perso_Magie2(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoMagie2, Indice, , Perso.RecuperationMagie, , IDJoueur
End Sub
Public Sub Perso_Magie3(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoMagie3, Indice, , Perso.MaxMagie, , IDJoueur
End Sub
Public Sub Perso_MagieExp(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoMagieExp, Indice, , Perso.ExpMagie, , IDJoueur
End Sub
Public Sub Perso_Moral(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoMoral, Indice, , Perso.Moral, Perso.BonusMaxMoral, IDJoueur
End Sub
Public Sub Perso_Moral2(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoMoral2, Indice, , Perso.RecuperationMoral, , IDJoueur
End Sub
Public Sub Perso_Moral3(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoMoral3, Indice, , Perso.MaxMoral, , IDJoueur
End Sub
Public Sub Perso_MoralExp(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoMoralExp, Indice, , Perso.ExpMoral, , IDJoueur
End Sub
Public Sub Perso_Attaque(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoAttaque, Indice, , Perso.Attaque, Perso.BonusMaxAttaque, IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_Attaque2(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoAttaque2, Indice, , Perso.MaxAttaque, , IDJoueur
End Sub
Public Sub Perso_AttaqueRecharge(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoAttaqueRecharge, Indice, , Perso.DureeRechargeAttaque, Perso.DureeRechargeAttaqueMax, IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_AttaqueExp(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoAttaqueExp, Indice, , Perso.ExpAttaque, , IDJoueur
End Sub
Public Sub Perso_Defense(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoDefense, Indice, , Perso.Defense, Perso.BonusMaxDefense, IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_Defense2(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoDefense2, Indice, , Perso.MaxDefense, , IDJoueur
End Sub
Public Sub Perso_DefenseExp(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoDefenseExp, Indice, , Perso.ExpDefense, , IDJoueur
End Sub
Public Sub Perso_Joueur(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoJoueur, Indice, , Perso.Joueur, , IDJoueur
End Sub
Public Sub Perso_Equipe(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoEquipe, Indice, , Perso.NumeroEquipe, , IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_IA(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoIA, Indice, , Perso.IA, , IDJoueur, ToutJoueurs
End Sub
'Public Sub Perso_Pas(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
'    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoPas, Indice, , Perso.NumeroPas, , IDJoueur, ToutJoueurs
'End Sub
Public Sub Perso_Direction(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoDirection, Indice, , Perso.Direction, , IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_Feminin(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoFeminin, Indice, , Perso.Feminin, , IDJoueur, ToutJoueurs
End Sub
'Public Sub Perso_Vivant(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
'    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoVivant, Indice, , Perso.Vivant, , IDJoueur
'End Sub
Public Sub Perso_ChoisirApparence(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoChoisirApparence, Indice, , Perso.ChoisirApparence, , IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_FichierApparence(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoFichierApparence, Indice, , Perso.FichierApparence, , IDJoueur
End Sub
Public Sub Perso_ApparenceChoisie(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoApparenceChoisie, Indice, , Perso.ApparenceChoisie, , IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_Nombre_ApparenceBase(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoNombre_ApparenceBase, Indice, , Perso.Nombre_ApparenceBase, , IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_Definir_ApparenceBase(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoDefinir_ApparenceBase, Indice1, Indice2, Perso.Definir_ApparenceBase(Indice2), , IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_Definir_ApparenceDecalage(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoDefinir_ApparenceDecalage, Indice1, Indice2, Perso.Definir_ApparenceDecalageX(Indice2), Perso.Definir_ApparenceDecalageYFixe(Indice2), IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_Definir_ApparenceAvantPlan(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoDefinir_ApparenceAvantPlan, Indice1, Indice2, Perso.Definir_ApparenceAvantPlan(Indice2), , IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_Definir_FichierApparenceBase(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoDefinir_FichierApparenceBase, Indice1, Indice2, Perso.Definir_FichierApparenceBase(Indice2), , IDJoueur
End Sub
Public Sub Perso_ApparenceObjetAfficher(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoApparenceObjetAfficher, Indice, , Perso.ApparenceObjetAfficher, , IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_NombreApparenceObjet(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoNombreApparenceObjet, Indice, Perso.Nombre_ApparenceObjet, , , IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_Definir_ApparenceObjet(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoDefinir_ApparenceObjet, Indice1, Indice2, Perso.Definir_ApparenceObjet(Indice2), , IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_DansUneMaison(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoDansUneMaison, Indice, , Perso.DansSaMaison, Perso.DansUneMaison, IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_DansUnChateau(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoDansUnChateau, Indice, , Perso.DansSonChateau, Perso.DansUnChateau, IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_ChangerNombre(ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoChangerNombre, UBound(Persos()), , , , IDJoueur
End Sub
Public Sub Perso_ChoisirIA(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoChoisirIA, Indice, , Perso.ChoisirIA, Perso.FichierIA, IDJoueur
End Sub
Public Sub Perso_ChoisirIAScript(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoChoisirIAScript, Indice, , Perso.ChoisirIAScript, Perso.FichierIAScript, IDJoueur
End Sub
Public Sub Perso_NumeroChef(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoNumeroChef, Indice, , Perso.NumeroChef, , IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_NombreSoldats(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoNombreSoldats, Indice, , Perso.NombreSoldats, , IDJoueur
End Sub
Public Sub Perso_AttitudeFormation(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoAttitudeFormation, Indice, , Perso.Attitude, Perso.Formation, IDJoueur
End Sub
Public Sub Perso_ExpVie(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoExpVie, Indice, , Perso.ExpVie, , IDJoueur
End Sub
Public Sub Perso_ExpEnergie(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoExpEnergie, Indice, , Perso.ExpEnergie, , IDJoueur
End Sub
Public Sub Perso_ExpMagie(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoExpMagie, Indice, , Perso.ExpMagie, , IDJoueur
End Sub
Public Sub Perso_ExpMoral(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoExpMoral, Indice, , Perso.ExpMoral, , IDJoueur
End Sub
Public Sub Perso_ExpAttaque(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoExpAttaque, Indice, , Perso.ExpAttaque, , IDJoueur
End Sub
Public Sub Perso_ExpDefense(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoExpDefense, Indice, , Perso.ExpDefense, , IDJoueur
End Sub
Public Sub Perso_CompetencesSpeciales(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoCompetencesSpeciales, Indice1, Indice2, Perso.Definir_Carac_CompetenceSpeciales(Indice2), Perso.Definir_Niveau_CompetenceSpeciales(Indice2), IDJoueur
End Sub
Public Sub Perso_CompetencesRessources(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoCompetencesRessources, Indice1, Indice2, Perso.Definir_Carac_Ressources(Indice2), Perso.Definir_Niveau_CompetenceRessources(Indice2), IDJoueur
End Sub
Public Sub Perso_CompetencesServices(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoCompetencesServices, Indice1, Indice2, Perso.Definir_Carac_CompetenceServices(Indice2), Perso.Definir_Niveau_CompetenceServices(Indice2), IDJoueur
End Sub
Public Sub Perso_CompetencesObjets(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoCompetencesObjets, Indice1, Indice2, Perso.Definir_Carac_CompetenceObjets(Indice2), Perso.Definir_Niveau_CompetenceObjets(Indice2), IDJoueur
End Sub
Public Sub Perso_CompetencesSpecialesExp(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoCompetencesSpecialesExp, Indice1, Indice2, Perso.Definir_ExpSpeciales(Indice2), , IDJoueur
End Sub
Public Sub Perso_CompetencesRessourcesExp(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoCompetencesRessourcesExp, Indice1, Indice2, Perso.Definir_ExpRessources(Indice2), , IDJoueur
End Sub
Public Sub Perso_CompetencesServicesExp(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoCompetencesServicesExp, Indice1, Indice2, Perso.Definir_ExpServices(Indice2), , IDJoueur
End Sub
Public Sub Perso_CompetencesObjetsExp(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoCompetencesObjetsExp, Indice1, Indice2, Perso.Definir_ExpObjets(Indice2), , IDJoueur
End Sub
Public Sub Perso_Bestiaire(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoBestiaire, Indice1, Indice2, Perso.Definir_Bestiaire(Indice2), , IDJoueur
End Sub
Public Sub Perso_Action(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoAction, Indice1, , Perso.Action, , IDJoueur
End Sub
Public Sub Perso_EtatFabricationObjet(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoEtatFabricationObjet, Indice1, , Perso.IndiceObjet, Perso.EtatFabricationObjet, IDJoueur
End Sub
Public Sub Perso_Cibles1(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoCibles1, Indice1, , Perso.RessourceCourante, Perso.IndiceRessource, IDJoueur
End Sub
Public Sub Perso_Cibles2(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoCibles2, Indice1, , Perso.IndicePerso, Perso.IndiceTresor, IDJoueur
End Sub
Public Sub Perso_Cibles3(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoCibles3, Indice1, , Perso.IndiceMaison, Perso.IndiceChateau, IDJoueur
End Sub
Public Sub Perso_Argent(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoArgent, Indice, , Perso.Definir_Argent, Perso.ArgentSelectionne, IDJoueur
End Sub
Public Sub Perso_Ressources(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoRessources, Indice1, Indice2, Perso.Definir_ressources(Indice2), Perso.Definir_MaxRessources(Indice2), IDJoueur
End Sub
Public Sub Perso_RessourcesSelectionnes(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoRessourcesSelectionnees, Indice1, Indice2, Perso.Definir_RessourcesSelectionnes(Indice2), , IDJoueur
End Sub
Public Sub Perso_ObjetSelectionne(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoObjetSelectionne, Indice, , Perso.ObjetSelectionne, , IDJoueur
End Sub
Public Sub Perso_ObjetsEquipes(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoObjetsEquipes, Indice1, Indice2, Perso.Objet_Equipes_Type(Indice2), , IDJoueur
End Sub
Public Sub Perso_ObjetsInventaire(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoObjetsInventaire, Indice1, Indice2, Perso.Objet_Inventaire_Type(Indice2), Perso.Objet_Inventaire_Nombre(Indice2), IDJoueur
End Sub
Public Sub Perso_Armure(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoArmure, Indice, , Perso.Definir_Armure, , IDJoueur
End Sub
Public Sub Perso_Vitalite(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoVitalite, Indice, , Perso.Vitalite, , IDJoueur
End Sub
Public Sub Perso_Vitesse(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoVitesse, Indice, , Perso.Longueur_PasTemp, Perso.BonusDeplacement, IDJoueur
End Sub
Public Sub Perso_EffetTempNombre(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoEffetTempNombre, Indice, , Perso.Nombre_EffetsTemp, Perso.Nombre_EffetsRecharge
End Sub
Public Sub Perso_EffetTempNumero(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoEffetTempNumero, Indice1, Indice2, Perso.Definir_EffetTemp_Numero(Indice2), , IDJoueur
End Sub
Public Sub Perso_EffetTempDuree(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoEffetTempDuree, Indice1, Indice2, Perso.Definir_EffetTemp_Duree(Indice2), , IDJoueur
End Sub
Public Sub Perso_EffetTempProprietaire(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoEffetTempProprietaire, Indice1, Indice2, Perso.Definir_EffetTemp_Proprietaire(Indice2), , IDJoueur
End Sub
Public Sub Perso_EffetTempAjouter(ByVal NoPerso As Long, _
                                  ByVal No As Long, _
                                  ByVal Duree As Long, _
                                  ByVal DureeMax As Long, _
                                  ByVal Proprietaire As Long, _
                                  ByVal IDJoueur As Long, _
                                  Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoEffetTempAjouter, NoPerso, No, Duree & "/" & DureeMax, Proprietaire, IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_EffetTempSupprimer(ByVal NoPerso As Long, _
                                    ByVal No As Long, _
                                    ByVal IDJoueur As Long, _
                                    Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoEffetTempSupprimer, NoPerso, No, , , IDJoueur
End Sub
Public Sub Perso_EffetRecharge(ByVal Perso As ClsJeuPerso, ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoEffetRecharge, Indice1, Indice2, Perso.Definir_EffetsRecharge_Numero(Indice2), Perso.Definir_EffetsRecharge_Duree(Indice2), IDJoueur
End Sub
Public Sub Perso_EffetRechargeAjouter(ByVal NoPerso As Long, ByVal No As Long, ByVal Duree As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoEffetRechargeAjouter, NoPerso, No, Duree, , IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_EffetRechargeSupprimer(ByVal NoPerso As Long, ByVal No As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoEffetRechargeSupprimer, NoPerso, No, , , IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_Commentaires(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoCommentaires, Indice, , Perso.Definir_Commentaires, , IDJoueur, ToutJoueurs
End Sub
Public Sub Perso_DureeResurrection(ByVal Perso As ClsJeuPerso, ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoDureeResurrection, Indice, , Perso.DureeRessurection, , IDJoueur
End Sub
Public Sub Perso_Actualiser(ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoActualiser, Indice, IDJoueur, , , IDJoueur
End Sub
Public Sub Perso_ActualiserBonus(ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoActualiserBonus, Indice, IDJoueur, , , IDJoueur
End Sub
Public Sub Perso_Enregistrer(ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Perso, Categorie2PersoEnregistrer, Indice, , , , IDJoueur
End Sub

' ************************************************************************************
' *                     Liste des messages concernant les maisons                    *
' ************************************************************************************
Public Sub Maison_Type(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Maison, Categorie2MaisonType, Indice, Maisons(Indice).TypeBatiment, , , IDJoueur, ToutJoueurs
End Sub
Public Sub Maison_EnConstruction(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Maison, Categorie2MaisonEnConstruction, Indice, , Maisons(Indice).EnConstruction, Maisons(Indice).Construit, IDJoueur, ToutJoueurs
End Sub
Public Sub Maison_Positions(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Maison, Categorie2MaisonPositions, Indice, , Maisons(Indice).PositionX, Maisons(Indice).PositionY, IDJoueur, ToutJoueurs
End Sub
Public Sub Maison_Vie(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Maison, Categorie2MaisonVie, Indice, , Maisons(Indice).Vie, Maisons(Indice).MaxVie, IDJoueur, ToutJoueurs
End Sub
Public Sub Maison_Stock(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Maison, Categorie2MaisonStock, Indice, , Maisons(Indice).Definir_Stock, Maisons(Indice).MaxStock, IDJoueur, ToutJoueurs
    'Messages(0).Ajouter_Message Categorie1Maison & " : " & Categorie2MaisonStock & " : " & Indice & " : " & "0" & " : " & Maisons(Indice).Definir_Stock & " : " & Maisons(Indice).MaxStock & " : " & IDJoueur
End Sub
Public Sub Maison_StockObjet(ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Maison, Categorie2MaisonStockObjet, Indice1, Indice2, Maisons(Indice1).Stock_Objets(Indice2), , IDJoueur
    'Messages(0).Ajouter_Message Categorie1Maison & " : " & Categorie2MaisonStockObjet & " : " & Indice1 & " : " & Indice2 & " : " & Maisons(Indice1).Stock_Objets(Indice2) & " : 0 : " & IDJoueur
End Sub
Public Sub Maison_Inventaire(ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Maison, Categorie2MaisonInventaire, Indice1, Indice2, Maisons(Indice1).Definir_ObjetsInventaire(Indice2), , IDJoueur
    'Messages(0).Ajouter_Message Categorie1Maison & " : " & Categorie2MaisonInventaire & " : " & Indice1 & " : " & Indice2 & " : " & Maisons(Indice1).Definir_ObjetsInventaire(Indice2) & " : 0 : " & IDJoueur
End Sub

' ************************************************************************************
' *                     Liste des messages concernant les châteaux                   *
' ************************************************************************************
Public Sub Chateau_Positions(ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Chateau, Categorie2ChateauPositions, Indice, , Chateaux(Indice).PositionX, Chateaux(Indice).PositionY, IDJoueur
End Sub
Public Sub Chateau_Vie(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Chateau, Categorie2ChateauVie, Indice, , Chateaux(Indice).Vie, Chateaux(Indice).MaxVie, IDJoueur, ToutJoueurs
End Sub
Public Sub Chateau_RecuperationVie(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Chateau, Categorie2ChateauRecuperationVie, Indice, , Chateaux(Indice).RecuperationVie, , IDJoueur, ToutJoueurs
End Sub
Public Sub Chateau_Ressources(ByVal Indice1 As Long, ByVal Indice2 As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Chateau, Categorie2ChateauRessources, Indice1, Indice2, Fiefs(Indice1).Definir_ressources(Indice2), , IDJoueur
End Sub

' ************************************************************************************
' *                      Liste des messages concernant les fiefs                     *
' ************************************************************************************
Public Sub Fief_Nombre(ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Fief, Categorie2FiefNombre, UBound(Fiefs()), , , , IDJoueur
End Sub
Public Sub Fief_Peuple(ByVal Indice As Long, ByVal IDJoueur As Long)
    ComReseau.Message_Envoyer Categorie1Fief, Categorie2FiefPeuple, Indice, , Fiefs(Indice).Definir_TypePeuple, , IDJoueur
End Sub
Public Sub Fief_Epoque(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Fief, Categorie2FiefEpoque, Indice, , Fiefs(Indice).Epoque, , IDJoueur, ToutJoueurs
End Sub
Public Sub Fief_Equipe(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Fief, Categorie2FiefEquipe, Indice, , Fiefs(Indice).NumeroEquipe, , IDJoueur, ToutJoueurs
End Sub
Public Sub Fief_TempsChangementEpoque(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Fief, Categorie2FiefTempsChangementEpoque, Indice, , Fiefs(Indice).TempsChangementEpoque, , IDJoueur, ToutJoueurs
End Sub
Public Sub Fief_NombreCitoyens(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Fief, Categorie2FiefNombreCitoyens, Indice, , Fiefs(Indice).NombreCitoyens, Fiefs(Indice).NombreMorts, IDJoueur, ToutJoueurs
End Sub
Public Sub Fief_Empereur(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Fief, Categorie2FiefEmpereur, Indice, , Fiefs(Indice).Empereur, , IDJoueur, ToutJoueurs
End Sub

' ************************************************************************************
' *                     Liste des messages concernant les trésors                    *
' ************************************************************************************
Public Sub Tresor_Nombre(ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Tresor, Categorie2TresorNombre, , , Tresor_DefinirNombre, , IDJoueur, ToutJoueurs
End Sub
Public Sub Tresor_Apparence(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Tresor, Categorie2TresorApparence, Indice, , Tresors(Indice).IndiceImage, , IDJoueur, ToutJoueurs
End Sub
Public Sub Tresor_Positions(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Tresor, Categorie2TresorPositions, Indice, , Tresors(Indice).X, Tresors(Indice).Y, IDJoueur, ToutJoueurs
End Sub
Public Sub Tresor_Ajouter(ByVal PositionX As Long, ByVal PositionY As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Tresor, Categorie2TresorAjouter, , , PositionX, PositionY, IDJoueur, ToutJoueurs
End Sub
Public Sub Tresor_Supprimer(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Tresor, Categorie2TresorSupprimer, Indice, , , , IDJoueur, ToutJoueurs
End Sub

' ************************************************************************************
' *                     Liste des messages concernant les décors                     *
' ************************************************************************************
Public Sub Decor_Nombre(ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Decor, Categorie2DecorNombre, , , Decor_DefinirNombre, , IDJoueur, ToutJoueurs
End Sub
Public Sub Decor_Nom(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Decor, Categorie2DecorNom, Indice, , Decors(Indice).Nom, , IDJoueur, ToutJoueurs
End Sub
Public Sub Decor_Apparence(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Decor, Categorie2DecorApparence, Indice, , Decors(Indice).Categorie, Decors(Indice).EtapeAnimation, IDJoueur, ToutJoueurs
End Sub
Public Sub Decor_Positions(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Decor, Categorie2DecorPositions, Indice, , Decors(Indice).X, Decors(Indice).Y, IDJoueur, ToutJoueurs
End Sub
Public Sub Decor_Ajouter(ByVal Categorie As Integer, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal Nom As String, _
                         ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Decor, Categorie2DecorAjouter, X, Y, Categorie, Nom, IDJoueur, ToutJoueurs
End Sub
Public Sub Decor_Supprimer(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Decor, Categorie2DecorSupprimer, Indice, , , , IDJoueur, ToutJoueurs
End Sub

' ************************************************************************************
' *                     Liste des messages concernant les dégats                     *
' ************************************************************************************
Public Sub Degat_Texte(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Degat, Categorie2DegatTexte, Indice, , AffDegat.Degats(Indice).Score, , IDJoueur, ToutJoueurs
End Sub
Public Sub Degat_Positions(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Degat, Categorie2DegatPositions, Indice, , AffDegat.Degats(Indice).X, AffDegat.Degats(Indice).Y, IDJoueur, ToutJoueurs
End Sub
Public Sub Degat_Visuel(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1Degat, Categorie2DegatVisuel, Indice, , AffDegat.Degats(Indice).Couleur, AffDegat.Degats(Indice).Critique, IDJoueur, ToutJoueurs
End Sub

' ************************************************************************************
' *                     Liste des messages concernant les effets                     *
' ************************************************************************************
Public Sub Effet_Apparence(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1EffetsVisuels, Categorie2EffetApparence, Indice, , Effets(Indice).Indice, , IDJoueur, ToutJoueurs
End Sub
Public Sub Effet_Positions(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
    ComReseau.Message_Envoyer Categorie1EffetsVisuels, Categorie2EffetPositions, Indice, , Effets(Indice).PositionX, Effets(Indice).PositionY, IDJoueur, ToutJoueurs
End Sub
'Public Sub Effet_ApparencePerso(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
'    ComReseau.Message_Envoyer Categorie1EffetsVisuels, Categorie2EffetApparencePerso, Indice, , EffetsPersos(Indice).Indice, , IDJoueur, ToutJoueurs
'End Sub
'Public Sub Effet_PositionsPerso(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
'    ComReseau.Message_Envoyer Categorie1EffetsVisuels, Categorie2EffetPositionsPerso, Indice, , EffetsPersos(Indice).PositionX, EffetsPersos(Indice).PositionY, IDJoueur, ToutJoueurs
'End Sub
'Public Sub Effet_ApparenceMaison(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
'    ComReseau.Message_Envoyer Categorie1EffetsVisuels, Categorie2EffetApparenceMaison, Indice, , EffetsMaisons(Indice).Indice, , IDJoueur, ToutJoueurs
'End Sub
'Public Sub Effet_PositionsMaison(ByVal Indice As Long, ByVal IDJoueur As Long, Optional ByVal ToutJoueurs As Boolean)
'    ComReseau.Message_Envoyer Categorie1EffetsVisuels, Categorie2EffetPositionsMaison, Indice, , EffetsMaisons(Indice).PositionX, EffetsMaisons(Indice).PositionY, IDJoueur, ToutJoueurs
'End Sub

' ************************************************************************************
' *                        Fonctions qui analysent les messages                      *
' ************************************************************************************
Public Sub Analyser(ByVal Categorie1 As Integer, _
                    ByVal Categorie2 As Integer, _
                    Optional Indice1 As Long, _
                    Optional Indice2 As Long, _
                    Optional ByVal Var1 As String, _
                    Optional ByVal Var2 As String, _
                    Optional ByVal ChatSeulement As Boolean)
    'Messages servants à chatter.
    'If Not ChatSeulement Or (Categorie1 = Categorie1MessageChat Or Categorie1 = Categorie1Partie) Then
    If Not ChatSeulement Or (Categorie1 = Categorie1MessageChat Or Categorie1 = Categorie1Infos Or Categorie1 = Categorie1Partie) Then
    Select Case Categorie1
    Case Categorie1MessageChat:
        FrmParam.LstMessages.AddItem Var1
        'Beep
    Case Categorie1Commandes:
        Ajouter_Commande Categorie2, Indice1, Var1, Var2
    'Messages de demande d'information.
    Case Categorie1Infos:
        Analyser_Infos Categorie2, Indice1, Indice2, Var1, Var2
    'Messages d'information de la partie.
    Case Categorie1Partie:
        Select Case Categorie2
        Case Categorie2PartieMode: Partie.IndiceMode = Var1
        Case Categorie2PartieEnCours: ComReseau.Afficher_Information_Reseau_PartieEnCours Var1
                                      Partie.EnCours = Var1
        End Select
    'Messages d'information du monde.
    Case Categorie1Monde:
        Analyser_Monde Categorie2, Indice1, Indice2, Var1, Var2
    'Messages d'information sur les ressources.
    Case Categorie1Ressource:
        Analyser_Ressource Categorie2, Indice1, Indice2, Var1, Var2
    'Messages d'information sur les personnages.
    Case Categorie1Perso:
        Analyser_Perso Categorie2, Indice1, Indice2, Var1, Var2
    'Messages d'information sur les maisons.
    Case Categorie1Maison:
        Analyser_Maison Categorie2, Indice1, Indice2, Var1, Var2
    'Messages d'information sur les châteaux.
    Case Categorie1Chateau:
        Analyser_Chateau Categorie2, Indice1, Indice2, Var1, Var2
    'Messages d'information sur les fiefs.
    Case Categorie1Fief:
        Analyser_Fief Categorie2, Indice1, Indice2, Var1, Var2
    'Messages d'information sur les trésors.
    Case Categorie1Tresor:
        Analyser_Tresor Categorie2, Indice1, Indice2, Var1, Var2
    'Messages d'information sur les décors.
    Case Categorie1Decor: Analyser_Decor Categorie2, Indice1, Indice2, Var1, Var2
    Case Categorie1Degat: Analyser_Degat Categorie2, Indice1, Indice2, Var1, Var2
    Case Categorie1EffetsVisuels: Analyser_Effet Categorie2, Indice1, Indice2, Var1, Var2
    End Select
    End If
End Sub

Private Sub Analyser_Infos(ByVal Categorie2 As Integer, _
                          Optional Indice1 As Long, _
                          Optional Indice2 As Long, _
                          Optional ByVal Var1 As String, _
                          Optional ByVal Var2 As String)
    Dim i As Long
    'Actualiser_Memoire
    Select Case Categorie2
    Case Categorie2InfosPartie: Partie_ Indice2
    Case Categorie2InfosPartieEnCours: Partie_EnCours Indice2
    Case Categorie2InfosMonde: Monde_Debut Indice2
    Case Categorie2InfosMondeEnCours: 'If Not ComReseau.Joueur_InfosAJour(Indice2) Then Monde_EnCours Indice2
                                      ComReseau.Joueur_InfosAJour(Indice2) = True
    Case Categorie2InfosPersoCaracs: If Not ComReseau.Joueur_InfosAJour(Indice2) Then Perso_Caracs Indice1, Indice2
    Case Categorie2InfosPersoCaracs2: If Not ComReseau.Joueur_InfosAJour(Indice2) Then Perso_Caracs2 Indice1, Indice2
    Case Categorie2InfosPersoListeRessources: If Not ComReseau.Joueur_InfosAJour(Indice2) Then Perso_ListeRessources Indice1, Indice2
    Case Categorie2InfosPersoInventaire: If Not ComReseau.Joueur_InfosAJour(Indice2) Then Perso_Inventaire Indice1, Indice2
    Case Categorie2InfosPersoInventaireMaison: If Not ComReseau.Joueur_InfosAJour(Indice2) Then Perso_InventaireMaison Indice1, Indice2
    Case Categorie2InfosPersoCompetences: If Not ComReseau.Joueur_InfosAJour(Indice2) Then Perso_Competences Indice1, Indice2
    Case Categorie2InfosPersoCompetences2: If Not ComReseau.Joueur_InfosAJour(Indice2) Then Perso_Competences2 Indice1, Indice2
    Case Categorie2InfosPersoListeBestiaire: If Not ComReseau.Joueur_InfosAJour(Indice2) Then Perso_ListeBestiaire Indice1, Indice2
    Case Categorie2InfosMaison: If Not ComReseau.Joueur_InfosAJour(Indice2) Then Maison_ListeStock Indice1, Indice2
    Case Categorie2InfosChateau: If Not ComReseau.Joueur_InfosAJour(Indice2) Then Chateau_ListeStock Indice1, Indice2
    
    Case Categorie2InfosPersoIndiceNouveau: 'Lu par le serveur.
        'Créé un nouveau joueur et renvoie son indice.
        i = Perso_Ajouter_Nouveau(Indice1)
        Persos(i).Changer_Race_Personnage_Equipement Parametres, Commentaires
        Actualiser_Infos_Fiefs 'Réactualise les nombres de citoyens des fiefs.
        Infos_Perso_EstTonJoueur i, Var1, Indice1
        'Actualiser_Memoire
        'MsgBox "Nouveau personnage demandé : Numero = " & i
    Case Categorie2InfosPersoIndiceSupprimer: 'Lu par le serveur quand le client quitte la partie.
        'Supprime tous les personnages du joueur dont l'ID est donnée.
        Perso_Supprimer_Nouveau Indice1
    Case Categorie2InfosPersoEstTonJoueur: 'Lu par le client.
        'Reçoit l'indice d'un joueur dans lequel on peut jouer. Et un insert notre personnage.
        ChargementIndicePerso = Indice1
        NoPerso = Indice1
        Nopersos(0) = Indice1
        'Charger_Sauvegarde_Perso Var1, Persos(Indice1)
        Perso_Envoyer 0, Indice1, ComReseau.ServeurID
        'MsgBox "Personnage envoyé"
    Case Categorie2InfosPersoChargementTermine: 'Lu par le client.
        'Le serveur a fini de charger le personnage du joueur.
        ChargementPersoTermine = True
        'MsgBox "Personnage chargement terminé"
    End Select
End Sub

Private Sub Analyser_Monde(ByVal Categorie2 As Integer, _
                          Optional Indice1 As Long, _
                          Optional Indice2 As Long, _
                          Optional ByVal Var1 As String, _
                          Optional ByVal Var2 As String)
    Select Case Categorie2
    Case Categorie2MondeDimensions:
        Monde.Largeur = Indice1: Monde.Hauteur = Indice2
        Monde.NombreCases = Var1
    Case Categorie2MondeCarte:
        Monde.Largeur = Indice1: Monde.Hauteur = Indice2
        Analyser_Carte Var1
        Monde.Reformer_Terrain
        ChargementTermine = True
    Case Categorie2MondeTerrain:
        Monde.Param_TerrainCategorie(Indice1) = Indice2
    Case Categorie2MondeTerrainFin:
        Monde.Reformer_Terrain
        ChargementTermine = True
    Case Categorie2MondeNomImageCarte:
        Cartes.NomImageCarte = Var1
        Cartes.Chargee = True
    End Select
End Sub

Private Sub Analyser_Carte(ByVal Carte As String)
    Dim i As Long
    With Monde
    For i = 0 To Len(Carte) - 1
        '.Param_TerrainCategorie(i) = Asc(Left(Carte, 1))
        'Carte = Right(Carte, Len(Carte) - 1)
        .Param_TerrainCategorie(i) = Asc(Mid(Carte, i + 1, 1))
    Next i
    End With
End Sub

Private Sub Analyser_Ressource(ByVal Categorie2 As Integer, _
                              Optional Indice1 As Long, _
                              Optional Indice2 As Long, _
                              Optional ByVal Var1 As String, _
                              Optional ByVal Var2 As String)
    Dim i As Long
    Select Case Categorie2
    Case Categorie2RessourceNombre:
        ReDim Preserve Ressources(Indice1)
        For i = 0 To UBound(Ressources())
            Set Ressources(i) = New ClsJeuRessources
        Next i
    Case Categorie2RessourceInfos:
        Ressources(Indice1).PositionX = Var1
        Ressources(Indice1).PositionY = Var2
        Ressources(Indice1).Definir_Type = Indice2
    End Select
End Sub

Private Sub Analyser_Perso(ByVal Categorie2 As Integer, _
                          Optional Indice1 As Long, _
                          Optional Indice2 As Long, _
                          Optional ByVal Var1 As String, _
                          Optional ByVal Var2 As String)
    Dim i As Long
    Select Case Categorie2
    Case Categorie2PersoNombre:
        ReDim Persos(Indice1)
        ReDim Maisons(Indice1)
        ReDim EffetsPersos(UBound(Persos()))
        ReDim EffetsMaisons(UBound(Maisons()))
        For i = 0 To UBound(Persos())
            Set Persos(i) = New ClsJeuPerso
            Persos(i).Init i, Parametres
            Set Maisons(i) = New ClsJeuBatiment
            Set EffetsMaisons(i) = New ClsJeuEffet
            Set EffetsPersos(i) = New ClsJeuEffet
        Next i
    Case Categorie2PersoNom: Persos(Indice1).Nom = Var1
    Case Categorie2PersoFief: Persos(Indice1).NumeroFief = Var1
    Case Categorie2PersoRace: Persos(Indice1).Changer_Race_Personnage Var1, Parametres, True
    Case Categorie2PersoPositions:
        Persos(Indice1).PositionX = Var1: Persos(Indice1).PositionY = Var2
        Persos(Indice1).EnDeplacement = True
    Case Categorie2PersoVie: Persos(Indice1).Vie = Var1: Persos(Indice1).BonusMaxVie = Var2
    Case Categorie2PersoVie2: Persos(Indice1).RecuperationVie = Var1
    Case Categorie2PersoVie3: Persos(Indice1).MaxVie = Var1
    Case Categorie2PersoVieExp: Persos(Indice1).ExpVie = Var1
    Case Categorie2PersoEnergie: Persos(Indice1).Energie = Var1: Persos(Indice1).BonusMaxEnergie = Var2
    Case Categorie2PersoEnergie2: Persos(Indice1).RecuperationEnergie = Var1
    Case Categorie2PersoEnergie3: Persos(Indice1).MaxEnergie = Var1
    Case Categorie2PersoEnergieExp: Persos(Indice1).ExpEnergie = Var1
    Case Categorie2PersoMagie: Persos(Indice1).Magie = Var1: Persos(Indice1).BonusMaxMagie = Var2
    Case Categorie2PersoMagie2: Persos(Indice1).RecuperationMagie = Var1
    Case Categorie2PersoMagie3: Persos(Indice1).MaxMagie = Var1
    Case Categorie2PersoMagieExp: Persos(Indice1).ExpMagie = Var1
    Case Categorie2PersoMoral: Persos(Indice1).Moral = Var1: Persos(Indice1).BonusMaxMoral = Var2
    Case Categorie2PersoMoral2: Persos(Indice1).RecuperationMoral = Var1
    Case Categorie2PersoMoral3: Persos(Indice1).MaxMoral = Var1
    Case Categorie2PersoMoralExp: Persos(Indice1).ExpMoral = Var1
    Case Categorie2PersoAttaque: Persos(Indice1).Attaque = Var1: Persos(Indice1).BonusMaxAttaque = Var2
    Case Categorie2PersoAttaque2: Persos(Indice1).MaxAttaque = Var1
    Case Categorie2PersoAttaqueRecharge: Persos(Indice1).DureeRechargeAttaque = Var1
                                         Persos(Indice1).DureeRechargeAttaqueMax = Var2
    Case Categorie2PersoAttaqueExp: Persos(Indice1).ExpAttaque = Var1
    Case Categorie2PersoDefense: Persos(Indice1).Defense = Var1: Persos(Indice1).BonusMaxDefense = Var2
    Case Categorie2PersoDefense2: Persos(Indice1).MaxDefense = Var1
    Case Categorie2PersoDefenseExp: Persos(Indice1).ExpDefense = Var1
    Case Categorie2PersoJoueur: Persos(Indice1).Joueur = Var1
    Case Categorie2PersoEquipe: Persos(Indice1).NumeroEquipe = Var1
    Case Categorie2PersoIA: Persos(Indice1).IA = Var1
    'Case Categorie2PersoPas: Persos(Indice1).NumeroPas = Var1
    Case Categorie2PersoDirection: Persos(Indice1).Direction = Var1
    Case Categorie2PersoFeminin: Persos(Indice1).Feminin = Var1
    'Case Categorie2PersoVivant: Persos(Indice1).Vivant = Var1
    Case Categorie2PersoChoisirApparence: Persos(Indice1).ChoisirApparence = Var1
    Case Categorie2PersoFichierApparence: Persos(Indice1).FichierApparence = Var1
    Case Categorie2PersoApparenceChoisie: Persos(Indice1).ApparenceChoisie = Var1
    Case Categorie2PersoNombre_ApparenceBase: Persos(Indice1).Nombre_ApparenceBase = Var1
    Case Categorie2PersoDefinir_ApparenceBase: Persos(Indice1).Definir_ApparenceBase(Indice2) = Var1
    Case Categorie2PersoDefinir_ApparenceDecalage: Persos(Indice1).Definir_ApparenceDecalageX(Indice2) = Var1
                                                   Persos(Indice1).Definir_ApparenceDecalageX(Indice2) = Var2
    Case Categorie2PersoDefinir_ApparenceAvantPlan: Persos(Indice1).Definir_ApparenceAvantPlan(Indice2) = Var1
    Case Categorie2PersoDefinir_FichierApparenceBase:
        Persos(Indice1).Definir_FichierApparenceBase(Indice2) = Var1
    Case Categorie2PersoApparenceObjetAfficher: Persos(Indice1).ApparenceObjetAfficher = Var1
    Case Categorie2PersoNombreApparenceObjet: Persos(Indice1).Nombre_ApparenceObjet = Indice2
    Case Categorie2PersoDefinir_ApparenceObjet: Persos(Indice1).Definir_ApparenceObjet(Indice2) = Var1
    Case Categorie2PersoDansUneMaison: Persos(Indice1).DansSaMaison = Var1: Persos(Indice1).DansUneMaison = Var2
    Case Categorie2PersoDansUnChateau: Persos(Indice1).DansSonChateau = Var1: Persos(Indice1).DansUnChateau = Var2
    Case Categorie2PersoChangerNombre: While Indice1 > UBound(Persos()):           Perso_Ajouter_Nouveau:        Wend
    Case Categorie2PersoChoisirIA: Persos(Indice1).ChoisirIA = Var1: Persos(Indice1).FichierIA = Var2
    Case Categorie2PersoChoisirIAScript: Persos(Indice1).ChoisirIAScript = Var1: Persos(Indice1).FichierIAScript = Var2
    Case Categorie2PersoNumeroChef: Persos(Indice1).NumeroChef = Var1
    Case Categorie2PersoNombreSoldats: Persos(Indice1).NombreSoldats = Var1
    Case Categorie2PersoAttitudeFormation: Persos(Indice1).Attitude = Var1: Persos(Indice1).Formation = Var2
    Case Categorie2PersoExpVie: Persos(Indice1).ExpVie = Var1
    Case Categorie2PersoExpEnergie: Persos(Indice1).ExpEnergie = Var1
    Case Categorie2PersoExpMagie: Persos(Indice1).ExpMagie = Var1
    Case Categorie2PersoExpMoral: Persos(Indice1).ExpMoral = Var1
    Case Categorie2PersoExpAttaque: Persos(Indice1).ExpAttaque = Var1
    Case Categorie2PersoExpDefense: Persos(Indice1).ExpDefense = Var1
    Case Categorie2PersoCompetencesSpeciales: Persos(Indice1).Definir_Carac_CompetenceSpeciales(Indice2) = Var1
                                              Persos(Indice1).Definir_Niveau_CompetenceSpeciales(Indice2) = Var2
    Case Categorie2PersoCompetencesRessources: Persos(Indice1).Definir_Carac_Ressources(Indice2) = Var1
                                               Persos(Indice1).Definir_Niveau_CompetenceRessources(Indice2) = Var2
    Case Categorie2PersoCompetencesServices: Persos(Indice1).Definir_Carac_CompetenceServices(Indice2) = Var1
                                             Persos(Indice1).Definir_Niveau_CompetenceServices(Indice2) = Var2
    Case Categorie2PersoCompetencesObjets: Persos(Indice1).Definir_Carac_CompetenceObjets(Indice2) = Var1
                                           Persos(Indice1).Definir_Niveau_CompetenceObjets(Indice2) = Var2
    Case Categorie2PersoCompetencesSpecialesExp: Persos(Indice1).Definir_ExpSpeciales(Indice2) = Var1
    Case Categorie2PersoCompetencesRessourcesExp: Persos(Indice1).Definir_ExpRessources(Indice2) = Var1
    Case Categorie2PersoCompetencesServicesExp: Persos(Indice1).Definir_ExpServices(Indice2) = Var1
    Case Categorie2PersoCompetencesObjetsExp: Persos(Indice1).Definir_ExpObjets(Indice2) = Var1
    Case Categorie2PersoBestiaire: Persos(Indice1).Definir_Bestiaire(Indice2) = Var1
    Case Categorie2PersoAction: Persos(Indice1).Action = Var1
    Case Categorie2PersoEtatFabricationObjet: Persos(Indice1).IndiceObjet = Var1
                                              Persos(Indice1).EtatFabricationObjet = Var2
    Case Categorie2PersoCibles1: Persos(Indice1).RessourceCourante = Var1
                                 Persos(Indice1).IndiceRessource = Var2
    Case Categorie2PersoCibles2: Persos(Indice1).IndicePerso = Var1
                                 Persos(Indice1).IndiceTresor = Var2
    Case Categorie2PersoCibles3: Persos(Indice1).IndiceMaison = Var1
                                 Persos(Indice1).IndiceChateau = Var2
    Case Categorie2PersoArgent: Persos(Indice1).Definir_Argent = Var1
                                Persos(Indice1).ArgentSelectionne = Var2
    Case Categorie2PersoRessources: Persos(Indice1).Definir_ressources(Indice2) = Var1
                                    Persos(Indice1).Definir_MaxRessources(Indice2) = Var2
    Case Categorie2PersoRessourcesSelectionnees: Persos(Indice1).Definir_RessourcesSelectionnes(Indice2) = Var1
    Case Categorie2PersoObjetSelectionne: Persos(Indice1).ObjetSelectionne = Var1
    Case Categorie2PersoObjetsEquipes: Persos(Indice1).Objet_Equipes_Type(Indice2) = Var1
    Case Categorie2PersoObjetsInventaire: Persos(Indice1).Objet_Inventaire_Type(Indice2) = Var1
                                          Persos(Indice1).Objet_Inventaire_Nombre(Indice2) = Var2
    Case Categorie2PersoArmure: Persos(Indice1).Definir_Armure = Var1
    Case Categorie2PersoVitalite: Persos(Indice1).Vitalite = Var1
    Case Categorie2PersoVitesse: Persos(Indice1).Longueur_PasTemp = Var1
                                 Persos(Indice1).BonusDeplacement = Var2
    Case Categorie2PersoEffetTempNombre: Persos(Indice1).Nombre_EffetsTemp = Var1
                                         Persos(Indice1).Nombre_EffetsRecharge = Var2
    Case Categorie2PersoEffetTempNumero: Persos(Indice1).Definir_EffetTemp_Numero(Indice2) = Var1
    Case Categorie2PersoEffetTempDuree: Persos(Indice1).Definir_EffetTemp_Duree(Indice2) = Var1
    Case Categorie2PersoEffetTempProprietaire: Persos(Indice1).Definir_EffetTemp_Proprietaire(Indice2) = Var1
    Case Categorie2PersoEffetTempAjouter: Persos(Indice1).EffetTemp_Ajouter Indice2, CLng(Left(Var1, InStr(Var1, "/") - 1)), CLng(Right(Var1, InStr(Var1, "/") - 1)), CLng(Var2)
    Case Categorie2PersoEffetTempSupprimer: Persos(Indice1).EffetTemp_Effacer Indice2
    Case Categorie2PersoEffetRecharge: Persos(Indice1).Definir_EffetsRecharge_Numero(Indice2) = Var1
                                       Persos(Indice1).Definir_EffetsRecharge_Duree(Indice2) = Var2
    Case Categorie2PersoEffetRechargeAjouter: Persos(Indice1).EffetRecharge_Ajouter Indice2, Var1
    Case Categorie2PersoEffetRechargeSupprimer: Persos(Indice1).EffetRecharge_Effacer Indice2
    Case Categorie2PersoCommentaires: Persos(Indice1).Definir_Commentaires = Var1
    Case Categorie2PersoDureeResurrection: Persos(Indice1).DureeRessurection = Var1
    Case Categorie2PersoActualiser:
        Persos(Indice1).Gestion_Tous_Niveaux
        Persos(Indice1).Actualiser_Bonus
        Persos(Indice1).Ressources_Verifications
    Case Categorie2PersoActualiserBonus:
        Persos(Indice1).Gestion_Tous_Niveaux
        Persos(Indice1).Actualiser_Bonus
        Perso_Finalier Indice1
        Persos(Indice1).NumeroEquipe = Fiefs(Persos(Indice1).NumeroFief).NumeroEquipe
        Persos(Indice1).PositionX = Chateaux(Persos(Indice1).NumeroFief).PositionX + Chateaux(Persos(Indice1).NumeroFief).Largeur / 2 - Persos(Indice1).Largeur / 2 + Chateaux(Persos(Indice1).NumeroFief).EntreePositionX
        Persos(Indice1).PositionY = Chateaux(Persos(Indice1).NumeroFief).PositionY + Chateaux(Persos(Indice1).NumeroFief).Hauteur / 2 - Persos(Indice1).Hauteur / 2 + Chateaux(Persos(Indice1).NumeroFief).EntreePositionY
        Perso_Charger_Ressources_Depart Persos(Indice1)
        Persos(Indice1).DansSonChateau = True
        Infos_Perso_ChargementTermine Indice1, Indice2
    Case Categorie2PersoEnregistrer: Enregistrer_Sauvegarde_Perso Persos(Indice1)
    End Select
End Sub

Private Sub Analyser_Maison(ByVal Categorie2 As Integer, _
                           Optional Indice1 As Long, _
                           Optional Indice2 As Long, _
                           Optional ByVal Var1 As String, _
                           Optional ByVal Var2 As String)
    On Error GoTo Erreur
    Dim i As Long
    Select Case Categorie2
    Case Categorie2MaisonType: If Maisons(Indice1).TypeBatiment <> Indice2 Or Indice2 = 0 Then Maisons(Indice1).Definir_TypeBatiment = Indice2
    Case Categorie2MaisonEnConstruction: Maisons(Indice1).EnConstruction = Var1
                                         Maisons(Indice1).Construit = Var2
    Case Categorie2MaisonPositions: Maisons(Indice1).PositionX = Var1
                                    Maisons(Indice1).PositionY = Var2
    Case Categorie2MaisonVie: Maisons(Indice1).MaxVie = Var2
                              Maisons(Indice1).Definir_Vie = Var1
    Case Categorie2MaisonStock: Maisons(Indice1).Definir_Stock = Var1
                                Maisons(Indice1).MaxStock = Var2
    Case Categorie2MaisonStockObjet: Maisons(Indice1).Stock_Objets(Indice2) = Var1
    Case Categorie2MaisonInventaire: Maisons(Indice1).Definir_ObjetsInventaire(Indice2) = Var1
    End Select
    Exit Sub
Erreur:
End Sub

Private Sub Analyser_Chateau(ByVal Categorie2 As Integer, _
                           Optional Indice1 As Long, _
                           Optional Indice2 As Long, _
                           Optional ByVal Var1 As String, _
                           Optional ByVal Var2 As String)
    Select Case Categorie2
    Case Categorie2ChateauPositions: Chateaux(Indice1).PositionX = Var1
                                     Chateaux(Indice1).PositionY = Var2
    Case Categorie2ChateauVie: Chateaux(Indice1).Definir_Vie = Var1
                               Chateaux(Indice1).MaxVie = Var2
    Case Categorie2ChateauRecuperationVie: Chateaux(Indice1).RecuperationVie = Var1
    Case Categorie2ChateauRessources: Fiefs(Indice1).Definir_ressources(Indice2) = Var1
    End Select
End Sub

Private Sub Analyser_Fief(ByVal Categorie2 As Integer, _
                         Optional Indice1 As Long, _
                         Optional Indice2 As Long, _
                         Optional ByVal Var1 As String, _
                         Optional ByVal Var2 As String)
    Dim i As Long
    Select Case Categorie2
    Case Categorie2FiefNombre:
        ReDim Fiefs(Indice1)
        ReDim Chateaux(Indice1)
        For i = 0 To UBound(Fiefs())
            Set Fiefs(i) = New ClsJeuFief
            Fiefs(i).Init i, Parametres
            Set Chateaux(i) = New ClsJeuBatiment
        Next i
    Case Categorie2FiefPeuple:
        Fiefs(Indice1).Definir_TypePeuple = Var1
        Chateaux(Indice1).Largeur = Parametres.Peuples_ChateauLargeur(Fiefs(Indice1).Definir_TypePeuple)
        Chateaux(Indice1).Hauteur = Parametres.Peuples_ChateauHauteur(Fiefs(Indice1).Definir_TypePeuple)
        For i = 0 To Fiefs(Indice1).Epoque
            Chateaux(Indice1).MaxVie = Chateaux(i).MaxVie + Parametres.Epoque_BonusVieChateau(i)
        Next i
    Case Categorie2FiefEpoque: Fiefs(Indice1).Epoque = Var1
                               'Fiefs(Indice1).TempsChangementEpoque = Var2
    Case Categorie2FiefEquipe: Fiefs(Indice1).NumeroEquipe = Var1
    Case Categorie2FiefTempsChangementEpoque: Fiefs(Indice1).TempsChangementEpoque = Var1
    Case Categorie2FiefNombreCitoyens: Fiefs(Indice1).NombreCitoyens = Var1
                                       Fiefs(Indice1).NombreMorts = Var2
    Case Categorie2FiefEmpereur: Fiefs(Indice1).Empereur_Changer Var1
    End Select
End Sub

Private Sub Analyser_Tresor(ByVal Categorie2 As Integer, _
                           Optional Indice1 As Long, _
                           Optional Indice2 As Long, _
                           Optional ByVal Var1 As String, _
                           Optional ByVal Var2 As String)
    Dim i As Long
    Select Case Categorie2
    Case Categorie2TresorNombre: Tresor_DefinirNombre = Var1
    Case Categorie2TresorApparence: Tresors(Indice1).IndiceImage = Var1
    Case Categorie2TresorPositions: Tresors(Indice1).X = Var1
                                    Tresors(Indice1).Y = Var2
    Case Categorie2TresorAjouter: Tresor_Creer Var1, Var2
    Case Categorie2TresorSupprimer: Tresors.Remove Indice1
    End Select
End Sub

Private Sub Analyser_Decor(ByVal Categorie2 As Integer, _
                          Optional Indice1 As Long, _
                          Optional Indice2 As Long, _
                          Optional ByVal Var1 As String, _
                          Optional ByVal Var2 As String)
    Dim i As Long
    Select Case Categorie2
    Case Categorie2DecorNombre: Decor_DefinirNombre = Var1
    Case Categorie2DecorNom: Decors(Indice1).Nom = Var1
    Case Categorie2DecorApparence: Decors(Indice1).Categorie = Var1
                                   Decors(Indice1).EtapeAnimation = Var2
    Case Categorie2DecorPositions: Decors(Indice1).X = Var1
                                   Decors(Indice1).Y = Var2
    Case Categorie2DecorAjouter: Decor_Creer Var1, Indice1, Indice2, Var2
    Case Categorie2DecorSupprimer: If Indice1 <= Decors.Count Then Decors.Remove Indice1
    End Select
End Sub

Private Sub Analyser_Degat(ByVal Categorie2 As Integer, _
                          Optional Indice1 As Long, _
                          Optional Indice2 As Long, _
                          Optional ByVal Var1 As String, _
                          Optional ByVal Var2 As String)
    Dim i As Long
    Select Case Categorie2
    Case Categorie2DegatTexte: AffDegat.Ajouter_Degat 0, 0, Var1
    Case Categorie2DegatPositions: AffDegat.Degats(AffDegat.Degats.Count).X = Var1
                                   AffDegat.Degats(AffDegat.Degats.Count).Y = Var2
    Case Categorie2DegatVisuel: AffDegat.Degats(AffDegat.Degats.Count).Couleur = Var1
                                AffDegat.Degats(AffDegat.Degats.Count).Critique = CBool(Var2)
    End Select
End Sub

Private Sub Analyser_Effet(ByVal Categorie2 As Integer, _
                          Optional Indice1 As Long, _
                          Optional Indice2 As Long, _
                          Optional ByVal Var1 As String, _
                          Optional ByVal Var2 As String)
    Dim i As Long
    Select Case Categorie2
    Case Categorie2EffetApparence: Activer_Effet Var1, 0, 0, 0, 0, Parametres
    Case Categorie2EffetPositions: Effets(Effets.Count).PositionX = Var1
                                   Effets(Effets.Count).PositionY = Var2
    'Case Categorie2EffetApparencePerso: EffetsPersos(Effets.Count).Activer_Effet Var1, 0, 0, 0, 0, Parametres
    'Case Categorie2EffetPositionsPerso: EffetsPersos(Effets.Count).X = Var1
    '                                    EffetsPersos(Effets.Count).Y = Var2
    'Case Categorie2EffetApparence: EffetsMaisons.Activer_Effet Var1, 0, 0, 0, 0, Parametres
    'Case Categorie2EffetPositions: EffetsMaisons(Effets.Count).X = Var1
    '                               EffetsMaisons(Effets.Count).Y = Var2
    End Select
End Sub

Public Sub Actualiser_Memoire()
    Dim i As Long
    Dim X As Long
    Dim Y As Long
'    Dim TempTresor As ClsComTresor
    'Dim TempDecor As ClsComDecor
    X = UBound(Persos())
    Y = NombreComPersos - 1
    If X > Y Then
        ReDim Preserve ComPersos(X)
        For i = Y To X
            Set ComPersos(i) = New ClsComPerso
            ComPersos(i).Init Parametres
        Next i
    End If
    X = UBound(Fiefs())
    Y = NombreComFiefs - 1
    If X > Y Then
        ReDim Preserve ComFiefs(X)
        For i = Y To X
            Set ComFiefs(i) = New ClsComFief
            ComFiefs(i).Init Parametres
        Next i
        NombreComFiefs = X + 1
    End If
'    If Tresors.Count > NombreComTresors Then
'        Do
'            ComTresors.Remove 1
'            NombreComTresors = NombreComTresors - 1
'        Loop Until Tresors.Count = NombreComTresors
'    ElseIf Tresors.Count < NombreComTresors Then
'        Do
'            Set TempTresor = New ClsComTresor
'            ComTresors.Add TempTresor
'            NombreComTresors = NombreComTresors + 1
'            Set TempTresor = Nothing
'        Loop Until Tresors.Count = NombreComTresors
'    End If
'    If Decors.Count > NombreComDecors Then
'        Do
'            ComDecors.Remove 1
'            NombreComDecors = NombreComDecors - 1
'        Loop Until Decors.Count = NombreComDecors
'    ElseIf Decors.Count < NombreComDecors Then
'        Do
'            Set TempDecor = New ClsComDecor
'            ComDecors.Add TempDecor
'            NombreComDecors = NombreComDecors + 1
'            Set TempDecor = Nothing
'        Loop Until Decors.Count = NombreComDecors
'    End If
End Sub

Public Property Get Definir_CategorieCommandes() As Integer
    Definir_CategorieCommandes = Categorie1Commandes
End Property

