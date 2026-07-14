Attribute VB_Name = "ModSAVMonde"
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

Const FichierINI = "SavMonde"
Const SectionINI = "SavMonde"

Private SavMondeEnQuittant As Boolean
Private SavImageEnQuittant As Boolean
Public SavMondeEnQuittantNom As String

Private FichierVierge As String
Public CheminSavMonde As String
Public ExtensionFichiersMonde As String
Public ExtensionFichiersAperçu As String
Public ExtensionFichiersTexte  As String

Private InfoDate As String
Private InfoHeure As String
Private InfoDimensions As String
Private InfoJoueur As String
Private InfoJoueursMax As String
Private InfoPeuple As String
Private InfoEpoque As String
Private InfoFiefs As String
Private InfoPersonnages As String

Public Sub Init_Sauvegarde_Monde()
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    
    SavMondeEnQuittant = Val(FicIni.Parametre("SavMondeEnQuittant"))
    SavImageEnQuittant = Val(FicIni.Parametre("SavImageEnQuittant"))
    SavMondeEnQuittantNom = FicIni.Parametre("SavMondeEnQuittantNom")
    FichierVierge = FicIni.Parametre("FichierVierge")
    CheminSavMonde = FicIni.Parametre("CheminSavMonde")
    ExtensionFichiersMonde = FicIni.Parametre("ExtensionFichiersMonde")
    ExtensionFichiersAperçu = FicIni.Parametre("ExtensionFichiersAperçu")
    ExtensionFichiersTexte = FicIni.Parametre("ExtensionFichiersTexte")
    
    InfoDate = FicIni.Parametre("InfoDate")
    InfoHeure = FicIni.Parametre("InfoHeure")
    InfoDimensions = FicIni.Parametre("InfoDimensions")
    InfoJoueur = FicIni.Parametre("InfoJoueur")
    InfoJoueursMax = FicIni.Parametre("InfoJoueursMax")
    InfoPeuple = FicIni.Parametre("InfoPeuple")
    InfoEpoque = FicIni.Parametre("InfoEpoque")
    InfoFiefs = FicIni.Parametre("InfoFiefs")
    InfoPersonnages = FicIni.Parametre("InfoPersonnages")
End Sub

Public Sub Charger_Sauvegarde_Monde(ByVal Nom As String)
    Dim i As Long
    Dim j As Integer, k As Integer
    Dim Temp, Temp2 As String
    Dim Temp3, Temp4 As String
    'On Error GoTo Erreur
    
    Open App.Path & CheminSavMonde & Nom For Input As #1
    
    'Sauvegarde les paramètres pour vérification.
    Line Input #1, Temp
    If Parametres.NombreCompetencesObjets <> Val(Temp) Then GoTo ErreurParam
    Line Input #1, Temp
    If Parametres.NombreCompetencesSpeciales <> Val(Temp) Then GoTo ErreurParam
    Line Input #1, Temp
    If Parametres.NombreDecors <> Val(Temp) Then GoTo ErreurParam
    Line Input #1, Temp
    If Parametres.NombreEffets <> Val(Temp) Then GoTo ErreurParam
    Line Input #1, Temp
    If Parametres.NombreEpoques <> Val(Temp) Then GoTo ErreurParam
    Line Input #1, Temp
    If Parametres.NombreObjets <> Val(Temp) Then GoTo ErreurParam
    Line Input #1, Temp
    If Parametres.NombrePeuples <> Val(Temp) Then GoTo ErreurParam
    Line Input #1, Temp
    If Parametres.NombreRaces <> Val(Temp) Then GoTo ErreurParam
    Line Input #1, Temp
    If Parametres.NombreRacesJouables <> Val(Temp) Then GoTo ErreurParam
    Line Input #1, Temp
    If Parametres.NombreResistances <> Val(Temp) Then GoTo ErreurParam
    Line Input #1, Temp
    If Parametres.NombreRessources <> Val(Temp) Then GoTo ErreurParam
    Line Input #1, Temp
    If Parametres.NombreRessourcesDepart <> Val(Temp) Then GoTo ErreurParam
    Line Input #1, Temp
    If Parametres.NombreServices <> Val(Temp) Then GoTo ErreurParam
    Line Input #1, Temp
    If Parametres.NombreStatuts <> Val(Temp) Then GoTo ErreurParam
    Line Input #1, Temp
    If Parametres.NombreTerrainRessource <> Val(Temp) Then GoTo ErreurParam
    Line Input #1, Temp
    If Parametres.NombreTresors <> Val(Temp) Then GoTo ErreurParam
    Line Input #1, Temp
    If Parametres.NombreTypeBatiments <> Val(Temp) Then GoTo ErreurParam
    Line Input #1, Temp
    If Parametres.NombreTypeBatimentsJouables <> Val(Temp) Then GoTo ErreurParam
    Line Input #1, Temp
    If Parametres.NombreVitesseEpoques <> Val(Temp) Then GoTo ErreurParam
    
    Line Input #1, Temp
    ReDim Fiefs(Val(Temp) - 1)
    Line Input #1, Temp
    ReDim Chateaux(Val(Temp) - 1)
    Line Input #1, Temp
    ReDim Ressources(Val(Temp) - 1)
    Line Input #1, Temp
    For i = 1 To Val(Temp)
        Tresor_Creer 0, 0
    Next i
    Line Input #1, Temp
    For i = 1 To Val(Temp)
        Decor_Creer 0, 0, 0
    Next i
    Line Input #1, Temp
    'ReDim Preserve Persos(Val(Temp) - 1)
    ReDim Persos(Val(Temp) - 1)
    ReDim EffetsPersos(UBound(Persos()))
    Line Input #1, Temp
    ReDim Maisons(Val(Temp) - 1)
    ReDim EffetsMaisons(UBound(Maisons()))
    
    Line Input #1, Temp: Parametres.VitesseEpoqueSelectionnee = Temp
    'Charge la mini-carte.
    Line Input #1, Temp
    If Temp <> "" Then
        Cartes.DerniereCarte = Temp
        Cartes.Chargee = True
    End If
    Line Input #1, Temp: Monde.Largeur = Temp
    Line Input #1, Temp: Monde.Hauteur = Temp
    Monde.Calculer_Diagonale
    
    For i = 0 To (Monde.Largeur + 1) * (Monde.Hauteur + 1) - 1
        Line Input #1, Temp: Monde.Param_TerrainApparence(i) = Temp
        Line Input #1, Temp: Monde.Param_TerrainCategorie(i) = Temp
    Next i
    
    'Paramètres du jeu.
    Line Input #1, Temp: AffTemps.Heures = Temp: AffTemps.Actualiser_Temps
    Line Input #1, Temp: AffTemps.Minutes = Temp
    Line Input #1, Temp: Jeu.Chrono = Temp
    Line Input #1, Temp: Jeu.IntituleChrono = Temp
    
    For i = 0 To UBound(Fiefs())
        Set Fiefs(i) = New ClsJeuFief
        Fiefs(i).Init i, Parametres
        Line Input #1, Temp: Fiefs(i).NumeroEquipe = Temp
        Line Input #1, Temp: Fiefs(i).Definir_TypePeuple = Temp
        Line Input #1, Temp: Fiefs(i).Nom = Temp
        Line Input #1, Temp: Fiefs(i).Epoque = Temp
        Line Input #1, Temp: Fiefs(i).ChangerEpoque = Temp
        Line Input #1, Temp: Fiefs(i).TempsChangementEpoque = Temp
        Line Input #1, Temp: Fiefs(i).NombreCitoyens = Temp
        Line Input #1, Temp: Fiefs(i).NombreMorts = Temp
        Line Input #1, Temp: Fiefs(i).Empereur = Temp
        Line Input #1, Temp: Fiefs(i).Argent = Temp
        For j = 0 To Parametres.NombreRessources - 1
            Line Input #1, Temp: Fiefs(i).Definir_ressources(j) = Temp
        Next j
    Next i
    For i = 0 To UBound(Chateaux())
        Set Chateaux(i) = New ClsJeuBatiment
        Line Input #1, Temp: Chateaux(i).Numero = Temp
        Line Input #1, Temp: Chateaux(i).PositionX = Temp
        Line Input #1, Temp: Chateaux(i).PositionY = Temp
        Line Input #1, Temp: Chateaux(i).Largeur = Temp
        Line Input #1, Temp: Chateaux(i).Hauteur = Temp
        Line Input #1, Temp: Chateaux(i).EntreePositionX = Temp
        Line Input #1, Temp: Chateaux(i).EntreePositionY = Temp
        Line Input #1, Temp: Chateaux(i).Construit = Temp
        Line Input #1, Temp: Chateaux(i).EnConstruction = Temp
        Line Input #1, Temp: Chateaux(i).Vie = Temp
        Line Input #1, Temp: Chateaux(i).RecuperationVie = Temp
        Line Input #1, Temp: Chateaux(i).MaxVie = Temp
    Next i
    For i = 0 To UBound(Ressources())
        Set Ressources(i) = New ClsJeuRessources
        Line Input #1, Temp: Ressources(i).PositionX = Temp
        Line Input #1, Temp: Ressources(i).PositionY = Temp
        Line Input #1, Temp: Ressources(i).Nom = Temp
        Line Input #1, Temp: Ressources(i).Apparence = Temp
        Line Input #1, Temp: Ressources(i).TypeRessource = Temp
        Line Input #1, Temp: Ressources(i).Quantite = Temp
        Line Input #1, Temp: Ressources(i).MaxQuantite = Temp
    Next i
    'Trésors.
    For i = 1 To Tresors.Count
        Line Input #1, Temp: Tresors(i).X = Temp
        Line Input #1, Temp: Tresors(i).Y = Temp
        Line Input #1, Temp: Tresors(i).Definir_Fortune = Temp
        Line Input #1, Temp: Tresors(i).Definir_Argent = Temp
        For j = 0 To Parametres.NombreRessources - 1
            Line Input #1, Temp: Tresors(i).Definir_ressources(j) = Temp
        Next j
        Line Input #1, Temp: Temp2 = Temp
        For j = 1 To Val(Temp2)
            Line Input #1, Temp: Tresors(i).Ajouter_Objet Temp, Parametres
        Next j
    Next i
    'Décors.
    For i = 1 To Decors.Count
        Line Input #1, Temp: Decors(i).Categorie = Temp
        Line Input #1, Temp: Decors(i).X = Temp
        Line Input #1, Temp: Decors(i).Y = Temp
        Line Input #1, Temp: Decors(i).Largeur = Temp
        Line Input #1, Temp: Decors(i).Hauteur = Temp
        Line Input #1, Temp: Decors(i).EtapeAnimation = Temp
        Line Input #1, Temp: Decors(i).NombreEtapeAnimation = Temp
        Line Input #1, Temp: Decors(i).Nom = Temp
        Line Input #1, Temp: Decors(i).Temporaire = Temp
        Line Input #1, Temp: Decors(i).Duree = Temp
    Next i
    'Sauvegarde des personnages.
    For i = 0 To UBound(Persos())
        Set Persos(i) = New ClsJeuPerso
        Set EffetsPersos(i) = New ClsJeuEffet
        Persos(i).Init i, Parametres
        'Position actuelle du personnage.
        Line Input #1, Temp: Persos(i).PositionX = Temp
        Line Input #1, Temp: Persos(i).PositionY = Temp
        Line Input #1, Temp: Persos(i).CibleX = Temp
        Line Input #1, Temp: Persos(i).CibleY = Temp
        Line Input #1, Temp: Persos(i).LongueurPas = Temp
        Line Input #1, Temp: Persos(i).Acceleration = Temp
        Line Input #1, Temp: Persos(i).BonusDeplacement = Temp
        Line Input #1, Temp: Persos(i).NumeroPas = Temp
        Line Input #1, Temp: Persos(i).Apparence = Temp
        'Position où doit se rendre le personnage.
        Line Input #1, Temp: Persos(i).DirectionX = Temp
        Line Input #1, Temp: Persos(i).DirectionY = Temp
        'Caractéristiques :
        Line Input #1, Temp: Persos(i).Numero = Temp
        Line Input #1, Temp: Persos(i).Nom = Temp
        Line Input #1, Temp: Persos(i).Largeur = Temp
        Line Input #1, Temp: Persos(i).Hauteur = Temp
        Line Input #1, Temp: Persos(i).NumeroFief = Temp
        Line Input #1, Temp: Persos(i).NumeroEquipe = Temp
        Line Input #1, Temp: Persos(i).NumeroChef = Temp
        Line Input #1, Temp: Persos(i).Incommandable = Temp
        Line Input #1, Temp: Persos(i).ChefDecalageX = Temp
        Line Input #1, Temp: Persos(i).ChefDecalageY = Temp
        Line Input #1, Temp: Persos(i).ChefPositionX = Temp
        Line Input #1, Temp: Persos(i).ChefPositionY = Temp
        Line Input #1, Temp: Persos(i).Attitude = Temp
        Line Input #1, Temp: Persos(i).Formation = Temp
        Line Input #1, Temp: Persos(i).Empereur = Temp
        Line Input #1, Temp: Persos(i).NombreSoldats = Temp
        Line Input #1, Temp: Persos(i).Race = Temp
        'Persos(i).Changer_Race_Personnage Persos(i).Race, Parametres, True
        Line Input #1, Temp: Persos(i).Parole = Temp
        Line Input #1, Temp: Persos(i).Definir_Volant = Temp
        Line Input #1, Temp: Persos(i).ChoisirApparence = Temp
        Line Input #1, Temp: Persos(i).FichierApparence = Temp
        Line Input #1, Temp: Persos(i).ApparenceChoisie = Temp
        Line Input #1, Temp: Persos(i).Titre = Temp
        Line Input #1, Temp: Persos(i).Feminin = Temp
        Line Input #1, Temp: Persos(i).GainExperience = Temp
        Line Input #1, Temp: Persos(i).DureeRessurection = Temp
        Line Input #1, Temp: Persos(i).ResurrectionNombre = Temp
        Line Input #1, Temp: Persos(i).VitesseResurrection = Temp
        Line Input #1, Temp: Persos(i).Definir_Commentaires = Temp
        Line Input #1, Temp: Persos(i).CommentairesSpecifique = Temp
        Line Input #1, Temp: Persos(i).CommentairesSpecifiqueIndice = Temp
        Line Input #1, Temp: Persos(i).TempsDernierCommentaire = Temp
        
        Line Input #1, Temp: Persos(i).VitesseRegenerationVie = Temp
        Line Input #1, Temp: Persos(i).VitesseRegenerationEnergie = Temp
        Line Input #1, Temp: Persos(i).VitesseRegenerationMagie = Temp
        Line Input #1, Temp: Persos(i).VitesseRegenerationMoral = Temp
        
        Line Input #1, Temp: Persos(i).Vie = Temp
        Line Input #1, Temp: Persos(i).MaxVie = Temp
        Line Input #1, Temp: Persos(i).RecuperationVie = Temp
        Line Input #1, Temp: Persos(i).ExpVie = Temp
        Line Input #1, Temp: Persos(i).NivVie = Temp
        Line Input #1, Temp: Persos(i).ExpNivVie = Temp
        Line Input #1, Temp: Persos(i).ExpNivPlusVie = Temp
        
        Line Input #1, Temp: Persos(i).Energie = Temp
        Line Input #1, Temp: Persos(i).MaxEnergie = Temp
        Line Input #1, Temp: Persos(i).RecuperationEnergie = Temp
        Line Input #1, Temp: Persos(i).ExpEnergie = Temp
        Line Input #1, Temp: Persos(i).NivEnergie = Temp
        Line Input #1, Temp: Persos(i).ExpNivEnergie = Temp
        Line Input #1, Temp: Persos(i).ExpNivPlusEnergie = Temp
        
        Line Input #1, Temp: Persos(i).Magie = Temp
        Line Input #1, Temp: Persos(i).MaxMagie = Temp
        Line Input #1, Temp: Persos(i).RecuperationMagie = Temp
        Line Input #1, Temp: Persos(i).ExpMagie = Temp
        Line Input #1, Temp: Persos(i).NivMagie = Temp
        Line Input #1, Temp: Persos(i).ExpNivMagie = Temp
        Line Input #1, Temp: Persos(i).ExpNivPlusMagie = Temp
        
        Line Input #1, Temp: Persos(i).Moral = Temp
        Line Input #1, Temp: Persos(i).MaxMoral = Temp
        Line Input #1, Temp: Persos(i).RecuperationMoral = Temp
        Line Input #1, Temp: Persos(i).ExpMoral = Temp
        Line Input #1, Temp: Persos(i).NivMoral = Temp
        Line Input #1, Temp: Persos(i).ExpNivMoral = Temp
        Line Input #1, Temp: Persos(i).ExpNivPlusMoral = Temp
        
        Line Input #1, Temp: Persos(i).Attaque = Temp
        Line Input #1, Temp: Persos(i).MaxAttaque = Temp
        Line Input #1, Temp: Persos(i).ExpAttaque = Temp
        Line Input #1, Temp: Persos(i).NivAttaque = Temp
        Line Input #1, Temp: Persos(i).ExpNivAttaque = Temp
        Line Input #1, Temp: Persos(i).ExpNivPlusAttaque = Temp
        Line Input #1, Temp: Persos(i).PorteeAttaque = Temp
        Line Input #1, Temp: Persos(i).EffetAttaque = Temp
        Line Input #1, Temp: Persos(i).ForcerAttaque = Temp
        Line Input #1, Temp: Persos(i).Kamikaze = Temp
        Line Input #1, Temp: Persos(i).DureeRechargeAttaque = Temp
        Line Input #1, Temp: Persos(i).DureeRechargeAttaqueMax = Temp
        
        Line Input #1, Temp: Persos(i).Defense = Temp
        Line Input #1, Temp: Persos(i).MaxDefense = Temp
        Line Input #1, Temp: Persos(i).ExpDefense = Temp
        Line Input #1, Temp: Persos(i).NivDefense = Temp
        Line Input #1, Temp: Persos(i).ExpNivDefense = Temp
        Line Input #1, Temp: Persos(i).ExpNivPlusDefense = Temp
        Line Input #1, Temp: Persos(i).Armure = Temp
        
        Line Input #1, Temp: Persos(i).Projectile = Temp
        Line Input #1, Temp: Persos(i).ProjectileCible = Temp
        Line Input #1, Temp: Persos(i).Pietinement = Temp
        Line Input #1, Temp: Persos(i).Inattaquable = Temp
        
        Line Input #1, Temp: Persos(i).ChoisirIA = Temp
        Line Input #1, Temp: Persos(i).FichierIA = Temp
        Line Input #1, Temp: Persos(i).IA_Equiper = Temp
        Line Input #1, Temp: Persos(i).IA_Manger = Temp
        Line Input #1, Temp: Persos(i).IA_Pacifique = Temp
        Line Input #1, Temp: Persos(i).IA_Temeraire = Temp
        Line Input #1, Temp: Persos(i).IA_Berserk = Temp
        Line Input #1, Temp: Persos(i).IA_Ordre = Temp
        Line Input #1, Temp: Persos(i).IA_Cible = Temp
        Line Input #1, Temp: Persos(i).ControleJoueur = Temp
        Line Input #1, Temp: Persos(i).Definir_Argent = Temp
        Line Input #1, Temp: Persos(i).Action = Temp
        Line Input #1, Temp: Persos(i).Direction = Temp
        Line Input #1, Temp: Persos(i).EnDeplacement = Temp
        Line Input #1, Temp: Persos(i).DureeImmobilisation = Temp
        Line Input #1, Temp: Persos(i).DansUneMaison = Temp
        Line Input #1, Temp: Persos(i).DansSaMaison = Temp
        Line Input #1, Temp: Persos(i).IndiceMaison = Temp
        Line Input #1, Temp: Persos(i).DansUnChateau = Temp
        Line Input #1, Temp: Persos(i).DansSonChateau = Temp
        
        'Mémorise les indices des destinations vers lesquels le joueur se dirige.
        Line Input #1, Temp: Persos(i).IndiceRessource = Temp
        Line Input #1, Temp: Persos(i).IndiceChateau = Temp
        Line Input #1, Temp: Persos(i).IndicePerso = Temp
        Line Input #1, Temp: Persos(i).IndiceTresor = Temp
        
        'Variables de ressources.
        Line Input #1, Temp: Persos(i).RessourceCourante = Temp
        Line Input #1, Temp: Persos(i).ServiceCourant = Temp
        
        'Compétences spéciales.
        For j = 0 To Parametres.NombreCompetencesSpeciales - 1
            Line Input #1, Temp: Persos(i).Definir_Carac_CompetenceSpeciales(j) = Temp
            Line Input #1, Temp: Persos(i).Definir_ExpSpeciales(j) = Temp
            Line Input #1, Temp: Persos(i).Niveau_CompetenceSpeciales(j) = Temp
            Line Input #1, Temp: Persos(i).Definir_ExpNivSpeciales(j) = Temp
            Line Input #1, Temp: Persos(i).Definir_ExpNivPlusSpeciales(j) = Temp
        Next j
        'Compétences des ressources.
        For j = 0 To Parametres.NombreRessources - 1
            Line Input #1, Temp: Persos(i).Definir_Carac_Ressources(j) = Temp
            Line Input #1, Temp: Persos(i).Definir_ExpRessources(j) = Temp
            Line Input #1, Temp: Persos(i).Niveau_CompetenceRessources(j) = Temp
            Line Input #1, Temp: Persos(i).Definir_ExpNivRessources(j) = Temp
            Line Input #1, Temp: Persos(i).Definir_ExpNivPlusRessources(j) = Temp
        Next j
        'Compétences des Services.
        For j = 0 To Parametres.NombreServices - 2
            Line Input #1, Temp: Persos(i).Definir_Carac_CompetenceServices(j) = Temp
            Line Input #1, Temp: Persos(i).Definir_ExpServices(j) = Temp
            Line Input #1, Temp: Persos(i).Niveau_CompetenceServices(j) = Temp
            Line Input #1, Temp: Persos(i).Definir_ExpNivServices(j) = Temp
            Line Input #1, Temp: Persos(i).Definir_ExpNivPlusServices(j) = Temp
        Next j
        'Compétences des objets.
        For j = 0 To Parametres.NombreCompetencesObjets - 1
            Line Input #1, Temp: Persos(i).Definir_Carac_CompetenceObjets(j) = Temp
            Line Input #1, Temp: Persos(i).Definir_ExpObjets(j) = Temp
            Line Input #1, Temp: Persos(i).Niveau_CompetenceObjets(j) = Temp
            Line Input #1, Temp: Persos(i).Definir_ExpNivObjets(j) = Temp
            Line Input #1, Temp: Persos(i).Definir_ExpNivPlusObjets(j) = Temp
        Next j
        
        'Modificateurs de la race
        Line Input #1, Temp: Persos(i).CoefRaceVie = Temp
        Line Input #1, Temp: Persos(i).CoefRaceEnergie = Temp
        Line Input #1, Temp: Persos(i).CoefRaceMagie = Temp
        Line Input #1, Temp: Persos(i).CoefRaceMoral = Temp
        Line Input #1, Temp: Persos(i).CoefRaceAttaque = Temp
        Line Input #1, Temp: Persos(i).CoefRaceDefense = Temp
        For j = 0 To Parametres.NombreCompetencesSpeciales - 1
            Line Input #1, Temp: Persos(i).Definir_CoefRaceCompSpeciales(j) = Temp 'Compétences spéciales.
        Next j
        For j = 0 To Parametres.NombreRessources - 1
            Line Input #1, Temp: Persos(i).Definir_CoefRaceCompRessources(j) = Temp 'Compétences des ressources.
        Next j
        For j = 0 To Parametres.NombreServices - 2
            Line Input #1, Temp: Persos(i).Definir_CoefRaceCompServices(j) = Temp 'Compétences des Services.
        Next j
        For j = 0 To Parametres.NombreCompetencesObjets - 1
            Line Input #1, Temp: Persos(i).Definir_CoefRaceCompObjets(j) = Temp 'Compétences des objets.
        Next j
        Line Input #1, Temp: Persos(i).BonusMaxRessourcesRace = Temp
        
        Persos(i).Gestion_Tous_Niveaux
        Persos(i).Ajuster_Bonus
        
        'Objets.
        Line Input #1, Temp: Persos(i).IndiceObjet = Temp
        Line Input #1, Temp: Persos(i).ObjetSelectionne = Temp
        Line Input #1, Temp: Persos(i).EtatFabricationObjet = Temp
        For j = 0 To Persos(i).Nombre_Objets_Equipes - 1
            Line Input #1, Temp: Persos(i).Objet_Equipes_Actif(j) = Temp
            Line Input #1, Temp: Persos(i).Objet_Equipes_Type(j) = Temp
            If Persos(i).Objet_Equipes_Actif(j) Then
                Persos(i).Objet_Equipe_Charger j, Persos(i).Objet_Equipes_Type(j), Parametres
            End If
        Next j
        For j = 0 To Persos(i).Nombre_Objets_Inventaire - 1
            Line Input #1, Temp: Persos(i).Objet_Inventaire_Actif(j) = Temp
            Line Input #1, Temp: Persos(i).Objet_Inventaire_Type(j) = Temp
            If Persos(i).Objet_Inventaire_Actif(j) Then
                Persos(i).Objet_Inventaire_Charger j, Persos(i).Objet_Inventaire_Type(j), Parametres
            End If
        Next j
        Persos(i).Actualiser_Bonus
'        If Persos(i).CoefRaceVie < 1 Then
'            Persos(i).Vie = Persos(i).Vie / Persos(i).CoefRaceVie
'        Else
'            Persos(i).Vie = Persos(i).Vie * Persos(i).CoefRaceVie
'        End If
'        If Persos(i).CoefRaceEnergie < 1 Then
'            Persos(i).Energie = Persos(i).Energie / Persos(i).CoefRaceEnergie
'        Else
'            Persos(i).Energie = Persos(i).Energie * Persos(i).CoefRaceEnergie
'        End If
'        If Persos(i).CoefRaceMagie < 1 Then
'            Persos(i).Magie = Persos(i).Magie / Persos(i).CoefRaceMagie
'        Else
'            Persos(i).Magie = Persos(i).Magie * Persos(i).CoefRaceMagie
'        End If
'        If Persos(i).CoefRaceMoral < 1 Then
'            Persos(i).Moral = Persos(i).Moral / Persos(i).CoefRaceMoral
'        Else
'            Persos(i).Moral = Persos(i).Moral * Persos(i).CoefRaceMoral
'        End If
'        If Persos(i).CoefRaceAttaque < 1 Then
'            Persos(i).Attaque = Persos(i).Attaque / Persos(i).CoefRaceAttaque
'        Else
'            Persos(i).Attaque = Persos(i).Attaque * Persos(i).CoefRaceAttaque
'        End If
'        If Persos(i).CoefRaceDefense < 1 Then
'            Persos(i).Defense = Persos(i).Defense / Persos(i).CoefRaceDefense
'        Else
'            Persos(i).Defense = Persos(i).Defense * Persos(i).CoefRaceDefense
'        End If
'        If Persos(i).Vie > Persos(i).Definir_MaxVie Then Persos(i).Vie = Persos(i).Definir_MaxVie
'        If Persos(i).Energie > Persos(i).Definir_MaxEnergie Then Persos(i).Energie = Persos(i).Definir_MaxEnergie
'        If Persos(i).Magie > Persos(i).Definir_MaxVie Then Persos(i).Magie = Persos(i).Definir_MaxMagie
'        If Persos(i).Moral > Persos(i).Definir_MaxMoral Then Persos(i).Moral = Persos(i).Definir_MaxMoral
'        If Persos(i).Attaque > Persos(i).Definir_MaxAttaque Then Persos(i).Attaque = Persos(i).Definir_MaxAttaque
'        If Persos(i).Defense > Persos(i).Definir_MaxDefense Then Persos(i).Defense = Persos(i).Definir_MaxDefense
        
        'Stock de ressources.
        For j = 0 To Parametres.NombreRessources - 1
            Line Input #1, Temp: Persos(i).Definir_ressources(j) = Temp
        Next j
        
        'Bestiaire.
        For j = 0 To Parametres.NombreRaces - 1
            Line Input #1, Temp: Persos(i).Definir_Bestiaire(j) = Temp
        Next j
        
        'Effets temporaires.
        Line Input #1, Temp
        k = Val(Temp)
        If k > 0 Then
            For j = 1 To k
                Line Input #1, Temp: Line Input #1, Temp2
                Line Input #1, Temp3: Line Input #1, Temp4
                If Not Parametres.Objet_Unique(Temp) Then
                    Persos(i).EffetTemp_Ajouter Temp, Temp2, CLng(Temp3), CLng(Temp4)
                End If
            Next j
        End If
        
        'Effets qui se rechargent.
        Line Input #1, Temp
        k = Val(Temp)
        If k > 0 Then
            For j = 1 To k
                Line Input #1, Temp: Line Input #1, Temp2
                Persos(i).EffetRecharge_Ajouter Temp, Temp2
            Next j
        End If
        
        'Coefficients de caractéristiques.
        Line Input #1, Temp: Persos(i).Vie = CDbl(Temp) * Persos(i).Definir_MaxVie
        Line Input #1, Temp: Persos(i).Energie = CDbl(Temp) * Persos(i).Definir_MaxEnergie
        Line Input #1, Temp: Persos(i).Magie = CDbl(Temp) * Persos(i).Definir_MaxMagie
        Line Input #1, Temp: Persos(i).Moral = CDbl(Temp) * Persos(i).Definir_MaxMoral
        Line Input #1, Temp: Persos(i).Attaque = CDbl(Temp) * Persos(i).Definir_MaxAttaque
        Line Input #1, Temp: Persos(i).Defense = CDbl(Temp) * Persos(i).Definir_MaxDefense
    Next i
    For i = 0 To UBound(Maisons())
        Set Maisons(i) = New ClsJeuBatiment
        Set EffetsMaisons(i) = New ClsJeuEffet
        Line Input #1, Temp: Maisons(i).Numero = Temp
        Line Input #1, Temp: Maisons(i).TypeBatiment = Temp
        Line Input #1, Temp: Maisons(i).PositionX = Temp
        Line Input #1, Temp: Maisons(i).PositionY = Temp
        Line Input #1, Temp: Maisons(i).Largeur = Temp
        Line Input #1, Temp: Maisons(i).Hauteur = Temp
        Line Input #1, Temp: Maisons(i).EntreePositionX = Temp
        Line Input #1, Temp: Maisons(i).EntreePositionY = Temp
        Line Input #1, Temp: Maisons(i).Construit = Temp
        Line Input #1, Temp: Maisons(i).EnConstruction = Temp
        Line Input #1, Temp: Maisons(i).Vie = Temp
        Line Input #1, Temp: Maisons(i).RecuperationVie = Temp
        Line Input #1, Temp: Maisons(i).MaxVie = Temp
        Line Input #1, Temp: Maisons(i).Argent = Temp
        Line Input #1, Temp: Maisons(i).MaxStock = Temp
        
        Maisons(i).Definir_Nombre_Stock_Objet = Parametres.Batiment_NombreObjetsFabriques(Maisons(i).TypeBatiment)
        For j = 0 To Parametres.Batiment_NombreObjetsFabriques(Maisons(i).TypeBatiment) - 1
            Line Input #1, Temp: Maisons(i).Stock_Objets(j) = Temp
        Next j
        Line Input #1, Temp: Maisons(i).Definir_Stock = Val(Temp)
        
        For j = 0 To Maisons(i).Definir_Nombre_Objet_Inventaire - 1
            Line Input #1, Temp: Maisons(i).Definir_ObjetsInventaire(j) = Val(Temp)
        Next j
        
        Line Input #1, Temp: Maisons(i).Magasin = Temp
        Line Input #1, Temp: Maisons(i).Fabrique = Temp
        Line Input #1, Temp: Maisons(i).Service = Temp
        Line Input #1, Temp: Maisons(i).Marche = Temp
    Next i
    Close #1
    
'    For i = 0 To UBound(Persos())
'        Persos(i).Calculer_Apparence Maisons(i), Parametres
'    Next i
    Exit Sub
Erreur:
    MsgBox "Le fichier de monde " & Nom & " est invalide", vbCritical
    End
ErreurParam:
    MsgBox "Le fichier de monde " & Nom & " est invalide" & Chr(10) & Chr(13) & "La version ou les mods dans cette sauvegarde" & Chr(10) & Chr(13) & "sont différents de ceux actuels", vbCritical
    End
End Sub

Public Sub Enregistrer_Sauvegarde_Monde(Optional ByVal Nom As String)
    If Not ComReseau.Client Then
    Dim i As Long
    Dim j As Integer
    If Nom = "" Then
        Nom = Format(Date, "dd_mm_yyyy") & "__" & Format(Time, "hh_mm")
    End If
    
    If SavImageEnQuittant Then Capture_Ecran Nom
    Capture_Texte Nom
    
    Nom = Nom & ExtensionFichiersMonde
    
    FileCopy App.Path & FichierVierge, App.Path & CheminSavMonde & Nom
    
    Open App.Path & CheminSavMonde & Nom For Output As #1
    
    'Sauvegarde les paramètres pour vérification.
    Print #1, Parametres.NombreCompetencesObjets
    Print #1, Parametres.NombreCompetencesSpeciales
    Print #1, Parametres.NombreDecors
    Print #1, Parametres.NombreEffets
    Print #1, Parametres.NombreEpoques
    Print #1, Parametres.NombreObjets
    Print #1, Parametres.NombrePeuples
    Print #1, Parametres.NombreRaces
    Print #1, Parametres.NombreRacesJouables
    Print #1, Parametres.NombreResistances
    Print #1, Parametres.NombreRessources
    Print #1, Parametres.NombreRessourcesDepart
    Print #1, Parametres.NombreServices
    Print #1, Parametres.NombreStatuts
    Print #1, Parametres.NombreTerrainRessource
    Print #1, Parametres.NombreTresors
    Print #1, Parametres.NombreTypeBatiments
    Print #1, Parametres.NombreTypeBatimentsJouables
    Print #1, Parametres.NombreVitesseEpoques
    
    Print #1, UBound(Fiefs()) + 1
    Print #1, UBound(Chateaux()) + 1
    Print #1, UBound(Ressources()) + 1
    Print #1, Tresors.Count
    Print #1, Decors.Count
    Print #1, UBound(Persos()) + 1
    Print #1, UBound(Maisons()) + 1
    
    Print #1, Parametres.VitesseEpoqueSelectionnee
    'Enregistre le chemin de la mini-map.
    If Cartes.Chargee And Cartes.Definir_ImageCarte_Existe Then
        Print #1, Cartes.DerniereCarte
    Else
        Print #1, ""
    End If
    Print #1, Monde.Largeur
    Print #1, Monde.Hauteur
    For i = 0 To (Monde.Largeur + 1) * (Monde.Hauteur + 1) - 1
        Print #1, Monde.Param_TerrainApparence(i)
        Print #1, Monde.Param_TerrainCategorie(i)
    Next i
    
    'Paramètres du jeu.
    Print #1, AffTemps.Heures
    Print #1, AffTemps.Minutes
    Print #1, Jeu.Chrono
    Print #1, Jeu.IntituleChrono
    
    For i = 0 To UBound(Fiefs())
        Print #1, Fiefs(i).NumeroEquipe
        Print #1, Fiefs(i).Definir_TypePeuple
        Print #1, Fiefs(i).Nom
        Print #1, Fiefs(i).Epoque
        Print #1, Fiefs(i).ChangerEpoque
        Print #1, Fiefs(i).TempsChangementEpoque
        Print #1, Fiefs(i).NombreCitoyens
        Print #1, Fiefs(i).NombreMorts
        Print #1, Fiefs(i).Empereur
        Print #1, Fiefs(i).Argent
        For j = 0 To Parametres.NombreRessources - 1
            Print #1, Fiefs(i).Definir_ressources(j)
        Next j
    Next i
    For i = 0 To UBound(Chateaux())
        Print #1, Chateaux(i).Numero
        Print #1, Chateaux(i).PositionX
        Print #1, Chateaux(i).PositionY
        Print #1, Chateaux(i).Largeur
        Print #1, Chateaux(i).Hauteur
        Print #1, Chateaux(i).EntreePositionX
        Print #1, Chateaux(i).EntreePositionY
        Print #1, Chateaux(i).Construit
        Print #1, Chateaux(i).EnConstruction
        Print #1, Chateaux(i).Vie
        Print #1, Chateaux(i).RecuperationVie
        Print #1, Chateaux(i).MaxVie
    Next i
    For i = 0 To UBound(Ressources())
        Print #1, Ressources(i).PositionX
        Print #1, Ressources(i).PositionY
        Print #1, Ressources(i).Nom
        Print #1, Ressources(i).Apparence
        Print #1, Ressources(i).TypeRessource
        Print #1, Ressources(i).Quantite
        Print #1, Ressources(i).MaxQuantite
    Next i
    For i = 1 To Tresors.Count
        Print #1, Tresors(i).X
        Print #1, Tresors(i).Y
        Print #1, Tresors(i).Definir_Fortune
        Print #1, Tresors(i).Definir_Argent
        For j = 0 To Parametres.NombreRessources - 1
            Print #1, Tresors(i).Definir_ressources(j)
        Next j
        Print #1, Tresors(i).Nombre_Objets
        For j = 1 To Tresors(i).Nombre_Objets
            Print #1, Tresors(i).Definir_Objet(j)
        Next j
    Next i
    For i = 1 To Decors.Count
        Print #1, Decors(i).Categorie
        Print #1, Decors(i).X
        Print #1, Decors(i).Y
        Print #1, Decors(i).Largeur
        Print #1, Decors(i).Hauteur
        Print #1, Decors(i).EtapeAnimation
        Print #1, Decors(i).NombreEtapeAnimation
        Print #1, Decors(i).Nom
        Print #1, -CDbl(Decors(i).Temporaire)
        Print #1, Decors(i).Duree
    Next i
    For i = 0 To UBound(Persos())
        'Position actuelle du personnage.
        Print #1, Persos(i).PositionX
        Print #1, Persos(i).PositionY
        Print #1, Persos(i).CibleX
        Print #1, Persos(i).CibleY
        Print #1, Persos(i).LongueurPas
        Print #1, Persos(i).Acceleration
        Print #1, Persos(i).BonusDeplacement
        Print #1, Persos(i).NumeroPas
        Print #1, Persos(i).Apparence
        'Position où doit se rendre le personnage.
        Print #1, Persos(i).DirectionX
        Print #1, Persos(i).DirectionY
        'Caractéristiques :
        Print #1, Persos(i).Numero
        Print #1, Persos(i).Nom
        Print #1, Persos(i).Largeur
        Print #1, Persos(i).Hauteur
        Print #1, Persos(i).NumeroFief
        Print #1, Persos(i).NumeroEquipe
        Print #1, Persos(i).NumeroChef
        Print #1, Persos(i).Incommandable
        Print #1, Persos(i).ChefDecalageX
        Print #1, Persos(i).ChefDecalageY
        Print #1, Persos(i).ChefPositionX
        Print #1, Persos(i).ChefPositionY
        Print #1, Persos(i).Attitude
        Print #1, Persos(i).Formation
        Print #1, Persos(i).Empereur
        Print #1, Persos(i).NombreSoldats
        Print #1, Persos(i).Race
        Print #1, Persos(i).Parole
        Print #1, Persos(i).Volant
        Print #1, Persos(i).ChoisirApparence
        Print #1, Persos(i).FichierApparence
        Print #1, Persos(i).ApparenceChoisie
        Print #1, Persos(i).Titre
        Print #1, Persos(i).Feminin
        Print #1, Persos(i).GainExperience
        Print #1, Persos(i).DureeRessurection
        Print #1, Persos(i).ResurrectionNombre
        Print #1, Persos(i).VitesseResurrection
        Print #1, Persos(i).Definir_Commentaires
        Print #1, Persos(i).CommentairesSpecifique
        Print #1, Persos(i).CommentairesSpecifiqueIndice
        Print #1, Persos(i).TempsDernierCommentaire
        
        Print #1, Persos(i).VitesseRegenerationVie
        Print #1, Persos(i).VitesseRegenerationEnergie
        Print #1, Persos(i).VitesseRegenerationMagie
        Print #1, Persos(i).VitesseRegenerationMoral
        
        Print #1, Persos(i).Vie
        Print #1, Persos(i).MaxVie
        Print #1, Persos(i).RecuperationVie
        Print #1, Persos(i).ExpVie
        Print #1, Persos(i).NivVie
        Print #1, Persos(i).ExpNivVie
        Print #1, Persos(i).ExpNivPlusVie
        
        Print #1, Persos(i).Energie
        Print #1, Persos(i).MaxEnergie
        Print #1, Persos(i).RecuperationEnergie
        Print #1, Persos(i).ExpEnergie
        Print #1, Persos(i).NivEnergie
        Print #1, Persos(i).ExpNivEnergie
        Print #1, Persos(i).ExpNivPlusEnergie
        
        Print #1, Persos(i).Magie
        Print #1, Persos(i).MaxMagie
        Print #1, Persos(i).RecuperationMagie
        Print #1, Persos(i).ExpMagie
        Print #1, Persos(i).NivMagie
        Print #1, Persos(i).ExpNivMagie
        Print #1, Persos(i).ExpNivPlusMagie
        
        Print #1, Persos(i).Moral
        Print #1, Persos(i).MaxMoral
        Print #1, Persos(i).RecuperationMoral
        Print #1, Persos(i).ExpMoral
        Print #1, Persos(i).NivMoral
        Print #1, Persos(i).ExpNivMoral
        Print #1, Persos(i).ExpNivPlusMoral
        
        Print #1, Persos(i).Attaque
        Print #1, Persos(i).MaxAttaque
        Print #1, Persos(i).ExpAttaque
        Print #1, Persos(i).NivAttaque
        Print #1, Persos(i).ExpNivAttaque
        Print #1, Persos(i).ExpNivPlusAttaque
        Print #1, Persos(i).PorteeAttaque
        Print #1, Persos(i).EffetAttaque
        Print #1, Persos(i).ForcerAttaque
        Print #1, Persos(i).Kamikaze
        Print #1, Persos(i).DureeRechargeAttaque
        Print #1, Persos(i).DureeRechargeAttaqueMax
        
        Print #1, Persos(i).Defense
        Print #1, Persos(i).MaxDefense
        Print #1, Persos(i).ExpDefense
        Print #1, Persos(i).NivDefense
        Print #1, Persos(i).ExpNivDefense
        Print #1, Persos(i).ExpNivPlusDefense
        Print #1, Persos(i).Armure
        
        Print #1, Persos(i).Projectile
        Print #1, Persos(i).ProjectileCible
        Print #1, Persos(i).Pietinement
        Print #1, Persos(i).Inattaquable
        
        Print #1, Persos(i).ChoisirIA
        Print #1, Persos(i).FichierIA
        Print #1, Persos(i).IA_Equiper
        Print #1, Persos(i).IA_Manger
        Print #1, Persos(i).IA_Pacifique
        Print #1, Persos(i).IA_Temeraire
        Print #1, Persos(i).IA_Berserk
        Print #1, Persos(i).IA_Ordre
        Print #1, Persos(i).IA_Cible
        Print #1, Persos(i).ControleJoueur
        Print #1, Persos(i).Definir_Argent
        Print #1, Persos(i).Action
        Print #1, Persos(i).Direction
        Print #1, Persos(i).EnDeplacement
        Print #1, Persos(i).DureeImmobilisation
        Print #1, Persos(i).DansUneMaison
        Print #1, Persos(i).DansSaMaison
        Print #1, Persos(i).IndiceMaison
        Print #1, Persos(i).DansUnChateau
        Print #1, Persos(i).DansSonChateau
        
        'Mémorise les indices des destinations vers lesquels le joueur se dirige.
        Print #1, Persos(i).IndiceRessource
        Print #1, Persos(i).IndiceChateau
        Print #1, Persos(i).IndicePerso
        Print #1, Persos(i).IndiceTresor
        
        'Variables de ressources.
        Print #1, Persos(i).RessourceCourante
        Print #1, Persos(i).ServiceCourant
        
        'Compétences spéciales.
        For j = 0 To Parametres.NombreCompetencesSpeciales - 1
            Print #1, Persos(i).Definir_Carac_CompetenceSpeciales(j)
            Print #1, Persos(i).Definir_ExpSpeciales(j)
            Print #1, Persos(i).Niveau_CompetenceSpeciales(j)
            Print #1, Persos(i).Definir_ExpNivSpeciales(j)
            Print #1, Persos(i).Definir_ExpNivPlusSpeciales(j)
        Next j
        'Compétences des ressources.
        For j = 0 To Parametres.NombreRessources - 1
            Print #1, Persos(i).Definir_Carac_Ressources(j)
            Print #1, Persos(i).Definir_ExpRessources(j)
            Print #1, Persos(i).Niveau_CompetenceRessources(j)
            Print #1, Persos(i).Definir_ExpNivRessources(j)
            Print #1, Persos(i).Definir_ExpNivPlusRessources(j)
        Next j
        'Compétences des Services.
        For j = 0 To Parametres.NombreServices - 2
            Print #1, Persos(i).Definir_Carac_CompetenceServices(j)
            Print #1, Persos(i).Definir_ExpServices(j)
            Print #1, Persos(i).Niveau_CompetenceServices(j)
            Print #1, Persos(i).Definir_ExpNivServices(j)
            Print #1, Persos(i).Definir_ExpNivPlusServices(j)
        Next j
        'Compétences des objets.
        For j = 0 To Parametres.NombreCompetencesObjets - 1
            Print #1, Persos(i).Definir_Carac_CompetenceObjets(j)
            Print #1, Persos(i).Definir_ExpObjets(j)
            Print #1, Persos(i).Niveau_CompetenceObjets(j)
            Print #1, Persos(i).Definir_ExpNivObjets(j)
            Print #1, Persos(i).Definir_ExpNivPlusObjets(j)
        Next j
        
        'Modificateurs de la race
        Print #1, Persos(i).CoefRaceVie
        Print #1, Persos(i).CoefRaceEnergie
        Print #1, Persos(i).CoefRaceMagie
        Print #1, Persos(i).CoefRaceMoral
        Print #1, Persos(i).CoefRaceAttaque
        Print #1, Persos(i).CoefRaceDefense
        For j = 0 To Parametres.NombreCompetencesSpeciales - 1
            Print #1, Persos(i).Definir_CoefRaceCompSpeciales(j)  'Compétences spéciales.
        Next j
        For j = 0 To Parametres.NombreRessources - 1
            Print #1, Persos(i).Definir_CoefRaceCompRessources(j)  'Compétences des ressources.
        Next j
        For j = 0 To Parametres.NombreServices - 2
            Print #1, Persos(i).Definir_CoefRaceCompServices(j)  'Compétences des Services.
        Next j
        For j = 0 To Parametres.NombreCompetencesObjets - 1
            Print #1, Persos(i).Definir_CoefRaceCompObjets(j)  'Compétences des objets.
        Next j
        Print #1, Persos(i).BonusMaxRessourcesRace
        
        'Objets.
        Print #1, Persos(i).IndiceObjet
        Print #1, Persos(i).ObjetSelectionne
        Print #1, Persos(i).EtatFabricationObjet
        For j = 0 To Persos(i).Nombre_Objets_Equipes - 1
            If Persos(i).Objet_Equipes_Actif(j) And _
               Parametres.Objet_Unique(Persos(i).Objet_Equipes_Type(j)) Then 'On ne sauvegarde pas les objets uniques.
                Print #1, False
                Print #1, -1
            Else
                Print #1, Persos(i).Objet_Equipes_Actif(j)
                Print #1, Persos(i).Objet_Equipes_Type(j)
            End If
        Next j
        For j = 0 To Persos(i).Nombre_Objets_Inventaire - 1
            If Persos(i).Objet_Inventaire_Actif(j) And _
               Parametres.Objet_Unique(Persos(i).Objet_Inventaire_Type(j)) Then 'On ne sauvegarde pas les objets uniques.
                Print #1, False
                Print #1, -1
            Else
                Print #1, Persos(i).Objet_Inventaire_Actif(j)
                Print #1, Persos(i).Objet_Inventaire_Type(j)
            End If
        Next j
        
        'Stock de ressources.
        For j = 0 To Parametres.NombreRessources - 1
            Print #1, Persos(i).Definir_ressources(j)
        Next j
        
        'Bestiaire.
        For j = 0 To Parametres.NombreRaces - 1
            Print #1, Persos(i).Definir_Bestiaire(j)
        Next j
        
        'Effets temporaires.
        Print #1, Persos(i).Nombre_EffetsTemp
        If Persos(i).Nombre_EffetsTemp > 0 Then
            For j = 1 To Persos(i).Nombre_EffetsTemp
                Print #1, Persos(i).Definir_EffetTemp_Numero(j)
                Print #1, Persos(i).Definir_EffetTemp_Duree(j)
                Print #1, Persos(i).Definir_EffetTemp_DureeMax(j)
                Print #1, Persos(i).Definir_EffetTemp_Proprietaire(j)
            Next j
        End If
        
        'Effets qui se rechargent.
        Print #1, Persos(i).Nombre_EffetsRecharge
        If Persos(i).Nombre_EffetsRecharge > 0 Then
            For j = 1 To Persos(i).Nombre_EffetsRecharge
                Print #1, Persos(i).Definir_EffetsRecharge_Numero(j)
                Print #1, Persos(i).Definir_EffetsRecharge_Duree(j)
            Next j
        End If
        
        'Coefficients de caractéristiques.
        Print #1, Persos(i).Definir_Vie / Persos(i).Definir_MaxVie
        Print #1, Persos(i).Definir_Energie / Persos(i).Definir_MaxEnergie
        Print #1, Persos(i).Definir_Magie / Persos(i).Definir_MaxMagie
        Print #1, Persos(i).Definir_Moral / Persos(i).Definir_MaxMoral
        Print #1, Persos(i).Definir_Attaque / Persos(i).Definir_MaxAttaque
        Print #1, Persos(i).Definir_Defense / Persos(i).Definir_MaxDefense
    Next i
    For i = 0 To UBound(Maisons())
        Print #1, Maisons(i).Numero
        Print #1, Maisons(i).TypeBatiment
        Print #1, Maisons(i).PositionX
        Print #1, Maisons(i).PositionY
        Print #1, Maisons(i).Largeur
        Print #1, Maisons(i).Hauteur
        Print #1, Maisons(i).EntreePositionX
        Print #1, Maisons(i).EntreePositionY
        Print #1, Maisons(i).Construit
        Print #1, Maisons(i).EnConstruction
        Print #1, Maisons(i).Vie
        Print #1, Maisons(i).RecuperationVie
        Print #1, Maisons(i).MaxVie
        Print #1, Maisons(i).Argent
        Print #1, Maisons(i).MaxStock
        
        For j = 0 To Parametres.Batiment_NombreObjetsFabriques(Maisons(i).TypeBatiment) - 1
            Print #1, Maisons(i).Stock_Objets(j)
        Next j
        
        Print #1, Maisons(i).Definir_Stock
        
        For j = 0 To Maisons(i).Definir_Nombre_Objet_Inventaire - 1
            Print #1, Maisons(i).Definir_ObjetsInventaire(j)
        Next j
        
        Print #1, Maisons(i).Magasin
        Print #1, Maisons(i).Fabrique
        Print #1, Maisons(i).Service
        Print #1, Maisons(i).Marche
    Next i
    Close #1
    Messages(0).Ajouter_Message_SauvegardeMonde
    End If
End Sub

Private Sub Capture_Ecran(ByVal Fichier As String)
    On Error GoTo Erreur
    'Clipboard.Clear
    'Clipboard.SetData screen
    GrabScreen
    SavePicture Clipboard.GetData, App.Path & CheminSavMonde & Fichier & ExtensionFichiersAperçu
    Exit Sub
Erreur:
End Sub
Private Sub Capture_Texte(ByVal Fichier As String)
    FileCopy App.Path & FichierVierge, App.Path & CheminSavMonde & Fichier & ExtensionFichiersTexte
    
    Open App.Path & CheminSavMonde & Fichier & ExtensionFichiersTexte For Output As #1
    Print #1, InfoDate & Format(Date, "dd/mm/yyyy") & InfoHeure & Format(Time, "hh:mm")
    Print #1, ""
    Print #1, InfoDimensions & Monde.Largeur & " * " & Monde.Hauteur
    Print #1, ""
    Print #1, InfoJoueur & Persos(NoPerso).Nom
    Print #1, InfoJoueursMax & Fiefs(Persos(NoPerso).NumeroFief).NombreCitoyens
    Print #1, InfoPeuple & Persos(NoPerso).NumeroFief + 1
    Print #1, InfoEpoque & Parametres.Epoque_Nom(Fiefs(Persos(NoPerso).NumeroFief).Epoque)
    Print #1, InfoFiefs & UBound(Fiefs()) + 1
    Print #1, InfoPersonnages & UBound(Persos()) + 1
    Close #1

End Sub

Public Sub Monde_Changer_Nom(ByVal Nom As String, ByVal Scenario As Boolean)
    On Error Resume Next
    Dim Fichier1 As String
    Dim Fichier2 As String
    Fichier1 = App.Path & IIf(Scenario, CheminScenario, CheminSavMonde) & Nom
    Fichier2 = InputBox("Entrez le nouveau nom du monde : " & Chr(13) & Nom, "Changer le nom du monde", Nom)
    If Fichier2 <> "" Then
        Fichier2 = App.Path & IIf(Scenario, CheminScenario, CheminSavMonde) & Fichier2
        FileCopy Fichier1 & ExtensionFichiersMonde, Fichier2 & ExtensionFichiersMonde
        FileCopy Fichier1 & ExtensionFichiersScenario, Fichier2 & ExtensionFichiersScenario
        FileCopy Fichier1 & ExtensionFichiersTexte, Fichier2 & ExtensionFichiersTexte
        FileCopy Fichier1 & ExtensionFichiersAperçu, Fichier2 & ExtensionFichiersAperçu
        FileCopy Fichier1 & ExtensionFichiersScript, Fichier2 & ExtensionFichiersScript
        Monde_Supprimer Nom, Scenario
    End If
End Sub

Public Sub Monde_Supprimer(ByVal Nom As String, ByVal Scenario As Boolean)
    On Error Resume Next
    Dim Fichier As String
    Fichier = App.Path & IIf(Scenario, CheminScenario, CheminSavMonde) & Nom
    Kill Fichier & ExtensionFichiersScenario
    Kill Fichier & ExtensionFichiersMonde
    Kill Fichier & ExtensionFichiersTexte
    Kill Fichier & ExtensionFichiersAperçu
    Kill Fichier & ExtensionFichiersScript
End Sub

Public Sub Monde_Changer_Commentaires(ByVal Nom As String, ByVal Scenario As Boolean)
    Editer_Texte App.Path & IIf(Scenario, CheminScenario, CheminSavMonde) & Nom & ExtensionFichiersTexte
End Sub

'Propriétés
Public Property Get Definir_SavMondeEnQuittant() As Boolean
    Definir_SavMondeEnQuittant = SavMondeEnQuittant
End Property
Public Property Let Definir_SavMondeEnQuittant(ByVal Valeur As Boolean)
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    SavMondeEnQuittant = Valeur
    FicIni.Parametre("SavMondeEnQuittant") = -Int(SavMondeEnQuittant)
End Property

Public Property Get Definir_SavImageEnQuittant() As Boolean
    Definir_SavImageEnQuittant = SavImageEnQuittant
End Property
Public Property Let Definir_SavImageEnQuittant(ByVal Valeur As Boolean)
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    SavImageEnQuittant = Valeur
    FicIni.Parametre("SavImageEnQuittant") = -Int(SavImageEnQuittant)
End Property
