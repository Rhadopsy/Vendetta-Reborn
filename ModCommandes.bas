Attribute VB_Name = "ModCommandes"
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
Const Debugger = False

'Module contenant tous les ordres envoyés.
Const ConstPersoDeplacer = 0
Const ConstPersoDeplacerSansCommentaire = 81
Const ConstPersoChercherRessource = 1
Const ConstPersoRentrerChateau = 2
Const ConstPersoRentrerMaison = 3
Const ConstPersoCreerMaisonCoordonnees = 4
Const ConstPersoCreerMaisonType = 5
Const ConstPersoDefinirIA = 6
Const ConstPersoAttaquerTerrainCoordonnees = 7
Const ConstPersoAttaquerTerrainCible = 8
Const ConstPersoVendreObjetsSelectionnees = 9
Const ConstPersoObjetDeposerSol = 11
Const ConstPersoRessourcesDeposerSol = 12
Const ConstPersoDefinirObjetSelectionne = 13
Const ConstPersoDefinirChefPosition = 14
Const ConstPersoDefinirIAOrdre = 15
Const ConstPersoDefinirIAAttitude = 78
Const ConstPersoDefinirCommentaires = 16
Const ConstPersoAllerVendre = 17
Const ConstPersoAllerMaison = 18
Const ConstPersoAllerChateau = 19
Const ConstPersoDefinirIACible = 20
Const ConstPersoAttaquerPerso = 21
Const ConstPersoAttaquerMaison = 22
Const ConstPersoAttaquerChateau = 23
Const ConstPersoAllerTresor = 24
Const ConstPersoAllerDonner = 25
Const ConstPersoObjetDeSelectionne = 26
Const ConstPersoOrdonnerChargez = 27
Const ConstPersoOrdonnerChargez2 = 28
Const ConstPersoOrdonnerFormation = 51
Const ConstPersoRecupererCarac = 29
Const ConstPersoObjetCreer = 30
Const ConstPersoObjetArreterFabrication = 61
Const ConstPersoObjetMaisonPrendre = 63
Const ConstPersoObjetMaisonDeposer = 69
Const ConstPersoObjetMaisonAcheter = 64
Const ConstPersoObjetChateauVendre = 70
Const ConstPersoUtiliserService = 31
Const ConstPersoDetruireMaison = 32
Const ConstPersoRessourceManger = 33
Const ConstPersoRessourcesDeposerMaison = 34
Const ConstPersoRessourcesAcheterMaison = 62
Const ConstPersoRessourcesAcheterChateau = 57
Const ConstPersoRessourcesVendreChateau = 35
Const ConstPersoTuerPerso = 36
Const ConstPersoTrouverRessource = 37
Const ConstPersoTrouverPersoEnnemi = 38
Const ConstPersoTrouverPersoEnnemi2 = 39
Const ConstPersoTrouverMaisonEnnemi = 40
Const ConstPersoTrouverMaisonEnnemi2 = 41
Const ConstPersoTrouverChateauEnnemi = 42
Const ConstPersoTrouverChateauEnnemi2 = 43
Const ConstPersoTrouverTresor = 44
Const ConstPersoRessourcePrendre = 45
Const ConstPersoRessourceLaisser = 46
Const ConstPersoArgentPrendre = 77
Const ConstPersoArgentMaisonPrendre = 47
Const ConstPersoArgentMaisonLaisser = 48
Const ConstPersoRessourceMaisonPrendre = 58
Const ConstPersoRessourceMaisonLaisser = 59
Const ConstPersoAppelerDelAide = 49
Const ConstPersoArreter = 50
Const ConstPersoObjetRanger = 52
Const ConstPersoObjetEquiper = 53
Const ConstPersoObjetEquiperADroite = 74
Const ConstPersoObjetSelectionne = 65
Const ConstPersoObjetEquipesRanger = 76
Const ConstPersoObjetEquipesDetruire = 68
Const ConstPersoObjetInventaireAjouter = 75
Const ConstPersoObjetInventaireCharger = 72
Const ConstPersoObjetInventaireDetruire = 66
Const ConstPersoObjetInventaireNombre = 71
Const ConstPersoObjetMaisonAjouter = 10
Const ConstPersoObjetMaisonCharger = 73
Const ConstPersoObjetMaisonDetruire = 67
Const ConstPersoVendreRessourceFief = 54
Const ConstPersoPrendreRessource = 55
Const ConstPersoDeposerRessource = 56
Const ConstPersoRestaurerChateau = 60
Const ConstPersoEffetTempEffacer = 79
Const ConstPersoEnvoyerCommentaire = 80
Const ConstPersoHostile = 82
Const ConstPersoDevenirEmpereur = 83
Const ConstPersoAttaquerAllies = 84

Private Commandes As Collection

Public Sub Init_Commandes()
    Set Commandes = New Collection
End Sub

Public Sub Ajouter_Commande(ByVal Categorie As Integer, _
                            ByVal Indice As Long, _
                            Optional ByVal Var1 As String, _
                            Optional ByVal Var2 As String)
    Dim Commande As ClsCommande
    Set Commande = New ClsCommande
    Commande.Categorie = Categorie
    Commande.Indice = Indice
    Commande.Var1 = Var1
    Commande.Var2 = Var2
    If ComReseau.Client Then
        ComReseau.Message_Envoyer ComMessages.Definir_CategorieCommandes, Categorie, Indice, , Var1, Var2
    Else
        Commandes.Add Commande
    End If
End Sub
Private Sub Supprimer_Commandes()
    Collection_Vider Commandes
End Sub

Public Sub Lire_Commandes()
    On Error GoTo Erreur
    Dim i As Integer
    'If Debugger Then Messages(0).Effacer
    For i = 1 To Commandes.Count
        With Commandes(i)
        If Debugger Then Messages(0).Ajouter_Message .Categorie & " : " & .Indice & " : " & .Var1 & " " & .Var2
        Select Case .Categorie
        Case ConstPersoDeplacer: Persos(.Indice).Aller_A .Var1, .Var2
        Case ConstPersoDeplacerSansCommentaire: Persos(.Indice).Aller_A .Var1, .Var2, True
        Case ConstPersoChercherRessource: Persos(.Indice).Aller_Ressource Ressources(.Var1), .Var1, AffRessources.Largeur, AffRessources.Hauteur
        Case ConstPersoRentrerChateau: Persos(.Indice).Rentrer_Chateau Chateaux(Persos(.Indice).NumeroFief)
        Case ConstPersoCreerMaisonCoordonnees: Maisons(.Indice).Creer .Var1, .Var2, Commandes(i + 1).Var1, .Indice, Parametres ': Persos(.Indice).Rentrer_Maison Maisons(.Indice)
        'Case ConstPersoCreerMaisonType:
        Case ConstPersoRentrerMaison: Persos(.Indice).Rentrer_Maison Maisons(.Indice), .Var1
        Case ConstPersoDefinirIA:
            Persos(.Indice).IA = .Var1
        Case ConstPersoAttaquerTerrainCoordonnees:
            Persos(.Indice).Attaquer_Terrain .Var1, .Var2, Commandes(i + 1).Var1
        'Case ConstPersoAttaquerTerrainCible:
        Case ConstPersoVendreObjetsSelectionnees: Fiefs(.Var1).Vendre_ObjetsSelectionnees Persos(.Indice), Parametres
        Case ConstPersoObjetDeposerSol: Persos(.Indice).Objet_Deposer_Sol .Var1, .Var2
        Case ConstPersoRessourcesDeposerSol: Persos(.Indice).Ressources_Deposer_Sol
        Case ConstPersoDefinirObjetSelectionne: Persos(.Indice).ObjetSelectionne = .Var1
        Case ConstPersoDefinirChefPosition: Persos(.Indice).ChefPositionX = .Var1: Persos(.Indice).ChefPositionY = .Var2
        Case ConstPersoDefinirIAOrdre: Persos(.Indice).IA_Ordre = .Var1
        Case ConstPersoDefinirIAAttitude:
            Persos(.Indice).Attitude = .Var1
            Persos(.Indice).IA_Ordre = 0
            Comportement_IA_Persos .Indice
        Case ConstPersoDefinirCommentaires: Persos(.Indice).Definir_Commentaires = .Var1
        Case ConstPersoAllerVendre: Persos(.Indice).Aller_Vendre .Var1, .Var2
        Case ConstPersoAllerMaison: Persos(.Indice).Aller_Maison Maisons(.Var1), .Var2
        Case ConstPersoAllerChateau: Persos(.Indice).Aller_Chateau Chateaux(.Var1), .Var1
        Case ConstPersoDefinirIACible: Persos(.Indice).IA_Cible = .Var1
        Case ConstPersoAttaquerPerso: Persos(.Indice).Attaquer_Perso .Var1
        Case ConstPersoAttaquerMaison: Persos(.Indice).Attaquer_Maison .Var1
        Case ConstPersoAttaquerChateau: Persos(.Indice).Attaquer_Chateau .Var1
        Case ConstPersoAllerTresor: Persos(.Indice).Aller_Tresor Tresors(Int(.Var1)), .Var1, AffTresor.Largeur, AffTresor.Hauteur
        Case ConstPersoAllerDonner: Persos(.Indice).Aller_Donner .Var1
        Case ConstPersoObjetDeSelectionne: Persos(.Indice).Definir_ObjetSelectionne = False
        Case ConstPersoOrdonnerChargez: Ordonner_Chargez .Indice, .Var1, .Var2, Commandes(i + 1).Indice, Commandes(i + 1).Var1, Commandes(i + 1).Var2
        'Case ConstPersoOrdonnerChargez2:
        Case ConstPersoOrdonnerFormation: Ordonner_Formation .Indice, .Var1
        Case ConstPersoRecupererCarac: Select Case .Var1: Case 0: Persos(.Indice).Recuperer_Vie Parametres: Case 1: Persos(.Indice).Recuperer_Energie Parametres: Case 2: Persos(.Indice).Recuperer_Magie Parametres: Case 3: Persos(.Indice).Recuperer_Moral Parametres: End Select
        Case ConstPersoObjetCreer: Persos(.Indice).Creer_Objet .Var1, Maisons(.Indice), Parametres
        Case ConstPersoObjetArreterFabrication: Persos(.Indice).Arreter
                                                Persos(.Indice).EtatFabricationObjet = 0
        Case ConstPersoObjetMaisonPrendre: Persos(.Indice).Objet_Ramasser_Maison .Var1, Maisons(.Indice), Parametres, Commentaires, .Var2
        Case ConstPersoObjetMaisonDeposer: Persos(.Indice).Objet_Deposer_Maison .Var1, Maisons(.Indice), Parametres, .Var2
        Case ConstPersoObjetMaisonAcheter: Persos(.Indice).Objet_Acheter_Maison .Var2, Maisons(.Var1), Parametres, Commentaires
        Case ConstPersoObjetChateauVendre: Persos(.Indice).Objet_Vendre_Chateau .Var1, Fiefs(Persos(.Indice).NumeroFief), Parametres, .Var2
        Case ConstPersoUtiliserService:  Maisons(.Var1).Service_Utiliser Persos(.Indice), Parametres
        Case ConstPersoDetruireMaison:
            Maisons(.Indice).Demolir Persos(.Indice), Parametres
        Case ConstPersoRessourceManger: Persos(.Indice).Manger_Ressources .Var1, Parametres
        Case ConstPersoRessourcesDeposerMaison: Maisons(.Indice).Transfert_Stock_Deposer .Var1, Persos(.Indice), Parametres, .Var2
        Case ConstPersoRessourcesAcheterMaison: Maisons(.Var1).Acheter .Var2, Persos(.Indice), Parametres
        Case ConstPersoRessourcesAcheterChateau:
            Fiefs(Persos(.Indice).NumeroFief).Acheter .Var1, .Var2, Persos(.Indice), Parametres
        Case ConstPersoRessourcesVendreChateau: Fiefs(Persos(.Indice).NumeroFief).Vendre .Var1, .Var2, Persos(.Indice), Parametres
        Case ConstPersoTuerPerso: 'Perso_Tuer_Perso Persos(.Indice), Persos(.Var1)
                                  Persos(.Var1).Vivant = False
        Case ConstPersoTrouverRessource: IA_Trouver_Ressource .Indice, .Var1
        Case ConstPersoTrouverPersoEnnemi:
            IA_Trouver_Ennemi .Indice, .Var1, .Var2, Commandes(i + 1).Var1, Commandes(i + 1).Var2
        'Case ConstPersoTrouverPersoEnnemi2:
        Case ConstPersoTrouverMaisonEnnemi: IA_Trouver_MaisonEnnemi .Indice, .Var1, .Var2, Commandes(i + 1).Var1, Commandes(i + 1).Var2
        'Case ConstPersoTrouverMaisonEnnemi2:
        Case ConstPersoTrouverChateauEnnemi: IA_Trouver_ChateauEnnemi .Indice, .Var1, .Var2, Commandes(i + 1).Var1, Commandes(i + 1).Var2
        'Case ConstPersoTrouverChateauEnnemi2:
        Case ConstPersoTrouverTresor: IA_Trouver_Tresor .Indice
        Case ConstPersoRessourcePrendre: Persos(.Indice).Ressources_Prendre .Var1, .Var2
        Case ConstPersoRessourceLaisser: Persos(.Indice).Ressources_Prendre .Var1, .Var2, , True
        Case ConstPersoArgentPrendre: Persos(.Indice).Argent_Prendre .Var1
        Case ConstPersoArgentMaisonPrendre: Maisons(.Indice).Transfert_Argent_Prendre .Var1, Persos(.Indice)
        Case ConstPersoArgentMaisonLaisser: Maisons(.Indice).Transfert_Argent_Deposer .Var1, Persos(.Indice)
        Case ConstPersoRessourceMaisonPrendre: Maisons(.Indice).Transfert_Stock_Prendre .Var1, Persos(.Indice), Parametres
        Case ConstPersoRessourceMaisonLaisser: Maisons(.Indice).Transfert_Stock_Deposer .Var1, Persos(.Indice), Parametres
        Case ConstPersoAppelerDelAide: Appeler_De_lAide Persos(.Indice)
        Case ConstPersoArreter: Persos(.Indice).Arreter
        Case ConstPersoObjetRanger: Persos(.Indice).Objet_Ranger .Var1, Parametres, Commentaires
        Case ConstPersoObjetEquiper: Persos(.Indice).Objet_Equiper .Var1, Parametres, Commentaires, .Var2
        Case ConstPersoObjetEquiperADroite: Persos(.Indice).Objet_Equiper .Var1, Parametres, Commentaires, .Var2, True
        Case ConstPersoObjetSelectionne: Persos(.Indice).ObjetSelectionne = .Var1
        Case ConstPersoObjetEquipesRanger: Persos(.Indice).Objet_Ranger .Var1, Parametres, Commentaires
        Case ConstPersoObjetEquipesDetruire: Persos(.Indice).Objet_Equipes_Detruire .Var1
        Case ConstPersoObjetInventaireAjouter: Persos(.Indice).Objet_Ajouter_Inventaire .Var1, Parametres, Commentaires, True
        Case ConstPersoObjetInventaireCharger: Persos(.Indice).Objet_Inventaire_Charger .Var1, .Var2, Parametres
        Case ConstPersoObjetInventaireDetruire: Persos(.Indice).Objet_Inventaire_Detruire .Var1
        Case ConstPersoObjetInventaireNombre: Persos(.Indice).Objet_Inventaire_Nombre(.Var1) = .Var2
        Case ConstPersoObjetMaisonAjouter: Persos(.Indice).Objet_DeposerInventaire_Maison .Var1, Maisons(.Indice), Parametres
        Case ConstPersoObjetMaisonCharger: Maisons(.Indice).Definir_ObjetsInventaire(.Var1) = .Var2
        Case ConstPersoObjetMaisonDetruire: Maisons(.Indice).Definir_ObjetsInventaire(.Var1) = -1
        Case ConstPersoPrendreRessource: Persos(.Indice).Ressources_Prendre .Var1, .Var2
        Case ConstPersoDeposerRessource: Persos(.Indice).Ressources_Prendre .Var1, .Var2, False, True
        Case ConstPersoRestaurerChateau: Fiefs(Persos(.Indice).NumeroFief).Restaurer_Chateau .Var1, Persos(.Indice)
        Case ConstPersoEffetTempEffacer: Persos(.Indice).EffetTemp_Effacer .Var1
        Case ConstPersoEnvoyerCommentaire: Persos(.Indice).Definir_Commentaires = ""
                                           Persos(.Indice).Definir_Commentaires = .Var1
        Case ConstPersoHostile: Perso_Hostile Persos(.Indice)
        Case ConstPersoDevenirEmpereur: Fiefs(Persos(.Indice).NumeroFief).Empereur_Changer .Indice, True
        Case ConstPersoAttaquerAllies: Persos(.Indice).Attaquer_Allies True
        End Select
        End With
    Next i
    Supprimer_Commandes
    Exit Sub
Erreur:
    Supprimer_Commandes
End Sub

Public Sub Cmd_Perso_Deplacer(ByVal i As Long, _
                              ByVal X As Long, _
                              ByVal Y As Long)
    'Déplace le personnage i aux positions X, Y.
    Ajouter_Commande ConstPersoDeplacer, i, X, Y
End Sub

Public Sub Cmd_Perso_DeplacerSansCommentaire(ByVal i As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long)
    'Déplace le personnage i aux positions X, Y.
    Ajouter_Commande ConstPersoDeplacerSansCommentaire, i, X, Y
End Sub

Public Sub Cmd_Perso_Chercher_Ressource(ByVal i As Long, _
                                        ByVal NoRessource As Long)
    'Le personnage va à une ressource donnée.
    Ajouter_Commande ConstPersoChercherRessource, i, NoRessource
End Sub

Public Sub Cmd_Perso_Rentrer_Chateau(ByVal i As Long)
    'Renvoie le personnage dans sa maison.
    Ajouter_Commande ConstPersoRentrerChateau, i
End Sub

Public Sub Cmd_Perso_Creer_Maison(ByVal i As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal TypeMaison As Integer)
    'Le personnage crée sa maison.
    Ajouter_Commande ConstPersoCreerMaisonCoordonnees, i, X, Y
    Ajouter_Commande ConstPersoCreerMaisonType, i, TypeMaison
End Sub

Public Sub Cmd_Perso_Rentrer_Maison(ByVal i As Long, ByVal Rentrer As Boolean)
    'Renvoie le personnage dans sa maison.
    Ajouter_Commande ConstPersoRentrerMaison, i, Rentrer
End Sub

Public Sub Cmd_Perso_Definir_IA(ByVal i As Long, ByVal IA As Boolean)
    If Persos(i).ChoisirIA Then 'On ne peut activer une IA que si on en a déjà choisi une.
        Ajouter_Commande ConstPersoDefinirIA, i, IA
    End If
End Sub

Public Sub Cmd_Perso_AttaquerTerrain(ByVal i As Long, _
                                     ByVal X As Long, _
                                     ByVal Y As Long, _
                                     ByVal TypeCible As Integer)
    Ajouter_Commande ConstPersoAttaquerTerrainCoordonnees, i, X, Y
    Ajouter_Commande ConstPersoAttaquerTerrainCible, i, TypeCible
End Sub

Public Sub Cmd_Perso_VendreObjetsSelectionnees(ByVal i As Long, _
                                               ByVal Fief As Long)
    Ajouter_Commande ConstPersoVendreObjetsSelectionnees, i, Fief
End Sub

Public Sub Cmd_PersoObjetDeposerSol(ByVal i As Long, _
                                    ByVal Objet As Integer, _
                                    ByVal Inventaire As Boolean)
    Ajouter_Commande ConstPersoObjetDeposerSol, i, Objet, Inventaire
End Sub

Public Sub Cmd_PersoRessourcesDeposerSol(ByVal i As Long)
    Ajouter_Commande ConstPersoRessourcesDeposerSol, i
End Sub

Public Sub Cmd_Perso_DefinirObjetSelectionne(ByVal i As Long, _
                                             ByVal Objet As Integer)
    Ajouter_Commande ConstPersoDefinirObjetSelectionne, i, Objet
End Sub

Public Sub Cmd_Perso_DefinirChefPosition(ByVal i As Long, _
                                         ByVal X As Long, _
                                         ByVal Y As Long)
    Ajouter_Commande ConstPersoDefinirChefPosition, i, X, Y
End Sub

Public Sub Cmd_Perso_DefinirIAOrdre(ByVal i As Long, _
                                    ByVal Ordre As Integer)
    Ajouter_Commande ConstPersoDefinirIAOrdre, i, Ordre
End Sub

Public Sub Cmd_Perso_DefinirIAAttitude(ByVal i As Long, _
                                       ByVal Attitude As Integer)
    Ajouter_Commande ConstPersoDefinirIAAttitude, i, Attitude
End Sub

Public Sub Cmd_Perso_DefinirCommentaires(ByVal i As Long, _
                                         ByVal Commentaire As String)
    Ajouter_Commande ConstPersoDefinirCommentaires, i, Commentaire
End Sub

Public Sub Cmd_Perso_AllerVendre(ByVal i As Long, _
                                 ByVal Fief As Integer, ByVal Marche As Boolean)
    Ajouter_Commande ConstPersoAllerVendre, i, Fief, Marche
End Sub

Public Sub Cmd_Perso_AllerMaison(ByVal i As Long, _
                                 ByVal Maison As Long, _
                                 ByVal Reparer As Boolean)
    Ajouter_Commande ConstPersoAllerMaison, i, Maison, Reparer
End Sub

Public Sub Cmd_Perso_AllerChateau(ByVal i As Long, _
                                  ByVal Chateau As Integer)
    Ajouter_Commande ConstPersoAllerChateau, i, Chateau
End Sub

Public Sub Cmd_Perso_DefinirIACible(ByVal i As Long, _
                                    ByVal Cible As Long)
    Ajouter_Commande ConstPersoDefinirIACible, i, Cible
End Sub

Public Sub Cmd_Perso_AttaquerPerso(ByVal i As Long, _
                                   ByVal Cible As Long)
    Ajouter_Commande ConstPersoAttaquerPerso, i, Cible
End Sub

Public Sub Cmd_Perso_AttaquerMaison(ByVal i As Long, _
                                     ByVal Cible As Long)
    Ajouter_Commande ConstPersoAttaquerMaison, i, Cible
End Sub

Public Sub Cmd_Perso_AttaquerChateau(ByVal i As Long, _
                                     ByVal Cible As Long)
    Ajouter_Commande ConstPersoAttaquerChateau, i, Cible
End Sub

Public Sub Cmd_Perso_AllerTresor(ByVal i As Long, _
                                 ByVal Cible As Long)
    Ajouter_Commande ConstPersoAllerTresor, i, Cible
End Sub

Public Sub Cmd_Perso_AllerDonner(ByVal i As Long, _
                                 ByVal Cible As Long)
    Ajouter_Commande ConstPersoAllerDonner, i, Cible
End Sub

Public Sub Cmd_Perso_ObjetDeSelectionne(ByVal i As Long)
    Ajouter_Commande ConstPersoObjetDeSelectionne, i
End Sub

Public Sub Cmd_PersoOrdonnerChargez(ByVal IndicePerso As Long, _
                                    Optional ByVal TypeCible As Integer = -1, _
                                    Optional ByVal IndiceCible As Long = -1, _
                                    Optional ByVal Ctrl As Boolean, _
                                    Optional ByVal Shift As Boolean, _
                                    Optional ByVal Alt As Boolean)
    Ajouter_Commande ConstPersoOrdonnerChargez, IndicePerso, TypeCible, IndiceCible
    Ajouter_Commande ConstPersoOrdonnerChargez2, Ctrl, Shift, Alt
End Sub

Public Sub Cmd_Perso_OrdonnerFormation(ByVal i As Long, _
                                       ByVal ChangerFormation As Boolean)
    Ajouter_Commande ConstPersoOrdonnerFormation, i, ChangerFormation
End Sub

Public Sub Cmd_Perso_RecupererCarac(ByVal i As Long, _
                                    ByVal Carac As Integer)
    Ajouter_Commande ConstPersoRecupererCarac, i, Carac
End Sub

Public Sub Cmd_Perso_CreerObjet(ByVal i As Long, _
                                ByVal Objet As Integer)
    Ajouter_Commande ConstPersoObjetCreer, i, Objet
End Sub

Public Sub Cmd_Perso_ObjetArreterFabrication(ByVal i As Long)
    Ajouter_Commande ConstPersoObjetArreterFabrication, i
End Sub

Public Sub Cmd_Perso_ObjetMaisonPrendre(ByVal i As Long, _
                                        ByVal IndiceObjet As Integer, _
                                        ByVal Quantite As Integer)
    Ajouter_Commande ConstPersoObjetMaisonPrendre, i, IndiceObjet, Quantite
End Sub

Public Sub Cmd_Perso_ObjetMaisonDeposer(ByVal i As Long, _
                                        ByVal IndiceObjet As Integer, _
                                        ByVal Inventaire As Boolean)
    Ajouter_Commande ConstPersoObjetMaisonDeposer, i, IndiceObjet, Inventaire
End Sub

Public Sub Cmd_Perso_ObjetMaisonAcheter(ByVal i As Long, _
                                        ByVal IndiceMagasin As Long, _
                                        ByVal IndiceObjet As Integer, _
                                        ByVal Quantite As Integer)
    Dim j As Integer
    For j = 1 To Quantite
        Ajouter_Commande ConstPersoObjetMaisonAcheter, i, IndiceMagasin, IndiceObjet
    Next j
End Sub

Public Sub Cmd_Perso_ObjetChateauVendre(ByVal i As Long, _
                                        ByVal IndiceObjet As Integer, _
                                        ByVal Inventaire As Boolean)
    Ajouter_Commande ConstPersoObjetChateauVendre, i, IndiceObjet, Inventaire
End Sub

Public Sub Cmd_Perso_UtiliserService(ByVal i As Long, _
                                     ByVal Maison As Long)
    Ajouter_Commande ConstPersoUtiliserService, i, Maison
End Sub

Public Sub Cmd_Perso_DetuireMaison(ByVal i As Long)
    Ajouter_Commande ConstPersoDetruireMaison, i
End Sub

Public Sub Cmd_Perso_RessourceManger(ByVal i As Long, _
                                     ByVal Ressource As Integer)
    Ajouter_Commande ConstPersoRessourceManger, i, Ressource
End Sub

Public Sub Cmd_Perso_RessourcesDeposerMaison(ByVal i As Long, _
                                             ByVal Quantite As Double, _
                                             Optional ByVal Max As Boolean)
    Ajouter_Commande ConstPersoRessourcesDeposerMaison, i, Quantite, Max
End Sub

Public Sub Cmd_Perso_RessourcesAcheterMaison(ByVal i As Long, _
                                             ByVal IndiceMaison As Long, _
                                             ByVal Quantite As Double)
    Ajouter_Commande ConstPersoRessourcesAcheterMaison, i, IndiceMaison, Quantite
End Sub

Public Sub Cmd_Perso_RessourcesAcheterChateau(ByVal i As Long, _
                                              ByVal Ressource As Integer, _
                                              ByVal Quantite As Double)
    Ajouter_Commande ConstPersoRessourcesAcheterChateau, i, Ressource, Quantite
End Sub

Public Sub Cmd_Perso_RessourcesVendreChateau(ByVal i As Long, _
                                             ByVal Ressource As Integer, _
                                             ByVal Quantite As Double)
    Ajouter_Commande ConstPersoRessourcesVendreChateau, i, Ressource, Quantite
End Sub

Public Sub Cmd_Perso_TuerPerso(ByVal i As Long, _
                               ByVal Cible As Long)
    Ajouter_Commande ConstPersoTuerPerso, i, Cible
End Sub

Public Sub Cmd_Perso_TrouverRessource(ByVal i As Long, _
                                      ByVal TypeRessource As Integer)
    Ajouter_Commande ConstPersoTrouverRessource, i, TypeRessource
End Sub

Public Sub Cmd_Perso_TrouverPersoEnnemi(ByVal i As Long, _
                                        Optional ByVal EnvoyerSoldats As Boolean, _
                                        Optional ByVal Ctrl As Boolean, _
                                        Optional ByVal Shift As Boolean, _
                                        Optional ByVal Alt As Boolean)
    Ajouter_Commande ConstPersoTrouverPersoEnnemi, i, EnvoyerSoldats, Ctrl
    Ajouter_Commande ConstPersoTrouverPersoEnnemi2, i, Shift, Alt
End Sub

Public Sub Cmd_Perso_TrouverMaisonEnnemi(ByVal i As Long, _
                                         Optional ByVal EnvoyerSoldats As Boolean, _
                                         Optional ByVal Ctrl As Boolean, _
                                         Optional ByVal Shift As Boolean, _
                                         Optional ByVal Alt As Boolean)
    Ajouter_Commande ConstPersoTrouverMaisonEnnemi, i, EnvoyerSoldats, Ctrl
    Ajouter_Commande ConstPersoTrouverMaisonEnnemi2, i, Shift, Alt
End Sub

Public Sub Cmd_Perso_TrouverChateauEnnemi(ByVal i As Long, _
                                          Optional ByVal EnvoyerSoldats As Boolean, _
                                          Optional ByVal Ctrl As Boolean, _
                                          Optional ByVal Shift As Boolean, _
                                          Optional ByVal Alt As Boolean)
    Ajouter_Commande ConstPersoTrouverChateauEnnemi, i, EnvoyerSoldats, Ctrl
    Ajouter_Commande ConstPersoTrouverChateauEnnemi2, i, Shift, Alt
End Sub

Public Sub Cmd_Perso_TrouverTresor(ByVal i As Long)
    Ajouter_Commande ConstPersoTrouverTresor, i
End Sub

Public Sub Cmd_Perso_RessourcePrendre(ByVal i As Long, _
                                      ByVal TypeRessource As Integer, _
                                      ByVal Quantite As Double)
    Ajouter_Commande ConstPersoRessourcePrendre, i, TypeRessource, Quantite
End Sub

Public Sub Cmd_Perso_RessourceLaisser(ByVal i As Long, _
                                      ByVal TypeRessource As Integer, _
                                      ByVal Quantite As Double)
    Ajouter_Commande ConstPersoRessourceLaisser, i, TypeRessource, Quantite
End Sub

Public Sub Cmd_Perso_ArgentPrendre(ByVal i As Long, _
                                   ByVal Quantite As Double)
    Ajouter_Commande ConstPersoArgentPrendre, i, Quantite
End Sub

Public Sub Cmd_Perso_ArgentMaisonPrendre(ByVal i As Long, _
                                         ByVal Quantite As Double)
    Ajouter_Commande ConstPersoArgentMaisonPrendre, i, Quantite
End Sub

Public Sub Cmd_Perso_ArgentMaisonLaisser(ByVal i As Long, _
                                   ByVal Quantite As Double)
    Ajouter_Commande ConstPersoArgentMaisonLaisser, i, Quantite
End Sub

Public Sub Cmd_Perso_RessourceMaisonPrendre(ByVal i As Long, _
                                            ByVal Quantite As Double)
    Ajouter_Commande ConstPersoRessourceMaisonPrendre, i, Quantite
End Sub

Public Sub Cmd_Perso_RessourceMaisonLaisser(ByVal i As Long, _
                                            ByVal Quantite As Double)
    Ajouter_Commande ConstPersoRessourceMaisonLaisser, i, Quantite
End Sub

Public Sub Cmd_Perso_AppelerDelAide(ByVal i As Long)
    Ajouter_Commande ConstPersoAppelerDelAide, i
End Sub

Public Sub Cmd_Perso_Arreter(ByVal i As Long)
    Ajouter_Commande ConstPersoArreter, i
End Sub

Public Sub Cmd_Perso_ObjetRanger(ByVal i As Long, ByVal NumeroObjetEquipe As Integer)
    Ajouter_Commande ConstPersoObjetRanger, i, NumeroObjetEquipe
End Sub

Public Sub Cmd_Perso_ObjetEquiper(ByVal i As Long, ByVal NumeroObjet As Integer, Optional ByVal Inventaire As Boolean = True)
    Ajouter_Commande ConstPersoObjetEquiper, i, NumeroObjet, Inventaire
End Sub

Public Sub Cmd_Perso_ObjetEquiperADroite(ByVal i As Long, ByVal NumeroObjet As Integer, Optional ByVal Inventaire As Boolean = True)
    Ajouter_Commande ConstPersoObjetEquiperADroite, i, NumeroObjet, Inventaire
End Sub

Public Sub Cmd_Perso_ObjetSelectionne(ByVal i As Long, ByVal IndiceObjet As Integer)
    Ajouter_Commande ConstPersoObjetSelectionne, i, IndiceObjet
End Sub

Public Sub Cmd_Perso_ObjetEquipesRanger(ByVal i As Long, ByVal IndiceObjetEquipes As Integer)
    Ajouter_Commande ConstPersoObjetEquipesRanger, i, IndiceObjetEquipes
End Sub

Public Sub Cmd_Perso_ObjetEquipesDetruire(ByVal i As Long, ByVal IndiceInventaire As Integer)
    Ajouter_Commande ConstPersoObjetEquipesDetruire, i, IndiceInventaire
End Sub

Public Sub Cmd_Perso_ObjetInventaireAjouter(ByVal i As Long, ByVal IndiceObjet As Integer)
    Ajouter_Commande ConstPersoObjetInventaireAjouter, i, IndiceObjet
End Sub

Public Sub Cmd_Perso_ObjetInventaireCharger(ByVal i As Long, ByVal IndiceInventaire As Integer, ByVal IndiceObjet As Integer)
    Ajouter_Commande ConstPersoObjetInventaireCharger, i, IndiceInventaire, IndiceObjet
End Sub

Public Sub Cmd_Perso_ObjetInventaireDetruire(ByVal i As Long, ByVal IndiceInventaire As Integer)
    Ajouter_Commande ConstPersoObjetInventaireDetruire, i, IndiceInventaire
End Sub

Public Sub Cmd_Perso_ObjetInventaireNombre(ByVal i As Long, ByVal IndiceInventaire As Integer, ByVal Quantite As Integer)
    Ajouter_Commande ConstPersoObjetInventaireNombre, i, IndiceInventaire, Quantite
End Sub

Public Sub Cmd_Perso_ObjetMaisonAjouter(ByVal i As Long, ByVal IndiceObjet As Integer)
    Ajouter_Commande ConstPersoObjetMaisonAjouter, i, IndiceObjet
End Sub

Public Sub Cmd_Perso_ObjetMaisonCharger(ByVal i As Long, ByVal IndiceInventaire As Integer, ByVal IndiceObjet As Integer)
    Ajouter_Commande ConstPersoObjetMaisonCharger, i, IndiceInventaire, IndiceObjet
End Sub

Public Sub Cmd_Perso_ObjetMaisonDetruire(ByVal i As Long, ByVal IndiceInventaire As Integer)
    Ajouter_Commande ConstPersoObjetMaisonDetruire, i, IndiceInventaire
End Sub

Public Sub Cmd_Perso_PrendreRessource(ByVal i As Long, _
                                      ByVal TypeRessource As Integer, _
                                      ByVal Quantite As Double, _
                                      Optional ByVal Negatif As Boolean)
    If Negatif Then
        Ajouter_Commande ConstPersoDeposerRessource, i, TypeRessource, Quantite
    Else
        Ajouter_Commande ConstPersoPrendreRessource, i, TypeRessource, Quantite
    End If
End Sub

Public Sub Cmd_Perso_RestaurerChateau(ByVal i As Long, ByVal QuantiteArgent As Double)
    Ajouter_Commande ConstPersoRestaurerChateau, i, QuantiteArgent
End Sub

Public Sub Cmd_Perso_EffetTempEffacer(ByVal i As Long, ByVal IndiceEffet As Integer)
    Ajouter_Commande ConstPersoEffetTempEffacer, i, IndiceEffet
End Sub

Public Sub Cmd_Perso_EnvoyerCommentaire(ByVal i As Long, ByVal Message As String)
    Ajouter_Commande ConstPersoEnvoyerCommentaire, i, Message
End Sub

Public Sub Cmd_Perso_Hostile(ByVal i As Long)
    Ajouter_Commande ConstPersoHostile, i
End Sub

Public Sub Cmd_Perso_DevenirEmpereur(ByVal i As Long)
    Ajouter_Commande ConstPersoDevenirEmpereur, i
End Sub

Public Sub Cmd_Perso_AttaquerAllies(ByVal i As Long)
    Ajouter_Commande ConstPersoAttaquerAllies, i
End Sub

