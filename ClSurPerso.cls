VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSurPerso"
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

Private Numero As Long

Public Sub Init(ByVal Perso As ClsJeuPerso)
    Numero = Perso.Numero
End Sub

'' ************************************************************************************
'' *                                     Propriétés                                   *
'' ************************************************************************************

'Régénérations
Public Property Get RegenerationVie() As Single
    RegenerationVie = Persos(Numero).BonusVitesseRegenerationVie
End Property
Public Property Get RegenerationEnergie() As Single
    RegenerationEnergie = Persos(Numero).BonusVitesseRegenerationEnergie
End Property
Public Property Get RegenerationMagie() As Single
    RegenerationMagie = Persos(Numero).BonusVitesseRegenerationMagie
End Property
Public Property Get RegenerationMoral() As Single
    RegenerationMoral = Persos(Numero).BonusVitesseRegenerationMoral
End Property

'Positions
Public Property Get PositionX() As Long
    PositionX = Persos(Numero).PositionX
End Property
Public Property Get PositionY() As Long
    PositionY = Persos(Numero).PositionY
End Property
Public Property Get Vitesse() As Long
    Vitesse = Persos(Numero).Longueur_Pas
End Property
Public Property Get DirectionX() As Long
    DirectionX = Persos(Numero).DirectionX
End Property
Public Property Get DirectionY() As Long
    DirectionY = Persos(Numero).DirectionY
End Property

Public Property Get Nom() As String
    Nom = Persos(Numero).Nom
End Property
Public Property Get Metier() As String
    Metier = Parametres.Batiment_Metier(Maisons(Numero).TypeBatiment, Persos(Numero).Feminin)
End Property

Public Property Get Largeur() As Integer
    Largeur = Persos(Numero).Largeur
End Property
Public Property Get Hauteur() As Integer
    Hauteur = Persos(Numero).Hauteur
End Property
Public Property Get NumeroFief() As Long
    NumeroFief = Persos(Numero).NumeroFief
End Property
Public Property Get NumeroEquipe() As Long
    NumeroEquipe = Persos(Numero).NumeroEquipe
End Property
Public Property Get Epoque() As Integer
    Epoque = Fiefs(Persos(Numero).NumeroFief).Epoque + 1
End Property
Public Property Get Invocation() As Boolean
    Invocation = Persos(Numero).Invocation
End Property
Public Property Get NomChef() As String
    NomChef = Persos(Persos(Numero).NumeroChef).Nom
End Property
Public Property Get Incommandable() As Boolean
    Incommandable = Persos(Numero).Incommandable
End Property
Public Property Get ChefDecalageX() As Long
    ChefDecalageX = Persos(Numero).ChefDecalageX
End Property
Public Property Get ChefDecalageY() As Long
    ChefDecalageY = Persos(Numero).ChefDecalageY
End Property
Public Property Get Attitude() As Integer
    Attitude = Persos(Numero).Attitude
End Property
Public Property Get Formation() As Integer
    Formation = Persos(Numero).Formation
End Property
Public Property Get Empereur() As Boolean
    Empereur = Persos(Numero).Empereur
End Property
Public Property Get NombreSoldats() As Long
    NombreSoldats = Persos(Numero).NombreSoldats
End Property
Public Property Get NomRace() As String
    NomRace = Parametres.Race_Nom(Persos(Numero).Race, Persos(Numero).Feminin)
End Property
Public Property Get Parole() As Boolean
    Parole = Persos(Numero).Parole
End Property
Public Property Get Volant() As Boolean
    Volant = Persos(Numero).Volant
End Property
Public Property Get ChoisirApparence() As Boolean
    ChoisirApparence = Persos(Numero).ChoisirApparence
End Property
Public Property Get FichierApparence() As String
    FichierApparence = Persos(Numero).FichierApparence
End Property
Public Property Get Titre() As Integer
    Titre = Persos(Numero).Titre
End Property
Public Property Get Feminin() As Boolean
    Feminin = Persos(Numero).Feminin
End Property
Public Property Get GainExperience() As Boolean
    GainExperience = Persos(Numero).GainExperience
End Property
Public Property Get DureeRessurection() As Double
    DureeRessurection = Persos(Numero).DureeRessurection
End Property
Public Property Get CommentairesSpecifiques() As Boolean
    CommentairesSpecifiques = Persos(Numero).CommentairesSpecifique
End Property

Public Property Get Vie() As Double
    Vie = Persos(Numero).Vie
End Property
Public Property Get MaxVie() As Double
    MaxVie = Persos(Numero).Definir_MaxVie
End Property
Public Property Get ExpVie() As Single
    ExpVie = Persos(Numero).Progression_Vie * 100
End Property

Public Property Get Energie() As Double
    Energie = Persos(Numero).Energie
End Property
Public Property Get MaxEnergie() As Double
    MaxEnergie = Persos(Numero).Definir_MaxEnergie
End Property
Public Property Get ExpEnergie() As Single
    ExpEnergie = Persos(Numero).Progression_Energie * 100
End Property

Public Property Get Magie() As Double
    Magie = Persos(Numero).Magie
End Property
Public Property Get MaxMagie() As Double
    MaxMagie = Persos(Numero).Definir_MaxMagie
End Property
Public Property Get ExpMagie() As Single
    ExpMagie = Persos(Numero).Progression_Magie * 100
End Property

Public Property Get Moral() As Double
    Moral = Persos(Numero).Moral
End Property
Public Property Get MaxMoral() As Double
    MaxMoral = Persos(Numero).Definir_MaxMoral
End Property
Public Property Get ExpMoral() As Single
    ExpMoral = Persos(Numero).Progression_Moral * 100
End Property

Public Property Get Attaque() As Double
    Attaque = Persos(Numero).Attaque
End Property
Public Property Get MaxAttaque() As Double
    MaxAttaque = Persos(Numero).Definir_MaxAttaque
End Property
Public Property Get ExpAttaque() As Single
    ExpAttaque = Persos(Numero).Progression_Attaque * 100
End Property

Public Property Get Defense() As Double
    Defense = Persos(Numero).Defense
End Property
Public Property Get MaxDefense() As Double
    MaxDefense = Persos(Numero).Definir_MaxDefense
End Property
Public Property Get ExpDefense() As Single
    ExpDefense = Persos(Numero).Progression_Defense * 100
End Property

Public Property Get PorteeAttaque() As Integer
    PorteeAttaque = Persos(Numero).Definir_PorteeAttaque
End Property
Public Property Get EffetAttaque() As String
    EffetAttaque = Parametres.Effet_Nom(Persos(Numero).EffetAttaque)
End Property
Public Property Get SurfaceAttaque() As Integer
    SurfaceAttaque = Persos(Numero).BonusSurfaceAttaque
End Property
Public Property Get Kamikaze() As Boolean
    Kamikaze = Persos(Numero).Kamikaze
End Property

Public Property Get PerteVie() As Double
    PerteVie = Persos(Numero).PerteVie
End Property
Public Property Get PerteEnergie() As Double
    PerteEnergie = Persos(Numero).PerteEnergie
End Property
Public Property Get PerteMagie() As Double
    PerteMagie = Persos(Numero).PerteMagie
End Property
Public Property Get PerteMoral() As Double
    PerteMoral = Persos(Numero).PerteMoral
End Property
Public Property Get PerteArgent() As Double
    PerteArgent = Persos(Numero).PerteArgent
End Property
Public Property Get PerteRessourceIndice() As Integer
    PerteRessourceIndice = Persos(Numero).PerteRessourceIndice
End Property
Public Property Get PerteRessourceQuantite() As Double
    PerteRessourceQuantite = Persos(Numero).PerteRessourceQuantite
End Property
Public Property Get Armure() As Integer
    Armure = Persos(Numero).Definir_Armure
End Property

'Public IA As Boolean
'Public ChoisirIA As Boolean
'Public ChoisirIAScript As Boolean
'Public IA_Equiper As integer
'Public IA_Manger As Boolean
'Public IA_Pacifique As Boolean
'Public IA_Temeraire As Boolean
'Public IA_Berserk As integer 'Attaque systématiquement à chaque tour.
'Public IA_Ordre As Integer '1=Garder Maison, 2=Garder Chateau.
'Public FichierIA As String
'Public FichierIAScript As String

'

'
'Public ControleJoueur As Boolean
'Private Argent As Double
'Public Action As Integer
'Public Direction As Integer
'Public EnDeplacement As Boolean
'Public DureeImmobilisation As Double 'La latence est le temps en nombre de pas où le personnage est immobilisé.
'Public DansUneMaison As Boolean 'Renvoie vrai si le personnage est à l'intérieur d'une maison.
'Public DansSaMaison As Boolean
'Public IndiceMaison As Long 'Mémorise l'indice de la maison vers laquelle le joueur se dirige.
'Public DansUnChateau As Boolean 'Renvoie vrai si le personnage est à l'intérieur d'un chateau.
'Public DansSonChateau As Boolean
'
''Mémorise les indices des destinations vers lesquels le joueur se dirige.
'Public IndiceRessource As Long
'Public IndiceChateau As Long 'Mémorise l'indice du chateau vers lequel le joueur se dirige.
Public Property Get PersoCible() As Long
    PersoCible = Persos(Numero).IndicePerso
End Property
Public Property Let PersoCible(ByVal Valeur As Long)
    Persos(Numero).IndicePerso = Valeur
End Property
'Public IndiceTresor As Long
'
''Variables de ressources.
''Indique l'indice+1 de la ressource que l'on est en train de récolter
''0 indique que le perso ne récolte pas de ressource.
'Public RessourceCourante As Long
'Public ServiceCourant As Long
''Stock de ressources.
'Private Ressources() As Double
'Private MaxRessources() As Long
'Private BonusMaxRessources As Long
'
''Compétences spéciales.
'Private CaracSpeciales() As Long
'Private ExpSpeciales() As Double
'Private NivSpeciales() As Long
'Private BonusNivSpeciales() As Long
'Private ExpNivSpeciales() As Double
'Private ExpNivPlusSpeciales() As Double
''Compétences des ressources.
'Private ExpRessources() As Double
'Private NivRessources() As Long
'Private BonusNivRessources() As Long
'Private ExpNivRessources() As Double
'Private ExpNivPlusRessources() As Double
''Compétences des Services.
'Private CaracServices() As Long
'Private ExpServices() As Double
'Private NivServices() As Long
'Private BonusNivServices() As Long
'Private ExpNivServices() As Double
'Private ExpNivPlusServices() As Double
''Compétences des objets.
'Public CaracObjetsDepart As Long
'Private CaracObjets() As Long
'Private ExpObjets() As Double
'Private NivObjets() As Long
'Private BonusNivObjets() As Long
'Private ExpNivObjets() As Double
'Private ExpNivPlusObjets() As Double
'
''Modificateurs de la race
'Public CoefRaceVie As Single
'Public CoefRaceEnergie As Single
'Public CoefRaceMagie As Single
'Public CoefRaceMoral As Single
'Public CoefRaceAttaque As Single
'Public CoefRaceDefense As Single
'Private CoefRaceCompSpeciales() As Single 'Compétences spéciales.
'Private CoefRaceCompRessources() As Single 'Compétences des ressources.
'Private CoefRaceCompServices() As Single 'Compétences des Services.
'Private CoefRaceCompObjets() As Single 'Compétences des objets.
'
''Objets.
'Public IndiceObjet As Integer 'Indice de l'objet que le joueur fabrique.
'Public ObjetSelectionne As Integer 'Objet saisie avec le clique gauche.
'Public RessourceSelectionne As Boolean 'Vrai si le personnage a saisi des ressources pour les donner ou les déposer.
'Private RessourcesSelectionnes() As Double
'Public ArgentSelectionne As Double
'Public EtatFabricationObjet As Double
'Private ObjetsEquipes() As ClsJeuObjet
'Private ObjetsInventaires() As ClsJeuObjet
'Public NumeroObjetUtilise As Integer 'Mémorise l'indice de l'objet utilisé.
'Private GenerationRessources As Boolean
'Private GainArgent As Single
'Private GainRessource() As Single
'
''Bestiaire.
'Private Bestiaire() As Long
'
''Variable temporaire.
'Public Vitalite As Double


''Stats de la partie
Public Property Get FragsPersos() As Long
    FragsPersos = Persos(Numero).FragsPersos
End Property
Public Property Get FragsBatiments() As Long
    FragsBatiments = Persos(Numero).FragsBatiments
End Property
Public Property Get FragsChateaux() As Long
    FragsChateaux = Persos(Numero).FragsChateaux
End Property
Public Property Get Morts() As Long
    Morts = Persos(Numero).Morts
End Property

'Maison.
'Position actuelle de la maison.
Public Property Get MaisonPositionX() As Long
    MaisonPositionX = Maisons(Numero).PositionX
End Property
Public Property Get MaisonPositionY() As Long
    MaisonPositionY = Maisons(Numero).PositionY
End Property
Public Property Get MaisonLargeur() As Integer
    MaisonLargeur = Maisons(Numero).Largeur
End Property
Public Property Get MaisonHauteur() As Integer
    MaisonHauteur = Maisons(Numero).Hauteur
End Property
Public Property Get MaisonEntreePositionX() As Integer
    MaisonEntreePositionX = Maisons(Numero).EntreePositionX
End Property
Public Property Get MaisonEntreePositionY() As Integer
    MaisonEntreePositionY = Maisons(Numero).EntreePositionY
End Property
Public Property Get MaisonConstruit() As Boolean
    MaisonConstruit = Maisons(Numero).Construit
End Property
Public Property Get MaisonEnConstruction() As Boolean
    MaisonEnConstruction = Maisons(Numero).Visible And Not (Maisons(Numero).Construit And Maisons(Numero).Definir_Vie = Maisons(Numero).MaxVie)
End Property
Public Property Get MaisonVie() As Double
    MaisonVie = Maisons(Numero).Vie
End Property
Public Property Get MaisonMaxVie() As Long
    MaisonMaxVie = Maisons(Numero).MaxVie
End Property
Public Property Get MaisonArgent() As Double
    MaisonArgent = Maisons(Numero).Argent
End Property
'Private Stock As Double
'Public MaxStock As Double
'Private StockObjets() As Double
'Public Magasin As Boolean
'Public Fabrique As Boolean
'Public Service As Integer
Public Property Get MaisonType() As Integer
    MaisonType = Maisons(Numero).TypeBatiment + 1
End Property
Public Function MaisonBatir(ByVal X As Long, _
                            ByVal Y As Long, _
                            ByVal TypeMaison As Integer) As Boolean
    If Persos(Numero).Vivant And _
       Obstacle(Numero, X, Y, _
                Parametres.Batiment_Largeur(TypeMaison), _
                Parametres.Batiment_Hauteur(TypeMaison)) = 0 And _
       Persos(Numero).PeutConstruire_Maison(TypeMaison, Parametres) Then
        Maisons(Numero).Creer X, Y, TypeMaison, Numero, Parametres
        'Persos(Numero).Rentrer_Maison Maisons(Numero)
    End If
End Function


'Public Sub Demolir(ByRef Perso As ClsJeuPerso, _
'                   ByRef Parametres As ClsJeuParam)
'    Dim i As Long
'    For i = 0 To Parametres.NombreRessources - 1
'        If Parametres.Batiment_Prix(TypeBatiment, i) > 0 Then
'            Perso.Definir_Ressources(i) = Perso.Definir_Ressources(i) + _
'                                          Parametres.Batiment_Prix(TypeBatiment, i) * _
'                                          CoefRecyclage * _
'                                          Vie / MaxVie
'        End If
'    Next i
'    Perso_Demolir_Batiment Perso, Me
'    Definir_Vie = 0
'    If Perso.DansSaMaison Then
'        Perso.DansSaMaison = False
'        Perso.Action = 0
'    End If
'    Perso.Calculer_Apparence Me, Parametres
'End Sub
'
'Public Sub Transfert_Argent_Prendre(ByVal Quantite As Double, _
'                                    ByRef Perso As ClsJeuPerso, _
'                                    Optional TransfertMax As Boolean)
'    If TransfertMax Or Quantite > Argent Then
'        Quantite = Argent
'    End If
'    Perso.Definir_Argent = Perso.Definir_Argent + Quantite
'    Argent = Argent - Quantite
'End Sub
'
'Public Sub Transfert_Argent_Deposer(ByVal Quantite As Double, _
'                                    ByRef Perso As ClsJeuPerso, _
'                                    Optional TransfertMax As Boolean)
'    If TransfertMax Or Quantite > Perso.Definir_Argent Then
'        Quantite = Perso.Definir_Argent
'    End If
'    Perso.Definir_Argent = Perso.Definir_Argent - Quantite
'    Argent = Argent + Quantite
'End Sub
'
'Public Sub Transfert_Stock_Prendre(ByVal Quantite As Double, _
'                                   ByRef Perso As ClsJeuPerso, _
'                                   ByRef Parametres As ClsJeuParam, _
'                                   Optional TransfertMax As Boolean)
'    If TransfertMax Or Quantite > Stock Then
'        Quantite = Stock
'    End If
'    'Ne prend pas plus de ressources que le joueur peut prendre.
'    If Quantite > Perso.Definir_MaxRessources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) - Int(Perso.Definir_Ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment))) Then
'       Quantite = Perso.Definir_MaxRessources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) - Int(Perso.Definir_Ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment)))
'    End If
'    Quantite = Int(Quantite)
'    'On transfert.
'    Perso.Definir_Ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) = Perso.Definir_Ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) + Quantite
'    Stock = Stock - Quantite
'End Sub
'
'Public Sub Transfert_Stock_Deposer(ByVal Quantite As Double, _
'                                   ByRef Perso As ClsJeuPerso, _
'                                   ByRef Parametres As ClsJeuParam, _
'                                   Optional TransfertMax As Boolean)
'    If TransfertMax Or Quantite > Perso.Definir_Ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) Then
'        Quantite = Perso.Definir_Ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment))
'    End If
'    'On ne dépose pas plus de ressources que le batiment peut contenir.
'    If Quantite > MaxStock - Stock Then
'       Quantite = MaxStock - Stock
'    End If
'    Quantite = Int(Quantite)
'    'On transfert.
'    Perso.Definir_Ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) = Perso.Definir_Ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) - Quantite
'    Stock = Stock + Quantite
'End Sub
'
''Public Sub Transfert_Objet_Prendre(ByVal IndiceObjet As Integer, _
'                                   ByRef Perso As ClsJeuPerso, _
'                                   ByRef Parametres As ClsJeuParam)
'    'On met l'objet dans l'inventaire.
'
''    Stock = Stock - Quantite
''End Sub
'
'
'Public Sub Acheter(ByVal Quantite As Double, _
'                   ByRef Perso As ClsJeuPerso, _
'                   ByRef Parametres As ClsJeuParam, _
'                   Optional AcheterMax As Boolean)
'    'Achat de ressources.
'    Dim Prix As Integer
'    If AcheterMax Or _
'       Perso.Definir_Argent < Quantite * Perso.Prix_Marchandee(Parametres.Ressources_PrixVente(Parametres.Batiment_RessourceFabrique(TypeBatiment)) * Parametres.CoefficientAchat, Parametres) Then
'        Quantite = Perso.Definir_Argent / Perso.Prix_Marchandee(Parametres.Ressources_PrixVente(Parametres.Batiment_RessourceFabrique(TypeBatiment)) * Parametres.CoefficientAchat, Parametres)
'    End If
'    If Quantite > Stock Then
'        Quantite = Stock
'    End If
'    If Quantite > Perso.Definir_MaxRessources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) - Int(Perso.Definir_Ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment))) Then
'        Quantite = Perso.Definir_MaxRessources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) - Int(Perso.Definir_Ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment)))
'    End If
'    Quantite = Int(Quantite)
'    Perso.Definir_Ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) = Perso.Definir_Ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) + Quantite
'    Stock = Stock - Quantite
'    Prix = Quantite * Perso.Prix_Marchandee(Parametres.Ressources_PrixVente(Parametres.Batiment_RessourceFabrique(TypeBatiment)) * Parametres.CoefficientAchat, Parametres)
'    Perso.Definir_Argent = Perso.Definir_Argent - Prix
'    Argent = Argent + Prix
'    'Compétence marchandage.
'    Perso.Gestion_Niveau_Comp_Speciale 3, Prix, Parametres
'End Sub
'
'Public Function Service_Utilisable(ByVal Perso As ClsJeuPerso, _
'                                   ByVal Parametres As ClsJeuParam) As Boolean
'    'Renvoie si le perso peut utiliser le service.
'    'Le temples ne ressuscitent pas les personnes vivantes.
'    'Le joueur doit avoir assez d'argent ou être propriétaire de la maison.
'    Service_Utilisable = Not (Service = 1 And Perso.Vivant) And _
'                         (Perso.Definir_Argent >= Perso.Prix_Marchandee(Parametres.Service_Prix(Service), Parametres) Or _
'                         Perso.Numero = Numero) And _
'                         Stock >= 1
'    'If Perso.Numero > 0 Then
'    '    Debug.Print Service & ", " & Numero & ", " & Int(Stock) & ", " & Int(Me.Definir_Stock)
'    'End If
'End Function
'
'Public Sub Service_Utiliser(ByRef Perso As ClsJeuPerso, _
'                            ByVal Parametres As ClsJeuParam, _
'                            Optional ByVal Score As Double)
'    Dim Prix As Integer
'    Select Case Service
'    Case 1:
'        Perso.Vivant = True
'        Perso.Vie = Parametres.Service_Coefficient(Service) * Perso.Definir_MaxVie
'        Stock = Stock - 1
'    Case 2:
'        Argent = Argent + Score
'        Perso.Gestion_Niveau_Prime Score, Parametres
'    End Select
'    'If Stock < 1 Then
'    '    Debug.Print Service_Utilisable(Perso, Parametres)
'    'End If
'    If Not (Perso.Numero = Numero) Then
'        'On paye le service utilisé.
'        Prix = Perso.Prix_Marchandee(Parametres.Service_Prix(Service), Parametres)
'        Perso.Definir_Argent = Perso.Definir_Argent - Prix
'        Argent = Argent + Prix
'        'Compétence marchandage.
'        Perso.Gestion_Niveau_Comp_Speciale 3, Prix, Parametres
'    End If
'    If Parametres.Service_Effet(Service) >= 0 Then
'        EffetsMaisons(Numero).Activer_Effet Parametres.Service_Effet(Service), PositionX, PositionY, Largeur, Hauteur, Parametres
'    End If
'End Sub
'
'
'Public Property Get Visible() As Boolean
'    Visible = Vie > 0
'End Property
'
'Public Property Get Definir_Vie() As Double
'    If Vie + RecuperationVie > MaxVie Then
'        Definir_Vie = MaxVie
'    Else
'        Definir_Vie = Vie + RecuperationVie
'    End If
'End Property
'
'Public Property Let Definir_Vie(ByVal Valeur As Double)
'    Vie = Valeur
'    If Vie >= MaxVie Then
'        Vie = MaxVie
'        EnConstruction = False
'        Construit = True
'    ElseIf Vie <= 0 Then
'        Construit = False
'        EnConstruction = False
'        RecuperationVie = 0 'Les chateaux ne se régénèrent plus.
'    Else
'        EnConstruction = True
'    End If
'End Property
'
'Public Property Get Etat(ByVal Parametres As ClsJeuParam) As Integer
'    If Construit Then
'        If (Magasin Or Fabrique Or (Service And MaxStock > 0)) And Stock < 1 Then
'            'Magasin Vide.
'            Etat = Parametres.Batiment_NombreEtapeFabrication(TypeBatiment) - 1
'        Else
'            Etat = Parametres.Batiment_NombreEtapeFabrication(TypeBatiment)
'        End If
'    Else
'        Etat = Int(Vie / MaxVie * Parametres.Batiment_NombreEtapeFabrication(TypeBatiment))
'    End If
'End Property
'
'Public Sub Actualiser_NumeroAnimation()
'    NumeroAnimation = NumeroAnimation + 1
'    If NumeroAnimation > Parametres.Batiment_NombreEtapeAnimation(TypeBatiment) - 1 Then
'        NumeroAnimation = 0
'    End If
'End Sub
'
'Public Property Get Definir_NumeroAnimation(ByVal Parametres As ClsJeuParam) As Integer
'    Definir_NumeroAnimation = NumeroAnimation
'End Property
'
'Public Property Get Definir_Stock() As Double
'    Definir_Stock = Stock
'End Property
'Public Property Let Definir_Stock(ByVal Valeur As Double)
'    Stock = Valeur
'    If Stock > MaxStock Then
'        Stock = MaxStock
'    End If
'End Property
'
'Public Property Get Stock_Objets(ByVal Indice As Integer) As Double
'    Stock_Objets = StockObjets(Indice)
'End Property
'Public Property Let Stock_Objets(ByVal Indice As Integer, ByVal Valeur As Double)
'    Stock = Stock + Valeur - StockObjets(Indice)
'    StockObjets(Indice) = Valeur
'End Property
'
'Public Property Get Definir_Nombre_Stock_Objet() As Integer
'    Definir_Nombre_Stock_Objet = UBound(StockObjets())
'End Property
'Public Property Let Definir_Nombre_Stock_Objet(ByVal Valeur As Integer)
'    ReDim StockObjets(Valeur)
'End Property


'' ************************************************************************************
'' *                          Fonctions du personnage                                 *
'' ************************************************************************************

Public Function Aller_A(ByVal X As Long, ByVal Y As Long) As Boolean
    If X >= 0 And _
       X < Monde.Largeur * AffTerrain.Largeur - Persos(Numero).Largeur And _
       Y >= 0 And _
       Y < Monde.Hauteur * AffTerrain.Hauteur - Persos(Numero).Hauteur Then
        Persos(Numero).Aller_A X, Y
        Aller_A = True
    End If
End Function

Public Function Arreter() As Boolean
    Persos(Numero).Arreter
    Arreter = True
End Function

Public Function Rentrer_Maison() As Boolean
    If Maisons(Numero).Visible Then
        Persos(Numero).Rentrer_Maison Maisons(Numero)
        Rentrer_Maison = True
    Else
        Rentrer_Maison = False
    End If
End Function

'Public Sub Aller_Maison(ByVal Maison As ClsJeuBatiment)
'    If Maison.Visible Then
'End Sub

Public Function Rentrer_Chateau() As Boolean
    Persos(Numero).Rentrer_Chateau Chateaux(Persos(Numero).NumeroFief)
    Rentrer_Chateau = Action = 202
End Function

'Public Sub Aller_Chateau(ByVal Chateau As ClsJeuBatiment, _
'                         ByVal Indice As Long)
'    If Chateau.Visible Then
'End Sub
'
'Public Sub Aller_Ressource(ByVal Ressource As ClsJeuRessources, _
'                           ByVal Indice As Long, _
'                           ByVal LargeurRessource As Long, _
'                           ByVal HauteurRessource As Long)
'    DirectionX = Ressource.PositionX + LargeurRessource / 2 - Largeur / 2
'    DirectionY = Ressource.PositionY + LargeurRessource / 2 - Hauteur / 2
'    IndiceRessource = Indice
'    Action = 3
'End Sub
'
'Public Sub Aller_Tresor(ByVal Tresor As ClsJeuTresors, _
'                        ByVal Indice As Long, _
'                        ByVal LargeurTresor As Long, _
'                        ByVal HauteurTresor As Long)
'    DirectionX = Tresor.X + LargeurTresor / 2 - Largeur / 2
'    DirectionY = Tresor.Y + LargeurTresor / 2 - Hauteur / 2
'    IndiceTresor = Indice
'    Action = 204
'    Visible = True
'End Sub
'
'Public Sub Aller_Donner(ByVal Indice As Long, ByVal TypeObjet As Integer)
'    IndicePerso = Indice
'    IndiceObjet = TypeObjet
'    Action = 205
'    Visible = True
'End Sub
Public Function Donner_Ressource(ByVal Indice As Integer, _
                                 ByVal Quantite As Double, _
                                 ByVal PersoCible As Long, _
                                 Optional Max As Boolean)
    If Persos(PersoCible).Vivant And Persos(PersoCible).Visible Then
        Persos(Numero).Ressources_Prendre Indice - 1, Quantite, Max
        Persos(Numero).Aller_Donner PersoCible
    End If
End Function
'Public Sub Attaquer_Perso(ByVal Indice As Long)
'    Action = 300
'    IndicePerso = Indice
'    Visible = True
'    ForcerAttaque = True
'End Sub
'
'Public Sub Attaquer_Maison(ByVal Indice As Long)
'    Action = 301
'    IndiceMaison = Indice
'    Visible = True
'    ForcerAttaque = True
'End Sub
'
'Public Sub Attaquer_Chateau(ByVal Indice As Long)
'    Action = 302
'    IndiceChateau = Indice
'    Visible = True
'    ForcerAttaque = True
'End Sub
'
'' ************************************************************************************
'' *                              Fonctions de récuparation                           *
'' ************************************************************************************
'
'Public Sub Recuperer_AutoVie(ByVal Parametres As ClsJeuParam)
'    If Definir_Vie < Definir_MaxVie * 1 Then
'        Recuperer_Vie Parametres
'    End If
'End Sub
'Public Sub Recuperer_Vie(ByVal Parametres As ClsJeuParam)
'    'Mange tout ce qu'il peut pour récuperer de la vie.
'    Dim X As Integer
'    For X = 0 To Parametres.NombreRessources - 1
'        If Parametres.Ressources_Mangeable(X) And _
'           Parametres.Ressources_GainVie(X) > 0 And _
'           Definir_Ressources(X) >= 1 Then
'            Do
'                Manger_Ressources X, Parametres
'            Loop Until Definir_Vie >= Definir_MaxVie - Parametres.Ressources_GainVie(X) * Manger_Ressources_Coef(0) Or _
'                       Definir_Ressources(X) < 1
'            X = Parametres.NombreRessources
'        End If
'    Next X
'End Sub
'Public Sub Recuperer_AutoEnergie(ByVal Parametres As ClsJeuParam)
'    If Definir_Energie < Definir_MaxEnergie * 0.5 Then
'        Recuperer_Energie Parametres
'    End If
'End Sub
'Public Sub Recuperer_Energie(ByVal Parametres As ClsJeuParam)
'    'Mange tout ce qu'il peut pour récuperer de l'énergie.
'    Dim X As Integer
'    For X = 0 To Parametres.NombreRessources - 1
'        If Parametres.Ressources_Mangeable(X) And _
'           Parametres.Ressources_GainEnergie(X) > 0 And _
'           Definir_Ressources(X) >= 1 Then
'            Do
'                Manger_Ressources X, Parametres
'            Loop Until Definir_Energie >= Definir_MaxEnergie - Parametres.Ressources_GainEnergie(X) * Manger_Ressources_Coef(1) Or _
'                       Definir_Ressources(X) < 1
'            X = Parametres.NombreRessources
'        End If
'    Next X
'End Sub
'Public Sub Recuperer_AutoMagie(ByVal Parametres As ClsJeuParam)
'    If Definir_Magie < Definir_MaxMagie * 0.5 Then
'        Recuperer_Magie Parametres
'    End If
'End Sub
'Public Sub Recuperer_Magie(ByVal Parametres As ClsJeuParam)
'    'Mange tout ce qu'il peut pour récuperer de la magie.
'    Dim X As Integer
'    For X = 0 To Parametres.NombreRessources - 1
'        If Parametres.Ressources_Mangeable(X) And _
'           Parametres.Ressources_GainMagie(X) > 0 And _
'           Definir_Ressources(X) >= 1 Then
'            Do
'                Me.Manger_Ressources X, Parametres
'            Loop Until Definir_Magie >= Definir_MaxMagie - Parametres.Ressources_GainMagie(X) * Manger_Ressources_Coef(2) Or _
'                       Definir_Ressources(X) < 1
'            X = Parametres.NombreRessources
'        End If
'    Next X
'End Sub
'Public Sub Recuperer_AutoMoral(ByVal Parametres As ClsJeuParam)
'    If Definir_Moral < Definir_MaxMoral * 0.5 Then
'        Recuperer_Moral Parametres
'    End If
'End Sub
'Public Sub Recuperer_Moral(ByVal Parametres As ClsJeuParam)
'    'Mange tout ce qu'il peut pour récuperer de la magie.
'    Dim X As Integer
'    For X = 0 To Parametres.NombreRessources - 1
'        If Parametres.Ressources_Mangeable(X) And _
'           Parametres.Ressources_GainMoral(X) > 0 And _
'           Definir_Ressources(X) >= 1 Then
'            Do
'                Manger_Ressources X, Parametres
'            Loop Until Definir_Moral >= Definir_MaxMoral - Parametres.Ressources_GainMoral(X) * Manger_Ressources_Coef(3) Or _
'                       Definir_Ressources(X) < 1
'            X = Parametres.NombreRessources
'        End If
'    Next X
'End Sub
'
'' ************************************************************************************
'' *                               Fonction principale                                *
'' ************************************************************************************
'
'Public Sub Deplacement(ByRef RessourceCible As ClsJeuRessources, _
'                       ByRef PersoCible As ClsJeuPerso, _
'                       ByRef MaisonCible As ClsJeuBatiment, _
'                       ByRef ChateauCible As ClsJeuBatiment, _
'                       ByRef Fief As ClsJeuFief, _
'                       ByRef NoCommentaire As Integer, _
'                       ByRef Commentaires As ClsJeuCommentaires, _
'                       ByVal Parametres As ClsJeuParam)
'    Calculer_Vitalite
'    If EnCombat Or Action = 205 Then
'        Select Case Action
'        Case 205:
'            If PersoCible.Visible And PersoCible.Vivant Then
'                DirectionX = PersoCible.PositionX + PersoCible.Largeur / 2 - Largeur / 2
'                DirectionY = PersoCible.PositionY + PersoCible.Hauteur / 2 - Hauteur / 2
'            Else
'                Arreter
'            End If
'        Case 300:
'            If PersoCible.Visible And PersoCible.Vivant Then
'                DirectionX = PersoCible.PositionX + PersoCible.Largeur / 2 - Largeur / 2
'                DirectionY = PersoCible.PositionY + PersoCible.Hauteur / 2 - Hauteur / 2
'            ElseIf PersoCible.Vivant Then
'                If IA Then 'Les IA poursuivent le réfugié.
'                    If PersoCible.DansSaMaison Then
'                        Attaquer_Maison PersoCible.Numero
'                    ElseIf PersoCible.DansUneMaison Then
'                        Attaquer_Maison PersoCible.IndiceMaison
'                    ElseIf PersoCible.DansSonChateau Then
'                        Attaquer_Chateau PersoCible.NumeroFief
'                    ElseIf PersoCible.DansUnChateau Then
'                        Attaquer_Chateau PersoCible.IndiceChateau
'                    End If
'                Else
'                    Arreter
'                End If
'                Definir_Commentaires = Commentaires.Message(15, Me)
'            Else
'                Arreter
'            End If
'        Case 301:
'            If MaisonCible.Visible Then
'                DirectionX = MaisonCible.PositionX + MaisonCible.Largeur / 2 - Largeur / 2
'                DirectionY = MaisonCible.PositionY + MaisonCible.Hauteur / 2 - Hauteur / 2
'            Else
'                Arreter
'            End If
'        Case 302:
'            If Chateaux(IndiceChateau).Visible Then
'                DirectionX = ChateauCible.PositionX + ChateauCible.Largeur / 2 - Largeur / 2
'                DirectionY = ChateauCible.PositionY + ChateauCible.Hauteur / 2 - Hauteur / 2
'            Else
'                Arreter
'            End If
'        End Select
'    End If
'    If Action = 205 And _
'        PersoCible.PositionX + PersoCible.Largeur >= PositionX And _
'        PersoCible.PositionY + PersoCible.Hauteur >= PositionY And _
'        PersoCible.PositionX <= PositionX + Largeur And _
'        PersoCible.PositionY <= PositionY + Hauteur _
'       Or _
'        Action = 300 And _
'        PersoCible.PositionX + PersoCible.Largeur >= PositionX - Definir_PorteeAttaque And _
'        PersoCible.PositionY + PersoCible.Hauteur >= PositionY - Definir_PorteeAttaque And _
'        PersoCible.PositionX <= PositionX + Largeur + Definir_PorteeAttaque And _
'        PersoCible.PositionY <= PositionY + Hauteur + Definir_PorteeAttaque _
'       Or _
'       Action = 301 And _
'        MaisonCible.PositionX + MaisonCible.Largeur >= PositionX - Definir_PorteeAttaque And _
'        MaisonCible.PositionY + MaisonCible.Hauteur >= PositionY - Definir_PorteeAttaque And _
'        MaisonCible.PositionX <= PositionX + Largeur + Definir_PorteeAttaque And _
'        MaisonCible.PositionY <= PositionY + Hauteur + Definir_PorteeAttaque _
'       Or _
'       Action = 302 And _
'        ChateauCible.PositionX + ChateauCible.Largeur >= PositionX - Definir_PorteeAttaque And _
'        ChateauCible.PositionY + ChateauCible.Hauteur >= PositionY - Definir_PorteeAttaque And _
'        ChateauCible.PositionX <= PositionX + Largeur + Definir_PorteeAttaque And _
'        ChateauCible.PositionY <= PositionY + Hauteur + Definir_PorteeAttaque _
'       Then
'        If Action = 205 Then
'            'Donne un objet.
'            If ObjetSelectionne >= 0 Then
'                PersoCible.Objet_Ajouter_Inventaire IndiceObjet, Parametres, Commentaires, True
'                ObjetSelectionne = -1
'            End If
'            'Donne des ressources.
'            Ressources_Donner PersoCible
'            PersoCible.Definir_Commentaires = Commentaires.Message(46, Me)
'            Arreter
'        Else
'            If Attaque = Definir_MaxAttaque Or _
'               ForcerAttaque Or _
'               (Action = 300 And _
'                IIf(ObjetsEquipes(0).PerteSiAttaque Or ObjetsEquipes(1).PerteSiAttaque, Attaque >= PersoCible.Vie / CoefDegatPerso + PersoCible.Defense, Attaque > PersoCible.Vie + PersoCible.Defense + 1)) Then
'                'S'il peut achever la cible, il l'achève.
'                If Vie > PerteVie Then
'                    If Energie > PerteEnergie Then
'                        If Magie > PerteMagie Then
'                            If Moral > PerteMoral Then
'                                If Definir_Argent >= PerteArgent Then
'                                    If Ressources(PerteRessourceIndice) >= PerteRessourceQuantite Then
'                                        'Le personnage attaque.
'                                        If BonusSurfaceAttaque > 0 Then
'                                            Select Case Action
'                                            Case 300: Perso_Attaquer_Surface Me, PersoCible.PositionX, PersoCible.PositionY, BonusSurfaceAttaque, ControleJoueur
'                                            Case 301: Perso_Attaquer_Surface Me, MaisonCible.PositionX, MaisonCible.PositionY, BonusSurfaceAttaque, ControleJoueur
'                                            Case 302: Perso_Attaquer_Surface Me, ChateauCible.PositionX, ChateauCible.PositionY, BonusSurfaceAttaque, ControleJoueur, IndiceChateau
'                                            End Select
'                                        Else
'                                            Select Case Action
'                                            Case 300: Perso_Attaquer_Perso Me, PersoCible, ControleJoueur
'                                            Case 301: Perso_Attaquer_Batiment Me, MaisonCible, Persos(IndiceMaison), ControleJoueur
'                                            Case 302: Perso_Attaquer_Batiment Me, ChateauCible, Persos(IndicePerso), ControleJoueur, IndiceChateau
'                                            End Select
'                                        End If
'                                        If Objet_Equipes_Actif(0) Then
'                                            Definir_Vie = Vie - ObjetsEquipes(0).PerteVie - ObjetsEquipes(1).PerteVie
'                                            Definir_Energie = Energie - ObjetsEquipes(0).PerteEnergie - ObjetsEquipes(1).PerteEnergie
'                                            Definir_Magie = Magie - ObjetsEquipes(0).PerteMagie - ObjetsEquipes(1).PerteMagie
'                                            Definir_Moral = Moral - ObjetsEquipes(0).PerteMoral - ObjetsEquipes(1).PerteMoral
'                                            'Les objets a utilisation unique disparaissent.
'                                            If ObjetsEquipes(0).UtilisationUnique Then
'                                                ObjetsEquipes(0).Decharger
'                                            End If
'                                        End If
'                                        ForcerAttaque = False
'                                        Definir_Argent = Definir_Argent - PerteArgent
'                                        Definir_Ressources(PerteRessourceIndice) = Definir_Ressources(PerteRessourceIndice) - PerteRessourceQuantite
'                                        'Les kamikazes meurent après avoir attaqué.
'                                        If Kamikaze Then Vivant = False
'                                    Else
'                                        'N'a plus de ressources pour attaquer.
'                                        Definir_Commentaires = Commentaires.Message(34, Me)
'                                        Objet_Ranger_Armes Commentaires, Parametres
'                                    End If
'                                Else
'                                    'N'a plus d'argent pour attaquer.
'                                    Definir_Commentaires = Commentaires.Message(45, Me)
'                                    Objet_Ranger_Armes Commentaires, Parametres
'                                End If
'                            Else
'                                'N'a plus de point de moral.
'                                Definir_Commentaires = Commentaires.Message(20, Me)
'                                Objet_Ranger_Armes Commentaires, Parametres
'                            End If
'                        Else
'                            'N'a plus de point de magie.
'                            Definir_Commentaires = Commentaires.Message(19, Me)
'                            Objet_Ranger_Armes Commentaires, Parametres
'                        End If
'                    Else
'                        'N'a plus de point d'énergie.
'                        Definir_Commentaires = Commentaires.Message(18, Me)
'                        Objet_Ranger_Armes Commentaires, Parametres
'                    End If
'                Else
'                    'N'a plus de point de vie.
'                    Definir_Commentaires = Commentaires.Message(17, Me)
'                    Objet_Ranger_Armes Commentaires, Parametres
'                End If
'            End If
'        End If
'        If Not Volant Then
'            Direction = 0
'        End If
'    ElseIf DureeImmobilisation <= 0 Then
'        Calculer_Longueur_Pas
'        'Cette fonction modifie la position du personnage.
'        If DirectionY < PositionY - Longueur_Pas / 2 Then
'            'Déplacement en haut.
'            If DirectionX < PositionX - Longueur_Pas / 2 Then
'                'Vers la gauche.
'                If Abs(DirectionY - PositionY) > Abs(DirectionX - PositionX) Then
'                    'Déplacement en haut à gauche. (+vers le haut)
'                    PositionX = PositionX - Longueur_PasChasse
'                    PositionY = PositionY - Longueur_PasNormal
'                    Direction = 1
'                Else
'                    'Déplacement en haut à gauche. (+ vers la gauche)
'                    PositionX = PositionX - Longueur_PasNormal
'                    PositionY = PositionY - Longueur_PasChasse
'                    Direction = 4
'                End If
'            ElseIf DirectionX > PositionX + Longueur_Pas / 2 Then
'                'Vers la droite.
'                If Abs(DirectionY - PositionY) > Abs(DirectionX - PositionX) Then
'                    'Déplacement en haut à droite. (+ vers le haut)
'                    PositionX = PositionX + Longueur_PasChasse
'                    PositionY = PositionY - Longueur_PasNormal
'                    Direction = 1
'                Else
'                    'Déplacement en haut à droite. (+ vers la droite)
'                    PositionX = PositionX + Longueur_PasNormal
'                    PositionY = PositionY - Longueur_PasChasse
'                    Direction = 2
'                End If
'            Else
'                PositionY = PositionY - Longueur_Pas
'                Direction = 1
'            End If
'            Visible = True
'            EnDeplacement = True
'        ElseIf DirectionY > PositionY + Longueur_Pas / 2 Then
'            'Déplacement en bas.
'            If DirectionX < PositionX - Longueur_Pas / 2 Then
'                'Vers la gauche.
'                If Abs(DirectionY - PositionY) > Abs(DirectionX - PositionX) Then
'                    'Déplacement en bas à gauche. (+vers le haut)
'                    PositionX = PositionX - Longueur_PasChasse
'                    PositionY = PositionY + Longueur_PasNormal
'                    Direction = 3
'                Else
'                    'Déplacement en haut à gauche. (+ vers la gauche)
'                    PositionX = PositionX - Longueur_PasNormal
'                    PositionY = PositionY + Longueur_PasChasse
'                    Direction = 4
'                End If
'            ElseIf DirectionX > PositionX + Longueur_Pas / 2 Then
'                'Vers la droite.
'                If Abs(DirectionY - PositionY) > Abs(DirectionX - PositionX) Then
'                    'Déplacement en bas à droite. (+ vers le bas)
'                    PositionX = PositionX + Longueur_PasChasse
'                    PositionY = PositionY + Longueur_PasNormal
'                    Direction = 3
'                Else
'                    'Déplacement en bas à droite. (+ vers la droite)
'                    PositionX = PositionX + Longueur_PasNormal
'                    PositionY = PositionY + Longueur_PasChasse
'                    Direction = 2
'                End If
'            Else
'                PositionY = PositionY + Longueur_Pas
'                Direction = 3
'            End If
'            Visible = True
'            EnDeplacement = True
'        ElseIf DirectionX < PositionX - Longueur_Pas / 2 Then
'            'Déplacement à gauche.
'            PositionX = PositionX - Longueur_Pas
'            Direction = 4
'            Visible = True
'            EnDeplacement = True
'        ElseIf DirectionX > PositionX + Me.Longueur_Pas / 2 Then
'            'Déplacement à droite.
'            PositionX = PositionX + Longueur_Pas
'            Direction = 2
'            Visible = True
'            EnDeplacement = True
'        Else
'            If Volant Then
'                Direction = 3
'                'EnDeplacement = True
'                'EnDeplacement = False
'            Else
'                'EnDeplacement = False
'                Direction = 0
'            End If
'            EnDeplacement = False
'            DirectionX = PositionX
'            DirectionY = PositionY
'            If Action Mod 2 = 1 And _
'               Action < 100 Then
'                Action = Action - 1
'            End If
'            If Vivant Then
'                Select Case Action
'                    'On ne bouge pas.
'                    Case 0: RessourceCourante = -1
'                            'MatierePremiere = 0
'                            If Not Chateaux(NumeroFief).Visible Then DansSonChateau = False
'                            If Not Chateaux(IndiceChateau).Visible Then DansUnChateau = False
'                            'If Not Maisons(Numero).Visible Then DansSaMaison = False
'                            If Not Maisons(IndiceMaison).Visible Or Not MaisonCible.Construit Then DansUneMaison = False
'                            'Maison.EnConstruction = False
'                            'On se repose dehors.
'                            Definir_Vie = Vie + VitesseRegenerationVieEnAttente * Definir_MaxVie
'                            Definir_Energie = Energie + VitesseRegenerationEnergieEnAttente * Definir_MaxEnergie
'                            Definir_Magie = Magie + VitesseRegenerationMagieEnAttente * Definir_MaxMagie
'                            Definir_Moral = Moral + VitesseRegenerationMoralEnAttente * Definir_MaxMoral
'
'                    'On récolte des ressources à l'extérieur.
'                    Case 2:
'                        RessourceCourante = RessourceCible.TypeRessource
'                        Recolter_Ressources RessourceCourante, Maisons(Numero), NoCommentaire, Parametres
'                        'Action = Parametres.Ressources_NomCompetence(RessourceCible.TypeRessource)
'                        DansUneMaison = False
'                    'On construit une maison.
'                    Case 6:
'                        If MaisonCible.Visible Then
'                            RessourceCourante = -1
'                            MaisonCible.EnConstruction = True
'                            Batir_Maison Fief, Chateaux(NumeroFief), MaisonCible, Parametres
'                        Else
'                            Arreter
'                        End If
'                    'Se repose.
'                    Case 100:
'                        If MaisonCible.Visible Then
'                            DansSaMaison = True
'                            RessourceCourante = -1
'                        Else
'                            DansSaMaison = False
'                        End If
'                        'On se repose à la maison.
'                        Definir_Vie = Vie + VitesseRegenerationVieEnRepos * Definir_MaxVie
'                        Definir_Energie = Energie + VitesseRegenerationEnergieEnRepos * Definir_MaxEnergie
'                        Definir_Magie = Magie + VitesseRegenerationMagieEnRepos * Definir_MaxMagie
'                        Definir_Moral = Moral + VitesseRegenerationMoralEnRepos * Definir_MaxMoral
'                    'Travail à la maison.
'                    Case 101, 102, 600:
'                        If MaisonCible.Visible Then
'                            DansSaMaison = True
'                            RessourceCourante = Parametres.Batiment_RessourceFabrique(MaisonCible.TypeBatiment)
'                            'MatierePremiere = Parametres.Batiment_TypeMatierePremiere(Maison.TypeBatiment)
'                            'Action = Parametres.Batiment_DesignationTravail(MaisonCible.TypeBatiment)
'                            If RessourceCourante >= 0 Then
'                                'Action = Parametres.Ressources_NomCompetence(RessourceCourante)
'                                Action = 102
'                                Recolter_Ressources RessourceCourante, MaisonCible, NoCommentaire, Parametres
'                            ElseIf MaisonCible.Service Then
'                                Preparer_Service MaisonCible, NoCommentaire, Parametres
'                            Else
'                                Arreter
'                            End If
'                        Else
'                            DansSaMaison = False
'                            Arreter
'                        End If
'                    'Va prendre un trésor.
'                    Case 204:
'                        Tresor_Ramasser IndiceTresor, Me
'                        DansSaMaison = False
'                        Arreter
'                        'Action = 0
'                    Case 500:
'                        If MaisonCible.Visible Then
'                            Fabriquer_Objet Parametres, Maisons(Numero), NoCommentaire
'                        Else
'                            DansSaMaison = False
'                            Arreter
'                        End If
'                    'Rentre dans un magasin.
'                    Case 201:
'                        DansUneMaison = MaisonCible.Visible And MaisonCible.Construit
'                        Arreter
'                    'Retourne à son chateau.
'                    Case 202:
'                        DansSonChateau = Chateaux(NumeroFief).Visible
'                        Action = 0
'                    'Entre dans un chateau allié.
'                    Case 203:
'                        DansUnChateau = Chateaux(IndiceChateau).Visible
'                        Action = 0
'                    Case Else: RessourceCourante = -1
'                               'MatierePremiere = 0
'                               'Maison.EnConstruction = False
'                End Select
'            Else
'                'Liste des actions que le joueur peut faire quand il est mort.
'                Select Case Action
'                    Case 0:
'                        If Not Chateaux(NumeroFief).Visible Then DansSonChateau = False
'                        If Not Chateaux(IndiceChateau).Visible Then DansUnChateau = False
'                        'If Not Maisons(Numero).Visible Then DansSaMaison = False
'                        If Not Maisons(IndiceMaison).Visible And Maisons(IndiceMaison).Construit Then DansUneMaison = False
'                    Case 100:
'                        DansSaMaison = MaisonCible.Visible
'                        Arreter
'                    Case 101, 102:
'                        DansSaMaison = MaisonCible.Visible
'                        Arreter
'                    Case 201:
'                        DansUneMaison = MaisonCible.Visible And MaisonCible.Construit
'                        Arreter
'                    'Retourne à son chateau.
'                    Case 202:
'                        DansSonChateau = Chateaux(NumeroFief).Visible
'                        Action = 0
'                End Select
'            End If
'        End If
'        'If Deplacement > 0 Then
'        'If Direction > 0 Then
'        If EnDeplacement Then
'            If Action > 99 Then
'                If Action Mod 2 = 0 And Action < 200 Then
'                    Action = Action + 1
'                End If
'                If Action < 100 Then
'                    Action = 1
'                End If
'            ElseIf Action = 7 And Not MaisonCible.Visible Then
'                Arreter
'            End If
'            MaisonCible.EnConstruction = False
'            'Visible = True
'            'RessourceCourante = -1
'            If Vivant Then
'                Definir_Energie = Energie - VitesseFatigueDeplacement
'                Definir_Moral = Moral - VitesseDemoralisationDeplacement
'            End If
'        End If
'    Else
'        DureeImmobilisation = DureeImmobilisation - 1
'    End If
'
'    'On incrémente le pas pour le prochain affichage.
'    If Direction > 0 Or Volant Then
'        NumeroPas = NumeroPas + 1
'        If NumeroPas > 3 Then NumeroPas = 0
'    End If
'
'    'Direction = Deplacement
'End Sub
'
'
Public Property Get Action() As String
    Action = Persos(Numero).Action
End Property
'    Select Case Action
'        Case 0: Definir_Action = "En attente"
'        Case 1: Definir_Action = "Se déplacer"
'        Case 2: Definir_Action = "Récolter ressources"
'        Case 3: Definir_Action = "Aller à"
'        Case 6: Definir_Action = "Construire"
'        Case 7: Definir_Action = "Aller construire"
'        Case 100: Definir_Action = "En repos"
'        Case 101: Definir_Action = "Rentrer à la maison"
'        Case 102: Definir_Action = "Au travail"
'        Case 201: Definir_Action = "Entrer dans"
'        Case 202: Definir_Action = "Retourner château"
'        Case 203: Definir_Action = "Aller château"
'        Case 204: Definir_Action = "Aller trésor"
'        Case 205: Definir_Action = "Donner"
'        Case 300: Definir_Action = "Attaquer"
'        Case 301: Definir_Action = "Détruire maison"
'        Case 302: Definir_Action = "Détruire château"
'        Case 400: Definir_Action = "Ressusciter"
'        Case 500: Definir_Action = "Fabriquer"
'        Case 600: Definir_Action = "Préparer Service"
'        Case Else: Definir_Action = "Au travail"
'    End Select
'End Property
'
'Public Sub Preparer_Service(ByVal Maison As ClsJeuBatiment, _
'                            ByRef NoCommentaire As Integer, _
'                            ByVal Parametres As ClsJeuParam)
'    Dim MatierePremiere As Integer
'    Dim VitesseCreation As Double
'    Dim i As Integer
'
'    ServiceCourant = Maison.Service
'    MatierePremiere = Parametres.Service_MatierePremiere(ServiceCourant)
'    VitesseCreation = 1 / VitesseAction * Me.Vitalite * _
'                      Definir_Niveau_CompetenceServices(ServiceCourant - 1)
'
'    'On stock le service dans la maison.
'    If Maison.Definir_Stock < Maison.MaxStock Then
'        If MatierePremiere >= 0 Then 'Il faut une autre ressource comme matière première.
'            If Me.Definir_Ressources(MatierePremiere) >= Parametres.Service_VitesseCreation(ServiceCourant) * VitesseCreation _
'                                                         * Parametres.Service_QuantiteMatierePremiere(ServiceCourant) Then
'                'On retranche la matière première.
'                Me.Definir_Ressources(MatierePremiere) = Me.Definir_Ressources(MatierePremiere) - Parametres.Service_VitesseCreation(ServiceCourant) * VitesseCreation _
'                                                         * Parametres.Service_QuantiteMatierePremiere(ServiceCourant)
'                'On augmente la service.
'                Maison.Definir_Stock = Maison.Definir_Stock + Parametres.Service_VitesseCreation(ServiceCourant) * VitesseCreation
'            Else
'                Action = 100
'                NoCommentaire = 1
'            End If
'        Else
'            Maison.Definir_Stock = Maison.Definir_Stock + Parametres.Service_VitesseCreation(ServiceCourant) * VitesseCreation
'        End If
'        Action = 600
'    Else
'        Arreter
'        NoCommentaire = 0
'    End If
'    'Si le personnage a travaillé, il se fatigue.
'    If Action <> 0 And Action <> 100 Then
'        Definir_Energie = Energie - VitesseFatigueTravail
'        Definir_Moral = Moral - VitesseDemoralisationTravail
'        'Il gagne de l'expérience dans sa compétence de service.
'        Gestion_Niveau CaracServices(ServiceCourant - 1), _
'                       Parametres.Service_VitesseCreation(ServiceCourant) * VitesseCreation * Parametres.Service_CoefExp(ServiceCourant), _
'                       ExpServices(ServiceCourant - 1), NivServices(ServiceCourant - 1), ExpNivServices(ServiceCourant - 1), ExpNivPlusServices(ServiceCourant - 1), 7
'        'Debug.Print CaracServices(ServiceCourant - 1) & ", " & ExpServices(ServiceCourant - 1) & ", " & NivServices(ServiceCourant - 1) & ", " & ExpNivServices(ServiceCourant - 1) & ", " & ExpNivPlusServices(ServiceCourant - 1)
'        'Il gagne de l'expérience avec les objets qu'il utilise.
'        If Bonus_CompetenceServices(ServiceCourant - 1) Then
'            For i = 0 To UBound(ObjetsEquipes())
'                If ObjetsEquipes(i).Param_BonusCompetencesServices(ServiceCourant) > 0 Then
'                    Gestion_Niveau CaracObjets(ObjetsEquipes(i).Competence), _
'                                   Parametres.Service_VitesseCreation(ServiceCourant) * VitesseCreation * Parametres.Service_CoefExp(ServiceCourant), _
'                                   ExpObjets(ObjetsEquipes(i).Competence), NivObjets(ObjetsEquipes(i).Competence), ExpNivObjets(ObjetsEquipes(i).Competence), ExpNivPlusObjets(ObjetsEquipes(i).Competence), 7
'                End If
'            Next i
'        End If
'    End If
'End Sub
'
'Public Sub Manger_Ressources(ByVal Index As Long, _
'                             ByVal Parametres As ClsJeuParam)
'    'Le personnage mange une unité de ressource et récupère ce que la ressource lui redonne.
'    If Ressources(Index) >= 1 Then
'        Ressources(Index) = Ressources(Index) - 1
'        RecuperationVie = RecuperationVie + Parametres.Ressources_GainVie(Index) * Manger_Ressources_Coef(0)
'        RecuperationEnergie = RecuperationEnergie + Parametres.Ressources_GainEnergie(Index) * Manger_Ressources_Coef(1)
'        RecuperationMagie = RecuperationMagie + Parametres.Ressources_GainMagie(Index) * Manger_Ressources_Coef(2)
'        RecuperationMoral = RecuperationMoral + Parametres.Ressources_GainMoral(Index) * Manger_Ressources_Coef(3)
'    End If
'End Sub
'
'Public Function Manger_Ressources_Coef(ByVal NoCarac As Integer) As Double
'    Select Case NoCarac
'    Case 0:
'        Manger_Ressources_Coef = BonusMaxVie / MaxVie
'    Case 1:
'        Manger_Ressources_Coef = BonusMaxEnergie / MaxEnergie
'    Case 2:
'        Manger_Ressources_Coef = BonusMaxMagie / MaxMagie
'    Case 3:
'        Manger_Ressources_Coef = BonusMaxMoral / MaxMoral
'    End Select
'End Function
'
'Public Function PeutConstruire_Maison(ByVal IndiceMaison As Long, _
'                                      ByVal Parametres As ClsJeuParam) As Boolean
'    Dim i As Long
'    If IndiceMaison >= 0 Then
'        PeutConstruire_Maison = True
'        For i = 0 To Parametres.NombreRessources - 1
'            'If IndiceMaison = 998 Then Debug.Print "PeutConstruire_Maison"
'            If Me.Definir_Ressources(i) < Parametres.Batiment_Prix(IndiceMaison, i) Then
'                PeutConstruire_Maison = False
'                i = Parametres.NombreRessources - 1
'            End If
'        Next i
'    End If
'End Function
'
'Public Sub Creer_Objet(ByVal NumeroObjet As Integer, _
'                       ByVal Maison As ClsJeuBatiment, _
'                       ByVal Parametres As ClsJeuParam)
'    IndiceObjet = NumeroObjet
'    EtatFabricationObjet = 0
'    Action = 500
'End Sub
'
'Private Sub Fabriquer_Objet(ByVal Parametres As ClsJeuParam, _
'                            ByRef Maison As ClsJeuBatiment, _
'                            ByRef NoCommentaire As Integer)
'    Dim i As Integer
'    Dim VitesseFabrication As Double
'    Dim Objet As Integer
'    Dim Competence As Integer
'
'    Objet = Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, IndiceObjet)
'    Competence = Parametres.Objet_Competence(Objet)
'    VitesseFabrication = VitesseFabricationObjet * Me.Vitalite * Definir_Niveau_CompetenceObjets(Competence)
'
'    EtatFabricationObjet = EtatFabricationObjet + VitesseFabrication
'    'Le personnage se fatigue.
'    Me.Definir_Energie = Energie - VitesseFatigueFabrication
'    Me.Definir_Moral = Moral - VitesseDemoralisationFabrication
'    'Il gagne de l'expérience dans sa compétence objet.
'    Gestion_Niveau CaracObjets(Competence), _
'                   VitesseFabrication * CoefExpCompObjets * CoefExpCompObjetsFabrication, _
'                   ExpObjets(Competence), NivObjets(Competence), ExpNivObjets(Competence), ExpNivPlusObjets(Competence), 8
'    'VitesseFabrication / Parametres.Objet_DureeFabrication(Objet) * CoefExpCompObjets,
'
'    'Il gagne de l'expérience avec les objets qu'il utilise et qui lui offrent un bonus.
'    If Bonus_CompetenceObjets(Competence) Then
'        For i = 0 To UBound(ObjetsEquipes())
'            If ObjetsEquipes(i).Param_BonusCompetencesObjets(Competence) > 0 Then
'                Gestion_Niveau CaracObjets(ObjetsEquipes(i).Competence), _
'                               VitesseFabrication * CoefExpCompObjets * CoefExpCompObjetsFabrication, _
'                               ExpObjets(ObjetsEquipes(i).Competence), NivObjets(ObjetsEquipes(i).Competence), ExpNivObjets(ObjetsEquipes(i).Competence), ExpNivPlusObjets(ObjetsEquipes(i).Competence), 8
'            End If
'        Next i
'    End If
'    If EtatFabricationObjet >= Parametres.Objet_DureeFabrication(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, IndiceObjet)) Then
'        'Terminé.
'        'EtatFabricationObjet = 0
'        Maison.Stock_Objets(IndiceObjet) = Maison.Stock_Objets(IndiceObjet) + 1
'        NoCommentaire = 0
'        Action = 100
'        'Le joueur paie son coût.
'        For i = 0 To Parametres.NombreRessources - 1
'            Me.Definir_Ressources(i) = Me.Definir_Ressources(i) _
'                                       - Parametres.Objet_PrixRessources(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, IndiceObjet), i)
'        Next i
'    Else
'        Action = 500
'    End If
'End Sub
'
'Public Function PeutConstruire_Objet(ByVal NumeroObjet As Long, _
'                                     ByVal Maison As ClsJeuBatiment, _
'                                     ByVal Parametres As ClsJeuParam) As Boolean
'    Dim i As Long
'    'Vérifie qu'il y a de la place pour stocker l'objet.
'    If NumeroObjet >= 0 And Maison.Definir_Stock <= Maison.MaxStock - 1 Then
'        PeutConstruire_Objet = True
'        For i = 0 To Parametres.NombreRessources - 1
'            If Me.Definir_Ressources(i) < Parametres.Objet_PrixRessources(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, NumeroObjet), i) Then
'                PeutConstruire_Objet = False
'                i = Parametres.NombreRessources - 1
'            End If
'        Next i
'    End If
'End Function
'
'' ************************************************************************************
'' *                             Manipulation de l'inventaire                         *
'' ************************************************************************************
'
'Public Sub Objet_Acheter_Maison(ByVal NumeroObjet As Integer, _
'                                ByVal Maison As ClsJeuBatiment, _
'                                ByVal Parametres As ClsJeuParam, _
'                                ByRef Commentaires As ClsJeuCommentaires)
'    Dim Prix As Integer
'    'Transfère un objet d'une maison vers son inventaire.
'    Prix = Prix_Marchandee(Parametres.Objet_PrixVente(Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, NumeroObjet)) * Parametres.CoefficientAchat, Parametres)
'    If Argent >= Prix Then
'        Argent = Argent - Prix
'        Maison.Argent = Maison.Argent + Prix
'        Objet_Ramasser_Maison NumeroObjet, Maison, Parametres, Commentaires
'        'Compétence marchandage.
'        Gestion_Niveau_Comp_Speciale 3, Prix, Parametres
'    End If
'End Sub
'Public Sub Objet_Ramasser_Maison(ByVal NumeroObjet As Integer, _
'                                 ByVal Maison As ClsJeuBatiment, _
'                                 ByVal Parametres As ClsJeuParam, _
'                                 ByRef Commentaires As ClsJeuCommentaires)
'    'Transfère un objet d'une maison dans son inventaire.
'    If Maison.Stock_Objets(NumeroObjet) >= 1 And _
'       Not Inventaire_Plein Then
'        'If Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, NumeroObjet) = -1 Then Debug.Print "Erreur : Objet_Ramasser_Maison"
'        Objet_Ajouter_Inventaire Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, NumeroObjet), Parametres, Commentaires, True
'        Maison.Stock_Objets(NumeroObjet) = Maison.Stock_Objets(NumeroObjet) - 1
'    End If
'End Sub
'Public Sub Objet_Deposer_Maison(ByVal NumeroObjet As Integer, _
'                                ByVal Maison As ClsJeuBatiment, _
'                                ByVal Parametres As ClsJeuParam, _
'                                Optional ByVal ObjetInventaire As Boolean = True)
'    Dim i As Integer
'    Dim ObjetDepose As Boolean
'    Dim TempObjet As Integer
'    If ObjetInventaire Then
'        TempObjet = ObjetsInventaires(NumeroObjet).Numero
'    Else
'        TempObjet = NumeroObjet
'    End If
'    'Transfère un objet de son inventaire vers sa maison
'    If Maison.Definir_Stock <= Maison.MaxStock - 1 Then
'        For i = 0 To Parametres.Batiment_NombreObjetsFabriques(Maison.TypeBatiment) - 1
'            If TempObjet = Parametres.Batiment_ObjetFabrique(Maison.TypeBatiment, i) Then
'                If ObjetInventaire Then
'                    ObjetsInventaires(NumeroObjet).Decharger
'                End If
'                Maison.Stock_Objets(i) = Maison.Stock_Objets(i) + 1
'                ObjetDepose = True
'                i = Parametres.Batiment_NombreObjetsFabriques(Maison.TypeBatiment)
'            End If
'        Next i
'    End If
'    If Not ObjetDepose Then
'        Objet_Deposer_Sol NumeroObjet, ObjetInventaire
'    End If
'End Sub
'Public Sub Objet_Deposer_Sol(ByVal NumeroObjet As Integer, _
'                             Optional ByVal ObjetInventaire As Boolean = True)
'    'Se débarasse de l'objet sélectionné.
'    Dim IndexTresor As Long
'    Dim TempObjet As Integer
'    If ObjetInventaire Then
'        TempObjet = ObjetsInventaires(NumeroObjet).Numero
'    Else
'        TempObjet = NumeroObjet
'    End If
'    'If Not ObjetsInventaires(NumeroObjet).Insepararable Then
'    If Not Parametres.Objet_Inseparable(TempObjet) Then
'        IndexTresor = Tresor_Creer(PositionX, PositionY)
'        Tresors(IndexTresor).Ajouter_Objet TempObjet, Parametres
'        If ObjetInventaire Then
'            ObjetsInventaires(NumeroObjet).Decharger
'        Else
'            ObjetSelectionne = -1
'        End If
'    End If
'    'Dépose également les ressources sélectionnées.
'    Ressources_Deposer_Sol IndexTresor
'End Sub
'Public Sub Objet_Vendre_Chateau(ByVal NumeroObjet As Integer, _
'                                ByVal Fief As ClsJeuFief, _
'                                ByVal Parametres As ClsJeuParam, _
'                                Optional ByVal ObjetInventaire As Boolean = True)
'    'Vend l'objet sélectionné au chateau.
'    Dim i  As Integer
'    Dim TempObjet As Integer
'    If ObjetInventaire Then
'        TempObjet = ObjetsInventaires(NumeroObjet).Numero
'    Else
'        TempObjet = NumeroObjet
'    End If
'    With Parametres
'    'If Not ObjetsInventaires(NumeroObjet).Insepararable Then
'    If Not .Objet_Inseparable(NumeroObjet) Then
'        'Definir_Argent = Definir_Argent + ObjetsInventaires(NumeroObjet).PrixVente
'        Definir_Argent = Definir_Argent + .Objet_PrixVente(NumeroObjet)
'        'Le fief recycle l'objet et le reconvertit en ressources.
'        For i = 0 To .NombreRessources - 1
'            'Fief.Definir_Ressources(i) = Fief.Definir_Ressources(i) + ObjetsInventaires(NumeroObjet).Param_PrixRessources(i)
'            Fief.Definir_Ressources(i) = Fief.Definir_Ressources(i) + .Objet_PrixRessources(NumeroObjet, i)
'        Next i
'        If ObjetInventaire Then
'            ObjetsInventaires(NumeroObjet).Decharger
'        End If
'    End If
'    End With
'End Sub
'Public Sub Objet_Equipes_Detruire(ByVal NumeroObjet As Integer)
'    'L'objet redonne une apparence normale au personnage.
'    If ObjetsEquipes(NumeroObjet).ChangerApparencePerso Then
'        ChoisirApparence = False
'        Apparence = 0
'    End If
'    'Détruit de l'objet équipé sélectionné.
'    ObjetsEquipes(NumeroObjet).Decharger
'    Actualiser_Bonus
'End Sub
'Public Sub Objet_Equipes_Vider_Tout()
'    'Se débarasse de tous les objets qu'il possède.
'    Dim i As Integer
'    For i = 0 To Nombre_Objets_Equipes - 1
'        Objet_Equipes_Detruire i
'    Next i
'End Sub
'Public Sub Objet_Inventaire_Detruire(ByVal NumeroObjet As Integer)
'    'Détruit de l'objet dans l'inventaire sélectionné.
'    ObjetsInventaires(NumeroObjet).Decharger
'End Sub
'Public Sub Objet_Inventaire_Vider_Tout()
'    'Se débarasse de tous les objets qu'il a dans son inventaire.
'    Dim i As Integer
'    For i = 0 To Nombre_Objets_Inventaire - 1
'        Objet_Inventaire_Detruire i
'    Next i
'End Sub
'Public Sub Objet_Ajouter_Inventaire(ByVal IndiceObjet As Integer, _
'                                    ByVal Parametres As ClsJeuParam, _
'                                    ByRef Commentaires As ClsJeuCommentaires, _
'                                    Optional ByVal NouveauObjet As Boolean)
'    'Place un nouvel objet dans son inventaire.
'    Dim i As Integer
'    For i = 0 To Nombre_Objets_Inventaire - 1
'        If Not ObjetsInventaires(i).Actif Then
'            ObjetsInventaires(i).Charger IndiceObjet, Parametres
'            'Si c'est un nouveau objet, on s'équipe directement avec à condition que le joueur ne soit pas équipé.
'            If NouveauObjet And Equipement_Libre(IndiceObjet, Parametres) And _
'               Not ObjetsInventaires(i).Utilisable Then  'Empêche les objets utilisables d'être utilisés directement.
'                Objet_Equiper i, Parametres, Commentaires
'            End If
'            i = Nombre_Objets_Inventaire
'        End If
'    Next i
'End Sub
'Public Function Inventaire_Plein() As Boolean
'    'Renvoie vrai s'il n'y a plus de place dans l'inventaire.
'    Dim i As Integer
'    Inventaire_Plein = True
'    For i = 0 To Nombre_Objets_Inventaire - 1
'        If Not ObjetsInventaires(i).Actif Then
'            Inventaire_Plein = False
'            i = Nombre_Objets_Inventaire
'        End If
'    Next i
'End Function
'Public Function Objet_Equiper(ByVal NumeroObjet As Integer, _
'                              ByVal Parametres As ClsJeuParam, _
'                              ByRef Commentaires As ClsJeuCommentaires, _
'                              Optional ByVal ObjetInventaire As Boolean = True, _
'                              Optional ByVal CoteDroit As Boolean) As Boolean
'    'Renvoie vrai, si l'objet est détruit.
'    'CoteDroit : s'il ont place l'arme ou l'anneau à droite.
'    Dim TempObjet, TempObjet1, TempObjet2 As Integer
'    TempObjet1 = -1
'    TempObjet2 = -1
'
'    If ObjetInventaire Then
'        TempObjet = ObjetsInventaires(NumeroObjet).Numero
'    Else
'        TempObjet = NumeroObjet
'    End If
'
'    With Parametres
'    'S'équipe avec l'objet de l'inventaire désigné.
'    'If Not ObjetsInventaires(NumeroObjet).Utilisable Then
'    If Not .Objet_Utilisable(TempObjet) Then
'        Objet_Equiper = True
'        If Not Equipement_Libre(TempObjet, Parametres) Then
'            Select Case .Objet_Categorie(TempObjet)
'            Case 0: 'Objet à une main.
'                If CoteDroit Then
'                    If ObjetsEquipes(0).Actif And ObjetsEquipes(1).Actif Then
'                        TempObjet1 = ObjetsEquipes(1).Numero
'                        Objet_Equipes_Detruire 1
'                    End If
'                ElseIf ObjetsEquipes(1).Actif Or ObjetsEquipes(0).Actif And ObjetsEquipes(0).Categorie = 1 Then
'                    TempObjet1 = ObjetsEquipes(0).Numero
'                    Objet_Equipes_Detruire 0
'                End If
'            Case 1: 'Arme à 2 mains
'                TempObjet1 = ObjetsEquipes(0).Numero
'                Objet_Equipes_Detruire 0
'                If ObjetsEquipes(1).Actif Then
'                    TempObjet2 = ObjetsEquipes(1).Numero
'                    Objet_Equipes_Detruire 1
'                End If
'            Case 6, 7: 'Anneaux.
'                If CoteDroit Then
'                    If ObjetsEquipes(6).Actif And ObjetsEquipes(7).Actif Then
'                        TempObjet1 = ObjetsEquipes(7).Numero
'                        Objet_Equipes_Detruire 7
'                    End If
'                ElseIf ObjetsEquipes(7).Actif Or ObjetsEquipes(6).Actif And ObjetsEquipes(6).Categorie = 7 Then
'                    TempObjet1 = ObjetsEquipes(6).Numero
'                    Objet_Equipes_Detruire 6
'                End If
'            Case Else
'                'TempObjet1 = ObjetsEquipes(ObjetsInventaires(NumeroObjet).Categorie).Numero
'                TempObjet1 = ObjetsEquipes(.Objet_Categorie(TempObjet)).Numero
'                'Objet_Equipes_Detruire ObjetsInventaires(NumeroObjet).Categorie
'                Objet_Equipes_Detruire .Objet_Categorie(TempObjet)
'            End Select
'        End If
'        'L'objet métamorphose le personnage.
'        'If ObjetsInventaires(NumeroObjet).ChangerApparencePerso Then
'        If .Objet_ChangerApparencePerso(TempObjet) Then
'            ChoisirApparence = True
'            Apparence = ApparenceChoisie
'        End If
'        'Select Case ObjetsInventaires(NumeroObjet).Categorie
'        Select Case .Objet_Categorie(TempObjet)
'        Case 0: 'Objet à 1 main.
'            If CoteDroit Then
'                If ObjetsEquipes(1).Actif Then
'                    If ObjetsEquipes(0).Actif Then
'                        ObjetsEquipes(1).Charger TempObjet, Parametres
'                    Else
'                        ObjetsEquipes(0).Charger TempObjet, Parametres
'                    End If
'                Else
'                    ObjetsEquipes(1).Charger TempObjet, Parametres
'                End If
'            Else
'                If ObjetsEquipes(0).Actif Then
'                    'ObjetsEquipes(1).Charger ObjetsInventaires(NumeroObjet).Numero, Parametres
'                    ObjetsEquipes(1).Charger TempObjet, Parametres
'                Else
'                    'ObjetsEquipes(0).Charger ObjetsInventaires(NumeroObjet).Numero, Parametres
'                    ObjetsEquipes(0).Charger TempObjet, Parametres
'                End If
'            End If
'        Case 1: 'Objet a 2 mains.
'            'ObjetsEquipes(0).Charger ObjetsInventaires(NumeroObjet).Numero, Parametres
'            ObjetsEquipes(0).Charger TempObjet, Parametres
'        Case 6, 7: 'Anneaux.
'            If CoteDroit Then
'                If ObjetsEquipes(7).Actif Then
'                    If ObjetsEquipes(6).Actif Then
'                        ObjetsEquipes(7).Charger TempObjet, Parametres
'                    Else
'                        ObjetsEquipes(6).Charger TempObjet, Parametres
'                    End If
'                Else
'                    ObjetsEquipes(7).Charger TempObjet, Parametres
'                End If
'            Else
'                If ObjetsEquipes(6).Actif Then
'                    'ObjetsEquipes(7).Charger ObjetsInventaires(NumeroObjet).Numero, Parametres
'                    ObjetsEquipes(7).Charger TempObjet, Parametres
'                Else
'                    'ObjetsEquipes(6).Charger ObjetsInventaires(NumeroObjet).Numero, Parametres
'                    ObjetsEquipes(6).Charger TempObjet, Parametres
'                End If
'            End If
'        Case Else
'            'ObjetsEquipes(ObjetsInventaires(NumeroObjet).Categorie).Charger ObjetsInventaires(NumeroObjet).Numero, Parametres
'            ObjetsEquipes(.Objet_Categorie(TempObjet)).Charger TempObjet, Parametres
'        End Select
'        If ObjetInventaire Then
'            ObjetsInventaires(NumeroObjet).Decharger
'        End If
'        'On récupère l'équipement qui a été remplacé.
'        If TempObjet1 >= 0 Then
'            Objet_Ajouter_Inventaire TempObjet1, Parametres, Commentaires
'        End If
'        If TempObjet2 >= 0 Then
'            Objet_Ajouter_Inventaire TempObjet2, Parametres, Commentaires
'        End If
'        'On supprime les armes en trop.
'        If ObjetsEquipes(0).Actif And ObjetsEquipes(0).Categorie = 1 And ObjetsEquipes(1).Actif Then
'            ObjetsEquipes(1).Decharger
'        End If
'        Actualiser_Bonus
'        Calculer_Apparence Maisons(Numero), Parametres
'    'ElseIf (Vivant Or ObjetsInventaires(NumeroObjet).UtilisationMort) And NumeroObjetUtilise = -1 And Not (Vivant And ObjetsInventaires(NumeroObjet).SelfResurrection) Then
'    ElseIf (Vivant Or .Objet_UtilisationMort(TempObjet)) And NumeroObjetUtilise = -1 And Not (Vivant And .Objet_SelfResurrection(TempObjet)) Then
'        'If Vie > ObjetsInventaires(NumeroObjet).PerteVie Or Not Vivant Then
'        If Vie > .Objet_PerteVie(TempObjet) Or Not Vivant Then
'            'If Energie > ObjetsInventaires(NumeroObjet).PerteEnergie Then
'            If Energie > .Objet_PerteEnergie(TempObjet) Then
'                'If Magie > ObjetsInventaires(NumeroObjet).PerteMagie Then
'                If Magie > .Objet_PerteMagie(TempObjet) Then
'                    'If Moral > ObjetsInventaires(NumeroObjet).PerteMoral Then
'                    If Moral > .Objet_PerteMoral(TempObjet) Then
'                        'If Argent >= ObjetsInventaires(NumeroObjet).PerteArgent Then
'                        If Argent >= .Objet_PerteArgent(TempObjet) Then
'                            'If Ressources(PerteRessourceIndice) >= ObjetsInventaires(NumeroObjet).PerteRessourceQuantite Then
'                            If Ressources(PerteRessourceIndice) >= .Objet_PerteRessourceQuantite(TempObjet) Then
'                                'On utilise l'objet.
'                                'NumeroObjetUtilise = ObjetsInventaires(NumeroObjet).Numero
'                                NumeroObjetUtilise = TempObjet
'                                'On gagne de l'expérience.
'                                'Gestion_Niveau CaracObjets(ObjetsInventaires(NumeroObjet).Competence), _
'                                               ObjetsInventaires(NumeroObjet).Experience, _
'                                               ExpObjets(ObjetsInventaires(NumeroObjet).Competence), NivObjets(ObjetsInventaires(NumeroObjet).Competence), ExpNivObjets(ObjetsInventaires(NumeroObjet).Competence), ExpNivPlusObjets(ObjetsInventaires(NumeroObjet).Competence), 8
'                                Gestion_Niveau CaracObjets(.Objet_Competence(TempObjet)), _
'                                               .Objet_Experience(TempObjet), _
'                                               ExpObjets(.Objet_Competence(TempObjet)), NivObjets(.Objet_Competence(TempObjet)), ExpNivObjets(.Objet_Competence(TempObjet)), ExpNivPlusObjets(.Objet_Competence(TempObjet)), 8
'                                'On perd des caractéristiques.
'                                'Definir_Vie = Vie - ObjetsInventaires(NumeroObjet).PerteVie
'                                Definir_Vie = Vie - .Objet_PerteVie(TempObjet)
'                                'Definir_Energie = Energie - ObjetsInventaires(NumeroObjet).PerteEnergie
'                                Definir_Energie = Energie - .Objet_PerteEnergie(TempObjet)
'                                'Definir_Magie = Magie - ObjetsInventaires(NumeroObjet).PerteMagie
'                                Definir_Magie = Magie - .Objet_PerteMagie(TempObjet)
'                                'Definir_Moral = Moral - ObjetsInventaires(NumeroObjet).PerteMoral
'                                Definir_Moral = Moral - .Objet_PerteMoral(TempObjet)
'                                'L'objet disparaît.
'                                'If ObjetsInventaires(NumeroObjet).UtilisationUnique Then
'                                '    ObjetsInventaires(NumeroObjet).Decharger
'                                'End If
'                                If .Objet_UtilisationUnique(TempObjet) And ObjetInventaire Then
'                                    ObjetsInventaires(NumeroObjet).Decharger
'                                End If
'                                'Definir_Argent = Definir_Argent - ObjetsInventaires(NumeroObjet).PerteArgent
'                                Definir_Argent = Definir_Argent - .Objet_PerteArgent(TempObjet)
'                                'If ObjetsInventaires(NumeroObjet).PerteRessourceQuantite > 0 Then
'                                If .Objet_PerteRessourceQuantite(TempObjet) > 0 Then
'                                    'Definir_Ressources(ObjetsInventaires(NumeroObjet).PerteRessourceIndice) = Definir_Ressources(ObjetsInventaires(NumeroObjet).PerteRessourceIndice) - ObjetsInventaires(NumeroObjet).PerteRessourceQuantite
'                                    Definir_Ressources(.Objet_PerteRessourceIndice(TempObjet)) = Definir_Ressources(.Objet_PerteRessourceIndice(TempObjet)) - .Objet_PerteRessourceQuantite(TempObjet)
'                                End If
'                                Objet_Equiper = True
'                            Else
'                                'N'a plus assez de ressources.
'                                Definir_Commentaires = Commentaires.Message(34, Me)
'                            End If
'                        Else
'                            'N'a plus d'argent.
'                            Definir_Commentaires = Commentaires.Message(45, Me)
'                        End If
'                    Else
'                        'N'a plus de point de moral.
'                        Definir_Commentaires = Commentaires.Message(20, Me)
'                    End If
'                Else
'                    'N'a plus de point de magie.
'                    Definir_Commentaires = Commentaires.Message(19, Me)
'                End If
'            Else
'                'N'a plus de point d'énergie.
'                Definir_Commentaires = Commentaires.Message(18, Me)
'            End If
'        Else
'            'N'a plus de point de vie.
'            Definir_Commentaires = Commentaires.Message(17, Me)
'        End If
'    End If
'    End With
'End Function
'Private Function Equipement_Libre(ByVal IndiceObjet As Integer, _
'                                  ByVal Parametres As ClsJeuParam) As Boolean
'    'Verifie s'il y a de la place pour s'équiper avec l'objet définie.
'    Select Case Parametres.Objet_Categorie(IndiceObjet)
'    Case 0: 'Arme à une main.
'        If Not (ObjetsEquipes(0).Actif And ObjetsEquipes(1).Actif) And _
'           Not (ObjetsEquipes(0).Actif And ObjetsEquipes(0).Categorie = 1) Then
'            Equipement_Libre = True
'        End If
'    Case 1: 'Arme à 2 mains
'        If Not (ObjetsEquipes(0).Actif Or ObjetsEquipes(1).Actif) Then
'            Equipement_Libre = True
'        End If
'    Case 6, 7: 'Anneau.
'        If Not (ObjetsEquipes(6).Actif And ObjetsEquipes(7).Actif) Then
'            Equipement_Libre = True
'        End If
'    Case Else:
'        If Not ObjetsEquipes(Parametres.Objet_Categorie(IndiceObjet)).Actif Then
'            Equipement_Libre = True
'        End If
'    End Select
'End Function
'
'Public Function Objet_Posseder(ByVal NumeroObjet As Integer) As Boolean
'    'Renvoie si le personnage possède un au moins un objet de ce type.
'    Dim i As Integer
'    Objet_Posseder = Objet_NombreType(NumeroObjet) > 0
'End Function
'
'Public Function Objet_NombreType(ByVal NumeroObjet As Integer) As Integer
'    'Renvoie le nombre d'objet de ce type que possède ce personnage.
'    Dim i As Integer
'    For i = 0 To UBound(ObjetsEquipes())
'        If ObjetsEquipes(i).Actif And _
'           ObjetsEquipes(i).Numero = NumeroObjet Then
'            Objet_NombreType = Objet_NombreType + 1
'        End If
'    Next i
'    For i = 0 To UBound(ObjetsInventaires())
'        If ObjetsInventaires(i).Actif And _
'           ObjetsInventaires(i).Numero = NumeroObjet Then
'            Objet_NombreType = Objet_NombreType + 1
'        End If
'    Next i
'End Function
'
'Public Sub Objet_Ranger(ByVal NumeroObjet As Integer, _
'                        ByVal Parametres As ClsJeuParam, _
'                        ByRef Commentaires As ClsJeuCommentaires)
'    'On passe l'indice de l'objet équipé.
'    'Remet l'objet équipé dans l'inventaire.
'    If Not Inventaire_Plein Then
'        Objet_Ajouter_Inventaire ObjetsEquipes(NumeroObjet).Numero, Parametres, Commentaires
'        Objet_Equipes_Detruire NumeroObjet
'        Actualiser_Bonus
'        Calculer_Apparence Maisons(Numero), Parametres
'        'Réajuste les caractéristiques de combat.
'        If Attaque > Definir_MaxAttaque Then
'            Attaque = Definir_MaxAttaque
'        End If
'        If Defense > Definir_MaxDefense Then
'            Defense = Definir_MaxDefense
'        End If
'    End If
'End Sub
'
'Public Sub Objet_Equipe_Charger(ByVal i As Integer, _
'                                ByVal IndiceObjet As Integer, _
'                                ByVal Parametres As ClsJeuParam)
'    'If IndiceObjet = -1 Then Debug.Print "Erreur : Objet_Equipe_Charger"
'    ObjetsEquipes(i).Charger IndiceObjet, Parametres
'End Sub
'Public Sub Objet_Inventaire_Charger(ByVal i As Integer, _
'                                    ByVal IndiceObjet As Integer, _
'                                    ByVal Parametres As ClsJeuParam)
'    'If IndiceObjet = -1 Then Debug.Print "Erreur : Objet_Inventaire_Charger"
'    ObjetsInventaires(i).Charger IndiceObjet, Parametres
'End Sub
'
'Public Sub Objet_Ranger_Armes(ByRef Commentaires As ClsJeuCommentaires, _
'                              ByVal Parametres As ClsJeuParam)
'    Arreter
'    If ObjetsEquipes(0).Actif Then
'        Objet_Ranger 0, Parametres, Commentaires
'    End If
'    If ObjetsEquipes(1).Actif Then
'        Objet_Ranger 1, Parametres, Commentaires
'    End If
'End Sub
'
'Public Sub Actualiser_Bonus()
'    'Utiliser cette fonction chaque fois que l'on s'équipe ou se deséquipe d'un objet ou au changement d'un niveau de compétence d'objet.
'    Dim i As Long
'    Dim j As Long
'    Dim TempBonusDeplacement As Long
'    Dim TempBonusMaxVie As Long
'    Dim TempBonusMaxEnergie As Long
'    Dim TempBonusMaxMagie As Long
'    Dim TempBonusMaxMoral As Long
'    Dim TempBonusMaxAttaque As Long
'    Dim TempBonusMaxDefense As Long
'
'    GenerationRessources = False
'    GainArgent = 0
'    For j = 0 To UBound(BonusNivRessources())
'        GainRessource(j) = 0
'    Next j
'
'    TempBonusDeplacement = BonusDeplacement
'    TempBonusMaxVie = BonusMaxVie
'    TempBonusMaxEnergie = BonusMaxEnergie
'    TempBonusMaxMagie = BonusMaxMagie
'    TempBonusMaxMoral = BonusMaxMoral
'    TempBonusMaxAttaque = BonusMaxAttaque
'    TempBonusMaxDefense = BonusMaxDefense
'
'    PerteVie = 0
'    PerteEnergie = 0
'    PerteMagie = 0
'    PerteMoral = 0
'    PerteArgent = 0
'    PerteRessourceIndice = 0
'    PerteRessourceQuantite = 0
'
'    Ajuster_Bonus
'
'    'On cherche en premier les bonus des objets.
'    For i = 0 To UBound(ObjetsEquipes())
'        If ObjetsEquipes(i).Actif Then
'            For j = 0 To UBound(BonusNivObjets())
'                BonusNivObjets(j) = BonusNivObjets(j) + Objet_Equipe_BonusSup(ObjetsEquipes(i).Competence, ObjetsEquipes(i).Param_BonusCompetencesObjets(j))
'            Next j
'        End If
'    Next i
'    'Ajustements de la race.
'    For i = 0 To UBound(BonusNivObjets())
'        BonusNivObjets(i) = BonusNivObjets(i) * CoefRaceCompObjets(i)
'    Next i
'
'    For i = 0 To UBound(ObjetsEquipes())
'        With ObjetsEquipes(i)
'        If ObjetsEquipes(i).Actif Then
'            If Not Volant Then
'                BonusDeplacement = BonusDeplacement + Objet_Equipe_BonusSup(ObjetsEquipes(i).Competence, ObjetsEquipes(i).BonusDeplacement)
'            End If
'            BonusMaxVie = BonusMaxVie + Objet_Equipe_BonusSup(ObjetsEquipes(i).Competence, ObjetsEquipes(i).BonusVie)
'            BonusMaxEnergie = BonusMaxEnergie + Objet_Equipe_BonusSup(ObjetsEquipes(i).Competence, ObjetsEquipes(i).BonusEnergie)
'            BonusMaxMagie = BonusMaxMagie + Objet_Equipe_BonusSup(ObjetsEquipes(i).Competence, ObjetsEquipes(i).BonusMagie)
'            BonusMaxMoral = BonusMaxMoral + Objet_Equipe_BonusSup(ObjetsEquipes(i).Competence, ObjetsEquipes(i).BonusMoral)
'            BonusMaxAttaque = BonusMaxAttaque + Objet_Equipe_BonusSup(ObjetsEquipes(i).Competence, ObjetsEquipes(i).BonusAttaque)
'            BonusMaxDefense = BonusMaxDefense + Objet_Equipe_BonusSup(ObjetsEquipes(i).Competence, ObjetsEquipes(i).BonusDefense)
'            BonusPorteeAttaque = BonusPorteeAttaque + Objet_Equipe_PorteeSup(ObjetsEquipes(i).Competence, ObjetsEquipes(i).BonusPorteeAttaque)
'            BonusSurfaceAttaque = BonusSurfaceAttaque + Objet_Equipe_SurfaceSup(ObjetsEquipes(i).Competence, ObjetsEquipes(i).SurfaceAttaque)
'
'            PerteVie = PerteVie + ObjetsEquipes(i).PerteVie
'            PerteEnergie = PerteEnergie + ObjetsEquipes(i).PerteEnergie
'            PerteMagie = PerteMagie + ObjetsEquipes(i).PerteMagie
'            PerteMoral = PerteMoral + ObjetsEquipes(i).PerteMoral
'            PerteArgent = PerteArgent + ObjetsEquipes(i).PerteArgent
'            If ObjetsEquipes(i).PerteRessourceIndice >= 0 Then
'                PerteRessourceIndice = ObjetsEquipes(i).PerteRessourceIndice
'            End If
'            PerteRessourceQuantite = PerteRessourceQuantite + ObjetsEquipes(i).PerteRessourceQuantite
'
'            For j = 0 To UBound(BonusNivSpeciales())
'                BonusNivSpeciales(j) = BonusNivSpeciales(j) + Objet_Equipe_BonusSup(ObjetsEquipes(i).Competence, ObjetsEquipes(i).Param_BonusCompetencesSpeciales(j))
'            Next j
'            For j = 0 To UBound(BonusNivRessources())
'                BonusNivRessources(j) = BonusNivRessources(j) + Objet_Equipe_BonusSup(ObjetsEquipes(i).Competence, ObjetsEquipes(i).Param_BonusCompetencesRessources(j))
'            Next j
'            For j = 0 To UBound(BonusNivServices())
'                BonusNivServices(j) = BonusNivServices(j) + Objet_Equipe_BonusSup(ObjetsEquipes(i).Competence, ObjetsEquipes(i).Param_BonusCompetencesServices(j))
'            Next j
'
'            If ObjetsEquipes(i).GainArgent > 0 Then
'                GainArgent = GainArgent + Objet_Equipe_GainSup(ObjetsEquipes(i).Competence, ObjetsEquipes(i).GainArgent)
'                GenerationRessources = True
'            End If
'            For j = 0 To UBound(GainRessource())
'                If ObjetsEquipes(i).Param_GainRessource(j) > 0 Then
'                    GainRessource(j) = GainRessource(j) + Objet_Equipe_GainSup(ObjetsEquipes(i).Competence, ObjetsEquipes(i).Param_GainRessource(j))
'                    GenerationRessources = True
'                End If
'            Next j
'            BonusArmure = BonusArmure + Objet_Equipe_ArmureSup(ObjetsEquipes(i).Competence, ObjetsEquipes(i).BonusArmure)
'            BonusMaxRessources = BonusMaxRessources + Objet_Equipe_BonusSup(ObjetsEquipes(i).Competence, ObjetsEquipes(i).BonusMaxRessources)
'            BonusVitesseRegenerationVie = BonusVitesseRegenerationVie + ObjetsEquipes(i).BonusVitesseRegenerationVie
'            BonusVitesseRegenerationEnergie = BonusVitesseRegenerationEnergie + ObjetsEquipes(i).BonusVitesseRegenerationEnergie
'            BonusVitesseRegenerationMagie = BonusVitesseRegenerationMagie + ObjetsEquipes(i).BonusVitesseRegenerationMagie
'            BonusVitesseRegenerationMoral = BonusVitesseRegenerationMoral + ObjetsEquipes(i).BonusVitesseRegenerationMoral
'        End If
'        End With
'    Next i
'
'    If Not ChoisirApparence Then
'        While ApparenceObjet.Count > 0
'            ApparenceObjet.Remove 1
'        Wend
'        For i = 0 To UBound(ObjetsEquipes())
'            With ObjetsEquipes(i)
'                If ObjetsEquipes(i).Actif Then
'                    If Parametres.Objet_ChangerApparence(.Numero) Then
'                        ApparenceObjet.Add Parametres.Objet_NumeroApparence(.Numero, Feminin)
'                    End If
'                End If
'            End With
'        Next i
'        NombreApparenceObjet = ApparenceObjet.Count
'        ApparenceObjetAfficher = NombreApparenceObjet > 0
'    End If
'
'    'Ajustements de la race.
'    BonusMaxVie = BonusMaxVie * CoefRaceVie
'    BonusMaxEnergie = BonusMaxEnergie * CoefRaceEnergie
'    BonusMaxMagie = BonusMaxMagie * CoefRaceMagie
'    BonusMaxMoral = BonusMaxMoral * CoefRaceMoral
'    BonusMaxAttaque = BonusMaxAttaque * CoefRaceAttaque
'    BonusMaxDefense = BonusMaxDefense * CoefRaceDefense
'    For i = 0 To UBound(BonusNivSpeciales())
'        BonusNivSpeciales(i) = BonusNivSpeciales(i) * CoefRaceCompSpeciales(i)
'    Next i
'    For i = 0 To UBound(BonusNivRessources())
'        BonusNivRessources(i) = BonusNivRessources(i) + Parametres.PersosMaxRessources * CoefRaceCompRessources(i) - Parametres.PersosMaxRessources
'    Next i
'    For i = 0 To UBound(BonusNivServices())
'        BonusNivServices(i) = BonusNivServices(i) * CoefRaceCompServices(i)
'    Next i
'
'    'On rééquilibre les bonus.
'    Vie = Vie * BonusMaxVie / TempBonusMaxVie
'    Energie = Energie * BonusMaxEnergie / TempBonusMaxEnergie
'    Magie = Magie * BonusMaxMagie / TempBonusMaxMagie
'    Moral = Moral * BonusMaxMoral / TempBonusMaxMoral
'    RecuperationVie = RecuperationVie * BonusMaxVie / TempBonusMaxVie
'    RecuperationEnergie = RecuperationEnergie * BonusMaxEnergie / TempBonusMaxEnergie
'    RecuperationMagie = RecuperationMagie * BonusMaxMagie / TempBonusMaxMagie
'    RecuperationMoral = RecuperationMoral * BonusMaxMoral / TempBonusMaxMoral
'    Attaque = Attaque * BonusMaxAttaque / TempBonusMaxAttaque
'    Defense = Defense * BonusMaxDefense / TempBonusMaxDefense
'
'    'On élimine le surplus de ressources.
'    For i = 0 To UBound(Ressources())
'        If Ressources(i) > Definir_MaxRessources(i) Then
'            Ressources(i) = Definir_MaxRessources(i)
'        End If
'    Next i
'End Sub
'
'
''Bonus supplémentaire de l'objet grâce à la compétence.
'Public Property Get Objet_Equipe_BonusSup(ByVal NumeroCompetenceObjet As Integer, _
'                                          ByVal ValeurBonus As Single) As Integer
'    Objet_Equipe_BonusSup = Int((((ValeurBonus * _
'                                Definir_Niveau_CompetenceObjets(NumeroCompetenceObjet) / _
'                                CaracObjetsDepart) - ValeurBonus) * _
'                                CoefExpCompObjetsBonus) + ValeurBonus)
'End Property
'Public Property Get Objet_Equipe_GainSup(ByVal NumeroCompetenceObjet As Integer, _
'                                         ByVal ValeurBonus As Single) As Single
'    Objet_Equipe_GainSup = (((ValeurBonus * _
'                              Definir_Niveau_CompetenceObjets(NumeroCompetenceObjet) / _
'                              CaracObjetsDepart) - ValeurBonus) * _
'                              CoefExpCompObjetsBonus) + ValeurBonus
'End Property
'Public Property Get Objet_Equipe_PorteeSup(ByVal NumeroCompetenceObjet As Integer, _
'                                           ByVal ValeurBonus As Single) As Integer
'    Objet_Equipe_PorteeSup = Int((((ValeurBonus * _
'                                 Definir_Niveau_CompetenceObjets(NumeroCompetenceObjet) / _
'                                 CaracObjetsDepart) - ValeurBonus) * _
'                                 CoefExpCompPorteeBonus) + ValeurBonus)
'End Property
'Public Property Get Objet_Equipe_SurfaceSup(ByVal NumeroCompetenceObjet As Integer, _
'                                            ByVal ValeurBonus As Single) As Integer
'    Objet_Equipe_SurfaceSup = Int((((ValeurBonus * _
'                                 Definir_Niveau_CompetenceObjets(NumeroCompetenceObjet) / _
'                                 CaracObjetsDepart) - ValeurBonus) * _
'                                 CoefExpCompSurfaceBonus) + ValeurBonus)
'End Property
'
'Public Property Get Objet_Equipe_ArmureSup(ByVal NumeroCompetenceObjet As Integer, _
'                                           ByVal ValeurBonus As Single) As Integer
'    Objet_Equipe_ArmureSup = Int((((ValeurBonus * _
'                                 Definir_Niveau_CompetenceObjets(NumeroCompetenceObjet) / _
'                                 CaracObjetsDepart) - ValeurBonus) * _
'                                 CoefExpCompArmureBonus) + ValeurBonus)
'End Property
'Public Property Get Progression_CompetenceSpeciale(ByVal Index As Integer) As Double
'    Progression_CompetenceSpeciale = (ExpSpeciales(Index) - ExpNivSpeciales(Index)) / (ExpNivPlusSpeciales(Index) - ExpNivSpeciales(Index))
'End Property
'Public Property Get Progression_CompetenceRessource(ByVal Index As Integer) As Double
'    Progression_CompetenceRessource = (ExpRessources(Index) - ExpNivRessources(Index)) / (ExpNivPlusRessources(Index) - ExpNivRessources(Index))
'End Property
'Public Property Get Progression_CompetenceService(ByVal Index As Integer) As Double
'    Progression_CompetenceService = (ExpServices(Index) - ExpNivServices(Index)) / (ExpNivPlusServices(Index) - ExpNivServices(Index))
'End Property
'Public Property Get Progression_CompetenceObjet(ByVal Index As Integer) As Double
'    Progression_CompetenceObjet = (ExpObjets(Index) - ExpNivObjets(Index)) / (ExpNivPlusObjets(Index) - ExpNivObjets(Index))
'End Property
Public Property Get QuantiteRessources(ByVal Indice As Integer) As Double
    If Indice > 0 And Indice < Persos(Numero).Nombre_CompetencesRessources Then
        QuantiteRessources = Persos(Numero).Definir_ressources(Indice - 1)
    End If
End Property
Public Property Get Inactif() As Boolean
    Inactif = Persos(Numero).Inactif
End Property
''Propriétés des comptétences Speciales.
'Public Property Get Nombre_CompetencesSpeciales() As Integer
'    Nombre_CompetencesSpeciales = UBound(CaracSpeciales()) + 1
'End Property
'Public Property Get Definir_ExpSpeciales(ByVal Index As Long) As Double
'    If Index >= 0 Then
'        Definir_ExpSpeciales = ExpSpeciales(Index)
'    End If
'End Property
'Public Property Let Definir_ExpSpeciales(ByVal Index As Long, ByVal Valeur As Double)
'    ExpSpeciales(Index) = Valeur
'End Property
'Public Property Get Definir_ExpNivSpeciales(ByVal Index As Long) As Double
'    Definir_ExpNivSpeciales = ExpNivSpeciales(Index)
'End Property
'Public Property Let Definir_ExpNivSpeciales(ByVal Index As Long, ByVal Valeur As Double)
'    ExpNivSpeciales(Index) = Valeur
'End Property
'Public Property Get Definir_ExpNivPlusSpeciales(ByVal Index As Long) As Double
'    Definir_ExpNivPlusSpeciales = ExpNivPlusSpeciales(Index)
'End Property
'Public Property Let Definir_ExpNivPlusSpeciales(ByVal Index As Long, ByVal Valeur As Double)
'    ExpNivPlusSpeciales(Index) = Valeur
'End Property
'Public Property Get Definir_Niveau_CompetenceSpeciales(ByVal Index As Long) As Integer
'    Definir_Niveau_CompetenceSpeciales = BonusNivSpeciales(Index)
'End Property
'Public Property Get Niveau_CompetenceSpeciales(ByVal Index As Long) As Integer
'    Niveau_CompetenceSpeciales = NivSpeciales(Index)
'End Property
'Public Property Let Niveau_CompetenceSpeciales(ByVal Index As Long, ByVal Valeur As Integer)
'    NivSpeciales(Index) = Valeur
'End Property
'Public Property Get Definir_Carac_CompetenceSpeciales(ByVal Index As Long) As Integer
'    Definir_Carac_CompetenceSpeciales = CaracSpeciales(Index)
'End Property
'Public Property Let Definir_Carac_CompetenceSpeciales(ByVal Index As Long, ByVal Valeur As Integer)
'    CaracSpeciales(Index) = Valeur
'End Property
'Public Property Get Bonus_CompetenceSpeciales(ByVal Index As Long) As Boolean
'    Bonus_CompetenceSpeciales = BonusNivSpeciales(Index) > CaracSpeciales(Index) * CoefRaceCompSpeciales(Index)
'End Property
'
''Propriété des ressources.
'Public Property Get Nombre_CompetencesRessources() As Integer
'    Nombre_CompetencesRessources = UBound(MaxRessources()) + 1
'End Property
'Public Property Get Definir_Carac_Ressources(ByVal Index As Long) As Long
'    Definir_Carac_Ressources = MaxRessources(Index)
'End Property
'Public Property Let Definir_Carac_Ressources(ByVal Index As Long, ByVal Valeur As Long)
'    MaxRessources(Index) = Valeur
'End Property
'Public Property Get Definir_MaxRessources(ByVal Index As Long) As Double
'    If Index >= 0 Then
'        Definir_MaxRessources = MaxRessources(Index) + BonusMaxRessources
'    End If
'End Property
'Public Property Let Definir_MaxRessources(ByVal Index As Long, ByVal Valeur As Double)
'    MaxRessources(Index) = Valeur
'End Property
'Public Property Get Definir_ExpRessources(ByVal Index As Long) As Double
'    If Index >= 0 Then
'        Definir_ExpRessources = ExpRessources(Index)
'    End If
'End Property
'Public Property Let Definir_ExpRessources(ByVal Index As Long, ByVal Valeur As Double)
'    ExpRessources(Index) = Valeur
'End Property
'Public Property Get Definir_ExpNivRessources(ByVal Index As Long) As Double
'    Definir_ExpNivRessources = ExpNivRessources(Index)
'End Property
'Public Property Let Definir_ExpNivRessources(ByVal Index As Long, ByVal Valeur As Double)
'    ExpNivRessources(Index) = Valeur
'End Property
'Public Property Get Definir_ExpNivPlusRessources(ByVal Index As Long) As Double
'    Definir_ExpNivPlusRessources = ExpNivPlusRessources(Index)
'End Property
'Public Property Let Definir_ExpNivPlusRessources(ByVal Index As Long, ByVal Valeur As Double)
'    ExpNivPlusRessources(Index) = Valeur
'End Property
'Public Property Get Definir_Niveau_CompetenceRessources(ByVal Index As Long) As Integer
'    Definir_Niveau_CompetenceRessources = BonusNivRessources(Index)
'End Property
'Public Property Get Niveau_CompetenceRessources(ByVal Index As Long) As Integer
'    Niveau_CompetenceRessources = NivRessources(Index)
'End Property
'Public Property Let Niveau_CompetenceRessources(ByVal Index As Long, ByVal Valeur As Integer)
'    NivRessources(Index) = Valeur
'End Property
'Public Property Get Bonus_CompetenceRessources(ByVal Index As Long) As Boolean
'    Bonus_CompetenceRessources = BonusNivRessources(Index) > NivRessources(Index) * CoefRaceCompRessources(Index)
'End Property
'
''Propriétés des comptétences de services.
'Public Property Get Nombre_CompetencesServices() As Integer
'    Nombre_CompetencesServices = UBound(CaracServices()) + 1
'End Property
'Public Property Get Definir_ExpServices(ByVal Index As Long) As Double
'    If Index >= 0 Then
'        Definir_ExpServices = ExpServices(Index)
'    End If
'End Property
'Public Property Let Definir_ExpServices(ByVal Index As Long, ByVal Valeur As Double)
'    ExpServices(Index) = Valeur
'End Property
'Public Property Get Definir_ExpNivServices(ByVal Index As Long) As Double
'    Definir_ExpNivServices = ExpNivServices(Index)
'End Property
'Public Property Let Definir_ExpNivServices(ByVal Index As Long, ByVal Valeur As Double)
'    ExpNivServices(Index) = Valeur
'End Property
'Public Property Get Definir_ExpNivPlusServices(ByVal Index As Long) As Double
'    Definir_ExpNivPlusServices = ExpNivPlusServices(Index)
'End Property
'Public Property Let Definir_ExpNivPlusServices(ByVal Index As Long, ByVal Valeur As Double)
'    ExpNivPlusServices(Index) = Valeur
'End Property
'Public Property Get Definir_Niveau_CompetenceServices(ByVal Index As Long) As Integer
'    Definir_Niveau_CompetenceServices = BonusNivServices(Index)
'End Property
'Public Property Get Niveau_CompetenceServices(ByVal Index As Long) As Integer
'    Niveau_CompetenceServices = NivServices(Index)
'End Property
'Public Property Let Niveau_CompetenceServices(ByVal Index As Long, ByVal Valeur As Integer)
'    NivServices(Index) = Valeur
'End Property
'Public Property Get Definir_Carac_CompetenceServices(ByVal Index As Long) As Integer
'    Definir_Carac_CompetenceServices = CaracServices(Index)
'End Property
'Public Property Let Definir_Carac_CompetenceServices(ByVal Index As Long, ByVal Valeur As Integer)
'    CaracServices(Index) = Valeur
'End Property
'Public Property Get Bonus_CompetenceServices(ByVal Index As Long) As Boolean
'    Bonus_CompetenceServices = BonusNivServices(Index) > CaracServices(Index) * CoefRaceCompServices(Index)
'End Property
'
''Propriétés des comptétences d'objets.
'Public Property Get Nombre_CompetencesObjets() As Integer
'    Nombre_CompetencesObjets = UBound(CaracObjets()) + 1
'End Property
'Public Property Get Definir_ExpObjets(ByVal Index As Long) As Double
'    If Index >= 0 Then
'        Definir_ExpObjets = ExpObjets(Index)
'    End If
'End Property
'Public Property Let Definir_ExpObjets(ByVal Index As Long, ByVal Valeur As Double)
'    ExpObjets(Index) = Valeur
'End Property
'Public Property Get Definir_ExpNivObjets(ByVal Index As Long) As Double
'    Definir_ExpNivObjets = ExpNivObjets(Index)
'End Property
'Public Property Let Definir_ExpNivObjets(ByVal Index As Long, ByVal Valeur As Double)
'    ExpNivObjets(Index) = Valeur
'End Property
'Public Property Get Definir_ExpNivPlusObjets(ByVal Index As Long) As Double
'    Definir_ExpNivPlusObjets = ExpNivPlusObjets(Index)
'End Property
'Public Property Let Definir_ExpNivPlusObjets(ByVal Index As Long, ByVal Valeur As Double)
'    ExpNivPlusObjets(Index) = Valeur
'End Property
'Public Property Get Definir_Niveau_CompetenceObjets(ByVal Index As Long) As Integer
'    Definir_Niveau_CompetenceObjets = BonusNivObjets(Index)
'End Property
'Public Property Get Niveau_CompetenceObjets(ByVal Index As Long) As Integer
'    Niveau_CompetenceObjets = NivObjets(Index)
'End Property
'Public Property Let Niveau_CompetenceObjets(ByVal Index As Long, ByVal Valeur As Integer)
'    NivObjets(Index) = Valeur
'End Property
'Public Property Get Definir_Carac_CompetenceObjets(ByVal Index As Long) As Integer
'    Definir_Carac_CompetenceObjets = CaracObjets(Index)
'End Property
'Public Property Let Definir_Carac_CompetenceObjets(ByVal Index As Long, ByVal Valeur As Integer)
'    CaracObjets(Index) = Valeur
'End Property
'Public Property Get Bonus_CompetenceObjets(ByVal Index As Long) As Boolean
'    Bonus_CompetenceObjets = BonusNivObjets(Index) > CaracObjets(Index) * CoefRaceCompObjets(Index)
'End Property
'Public Property Get Definir_Min_Competence(Optional ByVal Categorie As Integer) As Long
'    Dim i As Long
'    Dim Temp1, Temp2 As Long
'    Select Case Categorie
'    Case 0: Temp1 = Definir_Niveau_CompetenceSpeciales(0)
'    Case 1: Temp1 = Niveau_CompetenceSpeciales(0) + 1
'    Case 2: Temp1 = Definir_ExpSpeciales(0)
'    End Select
'    For i = 1 To Nombre_CompetencesSpeciales - 1
'        Select Case Categorie
'        Case 0: Temp2 = Definir_Niveau_CompetenceSpeciales(i)
'        Case 1: Temp2 = Niveau_CompetenceSpeciales(i) + 1
'        Case 2: Temp2 = Definir_ExpSpeciales(i)
'        End Select
'        If Temp2 < Temp1 Then
'            Temp1 = Temp2
'        End If
'    Next i
'    For i = 0 To Nombre_CompetencesRessources - 1
'        Select Case Categorie
'        Case 0: Temp2 = Persos(Noperso).Definir_MaxRessources(i)
'        Case 1: Temp2 = Persos(Noperso).Niveau_CompetenceRessources(i) + 1
'        Case 2: Temp2 = Persos(Noperso).Definir_ExpRessources(i)
'        End Select
'        If Temp2 < Temp1 Then
'            Temp1 = Temp2
'        End If
'    Next i
'    For i = 0 To Nombre_CompetencesServices - 1
'        Select Case Categorie
'        Case 0: Temp2 = Definir_Niveau_CompetenceServices(i)
'        Case 1: Temp2 = Niveau_CompetenceServices(i) + 1
'        Case 2: Temp2 = Definir_ExpServices(i)
'        End Select
'        If Temp2 < Temp1 Then
'            Temp1 = Temp2
'        End If
'    Next i
'    For i = 1 To Nombre_CompetencesObjets - 1
'        Select Case Categorie
'        Case 0: Temp2 = Definir_Niveau_CompetenceObjets(i)
'        Case 1: Temp2 = Niveau_CompetenceObjets(i) + 1
'        Case 2: Temp2 = Definir_ExpObjets(i)
'        End Select
'        If Temp2 < Temp1 Then
'            Temp1 = Temp2
'        End If
'    Next i
'    Definir_Min_Competence = Temp1
'End Property
'Public Property Get Definir_Moy_Competence(Optional ByVal Categorie As Integer) As Double
'    Dim i As Long
'    Dim Temp As Long
'    For i = 0 To Nombre_CompetencesSpeciales - 1
'        Select Case Categorie
'        Case 0: Temp = Temp + Definir_Niveau_CompetenceSpeciales(i)
'        Case 1: Temp = Temp + Niveau_CompetenceSpeciales(i) + 1
'        Case 2: Temp = Temp + Definir_ExpSpeciales(i)
'        End Select
'    Next i
'    For i = 0 To Nombre_CompetencesRessources - 1
'        Select Case Categorie
'        Case 0: Temp = Temp + Persos(Noperso).Definir_MaxRessources(i)
'        Case 1: Temp = Temp + Persos(Noperso).Niveau_CompetenceRessources(i) + 1
'        Case 2: Temp = Temp + Persos(Noperso).Definir_ExpRessources(i)
'        End Select
'    Next i
'    For i = 0 To Nombre_CompetencesServices - 1
'        Select Case Categorie
'        Case 0: Temp = Temp + Definir_Niveau_CompetenceServices(i)
'        Case 1: Temp = Temp + Niveau_CompetenceServices(i) + 1
'        Case 2: Temp = Temp + Definir_ExpServices(i)
'        End Select
'    Next i
'    For i = 1 To Nombre_CompetencesObjets - 1
'        Select Case Categorie
'        Case 0: Temp = Temp + Definir_Niveau_CompetenceObjets(i)
'        Case 1: Temp = Temp + Niveau_CompetenceObjets(i) + 1
'        Case 2: Temp = Temp + Definir_ExpObjets(i)
'        End Select
'    Next i
'    Definir_Moy_Competence = Temp / (Nombre_CompetencesSpeciales + Nombre_CompetencesRessources + Nombre_CompetencesServices + Nombre_CompetencesObjets - 1)
'End Property
'Public Property Get Definir_Max_Competence(Optional ByVal Categorie As Integer) As Long
'    Dim i As Long
'    Dim Temp1, Temp2 As Long
'    For i = 0 To Nombre_CompetencesSpeciales - 1
'        Select Case Categorie
'        Case 0: Temp2 = Definir_Niveau_CompetenceSpeciales(i)
'        Case 1: Temp2 = Niveau_CompetenceSpeciales(i) + 1
'        Case 2: Temp2 = Definir_ExpSpeciales(i)
'        End Select
'        If Temp2 > Temp1 Then
'            Temp1 = Temp2
'        End If
'    Next i
'    For i = 0 To Nombre_CompetencesRessources - 1
'        Select Case Categorie
'        Case 0: Temp2 = Niveau_CompetenceRessources(i) + Parametres.PersosMaxRessources
'        Case 1: Temp2 = Niveau_CompetenceRessources(i) + 1
'        Case 2: Temp2 = Definir_ExpRessources(i)
'        End Select
'        If Temp2 > Temp1 Then
'            Temp1 = Temp2
'        End If
'    Next i
'    For i = 0 To Nombre_CompetencesServices - 1
'        Select Case Categorie
'        Case 0: Temp2 = Definir_Niveau_CompetenceServices(i)
'        Case 1: Temp2 = Niveau_CompetenceServices(i) + 1
'        Case 2: Temp2 = Definir_ExpServices(i)
'        End Select
'        If Temp2 > Temp1 Then
'            Temp1 = Temp2
'        End If
'    Next i
'    For i = 1 To Nombre_CompetencesObjets - 1
'        Select Case Categorie
'        Case 0: Temp2 = Definir_Niveau_CompetenceObjets(i)
'        Case 1: Temp2 = Niveau_CompetenceObjets(i) + 1
'        Case 2: Temp2 = Definir_ExpObjets(i)
'        End Select
'        If Temp2 > Temp1 Then
'            Temp1 = Temp2
'        End If
'    Next i
'    Definir_Max_Competence = Temp1
'End Property
'
'' ************************************************************************************
'' *                               Autres propriétes                                  *
'' ************************************************************************************
'
'
'Public Property Get Definir_Argent() As Double
'    Definir_Argent = Argent
'End Property
'
'Public Property Get Definir_DureeImmobilisation() As Double
'    Definir_DureeImmobilisation = DureeImmobilisation
'End Property
'Public Property Let Definir_DureeImmobilisation(ByVal Valeur As Double)
'    DureeImmobilisation = Valeur
'    RessourceCourante = -1
'End Property
'
Public Property Get Visible() As Boolean
    Visible = Persos(Numero).Visible
End Property
'
'Public Property Let Visible(ByVal Valeur As Boolean)
'    DansSaMaison = Not Valeur
'    DansUneMaison = Not Valeur
'    DansSonChateau = Not Valeur
'    DansUnChateau = Not Valeur
'End Property
'
'Public Property Get Attaquable() As Boolean
'    Attaquable = Vivant And Visible
'End Property
'
Public Property Get Vivant() As Boolean
    Vivant = Persos(Numero).Vivant
End Property
Public Property Get EnCombat() As Boolean
    EnCombat = Persos(Numero).EnCombat
End Property
''Public Property Get EnDeplacement() As Boolean
''    EnDeplacement = DirectionY < PositionY - Me.Longueur_Pas / 2 Or _
'                    DirectionY > PositionY + Me.Longueur_Pas / 2 Or _
'                    DirectionX < PositionX - Me.Longueur_Pas / 2 Or _
'                    DirectionX > PositionX + Me.Longueur_Pas / 2
''End Property
'
'Public Property Get Invocation() As Boolean
'    Invocation = NumeroChef >= 0
'End Property
Public Property Get Niveau() As Double
    Niveau = Persos(Numero).Niveau
End Property
'Public Property Get Defnir_Score_Tresor() As Integer
'    Defnir_Score_Tresor = (Niveau - 1) * 2 + Definir_Bestiaire_Total / 450
'End Property
Public Property Get Classe() As String
    Classe = Persos(Numero).Metier(Parametres)
End Property
'
''Propriété des objets.
'Public Property Get Nombre_Objets_Inventaire() As Integer
'    Nombre_Objets_Inventaire = UBound(ObjetsInventaires()) + 1
'End Property
'Public Property Get Nombre_Objets_Equipes() As Integer
'    Nombre_Objets_Equipes = UBound(ObjetsEquipes()) + 1
'End Property
'Public Property Get Objet_Inventaire_Apparence(ByVal Indice As Integer) As Integer
'    Objet_Inventaire_Apparence = ObjetsInventaires(Indice).Numero
'End Property
'Public Property Get Objet_Equipes_Apparence(ByVal Indice As Integer) As Integer
'    Objet_Equipes_Apparence = ObjetsEquipes(Indice).Numero
'End Property
'Public Property Get Objet_Inventaire_Actif(ByVal Indice As Integer) As Boolean
'    Objet_Inventaire_Actif = ObjetsInventaires(Indice).Actif
'End Property
'Public Property Let Objet_Inventaire_Actif(ByVal Indice As Integer, ByVal Valeur As Boolean)
'    ObjetsInventaires(Indice).Actif = Valeur
'End Property
'Public Property Get Objet_Inventaire_Categorie(ByVal Indice As Integer) As Integer
'    Objet_Inventaire_Categorie = ObjetsInventaires(Indice).Categorie
'End Property
'Public Property Let Objet_Inventaire_Categorie(ByVal Indice As Integer, ByVal Valeur As Integer)
'    ObjetsInventaires(Indice).Categorie = Valeur
'End Property
'Public Property Get Objet_Inventaire_Type(ByVal NumeroObjet As Integer) As Integer
'    Objet_Inventaire_Type = ObjetsInventaires(NumeroObjet).Numero
'End Property
'Public Property Let Objet_Inventaire_Type(ByVal NumeroObjet As Integer, ByVal Valeur As Integer)
'    ObjetsInventaires(NumeroObjet).Numero = Valeur
'End Property
'Public Property Get Objet_Equipes_Actif(ByVal Indice As Integer) As Boolean
'    Objet_Equipes_Actif = ObjetsEquipes(Indice).Actif
'End Property
'Public Property Let Objet_Equipes_Actif(ByVal Indice As Integer, ByVal Valeur As Boolean)
'    ObjetsEquipes(Indice).Actif = Valeur
'End Property
'Public Property Get Objet_Equipes_Categorie(ByVal Indice As Integer) As Integer
'    Objet_Equipes_Categorie = ObjetsEquipes(Indice).Categorie
'End Property
'Public Property Let Objet_Equipes_Categorie(ByVal Indice As Integer, ByVal Valeur As Integer)
'    ObjetsEquipes(Indice).Categorie = Valeur
'End Property
'Public Property Get Objet_Equipes_Type(ByVal NumeroObjet As Integer) As Integer
'    Objet_Equipes_Type = ObjetsEquipes(NumeroObjet).Numero
'End Property
'Public Property Let Objet_Equipes_Type(ByVal NumeroObjet As Integer, ByVal Valeur As Integer)
'    ObjetsEquipes(NumeroObjet).Numero = Valeur
'End Property
'Public Property Get Objet_Equipes_Competence(ByVal NumeroObjet As Integer) As Integer
'    Objet_Equipes_Competence = ObjetsEquipes(NumeroObjet).Competence
'End Property
'
''Propriétés si l'objet en question est indestructible.
'Public Property Get Objet_Equipes_Indestructible(ByVal NumeroObjet As Integer) As Boolean
'    Objet_Equipes_Indestructible = ObjetsEquipes(NumeroObjet).Indesctructible
'End Property
'Public Property Get Objet_Inventaire_Indestructible(ByVal NumeroObjet As Integer) As Boolean
'    Objet_Inventaire_Indestructible = ObjetsInventaires(NumeroObjet).Indesctructible
'End Property
'
'Public Property Get Objet_Equipes_2_Mains() As Boolean
'    If ObjetsEquipes(0).Categorie = 1 Then
'        Objet_Equipes_2_Mains = True
'    Else
'        Objet_Equipes_2_Mains = False
'    End If
'End Property
Public Property Get Dire_Commentaires() As String
    Dire_Commentaires = Persos(Numero).Definir_Commentaires
End Property
Public Property Let Dire_Commentaires(ByVal Valeur As String)
    Persos(Numero).Definir_Commentaires = Valeur
End Property
''Propriétés du bestiaire.
'Public Property Get Nombre_de_lignes_du_bestiaire() As Integer
'    Nombre_de_lignes_du_bestiaire = UBound(Bestiaire()) + 1
'End Property
'Public Property Get Definir_Bestiaire_Total() As Long
'    Dim i As Long
'    Dim X As Long
'    For i = 0 To Nombre_de_lignes_du_bestiaire - 1
'        X = X + Definir_Bestiaire(i)
'    Next i
'    Definir_Bestiaire_Total = X
'End Property
'Public Property Get Definir_Bestiaire(ByVal Index As Integer) As Long
'    Definir_Bestiaire = Bestiaire(Index)
'End Property
'Public Property Let Definir_Bestiaire(ByVal Index As Integer, ByVal Valeur As Long)
'    Bestiaire(Index) = Valeur
'End Property
'
''Propriétés des objets saisies.
'Public Property Get Definir_ObjetSelectionne() As Boolean
'    Definir_ObjetSelectionne = ArgentSelectionne > 0 Or ObjetSelectionne >= 0 Or RessourceSelectionne
'End Property
'Public Property Let Definir_ObjetSelectionne(ByVal Valeur As Boolean)
'    Dim i As Integer
'    If Not Valeur Then
'        ArgentSelectionne = 0
'        For i = 0 To UBound(RessourcesSelectionnes())
'            RessourcesSelectionnes(i) = 0
'        Next i
'        RessourceSelectionne = False
'        If ObjetSelectionne >= 0 Then
'            For i = 0 To Nombre_Objets_Inventaire - 1
'                If Not ObjetsInventaires(i).Actif Then
'                    ObjetsInventaires(i).Charger ObjetSelectionne, Parametres
'                    i = Nombre_Objets_Inventaire
'                End If
'            Next i
'            'Objet_Ajouter_Inventaire ObjetSelectionne, Parametres, Commentaires, False
'            ObjetSelectionne = -1
'        End If
'    End If
'End Property
'Public Property Get Definir_RessourcesSelectionnes(ByVal Index As Integer) As Double
'    Definir_RessourcesSelectionnes = RessourcesSelectionnes(Index)
'End Property
'Public Property Let Definir_RessourcesSelectionnes(ByVal Index As Integer, ByVal Valeur As Double)
'    RessourcesSelectionnes(Index) = Valeur
'End Property
'Public Sub Argent_Prendre(ByVal Quantite As Double, Optional ByVal Maximum As Boolean)
'    Quantite = Quantite + ArgentSelectionne
'    If Maximum Then
'        Quantite = Argent
'    ElseIf Argent < Quantite Then
'        Quantite = Argent
'    End If
'    If Quantite >= 1 Then
'        ArgentSelectionne = Quantite
'    End If
'End Sub
'Public Sub Ressources_Prendre(ByVal Index As Integer, ByVal Quantite As Double, Optional ByVal Maximum As Boolean)
'    If Maximum Then
'        Quantite = Ressources(Index)
'    ElseIf Ressources(Index) < Quantite Then
'        Quantite = Ressources(Index)
'    End If
'    If Quantite >= 1 Then
'        RessourcesSelectionnes(Index) = RessourcesSelectionnes(Index) + Quantite
'        RessourceSelectionne = True
'    End If
'End Sub
'Public Sub Ressources_Verifications()
'    'Vérifie qu'il reste au joueur suffisamment de ressources.
'    Dim i As Integer
'    If ArgentSelectionne > Argent Then ArgentSelectionne = Argent
'    RessourceSelectionne = False
'    For i = 0 To UBound(RessourcesSelectionnes())
'        If RessourcesSelectionnes(i) > Definir_Ressources(i) Then
'            RessourcesSelectionnes(i) = Definir_Ressources(i)
'            If Definir_Ressources(i) < 1 Then
'                RessourcesSelectionnes(i) = 0
'            End If
'        End If
'        If RessourcesSelectionnes(i) > 0 Then RessourceSelectionne = True
'    Next i
'End Sub
'Private Sub Ressources_Donner(ByRef PersoCible As ClsJeuPerso)
'    Dim i As Integer
'    Dim QuantiteMax As Double 'Définit la quantité maximale que l'on peut donner.
'
'    Ressources_Verifications
'
'    'Donne l'argent à la cible.
'    PersoCible.Definir_Argent = PersoCible.Definir_Argent + ArgentSelectionne
'    Definir_Argent = Definir_Argent - ArgentSelectionne
'
'    For i = 0 To UBound(RessourcesSelectionnes())
'        QuantiteMax = RessourcesSelectionnes(i)
'        If QuantiteMax > PersoCible.Definir_MaxRessources(i) - PersoCible.Definir_Ressources(i) Then
'            QuantiteMax = PersoCible.Definir_MaxRessources(i) - PersoCible.Definir_Ressources(i)
'        End If
'        'Donne les ressources à la cible.
'        PersoCible.Definir_Ressources(i) = PersoCible.Definir_Ressources(i) + QuantiteMax
'        'Retire les ressources au porteur.
'        Definir_Ressources(i) = Definir_Ressources(i) - QuantiteMax
'    Next i
'    Definir_ObjetSelectionne = False
'End Sub
'Public Sub Ressources_Deposer_Sol(Optional ByVal IndexTresor As Long = -1)
'    'Se débarasse des ressources sélectionnées.
'    Dim i As Integer
'
'    Ressources_Verifications
'
'    If IndexTresor = -1 Then
'        IndexTresor = Tresor_Creer(PositionX, PositionY)
'    End If
'    Tresors(IndexTresor).Definir_Argent = Tresors(IndexTresor).Definir_Argent + ArgentSelectionne
'    Definir_Argent = Definir_Argent - ArgentSelectionne
'    For i = 0 To UBound(RessourcesSelectionnes())
'        Tresors(IndexTresor).Definir_Ressources(i) = Tresors(IndexTresor).Definir_Ressources(i) + RessourcesSelectionnes(i)
'        Definir_Ressources(i) = Definir_Ressources(i) - RessourcesSelectionnes(i)
'    Next i
'    Definir_ObjetSelectionne = False
'End Sub
'
''Propriétés de la race.
'Public Property Get Definir_CoefRaceCompSpeciales(ByVal Index As Integer) As Single 'Coeficient de la race de la compétences spéciales.
'    Definir_CoefRaceCompSpeciales = CoefRaceCompSpeciales(Index)
'End Property
'Public Property Let Definir_CoefRaceCompSpeciales(ByVal Index As Integer, ByVal Valeur As Single) 'Coeficient de la race de la compétences spéciales.
'    CoefRaceCompSpeciales(Index) = Valeur
'End Property
'Public Property Get Definir_CoefRaceCompRessources(ByVal Index As Integer) As Single 'Coeficient de la race de la compétences des ressources.
'    Definir_CoefRaceCompRessources = CoefRaceCompRessources(Index)
'End Property
'Public Property Let Definir_CoefRaceCompRessources(ByVal Index As Integer, ByVal Valeur As Single) 'Coeficient de la race de la compétences des ressources.
'    CoefRaceCompRessources(Index) = Valeur
'End Property
'Public Property Get Definir_CoefRaceCompServices(ByVal Index As Integer) As Single 'Coeficient de la race de la compétences des Services.
'    Definir_CoefRaceCompServices = CoefRaceCompServices(Index)
'End Property
'Public Property Let Definir_CoefRaceCompServices(ByVal Index As Integer, ByVal Valeur As Single) 'Coeficient de la race de la compétences des Services.
'    CoefRaceCompServices(Index) = Valeur
'End Property
'Public Property Get Definir_CoefRaceCompObjets(ByVal Index As Integer) As Single 'Coeficient de la race de la compétence des objets.
'    Definir_CoefRaceCompObjets = CoefRaceCompObjets(Index)
'End Property
'Public Property Let Definir_CoefRaceCompObjets(ByVal Index As Integer, ByVal Valeur As Single) 'Coeficient de la race de la compétence des objets.
'    CoefRaceCompObjets(Index) = Valeur
'End Property
'
'
'
''Propriétés pour les apparences.
'Public Property Let Nombre_ApparenceBase(ByVal Valeur As Integer)
'    ReDim Preserve ApparenceBase(Valeur - 1)
'    ReDim Preserve ApparenceBaseAfficher(Valeur - 1)
'    ReDim Preserve FichierApparenceBase(Valeur - 1)
'    NombreApparenceBase = Valeur
'End Property
'Public Property Get Nombre_ApparenceBase() As Integer
'    Nombre_ApparenceBase = NombreApparenceBase
'End Property
'Public Property Let Definir_ApparenceBase(ByVal Index As Integer, _
'                                          ByVal Valeur As Integer)
'    ApparenceBase(Index) = Valeur
'End Property
'Public Property Get Definir_ApparenceBaseAfficher(ByVal Index As Integer) As Boolean
'    Definir_ApparenceBaseAfficher = ApparenceBaseAfficher(Index)
'End Property
'Public Property Let Definir_ApparenceBaseAfficher(ByVal Index As Integer, _
'                                                  ByVal Valeur As Boolean)
'    ApparenceBaseAfficher(Index) = Valeur
'End Property
'Public Property Get Definir_ApparenceBase(ByVal Index As Integer) As Integer
'    Definir_ApparenceBase = ApparenceBase(Index)
'End Property
'Public Property Get Definir_ApparenceObjet(ByVal Index As Integer) As Integer
'    Definir_ApparenceObjet = ApparenceObjet(Index)
'End Property
'Public Property Let Definir_FichierApparenceBase(ByVal Index As Integer, _
'                                                 ByVal Valeur As String)
'    FichierApparenceBase(Index) = Valeur
'End Property
'Public Property Get Definir_FichierApparenceBase(ByVal Index As Integer) As String
'    Definir_FichierApparenceBase = FichierApparenceBase(Index)
'End Property
'
'


'' ************************************************************************************
'' *                          Fonctions usant de l'IA                                 *
'' ************************************************************************************

Public Function Deplacement_Aleatoire() As Boolean
    IA_Deplacement_Aleatoire Numero
    Deplacement_Aleatoire = True
End Function
Public Function Comportement_Agressif() As Boolean
    If Persos(Numero).Vivant Then
        IA_Agressif Numero
        Comportement_Agressif = True
    Else
        Comportement_Agressif = False
    End If
End Function

'Public Sub IA_Garder_Batiment(ByVal i As Long, ByVal Chateau As Boolean)
'    Dim X As Long
'    If Chateau And Chateaux(Persos(i).IndiceChateau).Visible Or _
'       Not Chateau And Maisons(Persos(i).IndiceMaison).Visible Then
'        If Not Persos(i).EnCombat Then
'            If IA_Attaquer_Assaillant(i, Persos(i).IndiceMaison, Chateau, X) Then
'                Persos(i).Attaquer_Perso X
'            Else 'If Persos(i).Inactif Then
'                If Chateau Then
'                    With Chateaux(Persos(i).IndiceChateau)
'                    Select Case RND
'                    Case Is < 0.25: 'Va au coin en haut à gauche du château.
'                        Persos(i).Aller_A .PositionX - Persos(i).Largeur + Int(RND * (.Largeur + Persos(i).Largeur)), .PositionY - Persos(i).Hauteur
'                    Case Is < 0.5: 'Va au coin en haut à droite du château.
'                        Persos(i).Aller_A .PositionX + .Largeur, .PositionY - Persos(i).Hauteur + Int(RND * (.Hauteur + Persos(i).Hauteur))
'                    Case Is < 0.75: 'Va au coin en bas à droite du château.
'                        Persos(i).Aller_A .PositionX + .Largeur - Int(RND * (.Largeur + Persos(i).Largeur)), .PositionY + .Hauteur
'                    Case Else  'Va au coin en bas à gauche du château.
'                        Persos(i).Aller_A .PositionX - Persos(i).Largeur, .PositionY + .Hauteur - Int(RND * (.Hauteur + Persos(i).Hauteur))
'                    End Select
'                    End With
'                Else
'                    With Maisons(Persos(i).IndiceMaison)
'                    Select Case RND
'                    Case Is < 0.25: 'Va au coin en haut à gauche de la maison.
'                        Persos(i).Aller_A .PositionX - Persos(i).Largeur + Int(RND * (.Largeur + Persos(i).Largeur)), .PositionY - Persos(i).Hauteur
'                    Case Is < 0.5: 'Va au coin en haut à droite de la maison.
'                        Persos(i).Aller_A .PositionX + .Largeur, .PositionY - Persos(i).Hauteur + Int(RND * (.Hauteur + Persos(i).Hauteur))
'                    Case Is < 0.75: 'Va au coin en bas à droite de la maison.
'                        Persos(i).Aller_A .PositionX + .Largeur - Int(RND * (.Largeur + Persos(i).Largeur)), .PositionY + .Hauteur
'                    Case Else  'Va au coin en bas à gauche de la maison.
'                        Persos(i).Aller_A .PositionX - Persos(i).Largeur, .PositionY + .Hauteur - Int(RND * (.Hauteur + Persos(i).Hauteur))
'                    End Select
'                    End With
'                End If
'            End If
'        End If
'    Else
'        Persos(i).IA_Ordre = 0
'    End If
'End Sub
Public Function Comportement_Berserk() As Boolean
    If Persos(Numero).Vivant Then
        IA_ComportementBerserk Numero
        Comportement_Berserk = True
    Else
        Comportement_Berserk = False
    End If
End Function
'Private Function IA_Attaquer_Perso(ByVal NumeroIA As Long, ByRef Indice As Long) As Boolean
'    'Renvoie l'indice de persos le plus proche à attaquer.
'    Dim i As Long, lb As Long
'    lb = UBound(Persos())
'    Indice = -1
'    For i = 0 To NombreRechercheParTour
'        If IndiceAdversaire(NumeroIA) > lb Then
'            IndiceAdversaire(NumeroIA) = 0
'            DistanceAdversaire(NumeroIA) = DistanceAdversaire(NumeroIA) + DistancePalierRecherchePerso
'            If DistanceAdversaire(NumeroIA) > DistanceAdversaireMax Then
'                DistanceAdversaire(NumeroIA) = DistanceAdversaireMin
'            End If
'            'Exit Function
'        End If
'        'i = CLng(RND * UBound(Persos()))
'        If DistanceAdversaire(NumeroIA) > Sqr((Persos(NumeroIA).PositionX - Persos(IndiceAdversaire(NumeroIA)).PositionX) ^ 2 + _
'                                              (Persos(NumeroIA).PositionY - Persos(IndiceAdversaire(NumeroIA)).PositionY) ^ 2) And _
'           Persos(IndiceAdversaire(NumeroIA)).NumeroEquipe <> Persos(NumeroIA).NumeroEquipe And _
'           Persos(IndiceAdversaire(NumeroIA)).Attaquable Then
'            IA_Attaquer_Perso = True
'            Indice = IndiceAdversaire(NumeroIA)
'            i = NombreRechercheParTour
'        End If
'        IndiceAdversaire(NumeroIA) = IndiceAdversaire(NumeroIA) + 1
'    Next i
'End Function
'
'Private Function IA_Attaquer_Assaillant(ByVal NumeroIA As Long, ByVal NumeroBatiment As Long, ByVal Chateau As Boolean, ByRef Indice As Long) As Boolean
'    'Renvoie l'indice d'un perso qui attaque une maison.
'    Dim i As Long, lb As Long
'    lb = UBound(Persos())
'    Indice = -1
'    For i = 0 To NombreRechercheParTour
'        If IndiceAdversaire(NumeroIA) > lb Then
'            IndiceAdversaire(NumeroIA) = 0
'        End If
'        If Chateau Then
'            If Persos(IndiceAdversaire(NumeroIA)).Action = 302 And _
'               Persos(IndiceAdversaire(NumeroIA)).IndiceChateau = NumeroBatiment Then
'                IA_Attaquer_Assaillant = True
'                Indice = IndiceAdversaire(NumeroIA)
'                i = NombreRechercheParTour
'            End If
'        Else
'            If Persos(IndiceAdversaire(NumeroIA)).Action = 301 And _
'               Persos(IndiceAdversaire(NumeroIA)).IndiceMaison = NumeroBatiment Then
'                IA_Attaquer_Assaillant = True
'                Indice = IndiceAdversaire(NumeroIA)
'                i = NombreRechercheParTour
'            End If
'        End If
'        IndiceAdversaire(NumeroIA) = IndiceAdversaire(NumeroIA) + 1
'    Next i
'End Function
'
'Private Function IA_Attaquer_Maison(ByVal NumeroIA As Long, ByRef Indice As Long) As Boolean
'    'Renvoie l'indice de persos le plus proche à attaquer.
'    Dim i As Long, lb As Long
'    Dim pIAx, pIAy
'    pIAx = Persos(NumeroIA).PositionX
'    pIAy = Persos(NumeroIA).PositionY
'
'    lb = UBound(Maisons())
'    Indice = -1
'    For i = 0 To NombreRechercheParTour
'        If IndiceMaisonEnnemi(NumeroIA) > lb Then
'            IndiceMaisonEnnemi(NumeroIA) = 0
'            DistanceMaisonEnnemi(NumeroIA) = DistanceMaisonEnnemi(NumeroIA) + DistancePalierRechercheMaison
'            If DistanceMaisonEnnemi(NumeroIA) > DistanceAdversaireMax Then
'                DistanceMaisonEnnemi(NumeroIA) = DistanceAdversaireMin
'            End If
'            'Exit Function
'        End If
'        'i = CLng(RND * UBound(Persos()))
'        If DistanceMaisonEnnemi(NumeroIA) > Sqr((pIAx - Maisons(IndiceMaisonEnnemi(NumeroIA)).PositionX) ^ 2 + _
'                                                 (pIAy - Maisons(IndiceMaisonEnnemi(NumeroIA)).PositionY) ^ 2) And _
'           Persos(IndiceMaisonEnnemi(NumeroIA)).NumeroEquipe <> Persos(NumeroIA).NumeroEquipe And _
'           Maisons(IndiceMaisonEnnemi(NumeroIA)).Visible Then
'            IA_Attaquer_Maison = True
'            Indice = IndiceMaisonEnnemi(NumeroIA)
'            i = NombreRechercheParTour
'        End If
'        IndiceMaisonEnnemi(NumeroIA) = IndiceMaisonEnnemi(NumeroIA) + 1
'    Next i
'End Function
'
'Private Function IA_Attaquer_Chateau(ByVal NumeroIA As Long, ByRef Indice As Long) As Boolean
'    'Renvoie l'indice de persos le plus proche à attaquer.
'    Dim i As Long, pIAx, pIAy, uChato
'    pIAx = Persos(NumeroIA).PositionX
'    pIAy = Persos(NumeroIA).PositionY
'    uChato = UBound(Chateaux())
'    Indice = -1
'    For i = 0 To NombreRechercheParTour
'        If IndiceChateauEnnemi(NumeroIA) > uChato Then
'            IndiceChateauEnnemi(NumeroIA) = 0
'            DistanceChateauEnnemi(NumeroIA) = DistanceChateauEnnemi(NumeroIA) + Parametres.Peuples_ChateauLargeur(0) * 2
'            If DistanceChateauEnnemi(NumeroIA) > DistanceAdversaireMax Then
'                DistanceChateauEnnemi(NumeroIA) = DistanceAdversaireMin
'            End If
'            'Exit Function
'        End If
'        If DistanceChateauEnnemi(NumeroIA) > Sqr((pIAx - Chateaux(IndiceChateauEnnemi(NumeroIA)).PositionX) ^ 2 + _
'                                                 (pIAy - Chateaux(IndiceChateauEnnemi(NumeroIA)).PositionY) ^ 2) And _
'           (Fiefs(IndiceChateauEnnemi(NumeroIA)).NumeroEquipe <> Persos(NumeroIA).NumeroEquipe And _
'           Chateaux(IndiceChateauEnnemi(NumeroIA)).Visible) Then
'            IA_Attaquer_Chateau = True
'            Indice = IndiceChateauEnnemi(NumeroIA)
'            i = NombreRechercheParTour
'        End If
'        IndiceChateauEnnemi(NumeroIA) = IndiceChateauEnnemi(NumeroIA) + 1
'    Next i
'End Function
Public Function Attaquer_Perso_Proche(Optional ByVal EnvoyerSoldats As Boolean, _
                                      Optional ByVal Ctrl As Boolean, _
                                      Optional ByVal Shift As Boolean, _
                                      Optional ByVal Alt As Boolean) As Boolean
    If Trouver_Ennemi(Numero) >= 0 And Persos(Numero).Vivant Then
        IA_Trouver_Ennemi Numero, EnvoyerSoldats, Ctrl, Shift, Alt
        Attaquer_Perso_Proche = True
    Else
        Attaquer_Perso_Proche = False
    End If
End Function
'
'Public Function Trouver_Ennemi(ByVal i As Long) As Long
'    'Renvoie l'indice de l'ennemi le plus proche du personnage n°i.
'    Dim j As Long
'    Dim Distance As Long
'    Trouver_Ennemi = -1
'    Distance = Monde.Largeur * AffTerrain.Largeur + Monde.Hauteur * AffTerrain.Hauteur
'    For j = 1 To UBound(Persos())
'        If Persos(j).NumeroEquipe <> Persos(i).NumeroEquipe And _
'           Persos(j).Attaquable And _
'           Distance > Sqr((Persos(i).PositionX - Persos(j).PositionX) ^ 2 + _
'                          (Persos(i).PositionY - Persos(j).PositionY) ^ 2) Then
'            Distance = Sqr((Persos(i).PositionX - Persos(j).PositionX) ^ 2 + _
'                           (Persos(i).PositionY - Persos(j).PositionY) ^ 2)
'            Trouver_Ennemi = j
'        End If
'    Next j
'End Function
'
'Public Sub Trouver_Maison(ByRef Perso As ClsJeuPerso, _
'                          ByVal TypeBatiment As Integer)
'    'Le personnage se rend à la maison alliée du type indiqué en paramètres.
'    Dim i, j As Long
'    j = Perso.IndiceMaison + 1
'    For i = 0 To UBound(Maisons())
'        If j > UBound(Maisons()) Then
'            j = 0
'        End If
'        If Perso.NumeroEquipe = Persos(j).NumeroEquipe And _
'           Maisons(j).TypeBatiment = TypeBatiment And _
'           Maisons(j).Visible Then
'            Perso.Aller_Maison Maisons(j)
'            i = UBound(Maisons())
'        Else
'            j = j + 1
'        End If
'    Next i
'End Sub
'
'
Public Function Attaquer_Maison_Proche(Optional ByVal EnvoyerSoldats As Boolean, _
                                       Optional ByVal Ctrl As Boolean, _
                                       Optional ByVal Shift As Boolean, _
                                       Optional ByVal Alt As Boolean) As Boolean
    If Trouver_MaisonEnnemi(Numero) >= 0 And Persos(Numero).Vivant Then
        IA_Trouver_MaisonEnnemi Numero, EnvoyerSoldats, Ctrl, Shift, Alt
        Attaquer_Maison_Proche = True
    Else
        Attaquer_Maison_Proche = False
    End If
End Function
'
'Public Function Trouver_MaisonEnnemi(ByVal i As Long) As Long
'    'Renvoie l'indice de la maison ennemi la plus proche du personnage n°i.
'    Dim j As Long
'    Dim Distance As Long
'    Trouver_MaisonEnnemi = -1
'    Distance = Monde.Largeur * AffTerrain.Largeur + Monde.Hauteur * AffTerrain.Hauteur
'    For j = 0 To UBound(Maisons())
'        If Persos(j).NumeroEquipe <> Persos(i).NumeroEquipe And _
'           Maisons(j).Visible And _
'           Distance > Sqr((Persos(i).PositionX - Maisons(j).PositionX) ^ 2 + _
'                          (Persos(i).PositionY - Maisons(j).PositionY) ^ 2) Then
'            Distance = Sqr((Persos(i).PositionX - Maisons(j).PositionX) ^ 2 + _
'                           (Persos(i).PositionY - Maisons(j).PositionY) ^ 2)
'            Trouver_MaisonEnnemi = j
'            'j = UBound(Maisons())
'        End If
'    Next j
'End Function
'
Public Function Attaquer_Chateau_Proche(Optional ByVal EnvoyerSoldats As Boolean, _
                                        Optional ByVal Ctrl As Boolean, _
                                        Optional ByVal Shift As Boolean, _
                                        Optional ByVal Alt As Boolean) As Boolean
    If Trouver_ChateauEnnemi(Numero) >= 0 And Persos(Numero).Vivant Then
        IA_Trouver_ChateauEnnemi Numero, EnvoyerSoldats, Ctrl, Shift, Alt
        Attaquer_Chateau_Proche = True
    Else
        Attaquer_Chateau_Proche = False
    End If
End Function
'
'Public Function Trouver_ChateauEnnemi(ByVal i As Long) As Long
'    'Renvoie l'indice du château ennemi le plus proche du personnage n°i.
'    Dim j As Long
'    Dim Distance As Long
'    Trouver_ChateauEnnemi = -1
'    Distance = Monde.Largeur * AffTerrain.Largeur + Monde.Hauteur * AffTerrain.Hauteur
'    For j = 0 To UBound(Chateaux())
'        If Persos(i).NumeroEquipe <> Fiefs(j).NumeroEquipe And _
'           Chateaux(j).Visible And _
'           Distance > Sqr((Persos(i).PositionX - Chateaux(j).PositionX) ^ 2 + _
'                          (Persos(i).PositionY - Chateaux(j).PositionY) ^ 2) Then
'            Distance = Sqr((Persos(i).PositionX - Chateaux(j).PositionX) ^ 2 + _
'                           (Persos(i).PositionY - Chateaux(j).PositionY) ^ 2)
'            Trouver_ChateauEnnemi = j
'            'j = UBound(Chateaux())
'        End If
'    Next j
'End Function
'
'Public Sub IA_Poursuivre_Fuyard(ByVal i As Long)
'    'Le fuyard qui était dans la maison ressort.
'    With Persos(i)
'    If (.Action = 301 Or .Action = 302) And _
'        Persos(.IndicePerso).Attaquable And _
'       .NumeroEquipe <> Persos(.IndicePerso).NumeroEquipe Then
'             .Attaquer_Perso .IndicePerso
'             .Definir_Commentaires = Commentaires.Message(16, Persos(i))
'    End If
'    End With
'End Sub
'
'Public Sub IA_Recuperation_Carac(ByVal i As Long)
'    With Persos(i)
'    If .IA_Manger Then
'    'Le personnage est blessé.
'    'Mange ce qu'il a.
'    .Recuperer_AutoVie Parametres
'    'N'a rien pour récuperer ses points de vie.
'    'Va en acheter.
'    If Not .EnCombat And _
'       Not (.DansUneMaison) And _
'       .Inactif And _
'       .Definir_Vie < .Definir_MaxVie Then
'        'S'il na pas d'argent, il va chercher des herbes médicinales.
'        If IA_Argent_Besoin(i) Then
'            IA_Ramasser_Herbe i
'        Else
'            If Trouver_Maison_Proche(Persos(i), i, 0, 1, 0, X) Then
'                If X >= 0 Then
'                    IA_Aller_Maison i, X
'                End If
'            Else
'                IA_Ramasser_Herbe i
'            End If
'        End If
'    End If
'
'    'Le personnage a faim.
'    'Mange ce qu'il a.
'    .Recuperer_AutoEnergie Parametres
'    'N'a rien a manger.
'    If Not .EnCombat And _
'       Not (.DansUneMaison) And _
'       .Definir_Energie < .Definir_MaxEnergie * 0.5 Then
'        'S'il na pas d'argent, il va chercher les fruits sur un arbre
'        If IA_Argent_Besoin(i) Then
'            IA_Cueillir_Fruits i
'        Else
'        'Sinon, va acheter à manger.
'            If Trouver_Maison_Proche(Persos(i), i, 0, 2, 0, X) Then
'                If X >= 0 Then
'                    IA_Aller_Maison i, X
'                End If
'            Else
'                IA_Cueillir_Fruits i
'            End If
'        End If
'    End If
'
'    'Le personnage manque de magie.
'    'Utilise ce qu'il a.
'    .Recuperer_AutoMagie Parametres
'    'N'a rien pour récupérer de la magie.
'    'Va en acheter.
'    If Not .EnCombat And _
'       Not (.DansUneMaison) And _
'       .Definir_Magie < .Definir_MaxMagie * 0.5 And _
'       Trouver_Maison_Proche(Persos(i), i, 0, 3, 0, X) And _
'       Not IA_Argent_Besoin(i) Then
'        If X >= 0 Then
'            IA_Aller_Maison i, X
'        End If
'    End If
'
'    'Le personnage est démoralisé.
'    'Utilise ce qu'il a.
'    .Recuperer_AutoMoral Parametres
'    'N'a rien pour retrouver le moral.
'    'Va en acheter.
'    If Not .EnCombat And _
'       Not (.DansUneMaison) And _
'       .Definir_Moral < .Definir_MaxMoral * 0.5 And _
'       Trouver_Maison_Proche(Persos(i), i, 0, 4, 0, X) And _
'       Not IA_Argent_Besoin(i) Then
'        If X >= 0 Then
'            IA_Aller_Maison i, X
'        End If
'    End If
'    End If
'
'    'N'a plus de ressources pour attaquer.
'    'Va en acheter.
'    If Not (.DansUneMaison) And _
'       .Inactif And _
'       .Definir_Ressources(.PerteRessourceIndice) < .PerteRessourceQuantite Then
'        'S'il na pas d'argent, il va rammasser des munitions.
'        If IA_Argent_Besoin(i) Then
'            IA_Chercher_Ressource i, .PerteRessourceIndice
'        Else
'            If Trouver_Maison_Proche(Persos(i), i, .PerteRessourceIndice, 0, 0, X) Then
'                If X >= 0 Then
'                    IA_Aller_Maison i, X
'                End If
'            Else
'                IA_Chercher_Ressource i, .PerteRessourceIndice
'            End If
'        End If
'    End If
'
'    'Va chercher un trésor.
'    If .Inactif And .IA_Equiper>0 Then
'        If Trouver_Tresor_Proche(Persos(i), i, X) Then
'            If X > 0 Then
'                IA_Aller_Tresor i, X
'            End If
'        End If
'    End If
'
'    End With
'End Sub
'
'Public Sub IA_Depenser_Argent(ByVal i As Long)
'    'L'IA dépense son surplus d'argent.
'    With Persos(i)
'    If .Inactif And _
'       Not .DansUneMaison And _
'       .Definir_Vie = .Definir_MaxVie And _
'       .Definir_Energie >= .Definir_MaxEnergie * 0.5 And _
'       .Definir_Moral >= .Definir_MaxMoral * 0.5 And _
'       IA_Argent_Riche(i) And _
'       Not .Objet_Inventaire_Actif(.Nombre_Objets_Inventaire - 1) And _
'       Trouver_Maison_Proche(Persos(i), i, 0, 5, 0, X) Then
'        If X >= 0 Then
'            IA_Aller_Maison i, X
'        End If
'    End If
'    End With
'End Sub
'
'Private Sub IA_Ramasser_Herbe(ByVal i As Long)
'    If IndiceHerbe >= 0 Then
'        IA_Chercher_Ressource i, IndiceHerbe
'    End If
'End Sub
'Private Sub IA_Cueillir_Fruits(ByVal i As Long)
'    If IndiceFruit >= 0 Then
'        IA_Chercher_Ressource i, IndiceFruit
'    End If
'End Sub
'Private Sub IA_Couper_Arbre(ByVal i As Long)
'    If IndiceBois >= 0 Then
'        IA_Chercher_Ressource i, IndiceBois
'    End If
'End Sub
'Private Sub IA_ChercherPierre(ByVal i As Long)
'    If IndicePierre >= 0 Then
'        IA_Chercher_Ressource i, IndicePierre
'    End If
'End Sub
'
'Public Sub IA_ChercherMatieresPremiere(ByVal i As Long)
'    With Persos(i)
'    If (.Inactif Or .Action = 202) And _
'                       Not .DansUneMaison And _
'                       Maisons(i).Construit And _
'                       .Definir_Ressources(Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)) < MatierePremiereMin And _
'                       Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment) >= 0 Then
'        'Va récolter ses matières premières.
'        If Parametres.Ressources_Recoltable(Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)) Then
'            IA_Chercher_Ressource i, Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)
'        Else
'            'Va en acheter.
'            If .Definir_Argent > 1 Then
'                If Trouver_Maison_Proche(Persos(i), i, Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment), 0, 0, X) Then
'                    If X >= 0 Then
'                        IA_Aller_Maison i, X
'                    End If
'                'ElseIf Not .DansSonChateau And _
'                       Not .DansSaMaison And _
'                       Not .DansUneMaison And _
'                       Not .Action = 201 And _
'                       Fiefs(.NumeroFief).Definir_Ressources(Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)) > 0 And _
'                       Chateaux(.NumeroFief).Visible Then
'                ElseIf QuitterBatiment(i) And _
'                       Not .Action = 201 And _
'                       Fiefs(.NumeroFief).Definir_Ressources(Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)) > MatierePremiereMin And _
'                       Chateaux(.NumeroFief).Visible Then
'                    IA_Rentrer_Chateau i
'                End If
'            End If
'        End If
'    End If
'    End With
'End Sub
'
'Private Sub IA_Chercher_Ressource_Pour_Construire(ByVal i As Long)
'    Dim j As Long
'    If BatimentAConstruire(i) >= 0 Then
'        For j = 0 To Parametres.NombreRessources - 1
'            If Parametres.Ressources_Recoltable(j) And _
'               Persos(i).Definir_Ressources(j) < Parametres.Batiment_Prix(BatimentAConstruire(i), j) Then
'                IA_Chercher_Ressource i, j
'                j = Parametres.NombreRessources
'            End If
'        Next j
'    End If
'End Sub
'Public Sub IA_Chercher_Ressource(ByVal i As Long, ByVal TypeRessource As Long)
'    Dim j As Long
'    'If Trouver_Ressource_Proche(Persos(i), i, TypeRessource, j) Then
'    '    Persos(i).Aller_Ressource Ressources(j), j, AffRessources.Largeur, AffRessources.Hauteur
'    'End If
'    IA_Trouver_Ressource i, Parametres.Ressources_TypeTerrainRessource(TypeRessource)
'End Sub

Public Function Aller_Ressource(ByVal TypeRessource As Long) As Boolean
    If Persos(Numero).Vivant Then
        IA_Trouver_Ressource Numero, TypeRessource
        If Action = 3 And _
           Ressources(Persos(Numero).IndiceRessource).Apparence = TypeRessource Then
            Aller_Ressource = True
        Else
            Aller_Ressource = False
        End If
    End If
End Function

'Private Sub IA_Acheter_Ressources_Pour_Travailler(ByVal i As Long)
'    Dim j As Integer
'    Dim k As Integer
'    Dim Mtb As Long
'    Mtb = Maisons(i).TypeBatiment
'    'If i = 0 Then Debug.Print Maisons(i).Visible & " : " & Parametres.Batiment_Fabrique(Mtb) & " : " & Parametres.Batiment_TypeMatierePremiere(Mtb)
'    If Maisons(i).Visible Then
'        If Parametres.Batiment_Fabrique(Mtb) Then
'            'Achête les ressources pour tous les objets qu'il fabrique.
'            For j = 0 To Parametres.NombreRessources - 1
'                If Not (Parametres.Ressources_Recoltable(j)) Then
'                    For k = 0 To Parametres.Batiment_NombreObjetsFabriques(Mtb) - 1
'                        If Parametres.Objet_PrixRessources(Parametres.Batiment_ObjetFabrique(Mtb, k), j) > 0 Then
'                            Fiefs(Persos(i).NumeroFief).Acheter j, 0, Persos(i), Parametres, True
'                            k = Parametres.Batiment_NombreObjetsFabriques(Mtb)
'                        End If
'                    Next k
'                End If
'            Next j
'        ElseIf Parametres.Batiment_TypeMatierePremiere(Mtb) >= 0 Then
'            If Not (Parametres.Ressources_Recoltable(Parametres.Batiment_TypeMatierePremiere(Mtb))) Then
'                Fiefs(Persos(i).NumeroFief).Acheter Parametres.Batiment_TypeMatierePremiere(Mtb), 0, Persos(i), Parametres, True
'            End If
'        End If
'    End If
'End Sub
'Private Sub IA_Acheter_Ressources_Pour_Construire(ByVal i As Long)
'    Dim j As Long
'    If BatimentAConstruire(i) >= 0 Then
'        For j = 0 To Parametres.NombreRessources - 1
'            'If Parametres.Ressources_NecessairePourConstruire(j) And _
'               j <> Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment) Then
'               'Fiefs(Persos(i).NumeroFief).Acheter j, QuantiteMaxTransfert * RND, Persos(i), Parametres
'            If Parametres.Batiment_Prix(BatimentAConstruire(i), j) > Persos(i).Definir_Ressources(j) And _
'               j <> Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment) And _
'               Not Parametres.Ressources_Recoltable(j) Then
'                Fiefs(Persos(i).NumeroFief).Acheter j, QuantiteMaxTransfert * RND, Persos(i), Parametres
'            End If
'        Next j
'    End If
'End Sub
'
'Public Sub IA_Equiper(ByVal i As Long)
'    Dim j As Long
'    With Persos(i)
'    If .IA_Equiper Then
'        For j = 0 To IIf(.EnCombat, VitesseEquipementCombat, VitesseEquipement)
'            If .Objet_Inventaire_Actif(IndiceObjet) Then
'                If Parametres.Objet_Utilisable(.Objet_Inventaire_Type(IndiceObjet)) Then
'                    'Objets utilisables.
'                    If Parametres.Objet_UtilisationUnique(.Objet_Inventaire_Type(IndiceObjet)) Then
'                        If .EnCombat Then
'                            'N'utilise pas les invocations hors-combat.
'                            .Objet_Equiper IndiceObjet, Parametres, Commentaires
'                        End If
'                    Else
'                        .Objet_Equiper IndiceObjet, Parametres, Commentaires
'                    End If
'                Else
'                    Select Case Parametres.Objet_Categorie(.Objet_Inventaire_Type(IndiceObjet))
'                    Case 0: 'Arme à une main.
'                        If Not .Objet_Equipes_Actif(0) Or Not .Objet_Equipes_Actif(1) Then
'                            .Objet_Equiper IndiceObjet, Parametres, Commentaires
'                        Else
'                            If Parametres.Objet_PrixVente(.Objet_Inventaire_Type(IndiceObjet)) _
'                               >= Parametres.Objet_PrixVente(.Objet_Equipes_Type(0)) Then
'                                .Objet_Equiper IndiceObjet, Parametres, Commentaires
'                            End If
'                        End If
'                    Case 1: 'Arme à 2 mains.
'                        If Not .Objet_Equipes_Actif(0) Then
'                            .Objet_Equiper IndiceObjet, Parametres, Commentaires
'                        Else
'                            If Parametres.Objet_PrixVente(.Objet_Inventaire_Type(IndiceObjet)) _
'                               >= Parametres.Objet_PrixVente(.Objet_Equipes_Type(0)) Then
'                                .Objet_Equiper IndiceObjet, Parametres, Commentaires
'                            End If
'                        End If
'                    Case Else 'Autre équipement.
'                        If Not .Objet_Equipes_Actif(Parametres.Objet_Categorie(.Objet_Inventaire_Type(IndiceObjet))) Then
'                            .Objet_Equiper IndiceObjet, Parametres, Commentaires
'                        Else
'                            If Parametres.Objet_PrixVente(.Objet_Inventaire_Type(IndiceObjet)) _
'                               >= Parametres.Objet_PrixVente(.Objet_Equipes_Type(Parametres.Objet_Categorie(.Objet_Inventaire_Type(IndiceObjet)))) Then
'                                .Objet_Equiper IndiceObjet, Parametres, Commentaires
'                            End If
'                        End If
'                    End Select
'                End If
'            End If
'
'            IndiceObjet = IndiceObjet + 1
'            If IndiceObjet > Persos(i).Nombre_Objets_Inventaire - 1 Then
'                IndiceObjet = 0
'            End If
'        Next j
'    End If
'    End With
'End Sub
'
'Public Sub IA_Mort_Equiper(ByVal i As Long)
'    Dim j As Long
'    With Persos(i)
'    If .IA_Equiper Then
'        For j = 0 To VitesseEquipement
'            If .Objet_Inventaire_Actif(IndiceObjet) Then
'                If Parametres.Objet_UtilisationMort(.Objet_Inventaire_Type(IndiceObjet)) Then
'                    'Objets utilisables quand il est mort.
'                    .Objet_Equiper IndiceObjet, Parametres, Commentaires
'                End If
'            End If
'            IndiceObjet = IndiceObjet + 1
'            If IndiceObjet > Persos(i).Nombre_Objets_Inventaire - 1 Then
'                IndiceObjet = 0
'            End If
'        Next j
'    End If
'    End With
'End Sub
'
'Private Function Nombre_Total_Batiments(ByVal NumeroFief As Long)
'    Dim i As Long
'    With Persos(i)
'    For i = 0 To UBound(Maisons())
'        If .NumeroFief = NumeroFief Then
'            Nombre_Total_Batiments = Nombre_Total_Batiments + 1
'        End If
'    Next i
'    End With
'End Function
'
'Private Function Nombre_Batiments(ByVal TypeBatiment As Long, ByVal NumeroFief As Long)
'    Dim i As Long
'    If TypeBatiment >= 0 Then
'        For i = 0 To UBound(Maisons())
'            'If Maisons(i).TypeBatiment = TypeBatiment And _
'            '   Persos(i).NumeroFief = NumeroFief Then
'            '        Nombre_Batiments = Nombre_Batiments + 1
'            'End If
'            Nombre_Batiments = Nombre_Batiments + _
'                               Abs(Maisons(i).TypeBatiment = TypeBatiment And _
'                               Persos(i).NumeroFief = NumeroFief)
'        Next i
'    End If
'End Function
'
'Public Function Choisir_Batiment(ByVal NumeroPerso As Long, _
'                                 ByVal NumeroFief As Long, _
'                                 Optional ByVal Niveau As Long = -1) As Integer
'    Dim i As Long
'    Dim j As Long
'    Dim k As Long
'    Choisir_Batiment = -1
'
'    'Cherche une maison de niveau supérieur.
'    If Niveau = -1 Then
'        Niveau = 0
'        For i = Int(RND * BatimentNiveauMax) + 1 To Parametres.Batiment_Niveau(Maisons(NumeroPerso).TypeBatiment) + 1 Step -1
'            If Choisir_Batiment = -1 Then
'                Choisir_Batiment = Choisir_Batiment(NumeroPerso, NumeroFief, i)
'                If Persos(NumeroPerso).PeutConstruire_Maison(Choisir_Batiment, Parametres) Then
'                    i = 0
'                Else
'                    Choisir_Batiment = -1
'                End If
'            End If
'        Next i
'    End If
'
'    'Cherche s'il dispose du monopole sur un batiment de niveau inférieur que le sien.
'    If Nombre_Batiments(Maisons(NumeroPerso).TypeBatiment, NumeroFief) > 1 Then 'Sauf s'il a déjà le monopole.
'        For i = 0 To Parametres.NombreTypeBatiments - 1
'            If Nombre_Batiments(i, NumeroFief) = 0 And _
'               Parametres.Batiment_Niveau(i) < Parametres.Batiment_Niveau(Maisons(NumeroPerso).TypeBatiment) Then
'            'If Nombre_Batiments(i, NumeroFief) = 0 Then
'                Choisir_Batiment = i
'                i = Parametres.NombreTypeBatiments
'            End If
'        Next i
'    End If
'    If Choisir_Batiment = -1 Then
'        For i = 0 To Parametres.NombreTypeBatiments - 1
'            If Parametres.Batiment_Niveau(i) = Niveau Then
'                k = k + 1
'            End If
'        Next i
'        k = Int(RND * k)
'        For i = 0 To Parametres.NombreTypeBatiments - 1
'            If Parametres.Batiment_Niveau(i) = Niveau Then
'                If j = k Then
'                    Choisir_Batiment = i
'                    i = Parametres.NombreTypeBatiments
'                Else
'                    j = j + 1
'                End If
'            End If
'        Next i
'    ElseIf Parametres.Batiment_NoEpoque(Choisir_Batiment) - 1 > Fiefs(NumeroFief).Epoque And _
'           Persos(NumeroPerso).PeutConstruire_Maison(Choisir_Batiment, Parametres) And _
'           Not Maisons(NumeroPerso).Visible Then
'        'Si l'IA n'a pas de maison et ne peut pas la fabriquer tout de suite, il en choisit une de niveau 0.
'        Choisir_Batiment = Choisir_Batiment(NumeroPerso, NumeroFief, 0)
'    End If
'End Function
'
'Private Function Trouver_Ressource_Proche(ByRef Personnage As ClsJeuPerso, _
'                                          ByVal NumeroIA As Long, _
'                                          ByVal TypeRessource As Long, _
'                                          ByRef Indice As Long) As Boolean
'    'X = CLng(RND * UBound(Ressources()))
'    Dim i As Long
'    'Dim j As Long
'    'For j = 0 To Monde.Longueur_Diagonale Step 64
'    DistanceRessource(NumeroIA) = DistanceRessource(NumeroIA) + AffRessources.Largeur * RND * 10
'    For i = 0 To UBound(Ressources())
'        'If IA_IndiceRessource(NumeroIA) > UBound(Ressources()) Then
'        '    DistanceRessource(NumeroIA) = DistanceRessource(NumeroIA) + AffRessources.Largeur * 2
'        '    IA_IndiceRessource(NumeroIA) = 0
'        'End If
'        If Ressources(i).TypeRessource = TypeRessource And _
'            DistanceRessource(NumeroIA) >= Sqr((Personnage.PositionX - Ressources(i).PositionX) ^ 2 + _
'                                               (Personnage.PositionY - Ressources(i).PositionY) ^ 2) Then
'            Indice = i
'            DistanceRessource(NumeroIA) = 0
'            'IA_IndiceRessource(NumeroIA) = 0
'            Trouver_Ressource_Proche = True
'            'Exit Function
'            i = UBound(Ressources())
'            'j = Monde.Longueur_Diagonale
'        'Else
'        '    IA_IndiceRessource(NumeroIA) = IA_IndiceRessource(NumeroIA) + 1
'        End If
'    'Debug.Print DistanceRessource & " / " & IA_IndiceRessource(NumeroIA)
'    Next i
'    'Next j
'End Function
'
'Private Function Trouver_Tresor_Proche(ByRef Perso As ClsJeuPerso, _
'                                       ByVal NumeroIA As Long, _
'                                       ByRef Indice As Long) As Boolean
'    'Recherche un trésor.
'    Dim i As Long
'    DistanceTresor(NumeroIA) = DistanceTresor(NumeroIA) + AffTresor.Largeur * RND * 10
'    If DistanceTresor(NumeroIA) > DistanceTresorMax Then
'        DistanceTresor(NumeroIA) = 0
'    Else
'        For i = 1 To Tresors.Count
'            If DistanceTresor(NumeroIA) >= Sqr((Perso.PositionX - Tresors(i).X) ^ 2 + _
'                                               (Perso.PositionY - Tresors(i).Y) ^ 2) Then
'                Indice = i
'                DistanceTresor(NumeroIA) = 0
'                Trouver_Tresor_Proche = True
'                i = Tresors.Count
'            End If
'        Next i
'    End If
'End Function
'
Public Function Ramasser_Tresor_Proche() As Boolean
    If Persos(Numero).Vivant And Tresors.Count > 0 Then
        IA_Trouver_Tresor Numero
        Ramasser_Tresor_Proche = True
    Else
        Ramasser_Tresor_Proche = False
    End If
End Function
'
'Public Function Trouver_Maison_Proche(ByRef Personnage As ClsJeuPerso, _
'                                      ByVal NumeroIA As Long, _
'                                      ByVal TypeRessource As Integer, _
'                                      ByVal TypeBesoin As Integer, _
'                                      ByVal TypeService As Integer, _
'                                      ByRef Indice As Long) _
'                                      As Boolean
'    Dim i As Long
'    Dim pbm As Long
'    Dim IMnia As Long, ubMaison As Long
'
'    Indice = -1
'    For i = 0 To NombreRechercheParTour
'        If IndiceMaison(NumeroIA) > UBound(Maisons()) Then
'           DistanceMaison(NumeroIA) = DistanceMaison(NumeroIA) + DistancePalierRechercheMaison
'            IndiceMaison(NumeroIA) = 0
'        End If
'
'        'Recherche les batiments de service.
'        With Maisons(IndiceMaison(NumeroIA))
'        If TypeService > 0 Then
'            'If .Service = TypeService And .Construit And .Definir_Stock >= 1 And _
'            '   Persos(IndiceMaison(NumeroIA)).NumeroFief = Persos(NumeroIA).NumeroFief And _
'            '   Abs(Personnage.PositionX - .PositionX) < DistanceMaison(NumeroIA) And _
'            '   Abs(Personnage.PositionY - .PositionY) < DistanceMaison(NumeroIA) Then
'            '    Trouver_Maison_Proche = True
'            'End If
'            Trouver_Maison_Proche = (.Service = TypeService And .Construit And .Definir_Stock >= 1 And _
'               Persos(IndiceMaison(NumeroIA)).NumeroFief = Persos(NumeroIA).NumeroFief And _
'               Abs(Personnage.PositionX - .PositionX) < DistanceMaison(NumeroIA) And _
'               Abs(Personnage.PositionY - .PositionY) < DistanceMaison(NumeroIA))
'            'If NumeroIA = 0 Then
'            '    Debug.Print IndiceMaison(NumeroIA) & " : " & DistanceMaison(NumeroIA) & " : " & Trouver_Maison_Proche & " / " & NombreRechercheParTour
'            '    If IndiceMaison(NumeroIA) = 0 Then
'            '        Debug.Print Maisons(IndiceMaison(NumeroIA)).Service & " = " & TypeService & " : " & Maisons(IndiceMaison(NumeroIA)).Construit & " : " & Maisons(IndiceMaison(NumeroIA)).Definir_Stock & " >= 1" & " : " & Persos(IndiceMaison(NumeroIA)).NumeroFief & " = " & Persos(NumeroIA).NumeroFief & " : " & Abs(Personnage.PositionX - Maisons(IndiceMaison(NumeroIA)).PositionX) & " < " & DistanceMaison(NumeroIA) & " : " & Abs(Personnage.PositionY - Maisons(IndiceMaison(NumeroIA)).PositionY) & " < " & DistanceMaison(NumeroIA)
'            '    End If
'            'End If
'
'        ElseIf TypeBesoin = 5 Then
'            'Le personnage cherche un objet.
'            'Cherche une maison où il peut acheter des objets.
'
'                Trouver_Maison_Proche = (.Fabrique And .Construit And .Definir_Stock >= 1 And _
'                        Persos(IndiceMaison(NumeroIA)).NumeroFief = Persos(NumeroIA).NumeroFief And _
'                        Abs(Personnage.PositionX - .PositionX) < DistanceMaison(NumeroIA) And _
'                        Abs(Personnage.PositionY - .PositionY) < DistanceMaison(NumeroIA))
'
'        'If Parametres.Batiment_RessourceFabrique(Maisons(IndiceMaison(NumeroIA)).TypeBatiment) >= 0 Then
'        ElseIf .Magasin Or .Fabrique Then
'            If TypeRessource = 0 Then
'
'                pbm = Parametres.Batiment_RessourceFabrique(.TypeBatiment)
'
'                If pbm >= 0 And _
'                   TypeBesoin < 5 And _
'                   .Construit And .Definir_Stock >= 1 And _
'                   Persos(IndiceMaison(NumeroIA)).NumeroFief = Persos(NumeroIA).NumeroFief And _
'                   Abs(Personnage.PositionX - .PositionX) < DistanceMaison(NumeroIA) And _
'                   Abs(Personnage.PositionY - .PositionY) < DistanceMaison(NumeroIA) Then
'                   'DistanceMaison(NumeroIA) >= Sqr((Personnage.PositionX - Maisons(IndiceMaison(NumeroIA)).PositionX) ^ 2 + _
'                                                   (Personnage.PositionY - Maisons(IndiceMaison(NumeroIA)).PositionY) ^ 2) Then
'
'
'                    Select Case TypeBesoin
'                        'Le personnage a besoin de vie.
'                        Case 1:
'                            If Parametres.Ressources_GainVie(pbm) > 0 Then
'                                Trouver_Maison_Proche = True
'                            End If
'                        'Le personnage a besoin de vie.
'                        Case 2:
'                            If Parametres.Ressources_GainEnergie(pbm) > 0 Then
'                                Trouver_Maison_Proche = True
'                            End If
'                        'Le personnage a besoin de magie.
'                        Case 3:
'                            If Parametres.Ressources_GainMagie(pbm) > 0 Then
'                                Trouver_Maison_Proche = True
'                            End If
'                        'Le personnage a besoin de moral.
'                        Case 4:
'                            If Parametres.Ressources_GainMoral(pbm) > 0 Then
'                                Trouver_Maison_Proche = True
'                            End If
'                    End Select
'                End If
'            'Recherche une maison où il peut trouver la ressource passée en paramètre.
'            ElseIf Parametres.Batiment_RessourceFabrique(.TypeBatiment) = TypeRessource And _
'                   .Construit And _
'                   .Definir_Stock >= 1 And _
'                   Persos(IndiceMaison(NumeroIA)).NumeroFief = Persos(NumeroIA).NumeroFief And _
'                   Abs(Personnage.PositionX - .PositionX) < DistanceMaison(NumeroIA) And _
'                   Abs(Personnage.PositionY - .PositionY) < DistanceMaison(NumeroIA) And _
'                   DistanceMaison(NumeroIA) >= Sqr((Personnage.PositionX - Maisons(IndiceMaison(NumeroIA)).PositionX) ^ 2 + _
'                                                    (Personnage.PositionY - Maisons(IndiceMaison(NumeroIA)).PositionY) ^ 2) Then
'                    Trouver_Maison_Proche = True
'                'End If
'            End If
'        End If
'        End With
'
'        If Trouver_Maison_Proche Then
'            Indice = IndiceMaison(NumeroIA)
'            DistanceMaison(NumeroIA) = DistancePalierRechercheMaison
'            i = UBound(Maisons())
'        End If
'        IndiceMaison(NumeroIA) = IndiceMaison(NumeroIA) + 1
'    Next i
'    'Trouver_Maison_Proche = Trouver_Maison_Proche Or Indice < 0
'End Function
'
'Public Function Trouver_Emplacement_Maison(ByVal NumeroJoueur As Long, _
'                                           ByRef X As Long, _
'                                           ByRef Y As Long, _
'                                           ByVal TypeMaison As Integer) As Boolean
'    Dim i As Long
'    For i = 0 To NombreRecherchesEmplacementMaison
'        With Persos(NumeroJoueur)
'        DistanceObstacle(.NumeroFief) = DistanceObstacle(.NumeroFief) + Monde.LargeurCase
'        X = Chateaux(.NumeroFief).PositionX + RND * (DistanceObstacle(.NumeroFief) * 2 + Chateaux(.NumeroFief).Largeur) - DistanceObstacle(.NumeroFief)
'        If X < 0 Then
'            X = RND * Monde.LargeurCase
'        ElseIf X > Monde.Largeur * Monde.LargeurCase - Parametres.Batiment_Largeur(TypeMaison) Then
'            X = Monde.Largeur * Monde.LargeurCase - Parametres.Batiment_Largeur(TypeMaison) - RND * Monde.LargeurCase
'        End If
'        Y = Chateaux(.NumeroFief).PositionY + RND * (DistanceObstacle(.NumeroFief) * 2 + Chateaux(.NumeroFief).Hauteur) - DistanceObstacle(.NumeroFief)
'        If Y < 0 Then
'            Y = RND * Monde.LargeurCase
'        ElseIf Y > Monde.Hauteur * Monde.HauteurCase - Parametres.Batiment_Hauteur(TypeMaison) Then
'            Y = Monde.Hauteur * Monde.HauteurCase - Parametres.Batiment_Hauteur(TypeMaison) - RND * Monde.HauteurCase
'        End If
'        If Obstacle(NumeroJoueur, X, Y, Parametres.Batiment_Largeur(TypeMaison), Parametres.Batiment_Hauteur(TypeMaison)) = 0 Then
'            DistanceObstacle(.NumeroFief) = Chateaux(.NumeroFief).Largeur
'            Trouver_Emplacement_Maison = True
'            i = NombreRecherchesEmplacementMaison
'        End If
'        End With
'    Next i
'End Function
'
'Public Sub IA_Travail(ByVal i As Long)
'    With Persos(i)
'    If .Inactif Then
'        'S'il est dans un magasin, il achète ce dont il a besoin.
'        If .DansUneMaison And Not QuitterBatiment(i) Then
'           'and Not IA_Argent_Besoin(i) Then
'            If Maisons(.IndiceMaison).Magasin Then
'                'Magasin de ressources.
'                If Parametres.Batiment_RessourceFabrique(Maisons(.IndiceMaison).TypeBatiment) = _
'                   Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment) Then
'                    'L'IA achète le plus de matières premières pour son travail.
'                    Maisons(.IndiceMaison).Acheter 0, Persos(i), Parametres, True
'                Else
'                    Maisons(.IndiceMaison).Acheter Int(QuantiteMaxTransfert * RND), Persos(i), Parametres
'                End If
'            ElseIf Maisons(.IndiceMaison).Fabrique Then
'                'Fabrique d'objets.
'                'Achete un objet au hasard.
'                Z = Int(RND * Parametres.Batiment_NombreObjetsFabriques(Maisons(Persos(i).IndiceMaison).TypeBatiment))
'                If Maisons(.IndiceMaison).Stock_Objets(Z) > 0 And _
'                   Not .Objet_Posseder(Parametres.Batiment_ObjetFabrique(Maisons(.IndiceMaison).TypeBatiment, Z)) Then
'                    .Objet_Acheter_Maison Z, Maisons(.IndiceMaison), Parametres, Commentaires
'                End If
'            Else
'                'Batiments de services.
'                If Maisons(.IndiceMaison).Service_Utilisable(Persos(i), Parametres) Then
'                    Maisons(.IndiceMaison).Service_Utiliser Persos(i), Parametres
'                End If
'            End If
'            IA_Sortir_Batiment i
'            'Persos(i).Aller_A Persos(i).PositionX, Persos(i).PositionY + Maisons(i).Hauteur / 2
'
'        'Le personnage est dans sa maison,
'        ElseIf .DansSaMaison And Not QuitterBatiment(i) Then
'            IA_Travailler_Maison i
'
'        'Fait ses transactions au chateau
'        ElseIf (.DansUnChateau Or .DansSonChateau) And _
'               Not QuitterBatiment(i) Then
'            IA_Transaction_Chateau i
'
'        'Travail.
'        'Va chercher de l'argent chez lui si nécessaire.
'        ElseIf IA_Argent_Besoin(i) And _
'               Maisons(i).Argent >= 1 Then
'           IA_Rentrer_Maison i
'
'        'Travail de mercenaire.
'        ElseIf Maisons(i).Service = 2 Then
'            If .Vie = .Definir_Vie Then
'                IA_Agressif i: IA_Agressif i: IA_Agressif i
'                'IA_Agressif i: IA_Agressif i: IA_Agressif i
'                If .Inactif Then
'                    If Maisons(i).Argent > 0 Then
'                        IA_Rentrer_Maison i
'                    End If
'                End If
'            End If
'
'        'Rentre à la maison pour travailler.
'        ElseIf .Definir_Ressources(Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment)) _
'               < MatierePremiereMin And _
'               Maisons(i).Definir_Stock <= Maisons(i).MaxStock And _
'               (Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment) >= 0 Or Maisons(i).Fabrique Or Maisons(i).Service > 0) Then
'                    IA_Rentrer_Maison i
'
'        'Va au château pour vendre des ressources.
'        ElseIf IA_Stock_Exedentaire(i) Or _
'               (IA_Argent_Besoin(i) Or _
'               (.Definir_Ressources(Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)) <= 1 And _
'               Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment) >= 0) And _
'               Fiefs(.NumeroFief).Definir_Ressources(Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)) >= 1) And _
'               Chateaux(.NumeroFief).Visible Then
'                    IA_Rentrer_Chateau i
'        Else
'            'S'il n'a rien à faire, il attaque tout ce qui bouge.
'            IA_Agressif i
'        End If
'    End If
'    End With
'End Sub
'
'Public Sub IA_Travailler_Maison(ByVal i As Long)
'    With Persos(i)
'    'Il récupère ses économies.
'    Maisons(i).Transfert_Argent_Prendre 0, Persos(i), True
'    If Maisons(i).Fabrique Then
'        'Fabrique des objets.
'        Do
'        Z = Int(RND * Parametres.Batiment_NombreObjetsFabriques(Maisons(i).TypeBatiment))
'        Loop Until Parametres.Objet_Epoque(Parametres.Batiment_ObjetFabrique(Maisons(i).TypeBatiment, Z)) - 1 <= Fiefs(Persos(i).NumeroFief).Epoque
'        For X = 0 To Parametres.Batiment_NombreObjetsFabriques(Maisons(i).TypeBatiment) - 1
'            If Maisons(i).Stock_Objets(X) = 0 And _
'               .PeutConstruire_Objet(X, Maisons(i), Parametres) Then
'                Z = X
'                X = Parametres.Batiment_NombreObjetsFabriques(Maisons(i).TypeBatiment)
'            End If
'        Next X
'        If .PeutConstruire_Objet(Z, Maisons(i), Parametres) Then
'            .Creer_Objet Z, Maisons(i), Parametres
'        Else
'            'Récupère ses objets pour les vendre.
'            If IA_Argent_Besoin(i) Then
'                For X = 0 To Parametres.Batiment_NombreObjetsFabriques(Maisons(i).TypeBatiment) - 1
'                    If Maisons(i).Stock_Objets(X) >= 1 Then
'                        .Objet_Ramasser_Maison X, Maisons(i), Parametres, Commentaires
'                        X = Parametres.Batiment_NombreObjetsFabriques(Maisons(i).TypeBatiment)
'                    End If
'                Next X
'            End If
'            IA_Rentrer_Chateau i
'            'IA_Sortir_Batiment i
'        End If
'    Else
'        'Si le stock du magasin est vide, il le remplit.
'        If Maisons(i).Magasin And (Maisons(i).Definir_Stock <= 1 Or Not IA_Argent_Besoin(i)) Then
'            'Maisons(i).Transfert_Stock_Deposer QuantiteMaxTransfert, Persos(i), Parametres
'            Maisons(i).Transfert_Stock_Deposer 0.5 * .Definir_Ressources(Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment)), Persos(i), Parametres
'        End If
'        IA_Sortir_Batiment i
'    End If
'    End With
'End Sub
'
'Public Sub IA_Transaction_Chateau(ByVal i As Long)
'    With Persos(i)
'    'Vend son stock s'il n'a pas de magasin, pas d'argent ou que les réserves de son magasin sont pleines.
'    If Not Maisons(i).Magasin Or _
'       IA_Argent_Besoin(i) Or _
'       Maisons(i).Definir_Stock >= Maisons(i).MaxStock - 1 Then
'         If Maisons(i).Fabrique Then
'             'Vend ses objets.
'             For X = 0 To .Nombre_Objets_Inventaire - 1
'                 For Y = 0 To Parametres.Batiment_NombreObjetsFabriques(Maisons(i).TypeBatiment) - 1
'                     If .Objet_Inventaire_Actif(X) And _
'                        .Objet_Inventaire_Type(X) = Parametres.Batiment_ObjetFabrique(Maisons(i).TypeBatiment, Y) Then
'                         .Objet_Vendre_Chateau X, Fiefs(.IndiceChateau), Parametres
'                     End If
'                 Next Y
'             Next X
'         ElseIf Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment) >= 0 Then
'             Fiefs(Persos(i).NumeroFief).Vendre Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment), 0, Persos(i), Parametres, True
'         End If
'    End If
'    'Achète les ressources dont il a besoin.
'    'If Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment) <> 1 Then 'Pas besoin d'acheter de bois.
'    'A besoin de ressources pour son métier.
'    IA_Acheter_Ressources_Pour_Travailler i
'    'Achete des ressources pour construire.
'    If Not IA_Argent_Besoin(i) Then
'        IA_Acheter_Ressources_Pour_Construire i
'    End If
'
'    IA_Restaurer_Chateau i
'
'     'Il finit par quitter le chateau.
'     IA_Sortir_Batiment i
'     End With
'End Sub
'
'Private Sub IA_Restaurer_Chateau(ByVal i As Long)
'    With Persos(i)
'    'Restaure le chateau.
'     If Chateaux(.NumeroFief).Definir_Vie < Chateaux(.NumeroFief).MaxVie Then
'         'Le chateau est endommagé.
'         'If Not IA_Argent_Besoin(i) Then
'         '    Fiefs(.NumeroFief).Restaurer_Chateau 5, Persos(i)
'         'ElseIf IA_Argent_Riche(i) Then
'         '    Fiefs(.NumeroFief).Restaurer_Chateau 50, Persos(i)
'         'End If
'         Fiefs(.NumeroFief).Restaurer_Chateau Int(.Definir_Argent - (.Definir_Argent * Chateaux(.NumeroFief).Definir_Vie / Chateaux(.NumeroFief).MaxVie)) + 1, Persos(i)
'     End If
'     End With
'End Sub
'
'Public Sub IA_Vente_Surplus_Objets(ByVal i As Long)
'    Dim j As Integer
'    With Persos(i)
'        If (.DansSonChateau Or .DansUnChateau) Then
'            For j = .Nombre_Objets_Inventaire - 1 To 0 Step -1
'                If .Objet_Inventaire_Actif(j) Then
'                   If Not Parametres.Objet_UtilisationUnique(.Objet_Inventaire_Type(j)) And _
'                      ((.Objet_NombreType(.Objet_Inventaire_Type(j)) > 1 And _
'                       Parametres.Objet_Categorie(.Objet_Inventaire_Type(j)) <> 0) Or _
'                      (.Objet_NombreType(.Objet_Inventaire_Type(j)) > 2 And _
'                       Parametres.Objet_Categorie(.Objet_Inventaire_Type(j)) = 0)) Then
'                        .Objet_Vendre_Chateau j, Fiefs(.NumeroFief), Parametres
'                    End If
'                End If
'            Next j
'        End If
'    End With
'End Sub

'Public Function Obstacle(ByVal NumeroJoueur As Long, _
'                         ByVal X As Long, _
'                         ByVal Y As Long, _
'                         Optional ByVal Largeur As Long = 0, _
'                         Optional ByVal Hauteur As Long = 0, _
'                         Optional ByVal MemoriserIndiceObstacle As Boolean) As Long
'    'Cette fonction renvoie l'obstacle situé en x, y.
'    Dim i As Long, Pne As Long
'    Pne = Persos(NumeroJoueur).NumeroEquipe
'
'    'Vérifie les personnages ennemis.
'    for i=0 to ubound(persos())
'        If Est_Superpose(X, _
'                         Y, _
'                         Persos(i).PositionX, _
'                         Persos(i).PositionY, _
'                         Largeur, _
'                         Hauteur, _
'                         Persos(i).Largeur, _
'                         Persos(i).Hauteur) Then
'            If Persos(i).NumeroEquipe <> Pne And _
'               Persos(i).Attaquable Then
'                Obstacle = 9
'                If MemoriserIndiceObstacle Then IndiceObstacle = i
'                i = UBound(Persos())
'            End If
'        End If
'    Next i
'
'    'Vérifie les trésors.
'    If Obstacle = 0 And Tresors.Count > 0 Then
'        For i = 1 To Tresors.Count
'            If Est_Superpose(X, _
'                             Y, _
'                             Tresors(i).X, _
'                             Tresors(i).Y, _
'                             Largeur, _
'                             Hauteur, _
'                             AffTresor.Largeur, _
'                             AffTresor.Hauteur) Then
'                Obstacle = 13
'                If MemoriserIndiceObstacle Then
'                    IndiceObstacle = i
'                End If
'                i = Tresors.Count
'            End If
'        Next i
'    End If
'
'    If Obstacle = 0 Then
'        'Vérifie les chateaux.
'        For i = 0 To UBound(Chateaux())
'            If Est_Superpose(X, _
'                             Y, _
'                             Chateaux(i).PositionX, _
'                             Chateaux(i).PositionY, _
'                             Largeur, _
'                             Hauteur, _
'                             Chateaux(i).Largeur, _
'                             Chateaux(i).Hauteur) And _
'               Chateaux(i).Visible Then
'                If i = Persos(NumeroJoueur).NumeroFief Then
'                    Obstacle = 6
'                ElseIf Fiefs(i).NumeroEquipe = Pne Then
'                       'Cas où s'est un chateau allié.
'                       Obstacle = 8
'                       If MemoriserIndiceObstacle Then IndiceObstacle = i
'                Else 'C'est pas un allié.
'                       Obstacle = 11
'                       If MemoriserIndiceObstacle Then IndiceObstacle = i
'                End If
'                i = UBound(Chateaux())
'            End If
'        Next i
'    End If
'
'    'Vérifie les maisons.
'    If Obstacle = 0 Then
'        For i = 0 To UBound(Maisons())
'            If Maisons(i).Visible And _
'                Est_Superpose(X, _
'                              Y, _
'                              Maisons(i).PositionX, _
'                              Maisons(i).PositionY, _
'                              Largeur, _
'                              Hauteur, _
'                              Maisons(i).Largeur, _
'                              Maisons(i).Hauteur) Then
'                If i = NumeroJoueur Then 'Cas où c'est la maison du joueur.
'                    If Maisons(i).Construit And Maisons(i).Definir_Vie = Maisons(i).MaxVie Then
'                        'Cas où c'est la maison du joueur à utiliser.
'                        Obstacle = 5
'                    Else
'                        'Cas où c'est la maison du joueur à construire.
'                        Obstacle = 4
'                    End If
'                Else
'                    'Cas ou c'est une maison qui n'appartient pas au joueur.
'                    If Persos(i).NumeroEquipe = Pne Then
'                        'Cas où c'est une maison alliée.
'                        'If (Maisons(i).Magasin Or Maisons(i).Fabrique Or Maisons(i).Service > 0) And _
'                            Maisons(i).Construit Then
'                            If Maisons(i).Service = 1 Then
'                                'Cas où c'est un temple.
'                                Obstacle = 12
'                            Else
'                                Obstacle = 7
'                            End If
'                        'Else
'                        '    Obstacle = 999
'                        'End If
'                    Else 'Ce n'est pas une maison alliée.
'                        Obstacle = 10
'                    End If
'                End If
'                If MemoriserIndiceObstacle Then IndiceObstacle = i
'                i = UBound(Maisons())
'            End If
'        Next i
'    End If
'
'    'Vérifie les ressources récoltables.
'    If Obstacle = 0 Then
'        For i = 0 To UBound(Ressources())
'            If Est_Superpose(X, _
'                             Y, _
'                             Ressources(i).PositionX, _
'                             Ressources(i).PositionY, _
'                             Largeur, _
'                             Hauteur, _
'                             AffRessources.Largeur, _
'                             AffRessources.Hauteur) Then
'                Obstacle = 1
'                If MemoriserIndiceObstacle Then
'                    IndiceObstacle = i
'                End If
'                i = UBound(Ressources())
'            End If
'        Next i
'    End If
'
'    'Vérifie les décors.
'    If Obstacle = 0 Then
'        For i = 1 To Decors.Count
'            If Est_Superpose(X, _
'                             Y, _
'                             Decors(i).X, _
'                             Decors(i).Y, _
'                             Largeur, _
'                             Hauteur, _
'                             Decors(i).Largeur, _
'                             Decors(i).Hauteur) Then
'                Obstacle = 999
'                If MemoriserIndiceObstacle Then
'                    IndiceObstacle = i
'                End If
'                i = UBound(Ressources())
'            End If
'        Next i
'    End If
'
'    'Vérifie le relief du terrain.
'    If Obstacle = 0 Then
'        If Not Monde.Constructible(X, Y, Largeur, Hauteur) Then
'            Obstacle = 999
'        End If
'    End If
'End Function
'
'Public Function ObstacleAmi(ByVal NumeroJoueur As Long, _
'                            ByVal X As Long, _
'                            ByVal Y As Long, _
'                            Optional ByVal Largeur As Long = 0, _
'                            Optional ByVal Hauteur As Long = 0, _
'                            Optional ByVal MemoriserIndiceObstacle As Boolean) As Boolean
'    'Cette fonction renvoie le personnage de la même équipe.
'    Dim i As Long, Pne As Long
'    Pne = Persos(NumeroJoueur).NumeroEquipe
'
'    'Vérifie les personnages ennemis.
'    for i=0 to ubound(persos())
'        With Persos(i)
'        If Est_Superpose(X, _
'                         Y, _
'                         .PositionX, _
'                         .PositionY, _
'                         Largeur, _
'                         Hauteur, _
'                         .Largeur, _
'                         .Hauteur) Then
'            If .NumeroEquipe = Pne And _
'               .Attaquable And _
'               i <> NumeroJoueur Then
'                If MemoriserIndiceObstacle Then
'                    IndiceObstacle = i
'                End If
'                ObstacleAmi = True
'                i = UBound(Persos())
'            End If
'        End If
'        End With
'    Next i
'End Function
'
'Public Function Est_Superpose(ByVal X1 As Long, _
'                              ByVal Y1 As Long, _
'                              ByVal X2 As Long, _
'                              ByVal Y2 As Long, _
'                              Optional ByVal Largeur1 As Long = 0, _
'                              Optional ByVal Hauteur1 As Long = 0, _
'                              Optional ByVal Largeur2 As Long = 0, _
'                              Optional ByVal Hauteur2 As Long = 0) As Boolean
'    'Vérifie si les 2 objets sont superposés.
'
'    'Est_Superpose = False
'
'    If Largeur1 = 0 And Hauteur1 = 0 Then
'        Est_Superpose = X1 >= X2 And _
'                        X1 <= X2 + Largeur2 And _
'                        Y1 >= Y2 And _
'                        Y1 <= Y2 + Hauteur2
'    Else
'        Est_Superpose = _
'          ((X1 <= X2 And X2 <= X1 + Largeur1 Or _
'            X1 <= X2 + Largeur2 And X2 + Largeur2 <= X1 + Largeur1) And _
'           (Y1 <= Y2 And Y2 <= Y1 + Hauteur1 Or _
'            Y1 <= Y2 + Hauteur2 And Y2 + Hauteur2 <= Y1 + Hauteur1) Or _
'           (X2 <= X1 And X1 <= X2 + Largeur2 Or _
'            X2 <= X1 + Largeur1 And X1 + Largeur1 <= X2 + Largeur2) And _
'           (Y2 <= Y1 And Y1 <= Y2 + Hauteur2 Or _
'            Y2 <= Y1 + Hauteur1 And Y1 + Hauteur1 <= Y2 + Hauteur2) Or _
'           (X1 >= X2 And X1 + Largeur1 <= X2 + Largeur2 And _
'            Y1 >= Y2 And Y1 + Largeur1 <= Y2 + Largeur2) Or _
'           (X2 >= X1 And X2 + Largeur2 <= X1 + Largeur1 And _
'            Y2 >= Y1 And Y2 + Largeur2 <= Y1 + Largeur1)) Or _
'           (X1 < X2 And X1 + Largeur1 > X2 + Largeur2 And _
'            Y2 < Y1 And Y1 + Hauteur1 < Y2 + Hauteur2 Or _
'            X2 < X1 And X2 + Largeur2 > X1 + Largeur1 And _
'            Y1 < Y2 And Y2 + Hauteur2 < Y1 + Hauteur1)
'    End If
'End Function
'
'
'Public Function ObstacleChateau(ByVal NumeroJoueur As Long, _
'                                ByVal PositionX As Long, _
'                                ByVal PositionY As Long, _
'                                ByVal Largeur As Long, _
'                                ByVal Hauteur As Long) As Boolean
'    Dim i As Long
'    'Vérifie le relief du terrain.
'    If Not ObstacleChateau Then
'        ObstacleChateau = Monde.Constructible(X, Y, Largeur, Hauteur)
'    End If
'
'    'Vérifie que le chateau n'est pas superposé à un autre.
'    For i = 0 To UBound(Chateaux())
'        If Est_Superpose(PositionX, _
'                         PositionY, _
'                         Chateaux(i).PositionX, _
'                         Chateaux(i).PositionY, _
'                         Largeur, _
'                         Hauteur, _
'                         Chateaux(i).Largeur, _
'                         Chateaux(i).Hauteur) And _
'           Chateaux(i).Visible Then
'            ObstacleChateau = True
'            i = UBound(Chateaux())
'        End If
'    Next i
'End Function
'
'Public Property Let Nombre_IAs(ByVal Valeur As Long)
'    LPerso = Valeur - 1
'
'    ReDim Preserve DistanceRessource(LPerso)
'    ReDim Preserve DistanceTresor(LPerso)
'    ReDim Preserve DistanceMaison(LPerso)
'    ReDim Preserve DistanceAdversaire(LPerso)
'    ReDim Preserve DistanceMaisonEnnemi(LPerso)
'    ReDim Preserve DistanceChateauEnnemi(LPerso)
'    ReDim Preserve IndiceMaison(LPerso)
'    ReDim Preserve IndiceAdversaire(LPerso)
'    ReDim Preserve IndiceMaisonEnnemi(LPerso)
'    ReDim Preserve IndiceChateauEnnemi(LPerso)
'
'    NombreIAParTour = LPerso * VitesseIA
'    If NombreIAParTour = 0 Then
'        NombreIAParTour = 1
'    End If
'    NombreRechercheParTour = LPerso * VitesseRechercheIA
'    If NombreRechercheParTour = 0 Then
'        NombreRechercheParTour = 1
'    End If
'    ReDim Preserve BatimentAConstruire(LPerso)
'    ReDim Preserve RessourceAFabriquer(LPerso)
'    ReDim Preserve ObjetAFabriquer(LPerso)
'    ReDim Preserve QuantiteAFabriquer(LPerso)
'    ReDim Preserve QuitterBatiment(LPerso)
'End Property

Public Function Fabriquer_Ressource(ByVal TypeRessource As Integer, _
                                    ByVal Quantite As Double) As Boolean
    If Persos(Numero).Vivant Then
        TypeRessource = TypeRessource - 1
        If Parametres.Ressources_Epoque(TypeRessource) <= Fiefs(Persos(Numero).NumeroFief).Epoque Then
            IA2_Fabriquer_Ressource Numero, TypeRessource, Quantite
            Fabriquer_Ressource = True
        End If
    End If
End Function
Public Function Fabriquer_Objet(ByVal TypeObjet As Integer, _
                                ByVal Quantite As Integer) As Boolean
    If Persos(Numero).Vivant Then
        TypeObjet = TypeObjet - 1
        If Parametres.Objet_Epoque(TypeObjet) <= Fiefs(Persos(Numero).NumeroFief).Epoque And _
           Not Parametres.Objet_Secret(TypeObjet) Then
            IA2_ChercherObjet Numero, TypeObjet, Quantite
            Fabriquer_Objet = True
        End If
    End If
End Function

Public Property Get Objet_Equipe_Type(ByVal Emplacement As Integer) As Integer
    'Renvoie le type d'objet équipé à l'emplacement indiqué.
    If Emplacement > 0 And Emplacement <= Persos(Numero).Nombre_Objets_Equipes Then
        Emplacement = Emplacement - 1
        If Persos(Numero).Objet_Equipes_Actif(Emplacement) Then
            Objet_Equipe_Type = Persos(Numero).Objet_Equipes_Type(Emplacement) + 1
        Else
            Objet_Equipe_Type = 0
        End If
    End If
End Property
Public Property Get Objet_Inventaire_Type(ByVal Emplacement As Integer) As Integer
    'Renvoie le type d'objet de l'inventaire à l'emplacement indiqué.
    If Emplacement > 0 And Emplacement <= Persos(Numero).Nombre_Objets_Inventaire Then
        Emplacement = Emplacement - 1
        If Persos(Numero).Objet_Inventaire_Actif(Emplacement) Then
            Objet_Inventaire_Type = Persos(Numero).Objet_Inventaire_Type(Emplacement) + 1
        Else
            Objet_Inventaire_Type = 0
        End If
    End If
End Property

Public Property Get Objet_TotalQuantite(ByVal TypeObjet As Integer) As Integer
    'Renvoie le nombre d'objets que possède le personnage sur lui.
    Objet_TotalQuantite = Persos(Numero).Objet_NombreType(TypeObjet - 1)
End Property
Public Property Get Objet_EquipeQuantite(ByVal TypeObjet As Integer) As Integer
    'Renvoie le nombre d'objets que le personnage a d'équipé sur lui.
    Objet_EquipeQuantite = Persos(Numero).Objet_EquipeNombreType(TypeObjet - 1)
End Property
Public Property Get Objet_SacQuantite(ByVal TypeObjet As Integer) As Integer
    'Renvoie le nombre d'objets que le personnage a d'équipé sur lui.
    Dim i As Integer
    TypeObjet = TypeObjet - 1
    For i = 0 To Persos(Numero).Nombre_Objets_Inventaire - 1
        If Persos(Numero).Objet_Inventaire_Actif(i) And _
           Persos(Numero).Objet_Inventaire_Type(i) = TypeObjet Then
            Objet_SacQuantite = Objet_SacQuantite + 1
        End If
    Next i
End Property
Public Function Objet_SEquiper(ByVal TypeObjet As Integer) As Boolean
    'Utilise ou s'équipe avec l'objet en paramètre.
    Dim i As Integer
    TypeObjet = TypeObjet - 1
    For i = 0 To Persos(Numero).Nombre_Objets_Inventaire - 1
        If Persos(Numero).Objet_Inventaire_Actif(i) And _
           Persos(Numero).Objet_Inventaire_Type(i) = TypeObjet Then
            Persos(Numero).Objet_Equiper i, Parametres, Commentaires
            i = Persos(Numero).Nombre_Objets_Inventaire
            Objet_SEquiper = True
        End If
    Next i
End Function
'Private Function IA_Changer_Maison(ByVal i As Long) As Boolean
'
'End Function
'Public Sub IA2_Choisir_Maison(ByVal i As Long, ByVal TypeBatiment As Integer)
'
'End Sub

