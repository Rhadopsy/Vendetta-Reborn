VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsComPerso"
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

'Cette classe sert à vérifier si les données ont été changées, on envoie un message réseau.

Public Nom As String
'Public NumeroFief As String
Public Empereur As Boolean
Public Race As Integer
Public IA As Boolean
Public PositionX As Long
Public PositionY As Long
Public Vie As Double
Public BonusMaxVie As Long
Public RecuperationVie As Double
Public MaxVie As Long
Public ExpVie As Double
Public Energie As Double
Public BonusMaxEnergie As Long
Public RecuperationEnergie As Double
Public MaxEnergie As Long
Public ExpEnergie As Double
Public Magie As Double
Public BonusMaxMagie As Long
Public RecuperationMagie As Double
Public MaxMagie As Long
Public ExpMagie As Double
Public Moral As Double
Public BonusMaxMoral As Long
Public RecuperationMoral As Double
Public MaxMoral As Long
Public ExpMoral As Double
Public Attaque As Double
Public DureeRechargeAttaque As Long
Public DureeRechargeAttaqueMax As Long
Public BonusMaxAttaque As Long
Public MaxAttaque As Long
Public ExpAttaque As Double
Public Defense As Double
Public BonusMaxDefense As Long
Public MaxDefense As Long
Public ExpDefense As Double
'    Joueur
Public NumeroEquipe As Integer
Public NumeroPas As Integer
Public Direction As Integer
Public Feminin As Boolean
'    Vivant
Public ChoisirApparence As Boolean
'    FichierApparence
Public ApparenceChoisie As Integer
Private NombreApparenceBase As Integer
Private ApparenceBase() As Integer
'    Case Categorie2PersoDefinir_FichierApparenceBase:
'    Definir_FichierApparenceBase()
Public ApparenceObjetAfficher As Boolean
'    Case Categorie2PersoNombreApparenceObjet:
'        If Persos(Indice1).Nombre_ApparenceObjet <> Var1 Then
'            Persos(Indice1).Nombre_ApparenceObjet = Var1
'        End If
Private ApparenceDecalageX() As Integer 'Décale l'apparence du personnage horizontalement.
Private ApparenceDecalageY() As Integer 'Décale l'apparence du personnage verticalement.
Private ApparenceAvantPlan() As Boolean 'Place l'apparence en question en arrière plan selon la direction du personnage comme indice.
Private ApparenceObjet As Collection
Public DansSaMaison As Boolean
Public DansUneMaison As Boolean
Public DansSonChateau As Boolean
Public DansUnChateau As Boolean
'    Case Categorie2PersoChangerNombre: While Indice1 > UBound(Persos()):           Perso_Ajouter_Nouveau:        Wend
'    ChoisirIA
'    FichierIA
'    ChoisirIAScript
'    FichierIAScript
Public NumeroChef  As Long
Public NombreSoldats As Long
'Public Attitude As Integer
'Public Formation As Integer
'Public ExpVie As Double
'    Case Categorie2PersoExpEnergie: Persos(Indice1).ExpEnergie = Var1
'    Case Categorie2PersoExpMagie: Persos(Indice1).ExpMagie = Var1
'    Case Categorie2PersoExpMoral: Persos(Indice1).ExpMoral = Var1
'    Case Categorie2PersoExpAttaque: Persos(Indice1).ExpAttaque = Var1
'    Case Categorie2PersoExpDefense: Persos(Indice1).ExpDefense = Var1

'Informations pour les compétences.
Private Carac_CompetenceSpeciales() As Long
Private Niveau_CompetenceSpeciales() As Long
Private ExpSpeciales() As Double
Private Carac_Ressources() As Long
Private Niveau_CompetenceRessources() As Long
Private ExpRessources() As Double
Private Carac_CompetenceServices() As Long
Private Niveau_CompetenceServices() As Long
Private ExpServices() As Double
Private Carac_CompetenceObjets() As Long
Private Niveau_CompetenceObjets() As Long
Private ExpObjets() As Double

Private Bestiaire() As Integer
Public Action As Integer
Public IndiceObjet As Integer
Public EtatFabricationObjet As Double
Public RessourceCourante As Long
Public IndiceRessource As Long
Public IndicePerso As Long
Public IndiceTresor As Long
Public IndiceMaison As Long
Public IndiceChateau As Long
Public Argent As Double
Public ArgentSelectionne As Double
Private Ressources() As Double
Private MaxRessources() As Double
Private RessourcesSelectionnes() As Double
Public ObjetSelectionne As Integer
Private ObjetEquipesType() As Integer
Private ObjetInventaireType() As Integer
Private ObjetInventaireNombre() As Integer
Public Armure As Integer
Public Vitalite  As Double
Public Longueur_PasTemp As Long
Public BonusDeplacement As Long
'    Case Categorie2PersoEffetTempNombre: Persos(Indice1).Nombre_EffetsTemp = Var1
'                                         Persos(Indice1).Nombre_EffetsRecharge = Var2
'    Case Categorie2PersoEffetTemp: Persos(Indice1).Definir_EffetTemp_Numero(Indice2) = Var1
'                                   Persos(Indice1).Definir_EffetTemp_Duree(Indice2) = Var2
'    Case Categorie2PersoEffetTempProprietaire: Persos(Indice1).Definir_EffetTemp_Proprietaire(Indice2) = Var1
'
'Private EffetsRecharge_Numero() As Integer
'Private EffetsRecharge_Duree() As Long
Public Commentaires As String
Public DureeResurection As Double

Private MaisonTypeBatiment As Integer
Public MaisonConstruit As Boolean
Public MaisonEnConstruction As Boolean
Public MaisonPositionX As Long
Public MaisonPositionY As Long
Public MaisonVie As Double
Public MaisonMaxVie As Long
Public MaisonStock As Long
Public MaisonMaxStock As Long
Private MaisonNombreStockObjet As Integer
Private MaisonStockObjets() As Long
Private MaisonObjetInventaireType() As Integer

Public Sub Init(ByVal Parametres As ClsJeuParam)
    Dim Nb As Integer 'Variable temporaire.
    Set ApparenceObjet = New Collection
    With Parametres
    Nb = .NombreRessources - 1
    ReDim Ressources(Nb)
    ReDim MaxRessources(Nb)
    ReDim RessourcesSelectionnes(Nb)
    Nb = .PersosNombreObjetsEquipes - 1
    ReDim ObjetEquipesType(Nb)
    Nb = .PersosNombreObjetsInventaire - 1
    ReDim ObjetInventaireType(Nb)
    ReDim ObjetInventaireNombre(Nb)
    ReDim MaisonObjetInventaireType(Nb)
    MaisonTypeBatiment = -1
    Maison_Nombre_Stock_Objet = 0
    
    Nb = .NombreCompetencesSpeciales - 1
    ReDim Carac_CompetenceSpeciales(Nb)
    ReDim Niveau_CompetenceSpeciales(Nb)
    ReDim ExpSpeciales(Nb)
    
    Nb = .NombreRessources - 1
    ReDim Carac_Ressources(Nb)
    ReDim Niveau_CompetenceRessources(Nb)
    ReDim ExpRessources(Nb)
    
    Nb = .NombreServices - 2
    ReDim Carac_CompetenceServices(Nb)
    ReDim Niveau_CompetenceServices(Nb)
    ReDim ExpServices(Nb)
    
    Nb = .NombreCompetencesObjets - 1
    ReDim Carac_CompetenceObjets(Nb)
    ReDim Niveau_CompetenceObjets(Nb)
    ReDim ExpObjets(Nb)
    
    ReDim Bestiaire(.NombreRaces - 1)
    
    ReDim ApparenceDecalageX(4)  'Quatre directions différentes.
    ReDim ApparenceDecalageY(4)  'Quatre directions différentes.
    ReDim ApparenceAvantPlan(4)
    
    End With
End Sub

Public Property Get Definir_NombreApparenceBase() As Integer
    Definir_NombreApparenceBase = NombreApparenceBase
End Property
Public Property Let Definir_NombreApparenceBase(ByVal Valeur As Integer)
    NombreApparenceBase = Valeur
    If NombreApparenceBase = 0 Then
        ReDim Preserve ApparenceBase(0)
    Else
        ReDim Preserve ApparenceBase(NombreApparenceBase - 1)
    End If
End Property

Public Property Get Definir_ApparenceBase(ByVal Index As Integer) As Integer
    Definir_ApparenceBase = ApparenceBase(Index)
End Property
Public Property Let Definir_ApparenceBase(ByVal Index As Integer, ByVal Valeur As Integer)
    ApparenceBase(Index) = Valeur
End Property

Public Property Let Definir_Nombre_ApparenceObjet(ByVal Valeur As Integer)
    While ApparenceObjet.Count > Valeur
        ApparenceObjet.Remove ApparenceObjet.Count
    Wend
    'Dim i As Integer
    'i = -1
    While Valeur > ApparenceObjet.Count
        ApparenceObjet.Add -1
    Wend
End Property
Public Property Get Definir_Nombre_ApparenceObjet() As Integer
    Definir_Nombre_ApparenceObjet = ApparenceObjet.Count
End Property
Public Property Get Definir_ApparenceObjet(ByVal Index As Integer) As Integer
    Definir_ApparenceObjet = ApparenceObjet(Index)
End Property
Public Property Let Definir_ApparenceObjet(ByVal Index As Integer, ByVal Valeur As Integer)
    ApparenceObjet.Remove Index
    ApparenceObjet.Add Valeur
    'ApparenceObjet.Item(Index) = Valeur
End Property

'Propriétés des compétences.
Public Property Get Definir_Carac_CompetenceSpeciales(ByVal Index As Integer) As Long
    Definir_Carac_CompetenceSpeciales = Carac_CompetenceSpeciales(Index)
End Property
Public Property Let Definir_Carac_CompetenceSpeciales(ByVal Index As Integer, ByVal Valeur As Long)
    Carac_CompetenceSpeciales(Index) = Valeur
End Property
Public Property Get Definir_Niveau_CompetenceSpeciales(ByVal Index As Integer) As Long
    Definir_Niveau_CompetenceSpeciales = Niveau_CompetenceSpeciales(Index)
End Property
Public Property Let Definir_Niveau_CompetenceSpeciales(ByVal Index As Integer, ByVal Valeur As Long)
    Niveau_CompetenceSpeciales(Index) = Valeur
End Property
Public Property Get Definir_ExpSpeciales(ByVal Index As Integer) As Double
    Definir_ExpSpeciales = ExpSpeciales(Index)
End Property
Public Property Let Definir_ExpSpeciales(ByVal Index As Integer, ByVal Valeur As Double)
    ExpSpeciales(Index) = Valeur
End Property
Public Property Get Definir_Carac_Ressources(ByVal Index As Integer) As Long
    Definir_Carac_Ressources = Carac_Ressources(Index)
End Property
Public Property Let Definir_Carac_Ressources(ByVal Index As Integer, ByVal Valeur As Long)
    Carac_Ressources(Index) = Valeur
End Property
Public Property Get Definir_Niveau_CompetenceRessources(ByVal Index As Integer) As Long
    Definir_Niveau_CompetenceRessources = Niveau_CompetenceRessources(Index)
End Property
Public Property Let Definir_Niveau_CompetenceRessources(ByVal Index As Integer, ByVal Valeur As Long)
    Niveau_CompetenceRessources(Index) = Valeur
End Property
Public Property Get Definir_ExpRessources(ByVal Index As Integer) As Double
    Definir_ExpRessources = ExpRessources(Index)
End Property
Public Property Let Definir_ExpRessources(ByVal Index As Integer, ByVal Valeur As Double)
    ExpRessources(Index) = Valeur
End Property
Public Property Get Definir_Carac_CompetenceServices(ByVal Index As Integer) As Long
    Definir_Carac_CompetenceServices = Carac_CompetenceServices(Index)
End Property
Public Property Let Definir_Carac_CompetenceServices(ByVal Index As Integer, ByVal Valeur As Long)
    Carac_CompetenceServices(Index) = Valeur
End Property
Public Property Get Definir_Niveau_CompetenceServices(ByVal Index As Integer) As Long
    Definir_Niveau_CompetenceServices = Niveau_CompetenceServices(Index)
End Property
Public Property Let Definir_Niveau_CompetenceServices(ByVal Index As Integer, ByVal Valeur As Long)
    Niveau_CompetenceServices(Index) = Valeur
End Property
Public Property Get Definir_ExpServices(ByVal Index As Integer) As Double
    Definir_ExpServices = ExpServices(Index)
End Property
Public Property Let Definir_ExpServices(ByVal Index As Integer, ByVal Valeur As Double)
    ExpServices(Index) = Valeur
End Property
Public Property Get Definir_Carac_CompetenceObjets(ByVal Index As Integer) As Long
    Definir_Carac_CompetenceObjets = Carac_CompetenceObjets(Index)
End Property
Public Property Let Definir_Carac_CompetenceObjets(ByVal Index As Integer, ByVal Valeur As Long)
    Carac_CompetenceObjets(Index) = Valeur
End Property
Public Property Get Definir_Niveau_CompetenceObjets(ByVal Index As Integer) As Long
    Definir_Niveau_CompetenceObjets = Niveau_CompetenceObjets(Index)
End Property
Public Property Let Definir_Niveau_CompetenceObjets(ByVal Index As Integer, ByVal Valeur As Long)
    Niveau_CompetenceObjets(Index) = Valeur
End Property
Public Property Get Definir_ExpObjets(ByVal Index As Integer) As Double
    Definir_ExpObjets = ExpObjets(Index)
End Property
Public Property Let Definir_ExpObjets(ByVal Index As Integer, ByVal Valeur As Double)
    ExpObjets(Index) = Valeur
End Property

'Propriétés des ressources.
Public Property Get Definir_ressources(ByVal Index As Long) As Double
    Definir_ressources = Ressources(Index)
End Property
Public Property Let Definir_ressources(ByVal Index As Long, ByVal Valeur As Double)
    Ressources(Index) = Valeur
End Property
Public Property Get Definir_MaxRessources(ByVal Index As Long) As Double
    Definir_MaxRessources = MaxRessources(Index)
End Property
Public Property Let Definir_MaxRessources(ByVal Index As Long, ByVal Valeur As Double)
    MaxRessources(Index) = Valeur
End Property
Public Property Get Definir_RessourcesSelectionnes(ByVal Index As Integer) As Double
    Definir_RessourcesSelectionnes = RessourcesSelectionnes(Index)
End Property
Public Property Let Definir_RessourcesSelectionnes(ByVal Index As Integer, ByVal Valeur As Double)
    RessourcesSelectionnes(Index) = Valeur
End Property

'Propriétés de l'inventaire.
Public Property Get Objet_Equipes_Type(ByVal Index As Integer) As Integer
    Objet_Equipes_Type = ObjetEquipesType(Index)
End Property
Public Property Let Objet_Equipes_Type(ByVal Index As Integer, ByVal Valeur As Integer)
    ObjetEquipesType(Index) = Valeur
End Property
Public Property Get Objet_Inventaire_Type(ByVal Index As Integer) As Integer
    Objet_Inventaire_Type = ObjetInventaireType(Index)
End Property
Public Property Let Objet_Inventaire_Type(ByVal Index As Integer, ByVal Valeur As Integer)
    ObjetInventaireType(Index) = Valeur
End Property
Public Property Get Objet_Inventaire_Nombre(ByVal Index As Integer) As Integer
    Objet_Inventaire_Nombre = ObjetInventaireNombre(Index)
End Property
Public Property Let Objet_Inventaire_Nombre(ByVal Index As Integer, ByVal Valeur As Integer)
    ObjetInventaireNombre(Index) = Valeur
End Property

'Propriétés du bestiaire.
Public Property Get Definir_Bestiaire(ByVal Index As Integer) As Integer
    Definir_Bestiaire = Bestiaire(Index)
End Property
Public Property Let Definir_Bestiaire(ByVal Index As Integer, ByVal Valeur As Integer)
    Bestiaire(Index) = Valeur
End Property

'Propriétés des apparences.
Public Property Get Definir_ApparenceDecalageX(ByVal Direction As Integer) As Integer
    Definir_ApparenceDecalageX = ApparenceDecalageX(Direction)
End Property
Public Property Let Definir_ApparenceDecalageX(ByVal Direction As Integer, ByVal Valeur As Integer)
    ApparenceDecalageX(Direction) = Valeur
End Property
Public Property Get Definir_ApparenceDecalageY(ByVal Direction As Integer) As Integer
    Definir_ApparenceDecalageY = ApparenceDecalageY(Direction)
End Property
Public Property Let Definir_ApparenceDecalageY(ByVal Direction As Integer, ByVal Valeur As Integer)
    ApparenceDecalageY(Direction) = Valeur
End Property
Public Property Get Definir_ApparenceAvantPlan(ByVal Direction As Integer) As Boolean
    Definir_ApparenceAvantPlan = ApparenceAvantPlan(Direction)
End Property
Public Property Let Definir_ApparenceAvantPlan(ByVal Direction As Integer, ByVal Valeur As Boolean)
    ApparenceAvantPlan(Direction) = Valeur
End Property

'Propriétés de la maison.
Public Property Get Maison_TypeBatiment() As Integer
    Maison_TypeBatiment = MaisonTypeBatiment
End Property
Public Property Let Maison_TypeBatiment(ByVal Valeur As Integer)
    MaisonTypeBatiment = Valeur
    If Parametres.Batiment_Fabrique(Valeur) Then
        If Maison_Nombre_Stock_Objet <> Parametres.Batiment_NombreObjetsFabriques(Valeur) - 1 Then
            Maison_Nombre_Stock_Objet = Parametres.Batiment_NombreObjetsFabriques(Valeur) - 1
        End If
    End If
End Property
Public Property Get Maison_Nombre_Stock_Objet() As Integer
    Maison_Nombre_Stock_Objet = MaisonNombreStockObjet
End Property
Public Property Let Maison_Nombre_Stock_Objet(ByVal Valeur As Integer)
    MaisonNombreStockObjet = Valeur
    ReDim MaisonStockObjets(Valeur)
End Property
Public Property Get MaisonStock_Objet(ByVal Index As Integer) As Long
    MaisonStock_Objet = MaisonStockObjets(Index)
End Property
Public Property Let MaisonStock_Objet(ByVal Index As Integer, ByVal Valeur As Long)
    MaisonStockObjets(Index) = Valeur
End Property
Public Property Get MaisonObjet_Inventaire_Type(ByVal Index As Integer) As Integer
    MaisonObjet_Inventaire_Type = MaisonObjetInventaireType(Index)
End Property
Public Property Let MaisonObjet_Inventaire_Type(ByVal Index As Integer, ByVal Valeur As Integer)
    MaisonObjetInventaireType(Index) = Valeur
End Property
