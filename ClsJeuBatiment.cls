VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsJeuBatiment"
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

Const CoefDepartVie As Double = 0.09
Const VitesseRecuperationVie As Double = 0.1
Const CoefRecyclage As Double = 0.5 'Quantité de ressources récupérées à la démolition de la maison.

Public Numero As Long

'Position actuelle du batiment.
Public PositionX As Long
Public PositionY As Long

Public Largeur As Integer
Public Hauteur As Integer

Public EntreePositionX As Integer
Public EntreePositionY As Integer

Public Construit As Boolean
Public EnConstruction As Boolean

Public Vie As Double
Public RecuperationVie As Double
Public MaxVie As Long

Public Argent As Double
Private Stock As Double
Public MaxStock As Double
Private StockObjets() As Double
Private NombreStockObjet As Integer

Private NombreObjetsInventaire As Integer
Private ObjetsInventaire() As Integer

Public Magasin As Boolean
Public Fabrique As Boolean
Public Service As Integer
Public Marche As Integer

Private NumeroAnimation As Integer

Public TypeBatiment As Integer
Public Tir As Boolean 'On peut tirer depuis ce bâtiment sans en sortir.

Private Sub Class_Initialize()
    NombreObjetsInventaire = Parametres.PersosNombreObjetsInventaireMaison
    ReDim ObjetsInventaire(NombreObjetsInventaire - 1)
    ReDim StockObjets(0)
    'TypeBatiment = -1
End Sub

Public Sub Creer(ByVal X As Long, _
                 ByVal Y As Long, _
                 ByVal Batiment As Integer, _
                 ByVal IndicePerso As Long, _
                 ByVal Parametres As ClsJeuParam)
    Dim i As Long
    If Parametres.Batiment_CreerProprietaire(Batiment) >= 0 And Not Persos(IndicePerso).Invocation Then
        'C'est une invoc qui s'occupe de cette maison.
        i = Creer_Invocation(Parametres.Batiment_CreerProprietaire(Batiment), 1, Numero, -1)
        Maisons(i).Creer X, Y, Batiment, i, Parametres
        'Persos(i).Rentrer_Maison Me
    Else
        'C'est le joueur qui change de maison.
        'On démolit la maison précédente.
        If Construit Then
            Demolir Persos(IndicePerso), Parametres, True
        End If
        
        'On fait naître le batiment.
        PositionX = X
        PositionY = Y
        Definir_TypeBatiment = Batiment
        Tir = Parametres.Batiment_Tir(Batiment)
        
        Construit = False
        'MaxVie = Parametres.Batiment_MaxVie(Batiment) * Perso.Definir_Carac_CompetenceSpeciales(0) / Parametres.PersosNiveauSpecialesDepart
        MaxVie = ((Persos(IndicePerso).Definir_Carac_CompetenceSpeciales(0) - Parametres.PersosNiveauSpecialesDepart) / Parametres.PersosNiveauSpecialesDepart * 0.5 + 1) * Parametres.Batiment_MaxVie(Batiment)
        Vie = MaxVie * CoefDepartVie
        
        Argent = 0
        Stock = 0
        MaxStock = Int(Parametres.Batiment_MaxStockVente(Batiment) * Persos(IndicePerso).Definir_Niveau_CompetenceSpeciales(0) / Parametres.PersosNiveauSpecialesDepart)
        
        
        For i = 0 To NombreObjetsInventaire - 1
            ObjetsInventaire(i) = -1
        Next i
        'EtatFabrication = 0
        
        Persos(IndicePerso).EtatFabricationObjet = 0
        Persos(IndicePerso).Visible = True
        Persos(IndicePerso).Rentrer_Maison Me
    End If
    If Not Persos(IndicePerso).Invocation Then
        'Le joueur paie son coût.
        Persos(IndicePerso).Definir_Argent = Persos(IndicePerso).Definir_Argent - Parametres.Batiment_PrixArgent(Batiment)
        For i = 0 To Parametres.NombreRessources - 1
            Persos(IndicePerso).Definir_ressources(i) = Persos(IndicePerso).Definir_ressources(i) - Parametres.Batiment_PrixRessource(Batiment, i)
        Next i
    End If
End Sub

Public Property Let Definir_TypeBatiment(ByVal Batiment As Integer)
    Largeur = Parametres.Batiment_Largeur(Batiment)
    Hauteur = Parametres.Batiment_Hauteur(Batiment)
    EntreePositionX = Parametres.Batiment_EntreePositionX(Batiment)
    EntreePositionY = Parametres.Batiment_EntreePositionY(Batiment)
    TypeBatiment = Batiment
    Magasin = Parametres.Batiment_Magasin(TypeBatiment)
    Fabrique = Parametres.Batiment_Fabrique(TypeBatiment)
    Service = Parametres.Batiment_Service(TypeBatiment)
    Marche = Parametres.Batiment_Marche(TypeBatiment)
    
    If Parametres.Batiment_Fabrique(TypeBatiment) Then
        If Definir_Nombre_Stock_Objet <> Parametres.Batiment_NombreObjetsFabriques(TypeBatiment) - 1 Then
            Definir_Nombre_Stock_Objet = Parametres.Batiment_NombreObjetsFabriques(TypeBatiment) - 1
        End If
    Else
        If Definir_Nombre_Stock_Objet <> 0 Then
            Definir_Nombre_Stock_Objet = 0
        End If
    End If
End Property

Public Sub Restauration()
    'Le chateau récupère.
    'Cette fonction doit être appelée à chaque tour.
    If RecuperationVie >= VitesseRecuperationVie Then
        Me.Definir_Vie = Vie + VitesseRecuperationVie
        RecuperationVie = RecuperationVie - VitesseRecuperationVie
    Else
        RecuperationVie = 0
    End If
End Sub

Public Sub Demolir(ByRef Perso As ClsJeuPerso, _
                   ByRef Parametres As ClsJeuParam, _
                   Optional ByVal MaisonPrincipale As Boolean)
    Dim i As Long
    If Perso.DansUneMaison And Persos(Perso.IndiceMaison).NumeroChef = Perso.Numero And Not MaisonPrincipale Then
        Maisons(Perso.IndiceMaison).Demolir Persos(Perso.IndiceMaison), Parametres
        If Persos(Perso.IndiceMaison).Vivant Then
            Perso_Tuer_Perso Perso, Persos(Perso.IndiceMaison)
            Persos(Perso.IndiceMaison).Vivant = False
        End If
    ElseIf MaxVie > 0 Then
        Perso.Definir_Argent = Perso.Definir_Argent + Parametres.Batiment_PrixArgent(TypeBatiment) * CoefRecyclage * Vie / MaxVie
        For i = 0 To Parametres.NombreRessources - 1
            If Parametres.Batiment_PrixRessource(TypeBatiment, i) > 0 Then
                Perso.Definir_ressources(i) = Perso.Definir_ressources(i) + _
                                              Parametres.Batiment_PrixRessource(TypeBatiment, i) * _
                                              CoefRecyclage * _
                                              Vie / MaxVie
            End If
        Next i
        If Perso.DansSaMaison Then
            Perso.DansSaMaison = False
            Perso.Action = 0
        End If
        Perso_Demolir_Batiment Perso, Me
        Definir_Vie = 0
        'Perso.Calculer_Apparence Me, Parametres
    End If
    'Efface les objets.
    For i = 0 To UBound(StockObjets())
        StockObjets(i) = 0
    Next i
End Sub

Public Sub Transfert_Argent_Prendre(ByVal Quantite As Double, _
                                    ByRef Perso As ClsJeuPerso, _
                                    Optional TransfertMax As Boolean)
    If TransfertMax Or Quantite > Argent Then
        Quantite = Argent
    End If
    Perso.Definir_Argent = Perso.Definir_Argent + Quantite
    Argent = Argent - Quantite
End Sub

Public Sub Transfert_Argent_Deposer(ByVal Quantite As Double, _
                                    ByRef Perso As ClsJeuPerso, _
                                    Optional TransfertMax As Boolean)
    If TransfertMax Or Quantite > Perso.Definir_Argent Then
        Quantite = Perso.Definir_Argent
    End If
    Perso.Definir_Argent = Perso.Definir_Argent - Quantite
    Argent = Argent + Quantite
End Sub

Public Sub Transfert_Stock_Prendre(ByVal Quantite As Double, _
                                   ByRef Perso As ClsJeuPerso, _
                                   ByRef Parametres As ClsJeuParam, _
                                   Optional TransfertMax As Boolean)
    If TransfertMax Or Quantite > Stock Then
        Quantite = Stock
    End If
    'Ne prend pas plus de ressources que le joueur peut prendre.
    If Quantite > Perso.Definir_MaxRessources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) - Int(Perso.Definir_ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment))) Then
       Quantite = Perso.Definir_MaxRessources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) - Int(Perso.Definir_ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment)))
    End If
    Quantite = Int(Quantite)
    'On transfert.
    Perso.Definir_ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) = Perso.Definir_ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) + Quantite
    Stock = Stock - Quantite
End Sub

Public Sub Transfert_Stock_Deposer(ByVal Quantite As Double, _
                                   ByRef Perso As ClsJeuPerso, _
                                   ByRef Parametres As ClsJeuParam, _
                                   Optional TransfertMax As Boolean)
    If TransfertMax Or Quantite > Perso.Definir_ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) Then
        Quantite = Perso.Definir_ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment))
    End If
    'On ne dépose pas plus de ressources que le batiment peut contenir.
    If Quantite > MaxStock - Stock Then
       Quantite = MaxStock - Stock
    End If
    Quantite = Int(Quantite)
    'On transfert.
    Perso.Definir_ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) = Perso.Definir_ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) - Quantite
    Stock = Stock + Quantite
End Sub

'Public Sub Transfert_Objet_Prendre(ByVal IndiceObjet As Integer, _
                                   ByRef Perso As ClsJeuPerso, _
                                   ByRef Parametres As ClsJeuParam)
    'On met l'objet dans l'inventaire.
    
'    Stock = Stock - Quantite
'End Sub


Public Sub Acheter(ByVal Quantite As Double, _
                   ByRef Perso As ClsJeuPerso, _
                   ByRef Parametres As ClsJeuParam, _
                   Optional AcheterMax As Boolean)
    'Achat de ressources.
    Dim Prix As Integer
    If AcheterMax Or _
       Perso.Definir_Argent < Quantite * Perso.Prix_Marchandee(Parametres.Ressources_PrixVente(Parametres.Batiment_RessourceFabrique(TypeBatiment)) * Parametres.CoefficientAchat, Parametres) Then
        Quantite = Perso.Definir_Argent / Perso.Prix_Marchandee(Parametres.Ressources_PrixVente(Parametres.Batiment_RessourceFabrique(TypeBatiment)) * Parametres.CoefficientAchat, Parametres)
    End If
    If Quantite > Stock Then
        Quantite = Stock
    End If
    If Quantite > Perso.Definir_MaxRessources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) - Int(Perso.Definir_ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment))) Then
        Quantite = Perso.Definir_MaxRessources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) - Int(Perso.Definir_ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment)))
    End If
    Quantite = Int(Quantite)
    Perso.Definir_ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) = Perso.Definir_ressources(Parametres.Batiment_RessourceFabrique(TypeBatiment)) + Quantite
    Stock = Stock - Quantite
    Prix = Quantite * Perso.Prix_Marchandee(Parametres.Ressources_PrixVente(Parametres.Batiment_RessourceFabrique(TypeBatiment)) * Parametres.CoefficientAchat, Parametres)
    Perso.Definir_Argent = Perso.Definir_Argent - Prix
    Argent = Argent + Prix
    'Affiche l'argent dépensé.
    AffDegat.Ajouter_Degat PositionX + Largeur / 2, PositionY, Prix, , 2
    'Compétence marchandage.
    Perso.Gestion_Niveau_Comp_Speciale 3, Prix, Parametres
End Sub

Public Function Service_Utilisable(ByVal Perso As ClsJeuPerso, _
                                   ByVal Parametres As ClsJeuParam) As Boolean
    'Renvoie si le perso peut utiliser le service.
    'Le temples ne ressuscitent pas les personnes vivantes.
    'Le joueur doit avoir assez d'argent ou être propriétaire de la maison.
    Service_Utilisable = Not (Service = 1 And Perso.Vivant) And _
                         (Perso.Definir_Argent >= Perso.Prix_Marchandee(Parametres.Service_Prix(Service), Parametres) Or _
                         Perso.Numero = Numero) And _
                         Stock >= 1
    'If Perso.Numero > 0 Then
    '    Debug.Print Service & ", " & Numero & ", " & Int(Stock) & ", " & Int(Me.Definir_Stock)
    'End If
End Function

Public Sub Service_Utiliser(ByRef Perso As ClsJeuPerso, _
                            ByVal Parametres As ClsJeuParam, _
                            Optional ByVal Score As Double, _
                            Optional ByVal Experience As Boolean)
    Dim Prix As Integer
    Select Case Service
    Case 1:
        Perso.Vivant = True
        Perso.Vie = Parametres.Service_Coefficient(Service) * Perso.Definir_MaxVie
        Stock = Stock - 1
    Case 2:
        Argent = Argent + Score
        If Experience Then
            Perso.Gestion_Niveau_Prime Score, Parametres
        End If
        'Affiche la quantité d'argent gagné à l'écran.
        AffDegat.Ajouter_Degat PositionX + Largeur / 2, PositionY, Score, , 2
    End Select
    'If Stock < 1 Then
    '    Debug.Print Service_Utilisable(Perso, Parametres)
    'End If
    If Not (Perso.Numero = Numero) Then
        'On paye le service utilisé.
        Prix = Perso.Prix_Marchandee(Parametres.Service_Prix(Service), Parametres)
        Perso.Definir_Argent = Perso.Definir_Argent - Prix
        Argent = Argent + Prix
        'Affiche l'argent dépensé.
        AffDegat.Ajouter_Degat PositionX + Largeur / 2, PositionY, Prix, , 2
        'Compétence marchandage.
        Perso.Gestion_Niveau_Comp_Speciale 3, Prix, Parametres
    End If
    If Parametres.Service_Effet(Service) >= 0 Then
        'EffetsMaisons(Numero).Activer_Effet Parametres.Service_Effet(Service), PositionX, PositionY, Largeur, Hauteur, Parametres
        Activer_Effet Parametres.Service_Effet(Service), PositionX, PositionY, Largeur, Hauteur, Parametres
    End If
End Sub


Public Property Get Visible() As Boolean
    Visible = Vie > 0
End Property

Public Property Get Definir_Vie() As Double
    If Vie + RecuperationVie > MaxVie Then
        Definir_Vie = MaxVie
    Else
        Definir_Vie = Vie + RecuperationVie
    End If
End Property

Public Property Let Definir_Vie(ByVal Valeur As Double)
    Vie = Valeur
    If Vie >= MaxVie Then
        Vie = MaxVie
        EnConstruction = False
        If Not Construit Then
            Construit = True
            If Marche Then
                Fiefs(Persos(Numero).NumeroFief).Marches_Ajouter Numero
            End If
            'Le joueur reçoit des objets.
            If Parametres.Batiment_GainObjet(TypeBatiment) >= 0 Then
                Persos(Numero).Objet_Ajouter_Inventaire Parametres.Batiment_GainObjet(TypeBatiment), Parametres, Commentaires, True
            End If
        End If
    ElseIf Vie <= 0 Then
        EnConstruction = False
        RecuperationVie = 0 'Les chateaux ne se régénèrent plus.
        If Construit Then
            Construit = False
            If Marche Then
                Fiefs(Persos(Numero).NumeroFief).Marches_Supprimer Numero
            End If
        End If
        'Le joueur perd des objets.
        If Parametres.Batiment_PerteObjet(TypeBatiment) >= 0 Then
            Persos(Numero).Objet_Supprimer Parametres.Batiment_PerteObjet(TypeBatiment)
        End If
    Else
        EnConstruction = True
    End If
End Property

Public Property Get Etat(ByVal Parametres As ClsJeuParam) As Integer
    If Construit Then
        If (Magasin Or Fabrique Or (Service And MaxStock > 0)) And Stock < 1 Then
            'Magasin Vide.
            Etat = Parametres.Batiment_NombreEtapeFabrication(TypeBatiment) - 1
        Else
            Etat = Parametres.Batiment_NombreEtapeFabrication(TypeBatiment)
        End If
    Else
        Etat = Int(Vie / MaxVie * Parametres.Batiment_NombreEtapeFabrication(TypeBatiment))
    End If
End Property

Public Sub Actualiser_NumeroAnimation()
    NumeroAnimation = NumeroAnimation + 1
    If NumeroAnimation > Parametres.Batiment_NombreEtapeAnimation(TypeBatiment) - 1 Then
        NumeroAnimation = 0
    End If
End Sub

Public Property Get Definir_NumeroAnimation(ByVal Parametres As ClsJeuParam) As Integer
    Definir_NumeroAnimation = NumeroAnimation
End Property

Public Property Get Definir_Stock() As Double
    Definir_Stock = Stock
End Property
Public Property Let Definir_Stock(ByVal Valeur As Double)
    Stock = Valeur
    If Stock > MaxStock Then
        Stock = MaxStock
    End If
End Property

Public Property Get Stock_Objets(ByVal Indice As Integer) As Double
    Stock_Objets = StockObjets(Indice)
End Property
Public Property Let Stock_Objets(ByVal Indice As Integer, ByVal Valeur As Double)
    Stock = Stock + Valeur - StockObjets(Indice)
    StockObjets(Indice) = Valeur
End Property

Public Property Get Definir_Nombre_Stock_Objet() As Integer
    Definir_Nombre_Stock_Objet = NombreStockObjet
End Property
Public Property Let Definir_Nombre_Stock_Objet(ByVal Valeur As Integer)
    NombreStockObjet = Valeur
    ReDim StockObjets(Valeur)
End Property

Public Property Get Definir_Nombre_Objet_Inventaire() As Integer
    Definir_Nombre_Objet_Inventaire = NombreObjetsInventaire
End Property

Public Property Get Definir_ObjetsInventaire(ByVal Index As Integer) As Integer
    Definir_ObjetsInventaire = ObjetsInventaire(Index)
End Property
Public Property Let Definir_ObjetsInventaire(ByVal Index As Integer, ByVal Valeur As Integer)
    ObjetsInventaire(Index) = Valeur
End Property

Public Property Get Definir_Nombre_Objets_Deposes() As Integer
    Dim i As Integer
    For i = 0 To NombreObjetsInventaire - 1
        If ObjetsInventaire(i) >= 0 Then
            Definir_Nombre_Objets_Deposes = Definir_Nombre_Objets_Deposes + 1
        End If
    Next i
End Property
