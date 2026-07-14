Attribute VB_Name = "ModIA"
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
'DefLng A-Z
Option Explicit
Option Compare Binary

Public Const NiveauIA = 1 '0.5

Const VitesseIA As Single = 0.01
Const VitesseRechercheIA As Single = 0.6 '0.2

Const NombreRecherchesEmplacementMaison = 10

Const VitesseEquipement = 1
Const VitesseEquipementCombat = 5

Const ChanceReparerMaison = 0.1

'Const MaxDureeIA As Double = 0.05

Const DistancePalierRecherchePerso As Single = 48
Const DistancePalierRechercheMaison As Single = 96 '64

Const DistanceAdversaireMin As Single = 64
Const DistanceAdversaireMax As Single = 640

Const DistanceTresorMax As Single = 724

Const QuantiteMaxTransfert  As Long = 10 '5
Const MatierePremiereMin As Long = 1 '5 '10

Const ArgentSeuilPauvre As Long = 300
Const ArgentSeuilRiche As Long = 500

Const IndiceHerbe As Integer = 0
Const IndiceFruit As Integer = 1
Const IndiceBois As Integer = 2
Const IndicePierre As Integer = 12

'Variables utilisées dans les fonctions IA.
Public DistanceObstacle() As Long
Public IndiceObstacle As Long

Public DistanceRessource() As Long
Public DistanceTresor() As Long
Public DistanceMaison() As Long
Public DistanceAdversaire() As Long
Public DistanceMaisonEnnemi() As Long
Public DistanceChateauEnnemi() As Long

Public IndicePerso As Long
Public IndiceMaison() As Long
Public IndiceAdversaire() As Long
Public IndiceMaisonEnnemi() As Long
Public IndiceChateauEnnemi() As Long
Private AnglePatrouille() As Single
Public IndiceObjet As Integer

Public BatimentAConstruire() As Integer 'Indice du type de batiment que l'IA veut construire.
Public RessourceAFabriquer() As Integer 'Indice de la ressource que l'IA veut fabriquer.
Public ObjetAFabriquer() As Integer 'Indice de l'objet que l'IA veut fabriquer.
Public QuantiteAFabriquer() As Integer 'Indice de la ressource que l'IA veut fabriquer.
Public QuitterBatiment() As Boolean

'Information sur les batiments.
Private BatimentNiveauMax As Integer

Public LPerso As Long
Private NombreIAParTour As Long
'Public NombreActionsParTour As Long
Private NombreRechercheParTour As Long

Private NombreBatimentsDeNiveau() As Integer

'Variables pour des infos temporaires.
Private X As Long
Private Y As Long
Private Z As Long
Private SriptRunner As MSScriptControl.ScriptControl
Private SurPerso1 As ClsSurPerso
Private SurPerso2 As ClsSurPerso
Public SurJeu As ClsSurJeu

Public Function IA_Init()
    Dim i As Long
    Dim j As Long
    Set SriptRunner = New MSScriptControl.ScriptControl
        SriptRunner.Language = "VBScript"
        SriptRunner.Reset
        SriptRunner.AllowUI = False
    Set SurPerso1 = New ClsSurPerso
    Set SurPerso2 = New ClsSurPerso
    Set SurJeu = New ClsSurJeu

    ReDim DistanceObstacle(UBound(Fiefs()))
    
    'ReDim IA_IndiceRessource(UBound(Persos()))
    
    Nombre_IAs = UBound(Persos()) + 1
    
    IndiceObjet = Rnd * (Persos(0).Nombre_Objets_Inventaire)
    If IndiceObjet = Persos(0).Nombre_Objets_Inventaire Then IndiceObjet = 0
    
    For i = 0 To LPerso - 1
        IA_InitPerso i
        AnglePatrouille(i) = Rnd * 6.28
    Next i
    
    'Mémorise les niveaux des batiments.
    For i = 0 To Parametres.NombreTypeBatiments - 1
        If Parametres.Batiment_Niveau(i) > BatimentNiveauMax Then
            BatimentNiveauMax = Parametres.Batiment_Niveau(i)
        End If
    Next i
    ReDim NombreBatimentsDeNiveau(BatimentNiveauMax + 1)
    For j = 0 To UBound(NombreBatimentsDeNiveau())
        For i = 0 To Parametres.NombreTypeBatiments - 1
            If Parametres.Batiment_Niveau(i) = j Then
                NombreBatimentsDeNiveau(j) = NombreBatimentsDeNiveau(j) + 1
            End If
        Next i
    Next j
End Function

Public Sub IA_InitPerso(ByVal i As Long)
    'Initialise les informations d'un personnage.
    DistanceMaison(i) = DistancePalierRechercheMaison
    BatimentAConstruire(i) = -1
    RessourceAFabriquer(i) = -1
    ObjetAFabriquer(i) = -1
    QuantiteAFabriquer(i) = 0
    If Persos(i).ChoisirIAScript Then
        Persos(i).JeuIA3.SetScriptFile App.Path & "\" & CheminIAScript & Persos(i).FichierIAScript
    End If
End Sub

Public Function Comportement_IA()
    Dim j As Long
    Dim Debut As Double
    
    Debut = Timer
    'Debug.Print LPerso
    For j = 0 To NombreIAParTour
        IndicePerso = IndicePerso + 1
        If IndicePerso > LPerso Then
            IndicePerso = 0
            'IndiceObjet = IndiceObjet + 1
            'If IndiceObjet > Persos(0).Nombre_Objets_Inventaire - 1 Then
            '    IndiceObjet = 0
            'End If
        End If
        'FrmMoteur2D.Actions_Persos IndicePerso
        Comportement_IA_Persos IndicePerso
        
        'If Abs(Timer - Debut) > MaxDureeIA Then
        '    j = NombreIAParTour
        'End If
    Next j
End Function

Public Sub Comportement_IA_Persos(ByVal i As Long, Optional ByVal Simple As Boolean)
    If Persos(i).IA Then
        If Persos(i).ChoisirIAScript Then
            SurPerso1.Init Persos(i)
            SurPerso2.Init Persos(Persos(i).IndicePerso)
            SriptRunner.Reset
            'SriptRunner.AddObject "Temp", Persos(i).Temp, True
            SriptRunner.AddObject "Me", SurPerso1, True
            SriptRunner.AddObject "Cible", SurPerso2, True
            SriptRunner.AddObject "Jeu", SurJeu, True
            'SriptRunner.AddObject "Chateaux", Chateaux(0), True
            SriptRunner.AddCode Persos(i).JeuIA3.GetScriptFile()
            SriptRunner.Run "Jouer"
        Else
            If Persos(i).Vivant Then
                'Choisit un équipement adapté.
                IA_Equiper i
                If Persos(i).IA_Berserk > 0 Then
                    'IA_ComportementBerserk i
                ElseIf Persos(i).ChoisirIA Then
                    'Comportement personnalisé.
                    Comportement_IA2 i, Simple
                Else
                    'Comportement des monstres.
                    If Persos(i).Invocation Then  'Cas où le monstre obéit à un maître.
                        IA_Recuperation_CaracSimple i
                        If Persos(i).Inactif Then
                        'If Persos(.NumeroChef).Vivant Then 'N'obéit que si son maître est vivant.
                            If Persos(i).IA_Ordre > 0 Then
                                'Obéit à l'ordre qu'on lui a affecté.
                                Select Case Persos(i).IA_Ordre
                                Case 1: 'Attaquer un personnage.
                                    If Not Persos(i).EnCombat Then
                                        If Persos(Persos(i).IA_Cible).Attaquable And _
                                           Persos(Persos(i).IA_Cible).NumeroEquipe <> Persos(i).NumeroEquipe Then
                                            Persos(i).Attaquer_Perso Persos(i).IA_Cible
                                        Else
                                            Persos(i).IA_Ordre = 0
                                        End If
                                    End If
                                Case 2: 'Attaquer un batiment.
                                    If Not Persos(i).EnCombat Then
                                        If Maisons(Persos(i).IA_Cible).Visible And _
                                           Persos(Persos(i).IA_Cible).NumeroEquipe <> Persos(i).NumeroEquipe Then
                                            Persos(i).Attaquer_Maison Persos(i).IA_Cible
                                        Else
                                            Persos(i).IA_Ordre = 0
                                        End If
                                    End If
                                Case 3: 'Attaquer un château.
                                    If Not Persos(i).EnCombat Then
                                        If Chateaux(Persos(i).IA_Cible).Visible Then
                                            Persos(i).Attaquer_Chateau Persos(i).IA_Cible
                                        Else
                                            Persos(i).IA_Ordre = 0
                                        End If
                                    End If
                                Case 4: 'Protège un personnage.
                                    If Not Persos(i).EnCombat Then
                                        IA_Garder_Perso i
                                    End If
                                End Select
                            ElseIf Persos(i).IA_Berserk = 0 Then
                                IA_Comportement_Soldats i
                            End If
                        'Else
                            'Fait ce qu'il veut.
                            'IA_Comportement_Monstre i
                        'End If
                        End If
                    Else
                        IA_Recuperer_Equipement_Monstre i 'Recharge son équipement au chateau.
                        IA_Comportement_Monstre i
                    End If
    
                End If
            Else
                IA_Mort_Equiper i
                If Persos(i).Invocation Then
                    'Les monstres qui obéissent à un maître ne font rien quand ils sont des fantômes.
                    If Maisons(Persos(i).Numero).Visible And Persos(Persos(i).NumeroChef).Vivant Then
                        'Les invocations avec une maisons peuvent ressusciter si leur chef est toujours en vie.
                        If Not Persos(i).DansSaMaison Then
                            IA_Rentrer_Maison i
                        End If
                    End If
                Else
                    'Comportement des fantômes.
                    'Se ressuscite au temple.
                    If (Persos(i).DansUneMaison Or Persos(i).DansSaMaison) And _
                       Maisons(Persos(i).IndiceMaison).Construit And _
                       Maisons(Persos(i).IndiceMaison).Service_Utilisable(Persos(i), Parametres) Then
                        Maisons(Persos(i).IndiceMaison).Service_Utiliser Persos(i), Parametres
                        IA_Sortir_Batiment i
                    'Va au temple.
                    ElseIf Persos(i).Action <> 1 And Persos(i).Action <> 201 Then
                        If Trouver_Maison_Proche(Persos(i), i, 0, 0, 1, X) And _
                           Not IA_Argent_Besoin(i) Then
                            If X >= 0 Then
                                IA_Aller_Maison i, X
                            End If
                        'Va au chateau.
                        ElseIf Chateaux(Persos(i).NumeroFief).Visible And _
                           Not (Persos(i).DansSonChateau) Then
                            IA_Rentrer_Chateau i, False
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Function Comportement_IA_Fiefs()
    'Gère le comportement des fiefs.
    Dim j As Long
    For j = 0 To UBound(Fiefs())
        Fiefs(j).Transaction_Ressources Parametres
        Fiefs(j).Restauration Chateaux(j)
        Chateaux(j).Restauration
    Next j
End Function

Private Sub IA_Comportement_Monstre(ByVal i As Long)
    With Persos(i)
    IA_Poursuivre_Fuyard i
    If .Inactif Then
        'If .Action <> 3 And .Action <> 500 Then IA_Recuperation_CaracSimple i
        IA_Recuperation_CaracSimple i
        If .Inactif Then IA_Agressif i
    ElseIf Rnd > 0.99 Then
        .Arreter
    End If
    End With
End Sub

'Fonction secondaires utilisée par l'IA.
'Private Sub IA_Comportement_Idiots(ByVal i As Long)
'    With Persos(i)
'    If Rnd > 0.985 Then
'    'If RND > 0.975 Then
'        Select Case Rnd * 25
'        Case 0:
'            IA_Deplacement_Aleatoire i
'        Case 1: 'S'arrête sans raison.
'            .Arreter
'        Case 2: 'Fait une pause.
'            .Arreter
'        Case 3: 'Rentre chez lui.
'            IA_Rentrer_Maison i
'        Case 4: 'Retourne au chateau.
'            IA_Rentrer_Chateau i
'        Case 5: 'Veut changer de maison.
'            Z = (Rnd * (Parametres.NombreTypeBatiments - 1))
'            If .PeutConstruire_Maison(Z, Parametres) Then
'                If Trouver_Emplacement_Maison(i, X, Y, Z) Then
'                    Maisons(i).Creer X, Y, Z, i, Parametres
'                    'IA_Rentrer_Maison i
'                End If
'            End If
'        Case 6: 'S'attaque à un ennemi.
'            If IA_Attaquer_Perso(i, X) Then
'                .Attaquer_Perso X
'            End If
'        Case 7: 'S'attaque à une maison.
'            If IA_Attaquer_Maison(i, X) Then
'                .Attaquer_Maison X
'            End If
'        Case Else: 'Va chercher des ressources.
'            'Va chercher du bois.
'            If .Definir_ressources(IndiceBois) < .Definir_MaxRessources(IndiceBois) Then
'                IA_Couper_Arbre i
'            ElseIf .Definir_ressources(IndicePierre) < .Definir_MaxRessources(IndicePierre) Then
'                IA_ChercherPierre i
'            'Va chercher des herbes.
'            ElseIf .Definir_ressources(IndiceHerbe) < .Definir_MaxRessources(IndiceHerbe) Then
'                IA_Ramasser_Herbe i
'            'Va chercher des pommes.
'            ElseIf .Definir_ressources(IndiceFruit) < .Definir_MaxRessources(IndiceFruit) Then
'                IA_Cueillir_Fruits i
'            End If
'        End Select
'    End If
'    End With
'End Sub
Public Sub IA_Aller_Tresor(ByVal i As Long, ByVal Indice As Long)
    Persos(i).Aller_Tresor Tresors(Indice), Indice, AffTresor.Largeur, AffTresor.Hauteur
    QuitterBatiment(i) = True
End Sub
Public Sub IA_Rentrer_Maison(ByVal i As Long)
    Persos(i).Rentrer_Maison Maisons(i)
    QuitterBatiment(i) = False
End Sub
Public Sub IA_Aller_Maison(ByVal i As Long, ByVal Indice As Long)
    If i = Indice Then 'C'est sa maison.
        IA_Rentrer_Maison i
    Else
        Persos(i).Aller_Maison Maisons(Indice), Rnd < ChanceReparerMaison
        QuitterBatiment(i) = False
    End If
End Sub
Public Sub IA_Rentrer_Chateau(ByVal i As Long, ByVal ChercherRessources As Boolean)
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim Nf As Long 'Variable temporaire enregistrant le numéro de fief du personnage.
    Dim Distance As Long
    Distance = Monde.Diagonale
    With Persos(i)
    Nf = .NumeroFief
    If ChercherRessources Then
        'Si le personnage cherche des ressources, autant aller au marché le plus proche.
        For j = 1 To Fiefs(Nf).Marches.Count
            k = Fiefs(Nf).Marches(j)
            If Distance_Positions(.PositionX, .PositionY, Maisons(k).PositionX, Maisons(k).PositionY) < Distance Then
                Distance = Distance_Positions(.PositionX, .PositionY, Maisons(k).PositionX, Maisons(k).PositionY)
                l = k
            End If
        Next j
        If Distance < Distance_Positions(.PositionX, .PositionY, Chateaux(Nf).PositionX, Chateaux(Nf).PositionY) Or Not Chateaux(Nf).Visible Then
            .Aller_Maison Maisons(l), Rnd < ChanceReparerMaison
        Else
            .Rentrer_Chateau Chateaux(Nf)
        End If
    Else
        .Rentrer_Chateau Chateaux(Nf)
    End If
    QuitterBatiment(i) = False
    End With
End Sub
Private Sub IA_Sortir_Batiment(ByVal i As Long)
    QuitterBatiment(i) = True
    'Persos(i).Aller_A Persos(i).PositionX, Persos(i).PositionY + Maisons(i).Hauteur / 2
End Sub

Private Function IA_Stock_Exedentaire(ByVal i As Long) As Boolean
    'Renvoie vrai si le personnage a plus de la moitié de ses ressources.
    'IA_Stock_Exedentaire = _
        (Persos(i).Definir_Ressources(Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment)) > _
        Persos(i).Definir_MaxRessources(Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment)) _
        * 0.5 And _
        Not Maisons(i).Magasin) Or _
        (Maisons(i).Definir_Stock > Maisons(i).MaxStock - 1 And Maisons(i).Magasin)
    IA_Stock_Exedentaire = _
        Persos(i).Definir_ressources(Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment) > 1 And _
        Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment)) >= 0 Or _
        (Maisons(i).Definir_Stock > Maisons(i).MaxStock - 1 And Maisons(i).Magasin)
End Function
Private Function IA_Argent_Besoin(ByVal i As Long) As Boolean
    IA_Argent_Besoin = Persos(i).Definir_Argent < ArgentSeuilPauvre * (Fiefs(Persos(i).NumeroFief).Epoque + 1)
End Function
Private Function IA_Argent_Riche(ByVal i As Long) As Boolean
    IA_Argent_Riche = Persos(i).Definir_Argent > ArgentSeuilRiche * (Fiefs(Persos(i).NumeroFief).Epoque + 1)
End Function

Public Sub IA_Deplacement_Aleatoire(ByVal i As Long)
    Dim X As Long
    Dim Y As Long
    Dim nx As Long, ny As Long
    Dim H As Integer
    Dim NbRelance As Integer 'Nombre de fois que l'IA cherche un endroit où aller.
    'Déplacement aléatoire à proximité.
    With Persos(i)
    nx = Monde.Largeur * AffTerrain.Largeur - .Largeur
    ny = Monde.Hauteur * AffTerrain.Hauteur - .Hauteur
    If .Volant Then
        X = .PositionX + (Rnd * 128 - 64) * AffTerrain.Largeur
        Y = .PositionY + (Rnd * 96 - 48) * AffTerrain.Hauteur
        'On évite de sortir de la carte.
        If X < 0 Then
            X = 0
        ElseIf X > nx Then
            X = nx
        End If
        If Y < 0 Then
            Y = 0
        ElseIf Y > ny Then
            Y = ny
        End If
    Else
        H = .Hauteur - AffTerrain.Hauteur
        Do
            X = .PositionX + (Rnd * 128 - 64) * AffTerrain.Largeur
            Y = .PositionY + (Rnd * 96 - 48) * AffTerrain.Hauteur
            'On évite de sortir de la carte.
            If X < 0 Then
                X = 0
            ElseIf X > nx Then
                X = nx
            End If
            If Y < 0 Then
                Y = 0
            ElseIf Y > ny Then
                Y = ny
            End If
            NbRelance = NbRelance + 1
            'Evite d'aller sur une mauvaise surface.
        Loop Until Persos(i).Volant Or Monde.CoefVitesseSimple(X, Y + H) >= 1 Or NbRelance > 40
    End If
    
    .Aller_A X, Y
    End With
End Sub
Public Sub IA_Recuperer_Equipement_Monstre(ByVal i As Long)
    Dim j, k As Integer
    Dim Temp As Integer
    With Persos(i)
    If .DansSonChateau And Not QuitterBatiment(i) Then
'        'Récupère ses ressources.
'        For j = 0 To Parametres.NombreRessources - 1
'            If .Definir_Ressources(j) < Parametres.Race_Ressources(.Race, j) Then
'                .Definir_Ressources(j) = Parametres.Race_Ressources(.Race, j)
'            End If
'        Next j
'        'Récupère son équipement.
'        For j = 0 To Parametres.Race_NombreEquipement - 1
'            'If Parametres.Monstre_Objet(.Race - 1, j) > 0 Then
'            'Temp = .Objet_NombreType(j)
'            If .Objet_NombreType(j) = 0 Then
'            'If Temp < Parametres.Monstre_Objet(.Race - 1, j) Then
'                'For k = 0 To Parametres.Monstre_Objet(.Race - 1, j) - Temp - 1
'                .Objet_Ajouter_Inventaire Parametres.Race_Equipement(j), Parametres, Commentaires, .IA_Equiper>0
'                'Next k
'            End If
'            'End If
'        Next j
'        'Récupère de la vie
'        .RecuperationVie = .Definir_MaxVie - .Vie
'        .RecuperationEnergie = .Definir_MaxEnergie - .Energie
'        .RecuperationMagie = .Definir_MaxMagie - .Magie
'        .RecuperationMoral = .Definir_MaxMoral - .Moral
        'Restaure le chateau.
        IA_Restaurer_Chateau i
        QuitterBatiment(i) = True
    End If
    End With
End Sub

Public Sub IA_Agressif(ByVal i As Long)
    Dim X As Long
    If Persos(i).IA_Pacifique Then
        Select Case Rnd
            Case Is < 0.1:
            If Persos(i).Invocation Then
                Persos(i).Aller_A Persos(Persos(i).NumeroChef).DirectionX + Persos(i).ChefDecalageX, _
                                  Persos(Persos(i).NumeroChef).DirectionY + Persos(i).ChefDecalageY
            Else
                IA_Rentrer_Chateau i, False
            End If
        Case Is < 0.2:
            IA_Deplacement_Aleatoire i
        End Select
    Else
        Select Case Rnd
        Case Is < 0.05:
            If IA_Attaquer_Chateau(i, X) Then Persos(i).Attaquer_Chateau X
        Case Is < 0.1:
            If Persos(i).Invocation Then
                Persos(i).Aller_A Persos(Persos(i).NumeroChef).DirectionX + Persos(i).ChefDecalageX, _
                                  Persos(Persos(i).NumeroChef).DirectionY + Persos(i).ChefDecalageY
            Else
                IA_Rentrer_Chateau i, False
            End If
        Case Is < 0.2:
            IA_Deplacement_Aleatoire i
        Case Is < 0.4:
            If IA_Attaquer_Maison(i, X) Then Persos(i).Attaquer_Maison X
        Case Else
            If IA_Attaquer_Perso(i, X) Then Persos(i).Attaquer_Perso X
        End Select
    End If
End Sub

Private Sub IA_Comportement_Soldats(ByVal i As Long)
    'Le soldat agit selon l'attitude qui lui a été donné.
    Dim j As Long
    With Persos(i)
    'If .Inactif Then
        Select Case .Attitude
        Case 0:
            'Mode couverture. (Suit et agit selon son chef.)
            'If Not Persos(.NumeroChef).Vivant Then
            '    If .Action <> 3 And .Action <> 500 Then IA_Recuperation_CaracSimple i
            '    If .Inactif Then IA_Agressif i
            'ElseIf Not .EnCombat Then
            '    Select Case Persos(.NumeroChef).Action
            '    Case 300: .Attaquer_Perso Persos(.NumeroChef).IndicePerso
            '    Case 301: .Attaquer_Maison Persos(.NumeroChef).IndiceMaison
            '    Case 302: .Attaquer_Chateau Persos(.NumeroChef).IndiceChateau
            '    Case Else:
            '    'If Persos(.NumeroChef).DansSaMaison Or DansUneMaison Then
                    If Not .EnCombat Then
                    .Aller_A Persos(.NumeroChef).DirectionX + .ChefDecalageX, _
                                    Persos(.NumeroChef).DirectionY + .ChefDecalageY
                    End If
            '    'End If
            '    End Select
            'End If
        Case 1:
            'Mode en attente.
            If Not .EnCombat Then
            .Aller_A .ChefPositionX, .ChefPositionY
            End If
        Case 2:
            'Mode en défense.
            j = Trouver_Ennemi(i, .ChefPositionX, .ChefPositionY)
            If j >= 0 Then
'                If Not .EnCombat Then
                If Sqr((.ChefPositionX - Persos(j).PositionX) ^ 2 + (.ChefPositionY - Persos(j).PositionY) ^ 2) <= 384 Then
                    .Attaquer_Perso j
                Else
                    .Aller_A .ChefPositionX, .ChefPositionY
                End If
'                Else
'                    .Aller_A .ChefPositionX, .ChefPositionY
'                End If
            Else
                .Aller_A .ChefPositionX, .ChefPositionY
            End If
        Case 3:
            'Mode en attaque.
            'If .Action <> 3 And .Action <> 500 Then IA_Recuperation_CaracSimple i
            'If .Inactif Then IA_Agressif i
            If Not .EnCombat Then
                j = Trouver_Ennemi(i)
                If j >= 0 Then
                    If Distance_Persos(Persos(i), Persos(j)) <= DistanceAdversaireMax Then
                        .Attaquer_Perso j
                    End If
                End If
            End If
        Case 4:
            'Mode en bersek.
            'If Persos(.NumeroChef).Vivant Then
                IA_ComportementBerserk i
            'Else
            '    If .Action <> 3 And .Action <> 500 Then IA_Recuperation_CaracSimple i
            '    If .Inactif Then IA_Agressif i
            'End If
        End Select
    'End If
    End With
End Sub

Public Sub IA_Garder_Perso(ByVal i As Long)
    Dim X As Long
    If Not Persos(i).EnCombat Then
        If Persos(Persos(i).IA_Cible).Vivant Then
            If IA_Attaquer_Assaillant_Perso(i, Persos(i).IA_Cible, X) Then
                Persos(i).Attaquer_Perso X
            Else 'If Persos(i).Inactif Then
                IA_Patrouiller_Cercle i, Persos(Persos(i).IA_Cible), 128
'                If Persos(Persos(i).IA_Cible).DansUnChateau Or Persos(Persos(i).IA_Cible).DansSonChateau Then
'                    With Chateaux(Persos(Persos(i).IA_Cible).IndiceChateau)
'                    Select Case Rnd
'                    Case Is < 0.25: 'Va au coin en haut à gauche du château.
'                        Persos(i).Aller_A .PositionX - Persos(i).Largeur + Int(Rnd * (.Largeur + Persos(i).Largeur)), .PositionY - Persos(i).Hauteur
'                    Case Is < 0.5: 'Va au coin en haut à droite du château.
'                        Persos(i).Aller_A .PositionX + .Largeur, .PositionY - Persos(i).Hauteur + Int(Rnd * (.Hauteur + Persos(i).Hauteur))
'                    Case Is < 0.75: 'Va au coin en bas à droite du château.
'                        Persos(i).Aller_A .PositionX + .Largeur - Int(Rnd * (.Largeur + Persos(i).Largeur)), .PositionY + .Hauteur
'                    Case Else  'Va au coin en bas à gauche du château.
'                        Persos(i).Aller_A .PositionX - Persos(i).Largeur, .PositionY + .Hauteur - Int(Rnd * (.Hauteur + Persos(i).Hauteur))
'                    End Select
'                    End With
'                ElseIf Persos(Persos(i).IA_Cible).DansUneMaison Or Persos(Persos(i).IA_Cible).DansSaMaison Then
'                    With Maisons(Persos(Persos(i).IA_Cible).IndiceMaison)
'                    Select Case Rnd
'                    Case Is < 0.25: 'Va au coin en haut à gauche de la maison.
'                        Persos(i).Aller_A .PositionX - Persos(i).Largeur + Int(Rnd * (.Largeur + Persos(i).Largeur)), .PositionY - Persos(i).Hauteur
'                    Case Is < 0.5: 'Va au coin en haut à droite de la maison.
'                        Persos(i).Aller_A .PositionX + .Largeur, .PositionY - Persos(i).Hauteur + Int(Rnd * (.Hauteur + Persos(i).Hauteur))
'                    Case Is < 0.75: 'Va au coin en bas à droite de la maison.
'                        Persos(i).Aller_A .PositionX + .Largeur - Int(Rnd * (.Largeur + Persos(i).Largeur)), .PositionY + .Hauteur
'                    Case Else  'Va au coin en bas à gauche de la maison.
'                        Persos(i).Aller_A .PositionX - Persos(i).Largeur, .PositionY + .Hauteur - Int(Rnd * (.Hauteur + Persos(i).Hauteur))
'                    End Select
'                    End With
'                Else
'                    With Persos(Persos(i).IA_Cible)
'                    Select Case Rnd
'                    Case Is < 0.25: 'Va au coin en haut à gauche du perso.
'                        Persos(i).Aller_A .PositionX - Persos(i).Largeur + Int(Rnd * (.Largeur + Persos(i).Largeur)), .PositionY - Persos(i).Hauteur
'                    Case Is < 0.5: 'Va au coin en haut à droite du perso.
'                        Persos(i).Aller_A .PositionX + .Largeur, .PositionY - Persos(i).Hauteur + Int(Rnd * (.Hauteur + Persos(i).Hauteur))
'                    Case Is < 0.75: 'Va au coin en bas à droite du perso.
'                        Persos(i).Aller_A .PositionX + .Largeur - Int(Rnd * (.Largeur + Persos(i).Largeur)), .PositionY + .Hauteur
'                    Case Else  'Va au coin en bas à gauche du perso.
'                        Persos(i).Aller_A .PositionX - Persos(i).Largeur, .PositionY + .Hauteur - Int(Rnd * (.Hauteur + Persos(i).Hauteur))
'                    End Select
'                    End With
'                End If
            End If
        Else
            Persos(i).IA_Ordre = 0
        End If
    End If
End Sub

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

Public Sub IA_ComportementBerserk(ByVal i As Long)
    'Comportement très agressif
    Dim j As Long
    With Persos(i)
    If Not .EnCombat Or .Longueur_Pas = 0 Then
        Select Case .IA_Berserk
        Case 0, 1: 'Cherche et détruit.
            IA_Trouver_Ennemi i
            If Not .EnCombat Then
            IA_Trouver_ChateauEnnemi i
                If Not .EnCombat Then
                IA_Trouver_MaisonEnnemi i
                    If Not .EnCombat And .Kamikaze Then
                        .Vivant = False
                    End If
                    If Not .EnDeplacement Then
                        IA_Deplacement_Aleatoire i
                    End If
                End If
            End If
        Case 2: 'Tourne autour du chef.
            If Not .EnCombat Then
                j = Trouver_Ennemi(i)
                If j >= 0 Then
                    If Distance_Persos(Persos(.NumeroChef), Persos(j)) < 384 Then
                        .Attaquer_Perso j
                    Else
                        IA_Patrouiller_Cercle i, Persos(.NumeroChef), 128
                    End If
                Else
                    IA_Patrouiller_Cercle i, Persos(.NumeroChef), 128
                End If
            End If
        Case 3: 'Cherche un fantôme à ressusciter.
            If IA_Trouver_Allie(i, False, True) Then
                If Not (.EnDeplacement) Then
                    If Parametres.Race_AutoResurrection(.Race) Then
                         Ressusciter_Perso .IndicePerso, .Effet_Attaque
                    End If
                    If .Kamikaze Then .Vivant = False
                End If
            Else 'If Not .EnDeplacement Then
                IA_Patrouiller_Cercle i, Persos(.NumeroChef), 128
            End If
        Case 4: 'Cherche des trésors à rammasser.
            If .NumeroChef > -1 Then
                If Persos(.NumeroChef).Vivant Then
                    'If .Definir_Argent > 0 Or .Definir_RessourcesTotales > 0 Or .Objet_Nombre_Total > 0 Then
                    If .Definir_Argent > 0 Then
                        'Va tout donner à son chef.
                        .ArgentSelectionne = .Definir_Argent
                        .Ressources_Prendre_Tout
                        .Objet_Prendre_Premier
                        .Aller_Donner .NumeroChef
                    End If
                End If
            End If
            If Not .EnDeplacement Then IA_Trouver_Tresor i
            If Not .EnDeplacement Then IA_Deplacement_Aleatoire i
        Case 5: 'Protège son bâtiment et les alentours de sa vie.
            'N'a plus de ressources pour attaquer.
            If .Inactif Then
                If .Definir_Vie < .Definir_MaxVie Then
                    If .IA_Manger Then
                        'Le personnage est blessé.
                        'Mange ce qu'il a.
                        .Recuperer_AutoVie Parametres
                        'N'a rien pour récuperer ses points de vie.
                        If Not .EnCombat And _
                           .Inactif Then
                                IA_Ramasser_Herbe i
                        End If
                    End If
                End If
            End If
            If .Inactif And _
               .Definir_ressources(.PerteRessourceIndice) < .PerteRessourceQuantite Then
                'S'il na pas d'argent, il va rammasser des munitions.
                If IA_Argent_Besoin(i) Then
                    IA_Chercher_Ressource i, .PerteRessourceIndice
                Else
                    If Trouver_Maison_Proche(Persos(i), i, .PerteRessourceIndice, 0, 0, X) Then
                        If X >= 0 Then
                            IA_Aller_Maison i, X
                        End If
                    Else
                        IA_Chercher_Ressource i, .PerteRessourceIndice
                    End If
                End If
            End If
            If .Inactif Then
                If Maisons(i).Visible Then
                    If Not .DansSaMaison Then
                        .Rentrer_Maison Maisons(i)
                    Else
                        'Prend une arme à distance.
                        If .Objet_Equipes_Actif(0) Then
                            If Parametres.Objet_BonusPorteeAttaque(.Objet_Equipes_Type(0)) = 0 Then
                                IA_Equiper i
                            End If
                        Else
                            IA_Equiper i
                        End If
                        'Envoie un projectile.
                        If .Attaque = .Definir_MaxAttaque And .ProjectileEnCours = 0 Then
                            IA_Lancer_projectile i
                            If .Inactif And Maisons(i).Definir_Vie < Maisons(i).MaxVie Then
                                .Rentrer_Maison Maisons(i)
                            End If
                        End If
                    End If
                Else
                    IA_Trouver_Ennemi i
                    If Not .EnCombat Then
                    IA_Trouver_ChateauEnnemi i
                        If Not .EnCombat Then
                        IA_Trouver_MaisonEnnemi i
                            If Not .EnCombat And .Kamikaze Then
                                .Vivant = False
                            End If
                            If Not .EnCombat Then
                                IA_Deplacement_Aleatoire i
                            End If
                        End If
                    End If
                End If
            End If
        Case 6: 'Fait des allés et retour au château en évitant les combats.
'            If .EnCombat And Not (.Action = 300 And Persos(.IndicePerso).Action = 300 And Persos(.IndicePerso).Numero = .Numero) Then
'                .Arreter
'            End If
            If .Inactif And _
               .Definir_ressources(.PerteRessourceIndice) < .PerteRessourceQuantite Then
                'Il va rammasser des munitions.
                IA_Chercher_Ressource i, .PerteRessourceIndice
                IA_Chercher_Ressource i, .PerteRessourceIndice
            End If
            If .Inactif Then
                If .Definir_Vie < .Definir_MaxVie Then
                    If .IA_Manger Then
                        'Le personnage est blessé.
                        'Mange ce qu'il a.
                        .Recuperer_AutoVie Parametres
                        'N'a rien pour récuperer ses points de vie.
                        If Not .EnCombat And _
                           .Inactif Then
                                IA_Ramasser_Herbe i
                        End If
                    End If
                ElseIf Not Maisons(i).Visible Then
                    'If Rnd < 0.005 Then IA_Deplacement_Aleatoire i
                    .Vivant = False
                ElseIf Maisons(i).Definir_Vie < Maisons(i).MaxVie Then
                    If Rnd < 0.01 Then .Rentrer_Maison Maisons(i)
                ElseIf .DansSonChateau Or .DansUneMaison Then
                    'If Rnd < 0.005 Then
                        j = Int(Distance_Positions(Maisons(.Numero).PositionX, Maisons(.Numero).PositionY, .PositionX, .PositionY) / 12) '16)
                        .Definir_Argent = .Definir_Argent + j
                        AffDegat.Ajouter_Degat .PositionX + .Largeur / 2, .PositionY, j, , 2
                        If Persos(.NumeroChef).Vivant Then
                            Persos(.NumeroChef).Definir_Argent = Persos(.NumeroChef).Definir_Argent + j
                            AffDegat.Ajouter_Degat Persos(.NumeroChef).PositionX + Persos(.NumeroChef).Largeur / 2, Persos(.NumeroChef).PositionY, j, , 2
                        End If
                        .Rentrer_Maison Maisons(i)
                    'End If
                ElseIf .DansSaMaison Then
                    If Rnd < 0.005 Then
                        If Chateaux(.NumeroFief).Visible And Rnd <= 1 / (Fiefs(.NumeroFief).Marches.Count + 1) Then
                            'Rentre au château.
                            .Rentrer_Chateau Chateaux(.NumeroFief)
                        Else
                            'Cherche un marché.
                            j = Int(Rnd * Fiefs(.NumeroFief).Marches.Count) 'On choisit un marché au hasard.
                            If Fiefs(.NumeroFief).Marches(j + 1) <> i Then 'Evite d'aller à son propre marché.
                                .Aller_Maison Maisons(Fiefs(.NumeroFief).Marches(j + 1)), Rnd < ChanceReparerMaison
                            End If
                        End If
'                        If .Definir_Argent > 1000 And Persos(.NumeroChef).Vivant Then
'                            Persos(.NumeroChef).Definir_Argent = Persos(.NumeroChef).Definir_Argent + .Definir_Argent
'                            'Affiche l'argent dépensé.
'                            AffDegat.Ajouter_Degat Persos(.NumeroChef).PositionX + Persos(.NumeroChef).Largeur / 2, Persos(.NumeroChef).PositionY, .Definir_Argent, True, 2
'                            .Definir_Argent = 0
'                        End If
                    End If
                    'Donne un tribut à son propriétaire.
                    'If Persos(.NumeroChef).DansUneMaison And Persos(.NumeroChef).IndiceMaison = .Numero Then
'                    If Chateaux(.NumeroFief).Visible Or Fiefs(.NumeroFief).Marches.Count > 1 Then
'                    End If
                    'End If
                ElseIf Not .DansSaMaison Then
                    .Rentrer_Maison Maisons(i)
                End If
            End If
        End Select
    End If
    End With
End Sub

Private Sub IA_Patrouiller_Cercle(ByVal NumeroIA As Long, ByVal Protege As ClsJeuPerso, ByVal Rayon As Long)
    Dim i As Long
    With Persos(NumeroIA)
    .Aller_A Cos(AnglePatrouille(NumeroIA)) * Rayon - .Largeur / 2 + Protege.PositionX + Protege.Largeur / 2, _
             Sin(AnglePatrouille(NumeroIA)) * Rayon - .Hauteur / 2 + Protege.PositionY + Protege.Hauteur / 2
    AnglePatrouille(NumeroIA) = AnglePatrouille(NumeroIA) + 6.28 * Rayon / .BonusDeplacement * 2
    If AnglePatrouille(NumeroIA) > 6.28 Then
        AnglePatrouille(NumeroIA) = AnglePatrouille(NumeroIA) - 6.28
    End If
    End With
End Sub

Private Sub IA_Lancer_projectile(ByVal NumeroIA As Long)
    'Cherche l'adversaire le plus proche et lui tir un projectile sans se déplacer.
    Dim i  As Long
    i = Trouver_Ennemi(NumeroIA)
    If i > -1 Then
        If Distance_Persos(Persos(i), Persos(NumeroIA)) <= Persos(NumeroIA).Definir_PorteeAttaque * 2 Then
            Persos(NumeroIA).Attaquer_Terrain Persos(i).PositionX + Persos(i).Largeur / 2, _
                                              Persos(i).PositionY + Persos(i).Hauteur / 2, _
                                              0
        End If
    End If
End Sub

Private Function IA_Attaquer_Perso(ByVal NumeroIA As Long, ByRef Indice As Long) As Boolean
    'Renvoie l'indice de persos le plus proche à attaquer.
    Dim i As Long, lb As Long
    lb = UBound(Persos())
    Indice = -1
    For i = 0 To NombreRechercheParTour
        If IndiceAdversaire(NumeroIA) > lb Then
            IndiceAdversaire(NumeroIA) = 0
            DistanceAdversaire(NumeroIA) = DistanceAdversaire(NumeroIA) + DistancePalierRecherchePerso
            If DistanceAdversaire(NumeroIA) > DistanceAdversaireMax Then
                DistanceAdversaire(NumeroIA) = DistanceAdversaireMin
            End If
            'Exit Function
        End If
        If DistanceAdversaire(NumeroIA) > Sqr((Persos(NumeroIA).PositionX - Persos(IndiceAdversaire(NumeroIA)).PositionX) ^ 2 + _
                                              (Persos(NumeroIA).PositionY - Persos(IndiceAdversaire(NumeroIA)).PositionY) ^ 2) And _
           Persos(IndiceAdversaire(NumeroIA)).NumeroEquipe <> Persos(NumeroIA).NumeroEquipe And _
           Persos(IndiceAdversaire(NumeroIA)).Attaquable Then
            IA_Attaquer_Perso = True
            Indice = IndiceAdversaire(NumeroIA)
            Exit For
        End If
        IndiceAdversaire(NumeroIA) = IndiceAdversaire(NumeroIA) + 1
    Next i
End Function

Private Function IA_Attaquer_Assaillant_Perso(ByVal NumeroIA As Long, ByVal NumeroCible As Long, ByRef IndiceAssaillant As Long) As Boolean
    'Renvoie l'indice d'un perso qui attaque une maison.
    Dim i As Long, lb As Long
    lb = UBound(Persos())
    IndiceAssaillant = -1
    For i = 0 To NombreRechercheParTour
        If IndiceAdversaire(NumeroIA) > lb Then
            IndiceAdversaire(NumeroIA) = 0
        End If
        If Persos(IndiceAdversaire(NumeroIA)).Action = 300 And _
           Persos(IndiceAdversaire(NumeroIA)).IndicePerso = NumeroCible Or _
           Persos(IndiceAdversaire(NumeroIA)).Action = 301 And _
           Persos(IndiceAdversaire(NumeroIA)).IndiceMaison = NumeroCible Or _
           Persos(IndiceAdversaire(NumeroIA)).Action = 302 And _
           Persos(NumeroCible).DansSonChateau And _
           Persos(IndiceAdversaire(NumeroIA)).IndiceChateau = Persos(NumeroCible).NumeroFief Then
            IA_Attaquer_Assaillant_Perso = True
            IndiceAssaillant = IndiceAdversaire(NumeroIA)
            Exit Function
        End If
        IndiceAdversaire(NumeroIA) = IndiceAdversaire(NumeroIA) + 1
    Next i
End Function

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

Private Function IA_Attaquer_Maison(ByVal NumeroIA As Long, ByRef Indice As Long) As Boolean
    'Renvoie l'indice de persos le plus proche à attaquer.
    Dim i As Long, lb As Long
    Dim pIAx As Long, pIAy As Long
    pIAx = Persos(NumeroIA).PositionX
    pIAy = Persos(NumeroIA).PositionY

    lb = UBound(Maisons())
    Indice = -1
    For i = 0 To NombreRechercheParTour
        If IndiceMaisonEnnemi(NumeroIA) > lb Then
            IndiceMaisonEnnemi(NumeroIA) = 0
            DistanceMaisonEnnemi(NumeroIA) = DistanceMaisonEnnemi(NumeroIA) + DistancePalierRechercheMaison
            If DistanceMaisonEnnemi(NumeroIA) > DistanceAdversaireMax Then
                DistanceMaisonEnnemi(NumeroIA) = DistanceAdversaireMin
            End If
            'Exit Function
        End If
        If DistanceMaisonEnnemi(NumeroIA) > Sqr((pIAx - Maisons(IndiceMaisonEnnemi(NumeroIA)).PositionX) ^ 2 + _
                                                 (pIAy - Maisons(IndiceMaisonEnnemi(NumeroIA)).PositionY) ^ 2) And _
           Persos(IndiceMaisonEnnemi(NumeroIA)).NumeroEquipe <> Persos(NumeroIA).NumeroEquipe And _
           Maisons(IndiceMaisonEnnemi(NumeroIA)).Visible Then
            IA_Attaquer_Maison = True
            Indice = IndiceMaisonEnnemi(NumeroIA)
            Exit Function
        End If
        IndiceMaisonEnnemi(NumeroIA) = IndiceMaisonEnnemi(NumeroIA) + 1
    Next i
End Function

Private Function IA_Attaquer_Chateau(ByVal NumeroIA As Long, ByRef Indice As Long) As Boolean
    'Renvoie l'indice de persos le plus proche à attaquer.
    Dim i As Long, pIAx As Long, pIAy As Long, uChato As Long
    pIAx = Persos(NumeroIA).PositionX
    pIAy = Persos(NumeroIA).PositionY
    uChato = UBound(Chateaux())
    Indice = -1
    For i = 0 To NombreRechercheParTour
        If IndiceChateauEnnemi(NumeroIA) > uChato Then
            IndiceChateauEnnemi(NumeroIA) = 0
            DistanceChateauEnnemi(NumeroIA) = DistanceChateauEnnemi(NumeroIA) + Parametres.Peuples_ChateauLargeur(0) * 2
            If DistanceChateauEnnemi(NumeroIA) > DistanceAdversaireMax Then
                DistanceChateauEnnemi(NumeroIA) = DistanceAdversaireMin
            End If
            'Exit Function
        End If
        If DistanceChateauEnnemi(NumeroIA) > Sqr((pIAx - Chateaux(IndiceChateauEnnemi(NumeroIA)).PositionX) ^ 2 + _
                                                 (pIAy - Chateaux(IndiceChateauEnnemi(NumeroIA)).PositionY) ^ 2) And _
           (Fiefs(IndiceChateauEnnemi(NumeroIA)).NumeroEquipe <> Persos(NumeroIA).NumeroEquipe And _
           Chateaux(IndiceChateauEnnemi(NumeroIA)).Visible) Then
            IA_Attaquer_Chateau = True
            Indice = IndiceChateauEnnemi(NumeroIA)
            i = NombreRechercheParTour
        End If
        IndiceChateauEnnemi(NumeroIA) = IndiceChateauEnnemi(NumeroIA) + 1
    Next i
End Function
Public Sub IA_Trouver_Ennemi(ByVal i As Long, _
                             Optional ByVal EnvoyerSoldats As Boolean, _
                             Optional ByVal Ctrl As Boolean, _
                             Optional ByVal Shift As Boolean, _
                             Optional ByVal Alt As Boolean)
    'Attaque l'ennemi le plus proche.
    Dim IndicePerso As Long
    IndicePerso = Trouver_Ennemi(i)
    
    If IndicePerso >= 0 Then
        If EnvoyerSoldats Then
            Ordonner_Chargez i, 0, IndicePerso, Ctrl, Shift, Alt
        Else
            Persos(i).Attaquer_Perso IndicePerso
        End If
    End If
End Sub

Public Function Trouver_Ennemi(ByVal i As Long, Optional ByVal X As Long = -1, Optional ByVal Y As Long = -1) As Long
    'Renvoie l'indice de l'ennemi le plus proche du personnage n°i.
    Dim j As Long
    Dim Distance As Long
    Dim TempDistance As Long
    Dim NumeroEquipe As Long
    If X = -1 And Y = -1 Then
        X = Persos(i).PositionX
        Y = Persos(i).PositionY
    End If
    NumeroEquipe = Persos(i).NumeroEquipe
    Trouver_Ennemi = -1
    Distance = Monde.Diagonale
    For j = 0 To UBound(Persos())
        With Persos(j)
        If .NumeroEquipe <> NumeroEquipe And _
           .Attaquable Then
            TempDistance = Distance_Positions(X, Y, Persos(j).PositionX, Persos(j).PositionY)
            If Distance > TempDistance Then
                Distance = TempDistance
                Trouver_Ennemi = j
            End If
        End If
        End With
    Next j
End Function


Private Function IA_Trouver_Allie(ByVal i As Long, _
                                  Optional ByVal Vivant As Boolean = True, _
                                  Optional ByVal Mort As Boolean) As Boolean
    'Cherche l'allié le plus proche.
    'Renvoie vrai, si on est prêt de cet allié.
    Dim IndicePerso As Long
    IndicePerso = Trouver_Allie(i, Vivant, Mort)
    
    With Persos(i)
    If IndicePerso >= 0 Then
        If Abs(.PositionX - Persos(IndicePerso).PositionX) < .Largeur And _
           Abs(.PositionY - Persos(IndicePerso).PositionY) < .Hauteur Then
            .EnDeplacement = False
        Else
            .Aller_A Persos(IndicePerso).PositionX, Persos(IndicePerso).PositionY
            .EnDeplacement = True
        End If
        IA_Trouver_Allie = True
        .IndicePerso = IndicePerso
    End If
    End With
End Function

Private Function Trouver_Allie(ByVal i As Long, _
                               Optional ByVal Vivant As Boolean = True, _
                               Optional ByVal Mort As Boolean) As Long
    'Renvoie l'indice l'allié le plus proche du personnage n°i.
    'Si "vivant" est vrai, on recherche les vivants uniquement.
    'Si "mort" est vrai, on recherche le fantôme uniquement.
    Dim j As Long
    Dim Distance As Long
    Dim TempDistance As Long
    Dim NumeroEquipe As Long 'Variables temporaires.
    NumeroEquipe = Persos(i).NumeroEquipe
    Trouver_Allie = -1
    Distance = Monde.Diagonale
    For j = 0 To UBound(Persos())
        With Persos(j)
        If .NumeroEquipe = NumeroEquipe Then
            If Not .Invocation Then
                If Not Vivant Or .Vivant Then
                    If Not (Mort And .Vivant) Then
                        TempDistance = Distance_Persos(Persos(i), Persos(j))
                        If Distance > TempDistance Then
                            Distance = TempDistance
                            Trouver_Allie = j
                        End If
                    End If
                End If
            End If
        End If
        End With
    Next j
End Function

Public Sub IA_Trouver_MaisonEnnemi(ByVal i As Long, _
                                   Optional ByVal EnvoyerSoldats As Boolean, _
                                   Optional ByVal Ctrl As Boolean, _
                                   Optional ByVal Shift As Boolean, _
                                   Optional ByVal Alt As Boolean)
    'Attaque la maison la plus proche.
    Dim IndiceMaison As Long
    IndiceMaison = Trouver_MaisonEnnemi(i)
    
    If IndiceMaison >= 0 Then
        If EnvoyerSoldats Then
            Ordonner_Chargez i, 1, IndiceMaison, Ctrl, Shift, Alt
        Else
            Persos(i).Attaquer_Maison IndiceMaison
        End If
    End If
End Sub

Public Function Trouver_MaisonEnnemi(ByVal i As Long) As Long
    'Renvoie l'indice de la maison ennemi la plus proche du personnage n°i.
    Dim j As Long
    Dim Distance As Long
    Trouver_MaisonEnnemi = -1
    Distance = Monde.Largeur * AffTerrain.Largeur + Monde.Hauteur * AffTerrain.Hauteur
    For j = 0 To UBound(Maisons())
        If Persos(j).NumeroEquipe <> Persos(i).NumeroEquipe And _
           Maisons(j).Visible And _
           Distance > Sqr((Persos(i).PositionX - Maisons(j).PositionX) ^ 2 + _
                          (Persos(i).PositionY - Maisons(j).PositionY) ^ 2) Then
            Distance = Sqr((Persos(i).PositionX - Maisons(j).PositionX) ^ 2 + _
                           (Persos(i).PositionY - Maisons(j).PositionY) ^ 2)
            Trouver_MaisonEnnemi = j
            'j = UBound(Maisons())
        End If
    Next j
End Function

Public Sub IA_Trouver_ChateauEnnemi(ByVal i As Long, _
                                    Optional ByVal EnvoyerSoldats As Boolean, _
                                    Optional ByVal Ctrl As Boolean, _
                                    Optional ByVal Shift As Boolean, _
                                    Optional ByVal Alt As Boolean)
    'Attaque l'ennemi le plus proche.
    Dim IndiceChateau As Long
    IndiceChateau = Trouver_ChateauEnnemi(i)
    
    If IndiceChateau >= 0 Then
        If EnvoyerSoldats Then
            Ordonner_Chargez i, 2, IndiceChateau, Ctrl, Shift, Alt
        Else
            Persos(i).Attaquer_Chateau IndiceChateau
        End If
    End If
End Sub

Public Function Trouver_ChateauEnnemi(ByVal i As Long) As Long
    'Renvoie l'indice du château ennemi le plus proche du personnage n°i.
    Dim j As Long
    Dim Distance As Long
    Trouver_ChateauEnnemi = -1
    Distance = Monde.Largeur * AffTerrain.Largeur + Monde.Hauteur * AffTerrain.Hauteur
    For j = 0 To UBound(Chateaux())
        If Persos(i).NumeroEquipe <> Fiefs(j).NumeroEquipe And _
           Chateaux(j).Visible And _
           Distance > Sqr((Persos(i).PositionX - Chateaux(j).PositionX) ^ 2 + _
                          (Persos(i).PositionY - Chateaux(j).PositionY) ^ 2) Then
            Distance = Sqr((Persos(i).PositionX - Chateaux(j).PositionX) ^ 2 + _
                           (Persos(i).PositionY - Chateaux(j).PositionY) ^ 2)
            Trouver_ChateauEnnemi = j
            'j = UBound(Chateaux())
        End If
    Next j
End Function

Public Sub IA_Poursuivre_Fuyard(ByVal i As Long)
    'Le fuyard qui était dans la maison ressort.
    With Persos(i)
    If (.Action = 301 Or .Action = 302) And _
        Persos(.IndicePerso).Attaquable And _
       .NumeroEquipe <> Persos(.IndicePerso).NumeroEquipe Then
             .Attaquer_Perso .IndicePerso
             .Definir_Commentaires = Commentaires.Message(16, Persos(i))
    End If
    End With
End Sub

Public Sub IA_Recuperation_Carac(ByVal i As Long)
    With Persos(i)
    If .IA_Manger Then
    'Le personnage est blessé.
    'Mange ce qu'il a.
    .Recuperer_AutoVie Parametres
    .Recuperer_AutoEnergie Parametres
    .Recuperer_AutoMagie Parametres
    .Recuperer_AutoMoral Parametres
    
    'Achete des matières prémières dans une maison.
    If .DansUneMaison Then
        If Maisons(.IndiceMaison).Magasin Then
            If Parametres.Ressources_Mangeable(Parametres.Batiment_RessourceFabrique(Maisons(.IndiceMaison).TypeBatiment)) Then
                Maisons(.IndiceMaison).Acheter 5, Persos(i), Parametres
                Exit Sub
            End If
        End If
    ElseIf Not .DansUneMaison Then
        'If Not .EnCombat Then
        If .Inactif Then
            'N'a rien pour récuperer ses points de vie.
            If .Definir_Vie < .Definir_MaxVie Then
                'S'il na pas d'argent, il va chercher des herbes médicinales.
                If IA_Argent_Besoin(i) Then
                    IA_Ramasser_Herbe i
                    Exit Sub
                ElseIf Trouver_Maison_Proche(Persos(i), i, 0, 1, 0, X) Then
                    If X >= 0 Then
                        IA_Aller_Maison i, X
                        Exit Sub
                    Else
                        IA_Ramasser_Herbe i
                        Exit Sub
                    End If
                End If
            End If
            
            'Le personnage a faim.
            If .Definir_Energie < .Definir_MaxEnergie * 0.5 Then
                'S'il na pas d'argent, il va chercher les fruits sur un arbre
                If IA_Argent_Besoin(i) Then
                    IA_Cueillir_Fruits i
                    Exit Sub
                'Sinon, va acheter à manger.
                ElseIf Trouver_Maison_Proche(Persos(i), i, 0, 2, 0, X) Then
                    If X >= 0 Then
                        IA_Aller_Maison i, X
                        Exit Sub
                    End If
                Else
                    IA_Cueillir_Fruits i
                    Exit Sub
                End If
            End If
            
            'Le personnage manque de magie.
            'Va en acheter.
            If .Definir_Magie < .Definir_MaxMagie * 0.5 Then
                If Not IA_Argent_Besoin(i) Then
                    If Trouver_Maison_Proche(Persos(i), i, 0, 3, 0, X) Then
                        If X >= 0 Then
                            IA_Aller_Maison i, X
                            Exit Sub
                        End If
                    End If
                End If
            End If
            
            'Le personnage est démoralisé.
            'Va en acheter.
            If .Definir_Moral < .Definir_MaxMoral * 0.5 Then
                If Not IA_Argent_Besoin(i) Then
                    If Trouver_Maison_Proche(Persos(i), i, 0, 4, 0, X) Then
                        If X >= 0 Then
                            IA_Aller_Maison i, X
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        
        'N'a plus de ressources pour attaquer.
        'Va en acheter.
        If .PerteRessourceIndice >= 0 Then
            If .Definir_ressources(.PerteRessourceIndice) < .PerteRessourceQuantite Then
                'S'il na pas d'argent, il va rammasser des munitions.
                If IA_Argent_Besoin(i) Then
                    IA_Chercher_Ressource i, .PerteRessourceIndice
                    Exit Sub
                Else
                    If Trouver_Maison_Proche(Persos(i), i, .PerteRessourceIndice, 0, 0, X) Then
                        If X >= 0 Then
                            IA_Aller_Maison i, X
                            Exit Sub
                        End If
                    Else
                        IA_Chercher_Ressource i, .PerteRessourceIndice
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    End If
    
    'Va chercher un trésor.
    If .Inactif And .IA_Equiper > 0 Then
        If Trouver_Tresor_Proche(Persos(i), i, X) Then
            If X > 0 Then
                IA_Aller_Tresor i, X
                Exit Sub
            End If
        End If
    End If
    End With
End Sub

Public Sub IA_Recuperation_CaracSimple(ByVal i As Long)
    With Persos(i)
    If .IA_Manger Then
    'Le personnage est blessé.
    'Mange ce qu'il a.
    .Recuperer_AutoVie Parametres
    .Recuperer_AutoEnergie Parametres
    .Recuperer_AutoMagie Parametres
    .Recuperer_AutoMoral Parametres
    'If Not .EnCombat Then
    If .Inactif Then
        'N'a rien pour récuperer ses points de vie.
        If .Definir_Vie < .Definir_MaxVie Then
            'S'il na pas d'argent, il va chercher des herbes médicinales.
            IA_Ramasser_Herbe i
            Exit Sub
        End If
        
        'Le personnage a faim.
        If .Definir_Energie < .Definir_MaxEnergie * 0.5 Then
            'S'il na pas d'argent, il va chercher les fruits sur un arbre
            IA_Cueillir_Fruits i
            Exit Sub
        End If
    End If
    
    'N'a plus de ressources pour attaquer.
    'Va en acheter.
    If .PerteRessourceIndice >= 0 Then
        If .Definir_ressources(.PerteRessourceIndice) < .PerteRessourceQuantite Then
            'S'il na pas d'argent, il va rammasser des munitions.
            IA_Chercher_Ressource i, .PerteRessourceIndice
            Exit Sub
        End If
    End If
    End If
    End With
End Sub

Public Sub IA_Depenser_Argent(ByVal i As Long)
    'L'IA dépense son surplus d'argent.
    With Persos(i)
    If .Inactif And _
       Not .DansUneMaison And _
       .Definir_Vie = .Definir_MaxVie And _
       .Definir_Energie >= .Definir_MaxEnergie * 0.5 And _
       .Definir_Moral >= .Definir_MaxMoral * 0.5 And _
       IA_Argent_Riche(i) And _
       Not .Objet_Inventaire_Actif(.Nombre_Objets_Inventaire - 1) And _
       Trouver_Maison_Proche(Persos(i), i, 0, 5, 0, X) Then
        If X >= 0 Then
            IA_Aller_Maison i, X
        End If
    End If
    End With
End Sub

Private Sub IA_Ramasser_Herbe(ByVal i As Long)
    If IndiceHerbe >= 0 Then
        IA_Chercher_Ressource i, IndiceHerbe
    End If
End Sub
Private Sub IA_Cueillir_Fruits(ByVal i As Long)
    If IndiceFruit >= 0 Then
        IA_Chercher_Ressource i, IndiceFruit
    End If
End Sub
Private Sub IA_Couper_Arbre(ByVal i As Long)
    If IndiceBois >= 0 Then
        IA_Chercher_Ressource i, IndiceBois
    End If
End Sub
Private Sub IA_ChercherPierre(ByVal i As Long)
    If IndicePierre >= 0 Then
        IA_Chercher_Ressource i, IndicePierre
    End If
End Sub

Public Sub IA_ChercherMatieresPremiere(ByVal i As Long)
    With Persos(i)
    If (.Inactif Or .Action = 202) And _
                       Not .DansUneMaison And _
                       Maisons(i).Construit And _
                       .Definir_ressources(Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)) < MatierePremiereMin And _
                       Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment) >= 0 Then
        'Va récolter ses matières premières.
        If Parametres.Ressources_Recoltable(Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)) Then
            IA_Chercher_Ressource i, Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)
        Else
            'Va en acheter.
            If .Definir_Argent > 1 Then
                If Trouver_Maison_Proche(Persos(i), i, Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment), 0, 0, X) Then
                    If X >= 0 Then
                        IA_Aller_Maison i, X
                    End If
                'ElseIf Not .DansSonChateau And _
                       Not .DansSaMaison And _
                       Not .DansUneMaison And _
                       Not .Action = 201 And _
                       Fiefs(.NumeroFief).Definir_Ressources(Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)) > 0 And _
                       Chateaux(.NumeroFief).Visible Then
                ElseIf QuitterBatiment(i) And _
                       Not .Action = 201 And _
                       Fiefs(.NumeroFief).Definir_ressources(Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)) > MatierePremiereMin And _
                       Chateaux(.NumeroFief).Visible Or Fiefs(.NumeroFief).Marches.Count > 0 Then
                    IA_Rentrer_Chateau i, True
                End If
            End If
        End If
    End If
    End With
End Sub

Private Sub IA_Chercher_Ressource_Pour_Construire(ByVal i As Long)
    Dim j As Long
    If BatimentAConstruire(i) >= 0 Then
        For j = 0 To Parametres.NombreRessources - 1
            If Parametres.Ressources_Recoltable(j) And _
               Persos(i).Definir_ressources(j) < Parametres.Batiment_PrixRessource(BatimentAConstruire(i), j) Then
                IA_Chercher_Ressource i, j
                j = Parametres.NombreRessources
            End If
        Next j
    End If
End Sub
Public Sub IA_Chercher_Ressource(ByVal i As Long, ByVal TypeRessource As Long)
    Dim j As Long
    'If Trouver_Ressource_Proche(Persos(i), i, TypeRessource, j) Then
    '    Persos(i).Aller_Ressource Ressources(j), j, AffRessources.Largeur, AffRessources.Hauteur
    'End If
    IA_Trouver_Ressource i, Parametres.Ressources_TypeTerrainRessource(TypeRessource)
End Sub

Public Sub IA_Trouver_Ressource(ByVal i As Long, ByVal TypeRessource As Long)
    'Même fonction que IA_Chercher_Ressource, sauf que la fonction continue jusqu'à temps de trouver la ressource.
    Dim j As Long
    Dim Distance As Long
    Dim IndiceRessource As Long
    
    Distance = Monde.Largeur * AffTerrain.Largeur + Monde.Hauteur * AffTerrain.Hauteur
    For j = 0 To UBound(Ressources())
        'Debug.Print Ressources(j).TypeRessource
        If Ressources(j).Apparence = TypeRessource And _
           Distance > Sqr((Persos(i).PositionX - Ressources(j).PositionX) ^ 2 + _
                          (Persos(i).PositionY - Ressources(j).PositionY) ^ 2) Then
           'Distance > Abs(Persos(i).PositionX - Ressources(j).PositionX) + _
                      Abs(Persos(i).PositionY - Ressources(j).PositionY) Then
            Distance = Sqr((Persos(i).PositionX - Ressources(j).PositionX) ^ 2 + _
                           (Persos(i).PositionY - Ressources(j).PositionY) ^ 2)
            'Distance = Abs(Persos(i).PositionX - Ressources(j).PositionX) + _
                       Abs(Persos(i).PositionY - Ressources(j).PositionY)
            IndiceRessource = j
        End If
    Next j
    'Debug.Print TypeRessource & " : " & IndiceRessource
    Persos(i).Aller_Ressource Ressources(IndiceRessource), IndiceRessource, AffRessources.Largeur, AffRessources.Hauteur
End Sub

Private Sub IA_Acheter_Ressources_Pour_Manger(ByVal i As Long)
    'Achète des ressources au château pour manger.
    Dim j As Integer
    If Not IA_Argent_Besoin(i) Then
        With Persos(i)
        If .Definir_Moral < .Definir_MaxMoral Then
            For j = Parametres.NombreRessources - 1 To 0 Step -1
                If Parametres.Ressources_GainMoral(j) > 0 Then
                    If Not Parametres.Ressources_Recoltable(j) Then
                        Fiefs(.NumeroFief).Acheter j, 0, Persos(i), Parametres, True
                        Exit Sub
                    End If
                End If
            Next j
        End If
        If .Definir_Vie < .Definir_MaxVie Then
            For j = Parametres.NombreRessources - 1 To 0 Step -1
                If Parametres.Ressources_GainVie(j) > 0 Then
                    If Not Parametres.Ressources_Recoltable(j) Then
                        Fiefs(.NumeroFief).Acheter j, 0, Persos(i), Parametres, True
                        Exit Sub
                    End If
                End If
            Next j
        End If
        If .Definir_Energie < .Definir_MaxEnergie Then
            For j = Parametres.NombreRessources - 1 To 0 Step -1
                If Parametres.Ressources_GainEnergie(j) > 0 Then
                    If Not Parametres.Ressources_Recoltable(j) Then
                        Fiefs(.NumeroFief).Acheter j, 0, Persos(i), Parametres, True
                        Exit Sub
                    End If
                End If
            Next j
        End If
        If .Definir_Magie < .Definir_MaxMagie Then
            For j = Parametres.NombreRessources - 1 To 0 Step -1
                If Parametres.Ressources_GainMagie(j) > 0 Then
                    If Not Parametres.Ressources_Recoltable(j) Then
                        Fiefs(.NumeroFief).Acheter j, 0, Persos(i), Parametres, True
                        Exit Sub
                    End If
                End If
            Next j
        End If
        End With
    End If
End Sub

Private Sub IA_Acheter_Ressources_Pour_Travailler(ByVal i As Long)
    Dim j As Integer
    Dim k As Integer
    Dim Mtb As Long
    Mtb = Maisons(i).TypeBatiment
    'If i = 0 Then Debug.Print Maisons(i).Visible & " : " & Parametres.Batiment_Fabrique(Mtb) & " : " & Parametres.Batiment_TypeMatierePremiere(Mtb)
    If Maisons(i).Visible Then
        If Parametres.Batiment_Fabrique(Mtb) Then
            'Achête les ressources pour tous les objets qu'il fabrique.
            For j = 0 To Parametres.NombreRessources - 1
                If Not (Parametres.Ressources_Recoltable(j)) Then
                    For k = 0 To Parametres.Batiment_NombreObjetsFabriques(Mtb) - 1
                        If Parametres.Objet_PrixRessources(Parametres.Batiment_ObjetFabrique(Mtb, k), j) > 0 Then
                            Fiefs(Persos(i).NumeroFief).Acheter j, 0, Persos(i), Parametres, True
                            k = Parametres.Batiment_NombreObjetsFabriques(Mtb)
                        End If
                    Next k
                End If
            Next j
        ElseIf Parametres.Batiment_TypeMatierePremiere(Mtb) >= 0 Then
            If Not (Parametres.Ressources_Recoltable(Parametres.Batiment_TypeMatierePremiere(Mtb))) Then
                Fiefs(Persos(i).NumeroFief).Acheter Parametres.Batiment_TypeMatierePremiere(Mtb), 0, Persos(i), Parametres, True
            End If
        End If
    End If
End Sub
Private Sub IA_Acheter_Ressources_Pour_Construire(ByVal i As Long)
    Dim j As Long
    If BatimentAConstruire(i) >= 0 Then
        For j = 0 To Parametres.NombreRessources - 1
            'If Parametres.Ressources_NecessairePourConstruire(j) And _
               j <> Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment) Then
               'Fiefs(Persos(i).NumeroFief).Acheter j, QuantiteMaxTransfert * RND, Persos(i), Parametres
            If Parametres.Batiment_PrixRessource(BatimentAConstruire(i), j) > Persos(i).Definir_ressources(j) And _
               j <> Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment) And _
               Not Parametres.Ressources_Recoltable(j) Then
                Fiefs(Persos(i).NumeroFief).Acheter j, QuantiteMaxTransfert * Rnd, Persos(i), Parametres
            End If
        Next j
    End If
End Sub

Public Sub IA_Equiper(ByVal i As Long)
    Dim j As Long
    With Persos(i)
    Select Case .IA_Equiper
    Case 1: 'Equipement selon la valeur.
        For j = 0 To IIf(.EnCombat, VitesseEquipementCombat, VitesseEquipement)
            If .Objet_Inventaire_Actif(IndiceObjet) Then
                If Parametres.Objet_Utilisable(.Objet_Inventaire_Type(IndiceObjet)) Then
                    'Objets utilisables.
                    If Parametres.Objet_UtilisationUnique(.Objet_Inventaire_Type(IndiceObjet)) Then
                        If .EnCombat Then
                            'N'utilise pas les invocations hors-combat.
                            .Objet_Equiper IndiceObjet, Parametres, Commentaires
                        End If
                    Else
                        .Objet_Equiper IndiceObjet, Parametres, Commentaires
                    End If
                Else
                    Select Case Parametres.Objet_Categorie(.Objet_Inventaire_Type(IndiceObjet))
                    Case 0: 'Arme à une main.
                        If Not .Objet_Equipes_Actif(0) Or Not .Objet_Equipes_Actif(1) Then
                            .Objet_Equiper IndiceObjet, Parametres, Commentaires
                        Else
                            If Parametres.Objet_PrixVente(.Objet_Inventaire_Type(IndiceObjet)) _
                               > Parametres.Objet_PrixVente(.Objet_Equipes_Type(0)) Then
                                .Objet_Equiper IndiceObjet, Parametres, Commentaires
                            End If
                        End If
                    Case 1: 'Arme à 2 mains.
                        If Not .Objet_Equipes_Actif(0) Then
                            .Objet_Equiper IndiceObjet, Parametres, Commentaires
                        ElseIf .DansUneMaison Or .DansSaMaison And Parametres.Batiment_Tir(Maisons(.IndiceMaison).TypeBatiment) And Parametres.Objet_BonusPorteeAttaque(.Objet_Inventaire_Type(IndiceObjet)) > Parametres.Objet_BonusPorteeAttaque(.Objet_Equipes_Type(0)) Then
                            'Dans une tourelle de défense, on prend une arme à distance de préférence.
                            .Objet_Equiper IndiceObjet, Parametres, Commentaires
                        Else
                            If Parametres.Objet_PrixVente(.Objet_Inventaire_Type(IndiceObjet)) _
                               > Parametres.Objet_PrixVente(.Objet_Equipes_Type(0)) Then
                                .Objet_Equiper IndiceObjet, Parametres, Commentaires
                            End If
                        End If
                    Case Else 'Autre équipement.
                        If Not .Objet_Equipes_Actif(Parametres.Objet_Categorie(.Objet_Inventaire_Type(IndiceObjet))) Then
                            .Objet_Equiper IndiceObjet, Parametres, Commentaires
                        Else
                            If Parametres.Objet_PrixVente(.Objet_Inventaire_Type(IndiceObjet)) _
                               > Parametres.Objet_PrixVente(.Objet_Equipes_Type(Parametres.Objet_Categorie(.Objet_Inventaire_Type(IndiceObjet)))) Then
                                .Objet_Equiper IndiceObjet, Parametres, Commentaires
                            End If
                        End If
                    End Select
                End If
            End If
            
            IndiceObjet = IndiceObjet + 1
            If IndiceObjet > Persos(i).Nombre_Objets_Inventaire - 1 Then
                IndiceObjet = 0
            End If
        Next j
    Case 2: 'Equipement selon la compétence. (Puis selon le prix)
        For j = 0 To IIf(.EnCombat, VitesseEquipementCombat, VitesseEquipement)
            If .Objet_Inventaire_Actif(IndiceObjet) Then
                If Parametres.Objet_Utilisable(.Objet_Inventaire_Type(IndiceObjet)) Then
                    'Objets utilisables.
                    If Parametres.Objet_UtilisationUnique(.Objet_Inventaire_Type(IndiceObjet)) Then
                        If .EnCombat Then
                            'N'utilise pas les invocations hors-combat.
                            .Objet_Equiper IndiceObjet, Parametres, Commentaires
                        End If
                    Else
                        .Objet_Equiper IndiceObjet, Parametres, Commentaires
                    End If
                Else
                    Select Case Parametres.Objet_Categorie(.Objet_Inventaire_Type(IndiceObjet))
                    Case 0: 'Arme à une main.
                        If Not .Objet_Equipes_Actif(0) Or Not .Objet_Equipes_Actif(1) Then
                            .Objet_Equiper IndiceObjet, Parametres, Commentaires
                        Else
                            If .Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(.Objet_Inventaire_Type(IndiceObjet))) > _
                               .Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(.Objet_Equipes_Type(0))) Then
                                .Objet_Equiper IndiceObjet, Parametres, Commentaires
                            ElseIf .Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(.Objet_Inventaire_Type(IndiceObjet))) = _
                                 .Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(.Objet_Equipes_Type(0))) Then
                                 If Parametres.Objet_PrixVente(.Objet_Inventaire_Type(IndiceObjet)) > _
                                     Parametres.Objet_PrixVente(.Objet_Equipes_Type(0)) Then
                                      .Objet_Equiper IndiceObjet, Parametres, Commentaires
                                 End If
                            End If
                        End If
                    Case 1: 'Arme à 2 mains.
                        If Not .Objet_Equipes_Actif(0) Then
                            .Objet_Equiper IndiceObjet, Parametres, Commentaires
                        ElseIf .DansUneMaison Or .DansSaMaison And Parametres.Batiment_Tir(Maisons(.IndiceMaison).TypeBatiment) And Parametres.Objet_BonusPorteeAttaque(.Objet_Inventaire_Type(IndiceObjet)) > Parametres.Objet_BonusPorteeAttaque(.Objet_Equipes_Type(0)) Then
                            'Dans une tourelle de défense, on prend une arme à distance de préférence.
                            .Objet_Equiper IndiceObjet, Parametres, Commentaires
                        Else
                            If .Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(.Objet_Inventaire_Type(IndiceObjet))) > _
                               .Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(.Objet_Equipes_Type(0))) Then
                                .Objet_Equiper IndiceObjet, Parametres, Commentaires
                            ElseIf .Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(.Objet_Inventaire_Type(IndiceObjet))) = _
                                   .Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(.Objet_Equipes_Type(0))) Then
                                If Parametres.Objet_PrixVente(.Objet_Inventaire_Type(IndiceObjet)) > _
                                   Parametres.Objet_PrixVente(.Objet_Equipes_Type(0)) Then
                                    .Objet_Equiper IndiceObjet, Parametres, Commentaires
                                End If
                            End If
                        End If
                    Case Else 'Autre équipement.
                        If Not .Objet_Equipes_Actif(Parametres.Objet_Categorie(.Objet_Inventaire_Type(IndiceObjet))) Then
                            .Objet_Equiper IndiceObjet, Parametres, Commentaires
                        Else
                            If .Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(.Objet_Inventaire_Type(IndiceObjet))) > _
                               .Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(.Objet_Equipes_Type(Parametres.Objet_Categorie(.Objet_Inventaire_Type(IndiceObjet))))) Then
                                .Objet_Equiper IndiceObjet, Parametres, Commentaires
                            ElseIf .Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(.Objet_Inventaire_Type(IndiceObjet))) = _
                                   .Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(.Objet_Equipes_Type(Parametres.Objet_Categorie(.Objet_Inventaire_Type(IndiceObjet))))) Then
                                If Parametres.Objet_PrixVente(.Objet_Inventaire_Type(IndiceObjet)) > _
                                   Parametres.Objet_PrixVente(.Objet_Equipes_Type(Parametres.Objet_Categorie(.Objet_Inventaire_Type(IndiceObjet)))) Then
                                    .Objet_Equiper IndiceObjet, Parametres, Commentaires
                                End If
                            End If
                        End If
                    End Select
                End If
            End If
            
            IndiceObjet = IndiceObjet + 1
            If IndiceObjet > Persos(i).Nombre_Objets_Inventaire - 1 Then
                IndiceObjet = 0
            End If
        Next j
    Case 3: 'Equipement intelligent
        'Travail (Ressources)
        'Travail (Objets)
        'Travail (Autres)
        'Combat (corps à corps)
        'Combat (distance)
    End Select
    End With
End Sub

Public Sub IA_Mort_Equiper(ByVal i As Long)
    Dim j As Long
    With Persos(i)
    If .IA_Equiper > 0 Then
        For j = 0 To VitesseEquipement
            If .Objet_Inventaire_Actif(IndiceObjet) Then
                If Parametres.Objet_UtilisationMort(.Objet_Inventaire_Type(IndiceObjet)) Then
                    'Objets utilisables quand il est mort.
                    .Objet_Equiper IndiceObjet, Parametres, Commentaires
                End If
            End If
            IndiceObjet = IndiceObjet + 1
            If IndiceObjet > Persos(i).Nombre_Objets_Inventaire - 1 Then
                IndiceObjet = 0
            End If
        Next j
    End If
    End With
End Sub

Private Function Nombre_Total_Batiments(ByVal NumeroFief As Long) As Integer
    Dim i As Long
    With Persos(i)
    For i = 0 To UBound(Maisons())
        If .NumeroFief = NumeroFief Then
            Nombre_Total_Batiments = Nombre_Total_Batiments + 1
        End If
    Next i
    End With
End Function

Private Function Nombre_Batiments(ByVal TypeBatiment As Long, ByVal NumeroFief As Long) As Integer
    Dim i As Long
    If TypeBatiment >= 0 Then
        For i = 0 To UBound(Maisons())
            'If Maisons(i).TypeBatiment = TypeBatiment And _
            '   Persos(i).NumeroFief = NumeroFief Then
            '        Nombre_Batiments = Nombre_Batiments + 1
            'End If
            Nombre_Batiments = Nombre_Batiments + _
                               Abs(Maisons(i).TypeBatiment = TypeBatiment And _
                               Persos(i).NumeroFief = NumeroFief)
        Next i
    End If
End Function

Public Function Choisir_Batiment(ByVal NumeroPerso As Long, _
                                 ByVal NumeroFief As Long, _
                                 Optional ByVal Niveau As Long = -1) As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Choisir_Batiment = -1
    
    'Cherche une maison de niveau supérieur.
    If Niveau = -1 Then
        Niveau = 0
        For i = Int(Rnd * BatimentNiveauMax) + 1 To Parametres.Batiment_Niveau(Maisons(NumeroPerso).TypeBatiment) + 1 Step -1
            If Choisir_Batiment = -1 Then
                Choisir_Batiment = Choisir_Batiment(NumeroPerso, NumeroFief, i)
                If Persos(NumeroPerso).PeutConstruire_Maison(Choisir_Batiment, Parametres) Then
                    i = 0
                Else
                    Choisir_Batiment = -1
                End If
            End If
        Next i
    End If
    
    'Cherche s'il dispose du monopole sur un batiment de niveau inférieur que le sien.
    If Nombre_Batiments(Maisons(NumeroPerso).TypeBatiment, NumeroFief) > 1 Then 'Sauf s'il a déjà le monopole.
        For i = 0 To Parametres.NombreTypeBatiments - 1
            If Nombre_Batiments(i, NumeroFief) = 0 And _
               Parametres.Batiment_Niveau(i) < Parametres.Batiment_Niveau(Maisons(NumeroPerso).TypeBatiment) Then
            'If Nombre_Batiments(i, NumeroFief) = 0 Then
                Choisir_Batiment = i
                i = Parametres.NombreTypeBatiments
            End If
        Next i
    End If
    If Choisir_Batiment = -1 Then
        For i = 0 To Parametres.NombreTypeBatiments - 1
            If Parametres.Batiment_Niveau(i) = Niveau Then
                k = k + 1
            End If
        Next i
        k = Int(Rnd * k)
        For i = 0 To Parametres.NombreTypeBatiments - 1
            If Parametres.Batiment_Niveau(i) = Niveau Then
                If j = k Then
                    Choisir_Batiment = i
                    i = Parametres.NombreTypeBatiments
                Else
                    j = j + 1
                End If
            End If
        Next i
    ElseIf Parametres.Batiment_NoEpoque(Choisir_Batiment) - 1 > Fiefs(NumeroFief).Epoque And _
           Persos(NumeroPerso).PeutConstruire_Maison(Choisir_Batiment, Parametres) And _
           Not Maisons(NumeroPerso).Visible Then
        'Si l'IA n'a pas de maison et ne peut pas la fabriquer tout de suite, il en choisit une de niveau 0.
        Choisir_Batiment = Choisir_Batiment(NumeroPerso, NumeroFief, 0)
    End If
End Function

Private Function Trouver_Ressource_Proche(ByRef Personnage As ClsJeuPerso, _
                                          ByVal NumeroIA As Long, _
                                          ByVal TypeRessource As Long, _
                                          ByRef Indice As Long) As Boolean
    'X = CLng(RND * UBound(Ressources()))
    Dim i As Long
    'Dim j As Long
    'For j = 0 To Monde.Longueur_Diagonale Step 64
    DistanceRessource(NumeroIA) = DistanceRessource(NumeroIA) + AffRessources.Largeur * Rnd * 10
    For i = 0 To UBound(Ressources())
        'If IA_IndiceRessource(NumeroIA) > UBound(Ressources()) Then
        '    DistanceRessource(NumeroIA) = DistanceRessource(NumeroIA) + AffRessources.Largeur * 2
        '    IA_IndiceRessource(NumeroIA) = 0
        'End If
        If Ressources(i).TypeRessource = TypeRessource And _
            DistanceRessource(NumeroIA) >= Sqr((Personnage.PositionX - Ressources(i).PositionX) ^ 2 + _
                                               (Personnage.PositionY - Ressources(i).PositionY) ^ 2) Then
            Indice = i
            DistanceRessource(NumeroIA) = 0
            'IA_IndiceRessource(NumeroIA) = 0
            Trouver_Ressource_Proche = True
            'Exit Function
            i = UBound(Ressources())
            'j = Monde.Longueur_Diagonale
        'Else
        '    IA_IndiceRessource(NumeroIA) = IA_IndiceRessource(NumeroIA) + 1
        End If
    'Debug.Print DistanceRessource & " / " & IA_IndiceRessource(NumeroIA)
    Next i
    'Next j
End Function

Private Function Trouver_Tresor_Proche(ByRef Perso As ClsJeuPerso, _
                                       ByVal NumeroIA As Long, _
                                       ByRef Indice As Long) As Boolean
    'Recherche un trésor.
    Dim i As Long
    DistanceTresor(NumeroIA) = DistanceTresor(NumeroIA) + AffTresor.Largeur * Rnd * 10
    If DistanceTresor(NumeroIA) > DistanceTresorMax Then
        DistanceTresor(NumeroIA) = 0
    Else
        For i = 1 To Tresors.Count
            If DistanceTresor(NumeroIA) >= Sqr((Perso.PositionX - Tresors(i).X) ^ 2 + _
                                               (Perso.PositionY - Tresors(i).Y) ^ 2) Then
                Indice = i
                DistanceTresor(NumeroIA) = 0
                Trouver_Tresor_Proche = True
                i = Tresors.Count
            End If
        Next i
    End If
End Function

Public Sub IA_Trouver_Tresor(ByVal i As Long)
    'Le Personnage n°i va chercher le trésor le plus proche.
    Dim j As Long
    Dim Distance As Long
    Dim IndiceTresor As Long
    
    If Tresors.Count > 0 Then
        Distance = Monde.Largeur * AffTerrain.Largeur + Monde.Hauteur * AffTerrain.Hauteur
        For j = 1 To Tresors.Count
            If Distance > Sqr((Persos(i).PositionX - Tresors(j).X) ^ 2 + _
                              (Persos(i).PositionY - Tresors(j).Y) ^ 2) Then
                Distance = Sqr((Persos(i).PositionX - Tresors(j).X) ^ 2 + _
                               (Persos(i).PositionY - Tresors(j).Y) ^ 2)
                IndiceTresor = j
            End If
        Next j
        Persos(i).Aller_Tresor Tresors(IndiceTresor), IndiceTresor, AffTresor.Largeur, AffTresor.Hauteur
    End If
End Sub

Public Function Trouver_Maison_Proche(ByRef Personnage As ClsJeuPerso, _
                                      ByVal NumeroIA As Long, _
                                      ByVal TypeRessource As Integer, _
                                      ByVal TypeBesoin As Integer, _
                                      ByVal TypeService As Integer, _
                                      ByRef Indice As Long) _
                                      As Boolean
    Dim i As Long
    Dim pbm As Long
    Dim IMnia As Long, ubMaison As Long
    
    Indice = -1
    For i = 0 To NombreRechercheParTour
        If IndiceMaison(NumeroIA) > UBound(Maisons()) Then
           DistanceMaison(NumeroIA) = DistanceMaison(NumeroIA) + DistancePalierRechercheMaison
            IndiceMaison(NumeroIA) = 0
        End If
        
        'Recherche les batiments de service.
        With Maisons(IndiceMaison(NumeroIA))
        If TypeService > 0 Then
            'If .Service = TypeService And .Construit And .Definir_Stock >= 1 And _
            '   Persos(IndiceMaison(NumeroIA)).NumeroFief = Persos(NumeroIA).NumeroFief And _
            '   Abs(Personnage.PositionX - .PositionX) < DistanceMaison(NumeroIA) And _
            '   Abs(Personnage.PositionY - .PositionY) < DistanceMaison(NumeroIA) Then
            '    Trouver_Maison_Proche = True
            'End If
            Trouver_Maison_Proche = (.Service = TypeService And .Construit And .Definir_Stock >= 1 And _
               Persos(IndiceMaison(NumeroIA)).NumeroFief = Persos(NumeroIA).NumeroFief And _
               Abs(Personnage.PositionX - .PositionX) < DistanceMaison(NumeroIA) And _
               Abs(Personnage.PositionY - .PositionY) < DistanceMaison(NumeroIA))
            'If NumeroIA = 0 Then
            '    Debug.Print IndiceMaison(NumeroIA) & " : " & DistanceMaison(NumeroIA) & " : " & Trouver_Maison_Proche & " / " & NombreRechercheParTour
            '    If IndiceMaison(NumeroIA) = 0 Then
            '        Debug.Print Maisons(IndiceMaison(NumeroIA)).Service & " = " & TypeService & " : " & Maisons(IndiceMaison(NumeroIA)).Construit & " : " & Maisons(IndiceMaison(NumeroIA)).Definir_Stock & " >= 1" & " : " & Persos(IndiceMaison(NumeroIA)).NumeroFief & " = " & Persos(NumeroIA).NumeroFief & " : " & Abs(Personnage.PositionX - Maisons(IndiceMaison(NumeroIA)).PositionX) & " < " & DistanceMaison(NumeroIA) & " : " & Abs(Personnage.PositionY - Maisons(IndiceMaison(NumeroIA)).PositionY) & " < " & DistanceMaison(NumeroIA)
            '    End If
            'End If
        
        ElseIf TypeBesoin = 5 Then
            'Le personnage cherche un objet.
            'Cherche une maison où il peut acheter des objets.
           
                Trouver_Maison_Proche = (.Fabrique And .Construit And .Definir_Stock >= 1 And _
                        Persos(IndiceMaison(NumeroIA)).NumeroFief = Persos(NumeroIA).NumeroFief And _
                        Abs(Personnage.PositionX - .PositionX) < DistanceMaison(NumeroIA) And _
                        Abs(Personnage.PositionY - .PositionY) < DistanceMaison(NumeroIA))
        
        'If Parametres.Batiment_RessourceFabrique(Maisons(IndiceMaison(NumeroIA)).TypeBatiment) >= 0 Then
        ElseIf .Magasin Or .Fabrique Then
            If TypeRessource = 0 Then
                
                pbm = Parametres.Batiment_RessourceFabrique(.TypeBatiment)
                
                If pbm >= 0 And _
                   TypeBesoin < 5 And _
                   .Construit And .Definir_Stock >= 1 And _
                   Persos(IndiceMaison(NumeroIA)).NumeroFief = Persos(NumeroIA).NumeroFief And _
                   Abs(Personnage.PositionX - .PositionX) < DistanceMaison(NumeroIA) And _
                   Abs(Personnage.PositionY - .PositionY) < DistanceMaison(NumeroIA) Then
                   'DistanceMaison(NumeroIA) >= Sqr((Personnage.PositionX - Maisons(IndiceMaison(NumeroIA)).PositionX) ^ 2 + _
                                                   (Personnage.PositionY - Maisons(IndiceMaison(NumeroIA)).PositionY) ^ 2) Then
                                                   
                    
                    Select Case TypeBesoin
                        'Le personnage a besoin de vie.
                        Case 1:
                            If Parametres.Ressources_GainVie(pbm) > 0 Then
                                Trouver_Maison_Proche = True
                            End If
                        'Le personnage a besoin de vie.
                        Case 2:
                            If Parametres.Ressources_GainEnergie(pbm) > 0 Then
                                Trouver_Maison_Proche = True
                            End If
                        'Le personnage a besoin de magie.
                        Case 3:
                            If Parametres.Ressources_GainMagie(pbm) > 0 Then
                                Trouver_Maison_Proche = True
                            End If
                        'Le personnage a besoin de moral.
                        Case 4:
                            If Parametres.Ressources_GainMoral(pbm) > 0 Then
                                Trouver_Maison_Proche = True
                            End If
                    End Select
                End If
            'Recherche une maison où il peut trouver la ressource passée en paramètre.
            ElseIf Parametres.Batiment_RessourceFabrique(.TypeBatiment) = TypeRessource And _
                   .Construit And _
                   .Definir_Stock >= 1 And _
                   Persos(IndiceMaison(NumeroIA)).NumeroFief = Persos(NumeroIA).NumeroFief And _
                   Abs(Personnage.PositionX - .PositionX) < DistanceMaison(NumeroIA) And _
                   Abs(Personnage.PositionY - .PositionY) < DistanceMaison(NumeroIA) And _
                   DistanceMaison(NumeroIA) >= Sqr((Personnage.PositionX - Maisons(IndiceMaison(NumeroIA)).PositionX) ^ 2 + _
                                                    (Personnage.PositionY - Maisons(IndiceMaison(NumeroIA)).PositionY) ^ 2) Then
                    Trouver_Maison_Proche = True
                'End If
            End If
        End If
        End With
        
        If Trouver_Maison_Proche Then
            Indice = IndiceMaison(NumeroIA)
            DistanceMaison(NumeroIA) = DistancePalierRechercheMaison
            i = UBound(Maisons())
        End If
        IndiceMaison(NumeroIA) = IndiceMaison(NumeroIA) + 1
    Next i
    'Trouver_Maison_Proche = Trouver_Maison_Proche Or Indice < 0
End Function

Public Function Trouver_Emplacement_Maison(ByVal NumeroJoueur As Long, _
                                           ByRef X As Long, _
                                           ByRef Y As Long, _
                                           ByVal TypeMaison As Integer) As Boolean
    Dim i As Long
    Dim j As Long
    Dim Nf As Long 'Variable temporaire utilisée pour mémorisée le numéro du fief.
    'Vérifie si l'emplacement de la maison précédente est réutilisable.
    
    With Persos(NumeroJoueur)
    Nf = .NumeroFief
    X = Maisons(NumeroJoueur).PositionX
    Y = Maisons(NumeroJoueur).PositionY
    If Obstacle(NumeroJoueur, X, Y, Parametres.Batiment_Largeur(TypeMaison), Parametres.Batiment_Hauteur(TypeMaison)) = 0 Then
        DistanceObstacle(Nf) = Chateaux(Nf).Largeur
        Trouver_Emplacement_Maison = True
        Exit Function
    End If
    For i = 0 To NombreRecherchesEmplacementMaison
        DistanceObstacle(Nf) = DistanceObstacle(Nf) + Monde.LargeurCase
        If .Invocation Then
            'Les invocations d'installent près de leur chef.
            If Maisons(.NumeroChef).Visible Then
                'S'installe près de la maison du chef.
                X = Maisons(.NumeroChef).PositionX + Rnd * (DistanceObstacle(Nf) * 2 + Maisons(.NumeroChef).Largeur) - DistanceObstacle(Nf)
                Y = Maisons(.NumeroChef).PositionY + Rnd * (DistanceObstacle(Nf) * 2 + Maisons(.NumeroChef).Hauteur) - DistanceObstacle(Nf)
            Else
                X = Persos(.NumeroChef).PositionX + Rnd * (DistanceObstacle(Nf) * 2 + Persos(.NumeroChef).Largeur) - DistanceObstacle(Nf)
                Y = Persos(.NumeroChef).PositionY + Rnd * (DistanceObstacle(Nf) * 2 + Persos(.NumeroChef).Hauteur) - DistanceObstacle(Nf)
            End If
        ElseIf Chateaux(Nf).Visible And Rnd <= 1 / (Fiefs(Nf).Marches.Count + 1) Or Not Chateaux(Nf).Visible And Fiefs(Nf).Marches.Count = 0 Then
            'Cherche à placer sa maison prêt du chateaux.
            X = Chateaux(Nf).PositionX + Rnd * (DistanceObstacle(Nf) * 2 + Chateaux(Nf).Largeur) - DistanceObstacle(Nf)
            Y = Chateaux(Nf).PositionY + Rnd * (DistanceObstacle(Nf) * 2 + Chateaux(Nf).Hauteur) - DistanceObstacle(Nf)
        Else
            'Cherche à placer sa maison prêt d'un marché.
            j = Int(Rnd * Fiefs(Nf).Marches.Count) 'On choisit un marché au hasard.
            j = Fiefs(Nf).Marches(j + 1)
            X = Maisons(j).PositionX + Rnd * (DistanceObstacle(Nf) * 2 + Maisons(j).Largeur) - DistanceObstacle(Nf)
            Y = Maisons(j).PositionY + Rnd * (DistanceObstacle(Nf) * 2 + Maisons(j).Hauteur) - DistanceObstacle(Nf)
        End If
        If X < 0 Then
            X = Rnd * Monde.LargeurCase
        ElseIf X > Monde.Largeur * Monde.LargeurCase - Parametres.Batiment_Largeur(TypeMaison) Then
            X = Monde.Largeur * Monde.LargeurCase - Parametres.Batiment_Largeur(TypeMaison) - Rnd * Monde.LargeurCase
        End If
        If Y < 0 Then
            Y = Rnd * Monde.LargeurCase
        ElseIf Y > Monde.Hauteur * Monde.HauteurCase - Parametres.Batiment_Hauteur(TypeMaison) Then
            Y = Monde.Hauteur * Monde.HauteurCase - Parametres.Batiment_Hauteur(TypeMaison) - Rnd * Monde.HauteurCase
        End If
        If Obstacle(NumeroJoueur, X, Y, Parametres.Batiment_Largeur(TypeMaison), Parametres.Batiment_Hauteur(TypeMaison)) = 0 Then
            DistanceObstacle(Nf) = Chateaux(Nf).Largeur
            Trouver_Emplacement_Maison = True
            Exit Function
        End If
    Next i
    End With
End Function

Public Sub IA_Travail(ByVal i As Long)
    With Persos(i)
    If .Inactif Then
        'S'il est dans un magasin, il achète ce dont il a besoin.
        If .DansUneMaison And Not QuitterBatiment(i) Then
           'and Not IA_Argent_Besoin(i) Then
            If Maisons(.IndiceMaison).Magasin Then
                'Magasin de ressources.
                If Parametres.Batiment_RessourceFabrique(Maisons(.IndiceMaison).TypeBatiment) = _
                   Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment) Then
                    'L'IA achète le plus de matières premières pour son travail.
                    Maisons(.IndiceMaison).Acheter 0, Persos(i), Parametres, True
                Else
                    Maisons(.IndiceMaison).Acheter Int(QuantiteMaxTransfert * Rnd), Persos(i), Parametres
                End If
            ElseIf Maisons(.IndiceMaison).Fabrique Then
                'Fabrique d'objets.
                'Achete un objet au hasard.
                Z = Int(Rnd * Parametres.Batiment_NombreObjetsFabriques(Maisons(Persos(i).IndiceMaison).TypeBatiment))
                If Maisons(.IndiceMaison).Stock_Objets(Z) > 0 And _
                   Not .Objet_Posseder(Parametres.Batiment_ObjetFabrique(Maisons(.IndiceMaison).TypeBatiment, Z)) Then
                    .Objet_Acheter_Maison Z, Maisons(.IndiceMaison), Parametres, Commentaires
                End If
            Else
                'Batiments de services.
                If Maisons(.IndiceMaison).Service_Utilisable(Persos(i), Parametres) Then
                    Maisons(.IndiceMaison).Service_Utiliser Persos(i), Parametres
                End If
            End If
            IA_Sortir_Batiment i
            'Persos(i).Aller_A Persos(i).PositionX, Persos(i).PositionY + Maisons(i).Hauteur / 2
    
        'Le personnage est dans sa maison,
        ElseIf .DansSaMaison And Not QuitterBatiment(i) Then
            IA_Travailler_Maison i
                
        'Fait ses transactions au chateau
        ElseIf (.DansUnChateau Or .DansSonChateau) And _
               Not QuitterBatiment(i) Or (.DansUneMaison And Maisons(.IndiceMaison).Marche) Then
            IA_Transaction_Chateau i
            IA_Sortir_Batiment i
        
        'Travail.
        'Va chercher de l'argent chez lui si nécessaire.
        ElseIf IA_Argent_Besoin(i) And _
               Maisons(i).Argent >= 1 Then
           IA_Rentrer_Maison i
           
        'Travail de mercenaire.
        ElseIf Maisons(i).Service = 2 Then
            If .Vie = .Definir_Vie Then
                IA_Agressif i: IA_Agressif i: IA_Agressif i
                'IA_Agressif i: IA_Agressif i: IA_Agressif i
                If .Inactif Then
                    If Maisons(i).Argent > 0 Then
                        IA_Rentrer_Maison i
                    End If
                End If
            End If
        
        'Rentre à la maison pour travailler.
        ElseIf .Definir_ressources(Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment)) _
               > MatierePremiereMin And _
               Maisons(i).Definir_Stock < Maisons(i).MaxStock And _
               (Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment) >= 0 Or Maisons(i).Fabrique Or Maisons(i).Service > 0) Then
                    IA_Rentrer_Maison i
        
        'Va au château pour vendre des ressources.
        ElseIf IA_Stock_Exedentaire(i) Or _
               (IA_Argent_Besoin(i) Or _
               (.Definir_ressources(Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)) <= 1 And _
               Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment) >= 0) And _
               Fiefs(.NumeroFief).Definir_ressources(Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)) >= 1) And _
               (Chateaux(.NumeroFief).Visible Or Fiefs(.NumeroFief).Marches.Count > 0) Then
                If .DansSonChateau Then
                    IA_Rentrer_Maison i
                Else
                    IA_Rentrer_Chateau i, True
                End If
        Else
            'S'il n'a rien à faire, il attaque tout ce qui bouge.
            IA_Agressif i
        End If
    End If
    End With
End Sub

Public Sub IA_Travailler_Maison(ByVal i As Long)
    With Persos(i)
    'Il récupère ses économies.
    Maisons(i).Transfert_Argent_Prendre 0, Persos(i), True
    If Maisons(i).Fabrique Then
        'Fabrique des objets.
        Do
            Z = Int(Rnd * Parametres.Batiment_NombreObjetsFabriques(Maisons(i).TypeBatiment))
        Loop Until Parametres.Objet_Epoque(Parametres.Batiment_ObjetFabrique(Maisons(i).TypeBatiment, Z)) - 1 <= Fiefs(Persos(i).NumeroFief).Epoque
        For X = 0 To Parametres.Batiment_NombreObjetsFabriques(Maisons(i).TypeBatiment) - 1
            If Maisons(i).Stock_Objets(X) = 0 And _
               .PeutConstruire_Objet(X, Maisons(i), Parametres) And Parametres.Objet_Epoque(Parametres.Batiment_ObjetFabrique(Maisons(i).TypeBatiment, X)) - 1 <= Fiefs(Persos(i).NumeroFief).Epoque Then
                Z = X
                Exit For
            End If
        Next X
        If .PeutConstruire_Objet(Z, Maisons(i), Parametres) And Maisons(i).Definir_Stock < Maisons(i).MaxStock Then
            .Creer_Objet Z, Maisons(i), Parametres
        Else
            'Récupère ses objets pour les vendre.
            If IA_Argent_Besoin(i) Or Maisons(i).Definir_Stock = Maisons(i).MaxStock Then
                For X = 0 To Parametres.Batiment_NombreObjetsFabriques(Maisons(i).TypeBatiment) - 1
                    If Maisons(i).Stock_Objets(X) >= 1 Then
                        .Objet_Ramasser_Maison X, Maisons(i), Parametres, Commentaires
                        Exit For
                    End If
                Next X
            End If
            IA_Rentrer_Chateau i, True
            'IA_Sortir_Batiment i
        End If
    Else
        'Si le stock du magasin est vide, il le remplit.
        If Maisons(i).Magasin And (Maisons(i).Definir_Stock <= 1 Or Not IA_Argent_Besoin(i)) Then
            If Maisons(i).Definir_Stock = Maisons(i).MaxStock Then
                IA_Rentrer_Chateau i, True
                'IA_Sortir_Batiment i
            Else
                'Maisons(i).Transfert_Stock_Deposer QuantiteMaxTransfert, Persos(i), Parametres
                Maisons(i).Transfert_Stock_Deposer 0.5 * .Definir_ressources(Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment)), Persos(i), Parametres
                IA_Rentrer_Chateau i, True
            End If
        Else
            IA_Rentrer_Chateau i, True
        End If
        'IA_Sortir_Batiment i
    End If
    End With
End Sub

Public Sub IA_Transaction_Chateau(ByVal i As Long)
    Dim NombreObjetsVendus As Integer
    With Persos(i)
    'Vend son stock s'il n'a pas de magasin, pas d'argent ou que les réserves de son magasin sont pleines.
    If Not Maisons(i).Magasin Or _
       IA_Argent_Besoin(i) Or _
       Maisons(i).Definir_Stock >= Maisons(i).MaxStock - 1 Then
         If Maisons(i).Fabrique Then
             'Vend ses objets.
             For X = 0 To .Nombre_Objets_Inventaire - 1
                 If .Objet_Inventaire_Actif(X) Then
                    For Y = 0 To Parametres.Batiment_NombreObjetsFabriques(Maisons(i).TypeBatiment) - 1
                        
                        If .Objet_Inventaire_Type(X) = Parametres.Batiment_ObjetFabrique(Maisons(i).TypeBatiment, Y) Then
                           .Objet_Vendre_Chateau X, Fiefs(.IndiceChateau), Parametres
                            NombreObjetsVendus = NombreObjetsVendus + 1
                        End If
                    Next Y
                 End If
             Next X
             'Vend les objets qu'il a sur lui si cela ne suffit pas.
'             If NombreObjetsVendus = 0 Then
'                For X = 0 To .Nombre_Objets_Equipes - 1
'                    If .Objet_Equipes_Actif(X) Then
'                        For Y = 0 To Parametres.Batiment_NombreObjetsFabriques(Maisons(i).TypeBatiment) - 1
'                            If .Objet_Inventaire_Type(X) = Parametres.Batiment_ObjetFabrique(Maisons(i).TypeBatiment, Y) Then
'                                .ObjetSelectionne = .Objet_Equipes_Actif(X)
'                                .Objet_Equipes_Detruire (X)
'                                .Objet_Vendre_Chateau .ObjetSelectionne, Fiefs(.IndiceChateau), Parametres, False
'                                NombreObjetsVendus = NombreObjetsVendus + 1
'                            End If
'                        Next Y
'                    End If
'                Next X
'             End If
             If NombreObjetsVendus = 0 Then
                'Si n'a pas d'argent et aucun objet à vendre : déclare la banqueroute.
                If Maisons(i).Argent = 0 And Maisons(i).Definir_Stock = 0 Then
                    Maisons(i).Demolir Persos(i), Parametres
                End If
             End If
         ElseIf Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment) >= 0 Then
             Fiefs(Persos(i).NumeroFief).Vendre Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment), 0, Persos(i), Parametres, True
         End If
    End If
    
    IA_Devenir_Empereur i
    
    'Achète les ressources dont il a besoin.
    IA_Acheter_Ressources_Pour_Manger i
    'If Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment) <> 1 Then 'Pas besoin d'acheter de bois.
    'A besoin de ressources pour son métier.
    IA_Acheter_Ressources_Pour_Travailler i
    'Achete des ressources pour construire.
    If Not IA_Argent_Besoin(i) Then
        IA_Acheter_Ressources_Pour_Construire i
    End If
    
    IA_Restaurer_Chateau i
     
     'Il finit par quitter le chateau.
     IA_Sortir_Batiment i
     End With
End Sub

Private Sub IA_Restaurer_Chateau(ByVal i As Long)
    With Persos(i)
    'Restaure le chateau.
    If Chateaux(.NumeroFief).Visible Then
        If Chateaux(.NumeroFief).Definir_Vie < Chateaux(.NumeroFief).MaxVie Then
            'Le chateau est endommagé.
            'If Not IA_Argent_Besoin(i) Then
            '    Fiefs(.NumeroFief).Restaurer_Chateau 5, Persos(i)
            'ElseIf IA_Argent_Riche(i) Then
            '    Fiefs(.NumeroFief).Restaurer_Chateau 50, Persos(i)
            'End If
            Fiefs(.NumeroFief).Restaurer_Chateau Int(.Definir_Argent - (.Definir_Argent * Chateaux(.NumeroFief).Definir_Vie / Chateaux(.NumeroFief).MaxVie)) + 1, Persos(i)
        End If
    End If
    End With
End Sub
Private Sub IA_Devenir_Empereur(ByVal i As Long)
    'Si l'IA le peut, elle devient l'empereur.
    Dim j As Integer
    Dim Nf As Long
    Nf = Persos(i).NumeroFief
    With Persos(i)
    If Chateaux(Nf).Visible Then
        'If Persos(i).PeutDevenir_Empereur(Fiefs(Persos(i).NumeroFief), Parametres) Then
        '    Fiefs(Persos(i).NumeroFief).Empereur_Changer i, True
        'End If
        If Fiefs(Nf).Empereur < 0 Then
            'Messages(0).Ajouter_Message Persos(i).Definir_Argent + .Ressources_TotalPrix + .Objets_TotalPrix & " / " & Fiefs(Nf).Empereur_Prix(Parametres)
            If .Definir_Argent >= Fiefs(Nf).Empereur_Prix(Parametres) Then
                Fiefs(Nf).Empereur_Changer i, True
            ElseIf Persos(i).Definir_Argent + .Ressources_TotalPrix + .Objets_TotalPrix >= Fiefs(Nf).Empereur_Prix(Parametres) Then
                'Vend ses ressources pour devenir empereur.
                Do
                    Fiefs(Nf).Vendre j, 0, Persos(i), Parametres, True
                    j = j + 1
                Loop Until j = .Nombre_CompetencesRessources Or .Definir_Argent >= Fiefs(Nf).Empereur_Prix(Parametres)
                If .Definir_Argent >= Fiefs(Nf).Empereur_Prix(Parametres) Then
                    Fiefs(Nf).Empereur_Changer i, True
                Else
                    'Vend son inventaire pour devenir empereur.
                    j = 0
                    Do
                        If .Objet_Inventaire_Actif(j) Then
                            .Objet_Vendre_Chateau j, Fiefs(Nf), Parametres, True
                        End If
                        j = j + 1
                    Loop Until j = .Nombre_Objets_Inventaire Or .Definir_Argent >= Fiefs(Nf).Empereur_Prix(Parametres)
                    If .Definir_Argent >= Fiefs(Nf).Empereur_Prix(Parametres) Then
                        Fiefs(Nf).Empereur_Changer i, True
                    Else
                        'Vend le reste de ses objets pour devenir empereur.
                        j = 0
                        Do
                            If .Objet_Equipes_Actif(j) Then
                                .Objet_Ranger j, Parametres, Commentaires
                                .Objet_Vendre_Chateau 0, Fiefs(Nf), Parametres, True
                            End If
                            j = j + 1
                        Loop Until j = .Nombre_Objets_Equipes Or .Definir_Argent >= Fiefs(Nf).Empereur_Prix(Parametres)
                    End If
                End If
            End If
        End If
    End If
    End With
End Sub

Public Sub IA_Vente_Surplus_Objets(ByVal i As Long)
    Dim j As Integer
    With Persos(i)
        If (.DansSonChateau Or .DansUnChateau) Or (.DansUneMaison And Maisons(.IndiceMaison).Marche) Then
            For j = .Nombre_Objets_Inventaire - 1 To 0 Step -1
                If .Objet_Inventaire_Actif(j) Then
                   If Not Parametres.Objet_UtilisationUnique(.Objet_Inventaire_Type(j)) And _
                      ((.Objet_NombreType(.Objet_Inventaire_Type(j)) > 1 And _
                       Parametres.Objet_Categorie(.Objet_Inventaire_Type(j)) <> 0) Or _
                      (.Objet_NombreType(.Objet_Inventaire_Type(j)) > 2 And _
                       Parametres.Objet_Categorie(.Objet_Inventaire_Type(j)) = 0)) Then
                        .Objet_Vendre_Chateau j, Fiefs(.NumeroFief), Parametres
                    End If
                End If
            Next j
        End If
    End With
End Sub

Public Function Obstacle(ByVal NumeroJoueur As Long, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         Optional ByVal Largeur As Integer = 0, _
                         Optional ByVal Hauteur As Integer = 0, _
                         Optional ByVal MemoriserIndiceObstacle As Boolean) As Integer
    'Cette fonction renvoie l'obstacle situé en x, y.
    Dim i As Long, Pne As Long
    Pne = Persos(NumeroJoueur).NumeroEquipe
    
    'Vérifie les personnages ennemis.
    For i = 0 To UBound(Persos())
        If Est_Superpose(X, _
                         Y, _
                         Persos(i).PositionX, _
                         Persos(i).PositionY, _
                         Largeur, _
                         Hauteur, _
                         Persos(i).Largeur, _
                         Persos(i).Hauteur) Then
            If Persos(i).NumeroEquipe <> Pne And _
               Persos(i).Attaquable Then
                If i <> NumeroJoueur Then
                    Obstacle = 9
                    If MemoriserIndiceObstacle Then IndiceObstacle = i
                    Exit For
                End If
            End If
        End If
    Next i
    
    'Vérifie les trésors.
    If Obstacle = 0 And Tresors.Count > 0 Then
        For i = 1 To Tresors.Count
            If Est_Superpose(X, _
                             Y, _
                             Tresors(i).X, _
                             Tresors(i).Y, _
                             Largeur, _
                             Hauteur, _
                             AffTresor.Largeur, _
                             AffTresor.Hauteur) Then
                Obstacle = 13
                If MemoriserIndiceObstacle Then
                    IndiceObstacle = i
                End If
                Exit For
            End If
        Next i
    End If
    
    If Obstacle = 0 Then
        'Vérifie les chateaux.
        For i = 0 To UBound(Chateaux())
            If Est_Superpose(X, _
                             Y, _
                             Chateaux(i).PositionX, _
                             Chateaux(i).PositionY, _
                             Largeur, _
                             Hauteur, _
                             Chateaux(i).Largeur, _
                             Chateaux(i).Hauteur) And _
               Chateaux(i).Visible Then
                If Fiefs(i).NumeroEquipe = Pne Then
                    If i = Persos(NumeroJoueur).NumeroFief Then
                        Obstacle = 6
                    Else
                       'Cas où s'est un chateau allié.
                       Obstacle = 8
                       If MemoriserIndiceObstacle Then IndiceObstacle = i
                    End If
                Else 'C'est pas un allié.
                       Obstacle = 11
                       If MemoriserIndiceObstacle Then IndiceObstacle = i
                End If
                Exit For
            End If
        Next i
    End If
    
    'Vérifie les maisons.
    If Obstacle = 0 Then
        For i = 0 To UBound(Maisons())
            If Maisons(i).Visible And _
                Est_Superpose(X, _
                              Y, _
                              Maisons(i).PositionX, _
                              Maisons(i).PositionY, _
                              Largeur, _
                              Hauteur, _
                              Maisons(i).Largeur, _
                              Maisons(i).Hauteur) Then
                If i = NumeroJoueur Then 'Cas où c'est la maison du joueur.
                    If Maisons(i).Construit And Maisons(i).Definir_Vie = Maisons(i).MaxVie Then
                        'Cas où c'est la maison du joueur à utiliser.
                        Obstacle = 5
                    Else
                        'Cas où c'est la maison du joueur à construire.
                        Obstacle = 4
                    End If
                Else
                    'Cas ou c'est une maison qui n'appartient pas au joueur.
                    If Persos(i).NumeroEquipe = Pne Then
                        'Cas où c'est une maison alliée.
                        'If (Maisons(i).Magasin Or Maisons(i).Fabrique Or Maisons(i).Service > 0) And _
                            Maisons(i).Construit Then
                            'If Maisons(i).Service = 1 Then
                            '    'Cas où c'est un temple.
                            '    Obstacle = 12
                            'Else
                                Obstacle = 7
                            'End If
                        'Else
                        '    Obstacle = 999
                        'End If
                    Else 'Ce n'est pas une maison alliée.
                        Obstacle = 10
                    End If
                End If
                If MemoriserIndiceObstacle Then IndiceObstacle = i
                Exit For
            End If
        Next i
    End If
    
    'Vérifie les ressources récoltables.
    If Obstacle = 0 Then
        For i = 0 To UBound(Ressources())
            If Est_Superpose(X, _
                             Y, _
                             Ressources(i).PositionX, _
                             Ressources(i).PositionY, _
                             Largeur, _
                             Hauteur, _
                             AffRessources.Largeur, _
                             AffRessources.Hauteur) Then
                Obstacle = 1
                If MemoriserIndiceObstacle Then
                    IndiceObstacle = i
                End If
                Exit For
            End If
        Next i
    End If
    
    'Vérifie les décors.
    If Obstacle = 0 Then
        For i = 1 To Decors.Count
            If Est_Superpose(X, _
                             Y, _
                             Decors(i).X, _
                             Decors(i).Y, _
                             Largeur, _
                             Hauteur, _
                             Decors(i).Largeur, _
                             Decors(i).Hauteur) Then
                Obstacle = 16
                If MemoriserIndiceObstacle Then
                    IndiceObstacle = i
                End If
                Exit For
            End If
        Next i
    End If
    
    'Vérifie le relief du terrain.
    If Obstacle = 0 Then
        If Not Monde.Constructible(X, Y, Largeur, Hauteur) Then
            Obstacle = 999
        End If
    End If
End Function

Public Function ObstacleMort(ByVal NumeroJoueur As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             Optional ByVal Largeur As Long = 0, _
                             Optional ByVal Hauteur As Long = 0, _
                             Optional ByVal MemoriserIndiceObstacle As Boolean) As Integer
    'Cette fonction renvoie tous les endroits où un mort peut aller.
    Dim i As Long, Pne As Long
    Pne = Persos(NumeroJoueur).NumeroEquipe
    
    'Vérifie les chateaux.
    For i = 0 To UBound(Chateaux())
        If Est_Superpose(X, _
                         Y, _
                         Chateaux(i).PositionX, _
                         Chateaux(i).PositionY, _
                         Largeur, _
                         Hauteur, _
                         Chateaux(i).Largeur, _
                         Chateaux(i).Hauteur) And _
           Chateaux(i).Visible Then
            If i = Persos(NumeroJoueur).NumeroFief Then
                ObstacleMort = 6
            'Else
                'ObstacleMort = 999
            End If
            If MemoriserIndiceObstacle Then IndiceObstacle = i
            Exit For
        End If
    Next i
    
    'Vérifie les temples éventuels.
    If ObstacleMort = 0 Then
        For i = 0 To UBound(Maisons())
            If Maisons(i).Visible And _
                Est_Superpose(X, _
                              Y, _
                              Maisons(i).PositionX, _
                              Maisons(i).PositionY, _
                              Largeur, _
                              Hauteur, _
                              Maisons(i).Largeur, _
                              Maisons(i).Hauteur) Then
                If i = NumeroJoueur Then 'Cas où c'est la maison du joueur.
                    If Maisons(i).Construit And _
                       Maisons(i).Service = 1 Then
                        'Cas où la maison du joueur est un temple.
                        ObstacleMort = 5
                    'Else
                        'Autres cas.
                        'ObstacleMort = 999
                    End If
                Else
                    'Cas ou c'est une maison qui n'appartient pas au joueur.
                    If Persos(i).NumeroEquipe = Pne And Maisons(i).Construit And _
                       Maisons(i).Service = 1 Then
                        'Cas où c'est un temple alliée.
                        ObstacleMort = 12
                    'Else 'Ce n'est pas une maison alliée.
                        'ObstacleMort = 999
                    End If
                End If
                If MemoriserIndiceObstacle Then IndiceObstacle = i
                Exit For
            End If
        Next i
    End If
    
    'Vérifie les tombes des collegues.
    If ObstacleMort = 0 Then
        For i = 1 To Decors.Count
            If Est_Superpose(X, _
                             Y, _
                             Decors(i).X, _
                             Decors(i).Y, _
                             Largeur, _
                             Hauteur, _
                             Decors(i).Largeur, _
                             Decors(i).Hauteur) Then
                If Decors(i).Nom <> "" Then
                    ObstacleMort = 16
                    If MemoriserIndiceObstacle Then
                        IndiceObstacle = i
                    End If
                    Exit For
                End If
            End If
        Next i
    End If
End Function

Public Function ObstacleAmi(ByVal NumeroJoueur As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long) As Integer  ', _
                            ByVal Invocation As Boolean) As Integer
    'Cette fonction renvoie le personnage de la même équipe.
    Dim i As Long, Pne As Long
    Pne = Persos(NumeroJoueur).NumeroEquipe
    
    'Vérifie les personnages alliés.
    For i = 0 To UBound(Persos())
        With Persos(i)
        If Est_Superpose(X, _
                         Y, _
                         .PositionX, _
                         .PositionY, _
                         , _
                         , _
                         .Largeur, _
                         .Hauteur) Then
            'If Invocation Or Not .Invocation Then
                If .NumeroEquipe = Pne And _
                   .Attaquable And _
                   i <> NumeroJoueur Then
                    IndiceObstacle = i
                    ObstacleAmi = 1
                    Exit Function
                End If
            'End If
        End If
        End With
    Next i
    
    'Vérifie les maisons alliées.
    If ObstacleAmi = 0 Then
        For i = 0 To UBound(Maisons())
            If Maisons(i).Visible And _
               Persos(i).NumeroEquipe = Pne And _
               i <> NumeroJoueur And _
                Est_Superpose(X, _
                              Y, _
                              Maisons(i).PositionX, _
                              Maisons(i).PositionY, _
                              , _
                              , _
                              Maisons(i).Largeur, _
                              Maisons(i).Hauteur) Then
                ObstacleAmi = 2
                IndiceObstacle = i
                Exit Function
            End If
        Next i
    End If
    
    If ObstacleAmi = 0 Then
        'Vérifie les chateaux amis.
        For i = 0 To UBound(Chateaux())
            If Chateaux(i).Visible And _
               Est_Superpose(X, _
                             Y, _
                             Chateaux(i).PositionX, _
                             Chateaux(i).PositionY, _
                             , _
                             , _
                             Chateaux(i).Largeur, _
                             Chateaux(i).Hauteur) Then
                If i = Persos(NumeroJoueur).NumeroFief Then
                    ObstacleAmi = 3
                    IndiceObstacle = i
                ElseIf Fiefs(i).NumeroEquipe = Pne Then
                    ObstacleAmi = 4
                    IndiceObstacle = i
                End If
                Exit Function
            End If
        Next i
    End If
End Function

Public Function ObstacleEnnemi(ByVal Perso As ClsJeuPerso, _
                               ByVal X As Long, _
                               ByVal Y As Long, _
                               Optional ByVal Largeur As Long = 0, _
                               Optional ByVal Hauteur As Long = 0) As Boolean
    'Cette fonction renvoie un personnage adverse.
    Dim i As Long, Pne As Long
    Pne = Perso.NumeroEquipe
    
    'Vérifie les personnages ennemis.
    For i = 0 To UBound(Persos())
        With Persos(i)
        If Est_Superpose(X, _
                         Y, _
                         .PositionX, _
                         .PositionY, _
                         Largeur, _
                         Hauteur, _
                         .Largeur, _
                         .Hauteur) Then
            If .NumeroEquipe <> Pne And _
               .Attaquable Then
                IndiceObstacle = i
                ObstacleEnnemi = True
                i = UBound(Persos())
            End If
        End If
        End With
    Next i
End Function

Public Function Est_Superpose(ByVal X1 As Long, _
                              ByVal Y1 As Long, _
                              ByVal X2 As Long, _
                              ByVal Y2 As Long, _
                              Optional ByVal Largeur1 As Long = 0, _
                              Optional ByVal Hauteur1 As Long = 0, _
                              Optional ByVal Largeur2 As Long = 0, _
                              Optional ByVal Hauteur2 As Long = 0) As Boolean
    'Vérifie si les 2 objets sont superposés.
    
    'Est_Superpose = False
    
    If Largeur1 = 0 And Hauteur1 = 0 Then
        Est_Superpose = X1 >= X2 And _
                        X1 <= X2 + Largeur2 And _
                        Y1 >= Y2 And _
                        Y1 <= Y2 + Hauteur2
    Else
        Est_Superpose = _
          ((X1 <= X2 And X2 <= X1 + Largeur1 Or _
            X1 <= X2 + Largeur2 And X2 + Largeur2 <= X1 + Largeur1) And _
           (Y1 <= Y2 And Y2 <= Y1 + Hauteur1 Or _
            Y1 <= Y2 + Hauteur2 And Y2 + Hauteur2 <= Y1 + Hauteur1) Or _
           (X2 <= X1 And X1 <= X2 + Largeur2 Or _
            X2 <= X1 + Largeur1 And X1 + Largeur1 <= X2 + Largeur2) And _
           (Y2 <= Y1 And Y1 <= Y2 + Hauteur2 Or _
            Y2 <= Y1 + Hauteur1 And Y1 + Hauteur1 <= Y2 + Hauteur2) Or _
           (X1 >= X2 And X1 + Largeur1 <= X2 + Largeur2 And _
            Y1 >= Y2 And Y1 + Largeur1 <= Y2 + Largeur2) Or _
           (X2 >= X1 And X2 + Largeur2 <= X1 + Largeur1 And _
            Y2 >= Y1 And Y2 + Largeur2 <= Y1 + Largeur1)) Or _
           (X1 < X2 And X1 + Largeur1 > X2 + Largeur2 And _
            Y2 < Y1 And Y1 + Hauteur1 < Y2 + Hauteur2 Or _
            X2 < X1 And X2 + Largeur2 > X1 + Largeur1 And _
            Y1 < Y2 And Y2 + Hauteur2 < Y1 + Hauteur1)
    End If
End Function


Public Function ObstacleChateau(ByVal NumeroJoueur As Long, _
                                ByVal PositionX As Long, _
                                ByVal PositionY As Long, _
                                ByVal Largeur As Long, _
                                ByVal Hauteur As Long) As Boolean
    Dim i As Long
    'Vérifie le relief du terrain.
    If Not ObstacleChateau Then
        ObstacleChateau = Monde.Constructible(X, Y, Largeur, Hauteur)
    End If
    
    'Vérifie que le chateau n'est pas superposé à un autre.
    For i = 0 To UBound(Chateaux())
        If Est_Superpose(PositionX, _
                         PositionY, _
                         Chateaux(i).PositionX, _
                         Chateaux(i).PositionY, _
                         Largeur, _
                         Hauteur, _
                         Chateaux(i).Largeur, _
                         Chateaux(i).Hauteur) And _
           Chateaux(i).Visible Then
            ObstacleChateau = True
            i = UBound(Chateaux())
        End If
    Next i
End Function

Public Property Let Nombre_IAs(ByVal Valeur As Long)
    LPerso = Valeur - 1
        
    ReDim Preserve DistanceRessource(LPerso)
    ReDim Preserve DistanceTresor(LPerso)
    ReDim Preserve DistanceMaison(LPerso)
    ReDim Preserve DistanceAdversaire(LPerso)
    ReDim Preserve DistanceMaisonEnnemi(LPerso)
    ReDim Preserve DistanceChateauEnnemi(LPerso)
    ReDim Preserve IndiceMaison(LPerso)
    ReDim Preserve IndiceAdversaire(LPerso)
    ReDim Preserve IndiceMaisonEnnemi(LPerso)
    ReDim Preserve IndiceChateauEnnemi(LPerso)
    ReDim Preserve AnglePatrouille(LPerso)
    AnglePatrouille(LPerso) = Rnd * 6.28 'Place l'angle de patrouille à un départ aléatoire.
        
    NombreIAParTour = LPerso * VitesseIA
    If NombreIAParTour = 0 Then
        NombreIAParTour = 1
    End If
    NombreRechercheParTour = LPerso * VitesseRechercheIA
    If NombreRechercheParTour = 0 Then
        NombreRechercheParTour = 1
    End If
    ReDim Preserve BatimentAConstruire(LPerso)
    ReDim Preserve RessourceAFabriquer(LPerso)
    ReDim Preserve ObjetAFabriquer(LPerso)
    ReDim Preserve QuantiteAFabriquer(LPerso)
    ReDim Preserve QuitterBatiment(LPerso)
End Property
