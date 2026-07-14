Attribute VB_Name = "ModCapacitesSpe"
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

Public Sub Gestion_Utilisation_Speciale()
    Dim i, j, k As Long
    Dim Objet As Integer
    Dim Competence As Integer
    Dim NiveauCompetence As Integer
    'Dim Temp As Double
    j = UBound(Persos())
    'Messages(0).Ajouter_Message Persos(Noperso).NumeroProjectileLance & " : " & Persos(Noperso).ProjectileEnCours
    For i = 0 To j
        With Persos(i)
        If .NumeroProjectileLance >= 0 And Not .ProjectileEnCours Then
            Creer_Invocation .NumeroProjectileLance, .NombreProjectiles, .Numero, .Effet_Attaque, True
            'Select Case .Action
            'Case 300:
                'Temp = Perso_Lancer_Projectile_Perso(Persos(i), Persos(.IndicePerso))
                'If Temp >= 1 Then
                '    Creer_Invocation .LanceProjectile, 1, .Numero, .Effet_Attaque, Temp + 1
                'End If
            'Case 301:
                'Temp = Perso_LancerProjectile_Batiment(Persos(i), Maisons(.IndiceMaison), Persos(.IndiceMaison))
                'If Temp >= 1 Then
                '    Creer_Invocation .LanceProjectile, 1, .Numero, .Effet_Attaque, Temp + 1
                
                'End If
            'Case 302:
                'Temp = Perso_LancerProjectile_Batiment(Persos(i), Chateaux(.IndiceChateau), Persos(.IndicePerso), .IndiceChateau)
                'If Temp >= 1 Then
                '    Creer_Invocation .LanceProjectile, 1, .Numero, .Effet_Attaque, Temp + 1
                'End If
            'End Select
            .NumeroProjectileLance = -1
        End If
        If .NumeroObjetUtilise >= 0 Then
            Objet = .NumeroObjetUtilise
            NiveauCompetence = .Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(Objet))
            
            Competence = Parametres.Objet_Competence(Objet)
            If Parametres.Objet_Invocation(Objet) >= 0 Then
                Creer_Invocation Parametres.Objet_Invocation(Objet), _
                                 .Objet_Equipe_InvocationSup(NiveauCompetence, Parametres.Objet_Quantite(Objet)), _
                                 i, _
                                 Parametres.Objet_EffetAttaque(Objet), , _
                                 Int(.Objet_Equipe_InvocationSup(NiveauCompetence, Parametres.Objet_QuantiteMax(Objet))), _
                                 Parametres.Objet_InvocationIA_Fichier(Objet)
            End If
            If Parametres.Objet_PorteeUtilisation(Objet) > 0 Then
                Effets_Proches Persos(i), _
                               Parametres.Objet_CreerObjetAllie(Objet), _
                               Parametres.Objet_CreerObjetEnnemi(Objet), _
                               .Objet_Equipe_PorteeSup(NiveauCompetence, Parametres.Objet_PorteeUtilisation(Objet)), _
                               .Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_GainVie(Objet)) * .Manger_Ressources_Coef(0), _
                               .Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_GainEnergie(Objet)) * .Manger_Ressources_Coef(1), _
                               .Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_GainMagie(Objet)) * .Manger_Ressources_Coef(2), _
                               .Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_GainMoral(Objet)) * .Manger_Ressources_Coef(3), _
                               Parametres.Objet_Resurrection(Objet), _
                               .Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_Quantite(Objet)), _
                               Parametres.Objet_EffetAttaque(Objet), _
                               NiveauCompetence, _
                               Parametres.Objet_AttaqueResistance(Objet)
            Else
                If Parametres.Objet_SelfResurrection(Objet) Then
                    'Auto resurrection.
                    Ressusciter_Perso i, Parametres.Objet_EffetAttaque(Objet)
                End If
                .RecuperationVie = .RecuperationVie + .Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_GainVie(Objet)) * .Manger_Ressources_Coef(0)
                If .Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_GainVie(Objet)) * .Manger_Ressources_Coef(0) > 0 Then
                    AffDegat.Ajouter_Degat .PositionX + .Largeur / 2, .PositionY, .Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_GainVie(Objet)) * .Manger_Ressources_Coef(0), , 1
                End If
                .RecuperationEnergie = .RecuperationEnergie + .Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_GainEnergie(Objet)) * .Manger_Ressources_Coef(1)
                .RecuperationMagie = .RecuperationMagie + .Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_GainMagie(Objet)) * .Manger_Ressources_Coef(2)
                .RecuperationMoral = .RecuperationMoral + .Objet_Equipe_BonusSup(NiveauCompetence, Parametres.Objet_GainMoral(Objet)) * .Manger_Ressources_Coef(3)
                If Parametres.Objet_Invocation(Objet) < 0 Then
                    'EffetsPersos(i).Activer_Effet Parametres.Objet_EffetAttaque(Objet), .PositionX, .PositionY, .Largeur, .Hauteur, Parametres
                    Activer_Effet Parametres.Objet_EffetAttaque(Objet), .PositionX, .PositionY, .Largeur, .Hauteur, Parametres
                End If
            End If
            If Parametres.Objet_Impot(Objet) Then Prelever_Impots Persos(i), Parametres.Objet_Quantite(Objet), Parametres.Objet_EffetAttaque(Objet), .Objet_Equipe_PorteeSup(NiveauCompetence, Parametres.Objet_PorteeUtilisation(Objet))
            If .Definir_Resistance(Parametres.Objet_AttaqueResistance(Objet)) > Parametres.ResistanceMinimum Then
                If Rnd < .Definir_Resistance(Parametres.Objet_AttaqueResistance(Objet)) Then
                    Effet_Sur_Perso Persos(i), Objet, NiveauCompetence, 0, i
                End If
            End If
            .NumeroObjetUtilise = -1
        End If
        End With
    Next i
End Sub

Private Function Prelever_Impots(ByVal Perso As ClsJeuPerso, _
                                 ByVal Quantite As Single, _
                                 ByVal Effet As Integer, _
                                 ByVal Portee As Long)
    Dim i As Long
    Dim Temp As Long
    Dim Total As Long
    Dim Nf As Long: Dim No As Long  'Variables temporaires pour gagner en vitesse.
    Nf = Perso.NumeroFief
    No = Perso.Numero
    'Prélève des impôts sur tous les personnages de son fief.
    For i = 0 To LPerso
        If Persos(i).NumeroFief = Nf Then
            If i <> No Then
                If Persos(i).Vivant Then
                    If Portee >= Distance_Persos(Persos(i), Perso) Then
                        Temp = Persos(i).Definir_Argent * Quantite
                        Persos(i).Definir_Argent = Persos(i).Definir_Argent - Temp
                        Total = Total + Temp
                        AffDegat.Ajouter_Degat Persos(i).PositionX + Persos(i).Largeur / 2, Persos(i).PositionY, -Temp, , 2
                    End If
                End If
                If Maisons(i).Visible Then
                    If Maisons(i).Argent >= 1 / Quantite Then
                    If Portee >= Sqr((Maisons(i).PositionX - Perso.PositionX) ^ 2 + (Maisons(i).PositionY - Perso.PositionY) ^ 2) Then
                        Temp = Maisons(i).Argent * Quantite
                        Maisons(i).Argent = Maisons(i).Argent - Temp
                        Total = Total + Temp
                        AffDegat.Ajouter_Degat Maisons(i).PositionX + Maisons(i).Largeur / 2, Maisons(i).PositionY, -Temp, , 2
                    End If
                    End If
                End If
            End If
        End If
    Next i
    Perso.Definir_Argent = Perso.Definir_Argent + Total
    Activer_Effet Effet, Perso.PositionX, Perso.PositionY, Perso.Largeur, Perso.Hauteur, Parametres
    AffDegat.Ajouter_Degat Perso.PositionX + Perso.Largeur / 2, Perso.PositionY, Total, True, 2
End Function

Public Function Creer_Invocation(ByVal NumeroInvocation As Integer, _
                                 ByVal Quantite As Single, _
                                 ByVal NumeroChef As Long, _
                                 ByVal EffetVisuel As Integer, _
                                 Optional ByVal Projectile As Boolean = False, _
                                 Optional ByVal QuantiteMax As Integer = 0, _
                                 Optional ByVal FichierIA As String = "") As Integer
    'Renvoie l'indice de l'invocation créée.
    Dim i As Long, j As Long, k As Long, l As Long
    Dim X, Y As Long
    'Dim NombreServiteurs As Long 'Mémorise le nombre de serviteurs que possède son chef.
    'On Error GoTo Erreur
    'Debug.Print NumeroInvocation & ", " & Quantite & ", " & NumeroChef
    'NombreServiteurs = Persos(NumeroChef).NombreSoldats
    Quantite = Quantite_Invocations_Aleatoire(Quantite)
    For i = 0 To Quantite - 1
        'Supprime les invocations en trop si l'on dépasse la quantité max.
        k = 0
        If QuantiteMax > 1 Then
            For j = 0 To UBound(Persos())
                If Persos(j).NumeroChef = NumeroChef Then
                    If Persos(j).Race = NumeroInvocation Then
                        k = k + 1
                        If k >= QuantiteMax Then
                            Persos(j).Vivant = False
                            Exit For
                        End If
                    End If
                End If
            Next j
        End If
    
        k = -1
        'On essaie de voir si dans le tableau de personnages, il y a déjà une invocation de morte à remplacer.
        For j = 0 To UBound(Persos())
            If Persos(j).Invocation And Not Persos(j).Vivant And Not Maisons(j).Visible Then
                k = j
                j = UBound(Persos())
            End If
        Next j
        If k = -1 Then 'Sinon, on ajoute un nouveau personnage.
            k = Perso_Ajouter_Nouveau
        Else
            Persos(k).Init k, Parametres
        End If
        
        With Persos(k)
        .NumeroFief = Persos(NumeroChef).NumeroFief
        .NumeroEquipe = Persos(NumeroChef).NumeroEquipe
        '.Changer_Race Parametres.Monstre_Peuple(NumeroInvocation), _
                               Parametres, _
                               Parametres.Monstre_Grade(NumeroInvocation), _
                               NumeroInvocation
        .Projectile = Projectile
        .Changer_Race_Personnage NumeroInvocation, Parametres
        .Changer_Race_Personnage_Equipement Parametres, Commentaires
        
        .IA = Not Projectile
        If FichierIA = "" Then
            .ChoisirIA = False
        Else
            'Définit une IA.
            .ChoisirIA = True
            .FichierIA = FichierIA
            .NumeroIA = Charger_IA(.FichierIA)
            IA_InitPerso k
        End If
        .IA_Ordre = 0
        .NumeroChef = NumeroChef
        
        If Projectile Then
            'Prend les caractéristiques du lanceur.
            .Definir_MaxAttaque = Persos(NumeroChef).Definir_MaxAttaque
            .Definir_Carac_CompetenceSpeciales(1) = Persos(NumeroChef).Definir_Niveau_CompetenceSpeciales(1)
            If Persos(NumeroChef).Definir_NombreAttaqueCauseEffet > 0 Then
                For l = 1 To Persos(NumeroChef).Definir_NombreAttaqueCauseEffet
                    .Definir_Carac_CompetenceObjets(Parametres.Objet_Competence(Persos(NumeroChef).Definir_AttaqueCauseEffet(l))) = _
                        Persos(NumeroChef).Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(Persos(NumeroChef).Definir_AttaqueCauseEffet(l)))
                Next l
            End If
        End If
        
        .Actualiser_Bonus
        '.Vie = .Definir_MaxVie
        .Vivant = True
        .Energie = .Definir_MaxEnergie
        .Magie = .Definir_MaxMagie
        .Moral = .Definir_MaxMoral
        '.Attaque = .Definir_MaxAttaque
        '.Defense = .Definir_MaxDefense
        If .ChoisirApparence Then
            If .FichierApparence = "" Then
                'Tire une apparence au hasard si le joueur n'en n'a pas.
                .FichierApparence = AffPerso.Tirer_Apparence_Au_Hasard(Parametres.Race_CheminApparence(.Race, .Feminin), Parametres, Parametres.Race_CheminAbsolu(.Race))
            End If
            .ApparenceChoisie = AffPerso.Charger_Hero(.FichierApparence, .Feminin)
            .Apparence = .ApparenceChoisie
            '.Definir_Apparence = .Apparence
        Else
            AffApparence.Tirer_Apparence_Au_Hasard Persos(k), Parametres
            AffApparence.Definir_Indices Persos(k)
        End If
        
        
        If .Projectile Then
            Persos(NumeroChef).ProjectileEnCours = Persos(NumeroChef).ProjectileEnCours + 1
            .PositionX = Persos(NumeroChef).PositionX + Persos(NumeroChef).Largeur / 2 - .Largeur / 2
            .PositionY = Persos(NumeroChef).PositionY + Persos(NumeroChef).Hauteur / 2 - .Hauteur / 2
            If i = 0 Then
                Select Case Persos(.NumeroChef).Action
                Case 300:
                    X = Persos(Persos(NumeroChef).IndicePerso).PositionX + Persos(Persos(NumeroChef).IndicePerso).Largeur / 2
                    Y = Persos(Persos(NumeroChef).IndicePerso).PositionY + Persos(Persos(NumeroChef).IndicePerso).Hauteur / 2
                    .ProjectileCible = 0
                Case 301:
                    X = Maisons(Persos(NumeroChef).IndiceMaison).PositionX + Maisons(Persos(NumeroChef).IndiceMaison).Largeur / 2
                    Y = Maisons(Persos(NumeroChef).IndiceMaison).PositionY + Maisons(Persos(NumeroChef).IndiceMaison).Hauteur / 2
                    .ProjectileCible = 1
                Case 302:
                    X = Chateaux(Persos(NumeroChef).IndiceChateau).PositionX + Chateaux(Persos(NumeroChef).IndiceChateau).Largeur / 2
                    Y = Chateaux(Persos(NumeroChef).IndiceChateau).PositionY + Chateaux(Persos(NumeroChef).IndiceChateau).Hauteur / 2
                    .ProjectileCible = 2
                Case Else:
                    X = Persos(NumeroChef).CibleX
                    Y = Persos(NumeroChef).CibleY
                    'Persos(NumeroChef).Arreter
                    Persos(NumeroChef).Action = 1 'Continue de marcher
                    'Persos(NumeroChef).Aller_A Persos(NumeroChef).DirectionX, Persos(NumeroChef).DirectionY
                    
                    .ProjectileCible = Persos(NumeroChef).ProjectileCible
                End Select
                l = Sqr((X - .PositionX) ^ 2 + (Y - .PositionY) ^ 2)
                If l = 0 Then l = 1
                .Aller_A Persos(NumeroChef).Definir_PorteeAttaque * 2 / l * (X - .PositionX) + .PositionX, _
                         Persos(NumeroChef).Definir_PorteeAttaque * 2 / l * (Y - .PositionY) + .PositionY
            Else 'Projectiles supplémentaires.
                Select Case Persos(.NumeroChef).Action
                Case 300:
                    .ProjectileCible = 0
                Case 301:
                    .ProjectileCible = 1
                Case 302:
                    .ProjectileCible = 2
                Case Else:
                    .ProjectileCible = Persos(NumeroChef).ProjectileCible
                End Select
                .Aller_A Persos(NumeroChef).Definir_PorteeAttaque * 2 / l * (X - .PositionX) + .PositionX + IIf(i Mod 2 = 0, i * 32, -i * 32), _
                         Persos(NumeroChef).Definir_PorteeAttaque * 2 / l * (Y - .PositionY) + .PositionY + IIf(i Mod 2 = 0, i * 32, -i * 32)
            End If
        Else 'On localise l'invocation en étoile à 8 branches.
            Select Case i Mod 8
            Case 0: 'A droite.
                .ChefDecalageX = Persos(NumeroChef).Largeur * (Int(i / 8) + 1)
                .ChefDecalageY = 0
            Case 1: 'A gauche.
                .ChefDecalageX = -Persos(NumeroChef).Largeur * (Int(i / 8) + 1)
                .ChefDecalageY = 0
            Case 2: 'En bas.
                .ChefDecalageX = 0
                .ChefDecalageY = Persos(NumeroChef).Hauteur * (Int(i / 8) + 1)
            Case 3: 'En haut.
                .ChefDecalageX = 0
                .ChefDecalageY = -Persos(NumeroChef).Hauteur * (Int(i / 8) + 1)
            Case 4: 'En haut à gauche.
                .ChefDecalageX = -Persos(NumeroChef).Largeur * (Int(i / 8) + 1)
                .ChefDecalageY = -Persos(NumeroChef).Hauteur * (Int(i / 8) + 1)
            Case 5: 'En haut à droite.
                .ChefDecalageX = Persos(NumeroChef).Largeur * (Int(i / 8) + 1)
                .ChefDecalageY = -Persos(NumeroChef).Hauteur * (Int(i / 8) + 1)
            Case 6: 'En bas à gauche.
                .ChefDecalageX = -Persos(NumeroChef).Largeur * (Int(i / 8) + 1)
                .ChefDecalageY = Persos(NumeroChef).Hauteur * (Int(i / 8) + 1)
            Case 7: 'En bas à droite.
                .ChefDecalageX = Persos(NumeroChef).Largeur * (Int(i / 8) + 1)
                .ChefDecalageY = Persos(NumeroChef).Hauteur * (Int(i / 8) + 1)
            End Select
            .PositionX = Persos(NumeroChef).PositionX + .ChefDecalageX
            .PositionY = Persos(NumeroChef).PositionY + .ChefDecalageY
            '.Arreter
            Select Case Persos(.NumeroChef).Action
                Case 300: .Attaquer_Perso Persos(NumeroChef).IndicePerso
                Case 301: .Attaquer_Maison Persos(NumeroChef).IndiceMaison
                Case 302:
                    .Attaquer_Chateau Persos(NumeroChef).IndiceChateau
                Case Else:
                    'If .IA_Berserk > 0 Then
                    '    IA_ComportementBerserk k 'Les berserkers attaquent tout de suite.
                    'Else
                        .Aller_A Persos(NumeroChef).DirectionX + .ChefDecalageX, _
                                 Persos(NumeroChef).DirectionY + .ChefDecalageY
                    'End If
            End Select
            'Petit effet visuel.
            If EffetVisuel >= 0 And Not Projectile Then
                'EffetsPersos(k).Activer_Effet EffetVisuel, .PositionX, .PositionY, Persos(k).Largeur, Persos(k).Hauteur, Parametres
                Activer_Effet EffetVisuel, .PositionX, .PositionY, Persos(k).Largeur, Persos(k).Hauteur, Parametres
            End If
            'L'invocation s'exprime.
            .Definir_Commentaires = Commentaires.Message(33, Persos(k))
            
            'NombreServiteurs = NombreServiteurs + 1
        End If
        End With
    Next i
    If Not Projectile Then
        'Persos(NumeroChef).NombreSoldats = NombreServiteurs ' Persos(NumeroChef).NombreSoldats + Quantite
        'Persos(NumeroChef).NombreSoldats = Persos(NumeroChef).NombreSoldats + Quantite
        Ordonner_Formation NumeroChef, , True
    End If
    Creer_Invocation = k
    'Exit Sub
'Erreur:
    'Debug.Print Err.Description
End Function

Private Function Quantite_Invocations_Aleatoire(ByVal Quantite As Single) As Integer
    'Renvoie le nombre d'invocation selon un tir aléatoire.
    'Ex : 2.75 renvoie 2 à (25%) ou 3 à (75%).
    Dim i As Integer
    i = Int(Quantite)
    Quantite_Invocations_Aleatoire = i + IIf(Rnd < Quantite - i, 1, 0)
End Function

Public Function Perso_Ajouter_Nouveau(Optional ByVal PlayerID As Long) As Long
    Dim i As Long
    'Ajoute un nouveau perso et renvoie son indice.
    i = UBound(Persos()) + 1
    ReDim Preserve Persos(i)
    Set Persos(i) = New ClsJeuPerso
    ReDim Preserve EffetsPersos(UBound(Persos()))
    Set EffetsPersos(i) = New ClsJeuEffet
    ReDim Preserve Maisons(UBound(Persos()))
    Set Maisons(i) = New ClsJeuBatiment
    Maisons(i).Numero = i
    ReDim Preserve EffetsMaisons(UBound(Persos()))
    Set EffetsMaisons(i) = New ClsJeuEffet
    Persos(i).Init i, Parametres
    Nombre_IAs = UBound(Persos()) + 1
    Persos(i).Joueur = PlayerID
    Perso_Ajouter_Nouveau = i
    If Jeu.Confrontation And PlayerID > 0 Then
        'Place le joueur de manière indépendant.
        Persos(i).PositionX = Rnd * (Monde.Largeur * Monde.LargeurCase - Persos(i).Largeur)
        Persos(i).PositionY = Rnd * (Monde.Hauteur * Monde.HauteurCase - Persos(i).Hauteur)
        'Ajoute un fief
        Persos(i).NumeroFief = Fief_Ajouter_Nouveau
        Persos(i).NumeroEquipe = i
        Fiefs(Persos(i).NumeroFief).NumeroEquipe = i
        Fiefs(Persos(i).NumeroFief).Epoque = Parametres.NombreEpoques - 1
    End If
End Function

Public Function Fief_Ajouter_Nouveau() As Long
    'Ajoute un nouveau fief et renvoie son numéro.
    Dim Nf As Long
    Nf = UBound(Fiefs()) + 1
    ReDim Preserve Fiefs(Nf)
    ReDim Preserve Chateaux(Nf)
    Set Fiefs(Nf) = New ClsJeuFief
    Fiefs(Nf).Init Nf, Parametres
    Set Chateaux(Nf) = New ClsJeuBatiment
    Chateaux(Nf).MaxVie = 1
    Fief_Ajouter_Nouveau = Nf
End Function

Public Function Perso_Supprimer_Nouveau(ByVal PlayerID As Long)
    Dim i As Long
    'Supprime tous les personnages du PlayerID indiqué en paramètres.
    For i = 0 To UBound(Persos())
        With Persos(i)
        If .Joueur = PlayerID Then
            .Joueur = 0
            .Invocation = True 'En le passant en invocation, on est sûr qu'il ne pourra pas ressusciter.
            .Vivant = False
        End If
        End With
    Next i
End Function

Private Sub Effets_Proches(ByVal Lanceur As ClsJeuPerso, _
                           ByVal ObjetAmi As Integer, _
                           ByVal ObjetEnnemi As Integer, _
                           ByVal ZoneEffet As Long, _
                           ByVal GainVie As Double, _
                           ByVal GainEnergie As Double, _
                           ByVal GainMagie As Double, _
                           ByVal GainMoral As Double, _
                           ByVal Resurrection As Long, _
                           ByVal Quantite As Long, _
                           ByVal EffetVisuel As Integer, _
                           ByVal NiveauCompetence As Integer, _
                           ByVal Resistance As Integer)
    Dim i As Long, j As Long
    Dim IndiceCible As Long
    Dim DistanceCible As Double
    Dim Nombre As Long
    Dim DejatAffecte As Collection 'Mémorise tous les joueurs qui ont déjà affecté au moins une fois.
    Dim TempDejatAffecte As Boolean
    Dim TempAmi As Boolean 'Vrai, l'effet affecte un allié. Faut, l'effet affecte un ennemi.

    If Quantite < 1 Then Quantite = 1
    Set DejatAffecte = New Collection
    For Nombre = 0 To Quantite - 1
        IndiceCible = -1
        DistanceCible = Sqr((Monde.Largeur * AffTerrain.Largeur) ^ 2 + _
                            (Monde.Hauteur * AffTerrain.Hauteur) ^ 2)
        For i = 0 To UBound(Persos())
            With Persos(i)
            If i <> Lanceur.Numero And _
               (IIf(ObjetAmi >= 0 And .Vivant, Lanceur.NumeroEquipe = .NumeroEquipe, False) Or IIf(ObjetEnnemi >= 0 And .Vivant, Lanceur.NumeroEquipe <> .NumeroEquipe, False) Or _
               (Lanceur.NumeroEquipe = .NumeroEquipe And (.Definir_Vie < .Definir_MaxVie And .Vivant And Resurrection = 0 Or _
                Not (.Vivant) And Resurrection > 0 And Not .Invocation))) Then
                If .Attaquable Then
                   'Persos(i).Definir_Vie < Persos(i).Definir_MaxVie Then
                    If DistanceCible > Sqr((Lanceur.PositionX - .PositionX) ^ 2 + _
                                           (Lanceur.PositionY - .PositionY) ^ 2) Then
                        'Résistance à l'effet.
                        If .Definir_Resistance(Resistance) > Parametres.ResistanceMinimum Then
                            If Rnd < .Definir_Resistance(Resistance) Then
                                'Verifie si la personne a déjà été affectée.
                                TempDejatAffecte = True
                                For j = 1 To DejatAffecte.Count
                                    If i = DejatAffecte(j) Then
                                        TempDejatAffecte = False
                                        Exit For
                                    End If
                                Next j
            '                    'Vérifie la résistance au type d'effet.
            '                    If Persos(i).Definir_Resistance(Parametres.Objet_AttaqueResistance(Parametres.Objet_AttaqueCauseEffet(i))) > Parametres.ResistanceMinimum Then 'Vérifie si immunité.
            '                        If Rnd < Persos(i).Definir_Resistance(Parametres.Objet_AttaqueResistance(Attaquant.Definir_AttaqueCauseEffet(i))) Then
            '
            '                        End If
            '                    End If
                                If TempDejatAffecte Then
                                    IndiceCible = i
                                    DistanceCible = Sqr((Lanceur.PositionX - .PositionX) ^ 2 + _
                                                    (Lanceur.PositionY - .PositionY) ^ 2)
                                    TempAmi = Lanceur.NumeroEquipe = Persos(i).NumeroEquipe
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            End With
        Next i
        'If IndiceCible >= 0 Then Debug.Print IndiceCible & " : " & Persos(IndiceCible).Nom
        If IndiceCible >= 0 And _
           DistanceCible < ZoneEffet Then
            If Not Persos(IndiceCible).Vivant And _
               Resurrection > 0 Then
                Ressusciter_Perso IndiceCible, EffetVisuel
            End If
            If GainVie + GainEnergie + GainMagie + GainMoral > 0 Then
                Persos(IndiceCible).RecuperationVie = Persos(IndiceCible).RecuperationVie + GainVie
                If GainVie > 0 Then
                    AffDegat.Ajouter_Degat Persos(IndiceCible).PositionX + Persos(IndiceCible).Largeur / 2, Persos(IndiceCible).PositionY, GainVie, , 1
                End If
                Persos(IndiceCible).RecuperationEnergie = Persos(IndiceCible).RecuperationEnergie + GainEnergie
                Persos(IndiceCible).RecuperationMagie = Persos(IndiceCible).RecuperationMagie + GainMagie
                Persos(IndiceCible).RecuperationMoral = Persos(IndiceCible).RecuperationMoral + GainMoral
            End If
            If TempAmi Then
                Effet_Sur_Perso Persos(IndiceCible), ObjetAmi, NiveauCompetence, 1, Lanceur.Numero
            Else
                Effet_Sur_Perso Persos(IndiceCible), ObjetEnnemi, NiveauCompetence, 2, Lanceur.Numero
            End If
            DejatAffecte.Add IndiceCible 'Ajoute cette cible à la liste de ceux qui ont été affectés.
            
            'Effet visuel.
            If EffetVisuel >= 0 Then
                'EffetsPersos(IndiceCible).Activer_Effet EffetVisuel, Persos(IndiceCible).PositionX, Persos(IndiceCible).PositionY, Persos(IndiceCible).Largeur, Persos(IndiceCible).Hauteur, Parametres
                Activer_Effet EffetVisuel, Persos(IndiceCible).PositionX, Persos(IndiceCible).PositionY, Persos(IndiceCible).Largeur, Persos(IndiceCible).Hauteur, Parametres
            End If
        Else
            'On n'a pas trouvé de cible.
            'Nombre = Quantite
            Exit For
        End If
    Next Nombre
End Sub

Private Sub Effet_Sur_Perso(ByRef Perso As ClsJeuPerso, ByVal Objet As Integer, ByVal NiveauCompetence As Integer, ByVal Indice As Integer, ByVal NoLanceur As Long)
    Dim k As Long
    Dim TempObjet As Integer
    If Objet > -1 Then
        Select Case Indice
        Case 0: TempObjet = Parametres.Objet_CreerObjet(Objet)
        Case 1: TempObjet = Parametres.Objet_CreerObjetAllie(Objet)
        Case 2: TempObjet = Parametres.Objet_CreerObjetEnnemi(Objet)
        End Select
        If TempObjet > -1 Then
            If Parametres.Objet_Duree(Objet) <> 0 Then
                Perso.EffetTemp_Ajouter TempObjet, IIf(Parametres.Objet_Duree(Objet) = -1, -1, Perso.Objet_Equipe_GainSup(NiveauCompetence, Parametres.Objet_Duree(Objet))), , NoLanceur
                'Messages(0).Ajouter_Message "Nouvel effet : " & TempObjet & ", " & IIf(Parametres.Objet_Duree(Objet) = -1, -1, Perso.Objet_Equipe_GainSup(NiveauCompetence, Parametres.Objet_Duree(Objet))) & " sur : " & Perso.Nom & " n°" & Perso.Numero
            Else
                If Parametres.Objet_Quantite(Objet) > 1 Then
                    For k = 1 To Parametres.Objet_Quantite(Objet)
                        Perso.Objet_Ajouter_Inventaire TempObjet, Parametres, Commentaires, True
                    Next k
                Else
                    Perso.Objet_Ajouter_Inventaire TempObjet, Parametres, Commentaires, True
                End If
            End If
        End If
    End If
End Sub

Public Sub Ressusciter_Perso(ByVal i As Long, ByVal EffetVisuel As Integer)
    Persos(i).Vivant = True
    Persos(i).Vie = Persos(i).Definir_MaxVie * 0.05
    Persos(i).RecuperationVie = Persos(i).Definir_MaxVie * (1 - 0.05)
    'Effet visuel.
    If EffetVisuel >= 0 Then
        'EffetsPersos(i).Activer_Effet EffetVisuel, Persos(i).PositionX, Persos(i).PositionY, Persos(i).Largeur, Persos(i).Hauteur, Parametres
        Activer_Effet EffetVisuel, Persos(i).PositionX, Persos(i).PositionY, Persos(i).Largeur, Persos(i).Hauteur, Parametres
    End If
End Sub
