Attribute VB_Name = "ModCombat"
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

Const FichierINI = "Combat"
Const SectionINI = "Combat"

Public CoefDegatPerso As Single
Private CoefDegatBatiment As Single
Private DureeImmobilisation As Integer
Const ImpactRalentissement = True

Private ChanceCoupCritique As Single
Private CoefDegatCoupCritique As Single
Private ChanceEchecCritique As Single
Private PerteEchecCritique As Single 'Nombre de points d'attaque perdus en cas d'échec critique.

Private MaxDegatBatiment As Single 'Nombre de dégât maximum que l'on inflige quand on tape sur un batiment.

'Quantité de ressources prises pendant un combat.
'Sur les personnages.
Private CoefPillageArgentPersoMort As Single
Private CoefPillageRessourcesPersoMort As Single
Private CoefPillageObjetsPersoMort As Single
Private CoefDestructionObjets As Single 'Chance qu'un objet soit détruit au lieu d'être pillé.
'Sur les bâtiments.
Private CoefPillageArgentMaison As Single
Private CoefPillageRessourcesMaison As Single
Private CoefPillageObjetsMaison As Single

Public Sub Init_Combat()
    FicIni.fichier = FicIni.chemin & FichierINI
    FicIni.Section = SectionINI
    
    CoefDegatPerso = Val(FicIni.Parametre("CoefDegatPerso"))
    CoefDegatBatiment = Val(FicIni.Parametre("CoefDegatBatiment"))
    DureeImmobilisation = Val(FicIni.Parametre("DureeImmobilisation"))
    
    ChanceCoupCritique = Val(FicIni.Parametre("ChanceCoupCritique"))
    CoefDegatCoupCritique = Val(FicIni.Parametre("CoefDegatCoupCritique"))
    ChanceEchecCritique = Val(FicIni.Parametre("ChanceEchecCritique"))
    PerteEchecCritique = Val(FicIni.Parametre("PerteEchecCritique"))
    
    MaxDegatBatiment = Val(FicIni.Parametre("MaxDegatBatiment"))
    
    CoefPillageArgentPersoMort = Val(FicIni.Parametre("CoefPillageArgentPersoMort"))
    CoefPillageRessourcesPersoMort = Val(FicIni.Parametre("CoefPillageRessourcesPersoMort"))
    CoefPillageObjetsPersoMort = Val(FicIni.Parametre("CoefPillageObjetsPersoMort"))
    CoefDestructionObjets = Val(FicIni.Parametre("CoefDestructionObjets"))
    
    CoefPillageArgentMaison = Val(FicIni.Parametre("CoefPillageArgentMaison"))
    CoefPillageRessourcesMaison = Val(FicIni.Parametre("CoefPillageRessourcesMaison"))
    CoefPillageObjetsMaison = Val(FicIni.Parametre("CoefPillageObjetsMaison"))
End Sub

Public Function Perso_Lancer_Projectile_Perso(ByRef Attaquant As ClsJeuPerso, _
                                              ByRef Defenseur As ClsJeuPerso) As Double
    Dim Temp As Double 'Mémorise les points d'attaque perdu l'hors du coup.
    
    Temp = (Defenseur.Vie + 1) / CoefDegatPerso + Defenseur.MaxDefense
    If Temp > Attaquant.Definir_Attaque Then Temp = Attaquant.Definir_Attaque
    If Temp >= 1 Then
        Attaquant.Definir_Attaque = Attaquant.Definir_Attaque - Temp
        Perso_Lancer_Projectile_Perso = Temp
    End If
End Function

Public Sub Perso_Attaquer_Perso(ByRef Attaquant As ClsJeuPerso, _
                                ByRef Defenseur As ClsJeuPerso, Optional ByVal Experience As Boolean = True)
    Dim Temp As Double 'Mémorise les points d'attaque perdu l'hors du coup.
    Dim Temp2 As Double 'Mémorise le nombre de points de vie perdu par la cible.
    Dim Attaque As Double 'Mémorise l'atttaque de l'attaquant.
    Dim Defense As Double 'Mémorise la défense du défenseur.
    Dim CoupDEtat As Boolean 'Si vrai, l'attaquant veut faire un ccoup d'état.
    If Attaquant.Projectile Then
        Attaque = Attaquant.Attaque
        If Attaque > Persos(Attaquant.NumeroChef).Attaque Then Attaque = Persos(Attaquant.NumeroChef).Attaque
    Else
        Attaque = Attaquant.Attaque
    End If
    Defense = Defenseur.Defense
    Dim i As Long
    If Rnd > ChanceEchecCritique Or (Attaquant.Kamikaze And Not Attaquant.Projectile) Then
        If Defenseur.Definir_Empereur Then 'On tente de s'attaquer à son empereur.
            If Attaquant.NumeroFief = Defenseur.NumeroFief Then
                CoupDEtat = True
            End If
        End If
        'Le personnage réussit son coup.
        'EffetsPersos(Defenseur.Numero).Activer_Effet Attaquant.Effet_Attaque, Defenseur.PositionX, Defenseur.PositionY, Defenseur.Largeur, Defenseur.Hauteur, Parametres
        Activer_Effet Attaquant.Effet_Attaque, Defenseur.PositionX, Defenseur.PositionY, Defenseur.Largeur, Defenseur.Hauteur, Parametres
        
        'Ralenti le personnage.
        Defenseur.DureeImmobilisation = DureeImmobilisation
        If ImpactRalentissement Then
            Defenseur.Longueur_Pas = 0
        End If
        
        Temp = Defenseur.Vie / CoefDegatPerso + Defenseur.Defense - 1
        If Attaque > Defense Then
            'Temp = (Defenseur.Vie + Defenseur.Defense)
            If Rnd > ChanceCoupCritique Then
                Temp2 = (Attaque - Defense) * CoefDegatPerso * Defenseur.Definir_Resistance(Attaquant.Definir_AttaqueResistance)
                If Temp2 > Defenseur.Vie Then Temp2 = Defenseur.Vie
                Defenseur.Definir_Vie = Defenseur.Vie _
                                        - Temp2
                AffDegat.Ajouter_Degat Defenseur.PositionX + Defenseur.Largeur / 2, Defenseur.PositionY, Temp2
            Else
                'Coup critique.
                If Attaquant.Projectile Then
                    'Temp2 = (Attaque * CoefDegatCoupCritique - Defense) _
                            * CoefDegatPerso _
                            * Persos(Attaquant.NumeroChef).Definir_Niveau_CompetenceSpeciales(2) / Parametres.PersosNiveauSpecialesDepart
                    Temp2 = (Attaque * (Persos(Attaquant.NumeroChef).Definir_Niveau_CompetenceSpeciales(2) / Parametres.PersosNiveauSpecialesDepart + CoefDegatCoupCritique) - Defense) _
                            * CoefDegatPerso _
                            * Defenseur.Definir_Resistance(Attaquant.Definir_AttaqueResistance)
                Else
                    Temp2 = (Attaque * (Attaquant.Definir_Niveau_CompetenceSpeciales(2) / Parametres.PersosNiveauSpecialesDepart + CoefDegatCoupCritique) - Defense) _
                            * CoefDegatPerso _
                            * Defenseur.Definir_Resistance(Attaquant.Definir_AttaqueResistance)
                End If
                If Temp2 > Defenseur.Vie Then Temp2 = Defenseur.Vie
                Defenseur.Definir_Vie = Defenseur.Vie _
                                        - Temp2
                AffDegat.Ajouter_Degat Defenseur.PositionX + Defenseur.Largeur / 2, Defenseur.PositionY, Temp2, True
                'Attaquant.Gestion_Niveau_Comp_Speciale 2 , Attaquant.Definir_Niveau_CompetenceSpeciales(2) * Temp, Parametres
                'Attaquant.Gestion_Niveau_Comp_Speciale 2, Attaquant.Attaque, Parametres
                'L'attaquant s'exclame de son coup critique.
                If Attaquant.Projectile Then
                    'Pour les projectiles, c'est le lanceur qui s'exclame.
                    Persos(Attaquant.NumeroChef).Definir_Commentaires = Commentaires.Message(2, Persos(Attaquant.NumeroChef))
                    If Experience Then
                        Persos(Attaquant.NumeroChef).Gestion_Niveau_Comp_Speciale 2, Temp2, Parametres
                    End If
                Else
                    Attaquant.Definir_Commentaires = Commentaires.Message(2, Attaquant)
                    If Experience Then
                        Attaquant.Gestion_Niveau_Comp_Speciale 2, Temp2, Parametres
                    End If
                End If
            End If
            
            'Gestion de l'expérience des résistances.
            Defenseur.Gestion_Niveau_Resistances Attaquant.Definir_AttaqueResistance, Temp2, Parametres
            
            'Debug.Print Attaquant.Attaque & "-" & Defenseur.Defense & "*" & CoefDegat & "=" & (Attaquant.Attaque + Defenseur.Defense) * CoefDegat
        Else
            'Si on ne traverse pas la défense, on n'inflige qu'1 point.
            Defenseur.Definir_Vie = Defenseur.Vie - 1
            AffDegat.Ajouter_Degat Defenseur.PositionX + Defenseur.Largeur / 2, Defenseur.PositionY, 1
            'Debug.Print 1
        End If
        
        Defenseur.Definir_Defense = Defense - Attaque
        If Not Defenseur.Vivant Then
            If Defenseur.Invocation Then
               If Not Attaquant.IA_Pacifique And _
                  Attaquant.Vitalite = 1 And _
                  Persos(Defenseur.NumeroChef).Vivant Then
                    If Attaquant.IA Then
                        'S'en prend à son invocateur.
                        Attaquant.Attaquer_Perso Defenseur.NumeroChef
                        Attaquant.Definir_Commentaires = Commentaires.Message(44, Attaquant)
                    End If
                    'L'invocateur envoie ses unités sur l'assaillant.
                    If Persos(Defenseur.NumeroChef).NombreSoldats > 0 Then
                        If Persos(Defenseur.NumeroChef).IA Then
                            Ordonner_Chargez Defenseur.NumeroChef, 0, Attaquant.Numero, Rnd > 0.5, Rnd > 0.5, Rnd > 0.8
                        End If
                    End If
                    
                'ElseIf Not Attaquant.Projectile Then
                '    Attaquant.Arreter
                End If
            'ElseIf Not Attaquant.Projectile Then
            '    Attaquant.Arreter
            End If
            'Defenseur.Arreter
            'Defenseur.Calculer_Apparence Maisons(Defenseur.Numero), Parametres
            'Ajout au bestiaire.
            If Attaquant.Invocation Then
                If Attaquant.Kamikaze Or Attaquant.Projectile Then  'Le bestiaire est compté pour le chef dans le cas d'un berserker.
                    Persos(Attaquant.NumeroChef).Definir_Bestiaire(Defenseur.Race) = Persos(Attaquant.NumeroChef).Definir_Bestiaire(Defenseur.Race) + 1
                    'Si l'attaquant est un mercenaire, il gagne une prime.
                    If Maisons(Attaquant.NumeroChef).Service = 2 And Maisons(Attaquant.NumeroChef).Construit Then
                        Maisons(Attaquant.NumeroChef).Service_Utiliser Persos(Attaquant.NumeroChef), Parametres, _
                            (Defenseur.Definir_MaxAttaque + Defenseur.Definir_MaxDefense + Defenseur.Definir_MaxVie) * _
                             Parametres.Service_Coefficient(2) * _
                            (Persos(Attaquant.NumeroChef).Definir_Niveau_CompetenceServices(2 - 1) + Parametres.PersosNiveauServicesDepart) / Parametres.PersosNiveauServicesDepart, True
                    End If
                Else
                    Attaquant.Definir_Bestiaire(Defenseur.Race) = Attaquant.Definir_Bestiaire(Defenseur.Race) + 1
                End If
                If Persos(Attaquant.NumeroChef).Vivant Then
                    'Les invocations offrent leur victoire à leur chef.
                    Perso_Tuer_Perso Persos(Attaquant.NumeroChef), Defenseur, CoupDEtat
                Else
                    Perso_Tuer_Perso Attaquant, Defenseur, CoupDEtat
                End If
                Persos(Attaquant.NumeroChef).FragsPersos = Persos(Attaquant.NumeroChef).FragsPersos + 1
            Else
                Perso_Tuer_Perso Attaquant, Defenseur, CoupDEtat
                Attaquant.Definir_Bestiaire(Defenseur.Race) = Attaquant.Definir_Bestiaire(Defenseur.Race) + 1
                Attaquant.FragsPersos = Attaquant.FragsPersos + 1
                'Si l'attaquant est un mercenaire, il gagne une prime.
                If Maisons(Attaquant.Numero).Service = 2 And Maisons(Attaquant.Numero).Construit Then
                    'Maisons(Attaquant.Numero).Service_Utiliser Attaquant, Parametres, _
                        (Defenseur.Definir_MaxAttaque + Defenseur.Definir_MaxDefense + Defenseur.Definir_MaxVie) * _
                         Parametres.Service_Coefficient(2) * _
                        (Attaquant.Definir_Niveau_CompetenceServices(2 - 1) + Parametres.PersosNiveauServicesDepart) / Parametres.PersosNiveauServicesDepart, True
                    Maisons(Attaquant.Numero).Service_Utiliser Attaquant, Parametres, _
                        (Defenseur.Definir_MaxAttaque + Defenseur.Definir_MaxDefense + Defenseur.Definir_MaxVie) * _
                         Parametres.Service_Coefficient(2) * _
                        ((Attaquant.Definir_Niveau_CompetenceServices(2 - 1) - Parametres.PersosNiveauServicesDepart + 1) / Parametres.PersosNiveauServicesDepart / 10 + 1), True
                End If
            End If
        Else
            'Se prend les effets temporaires de l'attaque.
            If Attaquant.Definir_NombreAttaqueCauseEffet > 0 Then
                For i = 1 To Attaquant.Definir_NombreAttaqueCauseEffet
                    If Defenseur.Definir_Resistance(Parametres.Objet_AttaqueResistance(Attaquant.Definir_AttaqueCauseEffet(i))) > Parametres.ResistanceMinimum Then
                        If Rnd < Defenseur.Definir_Resistance(Parametres.Objet_AttaqueResistance(Attaquant.Definir_AttaqueCauseEffet(i))) Then
                            Defenseur.EffetTemp_Ajouter Attaquant.Definir_AttaqueCauseEffet(i), _
                                                        IIf(Parametres.Objet_Duree(Attaquant.Definir_AttaqueCauseEffet(i)) = -1, -1, _
                                                        Attaquant.Objet_Equipe_GainSup(Attaquant.Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(Attaquant.Definir_AttaqueCauseEffet(i))), Parametres.Objet_Duree(Attaquant.Definir_AttaqueCauseEffet(i)))), _
                                                        IIf(Parametres.Objet_Duree(Attaquant.Definir_AttaqueCauseEffet(i)) = -1, -1, _
                                                        Attaquant.Objet_Equipe_GainSup(Attaquant.Definir_Niveau_CompetenceObjets(Parametres.Objet_Competence(Attaquant.Definir_AttaqueCauseEffet(i))), Parametres.Objet_Duree(Attaquant.Definir_AttaqueCauseEffet(i)))), _
                                                        IIf(Attaquant.Projectile, Attaquant.NumeroChef, Attaquant.Numero)
                            AffDegat.Ajouter_Texte Defenseur.PositionX + Defenseur.Largeur / 2, Defenseur.PositionY - 9 * (i - 1), Parametres.Objet_NomEffet(Attaquant.Definir_AttaqueCauseEffet(i), Defenseur.Feminin)
                        End If
                    End If
                Next i
            End If
            
            If Defenseur.IA Then
                If Defenseur.Action = 300 Then
                    'Ne s'en prend pas à l'attaquant si celui-ci est plus loin que sa cible actuelle.
                    If (Attaquant.Kamikaze And Attaquant.Invocation) Or Attaquant.Projectile Then
                        With Persos(Attaquant.NumeroChef)
                        If Abs(.PositionX - Defenseur.PositionX) + Abs(.PositionY - Defenseur.PositionY) _
                           < Abs(Persos(Defenseur.IndicePerso).PositionX - Defenseur.PositionX) + Abs(Persos(Defenseur.IndicePerso).PositionY - Defenseur.PositionY) Then
                            Defenseur.Attaquer_Perso .Numero
                        End If
                        End With
                    Else
                        With Attaquant
                        If Abs(.PositionX - Defenseur.PositionX) + Abs(.PositionY - Defenseur.PositionY) _
                           < Abs(Persos(Defenseur.IndicePerso).PositionX - Defenseur.PositionX) + Abs(Persos(Defenseur.IndicePerso).PositionY - Defenseur.PositionY) Then
                            Defenseur.Attaquer_Perso .Numero
                        End If
                        End With
                    End If
                Else
                    If (Attaquant.Kamikaze And Attaquant.Invocation) Or Attaquant.Projectile Then
                        Defenseur.Attaquer_Perso Attaquant.NumeroChef
                    Else
                        Defenseur.Attaquer_Perso Attaquant.Numero
                    End If
                End If
                If Defenseur.Vie < Defenseur.Definir_MaxVie / 2 And _
                   Defenseur.Vie < Attaquant.Vie And _
                   Not Defenseur.Invocation And _
                   Rnd > 0.2 Then
                    Appeler_De_lAide Defenseur.Numero
                    'Le défenseur fuit.
                    If Defenseur.Longueur_Pas >= Attaquant.Longueur_Pas And _
                       Defenseur.Vie < Attaquant.Vie And _
                       Rnd > 0.9 And _
                       Not Defenseur.IA_Temeraire Then
                        If Maisons(Defenseur.Numero).Construit Then
                            Defenseur.Rentrer_Maison Maisons(Defenseur.Numero)
                        ElseIf Chateaux(Defenseur.NumeroFief).Visible Then
                            Defenseur.Rentrer_Chateau Chateaux(Defenseur.NumeroFief)
                        Else
                            IA_Deplacement_Aleatoire Defenseur.Numero
                        End If
                        'Appeler_De_lAide Defenseur.Numero
                        Attaquant.Definir_Commentaires = Commentaires.Message(13, Attaquant)
                    End If
                End If
            ElseIf Defenseur.Inactif Or Defenseur.Action = 2 Then
                'Le personnage contre attaque automatiquement.
                If Attaquant.Kamikaze And Attaquant.Invocation Then
                    Defenseur.Attaquer_Perso Attaquant.NumeroChef
                Else
                    Defenseur.Attaquer_Perso Attaquant.Numero
                End If
            End If
        End If
    Else
        If Attaquant.Projectile Then
            'Pour les projectiles, c'est le lanceur qui s'exclame.
            Persos(Attaquant.NumeroChef).Definir_Commentaires = Commentaires.Message(3, Persos(Attaquant.NumeroChef))
        Else
            Attaquant.Definir_Commentaires = Commentaires.Message(3, Attaquant)
        End If
        Temp = PerteEchecCritique
    End If
    If Attaquant.Projectile Then
        If Experience Then
            Persos(Attaquant.NumeroChef).Definir_Attaque = Persos(Attaquant.NumeroChef).Definir_Attaque - Temp
        Else
            Persos(Attaquant.NumeroChef).Attaque = Persos(Attaquant.NumeroChef).Definir_Attaque - Temp
        End If
    ElseIf Experience Then
        Attaquant.Definir_Attaque = Attaquant.Definir_Attaque - Temp
    Else
        Attaquant.Attaque = Attaquant.Definir_Attaque - Temp
    End If
End Sub

Public Sub Perso_Tuer_Perso(ByRef Attaquant As ClsJeuPerso, _
                            ByRef Defenseur As ClsJeuPerso, Optional ByVal CoupDEtat As Boolean)
    Dim i As Integer
    Dim Pille As Boolean 'Vrai si l'attaquant récupère les ressources au défunt.
    Dim IndiceTresor As Integer 'Mémorise l'indice du trésor déposé.
    'Une personne qui était en danger le remercie.
    If Persos(Defenseur.IndicePerso).NumeroEquipe = Attaquant.NumeroEquipe And _
       Defenseur.IndicePerso <> Attaquant.Numero And _
       Defenseur.Action = 300 Then
        Persos(Defenseur.IndicePerso).Definir_Commentaires = ""
        Persos(Defenseur.IndicePerso).Definir_Commentaires = Commentaires.Message(14, Persos(Defenseur.IndicePerso))
    End If
    
    Pille = Defenseur.PositionX + Defenseur.Largeur >= Attaquant.PositionX And _
            Defenseur.PositionY + Defenseur.Hauteur >= Attaquant.PositionY And _
            Defenseur.PositionX <= Attaquant.PositionX + Attaquant.Largeur And _
            Defenseur.PositionY <= Attaquant.PositionY + Attaquant.Hauteur
    If Not Pille Then IndiceTresor = Tresor_Creer(Defenseur.PositionX + Defenseur.Largeur / 2 - AffTresor.Largeur / 2, _
                                                  Defenseur.PositionY + Defenseur.Hauteur / 2 - AffTresor.Hauteur / 2)
    
    'L'attaquant prend l'argent du défunt.
    If Pille Then
        Attaquant.Definir_Argent = Attaquant.Definir_Argent _
                                   + Defenseur.Definir_Argent * CoefPillageArgentPersoMort
    Else
        Tresors(IndiceTresor).Definir_Argent = Tresors(IndiceTresor).Definir_Argent _
                                   + Defenseur.Definir_Argent * CoefPillageArgentPersoMort
    End If
    Defenseur.Definir_Argent = Defenseur.Definir_Argent _
                               * (1 - CoefPillageArgentPersoMort)
    
    
    'L'attaquant prend les ressources du défunt.
    For i = 0 To Parametres.NombreRessources - 1
        If Pille Then
            Attaquant.Definir_ressources(i) = Attaquant.Definir_ressources(i) _
                                   + Defenseur.Definir_ressources(i) * CoefPillageRessourcesPersoMort
        Else
            Tresors(IndiceTresor).Definir_ressources(i) = Tresors(IndiceTresor).Definir_ressources(i) _
                                   + Defenseur.Definir_ressources(i) * CoefPillageRessourcesPersoMort
        End If
        Defenseur.Definir_ressources(i) = Defenseur.Definir_ressources(i) _
                               * (1 - CoefPillageRessourcesPersoMort)
    Next i
    'L'attaquant prend des objets au défunt.
    For i = 0 To Parametres.PersosNombreObjetsEquipes - 1
        If Defenseur.Objet_Equipes_Actif(i) Then
            If Rnd < CoefPillageObjetsPersoMort And _
               Not Defenseur.Objet_Equipes_Indestructible(i) Then
                 'L'objet est volé selon un facteur de chance.
                 If Rnd > CoefDestructionObjets Then
                     If Pille Then
                         Attaquant.Objet_Ajouter_Inventaire Defenseur.Objet_Equipes_Type(i), Parametres, Commentaires, True
                     Else
                         Tresors(IndiceTresor).Ajouter_Objet Defenseur.Objet_Equipes_Type(i), Parametres
                     End If
                 End If
                 'L'objet est détruit.
                 Defenseur.Objet_Equipes_Detruire i
             End If
        End If
    Next i
    For i = 0 To Parametres.PersosNombreObjetsInventaire - 1
        If Defenseur.Objet_Inventaire_Actif(i) Then
            If Rnd < CoefPillageObjetsPersoMort And _
                Not Defenseur.Objet_Inventaire_Indestructible(i) Then
                'L'objet est volé selon un facteur de chance.
                If Rnd > CoefDestructionObjets Then
                    If Pille Then
                        Attaquant.Objet_Ajouter_Inventaire Defenseur.Objet_Inventaire_Type(i), Parametres, Commentaires, True
                    Else
                        Tresors(IndiceTresor).Ajouter_Objet Defenseur.Objet_Inventaire_Type(i), Parametres
                    End If
                End If
                'L'objet est détruit.
                Defenseur.Objet_Inventaire_Detruire i
            End If
        End If
    Next i
    
    'Si c'est le roi, on prend sa couronne.
    If CoupDEtat Then
        Fiefs(Attaquant.NumeroFief).Empereur_Changer Attaquant.Numero
        'Attaquant.Definir_Empereur = True 'Prend la place de l'empereur.
        Attaquant.Attaquer_Allies False 'N'est plus hostile aux siens.
    End If
    
    'Les trésors vides n'apparaissent pas.
    If Not Pille Then
        If Tresors(IndiceTresor).Vide Then
            Tresor_Effacer IndiceTresor
        Else
            Tresors_Verifier IndiceTresor
        End If
    End If
    
    'Place une tombe.
    If Parametres.Race_Tombe(Defenseur.Race) >= 0 Then
        Decor_Creer Parametres.Race_Tombe(Defenseur.Race), Defenseur.PositionX, Defenseur.PositionY, "&" & Defenseur.Nom & " (" & Parametres.EtiquetteNiveau & " " & Int(Defenseur.Niveau) & ")" & IIf(Defenseur.Feminin, Parametres.EtiquetteTombeFeminin, Parametres.EtiquetteTombeMasculin) & Attaquant.Nom
    End If
End Sub

'Public Function Perso_LancerProjectile_Batiment(ByRef Attaquant As ClsJeuPerso, _
'                                                ByRef Batiment As ClsJeuBatiment, _
'                                                ByRef Defenseur As ClsJeuPerso, _
'                                                Optional ByVal NoChateau As Long = -1) As Double
'    Dim i As Long
'    Dim Temp As Double 'Mémorise les points d'attaque perdu l'hors du coup.
'
'    Temp = Batiment.Vie
'    If Temp > MaxDegatBatiment And Not Attaquant.Kamikaze Then
'        Temp = MaxDegatBatiment
'    End If
'    If Temp > Attaquant.Definir_Attaque Then
'        Temp = Attaquant.Definir_Attaque
'    ElseIf Temp < 1 Then
'        Temp = 1
'    End If
'    If Temp >= 1 Then
'        Attaquant.Attaque = Attaquant.Attaque - Temp
'        Perso_LancerProjectile_Batiment = Temp
'    End If
'End Function

Public Sub Perso_Attaquer_Batiment(ByRef Attaquant As ClsJeuPerso, _
                                   ByRef Batiment As ClsJeuBatiment, _
                                   ByRef Defenseur As ClsJeuPerso, _
                                   Optional ByVal NoChateau As Long = -1)
    Dim i As Long
    Dim Temp As Double, Temp2 As Double
    Dim Attaque As Double 'Mémorise l'attaque du défenseur.
    Dim Defense As Double 'Mémorise la défense du défenseur.
    
    If Attaquant.Projectile Then
        Attaque = Persos(Attaquant.NumeroChef).Attaque
    Else
        Attaque = Attaquant.Attaque
    End If
    Defense = Attaquant.Defense
    
    Temp = Batiment.Vie
    If Temp > MaxDegatBatiment And (Not Attaquant.Kamikaze Or Attaquant.Projectile) Then
        Temp = MaxDegatBatiment
    End If
    If Temp > Attaque Then
        Temp = Attaque
    ElseIf Temp < 1 Then
        Temp = 1
    End If
    
    'If Attaquant.Projectile Then
    '    Temp2 = Temp _
                * CoefDegatBatiment _
                * Persos(Attaquant.NumeroChef).Definir_Niveau_CompetenceSpeciales(1) / Parametres.PersosNiveauSpecialesDepart
    'Else
        Temp2 = Temp _
                * CoefDegatBatiment _
                * Attaquant.Definir_Niveau_CompetenceSpeciales(1) / Parametres.PersosNiveauSpecialesDepart
    'End If
    Batiment.Definir_Vie = Batiment.Vie _
                           - Temp2
    'AffDegat.Ajouter_Degat Batiment.PositionX + Batiment.Largeur / 2, Batiment.PositionY, Temp2
    AffDegat.Ajouter_Degat Batiment.PositionX + Batiment.Largeur / 4 + Rnd * Batiment.Largeur / 2, Batiment.PositionY, Temp2
    
    'Gain d'expérience pour la compétence détruire batiment.
    If Attaquant.Projectile Then
        'Pas d'expérience dans les caractéristiques de combat.
        Persos(Attaquant.NumeroChef).Attaque = Persos(Attaquant.NumeroChef).Attaque - Temp
        Persos(Attaquant.NumeroChef).Gestion_Niveau_Comp_Speciale 1, Temp, Parametres
    Else
        'Pas d'expérience dans les caractéristiques de combat.
        Attaquant.Attaque = Attaquant.Attaque - Temp
        Attaquant.Gestion_Niveau_Comp_Speciale 1, Temp, Parametres
    End If
    
    If Not Batiment.Visible Then
        Attaquant.Arreter
        Batiment.EnConstruction = False
        'Commentaires de ceux qui on perdu le batiment.
        If NoChateau >= 0 Then
            'EffetsMaisons(Attaquant.Numero).Activer_Effet Parametres.EffetChateauDetruit, Batiment.PositionX, Batiment.PositionY, Batiment.Largeur, Batiment.Hauteur, Parametres, True
            Activer_Effet Parametres.EffetChateauDetruit, Batiment.PositionX, Batiment.PositionY, Batiment.Largeur, Batiment.Hauteur, Parametres
            For i = 0 To UBound(Persos())
                With Persos(i)
                If .Vivant And _
                   .NumeroFief = NoChateau And _
                   Rnd > 0.6 Then
                    'Le chateau est détruit.
                    .Definir_Commentaires = ""
                    .Definir_Commentaires = Commentaires.Message(12, Persos(i))
                End If
                End With
            Next i
            If Attaquant.Invocation Then
                Persos(Attaquant.NumeroChef).FragsChateaux = Persos(Attaquant.NumeroChef).FragsChateaux + 1
            Else
                Attaquant.FragsChateaux = Attaquant.FragsChateaux + 1
            End If
        Else
            'La maison est détruite.
            'Pille le batiment.
            If Attaquant.Invocation Then
                'Les invocations offrent leur victoire à leur chef.
                Perso_Piller_Batiment Persos(Attaquant.NumeroChef), Batiment
                Persos(Attaquant.NumeroChef).FragsBatiments = Persos(Attaquant.NumeroChef).FragsBatiments + 1
            Else
                Perso_Piller_Batiment Attaquant, Batiment
                Attaquant.FragsBatiments = Attaquant.FragsBatiments + 1
            End If
            
            'Defenseur.Calculer_Apparence Batiment, Parametres
            Defenseur.Definir_Commentaires = ""
            Defenseur.Definir_Commentaires = Commentaires.Message(5, Defenseur)
            'EffetsMaisons(Batiment.Numero).Activer_Effet Parametres.EffetMaisonDetruite, Batiment.PositionX, Batiment.PositionY, Batiment.Largeur, Batiment.Hauteur, Parametres, True
            Activer_Effet Parametres.EffetMaisonDetruite, Batiment.PositionX, Batiment.PositionY, Batiment.Largeur, Batiment.Hauteur, Parametres
            'L'attaquant s'en prend maintenant au propriétaire de la maison, s'il était dedans.
            If Attaquant.IA And Defenseur.DansSaMaison And Rnd < 0.9 Then
                Attaquant.Attaquer_Perso Batiment.Numero
            End If
        End If
    Else
        If NoChateau >= 0 Then
            'Effet visuel.
            'EffetsMaisons(Attaquant.Numero).Activer_Effet Attaquant.Effet_Attaque, Batiment.PositionX, Batiment.PositionY, Batiment.Largeur, Batiment.Hauteur, Parametres, True
            'Activer_Effet Attaquant.Effet_Attaque, Batiment.PositionX, Batiment.PositionY, Batiment.Largeur, Batiment.Hauteur, Parametres
            Activer_Effet Attaquant.Effet_Attaque, Batiment.PositionX + Rnd * Batiment.Largeur / 2 - Batiment.Largeur / 4, Batiment.PositionY + Rnd * Batiment.Hauteur / 2 - Batiment.Hauteur / 4, Batiment.Largeur, Batiment.Hauteur, Parametres
            For i = 0 To UBound(Persos())
                With Persos(i)
                'Persos(i).Action < 300 And
                'Tous les IA qui protègent le chateau contre attaque.
                If .Vivant Then
                If .NumeroFief = NoChateau Then
                If Not .Invocation And Not .IA_Egoiste Then
                    If .IA And _
                       Rnd > 0.92 Then
                        If Attaquant.Kamikaze And Attaquant.Invocation Then
                            .Attaquer_Perso Attaquant.NumeroChef
                        Else
                            .Attaquer_Perso Attaquant.Numero
                        End If
                    End If
                    If Rnd > 0.9 Then
                        .Definir_Commentaires = Commentaires.Message(11, Persos(i))
                    End If
                End If
                End If
                End If
                End With
            Next i
            'L'attaquant appelle des renforts.
            'If Attaquant.IA And _
               Rnd > 0.9 Then
            If Rnd > 0.9 Then
                Appeler_De_lAide Attaquant.Numero
            End If
        Else
            'Effet visuel.
            'EffetsMaisons(Defenseur.Numero).Activer_Effet Attaquant.Effet_Attaque, Batiment.PositionX, Batiment.PositionY, Batiment.Largeur, Batiment.Hauteur, Parametres, True
            Activer_Effet Attaquant.Effet_Attaque, Batiment.PositionX + Rnd * Batiment.Largeur - Batiment.Largeur / 2, Batiment.PositionY + Rnd * Batiment.Hauteur - Batiment.Hauteur / 2, Batiment.Largeur, Batiment.Hauteur, Parametres, True
            
            If Defenseur.Vivant And _
               Defenseur.IA And Not Defenseur.Invocation Then
                If Defenseur.Action <> 300 And _
                   Rnd > 0.75 Then
                    Defenseur.Attaquer_Perso Attaquant.Numero
                'ElseIf Batiment.Definir_Vie < Batiment.MaxVie / 2 And _
                   Rnd > 0.65 Then
                '    Appeler_De_lAide Defenseur.Numero
                End If
            End If
            If Batiment.Definir_Vie < Batiment.MaxVie / 2 And _
               Rnd > 0.65 Then
                Appeler_De_lAide Defenseur.Numero
            End If
            Defenseur.Definir_Commentaires = Commentaires.Message(4, Defenseur)
        End If
    End If
End Sub

Public Sub Perso_Attaquer_Surface(ByRef Attaquant As ClsJeuPerso, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal Rayon As Long, _
                                  ByVal Joueur As Boolean, _
                                  Optional ByVal NoChateau As Long = -1)
    'Détruit tout dans la zone de centre X,Y et de rayon entré en paramètre.
    Dim i As Long
    Dim TempAttaqueInit As Double 'Garde en mémoire l'attaque initiale réinitialisé entre chaque coup.
    Dim TempAttaqueMin As Double 'Garde en mémoire le score d'attaque le plus bas. Qui sera le score final.
    TempAttaqueInit = Attaquant.Attaque
    'Détruit les personnages.
    For i = 0 To UBound(Persos())
        If Persos(i).NumeroEquipe <> Attaquant.NumeroEquipe And _
           Persos(i).Attaquable And _
           (Rayon > Sqr((Persos(i).PositionX - Attaquant.PositionX) ^ 2 + _
                       (Persos(i).PositionY - Attaquant.PositionY) ^ 2) Or _
           i = Attaquant.IndicePerso) Then
            Perso_Attaquer_Perso Attaquant, Persos(i), False
            If TempAttaqueMin > Attaquant.Attaque Then
                TempAttaqueMin = Attaquant.Attaque
            End If
            Attaquant.Attaque = TempAttaqueInit
        End If
    Next i
    'Détruit les maisons.
    For i = 0 To UBound(Maisons())
        If Persos(i).NumeroEquipe <> Attaquant.NumeroEquipe And _
           Maisons(i).Visible And _
           (Rayon > Sqr(((Maisons(i).PositionX + Maisons(i).Largeur / 2) - Attaquant.PositionX) ^ 2 + _
                       ((Maisons(i).PositionY + Maisons(i).Largeur / 2) - Attaquant.PositionY) ^ 2) Or _
           i = Attaquant.IndiceMaison) Then
            Perso_Attaquer_Batiment Attaquant, Maisons(i), Persos(i)
            If TempAttaqueMin > Attaquant.Attaque Then
                TempAttaqueMin = Attaquant.Attaque
            End If
            Attaquant.Attaque = TempAttaqueInit
        End If
    Next i
    'Détruit les chateaux.
    For i = 0 To UBound(Chateaux())
        If Fiefs(i).NumeroEquipe <> Attaquant.NumeroEquipe And _
           Chateaux(i).Visible And _
           (Rayon > Sqr((Chateaux(i).PositionX - Attaquant.PositionX) ^ 2 + _
                       (Chateaux(i).PositionY - Attaquant.PositionY) ^ 2) Or _
           i = Attaquant.IndiceChateau) Then
            Perso_Attaquer_Batiment Attaquant, Chateaux(i), Persos(Attaquant.IndicePerso), NoChateau
            If TempAttaqueMin > Attaquant.Attaque Then
                TempAttaqueMin = Attaquant.Attaque
            End If
            Attaquant.Attaque = TempAttaqueInit
        End If
    Next i
    'On réduit le score d'attaque aux dégâts les plus importants.
    Attaquant.Definir_Attaque = TempAttaqueMin
End Sub

Public Sub Perso_Piller_Batiment(ByRef Attaquant As ClsJeuPerso, _
                                 ByRef Batiment As ClsJeuBatiment)
    Dim i As Integer, j As Integer
    'Dim Pille As Boolean
    Dim IndiceTresor  As Long
    
    'Pille = Batiment.PositionX + Batiment.Largeur >= Attaquant.PositionX And _
            Batiment.PositionY + Batiment.Hauteur >= Attaquant.PositionY And _
            Batiment.PositionX <= Attaquant.PositionX + attaquant.Largeur And _
            Batiment.PositionY <= Attaquant.PositionY + attaquant.Hauteur
    'If Not Pille Then IndiceTresor = Tresor_Creer(Batiment.PositionX + Batiment.Largeur / 2 - AffTresor.Largeur / 2, Batiment.PositionY + Batiment.Largeur / 2 - AffTresor.Largeur / 2)
    IndiceTresor = Tresor_Creer(Batiment.PositionX + Batiment.Largeur / 2 - AffTresor.Largeur / 2, Batiment.PositionY + Batiment.Largeur / 2 - AffTresor.Largeur / 2)
    
    'L'attaquant pille de l'argent dans le batiment.
        'If Pille Then
        '    Attaquant.Definir_Argent = Attaquant.Definir_Argent + Batiment.Argent * CoefPillageArgentMaison
        'Else
        Tresors(IndiceTresor).Definir_Argent = Tresors(IndiceTresor).Definir_Argent + Batiment.Argent * CoefPillageArgentMaison
        'End If
        
    'L'attaquant prend des ressources dans le batiment.
    If Batiment.Magasin And _
       Parametres.Batiment_RessourceFabrique(Batiment.TypeBatiment) > 0 And _
       Batiment.Definir_Stock > 0 Then
        'If Pille Then
        '    Attaquant.Definir_Ressources(Parametres.Batiment_RessourceFabrique(Batiment.TypeBatiment)) = Attaquant.Definir_Ressources(Parametres.Batiment_RessourceFabrique(Batiment.TypeBatiment)) _
                                   + Batiment.Definir_Stock * CoefPillageRessourcesMaison
        'Else
        Tresors(IndiceTresor).Definir_ressources(Parametres.Batiment_RessourceFabrique(Batiment.TypeBatiment)) = Tresors(IndiceTresor).Definir_ressources(Parametres.Batiment_RessourceFabrique(Batiment.TypeBatiment)) _
                               + Batiment.Definir_Stock * CoefPillageRessourcesMaison
        'End If
    End If
    'L'attaquant prend des objets dans le bâtiment.
    If Batiment.Fabrique And _
       Batiment.Definir_Stock > 0 Then
        For i = 0 To Batiment.Definir_Nombre_Stock_Objet - 1
            If Batiment.Stock_Objets(i) > 0 Then
                For j = 0 To Batiment.Stock_Objets(i) - 1
                    If CoefPillageObjetsMaison > Rnd Then
                        'If Pille Then
                        '    Attaquant.Objet_Ramasser_Maison i, Batiment, Parametres, Commentaires
                        'Else
                        Tresors(IndiceTresor).Ajouter_Objet Parametres.Batiment_ObjetFabrique(Batiment.TypeBatiment, i), Parametres
                        'End If
                    End If
                Next j
            End If
        Next i
    End If
    
    'Les trésors vides n'apparaissent pas.
    'If Not Pille Then
        If Tresors(IndiceTresor).Vide Then
            Tresor_Effacer IndiceTresor
        Else
            Tresors_Verifier IndiceTresor
        End If
    'End If
End Sub

Public Sub Perso_Demolir_Batiment(ByRef Attaquant As ClsJeuPerso, _
                                  ByRef Batiment As ClsJeuBatiment)
    Dim i As Integer, j As Integer
    Dim Pille As Boolean
    Dim IndiceTresor  As Long
    
    Pille = Batiment.PositionX + Batiment.Largeur >= Attaquant.PositionX And _
            Batiment.PositionY + Batiment.Hauteur >= Attaquant.PositionY And _
            Batiment.PositionX <= Attaquant.PositionX + Attaquant.Largeur And _
            Batiment.PositionY <= Attaquant.PositionY + Attaquant.Hauteur
    If Not Pille Then
        IndiceTresor = Tresor_Creer(Batiment.PositionX + Batiment.Largeur / 2 - AffTresor.Largeur / 2, Batiment.PositionY + Batiment.Largeur / 2 - AffTresor.Largeur / 2)
    End If
    'IndiceTresor = Tresor_Creer(Batiment.PositionX + Batiment.Largeur / 2 - AffTresor.Largeur / 2, Batiment.PositionY + Batiment.Largeur / 2 - AffTresor.Largeur / 2)
    
    'L'attaquant pille de l'argent dans le batiment.
    If Pille Then
        Attaquant.Definir_Argent = Attaquant.Definir_Argent + Batiment.Argent
    Else
        Tresors(IndiceTresor).Definir_Argent = Tresors(IndiceTresor).Definir_Argent + Batiment.Argent
    End If
        
    'L'attaquant prend des ressources dans le batiment.
    If Batiment.Magasin And _
       Parametres.Batiment_RessourceFabrique(Batiment.TypeBatiment) > 0 And _
       Batiment.Definir_Stock > 0 Then
        If Pille Then
            Attaquant.Definir_ressources(Parametres.Batiment_RessourceFabrique(Batiment.TypeBatiment)) = Attaquant.Definir_ressources(Parametres.Batiment_RessourceFabrique(Batiment.TypeBatiment)) _
                                   + Batiment.Definir_Stock
        Else
            Tresors(IndiceTresor).Definir_ressources(Parametres.Batiment_RessourceFabrique(Batiment.TypeBatiment)) = Tresors(IndiceTresor).Definir_ressources(Parametres.Batiment_RessourceFabrique(Batiment.TypeBatiment)) _
                                   + Batiment.Definir_Stock
        End If
    End If
    'L'attaquant prend des objets dans le bâtiment.
    If Batiment.Fabrique And _
       Batiment.Definir_Stock > 0 Then
        For i = 0 To Batiment.Definir_Nombre_Stock_Objet
            If Batiment.Stock_Objets(i) > 0 Then
                For j = 0 To Batiment.Stock_Objets(i) - 1
                    If Pille Then
                        Attaquant.Objet_Ramasser_Maison i, Batiment, Parametres, Commentaires
                    Else
                        Tresors(IndiceTresor).Ajouter_Objet Parametres.Batiment_ObjetFabrique(Batiment.TypeBatiment, i), Parametres
                    End If
                Next j
            End If
        Next i
    End If
    'L'attaquant prend les objets de l'inventaire du bâtiment.
    For i = 0 To Batiment.Definir_Nombre_Objet_Inventaire - 1
        If Batiment.Definir_ObjetsInventaire(i) > -1 Then
            If Pille Then
                Attaquant.Objet_Ajouter_Inventaire Batiment.Definir_ObjetsInventaire(i), Parametres, Commentaires, True
            Else
                Tresors(IndiceTresor).Ajouter_Objet Batiment.Definir_ObjetsInventaire(i), Parametres
            End If
        End If
    Next i
    
    'Les trésors vides n'apparaissent pas.
    If Not Pille Then
        If Tresors(IndiceTresor).Vide Then
            Tresor_Effacer IndiceTresor
        End If
    End If
    
    'Toutes les IAs qui s'en prenaient au bâtiment attaquent maintenant le propriétaire.
'    For i = 0 To UBound(Persos())
'        If Persos(i).IA Then
'            IA_Poursuivre_Fuyard i
'        End If
'    Next i
End Sub

Public Sub Perso_Pietinement(ByRef Perso As ClsJeuPerso)
    'Le personnage écrase tout ce qui se trouve sous lui.
    Dim i As Long, Pne As Long
    Dim X As Long, Y As Long
    Dim Largeur As Integer, Hauteur As Integer
    X = Perso.PositionX: Y = Perso.PositionY
    Largeur = Perso.Largeur: Hauteur = Perso.Hauteur
    Pne = Perso.NumeroEquipe
    
    'Ecrase les personnages ennemis.
    For i = 0 To UBound(Persos())
        If Perso.Attaque < Perso.Pietinement Then
            Perso.Definir_Attaque = Perso.Pietinement
        End If
        With Persos(i)
        If Persos(i).NumeroEquipe <> Pne Then
            If Persos(i).Attaquable Then
                If Est_Superpose(X, _
                                 Y, _
                                 .PositionX, _
                                 .PositionY, _
                                 Largeur, _
                                 Hauteur, _
                                 .Largeur, _
                                 .Hauteur) Then
                    If Perso.Pietinement > .Pietinement Then
                        'N'écrase que les personnes qu'on une valeur de piétinement inférieure.
                        Perso_Attaquer_Perso Perso, Persos(i), False
                    End If
                End If
            End If
        End If
        End With
    Next i
    'Ecrase les bâtiments ennemis.
    For i = 0 To UBound(Maisons())
        If Perso.Attaque < Perso.Pietinement Then
            Perso.Definir_Attaque = Perso.Pietinement
        End If
        With Maisons(i)
        If Persos(i).NumeroEquipe <> Pne Then
            If Maisons(i).Visible Then
                If Est_Superpose(X, _
                                 Y, _
                                 .PositionX, _
                                 .PositionY, _
                                 Largeur, _
                                 Hauteur, _
                                 .Largeur, _
                                 .Hauteur) Then
                    Perso_Attaquer_Batiment Perso, Maisons(i), Persos(i)
                End If
            End If
        End If
        End With
    Next i
    'Ecrase les chateaux.
    For i = 0 To UBound(Chateaux())
        With Chateaux(i)
        If Est_Superpose(X, _
                         Y, _
                         .PositionX, _
                         .PositionY, _
                         Largeur, _
                         Hauteur, _
                         .Largeur, _
                         .Hauteur) Then
            If .Visible Then
                 If Fiefs(i).NumeroEquipe <> Pne Then
                     Perso_Attaquer_Batiment Perso, Chateaux(i), Persos(Perso.IndicePerso), i
                 End If
            End If
        End If
        End With
    Next i
End Sub

Public Sub Perso_Hostile(ByRef Perso As ClsJeuPerso)
    'Le personnage devient hostile et est l'ennemi de tout le monde.
    Dim Nf As Long
    Dim i, j As Long
    'Cherche la dernière équipe.
    For i = 0 To UBound(Fiefs())
        If Fiefs(i).NumeroEquipe > j Then
            j = Fiefs(i).NumeroEquipe
        End If
    Next i
    Nf = Fief_Ajouter_Nouveau
    With Fiefs(Nf)
        .NumeroEquipe = j + 1
        
        .Definir_TypePeuple = Fiefs(Perso.NumeroFief).Definir_TypePeuple
        .Nom = Perso.Nom
        
        .Epoque = Fiefs(Perso.NumeroFief).Epoque
        .ChangerEpoque = Fiefs(Perso.NumeroFief).ChangerEpoque
        .TempsChangementEpoque = Fiefs(Perso.NumeroFief).TempsChangementEpoque
        
        .NombreCitoyens = 1
        .NombreMorts = IIf(Perso.Vivant, 0, 1)
    End With
    Fiefs(Perso.NumeroFief).NombreCitoyens = Fiefs(Perso.NumeroFief).NombreCitoyens - 1
    Fiefs(Perso.NumeroFief).NombreMorts = Fiefs(Perso.NumeroFief).NombreMorts - IIf(Perso.Vivant, 0, 1)
    Perso.NumeroFief = Nf
    Perso.NumeroEquipe = Fiefs(Nf).NumeroEquipe
    Perso.Arreter
    'Convertit les soldats.
    For i = 0 To UBound(Persos())
        With Persos(i)
            If .NumeroChef = Perso.Numero Then
                .NumeroFief = Nf
                .NumeroEquipe = Fiefs(Nf).NumeroEquipe
                If .IA_Ordre = 4 Then .IA_Ordre = 0
            End If
        End With
    Next i
    Perso.DansUneMaison = False
    Perso.DansSonChateau = False
    Perso.DansUnChateau = False
End Sub

Public Function Perso_Allie(ByVal Perso1 As ClsJeuPerso, ByVal Perso2 As ClsJeuPerso) As Boolean
    'Renvoie vrai si les 2 personnages sont alliés.
    Perso_Allie = Perso1.NumeroEquipe = Perso2.NumeroEquipe
End Function

Public Function Distance_Persos(ByVal Perso1 As ClsJeuPerso, _
                                ByVal Perso2 As ClsJeuPerso) As Double
    Distance_Persos = Sqr((Perso1.PositionX - Perso2.PositionX) ^ 2 + (Perso1.PositionY - Perso2.PositionY) ^ 2)
End Function
Public Function Distance_Positions(ByVal X1 As Long, _
                                   ByVal Y1 As Long, _
                                   ByVal X2 As Long, _
                                   ByVal Y2 As Long) As Double
    Distance_Positions = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
End Function
