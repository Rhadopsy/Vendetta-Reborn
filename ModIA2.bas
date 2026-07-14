Attribute VB_Name = "ModIA2"
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

Public Const CheminIA = "IAs\"
Public Const CheminIAScript = "IAScripts\"
Public Const ExtentionFichierIA = ".ia"
Public Const ExtentionFichierIAScipt = ".txt"
Public Const FichierIADefaut = "_Defaut"
'Public Const CoefficientIADefaut = 0.25

Const IAEmpereur = "Empereur.ia"

Public IA2s() As ClsJeuIA2

'Variables pour des infos temporaires.
Private X As Long
Private Y As Long
Private TempEpoque As Integer

Public Sub Nouvelle_IA(ByVal Nom As String)
    Dim i As Integer
    Dim Creer As Boolean
    'On vérifie si le nom entré est déjà présent.
    Creer = True
    For i = 0 To UBound(IA2s())
        If IA2s(i).Nom = Nom Then
            Creer = False
        End If
    Next i
    
    If Creer Then
        ReDim Preserve IA2s(UBound(IA2s()) + 1)
        Set IA2s(UBound(IA2s())) = New ClsJeuIA2
        IA2s(UBound(IA2s())).Creer_IA_Vierge App.Path & "\" & CheminIA, Nom, Parametres
        Charger_Dossiers_IAs
    Else
        MsgBox "Le fichier " & Nom & " existe déjà", vbCritical, "Création impossible"
    End If
End Sub

Public Sub Renommer_IA(ByVal Nom1 As String, ByVal Nom2 As String)
    On Error GoTo Erreur
    Enregistrer_IAs
    FileCopy App.Path & "\" & CheminIA & Nom1, App.Path & "\" & CheminIA & Nom2
    Supprimer_IA Nom1
    Exit Sub
Erreur:
    MsgBox "Nom de fichier invalide.", vbCritical, "Erreur"
End Sub

Public Sub Duppliquer_IA(ByVal Nom1 As String, ByVal Nom2 As String)
    Enregistrer_IAs
    FileCopy App.Path & "\" & CheminIA & Nom1, App.Path & "\" & CheminIA & Nom2
    Charger_Dossiers_IAs
End Sub

Public Sub Supprimer_IA(ByVal Nom As String)
    Kill App.Path & "\" & CheminIA & Nom
    Charger_Dossiers_IAs
End Sub

Public Sub Charger_Dossiers_IAs()
    'Charge toutes les IAs situées dans le répertoire et les sous répertoires du dossier des IAs.
    'On Error GoTo Erreur
    Dim Temp As String
    Dim Dossiers As Collection 'Mémorise la liste des dossiers où sont les ias.
    Dim i As Integer
    Dim j As Integer
    Set Dossiers = New Collection
    ReDim IA2s(0)
    Set IA2s(0) = New ClsJeuIA2
    
    Temp = Dir(App.Path & "\" & CheminIA, vbDirectory)
    While Temp <> ""
        If InStr(Temp, ".") = 0 Then
            Dossiers.Add Temp
        End If
        Temp = Dir()
    Wend
    
    Charger_IAs App.Path & "\" & CheminIA, i, True
    For j = 1 To Dossiers.Count
        Charger_IAs App.Path & "\" & CheminIA & Dossiers.Item(j) & "\", i, False
    Next j
    'Exit Sub
'Erreur:
End Sub

Public Sub Charger_IAs(ByVal Dossier As String, ByRef i As Integer, ByVal Joueur As Boolean)
    'Charge toutes les IAs dans un répertoire.
    On Error GoTo Erreur
    Dim Temp As String
    
    Temp = Dir(Dossier, vbNormal)
    While Temp <> ""
        ReDim Preserve IA2s(i)
        Set IA2s(i) = New ClsJeuIA2
        IA2s(i).Charger_IA Dossier, Temp, Parametres, Joueur
        i = i + 1
        Temp = Dir()
    Wend
    Exit Sub
Erreur:
    MsgBox "Les fichiers d'IAs sont introuvables.", vbCritical, "Erreur"
    End
End Sub

Public Sub Enregistrer_IAs()
    Dim i As Integer
    For i = 0 To UBound(IA2s())
        If IA2s(i).Joueur Then
            IA2s(i).Sauvegarder_IA App.Path & "\" & CheminIA, IA2s(i).Nom, Parametres
        End If
    Next i
End Sub

Public Function Charger_IA(ByVal FichierIA As String) As Integer
    Dim i As Long
    Dim Temp As String
    Charger_IA = -1
    If FichierIA <> "" Then
        'Temp = Dir(App.Path & "\" & CheminIA, vbNormal)
        'While Temp <> ""
        '    If LCase(Temp) = LCase(FichierIA) Then
        '        Charger_IA = i
        '    End If
        '    i = i + 1
        '    Temp = Dir()
        'Wend
        For i = 0 To UBound(IA2s())
            If LCase(IA2s(i).Nom) = LCase(FichierIA) Then
                Charger_IA = i
                Exit For
            End If
        Next i
    End If
End Function

Public Sub Comportement_IA2(ByVal i As Long, Optional ByVal Simple As Boolean)
    Dim j As Integer
    Dim NumeroIA As Integer
    
'    If i = NoPerso Then
'        i = i
'    End If
    
    'Comportement urgent.
    IA_Poursuivre_Fuyard i
    If Persos(i).Action <> 3 And _
       Persos(i).Action <> 500 Then
        IA_Recuperation_Carac i
    End If
    
    If Not Persos(i).EnCombat And Persos(i).Action <> 2 And _
                         Persos(i).Action <> 3 And _
                         Persos(i).Action <> 102 And _
                         Persos(i).Action <> 205 And _
                         Persos(i).Action <> 500 And _
                         Persos(i).Action <> 600 Then
        'N'est pas en train de combattre ou de travailler.
    'If .Inactif Then
        'N'est pas occupé.
        IA_Vente_Surplus_Objets i
        'If .Definir_Vie = .Definir_MaxVie And _
           .Vitalite = 1 And _
           .Definir_Magie >= .Definir_MaxMagie * 0.5 Then
        'Fabrique sa maison.
        If Maisons(i).Visible And Maisons(i).Vie < Maisons(i).MaxVie Then
            'Va finir de construire sa maison.
            IA_Rentrer_Maison i
            'Messages.Ajouter_Message "IA_Réparer Maison"
        ElseIf BatimentAConstruire(i) >= 0 And _
           (Maisons(i).TypeBatiment <> BatimentAConstruire(i) Or Not Maisons(i).Visible) Then
           'S'occupe de fabriquer sa maison.
            'If Maisons(i).Visible Then
                'Détuit sa maison
                'Maisons(i).Demolir Persos(i), Parametres
                'Messages.Ajouter_Message "IA_Démolit Maison"
            If Persos(i).PeutConstruire_Maison(BatimentAConstruire(i), Parametres) Then
                If IA_Changer_Maison(i) Then
                    'Parametres.Batiment_NoEpoque(BatimentAConstruire(i)) - 1 <= Fiefs(.NumeroFief).Epoque Then
                    'Va construire une maison.
                    If Not Simple Then
                        If Trouver_Emplacement_Maison(i, X, Y, BatimentAConstruire(i)) Then
                            Maisons(i).Creer X, Y, BatimentAConstruire(i), i, Parametres
                            'IA_Rentrer_Maison i
                        End If
                    End If
                    'Messages.Ajouter_Message "IA_Construit : " & Parametres.Batiment_Nom(BatimentAConstruire(i))
                End If
            Else 'Cherche les matières premières nécessaire à sa contruction.
                For j = 0 To Parametres.NombreRessources - 1
                    If Persos(i).Definir_ressources(j) < Parametres.Batiment_PrixRessource(BatimentAConstruire(i), j) Then
                        IA2_ChercherMatieresPremiere i, j, Parametres.Batiment_PrixRessource(BatimentAConstruire(i), j)
                        Exit For
                    End If
                Next j
            End If
        ElseIf Maisons(i).Visible And _
               Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment) >= 0 And _
               Persos(i).Definir_ressources(Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment)) < Persos(i).Definir_MaxRessources(Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment)) And _
               IIf(Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment) = -1, False, Parametres.Ressources_Recoltable(IIf(Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment) = -1, 0, Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)))) Then
            'Cas où le personnage peut fabriquer des ressources avec sa maison sans dépenser de l'argent.
            If Persos(i).Definir_ressources(Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)) <= 1 Then
                'N'a plus de matières premières.
                IA_Chercher_Ressource i, Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)
               ' Messages.Ajouter_Message "Mat. 1ères"
            Else
                IA_Rentrer_Maison i
                'Messages.Ajouter_Message "Goto maison"
            End If
        ElseIf RessourceAFabriquer(i) >= 0 Then
            'S'occupe de fabriquer ses ressources.
            If Persos(i).Definir_ressources(RessourceAFabriquer(i)) < QuantiteAFabriquer(i) Then
                If Parametres.Ressources_Recoltable(RessourceAFabriquer(i)) Then
                    IA_Chercher_Ressource i, RessourceAFabriquer(i)
                Else
                    If Maisons(i).Visible And _
                       Parametres.Batiment_RessourceFabrique(Maisons(i).TypeBatiment) = RessourceAFabriquer(i) Then
                        If Not Persos(i).DansSaMaison Then
                            IA_Rentrer_Maison i
                        ElseIf Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment) >= 0 Then
                            'N'a plus de ressources.
                            If Persos(i).Definir_ressources(Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)) < 1 Then
                                IA2_ChercherMatieresPremiere i, Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment), Parametres.PersosMaxRessources ' (QuantiteAFabriquer(i) - Persos(i).Definir_Ressources(RessourceAFabriquer(i))) * Parametres.Ressources_QuantiteMatierePremiere(Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment))
                            End If
                        End If
                    ElseIf Persos(i).Action <> 202 And Persos(i).Action <> 203 Then
                        'Si n'est pas déjà en train de retourner au chateau.
                        IA2_ChercherMatieresPremiere i, RessourceAFabriquer(i), QuantiteAFabriquer(i)
                    End If
                End If
            Else
                RessourceAFabriquer(i) = -1
            End If
        ElseIf ObjetAFabriquer(i) >= 0 Then
            'Fabrique lui même son objet.
            If Not Persos(i).Action = 500 Then
                If Persos(i).Objet_NombreType(ObjetAFabriquer(i)) < QuantiteAFabriquer(i) Then
                    If Maisons(i).Visible Then
                        If Not Persos(i).DansSaMaison Then
                            IA2_ChercherObjet i, ObjetAFabriquer(i), QuantiteAFabriquer(i)
                            If Persos(i).Inactif Then IA_Rentrer_Maison i
                        ElseIf Maisons(i).Fabrique Then
                            For j = 0 To Parametres.Batiment_NombreObjetsFabriques(Maisons(i).TypeBatiment) - 1
                                If Parametres.Batiment_ObjetFabrique(Maisons(i).TypeBatiment, j) = ObjetAFabriquer(i) Then
                                    If Maisons(i).Stock_Objets(j) = 0 Then
                                        If Persos(i).PeutConstruire_Objet(j, Maisons(i), Parametres) Then
                                            Persos(i).Creer_Objet j, Maisons(i), Parametres
                                        Else
                                            IA2_ChercherObjet i, ObjetAFabriquer(i), QuantiteAFabriquer(i)
                                        End If
                                    Else
                                        Persos(i).Objet_Ramasser_Maison j, Maisons(i), Parametres, Commentaires
                                        ObjetAFabriquer(i) = -1
                                    End If
                                    j = Parametres.Batiment_NombreObjetsFabriques(Maisons(i).TypeBatiment)
                                End If
                            Next j
                        ElseIf IA_Changer_Maison(i) Then
                            'Cas où la maison ne lui permet pas de fabriquer son objet.
                            Maisons(i).Demolir Persos(i), Parametres
                            Comportement_IA2 i
                            BatimentAConstruire(i) = -1
                        End If
                    Else
                        IA2_ChercherObjet i, ObjetAFabriquer(i), QuantiteAFabriquer(i)
                    End If
                End If
            End If
        Else
            'Réalise les étapes de son comportement.
            If Persos(i).Definir_Empereur Then
                NumeroIA = Charger_IA(IAEmpereur)
            Else
                NumeroIA = Persos(i).NumeroIA
            End If
            TempEpoque = Fiefs(Persos(i).NumeroFief).Epoque
            For j = 0 To IA2s(NumeroIA).Definir_NombreEtapes(TempEpoque) - 1
                Select Case IA2s(NumeroIA).Definir_Categorie(TempEpoque, j)
                Case 0: 'Ressources.
                    If Persos(i).Definir_ressources(IA2s(NumeroIA).Definir_Valeur(TempEpoque, j)) < IA2s(NumeroIA).Definir_Quantite(TempEpoque, j) Then
                        'Si on a pas assez de ressources, on va en chercher.
                        IA2_ChercherMatieresPremiere i, IA2s(NumeroIA).Definir_Valeur(TempEpoque, j), IA2s(NumeroIA).Definir_Quantite(TempEpoque, j)
                        Exit For
                    End If
                Case 1: 'Batiment.
                    If Maisons(i).TypeBatiment <> IA2s(NumeroIA).Definir_Valeur(TempEpoque, j) Then
                        If BatimentAConstruire(i) <> IA2s(NumeroIA).Definir_Valeur(TempEpoque, j) Then
                            IA2_Choisir_Maison i, IA2s(NumeroIA).Definir_Valeur(TempEpoque, j)
                            Exit For
                        End If
                    End If
                Case 2: 'Objet.
                    If Persos(i).Objet_NombreType(IA2s(NumeroIA).Definir_Valeur(TempEpoque, j)) < IA2s(NumeroIA).Definir_Quantite(TempEpoque, j) Then
                        IA2_ChercherObjet i, IA2s(NumeroIA).Definir_Valeur(TempEpoque, j), IA2s(NumeroIA).Definir_Quantite(TempEpoque, j)
                        Exit For
                    End If
                Case 3: 'travailler.
                    If Maisons(i).Visible Then
                        If Chateaux(Persos(i).NumeroFief).Visible Or Fiefs(Persos(i).NumeroFief).Marches.Count > 0 Then
                            If IA2s(NumeroIA).Definir_BatimentDefini(TempEpoque) Then
                                'Sa maison est précisée dans le fichier.
                                IA_ChercherMatieresPremiere i: IA_Travail i
                            Else
                                'Choisit une maison de niveau supérieure.
                                BatimentAConstruire(i) = Choisir_Batiment(i, Persos(i).NumeroFief, Parametres.Batiment_Niveau(Maisons(i).TypeBatiment) + 1)
                                If BatimentAConstruire(i) >= 0 Then
                                    If Parametres.Batiment_NoEpoque(BatimentAConstruire(i)) - 1 > Fiefs(Persos(i).NumeroFief).Epoque Then
                                        BatimentAConstruire(i) = Maisons(i).TypeBatiment
                                        IA_ChercherMatieresPremiere i: IA_Travail i
                                    End If
                                Else
                                    IA_ChercherMatieresPremiere i: IA_Travail i
                                End If
                            End If
                        Else
                            IA_Agressif i
                        End If
                    Else
                        BatimentAConstruire(i) = Choisir_Batiment(i, Persos(i).NumeroFief)
                    End If
                    Exit For
                    'Messages.Ajouter_Message "IA_Travail"
                Case 4: 'patrouiller.
                    IA_Agressif i: IA_Agressif i: IA_Agressif i
                    Exit For
                Case 5: 'attaquer monstre.
                    If Victoire(0) Then
                        IA_Agressif i
                    ElseIf Not Persos(i).EnCombat And _
                           Persos(i).Definir_Attaque = Persos(i).Definir_MaxAttaque Then
                        IA_Trouver_Ennemi i, False, False, False, False
                        If Persos(i).NombreSoldats > 0 Then
                            Ordonner_Chargez i
                        End If
                        If Not Persos(i).EnCombat Then
                            IA_Agressif i
                        End If
                    End If
                    Exit For
                Case 6: 'attaquer donjon.
                    If Victoire(0) Then
                        IA_Agressif i
                    ElseIf Not Persos(i).EnCombat And _
                           Persos(i).Definir_Attaque = Persos(i).Definir_MaxAttaque Then
                        IA_Trouver_ChateauEnnemi i, False, False, False, False
                        If Persos(i).NombreSoldats > 0 Then
                            Ordonner_Chargez i
                        End If
                        If Not Persos(i).EnCombat And _
                           Persos(i).Definir_Attaque = Persos(i).Definir_MaxAttaque Then
                            IA_Trouver_Ennemi i, False, False, False, False
                            If Persos(i).NombreSoldats > 0 Then
                                Ordonner_Chargez i
                            End If
                            If Not Persos(i).EnCombat Then
                                IA_Agressif i
                            End If
                        End If
                    End If
                    Exit For
                Case 7:
                    Select Case IA2s(NumeroIA).Definir_Valeur(TempEpoque, j)
                    Case 0: Persos(i).IA_Pacifique = True
                    Case 1: Persos(i).IA_Pacifique = False
                    Case 2: If IA_Donner(i, True, True, True) Then Exit Sub
                    Case 3:
                        If IA_Donner(i, True, False, False) Then Exit Sub
                    Case 4: If IA_Donner(i, False, True, False) Then Exit Sub
                    Case 5: If IA_Donner(i, False, False, True) Then Exit Sub
                    Case 6:
                        If IA_Invoquer(i) Then Exit Sub
                    Case 7: Persos(i).IA_Protecteur = True
                    Case 8: Persos(i).IA_Protecteur = False
                    Case 9: Persos(i).IA_Egoiste = True
                    Case 10: Persos(i).IA_Egoiste = False
                    Case 11: Persos(i).IA_Equiper = 0
                    Case 12: Persos(i).IA_Equiper = 1
                    Case 13: Persos(i).IA_Equiper = 2
                    End Select
                Case 8:
                    If Persos(i).Definir_Argent < IA2s(NumeroIA).Definir_Quantite(TempEpoque, j) Then
                        'Collecte suffisament d'argent.
                        If Maisons(i).Visible And _
                           Maisons(i).TypeBatiment = IA2s(NumeroIA).Definir_Valeur(TempEpoque, j) Then
                            If Chateaux(Persos(i).NumeroFief).Visible Or Fiefs(Persos(i).NumeroFief).Marches.Count > 0 Then
                                'IA_ChercherMatieresPremiere i
                                IA_Travail i
                            Else
                                IA_Agressif i
                            End If
                        Else
                            IA2_Choisir_Maison i, IA2s(NumeroIA).Definir_Valeur(TempEpoque, j)
                            'BatimentAConstruire(i) = IA2s(.NumeroIA).Definir_Valeur(TempEpoque, j)
                        End If
                        Exit For
                    End If
                Case 9:
                    If Persos(i).Definir_Argent < IA2s(NumeroIA).Definir_Quantite(TempEpoque, j) Then
                        'Collecte suffisament d'argent en recoltant des ressources.
                        If Chateaux(Persos(i).NumeroFief).Visible Or Fiefs(Persos(i).NumeroFief).Marches.Count > 0 Then
                            If Persos(i).Definir_ressources(IA2s(NumeroIA).Definir_Valeur(TempEpoque, j)) < Persos(i).Definir_MaxRessources(IA2s(NumeroIA).Definir_Valeur(TempEpoque, j)) Then
                                'Si on a pas assez de ressources, on va en chercher.
                                IA_Chercher_Ressource i, IA2s(NumeroIA).Definir_Valeur(TempEpoque, j)
                            'Vend cette ressource au château.
                            ElseIf (Persos(i).DansUnChateau Or Persos(i).DansSonChateau) Or _
                                    (Persos(i).DansUneMaison And Maisons(Persos(i).IndiceMaison).Marche) Then
                                    Fiefs(Persos(i).NumeroFief).Vendre IA2s(NumeroIA).Definir_Valeur(TempEpoque, j), 0, Persos(i), Parametres, True
                            'Va au château pour vendre des ressources.
                            Else
                                IA_Rentrer_Chateau i, True
                            End If
                        Else
                            IA_Agressif i
                        End If
                        Exit For
                    End If
                End Select
            Next j
        End If
    End If
    
    'If i = 0 Then _
    Messages(0).Ajouter_Message "Debug_IA (" & Persos(i).FichierIA & ")   " & _
                             "M : " & IIf(BatimentAConstruire(i) = -1, "-", Parametres.Batiment_Nom(IIf(BatimentAConstruire(i) = -1, 0, BatimentAConstruire(i)))) & _
                             " | R : " & IIf(RessourceAFabriquer(i) = -1, "-", Parametres.Ressources_Nom(IIf(RessourceAFabriquer(i) = -1, 0, RessourceAFabriquer(i)))) & _
                             " | O : " & IIf(ObjetAFabriquer(i) = -1, "-", Parametres.Objet_Nom(IIf(ObjetAFabriquer(i) = -1, 0, ObjetAFabriquer(i)))) & _
                             " | Q : " & QuantiteAFabriquer(i) & _
                             " Pacifique = " & Persos(i).IA_Pacifique
    
End Sub

Private Sub IA2_ChercherMatieresPremiere(ByVal i As Long, _
                                         ByVal TypeRessource As Integer, _
                                         ByVal Quantite As Double)
    With Persos(i)
    'If Parametres.Ressources_TypeMatierePremiere(TypeRessource) >= 0 Then
        '.Definir_Ressources(Parametres.Ressources_TypeMatierePremiere(TypeRessource)) < Quantite Then
        'Recherche d'abord les matières premières de cette ressource. (fonction récursive)
        'IA2_ChercherMatieresPremiere i, Parametres.Ressources_TypeMatierePremiere(TypeRessource), Quantite * Parametres.Ressources_QuantiteMatierePremiere(TypeRessource)
        'RessourceAFabriquer(i) = Parametres.Ressources_TypeMatierePremiere(TypeRessource)
        'IA2_Fabriquer_Ressource i, RessourceAFabriquer(i), Quantite
    
'    If i = NoPerso Then
'        i = i
'    End If
    
    'Va récolter si c'est une matière première.
    If Parametres.Ressources_Recoltable(TypeRessource) Then
        IA_Chercher_Ressource i, TypeRessource
        RessourceAFabriquer(i) = TypeRessource
        QuantiteAFabriquer(i) = Quantite
    Else
        'Va en acheter.
        If .Definir_Argent > Quantite * Parametres.Ressources_PrixVente(TypeRessource) Then
            If .Action <> 201 And .Action <> 202 And Not (.DansUneMaison Or .DansSonChateau Or .DansUnChateau) Then
                If Trouver_Maison_Proche(Persos(i), i, TypeRessource, 0, 0, X) Then
                    'Se rend dans un magasin.
                    If X >= 0 Then
                        IA_Aller_Maison i, X
                    End If
                    RessourceAFabriquer(i) = TypeRessource
                    QuantiteAFabriquer(i) = Quantite
                ElseIf Fiefs(.NumeroFief).Definir_ressources(Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment)) > Quantite And _
                       (Chateaux(.NumeroFief).Visible Or _
                        Fiefs(.NumeroFief).Marches.Count > 0) Then
                    'Se rend au chateau.
                    IA_Rentrer_Chateau i, True
                    RessourceAFabriquer(i) = TypeRessource
                    QuantiteAFabriquer(i) = Quantite
                Else
                    IA2_Fabriquer_Ressource i, TypeRessource, Quantite
                End If
                'Messages.Ajouter_Message "IA_Aller_Acheter_Ressource : " & Parametres.Ressources_Nom(TypeRessource)
            Else
                If .DansUneMaison And Maisons(.IndiceMaison).Magasin Then
                    If Maisons(.IndiceMaison).Definir_Stock >= 1 Then
                        Maisons(.IndiceMaison).Acheter Quantite - .Definir_ressources(TypeRessource), Persos(i), Parametres
                        RessourceAFabriquer(i) = -1
                    Else
                        'Le magasin n'avait pas assez de ressources.
                        'Nouvelle recherche avec une fonction recursive.
                        .DansUneMaison = False
                        IA2_ChercherMatieresPremiere i, TypeRessource, Quantite
                    End If
                    'Messages.Ajouter_Message "IA_Transaction Magasin : " & Parametres.Ressources_Nom(TypeRessource)
                ElseIf (.DansUnChateau Or .DansSonChateau Or (.DansUneMaison And Maisons(.IndiceMaison).Marche)) And Fiefs(.NumeroFief).Definir_ressources(TypeRessource) >= Quantite Then
                    Fiefs(.NumeroFief).Acheter TypeRessource, Quantite - .Definir_ressources(TypeRessource), Persos(i), Parametres
                    RessourceAFabriquer(i) = -1
                    'Messages.Ajouter_Message "IA_Transaction Chateau : " & Parametres.Ressources_Nom(TypeRessource)
                Else
                    IA2_Fabriquer_Ressource i, TypeRessource, Quantite
                End If
                'Messages.Ajouter_Message "IA_Acheter Ressource : " & Parametres.Ressources_Nom(TypeRessource)
            End If
        Else
        'Doit le fabriquer soi-même.
            IA2_Fabriquer_Ressource i, TypeRessource, Quantite
        End If
    End If
    End With
    
    'Messages.Ajouter_Message "IA_Recherche : " & Parametres.Ressources_Nom(TypeRessource)
End Sub

Public Sub IA2_Fabriquer_Ressource(ByVal i As Long, _
                                   ByVal TypeRessource As Integer, _
                                   ByVal Quantite As Double)
    Dim j As Integer
    'Vérifie s'il faut commencer par la matière première.
    While Parametres.Ressources_TypeMatierePremiere(TypeRessource) >= 0 And _
       Persos(i).Definir_ressources(Parametres.Ressources_TypeMatierePremiere(TypeRessource)) < Quantite
        TypeRessource = Parametres.Ressources_TypeMatierePremiere(TypeRessource)
    Wend
    If Parametres.Ressources_Recoltable(TypeRessource) Then
        IA_Chercher_Ressource i, TypeRessource
    Else
        For j = 0 To Parametres.NombreTypeBatiments - 1
            If Parametres.Batiment_RessourceFabrique(j) = TypeRessource Then
                BatimentAConstruire(i) = j
                j = Parametres.NombreTypeBatiments
            End If
        Next j
    End If
    RessourceAFabriquer(i) = TypeRessource
    QuantiteAFabriquer(i) = Quantite
End Sub

Public Sub IA2_ChercherObjet(ByVal i As Long, _
                             ByVal TypeObjet As Integer, _
                             ByVal Quantite As Double)
    Dim j As Long, k As Long
    Dim Decision As Boolean
    With Persos(i)
    'Messages.Ajouter_Message "IA_Recherche : " & Parametres.Objet_Nom(TypeObjet)
    'Va en acheter l'objet en question.
    If Not .Action = 201 Then
        
        If .Definir_Argent >= Parametres.Objet_PrixVente(TypeObjet) Then
            For j = 0 To UBound(Maisons())
                If Maisons(j).Fabrique And Persos(j).NumeroEquipe = Persos(i).NumeroEquipe Then
                    For k = 0 To Parametres.Batiment_NombreObjetsFabriques(Maisons(j).TypeBatiment) - 1
                        If Parametres.Batiment_ObjetFabrique(Maisons(j).TypeBatiment, k) = TypeObjet And _
                           Maisons(j).Stock_Objets(k) >= Quantite Then
                            If i = j And .DansSaMaison Then
                                'Achète son objet chez lui.
                                .Objet_Ramasser_Maison k, Maisons(j), Parametres, Commentaires
                                ObjetAFabriquer(i) = -1
                            Else
                                'Prend l'objet dans un magasin.
                                If .DansUneMaison And .IndiceMaison = j Then
                                    .Objet_Acheter_Maison k, Maisons(j), Parametres, Commentaires
                                    ObjetAFabriquer(i) = -1
                                Else
                                    IA_Aller_Maison i, j
                                End If
                            End If
                            'Messages.Ajouter_Message "IA_Achète : " & Parametres.Objet_Nom(TypeObjet)
                            Decision = True
                            k = Parametres.Batiment_NombreObjetsFabriques(Maisons(j).TypeBatiment)
                            j = UBound(Maisons())
                        End If
                    Next k
                End If
            Next j
        End If
        
        If Not Decision Then
            'Vérifie si cet objet est déjà dans sa maison.
            If Maisons(i).Fabrique And _
               Maisons(i).Definir_Stock > 0 Then
                For j = 0 To Maisons(i).Definir_Nombre_Stock_Objet - 1
                    If Maisons(i).Stock_Objets(j) > 0 And _
                       Parametres.Batiment_ObjetFabrique(Maisons(i).TypeBatiment, j) = TypeObjet Then
                        IA_Rentrer_Maison i
                        Decision = True
                    End If
                Next j
            End If
        End If
            
        If Not Decision Then
            'Doit le fabriquer soi-même.
            ObjetAFabriquer(i) = TypeObjet
            QuantiteAFabriquer(i) = Quantite
            'Collecte les ressources nécessaires.
            For j = 0 To Parametres.NombreRessources - 1
                If .Definir_ressources(j) < Parametres.Objet_PrixRessources(TypeObjet, j) Then
                    IA2_ChercherMatieresPremiere i, j, Parametres.Objet_PrixRessources(TypeObjet, j)
                    j = Parametres.NombreRessources
                    Decision = True
                    'Messages.Ajouter_Message "IA_Fabrique : " & Parametres.Objet_Nom(TypeObjet)
                End If
                'Messages.Ajouter_Message .Definir_Ressources(j) & "<" & Parametres.Objet_PrixRessources(TypeObjet, j)
            Next j
        End If
        
        If Not Decision Then
            'Choisit la maison qui fabrique son objet.
            For j = 0 To Parametres.NombreTypeBatiments - 1
                If Parametres.Batiment_Fabrique(j) Then
                    For k = 0 To Parametres.Batiment_NombreObjetsFabriques(j) - 1
                        If Parametres.Batiment_ObjetFabrique(j, k) = TypeObjet Then
                            BatimentAConstruire(i) = j
                            ObjetAFabriquer(i) = TypeObjet
                            QuantiteAFabriquer(i) = Quantite
                            k = Parametres.Batiment_NombreObjetsFabriques(j)
                            j = Parametres.NombreTypeBatiments
                            'Messages.Ajouter_Message "IA_Construit_Fabrique : " & Parametres.Objet_Nom(TypeObjet)
                        End If
                    Next k
                End If
            Next j
        End If
    End If
    End With
End Sub

Public Sub IA2_Choisir_Maison(ByVal i As Long, ByVal TypeBatiment As Integer)
    Dim j As Integer, k As Integer
    With Parametres
    RessourceAFabriquer(i) = -1
    'On vérifie déjà que l'IA a déjà suffisamment de matières premières avant de travailler.
    If Parametres.Batiment_Fabrique(TypeBatiment) Then
        'Si c'est une fabrique d'objets.
        For j = 0 To .NombreRessources - 1
            If .Ressources_Epoque(j) <= Fiefs(Persos(i).NumeroFief).Epoque Then
                For k = 0 To .Batiment_NombreObjetsFabriques(TypeBatiment) - 1
                    If .Objet_PrixRessources(.Batiment_ObjetFabrique(TypeBatiment, k), j) > 0 And _
                       Persos(i).Definir_ressources(j) < .PersosMaxRessources Then
                        IA2_Fabriquer_Ressource i, j, .PersosMaxRessources
                        'RessourceAFabriquer(i) = j
                        j = .NombreRessources
                        k = .Batiment_NombreObjetsFabriques(TypeBatiment)
                    End If
                Next k
            End If
        Next j
    ElseIf .Batiment_Service(TypeBatiment) Then
        'Si c'est un batiment de service.
        j = .Service_MatierePremiere(.Batiment_Service(TypeBatiment))
        If j >= 0 And _
           Persos(i).Definir_ressources(j) < Persos(i).Definir_MaxRessources(j) Then
            IA2_Fabriquer_Ressource i, j, Persos(i).Definir_MaxRessources(j)
            'RessourceAFabriquer(i) = j
        End If
    ElseIf .Batiment_RessourceFabrique(TypeBatiment) >= 0 Then
        'Si c'est un batiment qui produit des ressources.
        j = .Ressources_TypeMatierePremiere(.Batiment_RessourceFabrique(TypeBatiment))
        If j >= 0 And _
           Persos(i).Definir_ressources(j) < Persos(i).Definir_MaxRessources(j) Then
            IA2_Fabriquer_Ressource i, j, Persos(i).Definir_MaxRessources(j)
            'RessourceAFabriquer(i) = j
        End If
    End If
    If RessourceAFabriquer(i) = -1 Then
        'QuantiteAFabriquer(i) = Persos(i).Definir_MaxRessources(RessourceAFabriquer(i))
        BatimentAConstruire(i) = TypeBatiment
    End If
    End With
End Sub

Private Function IA_Changer_Maison(ByVal i As Long) As Boolean
    'Revoie vrai si l'IA peut détruire sa maison ou en construire une autre.
    Dim j As Integer
    If Maisons(i).Visible Then
       If Maisons(i).Argent > 0 Or _
          (Maisons(i).Definir_Stock > 0 And Not Maisons(i).Service) Then
            If Persos(i).DansSaMaison Then
                'Récupère son argent.
                Maisons(i).Transfert_Argent_Prendre 0, Persos(i), True
                If Maisons(i).Magasin Then 'Récupère toutes ses ressources.
                    Maisons(i).Transfert_Stock_Prendre 0, Persos(i), Parametres, True
                ElseIf Maisons(i).Fabrique Then 'Récupère tous ses objets.
                    For j = 0 To Parametres.Batiment_NombreObjetsFabriques(Maisons(i).TypeBatiment) - 1
                        While Maisons(i).Stock_Objets(j) >= 1 And Not Persos(i).Inventaire_Plein
                            Persos(i).Objet_Ramasser_Maison j, Maisons(i), Parametres, Commentaires
                        Wend
                    Next j
                End If
                IA_Changer_Maison = True
            Else
                IA_Rentrer_Maison i
            End If
        Else
            IA_Changer_Maison = True
        End If
    Else
        IA_Changer_Maison = True
    End If
End Function

Private Function IA_Donner(ByVal i As Long, _
                           Optional ByVal DonnerArgent As Boolean, _
                           Optional ByVal DonnerRessources As Boolean, _
                           Optional ByVal DonnerObjet As Boolean) As Boolean
    'Va tout donner à son maître ou un allié au hasard
    Dim IndiceCible As Boolean 'Cherche à qui il doit donner.
    Dim j As Long: Dim k As Long
    With Persos(i)
    IndiceCible = -1
    If .Invocation Then
        If Persos(.NumeroChef).Vivant And _
           Distance_Persos(Persos(.NumeroChef), Persos(i)) <= 1024 And _
           (Persos(.NumeroChef).Visible Or Persos(.NumeroChef).DansSaMaison) Then
            IndiceCible = .NumeroChef
        End If
    Else
        For j = 0 To LPerso
            k = Int(Rnd * LPerso)
            If Perso_Allie(Persos(i), Persos(k)) Then
                IndiceCible = k
            End If
        Next j
    End If
    If IndiceCible > -1 Then
        'Donne son argent.
        If DonnerArgent Then
            If .Definir_Argent >= 1 Then
                .ArgentSelectionne = .Definir_Argent
                IA_Donner = True
            End If
        End If
        'Donne ses ressources.
        If DonnerRessources Then
            For j = 0 To Parametres.NombreRessources - 1
                If Persos(i).Definir_ressources(j) >= 1 Then 'Si l'a au moins une ressource.
                    If Persos(IndiceCible).Definir_MaxRessources(j) - Persos(IndiceCible).Definir_ressources(j) > 1 Then  'Si l'autre n'a pas ses ressources au maximum.
                    'If Persos(IndiceCible).Definir_ressources(j) < 20 Then '- Persos(IndiceCible).Definir_ressources(j) > 1 Then 'Si l'autre n'a pas ses ressources au maximum.
                        .Ressources_Prendre_Tout
                        IA_Donner = True
                        Exit For
                    End If
                End If
            Next j
        End If
        'Donne un objet.
        If DonnerObjet Then
            If Persos(i).Objet_Nombre_Total(True) >= 1 Then  'Si il a au moins un objet.
                If Not Persos(IndiceCible).Inventaire_Plein Then 'Si l'autre n'a pas déjà son inventaire de plein.
                    .Objet_Prendre_Premier
                    IA_Donner = True
                End If
            End If
        End If
        
        'Va donner.
        If IA_Donner Then .Aller_Donner IndiceCible
    End If
    End With
End Function

Private Function IA_Invoquer(ByVal i As Long) As Boolean
    'Invoque les créatures qu'il a dans son inventaire.
    Dim j As Integer
    With Persos(i)
    For j = 0 To .Nombre_Objets_Inventaire - 1
        If .Objet_Inventaire_Actif(j) Then
            If Parametres.Objet_Invocation(.Objet_Inventaire_Type(j)) >= 0 Then
                .Objet_Equiper j, Parametres, Commentaires
                If .NumeroObjetUtilise >= 0 Then
                    IA_Invoquer = True
                    Exit For
                End If
            End If
        End If
    Next j
    End With
End Function
