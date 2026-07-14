Attribute VB_Name = "ModScenario"
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

Public Const ExtensionFichiersScenario = ".snr"
Public Const CheminCampagnes = "\Sauvegardes\Campagnes\"
Public Const CheminScenario = "\Sauvegardes\Scenarios\"

Public Const ExtensionFichiersScript = ".scr"
Private SriptRunner As MSScriptControl.ScriptControl
Private Evenements As String

Public Sub Charger_Scenario(ByVal Nom As String, _
                            Optional ByVal Campagne As Boolean)
    Dim i As Long, j As Long
    Dim Temp As String
    Dim X As String
    Dim Y As String
    Dim PosX As Long
    Dim PosY As Long
    Dim Min As Long
    Dim Max As Long
    Dim Prop As String
    Dim Rangees As Long
    Dim ChoixTerrain As Integer
    Rangees = 1
    
    Parametres.VitesseEpoqueSelectionnee = Parametres.VitesseEpoqueDefaut
    If Campagne Then
        'Si c'est une campagne, on ne le prend pas dans le dossier des scénarios.
        Open App.Path & CheminCampagnes & Langues.Dossier & Nom & ExtensionFichiersScenario For Input As #1
    Else
        Open App.Path & CheminScenario & Langues.Dossier & Nom & ExtensionFichiersScenario For Input As #1
    End If
    
    While Not EOF(1)
        Line Input #1, Temp
        Temp = LCase(Temp)
        If InStr(Temp, "=") And Left(Temp, 1) <> "#" Then
            
            If InStr(Temp, ",") Then
                X = Mid(Temp, InStr(Temp, "=") + 1, InStr(Temp, ",") - InStr(Temp, "=") - 1)
                Y = Mid(Temp, InStr(Temp, ",") + 1, Len(Temp) - InStr(Temp, ","))
            Else
                X = Right(Temp, Len(Temp) - InStr(Temp, "="))
                Y = 0
            End If
            If InStr(Temp, "(") Then
                If InStr(Temp, "-") And InStr(Temp, "-") < InStr(Temp, ")") Then
                    Min = Val(Mid(Temp, InStr(Temp, "(") + 1, InStr(Temp, "-") - InStr(Temp, "(") - 1)) - 1
                    Max = Val(Mid(Temp, InStr(Temp, "-") + 1, InStr(Temp, ")") - InStr(Temp, "-") - 1)) - 1
                Else
                    If LCase(Mid(Temp, InStr(Temp, "(") + 1, InStr(Temp, ")") - InStr(Temp, "(") - 1)) = "tous" Then
                        Min = 0
                        If Temp Like "fief*" Then
                            Max = UBound(Fiefs())
                        Else
                            Max = UBound(Persos())
                        End If
                    Else
                        Min = Val(Mid(Temp, InStr(Temp, "(") + 1, InStr(Temp, ")") - InStr(Temp, "(") - 1)) - 1
                        Max = Min
                    End If
                End If
            End If
            If InStr(Temp, "_") Then
                Prop = Mid(Temp, InStr(Temp, "_") + 1, InStr(Temp, "=") - InStr(Temp, "_") - 1)
            End If
            
            Select Case Left(Temp, InStr(Temp, "=") - 1)
            Case "nombre_fiefs":
                ReDim Fiefs(X - 1)
                ReDim Chateaux(X - 1)
                For i = 0 To UBound(Fiefs())
                    Set Fiefs(i) = New ClsJeuFief
                    Fiefs(i).Init i, Parametres
                    Set Chateaux(i) = New ClsJeuBatiment
                Next i
            Case "nombre_ressources":
                ReDim Ressources(X - 1)
                For i = 0 To UBound(Ressources())
                    Set Ressources(i) = New ClsJeuRessources
                Next i
            Case "nombre_persos":
                ReDim Preserve Persos(X - 1)
                ReDim Persos(X - 1)
                'Déclare les variables pour les personnages.
                ReDim EffetsPersos(UBound(Persos()))
                For i = 0 To UBound(Persos())
                    Set EffetsPersos(i) = New ClsJeuEffet
                    'If i <> Noperso Then
                        Set Persos(i) = New ClsJeuPerso
                        With Persos(i)
                        .Init i, Parametres
                        End With
                    'End If
                Next i
                ReDim Maisons(UBound(Persos()))
                ReDim EffetsMaisons(UBound(Maisons()))
                For i = 0 To UBound(Maisons())
                    Set Maisons(i) = New ClsJeuBatiment
                    Set EffetsMaisons(i) = New ClsJeuEffet
                    Maisons(i).Numero = i
                Next i
            Case "monde_largeur":
                Monde.Largeur = X
            Case "monde_hauteur":
                Monde.Hauteur = X
            Case "type_terrain"
                Monde.TypeTerrain = X
            Case "epoque"
                For i = 0 To UBound(Fiefs())
                    Fiefs(i).Epoque = X - 1
                Next i
            Case "rangees"
                Rangees = X
            Case "choix_terrain"
                ChoixTerrain = X
            End Select
            'If Temp Like "terrain_Ligne*" Then
            '   'Trace une ligne de terrains.
            '    If Abs(Y - Max) < Abs(X - Min) Then
            '        For i = Min To X
            '            Monde.Param_TerrainCategorie(i + Monde.Largeur * Max) = ChoixTerrain
            '        Next i
            '    Else
            '
            '    End If
            'End If
            If Temp Like "terrain_carrevide*" Then
                'Trace un carré coté par coté.
                For i = Min + 1 To X
                    Monde.Param_TerrainCategorie(i + Monde.Largeur * (Max + 1)) = ChoixTerrain
                    Monde.Param_TerrainCategorie(i + Monde.Largeur * Y) = ChoixTerrain
                Next i
                For i = Max + 1 To Y
                    Monde.Param_TerrainCategorie(Min + 1 + Monde.Largeur * i) = ChoixTerrain
                    Monde.Param_TerrainCategorie(X + Monde.Largeur * i) = ChoixTerrain
                Next i
                Placer_Ressources
            End If
            If Temp Like "terrain_carreplein*" Then
                'Trace un rectangle.
                For i = Min + 1 To X
                    For j = Max + 1 To Y
                        Monde.Param_TerrainCategorie(i + Monde.Largeur * j) = ChoixTerrain
                    Next j
                Next i
                Placer_Ressources
            End If
            If Temp Like "decor*" Then
                Decor_Creer Min, X, Y
                Placer_Ressources
            End If
            If Temp Like "fief*" Then
                For i = Min To Max
                    Select Case Prop
                    Case "Nom": Fiefs(i).Nom = X
                    Case "peuple": Fiefs(i).Definir_TypePeuple = X - 1
                    Case "equipe": Fiefs(i).NumeroEquipe = X
                    Case "epoque": Fiefs(i).Epoque = X - 1
                    Case "maxvie": Chateaux(i).MaxVie = X: Chateaux(i).Vie = X
                    Case "vie": Chateaux(i).Vie = X
                    Case "ressources":
                        For j = 0 To Parametres.NombreRessources - 1
                            If Parametres.Ressources_Epoque(j) <= Fiefs(i).Epoque Then
                                Fiefs(i).Definir_ressources(j) = X
                            End If
                        Next j
                    Case "position"
                        Chateaux(i).PositionX = X
                        Chateaux(i).PositionY = Y
                    End Select
                Next i
            End If
            If Temp Like "perso*" Then
                For i = Min To Max
                    Select Case Prop
                    Case "ia": Persos(i).IA = X
                    Case "ia_equiper": If i <> NoPerso Then Persos(i).IA_Equiper = X
                    Case "ia_manger": If i <> NoPerso Then Persos(i).IA_Manger = X
                    Case "ia_pacifique": If i <> NoPerso Then Persos(i).IA_Pacifique = X
                    Case "ia_temeraire": If i <> NoPerso Then Persos(i).IA_Temeraire = X
                    Case "fichier_ia": If i <> NoPerso Then Persos(i).FichierIA = X
                                       If i <> NoPerso Then Persos(i).ChoisirIA = True
                    Case "fichier_iascript": If i <> NoPerso Then Persos(i).FichierIAScript = X
                                             If i <> NoPerso Then Persos(i).ChoisirIAScript = True
                    Case "position"
                        Persos(i).PositionX = X
                        Persos(i).PositionY = Y
                        Persos(i).DirectionX = X
                        Persos(i).DirectionY = Y
                        Y = Y + IntervalSoldatsY
                        If (i - Min + 1) Mod Rangees = 0 Then
                            Y = Y - IntervalSoldatsY * Rangees
                            X = X + IntervalSoldatsX
                        End If
                    Case "direction": Persos(i).Aller_A Persos(i).PositionX + X, Persos(i).PositionY + Y
                    Case "allerchateau": Persos(i).Aller_Chateau Chateaux(X - 1), X - 1
                    Case "allermaison": Persos(i).Aller_Maison Maisons(X - 1)
                    Case "attaquerperso": Persos(i).Attaquer_Perso X - 1
                    Case "attaquerpersos": Persos(i).Attaquer_Perso X - 1: X = X + 1
                    Case "attaquermaison": Persos(i).Attaquer_Maison X - 1
                    Case "attaquerchateau": Persos(i).Attaquer_Chateau X - 1
                    Case "nom": If i <> NoPerso Then Persos(i).Nom = UCase(Left(X, 1)) & Right(X, Len(X) - 1)
                    Case "chef": Persos(i).NumeroChef = X - 1
                                 Persos(i).ChefDecalageX = Persos(i).PositionX - Persos(Persos(i).NumeroChef).PositionX
                                 Persos(i).ChefDecalageY = Persos(i).PositionY - Persos(Persos(i).NumeroChef).PositionY
                    Case "maxvie": Persos(i).Definir_MaxVie = X: Persos(i).Ajuster_Bonus: Persos(i).Vie = X
                    Case "vie": Persos(i).Vie = X
                    Case "maxenergie": Persos(i).Definir_MaxEnergie = X: Persos(i).Ajuster_Bonus: Persos(i).Energie = X
                    Case "energie": Persos(i).Energie = X
                    Case "maxmagie": Persos(i).Definir_MaxMagie = X: Persos(i).Ajuster_Bonus: Persos(i).Magie = X
                    Case "magie": Persos(i).Magie = X
                    Case "maxmoral": Persos(i).Definir_MaxMoral = X: Persos(i).Ajuster_Bonus: Persos(i).Moral = X
                    Case "moral": Persos(i).Moral = X
                    Case "maxattaque": Persos(i).Definir_MaxAttaque = X: Persos(i).Ajuster_Bonus: Persos(i).Attaque = X
                    Case "attaque": Persos(i).Attaque = X
                    Case "maxdefense": Persos(i).Definir_MaxDefense = X: Persos(i).Ajuster_Bonus: Persos(i).Defense = X
                    Case "defense": Persos(i).Defense = X
                    Case "vitesse": Persos(i).LongueurPas = X: Persos(i).Ajuster_Bonus
                    Case "effetattaque": Persos(i).EffetAttaque = X - 1
                    Case "parole":
                        Persos(i).Parole = X
                    Case "grade":
                        'Persos(i).Changer_Race Persos(i).Race, Parametres, X - 1
                        Tirer_Monstre_Aleatoire i, Fiefs(Persos(i).NumeroFief).Definir_TypePeuple, X - 1
                        Persos(i).Actualiser_Bonus
                    Case "race":
                        Persos(i).Changer_Race_Personnage X - 1, Parametres
                        Persos(i).Actualiser_Bonus
                        'Tirer_Monstre_Aleatoire i, Fiefs(Persos(i).NumeroFief)
                    Case "sexe":  If i <> NoPerso Then Persos(i).Feminin = X
                    Case "fief"
                        Persos(i).NumeroFief = X - 1
                        Persos(i).NumeroEquipe = Fiefs(Persos(i).NumeroFief).NumeroEquipe
                        If i <> NoPerso Then
                            'Persos(i).Changer_Race Fiefs(Persos(i).NumeroFief).TypePeuple, Parametres
                            Tirer_Monstre_Aleatoire i, Fiefs(Persos(i).NumeroFief).Definir_TypePeuple
                        End If
                        Persos(i).Actualiser_Bonus
                        Persos(i).Definir_Attaque = Persos(i).Definir_MaxAttaque
                        Persos(i).Definir_Defense = Persos(i).Definir_MaxDefense
                    Case "dansunchateau"
                        Persos(i).PositionX = Chateaux(X - 1).PositionX + Chateaux(X - 1).Largeur / 2 - Persos(i).Largeur / 2 + Chateaux(X - 1).EntreePositionX
                        Persos(i).PositionY = Chateaux(X - 1).PositionY + Chateaux(X - 1).Hauteur / 2 - Persos(i).Hauteur / 2 + Chateaux(X - 1).EntreePositionY
                        If X - 1 = Persos(i).NumeroFief Then
                            Persos(i).DansSonChateau = True
                        Else
                            Persos(i).DansUnChateau = True
                        End If
                        Persos(i).Arreter
                    Case "argent":
                        Persos(i).Definir_Argent = IIf(Y = 0, X, (X - X * Y / 100) + (Rnd * (X * Y / 100)))
                    Case "ressource":
                        If X = 0 Then
                            For j = 0 To Parametres.NombreRessources - 1
                                If (Parametres.Ressources_Mangeable(j) Or _
                                   Parametres.Ressources_NecessairePourConstruire(j) Or _
                                   Parametres.Batiment_TypeMatierePremiere(Maisons(i).TypeBatiment) = j) And _
                                   Parametres.Ressources_Epoque(j) <= Fiefs(Persos(i).NumeroFief).Epoque Then
                                    Persos(i).Definir_ressources(j) = Persos(i).Definir_ressources(j) + _
                                                                      Rnd * Y
                                End If
                            Next j
                        Else
                            Persos(i).Definir_ressources(X - 1) = Y
                        End If
                    Case "objet":
                        For j = 0 To IIf(Y = 0, 1, Y) - 1
                            Persos(i).Objet_Ajouter_Inventaire X - 1, Parametres, Commentaires, True
                        Next j
                    Case "equipement":
                        For j = 0 To X
                            Perso_Tirer_Equipement_Hasard i, Y
                        Next j
                    End Select
                Next i
            End If
            If Temp Like "maison*" Then
                For i = Min To Max
                    Select Case Prop
                    Case "type"
                        If X = 0 Then
                            Do
                            Y = Int(Rnd * (Parametres.NombreTypeBatiments - 1)) + 1
                            Loop Until Parametres.Batiment_NoEpoque(Y - 1) - 1 <= Fiefs(Persos(i).NumeroFief).Epoque And Not Parametres.Batiment_Empereur(Y - 1) And Not Parametres.Batiment_CreerProprietaire(Y - 1) > 0
                            Y = Y - 1
                        Else
                            Y = Parametres.Batiment_Numero(X)
                        End If
                        'Debug.Print y
                        Do: Loop Until Trouver_Emplacement_Maison(i, PosX, PosY, Y)
                        Maisons(i).Creer PosX, PosY, Y, i, Parametres
                        Persos(i).Arreter
                        Maisons(i).Definir_Vie = Maisons(i).MaxVie
                    Case "position"
                        Maisons(i).PositionX = X
                        Maisons(i).PositionY = Y
                    Case "enfeu"
                        Maisons(i).Definir_Vie = CLng(Rnd * Maisons(i).MaxVie * 0.45)
                    Case "stock"
                        If Maisons(i).MaxStock > 0 Then
                            Maisons(i).Definir_Stock = Int(Rnd * Maisons(i).MaxStock * X / 100)
                            If Maisons(i).Fabrique Then
                                For j = 0 To Maisons(i).Definir_Stock - 1
                                    Y = Int(Rnd * (Maisons(i).Definir_Nombre_Stock_Objet + 1))
                                    Maisons(i).Definir_Stock = Maisons(i).Definir_Stock - 1
                                    Maisons(i).Stock_Objets(Y) = Maisons(i).Stock_Objets(Y) + 1
                                Next j
                            End If
                        End If
                    End Select
                Next i
            End If
            If Temp Like "ressource*" Then
                For i = Min To Max
                    Select Case Prop
                    Case "type"
                        Ressources(i).TypeRessource = Parametres.TerrainRessources_TypeRessource(X - 1)
                        Ressources(i).Nom = Parametres.TerrainRessources_Nom(X - 1)
                        Ressources(i).Apparence = X - 1
                    Case "position"
                        Ressources(i).PositionX = X
                        Ressources(i).PositionY = Y
                    End Select
                Next i
            End If
        Else
            Select Case Temp
            Case "debut"
                Monde.Generer_Monde ChoixTerrain
                Placer_Chateaux
                Placer_Ressources
                IA_Init
                For i = 0 To UBound(Persos())
                    Persos(i).PositionX = CLng(Rnd * (Monde.Largeur * Monde.LargeurCase - Chateaux(0).Largeur))
                    Persos(i).PositionY = CLng(Rnd * (Monde.Hauteur * Monde.HauteurCase - Chateaux(0).Hauteur))
                    Persos(i).Arreter
                Next i
            Case "fin"
                Monde.Reformer_Terrain
            End Select
        End If
    Wend
    
    Close #1
    
    Actualiser_Infos_Fiefs
    
    For i = 0 To UBound(Persos())
        With Persos(i)
        If Persos(i).Invocation Then
            Persos(Persos(i).NumeroChef).NombreSoldats = Persos(Persos(i).NumeroChef).NombreSoldats + 1
            Persos(i).ChoisirIA = False
        End If
        'Corrige les erreurs de caractéristiques éventuels.
        If .Vie > .Definir_MaxVie Then .Vie = .Definir_MaxVie
        If .Energie > .Definir_MaxEnergie Then .Energie = .Definir_MaxEnergie
        If .Magie > .Definir_MaxMagie Then .Magie = .Definir_MaxMagie
        If .Moral > .Definir_MaxMoral Then .Moral = .Definir_MaxMoral
        If .Attaque > .Definir_MaxAttaque Then .Attaque = .Definir_MaxAttaque
        If .Defense > .Definir_MaxDefense Then .Defense = .Definir_MaxDefense
        End With
    Next i
    
    Evenements_Charger Nom, Campagne
End Sub

Public Sub Actualiser_Infos_Fiefs()
    Dim i As Long, j As Long
    For i = 0 To UBound(Fiefs())
        Fiefs(i).NombreCitoyens = 0
        Fiefs(i).NombreMorts = 0
        For j = 0 To UBound(Persos())
            If Persos(j).NumeroFief = i And Not Persos(j).Invocation Then
                Fiefs(i).NombreCitoyens = Fiefs(i).NombreCitoyens + 1
                If Not Persos(j).Vivant Then
                    Fiefs(i).NombreMorts = Fiefs(i).NombreMorts + 1
                End If
            End If
        Next j
    Next i
End Sub

Private Sub Perso_Tirer_Equipement_Hasard(ByVal i As Long, ByVal Chance As Double)
    Dim X As Integer
    If Rnd < Chance / 100 Then
        Do
            X = Int(Rnd * (Parametres.NombreObjets - 1))
        Loop Until Not Parametres.Objet_Secret(X)
        Persos(i).Objet_Ajouter_Inventaire X, Parametres, Commentaires, True
    End If
End Sub

Public Sub Scenario_Editer(ByVal Nom As String)
    Editer_Texte App.Path & CheminScenario & Langues.Dossier & Nom & ExtensionFichiersScenario
End Sub

Public Sub Evenements_Initialiser()
    Evenements = ""
End Sub

Public Sub Evenements_Charger(ByVal Nom As String, _
                              Optional ByVal Campagne As Boolean)
    On Error GoTo Erreur
    Dim FreeNumber As Long
    Dim Buffer As String
    Evenements = ""
    FreeNumber = FreeFile()
    If Campagne Then
        Open App.Path & CheminCampagnes & Langues.Dossier & Nom & ExtensionFichiersScript For Input As #FreeNumber
    Else
        Open App.Path & CheminScenario & Langues.Dossier & Nom & ExtensionFichiersScript For Input As #FreeNumber
    End If
    While Not EOF(FreeNumber)
        Line Input #FreeNumber, Buffer
        Evenements = Evenements & Buffer & vbCrLf
    Wend
    Close #FreeNumber
    
    Set SriptRunner = New MSScriptControl.ScriptControl
    SriptRunner.Language = "VBScript"
    SriptRunner.Reset
    SriptRunner.AllowUI = False

    Exit Sub
Erreur:
End Sub

Public Sub Evenements_Lancer(ByVal Commande As String)
    Dim i As Long
    Dim Perso As Collection, Maison As Collection, Chateau As Collection, Fief As Collection
    If Not ComReseau.Client And Evenements <> "" Then
        Set Perso = New Collection
        Set Maison = New Collection
        Set Chateau = New Collection
        Set Fief = New Collection
        SriptRunner.Reset
        For i = 0 To UBound(Persos())
            Perso.Add Persos(i)
        Next i
        SriptRunner.AddObject "Persos", Perso, True
        For i = 0 To UBound(Maisons())
            Maison.Add Maisons(i)
        Next i
        SriptRunner.AddObject "Maisons", Maison, True
        For i = 0 To UBound(Chateaux())
            Chateau.Add Chateaux(i)
        Next i
        For i = 0 To UBound(Fiefs())
            Fief.Add Fiefs(i)
        Next i
        SriptRunner.AddObject "Chateaux", Chateau, True
        SriptRunner.AddObject "Fiefs", Fief, True
        SriptRunner.AddObject "Tresors", Tresors, True
        SriptRunner.AddObject "Jeu", Jeu, True
        SriptRunner.AddObject "Moteur", FrmMoteur2D, True
        SriptRunner.AddObject "Interface", Interface(0), True
        SriptRunner.AddObject "Parametres", Parametres, True
        SriptRunner.AddObject "Commentaires", Commentaires, True
        SriptRunner.AddObject "Musique", Musique, True
        SriptRunner.AddObject "Langue", Langues, True
        SriptRunner.AddCode Evenements
        SriptRunner.Run Commande
    End If
End Sub
