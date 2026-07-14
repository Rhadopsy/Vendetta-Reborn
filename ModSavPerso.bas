Attribute VB_Name = "ModSavPerso"
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

Const FichierINI = "SavPerso"
Const SectionINI = "SavPerso"

Const CleCRC = "XPBestiaire" 'La clé CRC s'appelle compétence pour qu'on ne voit pas que c'est une clé.

Public CheminSavPerso As String
Public ExtensionFichiersPerso As String
Public SectionPerso As String
Public SectionInterface As String
Public SectionTalents As String

Private HeureDerniereSauvegarde As Double
Public SauvegardePeriodique As Boolean
Private PeriodeSauvegarde As Integer

Public Sub Init_Sauvegarde_Perso()
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    CheminSavPerso = FicIni.Parametre("CheminSavPerso")
    ExtensionFichiersPerso = FicIni.Parametre("ExtensionFichiersPerso")
    FicPer.ExtensionFichierINI = ExtensionFichiersPerso
    SectionPerso = FicIni.Parametre("SectionPerso")
    SectionInterface = FicIni.Parametre("SectionInterface")
    SectionTalents = FicIni.Parametre("SectionTalents")
    
    SauvegardePeriodique = Val(FicIni.Parametre("SauvegardePeriodique"))
    PeriodeSauvegarde = Val(FicIni.Parametre("PeriodeSauvegarde"))
End Sub

Public Sub Charger_Sauvegarde_Perso(ByVal Nom As String, _
                                    ByRef Perso As ClsJeuPerso)
    Dim i As Integer
    
    'Variables mémorisant le pourcentage de caractéritiques.
    Dim CoefVie As Single
    Dim CoefEnergie As Single
    Dim CoefMagie As Single
    Dim CoefMoral As Single
    Dim CoefAttaque As Single
    Dim CoefDefense As Single
    
    'On Local Error GoTo Erreur
    
    With Perso
    'If MaxVie > 0 And MaxEnergie > 0 And MaxMagie > 0 And MaxMoral > 0 And MaxAttaque > 0 And MaxDefense > 0 Then
    CoefVie = .Vie / .Definir_MaxVie
    CoefEnergie = .Energie / .Definir_MaxEnergie
    CoefMagie = .Magie / .Definir_MaxMagie
    CoefMoral = .Moral / .Definir_MaxMoral
    CoefAttaque = .Attaque / .Definir_MaxAttaque
    CoefDefense = .Defense / .Definir_MaxDefense
    'Else
    '    CoefVie = 1
    '    CoefEnergie = 1
    '    CoefMagie = 1
    '    CoefMoral = 1
    '    CoefAttaque = 1
    '    CoefDefense = 1
    'End If

    'Perso.Init Perso.Numero, Parametres
    FicPer.Fichier = CheminSavPerso & Nom
    FicPer.Section = SectionPerso
    .Init_Carac Parametres
    .Changer_Race_Personnage Val(FicPer.Parametre("Race")), Parametres
    .Nom = Nom
      
    .Feminin = FicPer.Parametre("Sexe") = "Féminin"
    
    .NumeroFief = Val(FicPer.Parametre("Equipe"))
    
    .ApparenceChoisie = 0
    .ChoisirApparence = Val(FicPer.Parametre("Apparence")) = 1
    .FichierApparence = FicPer.Parametre("FichierApparence")
    
    '.Nombre_ApparenceBase = AffApparence.NombreCategoriesBase
    '.Nombre_ApparenceBase = Parametres.Race_DefinirNombreCategoriesApparences(.Race)
    For i = 0 To .Nombre_ApparenceBase - 1
        .Definir_FichierApparenceBase(i) = FicPer.Parametre("FichierApparenceBase" & i + 1)
    Next i
    
    .ChoisirIA = Val(FicPer.Parametre("IA")) = 1
    .FichierIA = FicPer.Parametre("FichierIA")
    .ChoisirIAScript = Val(FicPer.Parametre("ChoisirIAScript"))
    .FichierIAScript = FicPer.Parametre("FichierIAScript")
    
    .Attitude = Val(FicPer.Parametre("Attitude"))
    .Formation = Val(FicPer.Parametre("Formation"))
    
    'FicPer.Section = "Expérience des caractéristiques primaires"
    .ExpVie = Val(FicPer.Parametre("Vie"))
    .ExpEnergie = Val(FicPer.Parametre("Energie"))
    .ExpMagie = Val(FicPer.Parametre("Magie"))
    .ExpMoral = Val(FicPer.Parametre("Moral"))
    
    'FicPer.Section = "Expérience des caractéristiques de combat"
    .ExpAttaque = Val(FicPer.Parametre("Attaque"))
    .ExpDefense = Val(FicPer.Parametre("Défense"))
        
    For i = 0 To .Nombre_CompetencesSpeciales - 1
        .Definir_ExpSpeciales(i) = Val(FicPer.Parametre("CompétenceSpéciale" & i + 1))
    Next i
    For i = 0 To .Nombre_CompetencesRessources - 1
        .Definir_ExpRessources(i) = Val(FicPer.Parametre("CompétenceRessource" & i + 1))
    Next i
    For i = 0 To .Nombre_CompetencesServices - 1
        .Definir_ExpServices(i) = Val(FicPer.Parametre("CompétenceService" & i + 1))
    Next i
    For i = 0 To .Nombre_CompetencesObjets - 1
        .Definir_ExpObjets(i) = Val(FicPer.Parametre(Parametres.CompetenceObjet_SAVPerso(i)))
    Next i
    For i = 0 To .Nombre_de_lignes_du_bestiaire - 1
        '.Definir_Bestiaire(i) = Val(FicPer.Parametre("BestiaireMonstre" & i + 1))
        .Definir_Bestiaire(i) = _
            Val(FicPer.Parametre("BestiaireMonstre" & Parametres.Race_IndiceBestiaire(i)))
    Next i
    
    'Charge son équipement spécial.
    'Charger_Sauvegarde_Tresor Perso
    
    .Gestion_Tous_Niveaux
    '.Actualiser_Bonus
    'Perso.Ajuster_Bonus
    
    'Chargement des caractéristiques par défaut.
    .Vie = .Definir_MaxVie * CoefVie
    .Energie = .Definir_MaxEnergie * CoefEnergie
    .Magie = .Definir_MaxMagie * CoefMagie
    .Moral = .Definir_MaxMoral * CoefMoral
    .Attaque = .Definir_MaxAttaque * CoefAttaque
    .Defense = .Definir_MaxDefense * CoefDefense
    
    .Actualiser_Bonus
    
    'Vérifie la clé CRC.
    .Protege = IIf(CStr(PersoToCRC(Perso)) = FicPer.Parametre(CleCRC), True, False)
    'Debug.Print "Charger CRC : " & PersoToCRC(Perso) & " ?= " & FicPer.Parametre(CleCRC)
    'MsgBox .Nom & IIf(.Protege, " est protégé", " n'est pas protégé")
    
    End With
    Exit Sub
Erreur:
    MsgBox "Le fichier de personnage " & Nom & " est invalide", vbCritical
    End
End Sub

Public Sub Enregistrer_Sauvegarde_Perso(ByVal Perso As ClsJeuPerso, Optional ByVal NouveauPerso As Boolean)
    Dim i As Integer
    
    With Perso
    If .Nom <> "" Then
        
        FicPer.ExtensionFichierINI = ExtensionFichiersPerso
        FicPer.Fichier = CheminSavPerso & .Nom
        FicPer.Section = SectionPerso
        If .Feminin Then
            FicPer.Parametre("Sexe") = "Féminin"
        Else
            FicPer.Parametre("Sexe") = "Masculin"
        End If
        
        FicPer.Parametre("Race") = .Race
        'FicPer.Parametre("Equipe") = .NumeroFief
        
        FicPer.Parametre("Apparence") = -CDbl(.ChoisirApparence)
        FicPer.Parametre("FichierApparence") = .FichierApparence
        
        For i = 0 To .Nombre_ApparenceBase - 1
            FicPer.Parametre("FichierApparenceBase" & i + 1) = .Definir_FichierApparenceBase(i)
        Next i
        
        FicPer.Parametre("IA") = -CDbl(.ChoisirIA)
        FicPer.Parametre("FichierIA") = .FichierIA
        FicPer.Parametre("ChoisirIAScript") = -CDbl(.ChoisirIAScript)
        FicPer.Parametre("FichierIAScript") = .FichierIAScript
        
        FicPer.Parametre("Attitude") = .Attitude
        FicPer.Parametre("Formation") = .Formation
        
        FicPer.Parametre("Vie") = Val(.ExpVie)
        FicPer.Parametre("Energie") = Val(.ExpEnergie)
        FicPer.Parametre("Magie") = Val(.ExpMagie)
        FicPer.Parametre("Moral") = Val(.ExpMoral)
        
        FicPer.Parametre("Attaque") = Val(.ExpAttaque)
        FicPer.Parametre("Défense") = Val(.ExpDefense)
        
        For i = 0 To .Nombre_CompetencesSpeciales - 1
            FicPer.Parametre("CompétenceSpéciale" & i + 1) = Val(.Definir_ExpSpeciales(i))
        Next i
        For i = 0 To .Nombre_CompetencesRessources - 1
            FicPer.Parametre("CompétenceRessource" & i + 1) = Val(.Definir_ExpRessources(i))
        Next i
        For i = 0 To .Nombre_CompetencesServices - 1
            FicPer.Parametre("CompétenceService" & i + 1) = Val(.Definir_ExpServices(i))
        Next i
        For i = 0 To .Nombre_CompetencesObjets - 1
            FicPer.Parametre(Parametres.CompetenceObjet_SAVPerso(i)) = Val(Perso.Definir_ExpObjets(i))
        Next i
        For i = 0 To .Nombre_de_lignes_du_bestiaire - 1
            If Not Parametres.Race_FichierMod(i) Then
                FicPer.Parametre("BestiaireMonstre" & Parametres.Race_IndiceBestiaire(i)) = Perso.Definir_Bestiaire(i)
            End If
        Next i
        If NouveauPerso Or .Protege Then
            FicPer.Parametre(CleCRC) = PersoToCRC(Perso)
            'Debug.Print "Enregistrer CRC : " & PersoToCRC(Perso) & " -> " & FicPer.Parametre(CleCRC)
        End If
    End If
    End With
End Sub

Public Sub Charger_Sauvegarde_Tresor(ByVal Perso As ClsJeuPerso)
    'Charge son équipement spécial.
    Dim i As Integer
    
    With Perso
    FicPer.ExtensionFichierINI = ExtensionFichiersPerso
    FicPer.Fichier = CheminSavPerso & .Nom
    
    FicPer.Section = SectionTalents
    '.Objet_Equipes_Vider_Tout
    '.Objet_Inventaire_Vider_Tout
    If Parametres.Race_NombreObjets(.Race) > 0 Then
        For i = 0 To Parametres.Race_NombreObjets(.Race) - 1
            If Val(FicPer.Parametre("Tresor" & i + 1)) = 1 And _
               .Objet_NombreType(Parametres.Race_Objet(.Race, i)) = 0 Then 'Vérifie qu'il n'a pas déjà chargé cet objet.
                .Objet_Ajouter_Inventaire Parametres.Race_Objet(.Race, i), Parametres, ModMain.Commentaires, .IA_Equiper > 0
                .Objets_Utilisation_Automatique Parametres, ModMain.Commentaires
            End If
        Next i
    End If
    End With
End Sub

Public Sub Charger_Tableau_Tresor(ByVal Perso As ClsJeuPerso, _
                                  ByRef Tresors() As Boolean)
    'Renvoie un tableau de l'équipement spécial.
    Dim i As Integer
    
    With Perso
    FicPer.ExtensionFichierINI = ExtensionFichiersPerso
    FicPer.Fichier = CheminSavPerso & .Nom
    
    FicPer.Section = SectionTalents
    ReDim Tresors(Parametres.Race_NombreObjets(.Race))
    If Parametres.Race_NombreObjets(.Race) > 0 Then
        For i = 0 To Parametres.Race_NombreObjets(.Race) - 1
            Tresors(i) = Val(FicPer.Parametre("Tresor" & i + 1))
        Next i
    End If
    End With
End Sub

Public Sub Enregistrer_Sauvegarde_Tresor(ByVal Perso As ClsJeuPerso, _
                                         ByRef Tresors() As Boolean)
    'Enregiste les talents du personnage.
    Dim i As Integer
    
    With Perso
    FicPer.ExtensionFichierINI = ExtensionFichiersPerso
    FicPer.Fichier = CheminSavPerso & .Nom
    FicPer.Section = SectionTalents
    If Parametres.Race_NombreObjets(.Race) > 0 Then
        For i = 0 To Parametres.Race_NombreObjets(.Race) - 1
            FicPer.Parametre("Tresor" & i + 1) = -CDbl(Tresors(i))
        Next i
    End If
    End With
    
    'Charger_Sauvegarde_Tresor Perso
    Charger_Tableau_Tresor Perso, Tresors()
End Sub

Public Sub Charger_Sauvegarde_Interface(ByRef Perso As ClsJeuPerso, _
                                        ByRef Interface As ClsInterface)
    Dim i As Integer
    FicPer.ExtensionFichierINI = ExtensionFichiersPerso
    FicPer.Fichier = CheminSavPerso & Perso.Nom
    FicPer.Section = SectionInterface
    'FicPer.Section = SectionPerso

    'Sauvegarde des récupérations automatiques des barres de statut.
    For i = 0 To Interface.NombreBoutonsBarresStatut - 1
        Interface.AutoRecuperation_Carac(i) = CBool(Val(FicPer.Parametre("AutoRecuperationCarac" & i + 1)))
        If FicPer.Parametre("AutoRecuperationCaracPcent" & i + 1) = "" Then
            Interface.AutoRecuperation_CaracPcent(i) = IIf(i = 0, 1, 0.5)
        Else
            Interface.AutoRecuperation_CaracPcent(i) = Val(FicPer.Parametre("AutoRecuperationCaracPcent" & i + 1)) / 100
            If Interface.AutoRecuperation_CaracPcent(i) <= 0.01 Or _
               Interface.AutoRecuperation_CaracPcent(i) >= 1 Then
                Interface.AutoRecuperation_CaracPcent(i) = IIf(i = 0, 1, 0.5)
            End If
        End If
    Next i
    Interface.AutoBatimentService = CBool(Val(FicPer.Parametre("AutoBatimentService")))
    'Sauvegarde des touches de raccourci des objets.
    For i = 0 To Parametres.NombreObjets - 1
        Interface.Touches_Raccourci_Objets(i) = Val(FicPer.Parametre("ToucheRaccourciObjet" & i + 1)) - 1
    Next i

    'FicPer.Section = SectionInterface
    For i = 0 To Interface.NombreFenetres - 1
        If Val(FicPer.Parametre("Fenetre(" & i + 1 & ")_Deplace")) = 1 Then
            Interface.Fenetre_Deplace(i) = True
            Interface.Fenetre_PositionX(i) = Val(FicPer.Parametre("Fenetre(" & i + 1 & ")_X"))
            Interface.Fenetre_PositionY(i) = Val(FicPer.Parametre("Fenetre(" & i + 1 & ")_Y"))
        End If
    Next i
    
    'Sauvegarde les tris des compétences.
    Interface.CompetencesTri = Val(FicPer.Parametre("CompetencesTri"))
    Interface.CompetencesAfficherToutes = CBool(Val(FicPer.Parametre("CompetencesAfficherToutes")))
End Sub


Public Sub Enregistrer_Sauvegarde_Interface(ByVal Perso As ClsJeuPerso, _
                                            ByVal Interface As ClsInterface)
    Dim i As Integer
    FicPer.ExtensionFichierINI = ExtensionFichiersPerso
    FicPer.Fichier = CheminSavPerso & Perso.Nom
    'FicPer.Section = SectionPerso
    FicPer.Section = SectionInterface
    
    'Charge les récupérations automatiques des barres de statut.
    For i = 0 To Interface.NombreBoutonsBarresStatut - 1
        FicPer.Parametre("AutoRecuperationCarac" & i + 1) = -CDbl(Interface.AutoRecuperation_Carac(i))
        FicPer.Parametre("AutoRecuperationCaracPcent" & i + 1) = Interface.AutoRecuperation_CaracPcent(i) * 100
    Next i
    FicPer.Parametre("AutoBatimentService") = -CDbl(Interface.AutoBatimentService)
    'Charge les touches de raccourci des objets.
    For i = 0 To Parametres.NombreObjets - 1
        FicPer.Parametre("ToucheRaccourciObjet" & i + 1) = Interface.Touches_Raccourci_Objets(i) + 1
    Next i
    
    'FicPer.Section = SectionInterface
    For i = 0 To Interface.NombreFenetres - 1
        If Interface.Fenetre_Deplace(i) Then
            FicPer.Parametre("Fenetre(" & i + 1 & ")_Deplace") = 1
            FicPer.Parametre("Fenetre(" & i + 1 & ")_X") = Interface.Fenetre_PositionX(i)
            FicPer.Parametre("Fenetre(" & i + 1 & ")_Y") = Interface.Fenetre_PositionY(i)
        Else
            FicPer.Parametre("Fenetre(" & i + 1 & ")_Deplace") = 0
        End If
    Next i
    
    'Sauvegarde les tris des compétences.
    FicPer.Parametre("CompetencesTri") = Interface.CompetencesTri
    FicPer.Parametre("CompetencesAfficherToutes") = -CDbl(Interface.CompetencesAfficherToutes)
End Sub

Public Sub Enregistrer_Sauvegarde_Interface_Defaut(ByVal Perso As ClsJeuPerso)
    Dim i As Integer
    FicPer.ExtensionFichierINI = ExtensionFichiersPerso
    FicPer.Fichier = CheminSavPerso & Perso.Nom
    FicPer.Section = SectionInterface
    
    FicPer.Parametre("AutoBatimentService") = -CDbl(True)
    FicPer.Parametre("CompetencesTri") = 2
End Sub

Public Sub Supprimer_Sauvegarde_Perso(ByVal Nom As String)
    On Error GoTo Erreur
    Kill App.Path & "\" & CheminSavPerso & Nom & ExtensionFichiersPerso
    Kill App.Path & "\" & CheminSavPerso & Nom & ".bmp"
    Exit Sub
Erreur:
End Sub

Public Sub Enregistrement_Periodique(ByVal Perso As ClsJeuPerso, ByVal Temps As Double)
    If SauvegardePeriodique Then
        'If ComReseau.Client Then
        If Jeu.Confrontation Then
        Else
            If Abs(Temps - HeureDerniereSauvegarde) > PeriodeSauvegarde Then
                Enregistrer_Sauvegarde_Perso Perso
                HeureDerniereSauvegarde = Temps
                'Debug.Print "Sauvegarde à : " & Temps
            End If
        End If
    End If
End Sub

Public Function Charger_Sauvegarde_Perso_Equipe(ByVal Nom As String) As Long
    Dim i As Integer
    On Local Error GoTo Erreur
    
    FicPer.ExtensionFichierINI = ExtensionFichiersPerso
    FicPer.Fichier = CheminSavPerso & Nom
    FicPer.Section = SectionPerso

    Charger_Sauvegarde_Perso_Equipe = Val(FicPer.Parametre("Equipe"))
    Exit Function
Erreur:
    MsgBox "Le fichier de personnage " & Nom & " est invalide", vbCritical
    End
End Function

Public Sub Changer_Equipe_Perso(ByVal NomPerso As String, _
                                ByVal NumeroEquipe As Integer)
    FicPer.ExtensionFichierINI = ExtensionFichiersPerso
    FicPer.Fichier = CheminSavPerso & NomPerso
    FicPer.Section = SectionPerso
    FicPer.Parametre("Equipe") = NumeroEquipe
End Sub
                              
Public Sub Perso_Finalier(ByVal i As Integer)
    Dim j As Integer
    If Persos(i).ChoisirApparence Then
        If Persos(i).FichierApparence = "" Then
            'Tire une apparence au hasard si le joueur n'en n'a pas.
            Persos(i).FichierApparence = AffPerso.Tirer_Apparence_Au_Hasard(Parametres.Race_CheminApparence(Persos(i).Race, Persos(i).Feminin), Parametres, Parametres.Race_CheminAbsolu(Persos(i).Race))
        End If
        Persos(i).ApparenceChoisie = AffPerso.Charger_Hero(Persos(i).FichierApparence, Persos(i).Feminin)
        Persos(i).Apparence = Persos(i).ApparenceChoisie
    Else
        'Cherche si le personnage peut changer d'apparence.
        For j = 0 To Persos(i).Nombre_Objets_Inventaire - 1
            If Persos(i).Objet_Inventaire_Actif(j) Then
                If Parametres.Objet_ChangerApparencePerso(Persos(i).Objet_Inventaire_Type(j)) Then
                    If Persos(i).FichierApparence = "" Then
                        'Tire une apparence au hasard si le joueur n'en n'a pas.
                        Persos(i).FichierApparence = AffPerso.Tirer_Apparence_Au_Hasard(Parametres.Race_CheminApparence(Persos(i).Race, Persos(i).Feminin), Parametres, Parametres.Race_CheminAbsolu(Persos(i).Race))
                    End If
                    Persos(i).ApparenceChoisie = AffPerso.Charger_Hero(Persos(i).FichierApparence, Persos(i).Feminin)
                    Persos(i).Apparence = Persos(i).ApparenceChoisie
                    Exit For
                End If
            End If
        Next j
    End If
    If Persos(i).Joueur = 0 Then 'Le personnage n'a pas d'apparence choisie.
        AffApparence.Tirer_Apparence_Au_Hasard Persos(i), Parametres
    End If
    AffApparence.Definir_Indices Persos(i)
    If Persos(i).ChoisirIA Then
        Persos(i).NumeroIA = Charger_IA(Persos(i).FichierIA)
        If Persos(i).NumeroIA = -1 Then
            'Si le fichier d'IA n'a pas put être chargé, il est désactivé.
            Persos(i).ChoisirIA = False
        End If
    End If
End Sub

Function PersoToCRC(ByVal Perso As ClsJeuPerso) As Long
    'Crée une clé CRC à partir du fichier personnage entré en paramètre.
    Dim TempCRC As String
    Dim i As Long
    With Perso
    TempCRC = .Nom
    TempCRC = TempCRC & Hex(Val(.ChoisirIA))
    TempCRC = TempCRC & Hex(Val(.Feminin))
    TempCRC = TempCRC & Hex(Val(.ExpVie))
    TempCRC = TempCRC & Hex(Val(.ExpEnergie))
    TempCRC = TempCRC & Hex(Val(.ExpMagie))
    TempCRC = TempCRC & Hex(Val(.ExpMoral))
    'Debug.Print "Exp moral : " & .ExpMoral & " Hex(" & Hex(.ExpMoral); ")"
    TempCRC = TempCRC & Hex(Val(.ExpAttaque))
    TempCRC = TempCRC & Hex(Val(.ExpDefense))
    For i = 0 To .Nombre_CompetencesSpeciales - 1
        If .Definir_ExpSpeciales(i) >= 1 Then
            TempCRC = TempCRC & Hex(Val(.Definir_ExpSpeciales(i)))
        End If
    Next i
    For i = 0 To .Nombre_CompetencesRessources - 1
        If .Definir_ExpRessources(i) >= 1 Then
            TempCRC = TempCRC & Hex(Val(.Definir_ExpRessources(i)))
        End If
    Next i
    For i = 0 To .Nombre_CompetencesServices - 1
        If .Definir_ExpServices(i) >= 1 Then
            TempCRC = TempCRC & Hex(Val(.Definir_ExpServices(i)))
        End If
    Next i
    For i = 0 To .Nombre_CompetencesObjets - 1
        If .Definir_ExpObjets(i) >= 1 Then
            If Not Parametres.CompetenceObjet_FichierMod(i) Then
                TempCRC = TempCRC & Hex(Val(.Definir_ExpObjets(i)))
            End If
        End If
    Next i
    For i = 0 To .Nombre_de_lignes_du_bestiaire - 1
        If .Definir_Bestiaire(i) >= 1 Then
            If Not Parametres.Race_FichierMod(i) Then
                TempCRC = TempCRC & Hex(Val(.Definir_Bestiaire(i)))
            End If
        End If
    Next i
    PersoToCRC = CLng("&h" & CRC16(TempCRC))
    'Debug.Print TempCRC & " -> " & PersoToCRC
    End With
End Function

Function CRC16(Valeur As String) As String
    Dim l As String, S As Long, i As Long, B As Byte, CRC As Long, r As String
    S = 0
    For i = 1 To Len(Valeur) Step 1
        B = CByte(Asc(Mid(Valeur, i, 2)))
        S = S + B
    Next
    CRC = (Not (S Mod &H10000)) + 1
    CRC16 = Right$(Hex(CRC), 2) & Mid$(Hex(CRC), 5, 2)
End Function

'Propriétés
Public Property Get Definir_SAVPersoPeriodique() As Boolean
    Definir_SAVPersoPeriodique = SauvegardePeriodique
End Property
Public Property Let Definir_SAVPersoPeriodique(ByVal Valeur As Boolean)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    SauvegardePeriodique = Valeur
    FicIni.Parametre("SauvegardePeriodique") = -Int(SauvegardePeriodique)
End Property

