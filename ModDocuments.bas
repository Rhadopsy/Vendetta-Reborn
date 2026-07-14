Attribute VB_Name = "ModDocuments"
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

Const FichierINI = "Documents"
Const SectionINI = "Documents"

Private EnregistrerStastistiques As Boolean
Private AfficherStastistiquesEnQuittant As Boolean

Private MessageFichierManquantTexte As String
Private MessageFichierManquantTitre As String

Private FichierDocumentChemin As String
Private FichierDocumentExtension As String

Private FichierDocumentBatiments As String
Private FichierDocumentBestiaire As String
Private FichierDocumentObjets As String
Private FichierDocumentRessources As String
Private FichierDocumentPeuples As String
Private FichierDocumentStastistiques As String
Private FichierDocumentPerso As String
Private FichierDocumentResume As String
Private FichierDocumentArtworks As String

Private PageDebut  As String
Private PageFin  As String
Private LigneDebut  As String
Private LigneFin  As String
Private TexteGrasDebut  As String
Private TexteGrasFin  As String

Private RetourLigne  As String
Private Separateur  As String

Private NombreColonnes As Integer
Private LigneCourante As Integer
Private ColonneCourante As Integer

Public Sub Init_Documents()
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    EnregistrerStastistiques = Val(FicIni.Parametre("EnregistrerStastistiques"))
    AfficherStastistiquesEnQuittant = Val(FicIni.Parametre("AfficherStastistiquesEnQuittant"))

    MessageFichierManquantTexte = FicIni.Parametre("MessageFichierManquantTexte")
    MessageFichierManquantTitre = FicIni.Parametre("MessageFichierManquantTitre")

    FichierDocumentChemin = FicIni.Parametre("FichierDocumentChemin")
    FichierDocumentExtension = FicIni.Parametre("FichierDocumentExtension")

    FichierDocumentBatiments = FicIni.Parametre("FichierDocumentBatiments")
    FichierDocumentBestiaire = FicIni.Parametre("FichierDocumentBestiaire")
    FichierDocumentObjets = FicIni.Parametre("FichierDocumentObjets")
    FichierDocumentRessources = FicIni.Parametre("FichierDocumentRessources")
    FichierDocumentPeuples = FicIni.Parametre("FichierDocumentPeuples")
    FichierDocumentStastistiques = FicIni.Parametre("FichierDocumentStastistiques")
    FichierDocumentPerso = FicIni.Parametre("FichierDocumentPerso")
    FichierDocumentResume = FicIni.Parametre("FichierDocumentResume")
    FichierDocumentArtworks = FicIni.Parametre("FichierDocumentArtworks")

    PageDebut = FicIni.Parametre("PageDebut")
    PageFin = FicIni.Parametre("PageFin")
    LigneDebut = FicIni.Parametre("LigneDebut")
    LigneFin = FicIni.Parametre("LigneFin")
    TexteGrasDebut = FicIni.Parametre("TexteGrasDebut")
    TexteGrasFin = FicIni.Parametre("TexteGrasFin")

    RetourLigne = FicIni.Parametre("RetourLigne")
    Separateur = FicIni.Parametre("Separateur")
    
    NavigateurChoisi = Val(FicIni.Parametre("NavigateurChoisi")) - 1
    CheminNotePad = FicIni.Parametre("CheminNotePad")
    CheminExplorer = FicIni.Parametre("CheminExplorer")
    LogicielIExplorer = FicIni.Parametre("LogicielIExplorer")
    CheminIExplorer = FicIni.Parametre("CheminIExplorer")
    LogicielFireFox = FicIni.Parametre("LogicielFireFox")
    CheminFireFox = FicIni.Parametre("CheminFireFox")
End Sub

Public Sub Document_Batiments()
    Dim i As Integer, j As Integer, k As Integer
    Dim Temp As Boolean
    Dim TempEpoque As Integer
    Dim Etiquette As String
    
    Dim Fichier As String
    
    Fichier = App.Path & FichierDocumentChemin & FichierDocumentBatiments & FichierDocumentExtension
    Definir_Nombre_Colonnes = 8
    'Génère le fichier HTML des bâtiments.
    Open Fichier For Output As #1
    
    Print #1, PageDebut
    Afficher_Case UCase(Parametres.EtiquetteEpoque), True
    'Afficher_Case UCase(Parametres.EtiquetteImage), True
    Afficher_Case UCase(Parametres.EtiquetteNom), True
    Afficher_Case UCase(Parametres.EtiquettePrix), True
    Afficher_Case UCase(Parametres.EtiquetteMetier & " " & Left(Parametres.EtiquetteMasculin, 1) & " / " & Left(Parametres.EtiquetteFeminin, 1)), True
    Afficher_Case UCase(Parametres.EtiquetteTaille), True
    Afficher_Case UCase(Parametres.EtiquetteVie), True
    Afficher_Case UCase(Parametres.EtiquetteStock), True
    Afficher_Case UCase(Parametres.EtiquetteFabrique), True, False
    For k = 0 To Parametres.NombreTypeBatiments - 1
        i = Parametres.Batiment_Tri(k)
        'Affiche l'époque.
        If Parametres.Batiment_NoEpoque(i) <= Parametres.NombreEpoques Then
            Etiquette = Parametres.Epoque_Nom(Parametres.Batiment_NoEpoque(i) - 1)
        Else
            Etiquette = UCase(Parametres.EtiquetteSecret)
        End If
        If Parametres.Batiment_NoEpoque(i) > TempEpoque Then
            Afficher_Case Etiquette
            TempEpoque = Parametres.Batiment_NoEpoque(i)
        Else
            Afficher_Case ""
        End If
        
        'Affiche l'image.
        'Afficher_Case "<img src=" & Chr(34) & "Images/BoutonsConstruire/Construire" & Format(i, "000") & ".bmp width=28 height=28 >"
        
        'Affiche le nom.
        Afficher_Case Parametres.Batiment_Nom(i)
        
        'Affiche le coût en ressources.
        If Parametres.Batiment_PrixArgent(i) > 0 Then
            Etiquette = Parametres.Batiment_PrixArgent(i) & " " & Parametres.EtiquetteArgent
            Temp = True
        Else
            Etiquette = ""
            Temp = False
        End If
        For j = 0 To Parametres.NombreRessources - 1
            If Parametres.Batiment_PrixRessource(i, j) > 0 Then
                If Temp Then
                    Etiquette = Etiquette & RetourLigne
                End If
                Etiquette = Etiquette & _
                            Parametres.Batiment_PrixRessource(i, j) & " " & _
                            Parametres.Ressources_Nom(j)
                Temp = True
            End If
        Next j
        If Etiquette = "" Then Etiquette = "-"
        Afficher_Case Etiquette
        
        'Affiche le métier.
        Afficher_Case Parametres.Batiment_Metier(i, False) & " / " & Parametres.Batiment_Metier(i, True)
        
        'Affiche les dimensions.
        Afficher_Case Parametres.Batiment_Largeur(i) & "*" & Parametres.Batiment_Hauteur(i)
        
        'Affiche les points de vie.
        Afficher_Case Parametres.Batiment_MaxVie(i)
        
        'Affiche le stock
        If Parametres.Batiment_MaxStockVente(i) = 0 Then
            Afficher_Case "-"
        Else
            Afficher_Case Parametres.Batiment_MaxStockVente(i)
        End If
                
        'Affiche ce que le bâtiment fabrique.
        If Parametres.Batiment_RessourceFabrique(i) >= 0 Then
            Afficher_Case Parametres.Ressources_Nom(Parametres.Batiment_RessourceFabrique(i)), , False
        ElseIf Parametres.Batiment_Fabrique(i) Then
            Temp = False: Etiquette = ""
            For j = 0 To Parametres.Batiment_NombreObjetsFabriques(i) - 1
                If Temp Then
                    Etiquette = Etiquette & Separateur
                End If
                Etiquette = Etiquette & _
                            Parametres.Objet_Nom(Parametres.Batiment_ObjetFabrique(i, j))
                Temp = True
            Next j
            Afficher_Case Etiquette, , False
        ElseIf Parametres.Batiment_Service(i) > 0 Then
            Afficher_Case Parametres.Service_Nom(Parametres.Batiment_Service(i)), , False
        Else
            Afficher_Case Parametres.Batiment_Description(i), , False
        End If
    Next k
    Print #1, PageFin
    
    Close #1
    
    Ouvir_PageWeb Fichier
End Sub

Public Sub Document_Bestiaire()
    Dim i As Integer, j As Integer, X As Integer
    Dim Temp As Boolean
    Dim Temp2 As Integer
    Dim Etiquette As String
    
    Dim Fichier As String
    
    Fichier = App.Path & FichierDocumentChemin & FichierDocumentBestiaire & FichierDocumentExtension
    Definir_Nombre_Colonnes = 8
    'Génère le fichier HTML des monstres.
    Open Fichier For Output As #1
    
    Print #1, PageDebut
    Afficher_Case UCase(Parametres.EtiquetteNom), True
    Afficher_Case UCase(Parametres.EtiquetteVie), True
    Afficher_Case UCase(Parametres.EtiquetteAttaque), True
    Afficher_Case UCase(Parametres.EtiquetteDefense), True
    Afficher_Case UCase(Parametres.EtiquetteVitesse), True
    Afficher_Case UCase(Parametres.EtiquetteEffetAttaque), True
    Afficher_Case UCase(Parametres.EtiquetteXP), True
    Afficher_Case UCase(Parametres.EtiquetteParticularite), True, False
    With Parametres
    For i = 0 To .NombreRaces - 1
        X = .Bestiaire_Ordre_Alpha(i)
        
        If Not .Race_Inattaquable(X) Then
            Afficher_Case .Race_Nom(X)
            
            Afficher_Case Int(.Race_Vie(X) * .Race_CoefVie(X))
            
            Afficher_Case Int(.Race_Attaque(X) * .Race_CoefAttaque(X)) '& _
                          " (" & Int(.Race_CoefAttaque(x) * 100) & "%)"
            
            Afficher_Case Int(.Race_Defense(X) * .Race_CoefDefense(X)) '& _
                          " (" & Int(.Race_CoefDefense(x) * 100) & "%)"
            
            Afficher_Case .Race_Vitesse(X)
            
            If .Race_EffetAttaque(X) > 0 Then
                Afficher_Case .Effet_Nom(.Race_EffetAttaque(X)) & IIf(Parametres.Race_AttaqueResistance(X) > 0, RetourLigne & "(" & .Resistance_Attaque(Parametres.Race_AttaqueResistance(X)) & ")", "")
            Else
                Afficher_Case "-"
            End If
            
            If .Race_Experience(X) Then
                Afficher_Case "X"
            Else
                Afficher_Case "-"
            End If
            
            Temp = False: Etiquette = ""
            If .Race_Volant(X) Then
                If Temp Then
                    Etiquette = Etiquette & Separateur
                End If
                Etiquette = Etiquette & Parametres.EtiquetteVolant
                Temp = True
            End If
            
            If .Race_PorteeAttaque(X) > 0 Then
                If Temp Then
                    Etiquette = Etiquette & Separateur
                End If
                Etiquette = Etiquette & _
                            Parametres.EtiquettePorteeAttaque & " : " & .Race_PorteeAttaque(X)
                Temp = True
            End If
            
            If .Race_Armure(X) > 0 Then
                If Temp Then
                    Etiquette = Etiquette & Separateur
                End If
                Etiquette = Etiquette & Parametres.EtiquetteArmure & " : " & .Race_Armure(X)
                Temp = True
            End If
            If .Race_Pietinement(X) > 0 Then
                If Temp Then
                    Etiquette = Etiquette & Separateur
                End If
                Etiquette = Etiquette & Parametres.EtiquettePietinement & " : " & .Race_Pietinement(X)
                Temp = True
            End If
            If .Race_Infatigable(X) Then
                If Temp Then
                    Etiquette = Etiquette & Separateur
                End If
                Etiquette = Etiquette & Parametres.EtiquetteInfatiguable
                Temp = True
            End If
            For j = 0 To .NombreRessources - 1
                If .Race_Ressources(X, j) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & Separateur
                    End If
                    Etiquette = Etiquette & _
                                Int(.Race_Ressources(X, j)) & " " & _
                                Parametres.Ressources_Nom(j)
                    Temp = True
                End If
            Next j
            If .Race_NombreObjets(X) > 0 Then
                For j = 0 To .Race_NombreObjets(X) - 1
                    If Parametres.Objet_PrixTalents(.Race_Objet(X, j)) = 0 Then
                        If Temp Then
                            Etiquette = Etiquette & Separateur
                        End If
                        Etiquette = Etiquette & _
                                    Parametres.Objet_Nom(.Race_Objet(X, j))
                        Temp = True
                    End If
                Next j
            End If
            If .Race_NombreEquipement(X) > 0 Then
                For j = 0 To .Race_NombreEquipement(X) - 1
                    If Temp Then
                        Etiquette = Etiquette & Separateur
                    End If
                    Etiquette = Etiquette & _
                                Parametres.Objet_Nom(.Race_Equipement(X, j))
                    Temp = True
                Next j
            End If
            If .Race_VitesseRegenerationVie(X) > 0 Then
                If Temp Then
                    Etiquette = Etiquette & Separateur
                End If
                Etiquette = Etiquette & Parametres.EtiquetteRegeneration
                Temp = True
            End If
            If .Race_Parole(X) Then
                If Temp Then
                    Etiquette = Etiquette & Separateur
                End If
                Etiquette = Etiquette & Parametres.EtiquetteParole
                Temp = True
            End If
            If .Race_Equiper(X) Then
                If Temp Then
                    Etiquette = Etiquette & Separateur
                End If
                Etiquette = Etiquette & Parametres.EtiquetteSEquipe
                Temp = True
            End If
            If .Race_IA_Manger(X) Then
                If Temp Then
                    Etiquette = Etiquette & Separateur
                End If
                Etiquette = Etiquette & Parametres.EtiquetteMange
                Temp = True
            End If
            If .Race_IA_Pacifique(X) Then
                If Temp Then
                    Etiquette = Etiquette & Separateur
                End If
                Etiquette = Etiquette & Parametres.EtiquettePacifique
                Temp = True
            End If
            If .Race_IA_Temeraire(X) Then
                If Temp Then
                    Etiquette = Etiquette & Separateur
                End If
                Etiquette = Etiquette & Parametres.EtiquetteTemeraire
                Temp = True
            End If
            If .Race_IA_Berserk(X) = 1 Then
                If Temp Then
                    Etiquette = Etiquette & Separateur
                End If
                Etiquette = Etiquette & Parametres.EtiquetteBerserker
                Temp = True
            End If
            If .Race_AutoResurrection(X) Then
                If Temp Then
                    Etiquette = Etiquette & Separateur
                End If
                Etiquette = Etiquette & Parametres.EtiquetteRessuscite
                Temp = True
            End If
            If .Race_Kamikaze(X) Then
                If Temp Then
                    Etiquette = Etiquette & Separateur
                End If
                Etiquette = Etiquette & Parametres.EtiquetteKamikaze
                Temp = True
            End If
            If .Race_IA_Fichier(X) <> "" Then
                If Temp Then
                    Etiquette = Etiquette & Separateur
                End If
                Etiquette = Etiquette & _
                            .EtiquetteIA & ":" & Chr(34) & .Race_IA_Fichier(X) & Chr(34)
                Temp = True
            End If
            'Les résistances.
            For j = 0 To .NombreResistances - 1
                If .Race_Resistances(X, j) <> .ResistanceDefaut Then
                    If Temp Then
                        Etiquette = Etiquette & Separateur
                    End If
                    Etiquette = Etiquette & _
                                .Resistance_NomComplet(j) & " : " & IIf(.Race_Resistances(X, j) < .ResistanceDefaut, "+", "") & Format(-(.Race_Resistances(X, j) - 1) * 100, "0") & .EtiquettePourcent
                    Temp = True
                End If
            Next j
            
            If Etiquette = "" Then Etiquette = "-"
            Afficher_Case Etiquette, , False
        End If
    Next i
    End With
    Print #1, PageFin
    
    Close #1
    
    Ouvir_PageWeb Fichier
End Sub

Public Sub Document_Objets()
    Dim i As Integer, j As Integer, k As Integer, l As Integer, m As Integer
    Dim Temp As Boolean
    Dim Temp2 As Integer
    Dim Etiquette As String
    
    Dim Fichier As String
    
    Fichier = App.Path & FichierDocumentChemin & FichierDocumentObjets & FichierDocumentExtension
    Definir_Nombre_Colonnes = 9
    'Génère le fichier HTML des monstres.
    Open Fichier For Output As #1
    
    Print #1, PageDebut
    Afficher_Case UCase(Parametres.EtiquetteCompetence), True
    Afficher_Case UCase(Parametres.EtiquetteImage), True
    Afficher_Case UCase(Parametres.EtiquetteNom), True
    Afficher_Case UCase(Parametres.EtiquetteDisponibilite), True
    Afficher_Case UCase(Parametres.EtiquetteCout), True
    Afficher_Case UCase(Parametres.EtiquettePrix), True
    Afficher_Case UCase(Parametres.EtiquetteBonus), True
    Afficher_Case UCase(Parametres.EtiquetteEffet), True
    Afficher_Case UCase(Parametres.EtiquetteParticularite), True, False
    'Afficher_Case "DESCRIPTION", True, False
    Temp2 = -1
    For l = 0 To Parametres.NombreCompetencesObjets - 1
        k = Parametres.CompetenceObjet_NoListe(l)
        For i = 0 To Parametres.NombreObjets - 1
            'If (Not Parametres.Objet_Secret(i) Or CheckAfficherSecret.Value = 1) And _
               Parametres.Objet_Competence(i) = k Then
            If Parametres.Objet_Competence(i) = k Then
                Etiquette = Parametres.CompetenceObjet_Nom(Parametres.Objet_Competence(i))
                If k <> Temp2 Then
                    Afficher_Case Etiquette
                    Temp2 = k
                Else
                    Afficher_Case ""
                End If
                
                If Parametres.Objet_CheminImage(i) <> "" Then
                    Afficher_Case "<img src=" & Chr(34) & Parametres.Objet_CheminImage(i) & Parametres.Objet_Image(i) & ".bmp" & Chr(34) & " width=32 height=" & IIf(Parametres.Objet_Duree(i) <> 0, "64", "32") & " >"
                Else
                    Afficher_Case "<img src=" & Chr(34) & "Images/Objets/" & Parametres.Objet_Image(i) & ".bmp" & Chr(34) & " width=32 height=" & IIf(Parametres.Objet_Duree(i) <> 0, "64", "32") & " >"
                End If
                
                Afficher_Case Parametres.Objet_Nom(i) & IIf(Parametres.Objet_NomEffet(i) <> "", RetourLigne & "(" & Parametres.Objet_NomEffet(i) & ")", "")
                
                If Parametres.Objet_Unique(i) Then
                    Afficher_Case Parametres.EtiquetteTresor
                ElseIf Parametres.Objet_Secret(i) Then
                    Afficher_Case "-"
                ElseIf Parametres.Objet_Epoque(i) - 1 >= 0 Then
                    Afficher_Case Parametres.Epoque_Nom(Parametres.Objet_Epoque(i) - 1)
                Else
                    Afficher_Case "-"
                End If
                
                Etiquette = ""
                For j = 0 To Parametres.NombreRessources - 1
                    If Parametres.Objet_PrixRessources(i, j) > 0 Then
                        If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                        Etiquette = Etiquette & _
                                         Parametres.Objet_PrixRessources(i, j) & " " & _
                                         Parametres.Ressources_Nom(j)
                    End If
                Next j
                If Etiquette = "" Then Etiquette = "-"
                Afficher_Case Etiquette
                
                'Affiche le prix.
                If Parametres.Objet_PrixVente(i) > 0 Then
                    Afficher_Case Parametres.Objet_PrixVente(i)
                Else
                    Afficher_Case "-"
                End If
                
                Etiquette = ""
                If Parametres.Objet_BonusVie(i) <> 0 Then
                    'If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                    If Parametres.Objet_BonusVie(i) > 0 Then Etiquette = Etiquette & "+"
                    Etiquette = Etiquette & " " & Parametres.Objet_BonusVie(i) & " " & Parametres.EtiquetteVie
                End If
                
                If Parametres.Objet_BonusEnergie(i) <> 0 Then
                    If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                    If Parametres.Objet_BonusEnergie(i) > 0 Then Etiquette = Etiquette & "+"
                    Etiquette = Etiquette & " " & Parametres.Objet_BonusEnergie(i) & " " & Parametres.EtiquetteEnergie
                End If
                
                If Parametres.Objet_BonusMagie(i) <> 0 Then
                    If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                    If Parametres.Objet_BonusMagie(i) > 0 Then Etiquette = Etiquette & "+"
                    Etiquette = Etiquette & " " & Parametres.Objet_BonusMagie(i) & " " & Parametres.EtiquetteMagie
                End If
                
                If Parametres.Objet_BonusMoral(i) <> 0 Then
                    If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                    If Parametres.Objet_BonusMoral(i) > 0 Then Etiquette = Etiquette & "+"
                    Etiquette = Etiquette & " " & Parametres.Objet_BonusMoral(i) & " " & Parametres.EtiquetteMoral
                End If
                
                If Parametres.Objet_BonusAttaque(i) <> 0 Then
                    If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                    If Parametres.Objet_BonusAttaque(i) > 0 Then Etiquette = Etiquette & "+"
                    Etiquette = Etiquette & " " & Parametres.Objet_BonusAttaque(i) & " " & Parametres.EtiquetteAttaque
                End If
                If Parametres.Objet_DureeRechargeAttaque(i) > 0 Then
                    If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                    Etiquette = Etiquette & Parametres.EtiquetteDureeRechargeAttaque & " : " & Parametres.Objet_DureeRechargeAttaque(i)
                End If
                If Parametres.Objet_AttaqueResistance(i) > 0 Then
                    If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                    Etiquette = Etiquette & "(" & Parametres.Resistance_Attaque(Parametres.Objet_AttaqueResistance(i)) & ")"
                End If
                If Parametres.Objet_AttaqueCauseEffet(i) > 0 Then
                    If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                    Etiquette = Etiquette & "(" & Parametres.Objet_Nom(Parametres.Objet_AttaqueCauseEffet(i)) & " : " & Parametres.Objet_Duree(Parametres.Objet_AttaqueCauseEffet(i)) & Parametres.EtiquetteUniteDuree & ")"
                End If
                
                
                If Parametres.Objet_BonusDefense(i) <> 0 Then
                    If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                    If Parametres.Objet_BonusDefense(i) > 0 Then Etiquette = Etiquette & "+"
                    Etiquette = Etiquette & " " & Parametres.Objet_BonusDefense(i) & " " & Parametres.EtiquetteDefense
                End If
                
                If Parametres.Objet_BonusPorteeAttaque(i) > 0 Then
                    If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                    Etiquette = Etiquette & Parametres.EtiquettePortee & " : " & Parametres.Objet_BonusPorteeAttaque(i) & " " & Parametres.EtiquetteUniteDistance
                End If
                
                If Parametres.Objet_PorteeUtilisation(i) > 0 Then
                    If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                    Etiquette = Etiquette & Parametres.EtiquettePortee & " : " & Parametres.Objet_PorteeUtilisation(i) & " " & Parametres.EtiquetteUniteDistance
                End If
                
                If Etiquette = "" Then Etiquette = "-"
                Afficher_Case Etiquette
                
                If Parametres.Objet_EffetAttaque(i) > 0 Then
                    Etiquette = Parametres.Effet_Nom(Parametres.Objet_EffetAttaque(i))
                    Afficher_Case Etiquette
                Else
                    Afficher_Case "-"
                End If
                
                Temp = False: Etiquette = ""
                If Parametres.Objet_Monture(i) Then
                    Etiquette = Etiquette & Parametres.EtiquetteMonture
                    Temp = True
                End If
                If Parametres.Objet_BonusArmure(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & "+ " & _
                                    Parametres.Objet_BonusArmure(i) & " " & LCase(Parametres.EtiquetteArmure)
                    Temp = True
                End If
                If Parametres.Objet_BonusDeplacement(i) <> 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & IIf(Parametres.Objet_BonusDeplacement(i) > 0, "+ ", "") & _
                                    Parametres.Objet_BonusDeplacement(i) & " " & LCase(Parametres.EtiquetteDeplacement)
                    Temp = True
                End If
                If Parametres.Objet_BonusVitesse(i) <> 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    If Parametres.Objet_BonusVitesse(i) > 0 Then Etiquette = Etiquette & "+"
                    Etiquette = Etiquette & " " & _
                                    Parametres.Objet_BonusVitesse(i) & " " & LCase(Parametres.EtiquetteVitesse)
                    Temp = True
                End If
                If Parametres.Objet_BonusVitalite(i) <> 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    If Parametres.Objet_BonusVitalite(i) > 0 Then Etiquette = Etiquette & "+"
                    Etiquette = Etiquette & " " & _
                                    Parametres.Objet_BonusVitalite(i) * 100 & Parametres.EtiquettePourcent & " " & LCase(Parametres.EtiquetteVitalite)
                    Temp = True
                End If
                For j = 0 To Parametres.NombreCompetencesSpeciales - 1
                    If Parametres.Objet_BonusCompetencesSpeciales(i, j) > 0 Then
                        If Temp Then
                            Etiquette = Etiquette & RetourLigne
                        End If
                        Etiquette = Etiquette & "+ " & _
                                        Parametres.Objet_BonusCompetencesSpeciales(i, j) & " " & _
                                        Parametres.Speciales_NomCompetences(j)
                        Temp = True
                    End If
                Next j
                For j = 0 To Parametres.NombreRessources - 1
                    If Parametres.Objet_BonusCompetencesRessources(i, j) > 0 Then
                        If Temp Then
                            Etiquette = Etiquette & RetourLigne
                        End If
                        Etiquette = Etiquette & "+ " & _
                                         Parametres.Objet_BonusCompetencesRessources(i, j) & " " & _
                                         Parametres.Ressources_NomCompetence(j)
                        Temp = True
                    End If
                Next j
                For j = 0 To Parametres.NombreServices - 2
                    If Parametres.Objet_BonusCompetencesServices(i, j) > 0 Then
                        If Temp Then
                            Etiquette = Etiquette & RetourLigne
                        End If
                        Etiquette = Etiquette & "+ " & _
                                        Parametres.Objet_BonusCompetencesServices(i, j) & " " & _
                                        Parametres.Service_NomCompetence(j + 1)
                        Temp = True
                    End If
                Next j
                For j = 1 To Parametres.NombreCompetencesObjets - 1
                    m = Parametres.CompetenceObjet_NoListe(j)
                    If Parametres.Objet_BonusCompetencesObjets(i, m) > 0 Then
                        If Temp Then
                            Etiquette = Etiquette & RetourLigne
                        End If
                        Etiquette = Etiquette & "+ " & _
                                        Parametres.Objet_BonusCompetencesObjets(i, m) & " " & _
                                        Parametres.CompetenceObjet_Nom(m)
                        Temp = True
                    End If
                Next j
                If Parametres.Objet_SurfaceAttaque(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & _
                                               Parametres.EtiquetteSurfaceAttaque & " : " & _
                                               Parametres.Objet_SurfaceAttaque(i)
                    Temp = True
                End If
                If Parametres.Objet_Invocation(i) >= 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & _
                                Parametres.EtiquetteInvoque & " : " & _
                                Parametres.Objet_Quantite(i) & _
                                " " & Parametres.Race_Nom(Parametres.Objet_Invocation(i), , Parametres.Objet_Quantite(i) >= 2)
                    Temp = True
                    'Limite d'invocations.
                    If Parametres.Objet_QuantiteMax(i) > 0 Then
                        If Temp Then
                            Etiquette = Etiquette & RetourLigne
                        End If
                        Etiquette = Etiquette & _
                                    IIf(Parametres.Objet_QuantiteMax(i) >= 2, Parametres.EtiquetteInvoqueMax1Pluriel, Parametres.EtiquetteInvoqueMax1) & _
                                    Parametres.Objet_QuantiteMax(i) & " " & LCase(Parametres.Race_Nom(Parametres.Objet_Invocation(i), , Parametres.Objet_Quantite(i) >= 2)) & _
                                    IIf(Parametres.Objet_QuantiteMax(i) >= 2, Parametres.EtiquetteInvoqueMax2Pluriel, Parametres.EtiquetteInvoqueMax2)
                        Temp = True
                    End If
                End If
                
                If Parametres.Objet_PerteVie(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & "- " & _
                                               Parametres.Objet_PerteVie(i) & " " & Parametres.EtiquetteVie & " " & Parametres.EtiquetteParUtilisation
                    Temp = True
                End If
                If Parametres.Objet_PerteEnergie(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & "- " & _
                                               Parametres.Objet_PerteEnergie(i) & " " & Parametres.EtiquetteEnergie & " " & Parametres.EtiquetteParUtilisation
                    Temp = True
                End If
                If Parametres.Objet_PerteMagie(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & "- " & _
                                               Parametres.Objet_PerteMagie(i) & " " & Parametres.EtiquetteMagie & " " & Parametres.EtiquetteParUtilisation
                    Temp = True
                End If
                If Parametres.Objet_PerteMoral(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & "- " & _
                                               Parametres.Objet_PerteMoral(i) & " " & Parametres.EtiquetteMoral & " " & Parametres.EtiquetteParUtilisation
                    Temp = True
                End If
                If Parametres.Objet_PerteArgent(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & "- " & _
                                               Parametres.Objet_PerteArgent(i) & " " & Parametres.EtiquetteArgent & " " & Parametres.EtiquetteParUtilisation
                    Temp = True
                End If
                If Parametres.Objet_PerteRessourceQuantite(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & "- " & _
                                               Parametres.Objet_PerteRessourceQuantite(i) & " " & _
                                               Parametres.Ressources_Nom(Parametres.Objet_PerteRessourceIndice(i)) & _
                                               " " & Parametres.EtiquetteParUtilisation
                    Temp = True
                End If
                'Nombre de projectiles à la fois.
                If Parametres.Objet_ProjectileSupplementaire(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & Parametres.Objet_ProjectileSupplementaire(i) + 1 & _
                                            " " & Parametres.EtiquetteProjectilesSupplementaires
                    Temp = True
                End If
                If Parametres.Objet_Poison(i) >= 1 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & Parametres.EtiquettePoison & " : " & _
                                            Parametres.Objet_Poison(i) & " "
                    Temp = True
                End If
                If Parametres.Objet_GainVie(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & _
                                               "+ " & _
                                               Parametres.Objet_GainVie(i) & _
                                               " " & LCase(Parametres.EtiquetteVie)
                    Temp = True
                End If
                If Parametres.Objet_GainEnergie(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & _
                                               "+ " & _
                                               Parametres.Objet_GainEnergie(i) & _
                                               " " & LCase(Parametres.EtiquetteEnergie)
                    Temp = True
                End If
                If Parametres.Objet_GainMagie(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & _
                                               "+ " & _
                                               Parametres.Objet_GainMagie(i) & _
                                               " " & LCase(Parametres.EtiquetteMagie)
                    Temp = True
                End If
                If Parametres.Objet_GainMoral(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & _
                                               "+ " & _
                                               Parametres.Objet_GainMoral(i) & _
                                               " " & LCase(Parametres.EtiquetteMoral)
                    Temp = True
                End If
                'Génération de ressources.
                If Parametres.Objet_GainArgent(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & _
                                    Parametres.EtiquetteGenere & " " & _
                                    Parametres.Objet_GainArgent(i) & " argent"
                    Temp = True
                End If
                For j = 0 To Parametres.NombreRessources - 1
                    If Parametres.Objet_GainRessource(i, j) > 0 Then
                        If Temp Then
                            Etiquette = Etiquette & RetourLigne
                        End If
                        Etiquette = Etiquette & _
                                        Parametres.EtiquetteGenere & " " & _
                                        Parametres.Objet_GainRessource(i, j) & " " & _
                                        Parametres.Ressources_Nom(j)
                        Temp = True
                    End If
                Next j
                
                'Resurrection.
                If Parametres.Objet_Resurrection(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & Parametres.EtiquetteRessuscite & " " & _
                                               Parametres.Objet_Quantite(i) & _
                                               " " & _
                                               IIf(Parametres.Objet_Quantite(i) >= 2, Parametres.EtiquettePersonnes, Parametres.EtiquettePersonne)
                    Temp = True
                End If
                If Parametres.Objet_SelfResurrection(i) Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & Parametres.EtiquetteRessusciteSoi
                    Temp = True
                End If
                
                'Berserk
                If Parametres.Objet_Berserk(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & Parametres.EtiquetteBerserk
                    Temp = True
                End If
                
                'Bonus de transport de ressources.
                If Parametres.Objet_BonusMaxRessources(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & Parametres.EtiquetteTranport & " " & _
                                               Parametres.Objet_BonusMaxRessources(i) & _
                                               " " & Parametres.EtiquetteRessourcesSupplementaires
                    Temp = True
                End If
                
                'Bonus de régénération.
                If Parametres.Objet_BonusVitesseRegenerationVie(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & Parametres.EtiquetteRegenere & " " & _
                                               Format(Parametres.Objet_BonusVitesseRegenerationVie(i), "0.000000000") * 1000 & _
                                               Parametres.EtiquettePourcent & " " & Parametres.EtiquetteVie
                    Temp = True
                End If
                If Parametres.Objet_BonusVitesseRegenerationEnergie(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & Parametres.EtiquetteRegenere & " " & _
                                               Format(Parametres.Objet_BonusVitesseRegenerationEnergie(i), "0.000000000") * 1000 & _
                                               Parametres.EtiquettePourcent & " " & Parametres.EtiquetteEnergie
                    Temp = True
                End If
                If Parametres.Objet_BonusVitesseRegenerationMagie(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & Parametres.EtiquetteRegenere & " " & _
                                               Format(Parametres.Objet_BonusVitesseRegenerationMagie(i), "0.000000000") * 1000 & _
                                               Parametres.EtiquettePourcent & " " & Parametres.EtiquetteMagie
                    Temp = True
                End If
                If Parametres.Objet_BonusVitesseRegenerationMoral(i) > 0 Then
                    If Temp Then
                        Etiquette = Etiquette & RetourLigne
                    End If
                    Etiquette = Etiquette & Parametres.EtiquetteRegenere & " " & _
                                               Format(Parametres.Objet_BonusVitesseRegenerationMoral(i), "0.000000000") * 1000 & _
                                               Parametres.EtiquettePourcent & " " & Parametres.EtiquetteMoral
                    Temp = True
                End If
                
                'Informations sur les détails de son utilisation.
                If Parametres.Objet_UtilisationUnique(i) Then
                    If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                    Etiquette = Etiquette & Parametres.EtiquetteUtilisationUnique
                    Temp = True
                End If
                If Parametres.Objet_UtilisationMort(i) Then
                    If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                    Etiquette = Etiquette & Parametres.EtiquetteUtilisationMort
                    Temp = True
                End If
                
                'Indestructible.
                If Parametres.Objet_Indestructible(i) Then
                    If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                    Etiquette = Etiquette & Parametres.EtiquetteIndestructible
                    Temp = True
                End If
                
                'Résistances.
                For j = 0 To Parametres.NombreResistances - 1
                    If Parametres.Objet_BonusResistance(i, j) <> 0 Then
                        If Temp Then Etiquette = Etiquette & RetourLigne
                        Etiquette = Etiquette & _
                                    Parametres.Resistance_NomComplet(j) & " : " & IIf(Parametres.Objet_BonusResistance(i, j) > 0, "+", "") & Parametres.Objet_BonusResistance(i, j) * 100 & Parametres.EtiquettePourcent
                        Temp = True
                    End If
                Next j
                
                'Durée.
                If Parametres.Objet_Duree(i) > 0 Then
                    If Temp Then Etiquette = Etiquette & RetourLigne
                    Etiquette = Etiquette & _
                                Parametres.EtiquetteDuree & " : " & _
                                Int(Parametres.Objet_Duree(i)) & " " & Parametres.EtiquetteUniteDuree
                    Temp = True
                End If
                
                'Durée de rechargement.
                If Parametres.Objet_DureeRecharge(i) > 0 Then
                    If Temp Then Etiquette = Etiquette & RetourLigne
                    Etiquette = Etiquette & _
                                Parametres.EtiquetteDureeRecharge & " : " & _
                                Int(Parametres.Objet_DureeRecharge(i)) & " " & Parametres.EtiquetteUniteDuree
                    Temp = True
                    If Parametres.Objet_RechargesMultiples(i) > 0 Then
                        If Temp Then Etiquette = Etiquette & RetourLigne
                        Etiquette = Etiquette & Parametres.Objet_RechargesMultiples(i) & _
                                    " " & IIf(Parametres.Objet_RechargesMultiples(i) >= 2, Parametres.EtiquetteRechargesMultiples, Parametres.EtiquetteRechargeMultiple)
                        Temp = True
                    End If
                End If
                
                If Etiquette = "" Then Etiquette = "-"
                Afficher_Case Etiquette, False, False
                
                'Description.
                'If Parametres.Objet_Description(i) <> "" Then
                '    Afficher_Case Parametres.Objet_Description(i), , False
                'Else
                '    Afficher_Case "-", , False
                'End If
            End If
        Next i
    Next l
    Print #1, PageFin
    
    Close #1
    
    Ouvir_PageWeb Fichier
End Sub

Public Sub Document_Ressources()
    Dim i As Integer, j As Integer
    Dim Temp As Boolean
    Dim Etiquette As String
    
    Dim Fichier As String
    
    Fichier = App.Path & FichierDocumentChemin & FichierDocumentRessources & FichierDocumentExtension
    Definir_Nombre_Colonnes = 9
    'Génère le fichier HTML des ressources.
    Open Fichier For Output As #1
    
    Print #1, PageDebut
    Afficher_Case UCase(Parametres.EtiquetteImage), True
    Afficher_Case UCase(Parametres.EtiquetteNom), True
    Afficher_Case UCase(Parametres.EtiquetteCompetence), True
    Afficher_Case UCase(Parametres.EtiquetteMatieresPremieres), True
    Afficher_Case UCase(Parametres.EtiquetteValeur), True
    Afficher_Case UCase(Parametres.EtiquetteVitesseRecolte), True
    Afficher_Case UCase(Parametres.EtiquetteRegain), True
    Afficher_Case UCase(Parametres.EtiquetteParticularite), True, False
    Afficher_Case UCase(Parametres.EtiquetteRessourceObjets), True, False
    For i = 0 To Parametres.NombreRessources - 1
        If Parametres.Ressources_CheminImage(i) <> "" Then
            Afficher_Case "<img src=" & Chr(34) & Parametres.Ressources_CheminImage(i) & Parametres.Ressources_Image(i) & ".bmp" & Chr(34) & " width=32 height=32>"
        Else
            Afficher_Case "<img src=" & Chr(34) & "Images/Ressources/" & Parametres.Ressources_Image(i) & ".bmp" & Chr(34) & " width=32 height=32>"
        End If
    
        Afficher_Case Parametres.Ressources_Nom(i)
        
        Afficher_Case Parametres.Ressources_NomCompetence(i)
        
        If Parametres.Ressources_TypeMatierePremiere(i) < 0 Then
            Etiquette = "-"
        Else
            Etiquette = Parametres.Ressources_Nom(Parametres.Ressources_TypeMatierePremiere(i))
        End If
        If Parametres.Ressources_QuantiteMatierePremiere(i) <> 1 Then
            Etiquette = Etiquette & "(*" & Parametres.Ressources_QuantiteMatierePremiere(i) & ")"
        End If
        Afficher_Case Etiquette
        
        Afficher_Case Parametres.Ressources_PrixVente(i)
        
        Afficher_Case Parametres.Ressources_VitesseExtraction(i) * 100 & " " & Parametres.EtiquettePourcent
        
        If Parametres.Ressources_GainVie(i) > 0 Then
            Etiquette = "+ " & Parametres.Ressources_GainVie(i) & " " & Parametres.EtiquetteVie
        Else
            Etiquette = ""
        End If
        
        If Parametres.Ressources_GainEnergie(i) > 0 Then
            If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
            Etiquette = Etiquette & "+ " & Parametres.Ressources_GainEnergie(i) & " " & Parametres.EtiquetteEnergie
        End If
        
        If Parametres.Ressources_GainMagie(i) > 0 Then
            If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
            Etiquette = Etiquette & "+ " & Parametres.Ressources_GainMagie(i) & " " & Parametres.EtiquetteMagie
        End If
        
        If Parametres.Ressources_GainMoral(i) > 0 Then
            If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
            Etiquette = Etiquette & "+ " & Parametres.Ressources_GainMoral(i) & " " & Parametres.EtiquetteMoral
        End If
        If Etiquette = "" Then Etiquette = "-"
        Afficher_Case Etiquette
        
        Etiquette = ""
        For j = 0 To Parametres.NombreTerrainRessource - 1
            If Parametres.TerrainRessources_TypeRessource(j) = i Then
                If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                Etiquette = Etiquette & Parametres.EtiquetteRessourceOrigine & " : " & Parametres.TerrainRessources_Nom(j)
            End If
        Next j
        If Etiquette = "" Then Etiquette = "-"
        Afficher_Case Etiquette, , False
        
        'Objets utilisés pour booster la compétence.
        Etiquette = ""
        For j = 0 To Parametres.NombreObjets - 1
            If Parametres.Objet_BonusCompetencesRessources(j, i) > 0 Then
                If Etiquette <> "" Then Etiquette = Etiquette & Separateur
                Etiquette = Etiquette & Parametres.Objet_Nom(j)
            End If
        Next j
        If Etiquette = "" Then Etiquette = "-"
        Afficher_Case Etiquette, , False
    Next i
    Print #1, PageFin
    
    Close #1
    
    Ouvir_PageWeb Fichier
End Sub

Public Sub Document_Peuples()
    Dim i As Integer, j As Integer, k As Integer
    Dim Temp As Boolean
    Dim Etiquette As String
    Dim Chance As Single
    
    Dim Fichier As String
    
    Fichier = App.Path & FichierDocumentChemin & FichierDocumentPeuples & FichierDocumentExtension
    Definir_Nombre_Colonnes = 4
    'Génère le fichier HTML des ressources.
    Open Fichier For Output As #1
    
    Print #1, PageDebut
    Afficher_Case UCase(Parametres.EtiquetteNom), True
    Afficher_Case UCase(Parametres.EtiquetteForteresse), True
    Afficher_Case UCase(Parametres.EtiquetteCreatures), True
    Afficher_Case UCase(Parametres.EtiquetteIA), True
    With Parametres
    For i = 0 To .NombrePeuples - 1
    
        Afficher_Case .Peuples_Nom(i)
        
        Afficher_Case .Peuples_NomChateau(i)
        
        Temp = False
        If .Peuples_NombreMonstres(i) > 0 Then
            For j = 0 To .Peuples_NombreMonstres(i) - 1
                If Temp Then
                    Afficher_Case ""
                    Afficher_Case ""
                End If
                Temp = True
                Afficher_Case .Race_Nom(.Peuples_MonstreIndice(i, j), , True)
                
                'For k = 0 To j
                '   Select Case k
                '   Case 0:
                '       Chance = .Peuples_MonstreChance(i, k)
                '    Case Else:
                '        If Chance < 0.5 Then
                '            Chance = (1 - Chance) * .Peuples_MonstreChance(i, k)
                '        Else
                '            Chance = .Peuples_MonstreChance(i, k) * Chance
                '        End If
                '    End Select
                'Next k
                'Afficher_Case IIf(Chance < 0.01, Format(Chance, "0.0"), Int(Chance * 100)) & " %"
                
                Etiquette = .Peuples_MonstreFichierIA(i, j)
                If Etiquette = "" Then
                    Etiquette = .Race_IA_Fichier(.Peuples_MonstreIndice(i, j))
                End If
                Afficher_Case IIf(Etiquette = "", "-", Etiquette)
            Next j
        Else
            For j = 0 To NombreColonnes - 3
                Afficher_Case ""
            Next j
        End If
    Next i
    End With
    Print #1, PageFin
    
    Close #1
    
    Ouvir_PageWeb Fichier
End Sub


Public Sub Document_Stastistiques()
    Dim i As Integer, j As Integer, k As Integer, X As Integer
    Dim Temp As Boolean
    Dim TempEpoque As Integer
    Dim Etiquette As String
    
    Dim Fichier As String
    
    If EnregistrerStastistiques Then
    Fichier = App.Path & FichierDocumentChemin & FichierDocumentStastistiques & FichierDocumentExtension
    Definir_Nombre_Colonnes = 7
    'Génère le fichier HTML des stastistiques de fin de partie.
    Open Fichier For Output As #1
    
    Print #1, PageDebut
    
    'Print #1, "<B>" & "Partie du " & Format(Date, "DDDD") & " " & Format(Date, "dd") & " " & Format(Date, "MMMM") & " " & Format(Date, "YYYY") & "<B>"
    Print #1, "<B>" & Format(Date, "DDDD") & " " & Format(Date, "dd") & " " & Format(Date, "MMMM") & " " & Format(Date, "YYYY") & "<B>"
    Print #1, RetourLigne
    Print #1, RetourLigne
    
    Afficher_Case UCase(Parametres.EtiquettePersonnages), True
    Afficher_Case UCase(Parametres.EtiquetteFrags), True
    Afficher_Case UCase(Parametres.EtiquetteMorts), True
    Afficher_Case UCase(Parametres.EtiquetteNivPlus), True
    Afficher_Case UCase(Parametres.EtiquetteGainCaracs), True, False
    Afficher_Case UCase(Parametres.EtiquetteGainCompetences), True, False
    Afficher_Case UCase(Parametres.EtiquetteBestiaire), True, False
    For i = 0 To Stastistiques.NombrePersos - 1
        'Affiche le nom.
        Afficher_Case Stastistiques.Nom(i)
        
        'Affiche le score.
        If Stastistiques.Frags(i, 0) > 0 Then
            Etiquette = Stastistiques.Frags(i, 0) & " " & IIf(Stastistiques.Frags(i, 0) > 1, Parametres.EtiquettePersonnages, Parametres.EtiquettePersonnage)
        Else
            Etiquette = ""
        End If
        If Stastistiques.Frags(i, 1) > 0 Then
            If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
            Etiquette = Etiquette & Stastistiques.Frags(i, 1) & " " & IIf(Stastistiques.Frags(i, 1) > 1, Parametres.EtiquetteBatiments, Parametres.EtiquetteBatiment)
        End If
        If Stastistiques.Frags(i, 2) > 0 Then
            If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
            Etiquette = Etiquette & Stastistiques.Frags(i, 2) & " " & IIf(Stastistiques.Frags(i, 2) > 1, Parametres.EtiquetteChateaux, Parametres.EtiquetteChateau)
        End If
        If Etiquette = "" Then Etiquette = "-"
        Afficher_Case Etiquette
        
        Etiquette = Stastistiques.Morts(i)
        If Etiquette = "0" Then Etiquette = "-"
        Afficher_Case Etiquette
        
        'Affiche l'augmentation de niveau.
        Etiquette = Format(Stastistiques.Gain_Niveau(i), "0.00")
        If Stastistiques.Gain_Tresor(i) >= 1 Then
            Etiquette = Etiquette & RetourLigne & "(+" & Stastistiques.Gain_Tresor(i) & " " & LCase(Parametres.EtiquetteTresor) & ")"
        End If
        Afficher_Case Etiquette
        
        'Affiche les caractéristiques.
        Etiquette = ""
        If Stastistiques.Gain_Vie(i) > 0 Then
            If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
            Etiquette = Etiquette & "+ " & Stastistiques.Gain_Vie(i) & " " & Parametres.EtiquetteVie
        End If
        If Stastistiques.Gain_Energie(i) > 0 Then
            If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
            Etiquette = Etiquette & "+ " & Stastistiques.Gain_Energie(i) & " " & Parametres.EtiquetteEnergie
        End If
        If Stastistiques.Gain_Magie(i) > 0 Then
            If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
            Etiquette = Etiquette & "+ " & Stastistiques.Gain_Magie(i) & " " & Parametres.EtiquetteMagie
        End If
        If Stastistiques.Gain_Moral(i) > 0 Then
            If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
            Etiquette = Etiquette & "+ " & Stastistiques.Gain_Moral(i) & " " & Parametres.EtiquetteMoral
        End If
        If Stastistiques.Gain_Attaque(i) > 0 Then
            If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
            Etiquette = Etiquette & "+ " & Stastistiques.Gain_Attaque(i) & " " & Parametres.EtiquetteAttaque
        End If
        If Stastistiques.Gain_Defense(i) > 0 Then
            If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
            Etiquette = Etiquette & "+ " & Stastistiques.Gain_Defense(i) & " " & Parametres.EtiquetteDefense
        End If
        If Etiquette = "" Then Etiquette = "-"
        Afficher_Case Etiquette, , False
        
        'Affiche les compétences.
        Etiquette = ""
        For j = 0 To Parametres.NombreCompetencesSpeciales - 1
            If Stastistiques.Gain_CompetenceSpeciales(i, j) > 0 Then
                If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                Etiquette = Etiquette & "+ " & Stastistiques.Gain_CompetenceSpeciales(i, j) & " " & _
                            Parametres.Speciales_NomCompetences(j)
            End If
        Next j
        For j = 0 To Parametres.NombreRessources - 1
            If Stastistiques.Gain_CompetenceRessources(i, j) > 0 Then
                If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                Etiquette = Etiquette & "+ " & Stastistiques.Gain_CompetenceRessources(i, j) & " " & _
                            Parametres.Ressources_NomCompetence(j)
            End If
        Next j
        For j = 0 To Parametres.NombreServices - 2
            If Stastistiques.Gain_CompetenceServices(i, j) > 0 Then
                If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                Etiquette = Etiquette & "+ " & Stastistiques.Gain_CompetenceServices(i, j) & " " & _
                            Parametres.Service_NomCompetence(j + 1)
            End If
        Next j
        For j = 1 To Parametres.NombreCompetencesObjets - 1
            k = Parametres.CompetenceObjet_NoListe(j)
            If Stastistiques.Gain_CompetenceObjets(i, k) > 0 Then
                If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                Etiquette = Etiquette & "+ " & Stastistiques.Gain_CompetenceObjets(i, k) & " " & _
                            Parametres.CompetenceObjet_Nom(k)
            End If
        Next j
        If Etiquette = "" Then Etiquette = "-"
        Afficher_Case Etiquette, , False
        
        'Affiche le bestiaire.
        Etiquette = ""
        For j = 0 To Parametres.NombreRaces - 1
            X = Parametres.Bestiaire_Ordre_Alpha(j)
            If Stastistiques.Gain_Bestiaire(i, X) > 0 Then
                If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
                Etiquette = Etiquette & Stastistiques.Gain_Bestiaire(i, X) & " " & _
                            Parametres.Race_Nom(X, , Stastistiques.Gain_Bestiaire(i, X) >= 2)
            End If
        Next j
        If Etiquette = "" Then Etiquette = "-"
        Afficher_Case Etiquette, , False
        
    Next i
    Print #1, PageFin
    
    Close #1
    
    End If
End Sub

Public Sub Afficher_Stastistiques()
    On Error GoTo Erreur
    Dim Fichier As String
    Fichier = App.Path & FichierDocumentChemin & FichierDocumentStastistiques & FichierDocumentExtension
    Open Fichier For Input As #1
    Close #1
    Ouvir_PageWeb Fichier
    Exit Sub
Erreur:
    MsgBox MessageFichierManquantTexte, vbCritical, MessageFichierManquantTitre
End Sub

Public Sub Document_Perso(ByVal Perso As ClsJeuPerso)
    Dim i As Integer, j As Integer
    Dim Etiquette As String
    
    Dim Fichier As String
    
    Fichier = App.Path & FichierDocumentChemin & FichierDocumentPerso & Perso.Nom & FichierDocumentExtension
    Definir_Nombre_Colonnes = 1
    'Génère le fichier HTML des stastistiques de fin de partie.
    Open Fichier For Output As #1
    
    Print #1, PageDebut
    
    Etiquette = UCase(Parametres.EtiquetteNom) & " : " & UCase(Perso.Nom)
    Afficher_Case Etiquette, True, False
    
    Etiquette = Parametres.EtiquetteNiveau & " : " & FrmParam.LblNiveau & RetourLigne & _
                Parametres.EtiquetteRace & " : " & Parametres.Race_Nom(Perso.Race, Perso.Feminin) & RetourLigne & _
                Parametres.EtiquetteMetier & " : " & Perso.Metier(Parametres)
    Afficher_Case Etiquette, , False
    
    Afficher_Case "CARACTERISTIQUES", True, False
    Etiquette = Parametres.EtiquetteVie & " : " & Perso.Definir_MaxVie & RetourLigne & _
                Parametres.EtiquetteEnergie & " : " & Perso.Definir_MaxEnergie & RetourLigne & _
                Parametres.EtiquetteMagie & " : " & Perso.Definir_MaxMagie & RetourLigne & _
                Parametres.EtiquetteMoral & " : " & Perso.Definir_MaxMoral & RetourLigne & _
                Parametres.EtiquetteAttaque & " : " & Perso.Definir_MaxAttaque & RetourLigne & _
                Parametres.EtiquetteDefense & " : " & Perso.Definir_MaxDefense & RetourLigne
    
    Afficher_Case Etiquette, , False
    
    'Affiche les compétences.
    Afficher_Case UCase(Parametres.EtiquetteCompetences), True, False
    Etiquette = ""
'    For i = 0 To Parametres.NombreCompetencesSpeciales - 1
'        If Perso.Niveau_CompetenceSpeciales(i) > 0 Then
'            Etiquette = Etiquette & Parametres.Speciales_NomCompetences(i) & " : " & _
'                        Perso.Definir_Niveau_CompetenceSpeciales(i) & RetourLigne
'        End If
'    Next i
'    For i = 0 To Parametres.NombreRessources - 1
'        If Perso.Niveau_CompetenceRessources(i) > 0 Then
'            Etiquette = Etiquette & Parametres.Ressources_NomCompetence(i) & " : " & _
'                        Perso.Definir_Niveau_CompetenceRessources(i) + Parametres.PersosMaxRessources & RetourLigne
'        End If
'    Next i
'    For i = 0 To Perso.Nombre_CompetencesServices - 1
'        If Perso.Niveau_CompetenceServices(i) > 0 Then
'            Etiquette = Etiquette & Parametres.Service_NomCompetence(i + 1) & " : " & _
'                        Perso.Definir_Niveau_CompetenceServices(i) & RetourLigne
'        End If
'    Next i
'    For i = 1 To Parametres.NombreCompetencesObjets - 1
'        j = Parametres.CompetenceObjet_NoListe(i)
'        If Perso.Niveau_CompetenceObjets(j) > 0 Then
'            Etiquette = Etiquette & Parametres.CompetenceObjet_Nom(j) & " : " & _
'                        Perso.Definir_Niveau_CompetenceObjets(j) & RetourLigne
'        End If
'    Next i
'    If Etiquette = "" Then
'        Etiquette = "-"
'    Else
'        Etiquette = Etiquette & "MOYENNE : " & Format(Perso.Definir_Moy_Competence(MDIFrmMain.AffichageCompetences), "0.00")
'    End If
    With FrmParam.LstCompetences
    For i = 0 To .ListCount - 1
        If Left(.List(i), 1) <> "-" Then
            Etiquette = Etiquette & .List(i)
            If i < .ListCount - 1 Then Etiquette = Etiquette & RetourLigne
        End If
    Next i
    End With
    Afficher_Case Etiquette, , False
    
    'Affiche le bestiaire.
    Afficher_Case UCase(Parametres.EtiquetteBestiaire), True, False
    Etiquette = ""
    For i = 0 To Parametres.NombreRaces - 1
        j = Parametres.Bestiaire_Ordre_Alpha(i)
        If Perso.Definir_Bestiaire(j) > 0 Then
            If Etiquette <> "" Then Etiquette = Etiquette & RetourLigne
            Etiquette = Etiquette & _
                        Parametres.Race_Nom(j, , Perso.Definir_Bestiaire(j) > 1) & " : " & Perso.Definir_Bestiaire(j)
            
        End If
    Next i
    If Etiquette = "" Then
        Etiquette = "-"
    Else
        Etiquette = Etiquette & RetourLigne & UCase(Parametres.EtiquetteTotal) & " : " & Perso.Definir_Bestiaire_Total
    End If
    Afficher_Case Etiquette, , False
    
    Print #1, PageFin
    
    Close #1
    
    Ouvir_PageWeb Fichier
End Sub

Public Sub Document_Resume()
    Dim i As Long, j As Long
    Dim Fichier As String
    
    Fichier = App.Path & FichierDocumentChemin & FichierDocumentResume & FichierDocumentExtension
    'Génère le fichier HTML résumé du jeu.
    Open Fichier For Output As #1
    
    Afficher_Debut
    
    With Parametres
    'Titre.
    Afficher_Ligne App.Title & " " & _
                   App.Major & "." & _
                   App.Minor, True
    'Informations.
    Afficher_Ligne
    Afficher_Ligne .NombreEpoques & " " & .EtiquetteEpoque & "s"
    Afficher_Ligne .NombreRacesJouables & " " & Majuscule(.EtiquetteRacesJouables)
    Afficher_Ligne .NombreTypeBatimentsJouables & " " & .EtiquetteBatiment & "s"
    Afficher_Ligne
    Afficher_Ligne .NombreCompetencesSpeciales + .NombreRessources + .NombreServices - 1 + .NombreCompetencesObjets & " " & .EtiquetteCompetences
    Afficher_Ligne
    Afficher_Ligne .NombrePeuples & " " & .EtiquettePeuple & "s"
    For i = 0 To .NombreRaces - 1
        If Not .Race_Inattaquable(i) Then
            j = j + 1
        End If
    Next i
    
    Afficher_Ligne j - .NombreRacesJouables & " " & .EtiquetteMonstre & "s"
    Afficher_Ligne
    Afficher_Ligne .NombreRessources & " " & Majuscule(.EtiquetteRessources)
    Afficher_Ligne .NombreObjets & " " & Majuscule(.EtiquetteObjets)
    End With
    
    Afficher_Fin
    
    Close #1
    
    Ouvir_PageWeb Fichier
End Sub

Public Sub Document_Artworks()
    Dim i As Long, j As Long
    Dim Fichier As String
    
    Fichier = App.Path & FichierDocumentChemin & FichierDocumentArtworks & FichierDocumentExtension
    'Génère le fichier HTML résumé du jeu.
    Open Fichier For Output As #1
    
    Afficher_Debut
    
    Dim Temp As String
    
    Afficher_Ligne "<body bgcolor= & Chr(34) & #000000 & Chr(34) &  text= & Chr(34) & #A8B892 & Chr(34) & link= & Chr(34) & #FFCC99 & Chr(34) &  vlink= & Chr(34) & #CD853F & Chr(34) &  alink= & Chr(34) & #C0C0C0 & Chr(34) &  style= & Chr(34) & font-family: Arial & Chr(34) & >"
    Afficher_Ligne "<div align=center>"
    Temp = Dir(App.Path & "\Images\Chargement\", vbNormal)
    While Temp <> ""
        Afficher_Ligne "<img src=" & Chr(34) & "Images\Chargement\" & Temp & Chr(34) & ">"
        'Afficher_Ligne "<font color= & Chr(34) & #FF6666 & Chr(34) & >"
        Afficher_Ligne Temp, True
        'Afficher_Ligne "</font>"
        Afficher_Ligne
        Temp = Dir()
    Wend
    Afficher_Ligne "</div>"
    Afficher_Ligne "</body>"
    
    Afficher_Fin
    
    Close #1
    
    Ouvir_PageWeb Fichier
End Sub

Public Function Majuscule(ByVal Texte As String) As String
    'Renvoie le texte avec une majuscule à la première lettre
    Majuscule = UCase(Left(Texte, 1)) & Right(Texte, Len(Texte) - 1)
End Function

Private Sub Afficher_Debut()
    Print #1, PageDebut
End Sub
Private Sub Afficher_Fin()
    Print #1, PageFin
End Sub

Private Sub Afficher_Ligne(Optional ByVal Texte As String, _
                           Optional ByVal Gras As Boolean)
    Print #1, IIf(Gras, TexteGrasDebut, "") & Texte & RetourLigne & IIf(Gras, TexteGrasFin, "")
End Sub

Private Sub Afficher_Case(ByVal Texte As String, _
                          Optional ByVal Gras As Boolean = False, _
                          Optional ByVal Centre As Boolean = True)
    If LigneCourante = 0 Then
        Print #1, "<table width=100% border=1 cellspacing=0;>"
        LigneCourante = LigneCourante + 1
    End If
    If ColonneCourante = 0 Then
        Print #1, LigneDebut
    End If
    If Gras Then
        Texte = TexteGrasDebut & Texte & TexteGrasFin
    End If
    If Centre Then
        Texte = "<td align=center>" & Texte & "</td>"
    Else
        Texte = "<td>" & Texte & "</td>"
    End If
    Print #1, Texte
    ColonneCourante = ColonneCourante + 1
    If ColonneCourante = NombreColonnes Then
        ColonneCourante = 0
        LigneCourante = LigneCourante + 1
        Print #1, LigneFin
    End If
End Sub

Private Property Let Definir_Nombre_Colonnes(ByVal Valeur As Integer)
    NombreColonnes = Valeur
    LigneCourante = 0
    ColonneCourante = 0
End Property

'Propriétés
Public Property Get Definir_EnregistrerStastistiques() As Boolean
    Definir_EnregistrerStastistiques = EnregistrerStastistiques
End Property
Public Property Let Definir_EnregistrerStastistiques(ByVal Valeur As Boolean)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    EnregistrerStastistiques = Valeur
    FicIni.Parametre("EnregistrerStastistiques") = -CDbl(Valeur)
End Property

Public Property Get Definir_AfficherStastistiquesEnQuittant() As Boolean
    Definir_AfficherStastistiquesEnQuittant = AfficherStastistiquesEnQuittant
End Property
Public Property Let Definir_AfficherStastistiquesEnQuittant(ByVal Valeur As Boolean)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    AfficherStastistiquesEnQuittant = Valeur
    FicIni.Parametre("AfficherStastistiquesEnQuittant") = -CDbl(Valeur)
End Property

Public Property Get Definir_NavigateurWeb() As Integer
    Definir_NavigateurWeb = NavigateurChoisi
End Property
Public Property Let Definir_NavigateurWeb(ByVal Valeur As Integer)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    NavigateurChoisi = Valeur
    FicIni.Parametre("NavigateurChoisi") = Valeur + 1
End Property
