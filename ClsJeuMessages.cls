VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsJeuMessages"
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

Const FichierINI = "Messages"
Const SectionINI = "Messages"

Private FichierConseil As String
Private PrefixeConseil As String

Private Afficher As Boolean
Private NombreMaxMessages As Integer

Private AfficherConseil As Boolean
Private AfficherVie As Boolean
Private AfficherBestiaire As Boolean
Private AfficherArgent As Boolean
Private AfficherRessources As Boolean
Private AfficherObjets As Boolean
Private AfficherSoldats As Boolean
Private AfficherExperience As Boolean
Private AfficherEmpereur As Boolean

Public Duree As Double
Public Couleur As Long
Public PositionX As Integer
Public PositionY As Integer
Public DecalageX As Integer
Public DecalageY As Integer

Private PrefixeMessage As String
Private PrefixeExperience As String

Private MessageVivant As String
Private MessageMort As String
Private MessageArgent As String
Private MessageGain As String
Private MessagePerte As String
Private MessageBestiaire As String
Private MessageVie As String
Private MessageEnergie As String
Private MessageMagie As String
Private MessageMoral As String
Private MessageAttaque As String
Private MessageDefense As String
Private MessageSoldat As String
Private MessageSoldats As String
Private MessageCompetence As String
Private MessageEmpereurVieM  As String
Private MessageEmpereurVieF  As String
Private MessageEmpereurMortM  As String
Private MessageEmpereurMortF  As String

Private MessageSauvegardeMonde As String

Private Numero As Long
Private Vivant As Boolean
Private Argent As Double
Private Ressources() As Double
Private Objets() As Integer
Private Bestiaire() As Double
Private Soldats As Long
Private MaxVie As Long
Private MaxEnergie As Long
Private MaxMagie As Long
Private MaxMoral As Long
Private MaxAttaque As Long
Private MaxDefense As Long
Private NivSpeciales() As Long
Private NivRessources() As Long
Private NivServices() As Long
Private NivObjets() As Long
Private Empereur As Long 'Numéro de l'empereur.

Private AffMessage As ClsIntMessage
Public Messages As Collection

Private TempsDernierMessage As Double

Private Sub Class_Initialize()
    Dim i As Long
    Dim j As Long
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    
    FichierConseil = FicIni.Chemin & Langues.Dossier & FicIni.Parametre("FichierConseil")
    PrefixeConseil = FicIni.Parametre("PrefixeConseil")
    
    Afficher = CBool(FicIni.Parametre("Afficher"))
    NombreMaxMessages = Val(FicIni.Parametre("NombreMaxMessages"))
    
    AfficherConseil = CBool(FicIni.Parametre("AfficherConseil"))
    AfficherVie = CBool(FicIni.Parametre("AfficherVie"))
    AfficherBestiaire = CBool(FicIni.Parametre("AfficherBestiaire"))
    AfficherArgent = CBool(FicIni.Parametre("AfficherArgent"))
    AfficherRessources = CBool(FicIni.Parametre("AfficherRessources"))
    AfficherObjets = CBool(FicIni.Parametre("AfficherObjets"))
    AfficherSoldats = CBool(FicIni.Parametre("AfficherSoldats"))
    AfficherExperience = CBool(FicIni.Parametre("AfficherExperience"))
    AfficherEmpereur = CBool(FicIni.Parametre("AfficherEmpereur"))
    
    Duree = Val(FicIni.Parametre("Duree"))
    Couleur = Val(FicIni.Parametre("Couleur"))
    PositionX = Val(FicIni.Parametre("PositionX"))
    PositionY = Val(FicIni.Parametre("PositionY"))
    DecalageX = Val(FicIni.Parametre("DecalageX"))
    DecalageY = Val(FicIni.Parametre("DecalageY"))
    
    PrefixeMessage = FicIni.Parametre("PrefixeMessage") & " "
    PrefixeExperience = FicIni.Parametre("PrefixeExperience") & " "
    
    MessageVivant = FicIni.Parametre("MessageVivant")
    MessageMort = FicIni.Parametre("MessageMort")
    MessageArgent = " " & FicIni.Parametre("MessageArgent")
    MessageGain = FicIni.Parametre("MessageGain") & " "
    MessagePerte = FicIni.Parametre("MessagePerte") & " "
    MessageBestiaire = FicIni.Parametre("MessageBestiaire") & " "
    MessageVie = " " & FicIni.Parametre("MessageVie")
    MessageEnergie = " " & FicIni.Parametre("MessageEnergie")
    MessageMagie = " " & FicIni.Parametre("MessageMagie")
    MessageMoral = " " & FicIni.Parametre("MessageMoral")
    MessageAttaque = " " & FicIni.Parametre("MessageAttaque")
    MessageDefense = " " & FicIni.Parametre("MessageDefense")
    MessageSoldat = " " & FicIni.Parametre("MessageSoldat")
    MessageSoldats = " " & FicIni.Parametre("MessageSoldats")
    MessageCompetence = " " & FicIni.Parametre("MessageCompetence") & " "
    MessageEmpereurVieM = FicIni.Parametre("MessageEmpereurVieM")
    MessageEmpereurVieF = FicIni.Parametre("MessageEmpereurVieF")
    MessageEmpereurMortM = FicIni.Parametre("MessageEmpereurMortM")
    MessageEmpereurMortF = FicIni.Parametre("MessageEmpereurMortF")
    
    MessageSauvegardeMonde = FicIni.Parametre("MessageSauvegardeMonde")
    
    Set AffMessage = New ClsIntMessage
    Set Messages = New Collection
End Sub

Public Function Init(ByVal Parametres As ClsJeuParam, _
                     Optional ByVal DeltaX As Integer, _
                     Optional ByVal DeltaY As Integer)
    Numero = -1
    Argent = 0
    ReDim Ressources(Parametres.NombreRessources - 1)
    ReDim Objets(Parametres.NombreObjets - 1)
    ReDim Bestiaire(Parametres.NombreRaces - 1)
    ReDim NivSpeciales(Parametres.NombreCompetencesSpeciales - 1)
    ReDim NivRessources(Parametres.NombreRessources - 1)
    ReDim NivServices(Parametres.NombreServices - 2)
    ReDim NivObjets(Parametres.NombreCompetencesObjets - 1)
    Empereur = -1
    
    PositionX = PositionX + DeltaX
    PositionY = PositionY + DeltaY
End Function

Public Sub Actualiser_Messages(ByVal Perso As ClsJeuPerso, _
                               ByVal Fief As ClsJeuFief, _
                               ByVal Parametres As ClsJeuParam)
    Dim i As Integer
    Dim Message As String
    With Perso
    If Afficher Then
        'On ajoute les nouveaux messages.
        'Réinitialise si changement de personnage.
        If Numero <> .Numero Then
            Numero = .Numero
            Vivant = .Vivant
            Argent = .Definir_Argent
            For i = 0 To Parametres.NombreRessources - 1
                Ressources(i) = .Definir_ressources(i)
            Next i
            For i = 0 To Parametres.NombreRaces - 1
                Bestiaire(i) = .Definir_Bestiaire(i)
            Next i
            Soldats = .NombreSoldats
            MaxVie = .MaxVie
            MaxEnergie = .MaxEnergie
            MaxMagie = .MaxMagie
            MaxMoral = .MaxMoral
            MaxAttaque = .MaxAttaque
            MaxDefense = .MaxDefense
            For i = 0 To Parametres.NombreCompetencesSpeciales - 1
                NivSpeciales(i) = .Niveau_CompetenceSpeciales(i)
            Next i
            For i = 0 To Parametres.NombreRessources - 1
                NivRessources(i) = .Niveau_CompetenceRessources(i)
            Next i
            For i = 0 To Parametres.NombreObjets - 1
                Objets(i) = 0
            Next i
            For i = 0 To Parametres.PersosNombreObjetsEquipes - 1
                If .Objet_Equipes_Actif(i) Then
                    Objets(.Objet_Equipes_Type(i)) = Objets(.Objet_Equipes_Type(i)) + 1
                End If
            Next i
            For i = 0 To Parametres.PersosNombreObjetsInventaire - 1
                If .Objet_Inventaire_Actif(i) Then
                    Objets(.Objet_Inventaire_Type(i)) = Objets(.Objet_Inventaire_Type(i)) + 1
                End If
            Next i
            For i = 0 To Parametres.NombreServices - 2
                NivServices(i) = .Niveau_CompetenceServices(i)
            Next i
            For i = 0 To Parametres.NombreCompetencesObjets - 1
                NivObjets(i) = .Niveau_CompetenceObjets(i)
            Next i
            Empereur = Fief.Empereur
        Else
        
            'Inforamtions sur la vie du personnage.
            If AfficherVie Then
                If .Vivant = Not Vivant Then
                    If .Vivant Then
                        Ajouter_Message PrefixeMessage & MessageVivant
                    Else
                        Ajouter_Message PrefixeMessage & MessageMort
                    End If
                    Vivant = .Vivant
                End If
            End If
            
            'Messages à propos du bestiaire.
            If AfficherBestiaire Then
                For i = 0 To Parametres.NombreRaces - 1
                    If .Definir_Bestiaire(i) - Bestiaire(i) >= 1 Then
                        Ajouter_Message PrefixeMessage & MessageBestiaire & Int(.Definir_Bestiaire(i) - Bestiaire(i)) & " " & LCase(IIf(.Definir_Bestiaire(i) - Bestiaire(i) > 1, Parametres.Race_Nom(i, False, True), Parametres.Race_Nom(i, False, False))) & "."
                    End If
                    Bestiaire(i) = .Definir_Bestiaire(i)
                Next i
            End If
            
            'Messages à propos des soldats.
            If AfficherSoldats Then
                If .NombreSoldats - Soldats >= 1 Then
                    Ajouter_Message PrefixeMessage & MessageGain & .NombreSoldats - Soldats & IIf(.NombreSoldats - Soldats = 1, MessageSoldat, MessageSoldats)
                ElseIf Soldats - .NombreSoldats >= 1 Then
                    Ajouter_Message PrefixeMessage & MessagePerte & Soldats - .NombreSoldats & IIf(Soldats - .NombreSoldats = 1, MessageSoldat, MessageSoldats)
                End If
                Soldats = .NombreSoldats
            End If
            
            'Message à propos de l'argent.
            If AfficherArgent Then
                If .Definir_Argent - Argent - IIf(.Vivant, .Definir_GenerationArgent, 0) >= 1 Then
                    Ajouter_Message PrefixeMessage & MessageGain & Int(.Definir_Argent - Argent - IIf(.Vivant, .Definir_GenerationArgent, 0)) & MessageArgent
                ElseIf Argent + IIf(.Vivant, .Definir_GenerationArgent, 0) - .Definir_Argent >= 1 Then
                    Ajouter_Message PrefixeMessage & MessagePerte & Int(Argent - .Definir_Argent + IIf(.Vivant, .Definir_GenerationArgent, 0)) & MessageArgent
                End If
                Argent = .Definir_Argent
            End If
            
            'Messages à propos des ressources.
            If AfficherRessources Then
                For i = 0 To Parametres.NombreRessources - 1
                    If .Definir_ressources(i) - Ressources(i) - IIf(.Vivant, .Definir_GenerationRessource(i), 0) >= 1 And i <> .RessourceCourante Then
                        If Message = "" Then
                            Message = MessageGain
                        Else
                            Message = Message & ", "
                        End If
                        Message = Message & Int(.Definir_ressources(i) - Ressources(i) - IIf(.Vivant, .Definir_GenerationRessource(i), 0)) & " " & LCase(Parametres.Ressources_Nom(i))
                    End If
                Next i
                If Message <> "" Then
                    Ajouter_Message PrefixeMessage & Message & "."
                    Message = ""
                End If
                For i = 0 To Parametres.NombreRessources - 1
                    If Ressources(i) - .Definir_ressources(i) - IIf(.Vivant, .Definir_GenerationRessource(i), 0) >= 1 Then
                        If .RessourceCourante >= 0 Then
                            If i <> Parametres.Ressources_TypeMatierePremiere(.RessourceCourante) Then
                            If Message = "" Then
                                Message = MessagePerte
                            Else
                                Message = Message & ", "
                            End If
                            Message = Message & Int(Ressources(i) - .Definir_ressources(i) + IIf(.Vivant, .Definir_GenerationRessource(i), 0)) & " " & LCase(Parametres.Ressources_Nom(i))
                            End If
                        Else
                            If Message = "" Then
                                Message = MessagePerte
                            Else
                                Message = Message & ", "
                            End If
                            Message = Message & Int(Ressources(i) - .Definir_ressources(i) + IIf(.Vivant, .Definir_GenerationRessource(i), 0)) & " " & LCase(Parametres.Ressources_Nom(i))
                        End If
                    End If
                    Ressources(i) = .Definir_ressources(i)
                Next i
                If Message <> "" Then
                    Ajouter_Message PrefixeMessage & Message & "."
                    Message = ""
                End If
            End If
            
            'Messages à propos des objets.
            If AfficherObjets Then
                For i = 0 To Parametres.NombreObjets - 1
                    'On a gagné un objet.
                    If .Objet_NombreType(i) - Objets(i) + IIf(.ObjetSelectionne = i, 1, 0) >= 1 Then
                        If Message = "" Then
                            Message = MessageGain
                        Else
                            Message = Message & ", "
                        End If
                        Message = Message & Int(.Objet_NombreType(i) - Objets(i) + IIf(.ObjetSelectionne = i, 1, 0)) & " " & IIf(Int(.Objet_NombreType(i) - Objets(i) + IIf(.ObjetSelectionne = i, 1, 0)) >= 2, Parametres.Objet_NomPluriel(i), Parametres.Objet_Nom(i))
                    End If
                Next i
                If Message <> "" Then
                    Ajouter_Message PrefixeMessage & Message & "."
                    Message = ""
                End If
                For i = 0 To Parametres.NombreObjets - 1
                    'On a perdu un objet.
                    If Objets(i) - .Objet_NombreType(i) - IIf(.ObjetSelectionne = i, 1, 0) >= 1 Then
                        If Message = "" Then
                            Message = MessagePerte
                        Else
                            Message = Message & ", "
                        End If
                        Message = Message & Int(Objets(i) - .Objet_NombreType(i)) & " " & IIf(Objets(i) - .Objet_NombreType(i) >= 2, Parametres.Objet_NomPluriel(i), Parametres.Objet_Nom(i))
                    End If
                    Objets(i) = .Objet_NombreType(i) + IIf(.ObjetSelectionne = i, 1, 0)
                Next i
                If Message <> "" Then
                    Ajouter_Message PrefixeMessage & Message & "."
                    Message = ""
                End If
            End If
            
            'Messages à propos de l'expérience.
            If AfficherExperience Then
                If .MaxVie - MaxVie >= 1 Then
                    Ajouter_Message PrefixeExperience & MessageGain & Int(.MaxVie - MaxVie) & MessageVie
                End If
                MaxVie = .MaxVie
                If .MaxEnergie - MaxEnergie >= 1 Then
                    Ajouter_Message PrefixeExperience & MessageGain & Int(.MaxEnergie - MaxEnergie) & MessageEnergie
                End If
                MaxEnergie = .MaxEnergie
                If .MaxMagie - MaxMagie >= 1 Then
                    Ajouter_Message PrefixeExperience & MessageGain & Int(.MaxMagie - MaxMagie) & MessageMagie
                End If
                MaxMagie = .MaxMagie
                If .MaxMoral - MaxMoral >= 1 Then
                    Ajouter_Message PrefixeExperience & MessageGain & Int(.MaxMoral - MaxMoral) & MessageMoral
                End If
                MaxMoral = .MaxMoral
                If .MaxAttaque - MaxAttaque >= 1 Then
                    Ajouter_Message PrefixeExperience & MessageGain & Int(.MaxAttaque - MaxAttaque) & MessageAttaque
                End If
                MaxAttaque = .MaxAttaque
                If .MaxDefense - MaxDefense >= 1 Then
                    Ajouter_Message PrefixeExperience & MessageGain & Int(.MaxDefense - MaxDefense) & MessageDefense
                End If
                MaxDefense = .MaxDefense
                For i = 0 To Parametres.NombreCompetencesSpeciales - 1
                    If .Niveau_CompetenceSpeciales(i) - NivSpeciales(i) >= 1 Then
                        Ajouter_Message PrefixeExperience & MessageGain & Int(.Niveau_CompetenceSpeciales(i) - NivSpeciales(i)) & MessageCompetence & LCase(Parametres.Speciales_NomCompetences(i)) & "."
                    End If
                    NivSpeciales(i) = .Niveau_CompetenceSpeciales(i)
                Next i
                For i = 0 To Parametres.NombreRessources - 1
                    If .Niveau_CompetenceRessources(i) - NivRessources(i) >= 1 Then
                        Ajouter_Message PrefixeExperience & MessageGain & Int(.Niveau_CompetenceRessources(i) - NivRessources(i)) & MessageCompetence & LCase(Parametres.Ressources_NomCompetence(i)) & "."
                    End If
                    NivRessources(i) = .Niveau_CompetenceRessources(i)
                Next i
                For i = 0 To Parametres.NombreServices - 2
                    If .Niveau_CompetenceServices(i) - NivServices(i) >= 1 Then
                        Ajouter_Message PrefixeExperience & MessageGain & Int(.Niveau_CompetenceServices(i) - NivServices(i)) & MessageCompetence & LCase(Parametres.Service_NomCompetence(i + 1)) & "."
                    End If
                    NivServices(i) = .Niveau_CompetenceServices(i)
                Next i
                For i = 0 To Parametres.NombreCompetencesObjets - 1
                    If .Niveau_CompetenceObjets(i) - NivObjets(i) >= 1 Then
                        Ajouter_Message PrefixeExperience & MessageGain & Int(.Niveau_CompetenceObjets(i) - NivObjets(i)) & MessageCompetence & LCase(Parametres.CompetenceObjet_Nom(i)) & "."
                    End If
                    NivObjets(i) = .Niveau_CompetenceObjets(i)
                Next i
            End If
            'Messages à propos de l'expérience.
            If AfficherEmpereur Then
                If Empereur <> Fief.Empereur Then
                    'Le statut de l'empereur change.
                    If Fief.Empereur = -1 Then
                        'L'empereur est mort.
                        If Persos(Empereur).Feminin Then
                            Ajouter_Message PrefixeMessage & UCase(Parametres.Epoque_EmpereurNom(Fief.Epoque, Persos(Empereur).Feminin) & " " & Persos(Empereur).Nom & MessageEmpereurMortF)
                        Else
                            Ajouter_Message PrefixeMessage & UCase(Parametres.Epoque_EmpereurNom(Fief.Epoque, Persos(Empereur).Feminin) & " " & Persos(Empereur).Nom & MessageEmpereurMortM)
                        End If
                    Else
                        'Nouvel empereur.
                        If Persos(Fief.Empereur).Feminin Then
                            Ajouter_Message PrefixeMessage & UCase(Persos(Fief.Empereur).Nom & MessageEmpereurVieF & Parametres.Epoque_EmpereurNom(Fief.Epoque, Persos(Fief.Empereur).Feminin))
                        Else
                            Ajouter_Message PrefixeMessage & UCase(Persos(Fief.Empereur).Nom & MessageEmpereurVieM & Parametres.Epoque_EmpereurNom(Fief.Epoque, Persos(Fief.Empereur).Feminin))
                        End If
                    End If
                    Empereur = Fief.Empereur
                End If
            End If
        End If
    End If
    End With
End Sub

Public Sub Afficher_Messages(ByRef SurfBack As IDirectDrawSurface7)
    Dim i As Integer
    Dim Temps As Single 'Variable temporaire contenant le timer.
    If Afficher Then
        Temps = Timer
        'On efface les messages en trop.
        While Messages.Count > NombreMaxMessages
            Messages.Remove 1
            TempsDernierMessage = Temps
        Wend
        
        'On efface le chronomètre.
        If Jeu.IntituleChrono <> "" And Jeu.Chrono > 0 Then
            If Messages.Count > 0 Then
                Messages.Remove 1
            End If
        End If
        
        'On efface les anciens messages.
        If Messages.Count > 0 Then
            If Abs(Temps - TempsDernierMessage) > Len(Messages.Item(1)) * Duree Then
                Messages.Remove 1
                TempsDernierMessage = Temps
            End If
        End If
    
        'Chronomètre.
        If Jeu.IntituleChrono <> "" And Jeu.Chrono > 0 Then
            If Messages.Count = 0 Then
                Messages.Add Jeu.IntituleChrono & Jeu.Definir_Chrono
            Else
                Messages.Add Jeu.IntituleChrono & Jeu.Definir_Chrono, , 1
            End If
        End If

        'On affiche tous les messages.
        If Messages.Count > 0 Then
            SurfBack.SetForeColor Couleur
            For i = 0 To Messages.Count - 1
                AffMessage.Afficher PositionX + DecalageX * i, _
                                    PositionY + DecalageY * i, _
                                    Messages.Item(i + 1), _
                                    SurfBack
            Next i
            SurfBack.SetForeColor 0
        End If
    End If
End Sub

Public Sub Afficher_Conseil()
    Dim i As Long
    Dim Prefixe As String
    Dim Temp As String
    Dim Conseils() As String
    
    If AfficherConseil Then
        Open App.Path & FichierConseil For Input As #1
        ReDim Conseils(0)
        Line Input #1, Prefixe
        While Not EOF(1)
            Line Input #1, Temp
            If Left(Temp, 1) <> "#" And Len(Temp) > 0 Then
                If Left(Temp, 1) = PrefixeConseil Then
                    ReDim Preserve Conseils(i)
                    Conseils(i) = Temp
                    i = i + 1
                Else
                    Conseils(i - 1) = Conseils(i - 1) & AffMessage.Caractere_RetourLigne & Temp
                End If
            End If
        Wend
        Close #1
        Temp = Rnd * (i - 1)
        'Debug.Print Temp & " / " & UBound(Conseils())
        Ajouter_Message Prefixe & " " & Conseils(Val(Temp))
        'Ajouter_Message Prefixe & " " & Conseils(0)
    End If
End Sub

Public Sub Ajouter_Message(ByVal Message As String)
    Dim i As Integer

    If Messages.Count = 0 Or (Messages.Count = 1 And Jeu.IntituleChrono <> "" And Jeu.Chrono > 0) Then
        TempsDernierMessage = Timer
    End If
    Messages.Add Message
    If AffMessage.Nombre_Lignes(Message) > 1 Then
        For i = 1 To AffMessage.Nombre_Lignes(Message)
            Messages.Add ""
        Next i
    End If
End Sub

Public Sub Ajouter_Message_SauvegardeMonde()
    Ajouter_Message MessageSauvegardeMonde
End Sub

Public Sub Effacer()
    While Messages.Count > 0
        Messages.Remove 1
    Wend
End Sub

'Propriétés
Public Property Get Fichier_Conseil() As String
    Fichier_Conseil = FichierConseil
End Property

Public Property Get Definir_Afficher() As Boolean
    Definir_Afficher = Afficher
End Property
Public Property Let Definir_Afficher(ByVal Valeur As Boolean)
    Afficher = Valeur
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("Afficher") = -CDbl(Afficher)
End Property

Public Property Get Definir_AfficherConseil() As Boolean
    Definir_AfficherConseil = AfficherConseil
End Property
Public Property Let Definir_AfficherConseil(ByVal Valeur As Boolean)
    AfficherConseil = Valeur
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherConseil") = -CDbl(AfficherConseil)
End Property

Public Property Get Definir_AfficherVie() As Boolean
    Definir_AfficherVie = AfficherVie
End Property
Public Property Let Definir_AfficherVie(ByVal Valeur As Boolean)
    AfficherVie = Valeur
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherVie") = -CDbl(AfficherVie)
End Property

Public Property Get Definir_AfficherBestiaire() As Boolean
    Definir_AfficherBestiaire = AfficherBestiaire
End Property
Public Property Let Definir_AfficherBestiaire(ByVal Valeur As Boolean)
    AfficherBestiaire = Valeur
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherBestiaire") = -CDbl(AfficherBestiaire)
End Property

Public Property Get Definir_AfficherArgent() As Boolean
    Definir_AfficherArgent = AfficherArgent
End Property
Public Property Let Definir_AfficherArgent(ByVal Valeur As Boolean)
    AfficherArgent = Valeur
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherArgent") = -CDbl(AfficherArgent)
End Property

Public Property Get Definir_AfficherRessources() As Boolean
    Definir_AfficherRessources = AfficherRessources
End Property
Public Property Let Definir_AfficherRessources(ByVal Valeur As Boolean)
    AfficherRessources = Valeur
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherRessources") = -CDbl(AfficherRessources)
End Property

Public Property Get Definir_AfficherObjets() As Boolean
    Definir_AfficherObjets = AfficherObjets
End Property
Public Property Let Definir_AfficherObjets(ByVal Valeur As Boolean)
    AfficherObjets = Valeur
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherObjets") = -CDbl(AfficherObjets)
End Property

Public Property Get Definir_AfficherSoldats() As Boolean
    Definir_AfficherSoldats = AfficherSoldats
End Property
Public Property Let Definir_AfficherSoldats(ByVal Valeur As Boolean)
    AfficherSoldats = Valeur
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherSoldats") = -CDbl(AfficherSoldats)
End Property

Public Property Get Definir_AfficherExperience() As Boolean
    Definir_AfficherExperience = AfficherExperience
End Property
Public Property Let Definir_AfficherExperience(ByVal Valeur As Boolean)
    AfficherExperience = Valeur
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherExperience") = -CDbl(AfficherExperience)
End Property

Public Property Get Definir_AfficherEmpereur() As Boolean
    Definir_AfficherEmpereur = AfficherEmpereur
End Property
Public Property Let Definir_AfficherEmpereur(ByVal Valeur As Boolean)
    AfficherEmpereur = Valeur
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre("AfficherEmpereur") = -CDbl(AfficherEmpereur)
End Property
