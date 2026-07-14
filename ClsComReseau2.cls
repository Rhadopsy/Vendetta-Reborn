VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsComReseau2"
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
Const FichierINI = "Reseau"
Const SectionINI = "Reseau"
Const MyGUID = "{7F26304A-5D74-406F-B0A9-0C78FDB6156F}" 'Identifie de façon unique l'application.
Private NumeroPort As Long
Private OuverturePort As Boolean
'Private NumeroIP As String
Private NumeroIPServeur As String

Const DebugReseau = False
Const CheminReseau = "\"
Const FichierReseau = "Reseau.txt"

Private NomReseau As String
Private MessageErreurInitialisation As String
Private MessageErreurAdresses As String
Private MessageErreurRafraichirParties As String
Private MessageErreurSessionCreer As String
Private MessageErreurSessionRejoindre As String
Private MessageJoueurConnecter As String
Private MessageJoueurDeconnecter As String
Private MessageJoueurRejoindre As String
Private MessageJoueurQuitter As String
Private MessageServeurDebutPartie As String
Private MessageServeurFinPartie As String
Private MessageServeurBannir As String
Private MessageServeurDeconnecte As String

Private Actif As Boolean
Public ServiceProvider As Integer
Public NomJoueur As String
Public NomSession As String
Public NombreJoueurs As Long
Public NombreJoueursMin As Long
Public NombreJoueursMax As Long
Public MotDePasse As String

Public Connecte As Boolean
Public Serveur As Boolean
Public Client As Boolean

Public Adresse As Long

Public PlayerID As Long
Public ServeurID As Long

Dim Sock As Winsock

'Propriété utilisée dans la partie.
Private Type Joueur
    ID As Long
    Nom As String
    'Pret As Boolean 'Passe à vrai à chaque fin de tour.
    Serveur As Boolean 'Si vrai, il s'agit du serveur.
    InfosAJour As Boolean  'Vrai, le personnage a déjà toutes les informations.
End Type

Private Joueurs() As Joueur
Public NumeroJoueur As Integer

Private Message As ClsComMessage
Private NombreMessagesEmis As Long
Private NombreMessagesRecus As Long

Private Sub Class_Initialize()
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    
    NomReseau = FicIni.Parametre("NomReseau")
    MessageErreurInitialisation = FicIni.Parametre("MessageErreurInitialisation")
    MessageErreurAdresses = FicIni.Parametre("MessageErreurAdresses")
    MessageErreurRafraichirParties = FicIni.Parametre("MessageErreurRafraichirParties")
    MessageErreurSessionCreer = FicIni.Parametre("MessageErreurSessionCreer")
    MessageErreurSessionRejoindre = FicIni.Parametre("MessageErreurSessionRejoindre")
    MessageJoueurConnecter = FicIni.Parametre("MessageJoueurConnecter")
    MessageJoueurDeconnecter = FicIni.Parametre("MessageJoueurDeconnecter")
    MessageJoueurRejoindre = FicIni.Parametre("MessageJoueurRejoindre")
    MessageJoueurQuitter = FicIni.Parametre("MessageJoueurQuitter")
    MessageServeurDebutPartie = FicIni.Parametre("MessageServeurDebutPartie")
    MessageServeurFinPartie = FicIni.Parametre("MessageServeurFinPartie")
    MessageServeurBannir = FicIni.Parametre("MessageServeurBannir")
    MessageServeurDeconnecte = FicIni.Parametre("MessageServeurDeconnecte")
    
    FicIni.Fichier = FicIni.Chemin & FichierINI
    
    Actif = CBool(FicIni.Parametre("Actif"))
    ServiceProvider = Val(FicIni.Parametre("ServiceProvider"))
    NomJoueur = FicIni.Parametre("NomJoueur")
    NomSession = FicIni.Parametre("NomSession")
    NombreJoueurs = Val(FicIni.Parametre("NombreJoueurs")) + 1
    
    NombreJoueursMin = Val(FicIni.Parametre("NombreJoueursMin")) + 1
    NombreJoueursMax = Val(FicIni.Parametre("NombreJoueursMax")) + 1
    
    NumeroPort = Val(FicIni.Parametre("NumeroPort"))
    OuverturePort = Val(FicIni.Parametre("OuverturePort"))
    
    NumeroIPServeur = FicIni.Parametre("NumeroIPServeur")

    Initialiser_Connexion
    
    PlayerID = 1
End Sub

Private Sub InitAddress()
    
End Sub

Private Sub Initialiser_Connexion()
    
End Sub

Public Sub Rafraichir_Parties()
    Sock.Connect NumeroIPServeur, IIf(OuverturePort, NumeroPort, 0)
End Sub

Public Sub Rafraichir_Joueurs()
    
End Sub

Private Sub Afficher_Information_Reseau(ByVal Message As String)
    
End Sub

Public Sub Afficher_Information_Reseau_PartieEnCours(ByVal EnCours As Integer)
    
End Sub

Public Sub Session_Creer()
    
End Sub

Public Sub Session_Rejoindre(ByVal Index As Integer)
    
End Sub

Public Sub Session_Fermer()
    Sock.Close
End Sub

Public Sub Joueur_Creer(Optional ByVal EstServeur As Boolean)
    
End Sub

Public Sub Message_Envoyer(ByVal Categorie1 As Integer, _
                           ByVal Categorie2 As Integer, _
                           Optional Indice1 As Long, _
                           Optional Indice2 As Long, _
                           Optional ByVal Var1 As String, _
                           Optional ByVal Var2 As String, _
                           Optional ByVal IDJoueur As Long = -1, _
                           Optional ByVal ToutJoueurs As Boolean)
    Message.Temp1 = Categorie1
    Message.Temp2 = Categorie2
    Message.Temp3 = Indice1
    Message.Temp4 = Indice2
    Message.Temp5 = Var1
    Message.Temp6 = Var2
    Sock.SendData Message
    NombreMessagesEmis = NombreMessagesEmis + 1
End Sub

Public Sub Message_Lire(Optional ByVal ChatSeulement As Boolean)
    Sock.GetData Message
    ComMessages.Analyser Message.Temp1, Message.Temp2, Message.Temp3, Message.Temp4, Message.Temp5, Message.Temp6, ChatSeulement
End Sub

' ************************************************************************************
' *                                        Propriétés                                *
' ************************************************************************************
Public Property Get Definir_NombreMessagesRecus() As Long
    Definir_NombreMessagesRecus = NombreMessagesRecus
    NombreMessagesRecus = 0
End Property
Public Property Get Definir_NombreMessagesEmis() As Long
    Definir_NombreMessagesEmis = NombreMessagesEmis
    NombreMessagesEmis = 0
End Property

Public Property Get ServiceProvider_Nom(ByVal Index As Integer) As String
    
End Property

Public Property Get Partie_Nom(ByVal Index As Integer) As String
    'Renvoie le nom de la partie entrée en indice.
    
End Property
Public Property Get Partie_NombreJoueurs(ByVal Index As Integer) As Long
    'Renvoie le nombre de joueurs actuels de la partie entrée en indice.
    
End Property
Public Property Get Partie_NombreJoueursMax(ByVal Index As Integer) As Long
    'Renvoie le nombre de joueurs actuels de la partie entrée en indice.
    
End Property

Public Property Get Joueur_Nom(ByVal Index As Integer) As String
    'Renvoie le nom du joueur entré en indice.
    
End Property
Public Property Get Joueur_ID(ByVal Index As Integer) As Long
    'Renvoie l'indentifiant du joueur entré en indice.
    
End Property
Public Property Get Joueur_Serveur(ByVal Index As Integer) As Boolean
    'Renvoie si le joueur est un serveur.
    
End Property

Public Property Get Joueur_InfosAJour(ByVal ID As Long) As Boolean
    Dim i As Long
    'Renvoie si le joueur dont l'ID est entré, s'il est à jour.
    For i = 0 To UBound(Joueurs())
        If Joueurs(i).ID = ID Then
            Joueur_InfosAJour = Joueurs(i).InfosAJour
            Exit For
        End If
    Next i
End Property
Public Property Let Joueur_InfosAJour(ByVal ID As Long, ByVal Valeur As Boolean)
    Dim i As Long
    For i = 0 To UBound(Joueurs())
        If Joueurs(i).ID = ID Then
            Joueurs(i).InfosAJour = Valeur
            Exit For
        End If
    Next i
End Property
Public Property Let Joueur_Tous_InfosAJour(ByVal Valeur As Boolean)
    Dim i As Long
    For i = 0 To UBound(Joueurs())
        Joueurs(i).InfosAJour = Valeur
    Next i
End Property

Public Property Get Definir_Actif() As Boolean
    Definir_Actif = Actif
End Property
Public Property Let Definir_Actif(ByVal Valeur As Boolean)
    Actif = Valeur
    Changer_Parametre_Ini "Actif", -CDbl(Valeur)
End Property
Public Property Let Definir_NomJoueur(ByVal Valeur As String)
    NomJoueur = Valeur
    Changer_Parametre_Ini "NomJoueur", Valeur
End Property
Public Property Let Definir_NomSession(ByVal Valeur As String)
    NomSession = Valeur
    Changer_Parametre_Ini "NomSession", Valeur
    'AppDesc.SessionName = Valeur
End Property
Public Property Let Definir_NombreJoueurs(ByVal Valeur As Long)
    NombreJoueurs = Valeur
    Changer_Parametre_Ini "NombreJoueurs", Valeur - 1
    'AppDesc.lMaxPlayers = Valeur
End Property
Public Property Let Definir_ServiceProvider(ByVal Valeur As Integer)
    ServiceProvider = Valeur
    Changer_Parametre_Ini "ServiceProvider", Valeur
End Property
Public Property Get Definir_NumeroPort() As Long
    Definir_NumeroPort = NumeroPort
End Property
Public Property Let Definir_NumeroPort(ByVal Valeur As Long)
    NumeroPort = Valeur
    Changer_Parametre_Ini "NumeroPort", Valeur
End Property

Public Property Get Definir_OuverturePort() As Boolean
    Definir_OuverturePort = OuverturePort
End Property
Public Property Let Definir_OuverturePort(ByVal Valeur As Boolean)
    If Valeur <> OuverturePort Then
        OuverturePort = Valeur
        Changer_Parametre_Ini "OuverturePort", -CDbl(Valeur)
    End If
End Property

Public Property Get Definir_NumeroIPServeur() As String
    Definir_NumeroIPServeur = NumeroIPServeur
End Property
Public Property Let Definir_NumeroIPServeur(ByVal Valeur As String)
    NumeroIPServeur = Valeur
    Changer_Parametre_Ini "NumeroIPServeur", Valeur
End Property

Private Sub Changer_Parametre_Ini(ByVal Parametre As String, ByVal Valeur As String)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre(Parametre) = Valeur
End Sub


