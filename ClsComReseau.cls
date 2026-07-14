VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsComReseau"
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

Private URL_Serveur As String
Private URL_AjouterPartie As String
Private URL_SupprimerPartie As String
Private URL_MiseAJourPartie As String
Private URL_ListeParties As String
Private URL_NomPartie As String 'Nom de la partie
Private URL_NomJoueur As String 'Nom du créateur de la partie
Private URL_NoPort As String

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
Private JouerEnLigne As Boolean
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

Private Dplay As DirectPlay4

'DirectPlay.
'Private DPPeer As DirectPlay8Peer
'Private DPClient As DirectPlay8Client     '//This is the most important part
'Private DPServer As DirectPlay8Server     '//This is the most important part
'Private DPClientAddress As DirectPlay8Address '//Client IP Address, Port
'Private DPServerAddress As DirectPlay8Address '//Server IP Address, Port
'Private DPEvent As DirectPlay8Event
'Private DPLobbiedApplication As DirectPlay8LobbiedApplication

'Paramètres de communication.
Private dPlayEnumConnexions As DirectPlayEnumConnections
Public lngNumConnexions As Long
'Private DPlayEnumConnexions()  As DPN_SERVICE_PROVIDER_INFO
Private dPlayAdress As DirectPlayAddress

'Paramètres des sessions.
Private SessionData As DirectPlaySessionData
'Dim AppDesc As DPN_APPLICATION_DESC     '//Application description
Private dPlayEnumSession As DirectPlayEnumSessions
Public lngNumSession As Integer
Public SessionsEnLigne As XMLClass 'Charges toutes les parties en lignes à partir d'un serveur PHP dans un fichier XML.
Public ServeurEnLigneID As String 'Enregistre le nom du fichier d'information XML du serveur PHP.

'Paramètres des joueurs.
Private dPlayEnumJoueurs As DirectPlayEnumPlayers
Public lngNumJoueurs As Integer
Public PlayerID As Long
Public ServeurID As Long

'Paramètres des messages.
Private dpMessage As DirectPlayMessage

'Paramètres Internet.
Private dPlayLobby3 As DirectPlayLobby3
Private dPlayLobbyConnection As DirectPlayLobbyConnection
'Private dPlayAdressNet As DirectPlayAddress

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

Private NombreMessagesEmis As Long
Private NombreMessagesRecus As Long

Private Sub Class_Initialize()
    Set SessionsEnLigne = New XMLClass

    FicIni.fichier = FicIni.chemin & Langues.Dossier & FichierINI
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
    
    FicIni.fichier = FicIni.chemin & FichierINI
    
    Actif = CBool(FicIni.Parametre("Actif"))
    JouerEnLigne = CBool(FicIni.Parametre("JouerEnLigne"))
    ServiceProvider = Val(FicIni.Parametre("ServiceProvider"))
    NomJoueur = FicIni.Parametre("NomJoueur")
    NomSession = FicIni.Parametre("NomSession")
    NombreJoueurs = Val(FicIni.Parametre("NombreJoueurs")) + 1
    
    NombreJoueursMin = Val(FicIni.Parametre("NombreJoueursMin")) + 1
    NombreJoueursMax = Val(FicIni.Parametre("NombreJoueursMax")) + 1
    
    NumeroPort = Val(FicIni.Parametre("NumeroPort"))
    OuverturePort = Val(FicIni.Parametre("OuverturePort"))
    
    NumeroIPServeur = FicIni.Parametre("NumeroIPServeur")
    
    URL_Serveur = FicIni.Parametre("URL_Serveur")
    URL_AjouterPartie = FicIni.Parametre("URL_AjouterPartie")
    URL_SupprimerPartie = FicIni.Parametre("URL_SupprimerPartie")
    URL_MiseAJourPartie = FicIni.Parametre("URL_MiseAJourPartie")
    URL_ListeParties = FicIni.Parametre("URL_ListeParties")
    URL_NomPartie = FicIni.Parametre("URL_NomPartie")
    URL_NomJoueur = FicIni.Parametre("URL_NomJoueur")
    URL_NoPort = FicIni.Parametre("URL_NoPort")

    Initialiser_Connexion
    
    PlayerID = 1
End Sub

Private Sub InitDplay()
    'Set Dplay = dx.DirectPlayCreate("")
    'Set dPlayLobby3 = dx.DirectPlayLobbyCreate()
End Sub

Private Sub InitAddress()
    On Error GoTo Erreur
    'Set dPlayAdress = dPlayEnumConnexions.GetAddress(ServiceProvider)
    If OuverturePort Then
        If NumeroIPServeur <> "" Then
            Set dPlayAdress = dPlayLobby3.CreateINetAddress(NumeroIPServeur, NumeroPort)
        Else
            Set dPlayAdress = dPlayLobby3.CreateINetAddress("", NumeroPort)
        End If
    Else
        If NumeroIPServeur <> "" Then
            Set dPlayAdress = dPlayLobby3.CreateINetAddress(NumeroIPServeur, 0)
        Else
            Set dPlayAdress = dPlayLobby3.CreateINetAddress("", 0)
        End If
        'Set dPlayAdress = dPlayEnumConnexions.GetAddress(ServiceProvider)
    End If
    
    Dplay.InitializeConnection dPlayAdress
    Exit Sub
Erreur:
    MsgBox MessageErreurAdresses, vbCritical, NomReseau
End Sub

Private Sub Initialiser_Connexion()
    On Error GoTo Erreur
    Set Dplay = dx.DirectPlayCreate("")
    Set dPlayLobby3 = dx.DirectPlayLobbyCreate()
    
    Set dPlayEnumConnexions = Dplay.GetDPEnumConnections("", DPCONNECTION_DIRECTPLAY)
    lngNumConnexions = dPlayEnumConnexions.GetCount
    Exit Sub
Erreur:
    MsgBox MessageErreurInitialisation & vbCrLf & Err.Number & " : " & Err.Description, vbCritical, NomReseau
End Sub

Public Sub Rafraichir_Parties(Optional ByVal Locales As Boolean)
    On Error GoTo Erreur
    'Initialiser_Connexion
    If JouerEnLigne And Not Locales Then
        SessionsEnLigne.ouvrirXML2 (FrmParam.Inet1.OpenURL(URL_Serveur & URL_ListeParties))
    Else
        Dim TempSession As DirectPlaySessionData
        
        If Dplay Is Nothing Then InitDplay
        If dPlayAdress Is Nothing Or Locales Then InitAddress
        Set TempSession = Dplay.CreateSessionData()
        TempSession.SetGuidApplication MyGUID
        TempSession.SetSessionPassword MotDePasse
        'Set dPlayEnumSession = Dplay.GetDPEnumSessions(TempSession, 0, DPENUMSESSIONS_ASYNC Or DPENUMSESSIONS_ALL Or DPENUMSESSIONS_PASSWORDREQUIRED)
        Set dPlayEnumSession = Dplay.GetDPEnumSessions(TempSession, 0, DPENUMSESSIONS_AVAILABLE)
        DoEvents
        Sleep 1000
        Set dPlayEnumSession = Dplay.GetDPEnumSessions(TempSession, 0, DPENUMSESSIONS_STOPASYNC Or DPENUMSESSIONS_ALL Or DPENUMSESSIONS_PASSWORDREQUIRED)
        lngNumSession = dPlayEnumSession.GetCount
    End If
    Exit Sub
Erreur:
    MsgBox MessageErreurRafraichirParties & vbCrLf & Err.Number & " : " & Err.Description, vbCritical, NomReseau
    'Set dPlayEnumSession = Dplay.GetDPEnumSessions(TempSession, 0, DPENUMSESSIONS_AVAILABLE)
End Sub

Public Sub Rafraichir_Joueurs()
    On Error GoTo Erreur
    Dim i As Integer: Dim j As Integer
    Dim Temp As Boolean
    Dim TempConnecte As Boolean 'Mémorise la connection pour voir si l'on a été expulsé de la partie.
    Dim TempJoueurs As Collection 'Mémorise les joueurs déjà présents pour détecter les joueurs entrant et sortant.
    Set TempJoueurs = New Collection
    For i = 0 To lngNumJoueurs - 1
        TempJoueurs.Add Joueurs(i).Nom
    Next i
    TempConnecte = False
    Set dPlayEnumJoueurs = Nothing
    Set dPlayEnumJoueurs = Dplay.GetDPEnumPlayers("", DPENUMPLAYERS_ALL)
    lngNumJoueurs = dPlayEnumJoueurs.GetCount
    ReDim Joueurs(lngNumJoueurs - 1)
    For i = 0 To lngNumJoueurs - 1
        Joueurs(i).ID = Joueur_ID(i + 1)
        Joueurs(i).Nom = Joueur_Nom(i + 1)
        Joueurs(i).Serveur = Joueur_Serveur(i + 1)
        If Joueurs(i).Serveur Then
            ServeurID = Joueurs(i).ID
            TempConnecte = True
        End If
        If Joueurs(i).Nom = NomJoueur Then
            NumeroJoueur = i
        End If
    Next i
    If TempConnecte <> Connecte Then GoTo Erreur
    'Regarde quel joueur est sorti de la partie.
    For i = 1 To TempJoueurs.Count
        Temp = True 'Si temp reste vrai, on a pas retrouvé le joueur précédemment présent.
        For j = 0 To lngNumJoueurs - 1
            If Joueurs(j).Nom = TempJoueurs(i) Or Joueurs(j).ID = PlayerID Then
                Temp = False
                Exit For
            End If
        Next j
        If Temp Then
            Afficher_Information_Reseau Chr(34) & TempJoueurs(i) & Chr(34) & MessageJoueurDeconnecter
        End If
    Next i
    'Regarde quel joueur est entré dans la partie.
    For i = 0 To lngNumJoueurs - 1
        If Joueurs(i).ID <> PlayerID Then
            Temp = True 'Si temp reste vrai, on a pas retrouvé le joueur précédemment présent.
            For j = 1 To TempJoueurs.Count
                If Joueurs(i).Nom = TempJoueurs(j) Then
                    Temp = False
                    Exit For
                End If
            Next j
            If Temp Then
                Afficher_Information_Reseau Chr(34) & Joueurs(i).Nom & Chr(34) & MessageJoueurConnecter
            End If
        End If
    Next i
    Exit Sub
Erreur:
    Session_Fermer
End Sub

Private Sub Afficher_Information_Reseau(ByVal Message As String)
    'Affiche un message quand un évènement réseau a lieu.
    Select Case Partie.EnCours
    Case 0: FrmParam.LstMessagesInfos.AddItem Message
    Case 2: Messages(0).Ajouter_Message Message
    End Select
End Sub

Public Sub Afficher_Information_Reseau_PartieEnCours(ByVal EnCours As Integer)
    If EnCours <> Partie.EnCours Then
        Select Case EnCours
        Case 0: 'Afficher_Information_Reseau MessageServeurFinPartie
        Case 2: Afficher_Information_Reseau MessageServeurDebutPartie
        End Select
    End If
End Sub

Public Sub Session_Creer()
    On Error GoTo Erreur
    If Dplay Is Nothing Then Call InitDplay
    If dPlayAdress Is Nothing Then Call InitAddress
    
    Set SessionData = Dplay.CreateSessionData()
    
    With SessionData
        .SetGuidApplication MyGUID
        .SetMaxPlayers NombreJoueurs - 1
        .SetSessionName NomSession
        .SetSessionPassword MotDePasse
        .SetFlags DPSESSION_DIRECTPLAYPROTOCOL Or DPSESSION_MIGRATEHOST
    End With
    
    Dplay.Open SessionData, DPOPEN_CREATE
    
    If JouerEnLigne Then
        ServeurEnLigneID = FrmParam.Inet1.OpenURL(URL_Serveur & URL_AjouterPartie & "?" & _
                                         URLEncode(URL_NomPartie) & "=" & URLEncode(NomSession) & Chr(38) & _
                                         URLEncode(URL_NomJoueur) & "=" & URLEncode(NomJoueur) & Chr(38) & _
                                         URLEncode(URL_NoPort) & "=" & URLEncode(NumeroPort))
    End If
    
    Joueur_Creer True
    Connecte = True
    Serveur = True
    Exit Sub
Erreur:
    MsgBox MessageErreurSessionCreer & vbCrLf & Err.Number & " : " & Err.Description, vbCritical, NomReseau
    'MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "InitServer"
    'End
End Sub

Public Sub Session_Rejoindre(ByVal Index As Integer, Optional ByVal Locale As Boolean)
    On Error GoTo Erreur
    If JouerEnLigne And Not Locale Then
        Dim Temp3() As String
        Dim Temp4() As String
        Temp3 = SessionsEnLigne.lireDataXML2("GAME", "IP")
        Temp4 = SessionsEnLigne.lireDataXML2("GAME", "PORT")
        FrmParam.TxtNumeroIPServeur = Temp3(Index)
        FrmParam.TxtNumeroPort = Temp4(Index)
        Rafraichir_Parties True
        If lngNumSession > 0 Then
            Session_Rejoindre 0, True
        Else
            GoTo Erreur
        End If
    Else
        Set SessionData = dPlayEnumSession.GetItem(Index + 1)
        Dplay.Open SessionData, DPOPEN_JOIN
    End If
    Joueur_Creer
    Connecte = True
    Client = True
    Exit Sub
Erreur:
    MsgBox MessageErreurSessionRejoindre & vbCrLf & Err.Number & " : " & Err.Description, vbCritical, NomReseau
    'MsgBox "Impossible d'établir la connexion", vbCritical, "Erreur"
    Client = False
End Sub

Public Sub Session_Fermer()
    If Not Quitter And Not Serveur Then
        Quitter = True
        MsgBox MessageServeurDeconnecte, vbCritical, NomReseau
    End If
    On Error Resume Next
    If Serveur Then
        If JouerEnLigne Then
            FrmParam.Inet1.OpenURL (URL_Serveur & URL_SupprimerPartie & "?FILENAME=" & URLEncode(ServeurEnLigneID))
        End If
    '    objDPServer.CancelAsyncOperation 0, DPNCANCEL_ALL_OPERATIONS
    '    objDPServer.UnRegisterMessageHandler
    '    objDPServer.Close
    '    Set objDPServer = Nothing
    '    Set ObjDPServerAddress = Nothing
    '    Set objDX = Nothing
    End If
    Dplay.Close
    Connecte = False
    Serveur = False
    Client = False
    lngNumJoueurs = 0
End Sub

Public Sub Joueur_Creer(Optional ByVal EstServeur As Boolean)
    PlayerID = Dplay.CreatePlayer(NomJoueur, "", 0, IIf(EstServeur, DPPLAYER_SERVERPLAYER, DPPLAYER_DEFAULT))
    Set dpMessage = Dplay.CreateMessage
    'Dim PlayerInfo As DPN_PLAYER_INFO
    '//Set up peer info
    'PlayerInfo.Name = NomJoueur
    'PlayerInfo.lInfoFlags = DPNINFO_NAME
    'DPClient.SetClientInfo PlayerInfo, DPNOP_SYNC
End Sub

Public Sub Message_Envoyer(ByVal Categorie1 As Integer, _
                           ByVal Categorie2 As Integer, _
                           Optional Indice1 As Long, _
                           Optional Indice2 As Long, _
                           Optional ByVal Var1 As String, _
                           Optional ByVal Var2 As String, _
                           Optional ByVal IDJoueur As Long = -1, _
                           Optional ByVal ToutJoueurs As Boolean)
    On Error GoTo Erreur
    Set dpMessage = Dplay.CreateMessage
    dpMessage.WriteShort Categorie1
    dpMessage.WriteShort Categorie2
    dpMessage.WriteLong Indice1
    dpMessage.WriteLong Indice2
    If Var1 = "" Then Var1 = " "
    dpMessage.WriteString Var1
    If Var2 = "" Then Var2 = " "
    dpMessage.WriteString Var2
    If ToutJoueurs Then
        Dplay.SendEx PlayerID, DPID_ALLPLAYERS, DPSEND_GUARANTEED, dpMessage, 0, 0, 0
    ElseIf IDJoueur >= 0 Then
        Dplay.SendEx PlayerID, IDJoueur, DPSEND_GUARANTEED, dpMessage, 0, 0, 0
    ElseIf Serveur Then
        Dplay.SendEx PlayerID, DPID_ALLPLAYERS, DPSEND_GUARANTEED, dpMessage, 0, 0, 0
    Else
        Dplay.SendEx PlayerID, ServeurID, DPSEND_GUARANTEED, dpMessage, 0, 0, 0
    End If
    'dPlay.Send PlayerID, PlayerID, DPSEND_GUARANTEED, dpMessage
    'Message_Analyser PlayerID, Noperso, Texte, X, Y, Z 'Analyse également le message pour soi-même.
    'Debug.Print Categorie1 & " : " & Categorie2 & " : " & Indice1 & " : " & Indice2 & " : " & Var1 & " : " & Var2 & " : " & IDJoueur & " : " & ToutJoueurs
    If DebugReseau Then
        If Categorie1 = 2 Then
            Dim Temp As String
            Open App.Path & CheminReseau & FichierReseau For Append As #1
            Temp = "E : (" & Categorie1 & ", " & Categorie2 & ") (" & Indice1 & ", " & Indice2 & ") (" & Var1 & ", " & Var2 & ") - " & IDJoueur & "(" & ToutJoueurs & ")" & Chr(10) & Chr(13)
            Print #1, Temp
            Close #1
        End If
    End If
    NombreMessagesEmis = NombreMessagesEmis + 1
    Exit Sub
Erreur:
    Session_Fermer
End Sub

Public Sub Message_Lire(Optional ByVal ChatSeulement As Boolean)
    Dim lngFromPlayerID As Long     'ID of player that message is from
    Dim lngToPlayerID As Long       'ID of player to whom message was sent
    Dim lngMsgCount As Long 'How many messages are in the queue?
    Dim Ca1 As Long
    Dim Temp1 As Integer, Temp2 As Integer
    Dim Temp3 As Long, Temp4 As Long
    Dim Temp5 As String, Temp6 As String
    
    'Handle errors that may occur if the session or player have not yet been created
    On Error GoTo Erreur
    
    'Find out how many messages are waiting for the player indicated by lngPlayerID
    lngMsgCount = Dplay.GetMessageCount(PlayerID)
    
    NombreMessagesRecus = NombreMessagesRecus + lngMsgCount
    
    Do While lngMsgCount > 0
        Set dpMessage = Dplay.Receive(lngFromPlayerID, _
                                      lngToPlayerID, _
                                       DPRECEIVE_ALL)
        lngMsgCount = lngMsgCount - 1
        'Check if this message was sent by the system
        If lngFromPlayerID = DPID_SYSMSG Then
        '    FrmParam.CmdRafraichirJoueurs_Click
        Else
        'If not sent by system, this must be a application-defined message
            Temp1 = dpMessage.ReadShort
            Temp2 = dpMessage.ReadShort
            Temp3 = dpMessage.ReadLong
            Temp4 = dpMessage.ReadLong
            Temp5 = dpMessage.ReadString
            Temp6 = dpMessage.ReadString
            ComMessages.Analyser Temp1, Temp2, Temp3, Temp4, Temp5, Temp6, ChatSeulement
        End If
    Loop
    Exit Sub
Erreur:
    Exit Sub
End Sub

'Public Sub Tour_Debut()
    'Dim i As Integer
    'For i = 0 To lngNumJoueurs - 1
    '    Joueurs(i).Pret = False
    '    Messages(i).Effacer_Tout
    'Next i
'End Sub
'Public Sub Tour_Fin()
    'Message_Envoyer NumeroJoueur, "Prt"
    'Joueurs(NumeroJoueur).Pret = True
'End Sub
'Public Property Get Tour_Fini() As Boolean
'    'Si un seul joueur n'a pas fini, on l'attend pour finir le tour.
'    Dim i As Integer
'    Tour_Fini = True
'    For i = 0 To lngNumJoueurs - 1
'        If Joueurs(i).Pret = False Then
'            Tour_Fini = False
'            i = lngNumJoueurs
'        End If
'    Next i
'End Property
'Public Sub Tour_Suivant()
'    'Applique toutes les actions qui ont été faites dans le réseau.
'    Dim i As Long, j As Long
'    For i = 0 To lngNumJoueurs - 1
'        With Messages(i)
'        If .NombreMessages > 0 Then
'            For j = 0 To .NombreMessages - 1
'                Select Case .Categorie(j)
'                Case "PMV": 'Personnage Mouvement.
'                    Persos(.Noperso(j)).Aller_A .X(j), .Y(j)
'                Case "PCR": 'Personnage Chercher Ressource.
'                    Persos(.Noperso(j)).Aller_Ressource Ressources(.X(j)), .X(j), AffRessources.Largeur, AffRessources.Hauteur
'                Case "PRC": 'Personnage Rentrer Chateau.
'                    Persos(.Noperso(j)).Rentrer_Chateau Chateaux(.X(j))
'                Case "PCM": 'Personnage Creer Maison.
'                    Maisons(.Noperso(j)).Creer .X(j), .Y(j), _
'                                               .Z(j), .Noperso(j), Parametres
'                    Persos(.Noperso(j)).Rentrer_Maison Maisons(.Noperso(j))
'                Case "PRM": 'Personnage Rentrer Maison.
'                    Persos(.Noperso(j)).Rentrer_Maison Maisons(.X(j))
'                End Select
'            Next j
'        End If
'        End With
'    Next i
'End Sub

' convert a string so that it can be used on a URL query string
' Same effect as the Server.URLEncode method in ASP

Function URLEncode(ByVal Text As String) As String
    Dim i As Integer
    Dim acode As Integer
    Dim char As String
    
    URLEncode = Text
    
    For i = Len(URLEncode) To 1 Step -1
        acode = Asc(Mid$(URLEncode, i, 1))
        Select Case acode
            Case 48 To 57, 65 To 90, 97 To 122
                ' don't touch alphanumeric chars
            Case 32
                ' replace space with "+"
                Mid$(URLEncode, i, 1) = "+"
            Case Else
                ' replace punctuation chars with "%hex"
                URLEncode = Left$(URLEncode, i - 1) & "%" & Hex$(acode) & Mid$ _
                    (URLEncode, i + 1)
        End Select
    Next
    
End Function

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
    'Renvoie le nom du service provider entré en indice.
    ServiceProvider_Nom = dPlayEnumConnexions.GetName(Index)
    'ServiceProvider_Nom = dPlayEnumConnexions(Index).Name
End Property

Public Property Get Partie_Nom(ByVal Index As Integer) As String
    'Renvoie le nom de la partie entrée en indice.
    Partie_Nom = dPlayEnumSession.GetItem(Index).GetSessionName
End Property
Public Property Get Partie_NombreJoueurs(ByVal Index As Integer) As Long
    'Renvoie le nombre de joueurs actuels de la partie entrée en indice.
    Partie_NombreJoueurs = dPlayEnumSession.GetItem(Index).GetCurrentPlayers
End Property
Public Property Get Partie_NombreJoueursMax(ByVal Index As Integer) As Long
    'Renvoie le nombre de joueurs actuels de la partie entrée en indice.
    Partie_NombreJoueursMax = dPlayEnumSession.GetItem(Index).GetMaxPlayers
End Property

Public Property Get Joueur_Nom(ByVal Index As Integer) As String
    'Renvoie le nom du joueur entré en indice.
    Joueur_Nom = dPlayEnumJoueurs.GetShortName(Index)
End Property
Public Property Get Joueur_ID(ByVal Index As Integer) As Long
    'Renvoie l'indentifiant du joueur entré en indice.
    Joueur_ID = dPlayEnumJoueurs.GetDPID(Index)
End Property
Public Property Get Joueur_Serveur(ByVal Index As Integer) As Boolean
    'Renvoie si le joueur est un serveur.
    Joueur_Serveur = dPlayEnumJoueurs.GetFlags(Index) = DPPLAYER_SERVERPLAYER
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
Public Property Let Joueur_InfosAJour(ByVal ID As Long, ByVal valeur As Boolean)
    Dim i As Long
    For i = 0 To UBound(Joueurs())
        If Joueurs(i).ID = ID Then
            Joueurs(i).InfosAJour = valeur
            Exit For
        End If
    Next i
End Property
Public Property Let Joueur_Tous_InfosAJour(ByVal valeur As Boolean)
    Dim i As Long
    For i = 0 To UBound(Joueurs())
        Joueurs(i).InfosAJour = valeur
    Next i
End Property

Public Property Get Definir_Actif() As Boolean
    Definir_Actif = Actif
End Property
Public Property Let Definir_Actif(ByVal valeur As Boolean)
    Actif = valeur
    Changer_Parametre_Ini "Actif", -CDbl(valeur)
End Property
Public Property Get Definir_JouerEnLigne() As Boolean
    Definir_JouerEnLigne = JouerEnLigne
End Property
Public Property Let Definir_JouerEnLigne(ByVal valeur As Boolean)
    JouerEnLigne = valeur
    Changer_Parametre_Ini "JouerEnLigne", -CDbl(valeur)
End Property
Public Property Let Definir_NomJoueur(ByVal valeur As String)
    NomJoueur = valeur
    Changer_Parametre_Ini "NomJoueur", valeur
End Property
Public Property Let Definir_NomSession(ByVal valeur As String)
    NomSession = valeur
    Changer_Parametre_Ini "NomSession", valeur
    'AppDesc.SessionName = Valeur
End Property
Public Property Let Definir_NombreJoueurs(ByVal valeur As Long)
    NombreJoueurs = valeur
    Changer_Parametre_Ini "NombreJoueurs", valeur - 1
    'AppDesc.lMaxPlayers = Valeur
End Property
Public Property Let Definir_ServiceProvider(ByVal valeur As Integer)
    ServiceProvider = valeur
    Changer_Parametre_Ini "ServiceProvider", valeur
End Property
Public Property Get Definir_NumeroPort() As Long
    Definir_NumeroPort = NumeroPort
End Property
Public Property Let Definir_NumeroPort(ByVal valeur As Long)
    NumeroPort = valeur
    Changer_Parametre_Ini "NumeroPort", valeur
End Property

Public Property Get Definir_OuverturePort() As Boolean
    Definir_OuverturePort = OuverturePort
End Property
Public Property Let Definir_OuverturePort(ByVal valeur As Boolean)
    If valeur <> OuverturePort Then
        OuverturePort = valeur
        Changer_Parametre_Ini "OuverturePort", -CDbl(valeur)
        Set dPlayAdress = Nothing: InitAddress
    End If
End Property

Public Property Get Definir_NumeroIPServeur() As String
    Definir_NumeroIPServeur = NumeroIPServeur
End Property
Public Property Let Definir_NumeroIPServeur(ByVal valeur As String)
    NumeroIPServeur = valeur
    Changer_Parametre_Ini "NumeroIPServeur", valeur
End Property

Public Property Get Definir_URL_Serveur() As String
    Definir_URL_Serveur = URL_Serveur
End Property
Public Property Let Definir_URL_Serveur(ByVal valeur As String)
    URL_Serveur = valeur
    Changer_Parametre_Ini "URL_Serveur", valeur
End Property

Private Sub Changer_Parametre_Ini(ByVal Parametre As String, ByVal valeur As String)
    FicIni.fichier = FicIni.chemin & FichierINI
    FicIni.Section = SectionINI
    FicIni.Parametre(Parametre) = valeur
End Sub
