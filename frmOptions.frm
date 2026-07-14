VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5190
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6390
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.TreeView TreeViewOptions 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   7011
      _Version        =   327682
      Indentation     =   212
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Exemple 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Exemple 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Exemple 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Frame FrameOptions 
      Caption         =   "Reseau"
      Height          =   3975
      Index           =   9
      Left            =   2880
      TabIndex        =   85
      Top             =   120
      Width           =   3375
      Begin VB.TextBox TxtURL_Serveur 
         Height          =   285
         Left            =   120
         TabIndex        =   107
         Text            =   "Text1"
         Top             =   1800
         Width           =   3135
      End
      Begin VB.CheckBox CheckOuverturePort 
         Caption         =   "Ouvrir un port de communication. (1)"
         Height          =   255
         Left            =   360
         TabIndex        =   95
         Top             =   1080
         Width           =   2865
      End
      Begin VB.CheckBox CheckReseauActif 
         Caption         =   "Activer le moteur réseau"
         Height          =   255
         Left            =   360
         TabIndex        =   86
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label LblURL_Serveur 
         Caption         =   "Adresse du serveur :"
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Label12"
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   360
         TabIndex        =   88
         Top             =   3120
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   360
         TabIndex        =   87
         Top             =   2520
         Visible         =   0   'False
         Width           =   2775
      End
   End
   Begin VB.Frame FrameOptions 
      Caption         =   "Personnages"
      Height          =   3975
      Index           =   5
      Left            =   2880
      TabIndex        =   43
      Top             =   120
      Width           =   3375
      Begin VB.CheckBox CheckAfficherFantomes 
         Caption         =   "Afficher les fantômes alliés. (2)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   100
         Top             =   1080
         Width           =   2535
      End
      Begin VB.ComboBox ComboPathfinding 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CheckBox CheckAfficherDegats 
         Caption         =   "Afficher les dégats des combats."
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   480
         Width           =   2775
      End
      Begin VB.CheckBox CheckAfficherFantomes 
         Caption         =   "Afficher les fantômes alliés. (2)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   55
         Top             =   1320
         Width           =   2535
      End
      Begin ComctlLib.Slider SliderDureeCommentaires 
         Height          =   255
         Left            =   840
         TabIndex        =   46
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin VB.CheckBox CheckAfficherCommentaires 
         Caption         =   "Afficher les commentaires."
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label15 
         Caption         =   "Definition pathfinding"
         Height          =   255
         Left            =   240
         TabIndex        =   98
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Label Label14 
         Caption         =   "Pathfinding :"
         Height          =   255
         Left            =   240
         TabIndex        =   96
         Top             =   3045
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "s"
         Height          =   255
         Left            =   2660
         TabIndex        =   48
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label LblDureeCommentaires 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         Height          =   255
         Left            =   2280
         TabIndex        =   47
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Durée :"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   2160
         Width           =   975
      End
   End
   Begin VB.Frame FrameOptions 
      Caption         =   "Jeu"
      Height          =   3975
      Index           =   0
      Left            =   2880
      TabIndex        =   36
      Top             =   120
      Width           =   3375
      Begin VB.ComboBox ComboLangues 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   400
         Width           =   1455
      End
      Begin VB.CheckBox CheckAffichageRapide 
         Caption         =   "Désactiver la limite de FPS"
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   3240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin ComctlLib.Slider SliderVitesse 
         Height          =   255
         Left            =   840
         TabIndex        =   37
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin ComctlLib.Slider SliderFPS 
         Height          =   255
         Left            =   840
         TabIndex        =   52
         Top             =   3600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin VB.Label Label10 
         Caption         =   "Langue (1)  :"
         Height          =   255
         Left            =   240
         TabIndex        =   77
         Top             =   480
         Width           =   975
      End
      Begin VB.Label LblFPS 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Left            =   2640
         TabIndex        =   53
         Top             =   3600
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "FPS Max :"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Vitesse :"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label LblVitesse 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Left            =   2880
         TabIndex        =   38
         Top             =   1080
         Width           =   255
      End
   End
   Begin VB.Frame FrameOptions 
      Caption         =   "Affichage"
      Height          =   3975
      Index           =   2
      Left            =   2880
      TabIndex        =   19
      Top             =   120
      Width           =   3375
      Begin VB.CheckBox CheckModeFenetre 
         Caption         =   "ModeFenetre"
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   960
         Width           =   2535
      End
      Begin VB.ComboBox ComboResolution 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   435
         Width           =   1335
      End
      Begin VB.CheckBox CheckOptimiserAffichage 
         Caption         =   "Optimiser l'affichage."
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Résolution (1) (4) :"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame FrameOptions 
      Caption         =   "Interface"
      Height          =   3975
      Index           =   3
      Left            =   2880
      TabIndex        =   11
      Top             =   120
      Width           =   3375
      Begin VB.CheckBox CheckMasquerAutoBoutonsDeplacement 
         Caption         =   "Glisser les boutons de déplacement."
         Height          =   255
         Left            =   240
         TabIndex        =   76
         Top             =   2040
         Width           =   3015
      End
      Begin VB.CheckBox CheckAfficherBoutonsConstruire 
         Caption         =   "Afficher les boutons de construction."
         Height          =   255
         Left            =   240
         TabIndex        =   75
         Top             =   2280
         Width           =   3015
      End
      Begin VB.CheckBox CheckAfficherInfosBulles 
         Caption         =   "Afficher les infos bulles."
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   240
         Width           =   2535
      End
      Begin VB.CheckBox CheckAfficherStastistiques 
         Caption         =   "Afficher les informations des fiefs."
         Height          =   255
         Left            =   240
         TabIndex        =   71
         Top             =   3120
         Width           =   2775
      End
      Begin VB.CheckBox CheckAfficherEffetsTemporaires 
         Caption         =   "Afficher les effets temporaires."
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Top             =   3360
         Width           =   2895
      End
      Begin VB.CheckBox CheckAfficherRaccourciObjets 
         Caption         =   "Afficher les raccourcis des objets."
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   3600
         Width           =   3015
      End
      Begin VB.CheckBox CheckMasquerAutoBoutonsConstruire 
         Caption         =   "Glisser les boutons de construction."
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   2520
         Width           =   3015
      End
      Begin VB.CheckBox CheckAfficherInfoBulleTerrain 
         Caption         =   "Afficher les infos bulles sur le terrain."
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   480
         Width           =   2895
      End
      Begin VB.CheckBox CheckAfficherMessageModeAuto 
         Caption         =   "Afficher le mode auto."
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox CheckAfficherBoutonsDeplacement 
         Caption         =   "Afficher les boutons de déplacement."
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1800
         Width           =   3015
      End
      Begin VB.CheckBox CheckAfficherBoutonsInfos 
         Caption         =   "Afficher les boutons d'information."
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   2775
      End
      Begin VB.CheckBox CheckAfficherBarresStatut 
         Caption         =   "Afficher les barres de statut."
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox CheckAfficherAction 
         Caption         =   "Afficher les actions."
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   2535
      End
      Begin VB.CheckBox CheckAfficherFPS 
         Caption         =   "Afficher les frames par secondes."
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox CheckAfficherEpoques 
         Caption         =   "Afficher les époques"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2880
         Width           =   3015
      End
   End
   Begin VB.Frame FrameOptions 
      Caption         =   "Messages d'information"
      Height          =   3975
      Index           =   4
      Left            =   2880
      TabIndex        =   22
      Top             =   120
      Width           =   3375
      Begin VB.Frame FrameDetailsMessages 
         Caption         =   "Détails des messages d'information"
         Height          =   3135
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   3135
         Begin VB.CheckBox CheckMessagesAfficherEmpereur 
            Caption         =   "Afficher nouveau/mort empereur."
            Height          =   255
            Left            =   240
            TabIndex        =   101
            Top             =   2760
            Width           =   2775
         End
         Begin VB.CheckBox CheckMessagesAfficherSoldats 
            Caption         =   "Afficher gains/pertes soldats."
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   2160
            Width           =   2535
         End
         Begin VB.CheckBox CheckMessagesAfficherObjets 
            Caption         =   "Afficher gains/pertes objets."
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   1920
            Width           =   2535
         End
         Begin VB.CommandButton CmdConseils 
            Caption         =   "&Consulter les conseils"
            Height          =   495
            Left            =   1920
            TabIndex        =   31
            Top             =   320
            Width           =   1095
         End
         Begin VB.CheckBox CheckMessagesAfficherConseil 
            Caption         =   "Afficher un conseil en début de partie."
            Height          =   375
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox CheckMessagesAfficherExperience 
            Caption         =   "Afficher niveaux supérieures."
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   2400
            Width           =   2415
         End
         Begin VB.CheckBox CheckMessagesAfficherRessources 
            Caption         =   "Afficher gains/pertes ressources."
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   1680
            Width           =   2655
         End
         Begin VB.CheckBox CheckMessagesAfficherArgent 
            Caption         =   "Afficher gains/pertes argent."
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   1440
            Width           =   2415
         End
         Begin VB.CheckBox CheckMessagesAfficherBestiaire 
            Caption         =   "Afficher combats gagnés."
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CheckBox CheckMessagesAfficherVie 
            Caption         =   "Afficher Vie/Mort."
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   840
            Width           =   2535
         End
      End
      Begin VB.CheckBox CheckMessagesAfficher 
         Caption         =   "Afficher les messages d'information."
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame FrameOptions 
      Caption         =   "Terrain"
      Height          =   3975
      Index           =   6
      Left            =   2880
      TabIndex        =   56
      Top             =   120
      Width           =   3375
      Begin VB.CheckBox CheckAfficherTemps 
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   84
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CheckBox CheckTerrainAnime 
         Caption         =   "Surface animée. (1) (3)"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame FrameOptions 
      Caption         =   "Regles"
      Height          =   3975
      Index           =   1
      Left            =   2880
      TabIndex        =   92
      Top             =   120
      Width           =   3375
      Begin ComctlLib.Slider SliderRegle0 
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   960
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   327682
         Max             =   20
         SelStart        =   10
         Value           =   10
      End
      Begin VB.CheckBox CheckAfficherProjectiles 
         Caption         =   "Afficher les projectiles. (1)"
         Height          =   255
         Left            =   240
         TabIndex        =   93
         Top             =   2040
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CheckBox CheckRegles 
         Caption         =   "Regle2"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   94
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label16 
         Caption         =   "Label16"
         Height          =   255
         Left            =   3140
         TabIndex        =   105
         Top             =   980
         Width           =   175
      End
      Begin VB.Label LblRegle0 
         Caption         =   "Label17"
         Height          =   615
         Left            =   240
         TabIndex        =   103
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label LblValRegle0 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label16"
         Height          =   255
         Left            =   2760
         TabIndex        =   102
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.Frame FrameOptions 
      Caption         =   "Sons"
      Height          =   3975
      Index           =   7
      Left            =   2880
      TabIndex        =   9
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton CmdBruitages 
         Caption         =   "Command1"
         Height          =   495
         Left            =   2160
         TabIndex        =   99
         Top             =   360
         Width           =   1095
      End
      Begin VB.Frame Frame10 
         Caption         =   "Lecteur de musique (1)"
         Height          =   1215
         Left            =   240
         TabIndex        =   68
         Top             =   1920
         Width           =   3015
         Begin VB.OptionButton OptionLecteurMusique 
            Caption         =   "DirectMusic.             (Rapide, ne lit que les MIDI)"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   70
            Top             =   720
            Width           =   2295
         End
         Begin VB.OptionButton OptionLecteurMusique 
            Caption         =   "Windows Media Player.     (Lent, lit les MIDI, WAVE, MP3)"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   69
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.CommandButton CmdMusiques 
         Caption         =   "Accéder aux &musiques"
         Height          =   495
         Left            =   2160
         TabIndex        =   41
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox CheckJouerBruitages 
         Caption         =   "Jouer les bruitages."
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox CheckJouerMusiques 
         Caption         =   "Jouer les musiques."
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   2415
      End
   End
   Begin VB.Frame FrameOptions 
      Caption         =   "Sauvegardes"
      Height          =   3975
      Index           =   10
      Left            =   2880
      TabIndex        =   33
      Top             =   120
      Width           =   3375
      Begin VB.CheckBox CheckSavImageEnQuittant 
         Caption         =   "Sauvegarder un aperçu de la partie."
         Height          =   255
         Left            =   240
         TabIndex        =   82
         Top             =   1920
         Width           =   3015
      End
      Begin VB.CheckBox CheckSAVPersoPeriodique 
         Caption         =   "Sauvegarder le personnage actif toutes les minutes."
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox CheckSavMondeEnQuittant 
         Caption         =   "Sauvegarder la partie en quittant."
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1440
         Width           =   2775
      End
   End
   Begin VB.Frame FrameOptions 
      Caption         =   "Stastistiques"
      Height          =   3975
      Index           =   11
      Left            =   2880
      TabIndex        =   59
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton CommandConsulterStastistiques 
         Caption         =   "Consulter les stastistiques."
         Height          =   495
         Left            =   600
         TabIndex        =   83
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ComboBox ComboNavigateurWeb 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   2400
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox CheckEnregistrerStastistiques 
         Caption         =   "Enregistrer les stastistiques."
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   480
         Width           =   2775
      End
      Begin VB.CheckBox CheckAfficherStastistiquesEnQuittant 
         Caption         =   "Afficher en fin de partie."
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Left            =   240
         TabIndex        =   79
         Top             =   2445
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame FrameOptions 
      Caption         =   "Cartes"
      Height          =   3975
      Index           =   8
      Left            =   2880
      TabIndex        =   89
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton CmdCartes 
         Caption         =   "Accéder aux &Cartes"
         Height          =   495
         Left            =   1080
         TabIndex        =   91
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox CheckCreerFichierInformations 
         Caption         =   "Créer le &fichier ""Informations.ini"""
         Height          =   255
         Left            =   240
         TabIndex        =   90
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Label Label9 
      Caption         =   "(4) Certaines résolutions s'affichent mal quand ""Optimisier affichage"" est sélectionné."
      Height          =   255
      Left            =   0
      TabIndex        =   73
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Label LblRedemarrer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Redémarrez le jeu pour activer les modifications."
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   4080
      TabIndex        =   65
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "(2) Attention, cette option peut causer des ralentissments selon le nombre de fantômes affichés."
      Height          =   495
      Left            =   120
      TabIndex        =   63
      Top             =   4200
      Width           =   3975
   End
   Begin VB.Label Label6 
      Caption         =   "(1) Si cette option est modifiée, relancez le programme pour prendre en compte les changements."
      Height          =   495
      Left            =   120
      TabIndex        =   62
      Top             =   4200
      Width           =   3975
   End
   Begin VB.Label Label8 
      Caption         =   "(3) Cette option ralentit le chargement de l'optimisation de l'affichage et lui fait utiliser plus de mémoire."
      Height          =   495
      Left            =   120
      TabIndex        =   64
      Top             =   4680
      Width           =   3975
   End
   Begin VB.Menu MenuOptions 
      Caption         =   "O&ptions"
      Visible         =   0   'False
      Begin VB.Menu MenuOptionsAnnuler 
         Caption         =   "&Annuler"
      End
      Begin VB.Menu MenuOptionsOk 
         Caption         =   "&Ok"
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Const FichierINI = "Fenetre_Options"
Const SectionINI = "Fenetre_Options"

Const NombreCadres = 12
Const CadreJeu = 0
Const CadreRegles = 1
Const CadreAffichage = 2
Const CadreInterface = 3
Const CadreInformations = 4
Const CadrePersonnages = 5
Const CadreTerrain = 6
Const CadreSons = 7
Const CadreCartes = 8
Const CadreReseau = 9
Const CadreSauvegardes = 10
Const CadreStastistiques = 11

Dim TempBouton As ClsIntBouton

'Private Sub CheckOptimiserAffichage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If CheckOptimiserAffichage.Value = 1 Then
'        MsgBox "Attention, cette option utilise beaucoup de mémoire" & Chr(10) & _
'               "et ralentit énormément le temps de chargement." & Chr(10) & Chr(10) & _
'               "Il est préférable de désactiver la surface animée des terrains." & Chr(10) & _
'               "Si vous rencontrez des problèmes, décochez cette option.", _
'               vbInformation, "Avertissement sur : " & CheckOptimiserAffichage.Caption
'    End If
'End Sub

Private Sub CheckAfficherFantomes_Click(Index As Integer)
    If Index = 0 Then
        CheckAfficherFantomes(1).Enabled = -CDbl(CheckAfficherFantomes(0).Value)
    End If
End Sub

Private Sub CheckAfficherProjectiles_Click()
    LblRedemarrer.Visible = True
End Sub

Private Sub CheckEnregistrerStastistiques_Click()
    CheckAfficherStastistiquesEnQuittant.Enabled = CheckEnregistrerStastistiques.Value = 1
    If Not CheckAfficherStastistiquesEnQuittant.Enabled Then
        CheckAfficherStastistiquesEnQuittant.Value = 0
    End If
End Sub

Private Sub CheckModeFenetre_Click()
    LblRedemarrer.Visible = True
End Sub

Private Sub CheckOuverturePort_Click()
    LblRedemarrer.Visible = True
End Sub

Private Sub CheckRegles_Click(Index As Integer)
    If Index = 1 Then
        LblRedemarrer.Visible = True
    End If
End Sub

Private Sub CheckReseauActif_Click()
    LblRedemarrer.Visible = True
End Sub

Private Sub CheckTerrainAnime_Click()
    LblRedemarrer.Visible = True
End Sub

Private Sub CmdBruitages_Click()
    Editer_Repertoire App.Path & Son.chemin
End Sub

Private Sub CmdCartes_Click()
    Editer_Repertoire App.Path & Cartes.chemin
End Sub

Private Sub CmdMusiques_Click()
    Editer_Repertoire App.Path & Musique.chemin
End Sub

Private Sub ComboLangues_Click()
    LblRedemarrer.Visible = True
End Sub

Private Sub ComboResolution_Click()
    LblRedemarrer.Visible = True
End Sub

Private Sub CommandConsulterStastistiques_Click()
    Afficher_Stastistiques
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Set TempBouton = New ClsIntBouton
    FrmParam.Hide
    FrmInfoRaces.Hide
    
    FicIni.fichier = FicIni.chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI
    
    'Charge le fichier de langues.
    TreeViewOptions.Nodes.Add , , "Jeu", FicIni.Parametre("BandeJeu")
    TreeViewOptions.Nodes.Add , , "Regles", FicIni.Parametre("BandeRegles")
    TreeViewOptions.Nodes.Add , , "Affichage", FicIni.Parametre("BandeAffichage")
    TreeViewOptions.Nodes.Add , , "Interface", FicIni.Parametre("BandeInterface")
    TreeViewOptions.Nodes.Add , , "Informations", FicIni.Parametre("BandeInformations")
    TreeViewOptions.Nodes.Add , , "Personnages", FicIni.Parametre("BandePersonnages")
    TreeViewOptions.Nodes.Add , , "Terrain", FicIni.Parametre("BandeTerrain")
    TreeViewOptions.Nodes.Add , , "Sons", FicIni.Parametre("BandeSons")
    TreeViewOptions.Nodes.Add , , "Cartes", FicIni.Parametre("BandeCartes")
    TreeViewOptions.Nodes.Add , , "Reseau", FicIni.Parametre("BandeReseau")
    TreeViewOptions.Nodes.Add , , "Sauvegardes", FicIni.Parametre("BandeSauvegardes")
    TreeViewOptions.Nodes.Add , , "Stastistiques", FicIni.Parametre("BandeStastistiques")
    Set TreeViewOptions.Nodes.Item("Terrain").parent = TreeViewOptions.Nodes.Item("Affichage")
    Set TreeViewOptions.Nodes.Item("Personnages").parent = TreeViewOptions.Nodes.Item("Affichage")
    Set TreeViewOptions.Nodes.Item("Informations").parent = TreeViewOptions.Nodes.Item("Affichage")
    Set TreeViewOptions.Nodes.Item("Interface").parent = TreeViewOptions.Nodes.Item("Affichage")
    Set TreeViewOptions.Nodes.Item("Reseau").parent = TreeViewOptions.Nodes.Item("Jeu")
    Set TreeViewOptions.Nodes.Item("Cartes").parent = TreeViewOptions.Nodes.Item("Jeu")
    Set TreeViewOptions.Nodes.Item("Sons").parent = TreeViewOptions.Nodes.Item("Jeu")
    Set TreeViewOptions.Nodes.Item("Affichage").parent = TreeViewOptions.Nodes.Item("Jeu")
    Set TreeViewOptions.Nodes.Item("Regles").parent = TreeViewOptions.Nodes.Item("Jeu")
    TreeViewOptions.Nodes.Item("Jeu").Expanded = True
    TreeViewOptions.Nodes.Item("Affichage").Expanded = True
    TreeViewOptions.SelectedItem = TreeViewOptions.Nodes.Item("Jeu")
    TreeViewOptions_NodeClick TreeViewOptions.Nodes.Item("Jeu")
    
    Caption = FicIni.Parametre("EtiquetteNomFenetre")
    
    FrameOptions(CadreJeu).Caption = FicIni.Parametre("BandeJeu")
    Label3.Caption = FicIni.Parametre("EtiquetteVitesse")
    Label10.Caption = FicIni.Parametre("EtiquetteLangue")
    
    FrameOptions(CadreRegles).Caption = FicIni.Parametre("BandeRegles")
    'CheckRegles(0).Caption = FicIni.Parametre("EtiquetteRegle1")
    LblRegle0.Caption = FicIni.Parametre("EtiquetteRegle1")
    Label16.Caption = Parametres.EtiquettePourcent
    CheckRegles(1).Caption = FicIni.Parametre("EtiquetteRegle2")
    CheckAfficherProjectiles.Caption = FicIni.Parametre("EtiquettePersonnagesProjectiles")
    
    FrameOptions(CadreAffichage).Caption = FicIni.Parametre("BandeAffichage")
    CheckOptimiserAffichage.Caption = FicIni.Parametre("EtiquetteOptimiserAffichage")
    CheckAfficherTemps.Caption = FicIni.Parametre("EtiquetteAfficherTemps")
    Label1.Caption = FicIni.Parametre("EtiquetteResolutionsAffichage")
    CheckModeFenetre.Caption = FicIni.Parametre("EtiquetteModeFenetre")
    
    FrameOptions(CadreSons).Caption = FicIni.Parametre("BandeSons")
    CheckJouerBruitages.Caption = FicIni.Parametre("EtiquetteBruitages")
    CmdBruitages.Caption = FicIni.Parametre("BoutonBruitages")
    CheckJouerMusiques.Caption = FicIni.Parametre("EtiquetteMusiques")
    CmdMusiques.Caption = FicIni.Parametre("BoutonMusiques")
    Frame10.Caption = FicIni.Parametre("BandeLecteurMusique")
    OptionLecteurMusique(0).Caption = FicIni.Parametre("OptionWMP")
    OptionLecteurMusique(1).Caption = FicIni.Parametre("OptionDM")
    
    FrameOptions(CadreStastistiques).Caption = FicIni.Parametre("BandeStastistiques")
    CheckEnregistrerStastistiques.Caption = FicIni.Parametre("EtiquetteEnregistrerStastistiques")
    CommandConsulterStastistiques.Caption = FicIni.Parametre("BoutonConsulterStastistiques")
    CheckAfficherStastistiquesEnQuittant.Caption = FicIni.Parametre("EtiquetteAfficherStastistiques")
    Label11.Caption = FicIni.Parametre("EtiquetteNavigateurWeb")
    ComboNavigateurWeb.AddItem LogicielIExplorer
    ComboNavigateurWeb.AddItem LogicielFireFox
    ComboNavigateurWeb.ListIndex = Definir_NavigateurWeb
    
    FrameOptions(CadreInformations).Caption = FicIni.Parametre("BandeInformations")
    CheckMessagesAfficher.Caption = FicIni.Parametre("EtiquetteAfficherInformations")
    FrameDetailsMessages.Caption = FicIni.Parametre("BandeDetailsInformations")
    CheckMessagesAfficherConseil.Caption = FicIni.Parametre("EtiquetteInformationsConseil")
    CmdConseils.Caption = FicIni.Parametre("BoutonInformationsConseil")
    CheckMessagesAfficherVie.Caption = FicIni.Parametre("EtiquetteInformationsVie")
    CheckMessagesAfficherBestiaire.Caption = FicIni.Parametre("EtiquetteInformationsCombat")
    CheckMessagesAfficherArgent.Caption = FicIni.Parametre("EtiquetteInformationsArgent")
    CheckMessagesAfficherRessources.Caption = FicIni.Parametre("EtiquetteInformationsRessources")
    CheckMessagesAfficherObjets.Caption = FicIni.Parametre("EtiquetteInformationsObjets")
    CheckMessagesAfficherSoldats.Caption = FicIni.Parametre("EtiquetteInformationsSoldats")
    CheckMessagesAfficherExperience.Caption = FicIni.Parametre("EtiquetteInformationsNiveaux")
    CheckMessagesAfficherEmpereur.Caption = FicIni.Parametre("EtiquetteInformationsEmpereur")
    
    FrameOptions(CadreSauvegardes).Caption = FicIni.Parametre("BandeSauvegardes")
    CheckSAVPersoPeriodique.Caption = FicIni.Parametre("EtiquetteSauvegardePersonnage")
    CheckSavMondeEnQuittant.Caption = FicIni.Parametre("EtiquetteSauvegardePartie")
    CheckSavImageEnQuittant.Caption = FicIni.Parametre("EtiquetteSauvegardeImagePartie")
    
    FrameOptions(CadreTerrain).Caption = FicIni.Parametre("BandeTerrain")
    CheckTerrainAnime.Caption = FicIni.Parametre("EtiquetteTerrainAnimation")
    
    FrameOptions(CadrePersonnages).Caption = FicIni.Parametre("BandePersonnages")
    Label14.Caption = FicIni.Parametre("EtiquettePathfinding1")
    Label15.Caption = FicIni.Parametre("EtiquettePathfinding2")
    ComboPathfinding.AddItem FicIni.Parametre("OptionPathfinding1")
    ComboPathfinding.AddItem FicIni.Parametre("OptionPathfinding2")
    ComboPathfinding.AddItem FicIni.Parametre("OptionPathfinding3")
    CheckAfficherDegats.Caption = FicIni.Parametre("EtiquettePersonnagesDegats")
    CheckAfficherFantomes(0).Caption = FicIni.Parametre("EtiquettePersonnagesFantomes1")
    CheckAfficherFantomes(1).Caption = FicIni.Parametre("EtiquettePersonnagesFantomes2")
    CheckAfficherCommentaires.Caption = FicIni.Parametre("EtiquettePersonnagesCommentaires")
    Label2.Caption = FicIni.Parametre("EtiquettePersonnagesCommentairesDuree")
    Label4.Caption = FicIni.Parametre("EtiquettePersonnagesCommentairesTemps")
    
    FrameOptions(CadreInterface).Caption = FicIni.Parametre("BandeInterface")
    CheckAfficherMessageModeAuto.Caption = FicIni.Parametre("EtiquetteInterfaceModeAuto")
    CheckAfficherInfosBulles.Caption = FicIni.Parametre("EtiquetteInterfaceInfoBulles")
    CheckAfficherFPS.Caption = FicIni.Parametre("EtiquetteInterfaceFPS")
    CheckAfficherAction.Caption = FicIni.Parametre("EtiquetteInterfaceActions")
    CheckAfficherInfoBulleTerrain.Caption = FicIni.Parametre("EtiquetteInterfaceInfobullesTerrain")
    CheckAfficherBarresStatut.Caption = FicIni.Parametre("EtiquetteInterfaceStatuts")
    CheckAfficherBoutonsInfos.Caption = FicIni.Parametre("EtiquetteInterfaceBoutonsInfos")
    CheckAfficherBoutonsDeplacement.Caption = FicIni.Parametre("EtiquetteInterfaceBoutonsDeplacement")
    CheckMasquerAutoBoutonsDeplacement.Caption = FicIni.Parametre("EtiquetteInterfaceBoutonsDeplacementGlisser")
    CheckAfficherBoutonsConstruire.Caption = FicIni.Parametre("EtiquetteInterfaceBoutonsConstruction")
    CheckMasquerAutoBoutonsConstruire.Caption = FicIni.Parametre("EtiquetteInterfaceBoutonsConstructionGlisser")
    CheckAfficherStastistiques.Caption = FicIni.Parametre("EtiquetteInterfaceInfosFiefs")
    CheckAfficherEffetsTemporaires.Caption = FicIni.Parametre("EtiquetteInterfaceEffetsTemporaires")
    CheckAfficherRaccourciObjets.Caption = FicIni.Parametre("EtiquetteInterfaceRaccourcisObjets")
    CheckAfficherEpoques.Caption = FicIni.Parametre("EtiquetteInterfaceChangementEpoque")
    
    FrameOptions(CadreReseau).Caption = FicIni.Parametre("BandeReseau")
    CheckReseauActif.Caption = FicIni.Parametre("EtiquetteReseauActif")
    CheckOuverturePort.Caption = FicIni.Parametre("EtiquetteOuverturePort")
    LblURL_Serveur.Caption = FicIni.Parametre("EtiquetteURL_Serveur")
    Label12.Caption = FrmParam.Label44.Caption
    Label13.Caption = FrmParam.Label45.Caption
    
    FrameOptions(CadreCartes).Caption = FicIni.Parametre("BandeCartes")
    CheckCreerFichierInformations.Caption = FicIni.Parametre("EtiquetteCreerFichierInformations")
    CmdCartes.Caption = FicIni.Parametre("BoutonCartes")
    
    Label6.Caption = FicIni.Parametre("EtiquetteNotes1")
    Label7.Caption = FicIni.Parametre("EtiquetteNotes2")
    Label8.Caption = FicIni.Parametre("EtiquetteNotes3")
    Label9.Caption = FicIni.Parametre("EtiquetteNotes4")
    
    LblRedemarrer.Caption = FicIni.Parametre("EtiquetteAvertissementRedemarrer")
    
    cmdOK.Caption = FicIni.Parametre("BoutonOK")
    cmdCancel.Caption = FicIni.Parametre("BoutonAnnuler")
    
    'Centre la feuille.
    Move (MDIFrmMain.Width - Me.Width) / 2, _
         (MDIFrmMain.Height - Me.Height) / 2
    
    'LblValRegle0.Caption = Jeu.Definir_Regle(0)
    SliderRegle0.Value = Jeu.Definir_Regle(0) / 10
    SliderRegle0_Scroll
    CheckRegles(1).Value = Jeu.Definir_Regle(1)
    
    Interface(0).Affichage_PleinEcran = False
    
    CheckAfficherEpoques.Value = -CDbl(Interface(0).Definir_AfficherEpoques)
    CheckAfficherFPS.Value = -CDbl(Interface(0).Definir_AfficherFPS)
    CheckAfficherAction.Value = -CDbl(Interface(0).Definir_AfficherAction)
    CheckAfficherInfoBulleTerrain = -CDbl(Interface(0).Definir_AfficherInfoBulleTerrain)
    CheckAfficherBarresStatut.Value = -CDbl(Interface(0).Definir_AfficherBarresStatut)
    CheckAfficherBoutonsInfos.Value = -CDbl(Interface(0).Definir_AfficherBoutonsInfos)
    CheckAfficherBoutonsDeplacement.Value = -CDbl(Interface(0).Definir_AfficherBoutonsDeplacement)
    CheckAfficherBoutonsConstruire.Value = -CDbl(Interface(0).Definir_AfficherBoutonsConstruire)
    CheckMasquerAutoBoutonsDeplacement = -CDbl(Interface(0).Definir_MasquerAutoBoutonsDeplacement)
    CheckMasquerAutoBoutonsConstruire = -CDbl(Interface(0).Definir_MasquerAutoBoutonsConstruire)
    CheckAfficherMessageModeAuto.Value = -CDbl(Interface(0).Definir_AfficherMessageModeAuto)
    CheckAfficherRaccourciObjets = -CDbl(Interface(0).Definir_AfficherRaccourciObjets)
    CheckAfficherEffetsTemporaires = -CDbl(Interface(0).Definir_AfficherEffetsTemporaires)
    
    CheckAfficherCommentaires.Value = -CDbl(Commentaires.Definir_Afficher)
    SliderDureeCommentaires.Max = Commentaires.Definir_DureeMax / Commentaires.Definir_DureeIntervale
    SliderDureeCommentaires.Min = Commentaires.Definir_DureeMin / Commentaires.Definir_DureeIntervale
    SliderDureeCommentaires.Value = Commentaires.Definir_Duree / Commentaires.Definir_DureeIntervale
    SliderDureeCommentaires_Scroll
    
    ComboPathfinding.ListIndex = Jeu.Regle_Pathfinding
    CheckAfficherProjectiles = -CDbl(Jeu.Definir_AfficherProjectiles)
    CheckAfficherFantomes(0).Value = -CDbl(Jeu.Definir_AfficherFantomes)
    CheckAfficherFantomes(1).Value = -CDbl(Jeu.Definir_AfficherFantomesPNJ)
    CheckAfficherFantomes_Click 0
    CheckAfficherInfosBulles.Value = -CDbl(TempBouton.Definir_AfficherInfoBulle)
        
    CheckTerrainAnime = -CDbl(Monde.Definir_TerrainAnime)
        
    CheckJouerBruitages.Value = -CDbl(Son.Definir_Jouer)
    CheckJouerMusiques.Value = -CDbl(Musique.Definir_JouerMusique)
    
    OptionLecteurMusique(1).Value = Musique.Definir_UtiliserDirectMusic
    OptionLecteurMusique(0).Value = Not OptionLecteurMusique(1).Value
    
    CheckSAVPersoPeriodique.Value = -CDbl(Definir_SAVPersoPeriodique)
    CheckSavMondeEnQuittant.Value = -CDbl(Definir_SavMondeEnQuittant)
    CheckSavImageEnQuittant.Value = -CDbl(Definir_SavImageEnQuittant)
    
    CheckEnregistrerStastistiques = -CDbl(Definir_EnregistrerStastistiques)
    CheckAfficherStastistiquesEnQuittant = -CDbl(Definir_AfficherStastistiquesEnQuittant)
    CheckEnregistrerStastistiques_Click
    
    'Messsages d'informations.
    CheckAfficherDegats = -CDbl(AffDegat.Definir_AfficherDegats)
    CheckAfficherStastistiques = -CDbl(Interface(0).Definir_AfficherStastistiques)
    CheckMessagesAfficher.Value = -CDbl(Messages(0).Definir_Afficher)
    CheckMessagesAfficher_Click
    CheckMessagesAfficherConseil = -CDbl(Messages(0).Definir_AfficherConseil)
    CheckMessagesAfficherVie.Value = -CDbl(Messages(0).Definir_AfficherVie)
    CheckMessagesAfficherBestiaire.Value = -CDbl(Messages(0).Definir_AfficherBestiaire)
    CheckMessagesAfficherArgent.Value = -CDbl(Messages(0).Definir_AfficherArgent)
    CheckMessagesAfficherRessources.Value = -CDbl(Messages(0).Definir_AfficherRessources)
    CheckMessagesAfficherObjets.Value = -CDbl(Messages(0).Definir_AfficherObjets)
    CheckMessagesAfficherSoldats.Value = -CDbl(Messages(0).Definir_AfficherSoldats)
    CheckMessagesAfficherExperience.Value = -CDbl(Messages(0).Definir_AfficherExperience)
    CheckMessagesAfficherEmpereur.Value = -CDbl(Messages(0).Definir_AfficherEmpereur)
    
    'Vitesse de jeu.
    SliderVitesse.Min = Jeu.VitesseMin
    SliderVitesse.Max = Jeu.VitesseMax
    SliderVitesse = Jeu.Definir_Vitesse
    SliderVitesse_Scroll
    'FPS.
    SliderFPS.Min = Jeu.FPSMin
    SliderFPS.Max = Jeu.FPSMax
    SliderFPS = Jeu.Definir_FPS
    SliderFPS_Scroll
    CheckAffichageRapide = -CDbl(Jeu.Definir_AffichageRapide)
    CheckAffichageRapide_Click
    
    'Choix de la langue.
    For i = 0 To Langues.Nombre - 1
        ComboLangues.AddItem Langues.Definir_Nom(i)
    Next i
    ComboLangues.ListIndex = Langues.Definir_Langue
    
    'Résolutions d'écran.
    For i = 0 To Ecran.Nombre_Resolutions - 1
        ComboResolution.AddItem Ecran.Definir_ResolutionX(i) & " * " & Ecran.Definir_ResolutionY(i)
    Next i
    ComboResolution.ListIndex = Ecran.Resolution
    
    'Optimisation d'affichage.
    CheckOptimiserAffichage = -CDbl(Monde.Definir_TerrainOptimise)
    CheckAfficherTemps = -CDbl(AffTemps.Definir_Afficher)
    If Ecran.Definir_ModeFenetre Then
        CheckModeFenetre = 0
    Else
        CheckModeFenetre = 1
    End If
    
    'Options réseau.
    CheckReseauActif = -CDbl(ComReseau.Definir_Actif)
    CheckOuverturePort = -CDbl(ComReseau.Definir_OuverturePort)
    TxtURL_Serveur = ComReseau.Definir_URL_Serveur
    
    'Options cartes.
    CheckCreerFichierInformations = -CDbl(Cartes.Definir_CreerFichierInformations)
    
    LblRedemarrer.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If FrmChargement.Afficher Then
        FrmChargement.Show
    Else
        FrmParam.Show
    End If
End Sub

Private Sub MenuOptionsAnnuler_Click()
    cmdCancel_Click
End Sub

Private Sub MenuOptionsOk_Click()
    cmdOK_Click
End Sub

Private Sub OptionLecteurMusique_Click(Index As Integer)
    LblRedemarrer.Visible = True
End Sub

Private Sub SliderDureeCommentaires_Scroll()
    LblDureeCommentaires = Format(SliderDureeCommentaires * Commentaires.Definir_DureeIntervale, "0.00")
End Sub

Private Sub SliderFPS_Scroll()
    LblFPS = SliderFPS
End Sub

Private Sub SliderRegle0_Scroll()
    LblValRegle0.Caption = SliderRegle0 * 10
End Sub

Private Sub SliderVitesse_Scroll()
    LblVitesse = SliderVitesse
End Sub

Private Sub CheckAffichageRapide_Click()
    Label5.Visible = CheckAffichageRapide.Value <> 1
    SliderFPS.Visible = CheckAffichageRapide.Value <> 1
    LblFPS.Visible = CheckAffichageRapide.Value <> 1
End Sub

Private Sub CheckMessagesAfficher_Click()
    FrameDetailsMessages.Enabled = CheckMessagesAfficher.Value = 1
    CheckMessagesAfficherConseil.Enabled = FrameDetailsMessages.Enabled
    CheckMessagesAfficherVie.Enabled = FrameDetailsMessages.Enabled
    CheckMessagesAfficherBestiaire.Enabled = FrameDetailsMessages.Enabled
    CheckMessagesAfficherArgent.Enabled = FrameDetailsMessages.Enabled
    CheckMessagesAfficherObjets.Enabled = FrameDetailsMessages.Enabled
    CheckMessagesAfficherRessources.Enabled = FrameDetailsMessages.Enabled
    CheckMessagesAfficherSoldats.Enabled = FrameDetailsMessages.Enabled
    CheckMessagesAfficherExperience.Enabled = FrameDetailsMessages.Enabled
    CheckMessagesAfficherEmpereur.Enabled = FrameDetailsMessages.Enabled
End Sub

Private Sub CmdConseils_Click()
    Editer_Texte App.Path & Messages(0).Fichier_Conseil
End Sub

Private Sub cmdOK_Click()
    MousePointer = 11
    Interface(0).Definir_AfficherEpoques = CBool(CheckAfficherEpoques.Value)
    Interface(0).Definir_AfficherFPS = CBool(CheckAfficherFPS.Value)
    Interface(0).Definir_AfficherAction = CBool(CheckAfficherAction.Value)
    Interface(0).Definir_AfficherInfoBulleTerrain = CBool(CheckAfficherInfoBulleTerrain.Value)
    Interface(0).Definir_AfficherBarresStatut = CBool(CheckAfficherBarresStatut.Value)
    Interface(0).Definir_AfficherBoutonsInfos = CBool(CheckAfficherBoutonsInfos.Value)
    Interface(0).Definir_AfficherBoutonsDeplacement = CBool(CheckAfficherBoutonsDeplacement.Value)
    Interface(0).Definir_AfficherBoutonsConstruire = CBool(CheckAfficherBoutonsConstruire.Value)
    Interface(0).Definir_MasquerAutoBoutonsDeplacement = CBool(CheckMasquerAutoBoutonsDeplacement.Value)
    Interface(0).Definir_MasquerAutoBoutonsConstruire = CBool(CheckMasquerAutoBoutonsConstruire.Value)
    Interface(0).Definir_AfficherMessageModeAuto = CBool(CheckAfficherMessageModeAuto.Value)
    Interface(0).Definir_AfficherRaccourciObjets = CBool(CheckAfficherRaccourciObjets.Value)
    Interface(0).Definir_AfficherEffetsTemporaires = CBool(CheckAfficherEffetsTemporaires.Value)
    
    Jeu.Definir_Regle(0) = LblValRegle0.Caption 'CheckRegles(0).Value
    Jeu.Definir_Regle(1) = CheckRegles(1).Value
    Jeu.Definir_Regle(2) = ComboPathfinding.ListIndex
    Jeu.Definir_AfficherProjectiles = CBool(CheckAfficherProjectiles.Value)
    Jeu.Definir_AfficherFantomes = CBool(CheckAfficherFantomes(0).Value)
    Jeu.Definir_AfficherFantomesPNJ = CBool(CheckAfficherFantomes(1).Value)
    Commentaires.Definir_Afficher = CBool(CheckAfficherCommentaires.Value)
    Commentaires.Definir_Duree = SliderDureeCommentaires.Value * Commentaires.Definir_DureeIntervale
    TempBouton.Definir_AfficherInfoBulle = CBool(CheckAfficherInfosBulles.Value)
    
    Monde.Definir_TerrainAnime = CBool(CheckTerrainAnime.Value)
    
    Son.Definir_Jouer = CBool(CheckJouerBruitages.Value)
    Musique.Definir_JouerMusique = CBool(CheckJouerMusiques.Value)
    MDIFrmMain.MenuOptionMusique.Checked = Musique.Definir_JouerMusique
    
    Musique.Definir_UtiliserDirectMusic = OptionLecteurMusique(1).Value
    
    Definir_SAVPersoPeriodique = CBool(CheckSAVPersoPeriodique.Value)
    Definir_SavMondeEnQuittant = CBool(CheckSavMondeEnQuittant.Value)
    Definir_SavImageEnQuittant = CBool(CheckSavImageEnQuittant.Value)
    
    Definir_EnregistrerStastistiques = CBool(CheckEnregistrerStastistiques)
    Definir_AfficherStastistiquesEnQuittant = CBool(CheckAfficherStastistiquesEnQuittant)
    Definir_NavigateurWeb = ComboNavigateurWeb.ListIndex
    
    AffDegat.Definir_AfficherDegats = CBool(CheckAfficherDegats)
    Interface(0).Definir_AfficherStastistiques = CBool(CheckAfficherStastistiques)
    Messages(0).Definir_Afficher = CBool(CheckMessagesAfficher.Value)
    Messages(0).Definir_AfficherConseil = CBool(CheckMessagesAfficherConseil.Value)
    Messages(0).Definir_AfficherVie = CBool(CheckMessagesAfficherVie.Value)
    Messages(0).Definir_AfficherBestiaire = CBool(CheckMessagesAfficherBestiaire.Value)
    Messages(0).Definir_AfficherArgent = CBool(CheckMessagesAfficherArgent.Value)
    Messages(0).Definir_AfficherRessources = CBool(CheckMessagesAfficherRessources.Value)
    Messages(0).Definir_AfficherObjets = CBool(CheckMessagesAfficherObjets.Value)
    Messages(0).Definir_AfficherSoldats = CBool(CheckMessagesAfficherSoldats.Value)
    Messages(0).Definir_AfficherExperience = CBool(CheckMessagesAfficherExperience.Value)
    Messages(0).Definir_AfficherEmpereur = CBool(CheckMessagesAfficherEmpereur.Value)
    
    Jeu.Definir_Vitesse = SliderVitesse
    Jeu.Definir_FPS = SliderFPS
    Jeu.Definir_AffichageRapide = CBool(CheckAffichageRapide.Value)
    
    Langues.Definir_Langue = ComboLangues.ListIndex
    'If Ecran.Resolution <> ComboResolution.ListIndex Then
    Ecran.Definir_Resolution = ComboResolution.ListIndex
    '    MsgBox "Vous devez redémarrer " & App.Title & " pour que les changements graphiques soient pris en compte.", vbExclamation
    'End If
    
    Monde.Definir_TerrainOptimise = CBool(CheckOptimiserAffichage)
    AffTemps.Definir_Afficher = CBool(CheckAfficherTemps)
    Ecran.Definir_ModeFenetre = CheckModeFenetre = 0
    
    'Options réseau.
    ComReseau.Definir_Actif = CBool(CheckReseauActif)
    ComReseau.Definir_OuverturePort = CBool(CheckOuverturePort)
    ComReseau.Definir_URL_Serveur = TxtURL_Serveur
    
    'Options cartes.
    Cartes.Definir_CreerFichierInformations = CBool(CheckCreerFichierInformations)
    
    MousePointer = 0
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub TreeViewOptions_NodeClick(ByVal Node As ComctlLib.Node)
    Dim i As Integer
    Dim j As Integer
    j = TreeViewOptions.SelectedItem.Index
    For i = 0 To NombreCadres - 1
        FrameOptions(i).Visible = False
    Next i
    FrameOptions(j - 1).Visible = True
    Label6.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    Select Case j - 1
    Case CadreJeu:
        Label6.Visible = True
    Case CadreRegles:
        Label6.Visible = True
    Case CadreAffichage:
        Label6.Visible = True
    Case CadreInterface:
    Case CadreInformations:
    Case CadrePersonnages:
        'Label6.Visible = True
        Label7.Visible = True
    Case CadreTerrain
        Label6.Visible = True
        Label8.Visible = True
    Case CadreSons:
        Label6.Visible = True
    Case CadreCartes:
    Case CadreReseau:
        Label6.Visible = True
    Case CadreSauvegardes:
    Case CadreStastistiques:
    End Select
    'Label6.Visible = Not (j = CadreInformations Or j = CadrePersonnages Or j = CadreReseau Or j = CadreSauvegardes Or j = CadreStastistiques)
'    Select Case TreeViewOptions.SelectedItem.Key
'    Case "Jeu":
'    Case "Affichage":
'    Case "Interface":
'    Case "Informations":
'    Case "Personnages":
'    Case "Terrain":
'    Case "Sons":
'    Case "Sauvegardes":
'    Case "Stastistiques":
'    End Select
End Sub
