VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmParam 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10440
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   696
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   Begin VB.CommandButton CmdRetour 
      Caption         =   "&Retour"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      TabIndex        =   201
      Top             =   9840
      Width           =   1335
   End
   Begin VB.Frame FrameServeur 
      Caption         =   "Serveur"
      Height          =   2175
      Left            =   120
      TabIndex        =   75
      Top             =   360
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton CmdCreerServeur 
         Caption         =   "Créer un serveur"
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CommandButton CmdFermerServeur 
         Caption         =   "Fermer le Serveur"
         Height          =   375
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox TxtNomSession 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   80
         Top             =   600
         Width           =   3015
      End
      Begin ComctlLib.Slider SliderNombreJoueurs 
         Height          =   255
         Left            =   2520
         TabIndex        =   79
         Top             =   1080
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin VB.Label LblAdresseINet 
         Alignment       =   2  'Center
         Caption         =   "-"
         Height          =   255
         Left            =   2520
         TabIndex        =   173
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label43 
         Caption         =   "Adresse en ligne :"
         Height          =   255
         Left            =   240
         TabIndex        =   172
         Top             =   285
         Width           =   2175
      End
      Begin VB.Label Label30 
         Caption         =   "Nom de la session :"
         Height          =   255
         Left            =   240
         TabIndex        =   83
         Top             =   645
         Width           =   2175
      End
      Begin VB.Label Label29 
         Caption         =   "Nombre de joueurs maximum :"
         Height          =   255
         Left            =   240
         TabIndex        =   82
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label LblNombreJoueurs 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5160
         TabIndex        =   81
         Top             =   1080
         Width           =   375
      End
   End
   Begin VB.Frame FrameClient 
      Caption         =   "Parties en cours"
      Height          =   2175
      Left            =   120
      TabIndex        =   70
      Top             =   360
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox TxtNumeroIPServeur 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2400
         TabIndex        =   170
         Text            =   "0.0.0.0"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton CmdDeconnecter 
         Caption         =   "Se déconnecter"
         Height          =   375
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton CmdConnecter 
         Caption         =   "Se connecter"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   1020
         Width           =   1455
      End
      Begin VB.CommandButton CmdRafraichirParties 
         Caption         =   "Rafraichir"
         Height          =   375
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   360
         Width           =   1455
      End
      Begin VB.ListBox LstParties 
         Height          =   1230
         Left            =   120
         TabIndex        =   71
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label42 
         Caption         =   " IP du serveur (facultative) :"
         Height          =   255
         Left            =   120
         TabIndex        =   171
         Top             =   405
         Width           =   2175
      End
   End
   Begin VB.Timer TimerApparence 
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton CmdCommencer 
      Caption         =   "&Jouer"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13800
      TabIndex        =   1
      Top             =   9840
      Width           =   1335
   End
   Begin VB.Frame FrameReseau 
      Caption         =   "Paramètres réseaux"
      Height          =   2175
      Left            =   6000
      TabIndex        =   74
      Top             =   360
      Visible         =   0   'False
      Width           =   4335
      Begin VB.OptionButton OptionJouerEnLigne
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   205
         Top             =   1680
         Width           =   1935
      End
      Begin VB.OptionButton OptionJouerEnLigne 
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   204
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox TxtNumeroPort 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   169
         Text            =   "0000"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox LstServices 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   168
         Top             =   1680
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.TextBox TxtNomJoueur 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   76
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label41 
         Caption         =   "Numéro de port :"
         Height          =   255
         Left            =   240
         TabIndex        =   200
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label LblAdresseIP 
         Alignment       =   2  'Center
         Caption         =   "---"
         Height          =   615
         Left            =   2040
         TabIndex        =   167
         Top             =   285
         Width           =   1935
      End
      Begin VB.Label Label40 
         Caption         =   "Votre adresse IP locale :"
         Height          =   255
         Left            =   240
         TabIndex        =   166
         Top             =   285
         Width           =   1815
      End
      Begin VB.Label Label28 
         Caption         =   "Votre Pseudo :"
         Height          =   255
         Left            =   240
         TabIndex        =   78
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label27 
         Caption         =   "Sélectionnez un service provider :"
         Height          =   255
         Left            =   240
         TabIndex        =   77
         Top             =   1440
         Visible         =   0   'False
         Width           =   2895
      End
   End
   Begin VB.Frame FrameFiefs 
      Caption         =   "Fiefs"
      Height          =   7215
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Visible         =   0   'False
      Width           =   10215
      Begin VB.CheckBox CheckEquilibreFiefs 
         Caption         =   "Equilibrer les fiefs"
         Height          =   255
         Left            =   2400
         TabIndex        =   113
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox ComboEquipe 
         Height          =   315
         Index           =   0
         Left            =   9360
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   1080
         Width           =   615
      End
      Begin ComctlLib.Slider SliderVillageois 
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   10
         Top             =   1080
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   450
         _Version        =   327682
         Min             =   1
         Max             =   128
         SelStart        =   40
         Value           =   40
      End
      Begin VB.ComboBox ComboPeuple 
         Height          =   315
         Index           =   0
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin ComctlLib.Slider SliderFiefs 
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
         _Version        =   327682
         Min             =   2
         Max             =   16
         SelStart        =   2
         Value           =   2
      End
      Begin VB.Label Label26 
         Caption         =   "Equipe :"
         Height          =   255
         Left            =   9360
         TabIndex        =   67
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "Peuple :"
         Height          =   255
         Left            =   7800
         TabIndex        =   33
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label LblTotalVillageois 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "x"
         Height          =   255
         Left            =   7080
         TabIndex        =   13
         Top             =   6840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Total :"
         Height          =   255
         Left            =   6480
         TabIndex        =   12
         Top             =   6840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label LblVillageois 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "x"
         Height          =   255
         Index           =   0
         Left            =   7320
         TabIndex        =   11
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Nombre de villageois :"
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label6 
         Caption         =   "N°"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   255
      End
      Begin VB.Label LblNumeroFief 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label LblFiefs 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "x"
         Height          =   255
         Left            =   7320
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Nombre de fiefs :"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame FrameCampagnes 
      Caption         =   "Sélectionnez une campagne"
      Height          =   9855
      Left            =   120
      TabIndex        =   144
      Top             =   360
      Width           =   10215
      Begin VB.ListBox LstCampagnes 
         Height          =   2010
         Left            =   120
         TabIndex        =   149
         Top             =   240
         Width           =   3615
      End
      Begin VB.Frame FrameScenarioCampagnes 
         Caption         =   "Sélectionnez un scénario"
         Height          =   7335
         Left            =   240
         TabIndex        =   146
         Top             =   2400
         Width           =   9855
         Begin VB.ListBox LstCampagnesScenarios 
            Height          =   6690
            Left            =   360
            TabIndex        =   147
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label LblCampagnesScenariosDescription 
            Caption         =   "---"
            Height          =   6735
            Left            =   3960
            TabIndex        =   148
            Top             =   360
            Width           =   5775
         End
      End
      Begin VB.Label LblCampagneDescription 
         Caption         =   "---"
         Height          =   1935
         Left            =   3960
         TabIndex        =   145
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.Frame FrameMode2Joueurs 
      Caption         =   "Mode 2 joueurs"
      Height          =   9855
      Left            =   120
      TabIndex        =   126
      Top             =   360
      Visible         =   0   'False
      Width           =   10215
      Begin VB.Frame Frame8 
         Caption         =   "Joueur 1"
         Height          =   2175
         Index           =   1
         Left            =   240
         TabIndex        =   131
         Top             =   1320
         Width           =   4815
         Begin VB.ComboBox ComboNomJoueur 
            Height          =   315
            Index           =   0
            ItemData        =   "FrmParam.frx":0000
            Left            =   1560
            List            =   "FrmParam.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   136
            Top             =   360
            Width           =   1695
         End
         Begin VB.ListBox List1 
            Height          =   1230
            ItemData        =   "FrmParam.frx":0004
            Left            =   240
            List            =   "FrmParam.frx":001A
            TabIndex        =   132
            Top             =   720
            Width           =   4215
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Joueur 2"
         Height          =   2175
         Index           =   0
         Left            =   5160
         TabIndex        =   130
         Top             =   1320
         Width           =   4815
         Begin VB.ComboBox ComboNomJoueur 
            Height          =   315
            Index           =   1
            ItemData        =   "FrmParam.frx":0130
            Left            =   1560
            List            =   "FrmParam.frx":0132
            Style           =   2  'Dropdown List
            TabIndex        =   137
            Top             =   360
            Width           =   1695
         End
         Begin VB.ListBox List2 
            Height          =   1230
            ItemData        =   "FrmParam.frx":0134
            Left            =   240
            List            =   "FrmParam.frx":013E
            TabIndex        =   133
            Top             =   720
            Width           =   4215
         End
      End
      Begin VB.CheckBox CheckMaintenirEcransSepares 
         Caption         =   "Toujours maintenir les 2 écrans séparés."
         Height          =   255
         Left            =   240
         TabIndex        =   129
         Top             =   840
         Width           =   3615
      End
      Begin VB.CheckBox CheckJouera2 
         Caption         =   "Activer le mode 2 joueurs sur cet ordinateur."
         Height          =   255
         Left            =   240
         TabIndex        =   127
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label LblSelectionner2Personnages 
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3840
         TabIndex        =   128
         Top             =   480
         Width           =   5775
      End
   End
   Begin VB.Frame FrameMonde 
      Caption         =   "Monde"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   10215
      Begin VB.OptionButton OptionCarte 
         Caption         =   "Générer"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   192
         Top             =   240
         Width           =   1005
      End
      Begin VB.OptionButton OptionCarte 
         Caption         =   "Charger"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   191
         Top             =   240
         Width           =   1005
      End
      Begin VB.ComboBox ComboResurrections 
         Height          =   315
         Index           =   1
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   159
         Top             =   1680
         Width           =   1575
      End
      Begin VB.ComboBox ComboResurrections 
         Height          =   315
         Index           =   0
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   158
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox ComboRessourcesDepart 
         Height          =   315
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   155
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox ComboVitesseEpoque 
         Height          =   315
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   152
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox ComboDifficulte 
         Height          =   315
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   122
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ComboBox ComboEpoque 
         Height          =   315
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   600
         Width           =   1575
      End
      Begin VB.Frame FrameCarteCharger 
         Caption         =   "Carte"
         Height          =   2175
         Left            =   120
         TabIndex        =   193
         Top             =   240
         Width           =   6255
         Begin VB.CommandButton CmdChargerCarte 
            Caption         =   "Charger carte"
            Height          =   255
            Left            =   2280
            TabIndex        =   196
            Top             =   1080
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ListBox LstCartes 
            Height          =   1815
            Left            =   120
            TabIndex        =   195
            Top             =   280
            Width           =   1815
         End
         Begin VB.Image ImgCarte 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1815
            Left            =   2040
            Stretch         =   -1  'True
            Top             =   280
            Width           =   1815
         End
         Begin VB.Label LblCarteDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Caption         =   "-"
            ForeColor       =   &H80000008&
            Height          =   1780
            Left            =   4020
            TabIndex        =   194
            Top             =   300
            Width           =   2055
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H80000018&
            FillStyle       =   0  'Solid
            Height          =   1815
            Left            =   3960
            Top             =   285
            Width           =   2175
         End
      End
      Begin VB.Frame FrameCarteGeneree 
         Caption         =   "Carte"
         Height          =   2175
         Left            =   120
         TabIndex        =   175
         Top             =   240
         Width           =   6255
         Begin VB.ComboBox ComboTerrain 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   178
            Top             =   1245
            Width           =   1815
         End
         Begin VB.ComboBox ComboRessources 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   177
            Top             =   1725
            Width           =   1815
         End
         Begin VB.ComboBox ComboTaille 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   176
            Top             =   585
            Width           =   1815
         End
         Begin ComctlLib.Slider SliderNombreArbres 
            Height          =   255
            Left            =   2880
            TabIndex        =   179
            Top             =   1770
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   450
            _Version        =   327682
            Max             =   300
            SelStart        =   50
            Value           =   50
         End
         Begin ComctlLib.Slider SliderHauteurCarte 
            Height          =   255
            Left            =   3720
            TabIndex        =   180
            Top             =   840
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   327682
            Min             =   48
            Max             =   512
            SelStart        =   128
            Value           =   128
         End
         Begin ComctlLib.Slider SliderLargeurCarte 
            Height          =   255
            Left            =   3720
            TabIndex        =   181
            Top             =   450
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   327682
            Min             =   64
            Max             =   512
            SelStart        =   128
            Value           =   128
         End
         Begin VB.Label Label8 
            Caption         =   "Terrain :"
            Height          =   255
            Left            =   120
            TabIndex        =   190
            Top             =   1290
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "Ressources :"
            Height          =   255
            Left            =   120
            TabIndex        =   189
            Top             =   1770
            Width           =   1455
         End
         Begin VB.Label Label16 
            Caption         =   "Taille :"
            Height          =   255
            Left            =   120
            TabIndex        =   188
            Top             =   630
            Width           =   1335
         End
         Begin VB.Label LblNombreArbres 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "x"
            Height          =   255
            Left            =   5640
            TabIndex        =   187
            Top             =   1770
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Ressources :"
            Height          =   255
            Left            =   3000
            TabIndex        =   186
            Top             =   1770
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label LblHauteurCarte 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "x"
            Height          =   255
            Left            =   5640
            TabIndex        =   185
            Top             =   840
            Width           =   495
         End
         Begin VB.Label LblLargeurCarte 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "x"
            Height          =   255
            Left            =   5640
            TabIndex        =   184
            Top             =   450
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Hauteur :"
            Height          =   255
            Left            =   3000
            TabIndex        =   183
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Largeur :"
            Height          =   255
            Left            =   3000
            TabIndex        =   182
            Top             =   450
            Width           =   735
         End
      End
      Begin VB.Label Label39 
         Caption         =   "Resurrections :"
         Height          =   255
         Left            =   6480
         TabIndex        =   160
         Top             =   1725
         Width           =   2055
      End
      Begin VB.Label Label38 
         Caption         =   "Resurrections :"
         Height          =   255
         Left            =   6480
         TabIndex        =   157
         Top             =   1365
         Width           =   2055
      End
      Begin VB.Label Label36 
         Caption         =   "Ressources de départ :"
         Height          =   255
         Left            =   6480
         TabIndex        =   156
         Top             =   285
         Width           =   2055
      End
      Begin VB.Label Label37 
         Caption         =   "Changement d'époque :"
         Height          =   255
         Left            =   6480
         TabIndex        =   151
         Top             =   1005
         Width           =   2055
      End
      Begin VB.Label Label33 
         Caption         =   "Difficulté :"
         Height          =   255
         Left            =   6480
         TabIndex        =   121
         Top             =   2085
         Width           =   2055
      End
      Begin VB.Label Label24 
         Caption         =   "Epoque de départ :"
         Height          =   255
         Left            =   6480
         TabIndex        =   65
         Top             =   645
         Width           =   2055
      End
   End
   Begin VB.Frame FrameChargerMonde 
      Caption         =   "Sélectionnez une partie"
      Height          =   9855
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   10215
      Begin VB.ListBox LstMonde 
         Height          =   2010
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   3495
      End
      Begin VB.FileListBox FileMonde 
         Height          =   2040
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label LblDescriptionMonde 
         BackStyle       =   0  'Transparent
         Height          =   2055
         Left            =   3840
         TabIndex        =   26
         Top             =   360
         Width           =   6015
      End
      Begin VB.Image ImgApercu 
         Height          =   7335
         Left            =   240
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   9855
      End
   End
   Begin ComctlLib.TabStrip TabStripMonde2 
      Height          =   10335
      Left            =   0
      TabIndex        =   198
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   18230
      TabWidthStyle   =   2
      TabFixedWidth   =   3043
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   6
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Campagne"
            Key             =   "Campagnes"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Monde personnalisé"
            Key             =   "NouveauMonde"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Champ de bataille"
            Key             =   "NouvelleBataille"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Charger scénario"
            Key             =   "ChargerScenario"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Charger partie"
            Key             =   "ChargerMonde"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Mode 2 joueurs"
            Key             =   "Mode2Joueurs"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrameChargerPersonnage 
      Caption         =   "Sélectionnez un ou plusieurs personnages"
      Height          =   9255
      Left            =   10680
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton CmdApparence 
         Caption         =   "Personnaliser"
         Height          =   255
         Left            =   480
         TabIndex        =   140
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton OptionApparence 
         Caption         =   "Fixe"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   154
         Top             =   4440
         Width           =   1455
      End
      Begin VB.OptionButton OptionApparence 
         Caption         =   "Personnalisée"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   153
         Top             =   4200
         Width           =   1455
      End
      Begin VB.OptionButton OptionIAScript 
         Caption         =   "Scriptée"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   143
         Top             =   4440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton OptionIAScript 
         Caption         =   "Standard"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   142
         Top             =   4200
         Width           =   1455
      End
      Begin VB.ListBox LstPointsTalents 
         Height          =   450
         Left            =   1800
         TabIndex        =   120
         Top             =   4440
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ListBox LstTalents 
         Height          =   4110
         ItemData        =   "FrmParam.frx":0189
         Left            =   1800
         List            =   "FrmParam.frx":018B
         Style           =   1  'Checkbox
         TabIndex        =   119
         Top             =   4920
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton CmdSupprimerPerso 
         Caption         =   "&Supprimer personnage"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   3000
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ListBox LstPerso 
         Height          =   3180
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   64
         Top             =   360
         Width           =   1575
      End
      Begin VB.Frame FramePersonnage 
         Height          =   3495
         Left            =   1920
         TabIndex        =   34
         Top             =   240
         Width           =   2175
         Begin VB.ComboBox ComboEquipePerso 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   123
            Top             =   920
            Width           =   495
         End
         Begin VB.Frame Frame1 
            Height          =   1915
            Left            =   1680
            TabIndex        =   62
            Top             =   1440
            Width           =   360
            Begin VB.Image ImgApparence 
               Height          =   1920
               Left            =   0
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.Line Line10 
            X1              =   1560
            X2              =   2040
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label LblIA 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "IA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1840
            TabIndex        =   203
            Top             =   240
            Width           =   195
         End
         Begin VB.Image ImgLocker 
            Height          =   255
            Index           =   1
            Left            =   1600
            Top             =   220
            Width           =   255
         End
         Begin VB.Image ImgLocker 
            Height          =   255
            Index           =   0
            Left            =   1600
            Top             =   220
            Width           =   255
         End
         Begin VB.Line Line3 
            X1              =   120
            X2              =   2060
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line8 
            X1              =   2040
            X2              =   2040
            Y1              =   160
            Y2              =   3360
         End
         Begin VB.Line Line5 
            X1              =   120
            X2              =   2040
            Y1              =   160
            Y2              =   160
         End
         Begin VB.Line Line9 
            X1              =   1560
            X2              =   1560
            Y1              =   3360
            Y2              =   1200
         End
         Begin VB.Line Line7 
            X1              =   1560
            X2              =   1560
            Y1              =   960
            Y2              =   160
         End
         Begin VB.Line Line6 
            X1              =   120
            X2              =   1570
            Y1              =   675
            Y2              =   675
         End
         Begin VB.Line Line4 
            X1              =   120
            X2              =   120
            Y1              =   160
            Y2              =   3360
         End
         Begin VB.Line Line2 
            X1              =   1560
            X2              =   120
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Line Line1 
            X1              =   1560
            X2              =   120
            Y1              =   1725
            Y2              =   1725
         End
         Begin VB.Label LblNiveau 
            Alignment       =   1  'Right Justify
            Caption         =   "----"
            Height          =   255
            Left            =   960
            TabIndex        =   37
            Top             =   720
            Width           =   495
         End
         Begin VB.Label LblRace 
            Alignment       =   1  'Right Justify
            Caption         =   "----"
            Height          =   255
            Left            =   720
            TabIndex        =   95
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label LblSexe 
            Alignment       =   1  'Right Justify
            Caption         =   "----"
            Height          =   255
            Left            =   840
            TabIndex        =   39
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label LblDefense 
            Alignment       =   1  'Right Justify
            Caption         =   "----"
            Height          =   255
            Left            =   960
            TabIndex        =   46
            Top             =   3120
            Width           =   495
         End
         Begin VB.Label LblAttaque 
            Alignment       =   1  'Right Justify
            Caption         =   "----"
            Height          =   255
            Left            =   960
            TabIndex        =   47
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label LblMoral 
            Alignment       =   1  'Right Justify
            Caption         =   "----"
            Height          =   255
            Left            =   960
            TabIndex        =   48
            Top             =   2520
            Width           =   495
         End
         Begin VB.Label LblMagie 
            Alignment       =   1  'Right Justify
            Caption         =   "----"
            Height          =   255
            Left            =   960
            TabIndex        =   49
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label LblEnergie 
            Alignment       =   1  'Right Justify
            Caption         =   "----"
            Height          =   255
            Left            =   960
            TabIndex        =   50
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label LblVie 
            Alignment       =   1  'Right Justify
            Caption         =   "----"
            Height          =   255
            Left            =   960
            TabIndex        =   51
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label LblMetier 
            Alignment       =   2  'Center
            Caption         =   "--------------------------------------"
            Height          =   255
            Left            =   120
            TabIndex        =   125
            Top             =   435
            Width           =   1455
         End
         Begin VB.Label Label34 
            Caption         =   "Equipe :"
            Height          =   255
            Left            =   240
            TabIndex        =   124
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Energie :"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label31 
            Caption         =   "Race :"
            Height          =   255
            Left            =   240
            TabIndex        =   94
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "Attaque :"
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "Défense :"
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Vie :"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label12 
            Caption         =   "Moral :"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "Magie:"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label LblNivPlus 
            Alignment       =   2  'Center
            Caption         =   "----"
            Height          =   255
            Left            =   1620
            TabIndex        =   41
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            Caption         =   "Niv+"
            Height          =   255
            Left            =   1620
            TabIndex        =   40
            Top             =   520
            Width           =   375
         End
         Begin VB.Label Label17 
            Caption         =   "Sexe :"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label19 
            Caption         =   "Niveau :"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label LblNom 
            Alignment       =   2  'Center
            Caption         =   "--------------------------------------"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   195
            Width           =   1455
         End
      End
      Begin VB.FileListBox FilePerso 
         Height          =   3210
         Left            =   120
         MultiSelect     =   2  'Extended
         Pattern         =   "*"
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.FileListBox FileApparence 
         Height          =   4380
         Left            =   120
         TabIndex        =   23
         Top             =   4680
         Width           =   1695
      End
      Begin VB.FileListBox FileIA 
         Height          =   3990
         Left            =   120
         TabIndex        =   69
         Top             =   4920
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ListBox LstCompetences 
         Height          =   4545
         Left            =   1800
         TabIndex        =   20
         Top             =   4440
         Width           =   2535
      End
      Begin VB.ListBox LstBestiaire 
         Height          =   4545
         Left            =   1800
         TabIndex        =   21
         Top             =   4440
         Visible         =   0   'False
         Width           =   2535
      End
      Begin ComctlLib.TabStrip TabStripCompBest 
         Height          =   5175
         Left            =   1800
         TabIndex        =   24
         Top             =   3840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   9128
         MultiRow        =   -1  'True
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   3
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Compétences"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Bestiaire"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Trésors"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.FileListBox FileIAScript 
         Height          =   3990
         Left            =   120
         TabIndex        =   141
         Top             =   4920
         Width           =   1695
      End
      Begin VB.OptionButton OptionIAScript 
         Caption         =   "Aucune"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   202
         Top             =   4200
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
      Begin ComctlLib.TabStrip TabStripFichiers 
         Height          =   5175
         Left            =   120
         TabIndex        =   114
         Top             =   3840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   9128
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   2
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Apparence"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "IA"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label LblNombrePersos 
         Alignment       =   2  'Center
         Caption         =   "--------------------------------"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   3600
         Width           =   1575
      End
   End
   Begin VB.Frame FrameNouveauPersonnage 
      Caption         =   "Créez un nouveau personnage"
      Height          =   9255
      Left            =   10680
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Frame FrameGenerationPersonnage 
         Caption         =   "Génération automatiquement de personnages :"
         Height          =   2175
         Left            =   240
         TabIndex        =   107
         Top             =   6960
         Width           =   3855
         Begin VB.CommandButton CmdGenererPersos 
            Caption         =   "Générer &10 personnages"
            Height          =   375
            Index           =   1
            Left            =   720
            TabIndex        =   110
            Top             =   960
            Width           =   2295
         End
         Begin VB.CommandButton CmdGenererPersos 
            Caption         =   "&Générer 1 personnage"
            Height          =   375
            Index           =   0
            Left            =   720
            TabIndex        =   109
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton CmdGenererPersos 
            Caption         =   "Générer &50 personnages"
            Height          =   375
            Index           =   2
            Left            =   720
            TabIndex        =   108
            Top             =   1560
            Width           =   2295
         End
      End
      Begin VB.Frame FrameValidationPersonnage 
         Caption         =   "Validation :"
         Height          =   855
         Left            =   240
         TabIndex        =   105
         Top             =   5520
         Width           =   3855
         Begin VB.CommandButton CmdCreerPerso 
            Enabled         =   0   'False
            Height          =   495
            Left            =   600
            Picture         =   "FrmParam.frx":018D
            Style           =   1  'Graphical
            TabIndex        =   106
            Top             =   240
            Width           =   2580
         End
      End
      Begin VB.Frame FrameRacePersonnage 
         Caption         =   "Race :"
         Height          =   3015
         Left            =   240
         TabIndex        =   102
         Top             =   1440
         Width           =   3855
         Begin VB.CommandButton CmdInfoRace 
            Caption         =   "?"
            Height          =   375
            Left            =   3360
            TabIndex        =   111
            Top             =   1920
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ListBox LstRaces 
            Columns         =   4
            Height          =   1035
            Left            =   240
            TabIndex        =   104
            Top             =   240
            Width           =   3375
         End
         Begin VB.TextBox TxtRaceDecription 
            Height          =   1455
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   103
            Text            =   "FrmParam.frx":05CF
            Top             =   1440
            Width           =   3375
         End
      End
      Begin VB.Frame FrameSexePersonnage 
         Caption         =   "Sexe :"
         Height          =   855
         Left            =   240
         TabIndex        =   99
         Top             =   4560
         Width           =   3855
         Begin VB.OptionButton OptionSexe 
            Caption         =   "&Féminin"
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   101
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton OptionSexe 
            Caption         =   "&Masculin"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   100
            Top             =   480
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.Label LblRestictionSexe 
            Caption         =   "Ce choix est imposé par votre race."
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame FrameNomPersonnage 
         Caption         =   "Nom :"
         Height          =   1095
         Left            =   240
         TabIndex        =   96
         Top             =   240
         Width           =   3855
         Begin VB.TextBox TxtNomPerso 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   240
            TabIndex        =   97
            Top             =   600
            Width           =   3255
         End
         Begin VB.Label LblErreurNom 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   240
            Width           =   3735
         End
      End
   End
   Begin ComctlLib.TabStrip TabStripPersonnage 
      Height          =   9735
      Left            =   10560
      TabIndex        =   17
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   17171
      TabWidthStyle   =   2
      TabFixedWidth   =   3916
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Nouveau personnage"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Charger personnages"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrameInfoReseau 
      Caption         =   "Informations"
      Height          =   7575
      Left            =   120
      TabIndex        =   86
      Top             =   2640
      Visible         =   0   'False
      Width           =   10215
      Begin VB.Frame Frame2 
         Caption         =   "Détails de la partie"
         Height          =   2415
         Left            =   2040
         TabIndex        =   161
         Top             =   720
         Width           =   8055
         Begin VB.Label LblDetailPartie 
            Caption         =   "-"
            Height          =   1935
            Left            =   120
            TabIndex        =   164
            Top             =   360
            Width           =   3735
         End
      End
      Begin VB.ListBox LstMessagesInfos 
         Height          =   3765
         Left            =   6120
         TabIndex        =   174
         Top             =   3600
         Width           =   3975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Détails du personnage"
         Height          =   2415
         Left            =   2040
         TabIndex        =   162
         Top             =   720
         Width           =   3975
         Begin VB.Label LblDetailJoueur 
            Caption         =   "-"
            Height          =   1935
            Left            =   120
            TabIndex        =   163
            Top             =   360
            Width           =   3735
         End
      End
      Begin VB.CommandButton CmdEnvoyerMessage 
         Caption         =   "Envoyer"
         Height          =   255
         Left            =   9120
         TabIndex        =   92
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox TxtMessage 
         Height          =   285
         Left            =   3000
         TabIndex        =   91
         Top             =   3240
         Width           =   6015
      End
      Begin VB.ListBox LstMessages 
         Height          =   3765
         Left            =   2040
         TabIndex        =   90
         Top             =   3600
         Width           =   3975
      End
      Begin VB.CommandButton CmdRafraichirJoueurs 
         Caption         =   "Rafraichir"
         Height          =   375
         Left            =   120
         TabIndex        =   88
         Top             =   960
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ListBox LstJoueurs 
         Height          =   6495
         Left            =   120
         TabIndex        =   87
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label45 
         Caption         =   "Label45"
         Height          =   255
         Left            =   240
         TabIndex        =   199
         Top             =   440
         Width           =   9855
      End
      Begin VB.Label Label44 
         Caption         =   $"FrmParam.frx":05D5
         Height          =   255
         Left            =   240
         TabIndex        =   197
         Top             =   240
         Width           =   9855
      End
      Begin VB.Label Label32 
         Caption         =   "Messages :"
         Height          =   255
         Left            =   2040
         TabIndex        =   89
         Top             =   3240
         Width           =   1815
      End
   End
   Begin VB.Frame FrameArmees 
      Caption         =   "Armées"
      Height          =   7215
      Left            =   120
      TabIndex        =   28
      Top             =   3000
      Visible         =   0   'False
      Width           =   10215
      Begin VB.Frame FrameArmee 
         Caption         =   "Armée adverse"
         Height          =   6615
         Index           =   1
         Left            =   5280
         TabIndex        =   56
         Top             =   360
         Width           =   4695
         Begin VB.CheckBox CheckEquipementArmee 
            Caption         =   "Equipement :"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   139
            Top             =   1320
            Width           =   2175
         End
         Begin VB.ListBox LstObjets 
            Height          =   4785
            Index           =   1
            ItemData        =   "FrmParam.frx":06E9
            Left            =   240
            List            =   "FrmParam.frx":06EB
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   116
            Top             =   1560
            Width           =   4215
         End
         Begin VB.ComboBox ComboArmee 
            Height          =   315
            Index           =   1
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   600
            Width           =   1455
         End
         Begin ComctlLib.Slider SliderNombreSoldats 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   58
            Top             =   960
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   450
            _Version        =   327682
            Min             =   100
            Max             =   300
            SelStart        =   100
            Value           =   100
         End
         Begin VB.Label Label35 
            Caption         =   "Equipe : 2"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   135
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label25 
            Height          =   255
            Left            =   240
            TabIndex        =   118
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label LblNombreSoldats 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "x"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   61
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Nombre de soldats :"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   60
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label23 
            Caption         =   "Peuple :"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   59
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame FrameArmee 
         Caption         =   "Armée alliée"
         Height          =   6615
         Index           =   0
         Left            =   360
         TabIndex        =   29
         Top             =   360
         Width           =   4695
         Begin VB.CheckBox CheckEquipementArmee 
            Caption         =   "Equipement :"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   138
            Top             =   1320
            Width           =   2175
         End
         Begin VB.ListBox LstObjets 
            Height          =   4785
            Index           =   0
            ItemData        =   "FrmParam.frx":06ED
            Left            =   240
            List            =   "FrmParam.frx":06EF
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   115
            Top             =   1560
            Width           =   4215
         End
         Begin VB.ComboBox ComboArmee 
            Height          =   315
            Index           =   0
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   600
            Width           =   1455
         End
         Begin ComctlLib.Slider SliderNombreSoldats 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   960
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   450
            _Version        =   327682
            Min             =   100
            Max             =   300
            SelStart        =   100
            Value           =   100
         End
         Begin VB.Label Label35 
            Caption         =   "Equipe : 1"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   134
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label9 
            Height          =   255
            Left            =   240
            TabIndex        =   117
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label23 
            Caption         =   "Peuple :"
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   55
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label20 
            Caption         =   "Nombre de soldats :"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   32
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label LblNombreSoldats 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "x"
            Height          =   255
            Index           =   0
            Left            =   1920
            TabIndex        =   31
            Top             =   600
            Width           =   615
         End
      End
   End
   Begin ComctlLib.TabStrip TabStripMonde 
      Height          =   10335
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   18230
      TabFixedWidth   =   2249
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   8
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Campagne"
            Key             =   "Campagnes"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Monde personnalisé"
            Key             =   "NouveauMonde"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Champ de bataille"
            Key             =   "NouvelleBataille"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Charger scénario"
            Key             =   "ChargerScenario"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Charger partie"
            Key             =   "ChargerMonde"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Mode 2 joueurs"
            Key             =   "Mode2Joueurs"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Créer serveur"
            Key             =   "CreerServeur"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Rejoindre Partie"
            Key             =   "RejoindrePartie"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label LblReseau 
      Alignment       =   2  'Center
      Caption         =   "---"
      Height          =   495
      Left            =   12960
      TabIndex        =   165
      Top             =   9840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LblMode2J 
      Alignment       =   2  'Center
      Caption         =   "---"
      Height          =   495
      Left            =   12000
      TabIndex        =   150
      Top             =   9840
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmParam"
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

Const FichierINI = "Partie"
Const FichierFenetreINI = "Fenetre_Partie"
Const SectionFenetreINI = "Fenetre_Partie"
Const HauteurLignesFiefs = 380

Public ChargerPerso As Boolean
Public EnregistrerPerso As Boolean

Private ErreurNomPersonnage() As String

'Messages affichés occasionnellements.
Private EtiquetteMode2JImpossible As String
Private EtiquetteMode2JActive As String
Private EtiquetteMode2JDesActive As String
Private TitreMode2JImpossible As String
Private MessageMode2JImpossible As String

Private EtiquetteReseauClient As String
Private EtiquetteReseauServeur As String

'Paramètres des combos box de la carte.
Dim TailleLargeur() As Long
Dim TailleHauteur() As Long
Dim TailleRessources() As Long
Dim TailleFiefs() As Long
Dim RessourcesMultiplicateur() As Double
Dim DifficulteBonusXP() As Long
Dim DifficulteCoefVieChateaux() As Single
Dim ResurrectionsValeur() As Integer
Dim ResurrectionsCoefs() As Single

Dim CompteTimer As Integer

Dim DernierMonde As Integer
Dim DernierScenario As Integer
Dim DernierePartie As Integer
Dim FichiersIAScripts As Boolean 'Si faux, aucun fichiers d'IA script n'est disponible.
Public NombrePersonnages As Integer 'Indique le nombre de personnages sélectionnés.

Private Sub Check1_Click()

End Sub

Private Sub CheckEquilibreFiefs_Click()
    SliderVillageois_Scroll 0
End Sub

Private Sub CheckEquipementArmee_Click(Index As Integer)
    LstObjets(Index).Visible = CheckEquipementArmee(Index).Value = 1
End Sub

Private Sub CheckJouera2_Click()
    If CheckJouera2.Value = 1 Then
        If Not Interface(0).ModeDuoPossible(Ecran) Then
            MsgBox MessageMode2JImpossible, _
                   vbExclamation, TitreMode2JImpossible
            CheckJouera2.Value = 0
            LblMode2J.Visible = False
        Else
            LblMode2J.Visible = True
        End If
        ComboNomJoueur_Click 0
    Else
        LblMode2J.Visible = False
    End If
    Jeu.Definir_Mode2Joueurs = CBool(CheckJouera2.Value)
End Sub

Private Sub CheckMaintenirEcransSepares_Click()
    Jeu.Definir_Mode2JoueursEcranSepares = CBool(CheckMaintenirEcransSepares.Value)
End Sub

Private Sub CmdApparence_Click()
    'OptionApparence_Click 0
    FrmApparence.Show
End Sub

Public Sub CmdChargerCarte_Click()
    FrmMoteur2D.Afficher_Image_Chargement dd, 30, Parametres.Etiquette_Chargement(11)
    Cartes.Charger_Monde Monde
    FrmMoteur2D.Afficher_Image_Chargement dd, 30, Parametres.Etiquette_Chargement(12)
    Cartes.Charger_Chateaux Monde, Chateaux()
    FrmMoteur2D.Afficher_Image_Chargement dd, 30, Parametres.Etiquette_Chargement(13)
    Cartes.Charger_Ressources Monde, Ressources()
End Sub

Private Sub CmdConnecter_Click()
    MousePointer = 11
    If LstParties.ListIndex >= 0 Then
        ComReseau.Session_Rejoindre LstParties.ListIndex
        'CmdRafraichirJoueurs_Click
        LstParties.Clear
        Actualiser_Boutons_Reseau
    End If
    CmdConnecter.BackColor = &H8000000F
    LstMessages.Clear
    LstMessagesInfos.Clear
    MousePointer = 0
End Sub

Private Sub CmdCreerServeur_Click()
    MousePointer = 11
    ComReseau.Session_Creer
    Actualiser_Boutons_Reseau
    CmdCreerServeur.BackColor = &H8000000F
    CmdRafraichirParties.BackColor = &H8000000F
    LstMessages.Clear
    LstMessagesInfos.Clear
    MousePointer = 0
End Sub

Private Sub CmdDeconnecter_Click()
    MousePointer = 11
    ComReseau.Session_Fermer
    Actualiser_Boutons_Reseau
    LstJoueurs.Clear
    MousePointer = 0
End Sub

Private Sub CmdEnvoyerMessage_Click()
    If TxtMessage.Text <> "" Then
        'ComReseau.Message_Envoyer ComReseau.NumeroJoueur, "<" & ComReseau.NomJoueur & "> " & TxtMessage.Text
        ComMessages.MessageChat "<" & ComReseau.NomJoueur & "> " & TxtMessage.Text
        LstMessages.AddItem "<" & ComReseau.NomJoueur & "> " & TxtMessage.Text
        TxtMessage.Text = ""
    End If
End Sub

Private Sub CmdFermerServeur_Click()
    MousePointer = 11
    ComReseau.Session_Fermer
    LstJoueurs.Clear
    Actualiser_Boutons_Reseau
    MousePointer = 0
End Sub

Private Sub CmdGenererPersos_Click(Index As Integer)
    On Error GoTo Erreur
    Dim i As Long
    Dim j As Long
    Dim NombrePersos As Long
    MousePointer = 11
    Select Case Index
    Case 0:
        NombrePersos = 1
    Case 1:
        NombrePersos = 10
    Case 2:
        NombrePersos = 50
    End Select
    For i = 0 To NombrePersos - 1
        LstRaces.ListIndex = Int(Rnd * LstRaces.ListCount)
        If OptionSexe(0).Enabled Then
            If Rnd < 0.5 Then
                OptionSexe(0).Value = True
            Else
                OptionSexe(1).Value = True
            End If
        End If
        TxtNomPerso = Noms.Tirer_Nom_Aleatoire(OptionSexe(1).Value)
        If CmdCreerPerso.Enabled Then
            CmdCreerPerso_Click
            'Choisit une apparence au hasard.
            FileApparence.ListIndex = Rnd * (FileApparence.ListCount - 1)
            'Génère une IA au hasard.
            FileIA.ListIndex = Rnd * (FileIA.ListCount - 1)
        Else
            i = i - 1
        End If
        j = j + 1
        If j > 1000 Then i = NombrePersos
    Next i
    MousePointer = 0
    Exit Sub
Erreur:
    MousePointer = 0
    Exit Sub
End Sub

Private Sub CmdInfoRace_Click()
    If FrmParam.Visible Then
        MDIFrmMain.MenuAideRaces_Click
        FrmInfoRaces.LstRaces.ListIndex = FrmParam.LstRaces.ListIndex
    End If
End Sub

Public Sub CmdRafraichirJoueurs_Click()
    Dim i As Integer
    Dim Temp As Integer
    Temp = LstJoueurs.ListIndex
    LstJoueurs.Clear
    With ComReseau
    .Rafraichir_Joueurs
    For i = 1 To .lngNumJoueurs
        LstJoueurs.AddItem .Joueur_Nom(i) & IIf(.Joueur_Serveur(i), " (Serveur)", "")
    Next i
    If Temp >= 0 Then LstJoueurs.ListIndex = Temp
    End With
End Sub

Private Sub CmdRafraichirParties_Click()
    Dim i As Integer
    MousePointer = 11
    With ComReseau
    .Rafraichir_Parties
    LstParties.Clear
    If .Definir_JouerEnLigne Then
        Dim Temp1() As String
        Dim Temp2() As String
        Dim Temp3() As String
        Dim Temp4() As String
        Temp1 = .SessionsEnLigne.lireDataXML2("GAME", "NAME")
        Temp2 = .SessionsEnLigne.lireDataXML2("GAME", "PLAYER")
        Temp3 = .SessionsEnLigne.lireDataXML2("GAME", "IP")
        Temp4 = .SessionsEnLigne.lireDataXML2("GAME", "PORT")
        For i = 0 To UBound(Temp1())
            LstParties.AddItem Temp1(i) & " (" & Temp2(i) & ") - IP:" & Temp3(i) & " - PORT:" & Temp4(i)
        Next i
    Else
        For i = 1 To .lngNumSession
            LstParties.AddItem .Partie_Nom(i) & " (" & .Partie_NombreJoueurs(i) & "/" & .Partie_NombreJoueursMax(i) & ")"
        Next i
    End If
    End With
    If LstParties.ListCount > 0 Then
        LstParties.ListIndex = 0
        CmdRafraichirParties.BackColor = &H8000000F
    '    CmdConnecter.Enabled = True
    'Else
    '    CmdConnecter.Enabled = False
    End If
    Actualiser_Boutons_Reseau
    MousePointer = 0
End Sub

'Private Sub ComboEquipePerso_Click()
'    Dim i As Integer
'    If LblNombrePersos.Visible Then
'        'Cas où l'on a sélectionné plusieurs persos.
'        If MsgBox("Voulez-vous attribuer ce numéro d'équipe" & Chr(10) & _
'                  "à tous les personnages sélectionnés ?", vbYesNo, _
'                  "Changement d'équipe") = _
'                  vbYes Then
'            For i = 0 To LstPerso.ListCount - 1
'                If LstPerso.Selected(i) Then
'                    Changer_Equipe_Perso LstPerso.List(i), ComboEquipePerso.ListIndex
'                End If
'            Next i
'            Persos(Noperso).NumeroFief = ComboEquipePerso.ListIndex
'        Else
'            Persos(Noperso).NumeroFief = ComboEquipePerso.ListIndex
'            Changer_Equipe_Perso LblNom, ComboEquipePerso.ListIndex
'        End If
'    Else
'        Persos(Noperso).NumeroFief = ComboEquipePerso.ListIndex
'        Changer_Equipe_Perso LblNom, ComboEquipePerso.ListIndex
'    End If
'End Sub

Private Sub CmdRetour_Click()
    Me.Hide
    FrmChargement.Show
    FrmChargement.Afficher = True
    FrmChargement.Rafraichir_Boutons
End Sub

Private Sub ComboEquipePerso_Validate(Cancel As Boolean)
    Persos(NoPerso).NumeroFief = ComboEquipePerso.ListIndex
    Changer_Equipe_Perso LblNom, ComboEquipePerso.ListIndex
End Sub

Private Sub ComboNomJoueur_Click(Index As Integer)
    Dim i As Integer
    'If ComboNomJoueur(0).List(ComboNomJoueur(0).ListIndex) = ComboNomJoueur(1).List(ComboNomJoueur(1).ListIndex) Then
    If ComboNomJoueur(0).Text = ComboNomJoueur(1).Text And CheckJouera2.Value = 1 Then
        LblSelectionner2Personnages.Caption = EtiquetteMode2JImpossible
        CheckJouera2.Enabled = False
        LblMode2J.Caption = EtiquetteMode2JDesActive
    Else
        LblSelectionner2Personnages.Caption = ""
        CheckJouera2.Enabled = True
        LblMode2J.Caption = EtiquetteMode2JActive
    End If
End Sub

Private Sub ComboResurrections_Click(Index As Integer)
    'Label39.Visible = Definir_Resurrections_Nombre <> 0
    'ComboResurrections(1).Visible = Definir_Resurrections_Nombre <> 0
End Sub

Private Sub Command1_Click()
    Cartes.Charger_Monde Monde
End Sub

Private Sub FileIA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu MDIFrmMain.MenuOptions
    End If
End Sub

Private Sub FileIAScript_Click()
    If FileIAScript.ListCount > 0 And _
       FileIAScript.Filename <> "" Then
        Persos(NoPerso).FichierIAScript = FileIAScript.Filename
    End If
    If ChargerPerso Then
        Enregistrer_Sauvegarde_Perso Persos(NoPerso)
    End If
End Sub

Private Sub LstCampagnes_Click()
    Dim Temp As String
    On Error GoTo Erreur
    
    If LstCampagnes.ListIndex > -1 Then
        'Charge les scénarios.
        LstCampagnesScenarios.Clear
        Temp = Dir(App.Path & CheminCampagnes & Langues.Dossier & LstCampagnes.List(LstCampagnes.ListIndex) & "\", vbDirectory)
        While Temp <> ""
            If Right(Temp, Len(ExtensionFichiersScenario)) = ExtensionFichiersScenario Then
                LstCampagnesScenarios.AddItem Left(Temp, Len(Temp) - Len(ExtensionFichiersScenario))
            End If
            Temp = Dir()
        Wend
        
        If LstCampagnesScenarios.ListCount > 0 Then
            LstCampagnesScenarios.ListIndex = 0
        End If
        'Charge la description.
        Open App.Path & CheminCampagnes & Langues.Dossier & LstCampagnes.List(LstCampagnes.ListIndex) & "\" & LstCampagnes.List(LstCampagnes.ListIndex) & ExtensionFichiersTexte For Input As #1
        Line Input #1, Temp
        LblCampagneDescription.Caption = Temp
        If LblCampagneDescription.Caption <> "" Then
            While Not EOF(1)
                Line Input #1, Temp
                LblCampagneDescription.Caption = LblCampagneDescription.Caption & Chr(13) & Temp
            Wend
        End If
        Close #1
    Else
        LblCampagneDescription.Caption = ""
    End If
    Exit Sub
Erreur:
    LblCampagneDescription.Caption = ""
End Sub

Private Sub LstCampagnesScenarios_Click()
    On Error GoTo Erreur
    Dim Temp As String
    'Charge la description.
    Open App.Path & CheminCampagnes & Langues.Dossier & LstCampagnes.List(LstCampagnes.ListIndex) & "\" & LstCampagnesScenarios.List(LstCampagnesScenarios.ListIndex) & ExtensionFichiersTexte For Input As #1
    Line Input #1, Temp
    LblCampagnesScenariosDescription.Caption = Temp
    If LblCampagnesScenariosDescription <> "" Then
        While Not EOF(1)
            Line Input #1, Temp
            LblCampagnesScenariosDescription.Caption = LblCampagnesScenariosDescription.Caption & Chr(13) & Temp
        Wend
    End If
    Close #1
    Exit Sub
Erreur:
    LblCampagnesScenariosDescription.Caption = ""
End Sub

Private Sub LstCartes_Click()
    MousePointer = 11
    If LstCartes.ListIndex = -1 Then LstCartes.ListIndex = 0
    Cartes.Charger_Apercu LstCartes.List(LstCartes.ListIndex), ImgCarte
    Cartes.Charger_Description LstCartes.List(LstCartes.ListIndex), LblCarteDescription
    If LblCarteDescription.Caption <> "" Then
        LblCarteDescription.Caption = Right(LblCarteDescription.Caption, Len(LblCarteDescription.Caption) - 2)
    End If
    SliderFiefs.Value = Cartes.Definir_Nombre_Chateaux
    SliderFiefs_Scroll
    SliderFiefs.Enabled = False
    LblFiefs.BackColor = &H80000011
    MousePointer = 0
End Sub

Private Sub LstCompetences_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu MDIFrmMain.MenuPersonnageCompetences
    End If
End Sub

Private Sub LstParties_DblClick()
    If CmdConnecter.Enabled Then
        CmdConnecter_Click
    End If
End Sub

Private Sub LstPerso_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu MDIFrmMain.MenuPersonnage
    End If
End Sub

Private Sub LstMonde_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu MDIFrmMain.MenuMonde
    End If
End Sub

Private Sub Form_Load()
    'On Error GoTo Erreur
    Dim FicIni2 As ClsFicINI
    Dim i As Integer, j As Long
    Dim EtiquetteAleatoire As String
    
    Set FicIni2 = New ClsFicINI
    CompteTimer = 0
    'FichiersIAScripts = True
    
    FicIni.fichier = FicIni.chemin & Langues.Dossier & FichierFenetreINI
    FicIni.Section = SectionFenetreINI
    
    TabStripMonde.Tabs(1).Caption = FicIni.Parametre("BandeCampagne")
    TabStripMonde.Tabs(2).Caption = FicIni.Parametre("BandePersonnaliser")
    TabStripMonde.Tabs(3).Caption = FicIni.Parametre("BandeBataille")
    TabStripMonde.Tabs(4).Caption = FicIni.Parametre("BandeScenario")
    TabStripMonde.Tabs(5).Caption = FicIni.Parametre("BandeSauvegarde")
    TabStripMonde.Tabs(6).Caption = FicIni.Parametre("BandeMode2Joueurs")
    TabStripMonde.Tabs(7).Caption = FicIni.Parametre("BandeServeur")
    TabStripMonde.Tabs(8).Caption = FicIni.Parametre("BandeClient")
    TabStripMonde2.Tabs(1).Caption = FicIni.Parametre("BandeCampagne")
    TabStripMonde2.Tabs(2).Caption = FicIni.Parametre("BandePersonnaliser")
    TabStripMonde2.Tabs(3).Caption = FicIni.Parametre("BandeBataille")
    TabStripMonde2.Tabs(4).Caption = FicIni.Parametre("BandeScenario")
    TabStripMonde2.Tabs(5).Caption = FicIni.Parametre("BandeSauvegarde")
    TabStripMonde2.Tabs(6).Caption = FicIni.Parametre("BandeMode2Joueurs")
    TabStripMonde.Visible = ComReseau.Definir_Actif
    TabStripMonde2.Visible = Not ComReseau.Definir_Actif
    
    FrameCampagnes.Caption = FicIni.Parametre("CadreSelectionnerCampagne")
    FrameScenarioCampagnes.Caption = FicIni.Parametre("CadreSelectionnerScenarioCampagne")
    
    'Section monde personnalisé.
    FrameMonde.Caption = FicIni.Parametre("CadreMonde")
    FrameCarteGeneree.Caption = FicIni.Parametre("CadreCarte")
    FrameCarteCharger.Caption = FicIni.Parametre("CadreCarte")
    OptionCarte(0).Caption = FicIni.Parametre("EtiquetteMondeGenerer")
    OptionCarte(1).Caption = FicIni.Parametre("EtiquetteMondeChoisir")
    Label16.Caption = FicIni.Parametre("EtiquetteMondeTaille")
    Label8.Caption = FicIni.Parametre("EtiquetteMondeTerrain")
    Label18.Caption = FicIni.Parametre("EtiquetteMondeRessources")
    Label36.Caption = FicIni.Parametre("EtiquetteMondeRessourcesDepart")
    Label24.Caption = FicIni.Parametre("EtiquetteMondeEpoqueDepart")
    Label1.Caption = FicIni.Parametre("EtiquetteMondeLargeur")
    Label2.Caption = FicIni.Parametre("EtiquetteMondeHauteur")
    Label4.Caption = FicIni.Parametre("EtiquetteMondeRessourcesNombre")
    
    'Charge la liste des cartes.
    Cartes.Charger_Dans_Liste LstCartes
    
    Label37.Caption = FicIni.Parametre("EtiquetteMondeChangementEpoque")
    Label33.Caption = FicIni.Parametre("EtiquetteMondeDifficulte")
    Label38.Caption = FicIni.Parametre("EtiquetteMondeResurrectionsNombre")
    Label39.Caption = FicIni.Parametre("EtiquetteMondeResurrectionsVitesse")
    FrameFiefs.Caption = FicIni.Parametre("CadreFiefs")
    Label5.Caption = FicIni.Parametre("EtiquetteFiefsNombre")
    Label6.Caption = FicIni.Parametre("EtiquetteFiefsNumero")
    Label7.Caption = FicIni.Parametre("EtiquetteFiefsNombreVillageois")
    CheckEquilibreFiefs.Caption = FicIni.Parametre("BoutonFiefsEquilibrer")
    Label21.Caption = FicIni.Parametre("EtiquetteFiefsPeuples")
    Label26.Caption = FicIni.Parametre("EtiquetteFiefsEquipes")
    EtiquetteAleatoire = FicIni.Parametre("EtiquetteAleatoire")
    Label3.Caption = FicIni.Parametre("EtiquetteFiefsTotal")
    
    FrameArmees.Caption = FicIni.Parametre("CadreArmees")
    For i = 0 To 1
        FrameArmee(i).Caption = FicIni.Parametre("CadreArmee" & i + 1)
        Label35(i).Caption = FicIni.Parametre("EtiquetteArmeeEquipe") & (i + 1)
        Label20(i).Caption = FicIni.Parametre("EtiquetteArmeeNombre")
        Label23(i).Caption = FicIni.Parametre("EtiquetteArmeePeuple")
        CheckEquipementArmee(i).Caption = FicIni.Parametre("EtiquetteArmeeEquipement")
    Next i
    
    'Section scénario et sauvegarde.
    FrameChargerMonde.Caption = FicIni.Parametre("CadreSelectionnerScenario")
    
    'Section Réseau.
    FrameServeur.Caption = FicIni.Parametre("CadreServeur")
    Label43.Caption = FicIni.Parametre("EtiquetteAdresseEnLigne")
    Label30.Caption = FicIni.Parametre("EtiquetteNomSession")
    Label29.Caption = FicIni.Parametre("EtiquetteNombreJoueursMax")
    OptionJouerEnLigne(0).Caption = FicIni.Parametre("ChoixJouerEnLigne1")
    OptionJouerEnLigne(1).Caption = FicIni.Parametre("ChoixJouerEnLigne2")
    CmdCreerServeur.Caption = FicIni.Parametre("BoutonCreerServeur")
    CmdFermerServeur.Caption = FicIni.Parametre("BoutonFermerServeur")
    FrameClient.Caption = FicIni.Parametre("CadreClient")
    Label42.Caption = FicIni.Parametre("EtiquetteAdresseServeur")
    CmdRafraichirParties.Caption = FicIni.Parametre("BoutonRafraichirParties")
    CmdConnecter.Caption = FicIni.Parametre("BoutonConnecterParties")
    CmdDeconnecter.Caption = FicIni.Parametre("BoutonDeconnecterParties")
    FrameReseau.Caption = FicIni.Parametre("CadreParametresReseaux")
    Label28.Caption = FicIni.Parametre("EtiquettePseudoReseau")
    Label40.Caption = FicIni.Parametre("EtiquetteAdresseIP")
    Label41.Caption = FicIni.Parametre("EtiquetteNumeroPort")
    'CheckOuverturePort.Value = -CDbl(ComReseau.Definir_OuverturePort)
    Label41.Visible = ComReseau.Definir_OuverturePort
    TxtNumeroPort.Visible = ComReseau.Definir_OuverturePort
    Label27.Caption = FicIni.Parametre("EtiquetteServiceProvider")
    FrameInfoReseau.Caption = FicIni.Parametre("CadreInformationsReseaux")
    Label44.Caption = FicIni.Parametre("EtiquetteInformationsReseaux1")
    Label45.Caption = FicIni.Parametre("EtiquetteInformationsReseaux2")
    Frame3.Caption = FicIni.Parametre("CadreDetailPersonnageReseau")
    Frame2.Caption = FicIni.Parametre("CadreDetailPartieReseau")
    Label32.Caption = FicIni.Parametre("EtiquetteMessageReseau")
    CmdEnvoyerMessage.Caption = FicIni.Parametre("BoutonEnvoyerMessageReseau")

    
    'Section mode 2 joueurs.
    FrameMode2Joueurs.Caption = FicIni.Parametre("CadreMode2J")
    CheckJouera2.Caption = FicIni.Parametre("BoutonMode2JActiver")
    CheckMaintenirEcransSepares.Caption = FicIni.Parametre("BoutonMode2JSeparer")
    Frame8(1).Caption = FicIni.Parametre("CadreMode2JJoueur1")
    Frame8(0).Caption = FicIni.Parametre("CadreMode2JJoueur2")
    
    EtiquetteMode2JImpossible = FicIni.Parametre("EtiquetteMode2JImpossible")
    EtiquetteMode2JActive = FicIni.Parametre("EtiquetteMode2JActive")
    EtiquetteMode2JDesActive = FicIni.Parametre("EtiquetteMode2JDesActive")
    TitreMode2JImpossible = FicIni.Parametre("TitreMode2JIMpossible")
    MessageMode2JImpossible = FicIni.Parametre("MessageMode2JImpossible")
    
    EtiquetteReseauClient = FicIni.Parametre("EtiquetteReseauClient")
    EtiquetteReseauServeur = FicIni.Parametre("EtiquetteReseauServeur")
    
    List1.Clear: List2.Clear
    For i = 1 To 6
        List1.AddItem FicIni.Parametre("ListeMode2JJoueur1(" & i & ")")
        List2.AddItem FicIni.Parametre("ListeMode2JJoueur2(" & i & ")")
    Next i
    
    'Section personnage.
    TabStripPersonnage.Tabs(1).Caption = FicIni.Parametre("BandeNouveauPersonnage")
    TabStripPersonnage.Tabs(2).Caption = FicIni.Parametre("BandeChargerPersonnage")
    FrameNouveauPersonnage.Caption = FicIni.Parametre("CadreCreerPersonnage")
    FrameNomPersonnage.Caption = FicIni.Parametre("CadreNomPersonnage")
    ReDim ErreurNomPersonnage(2)
    ErreurNomPersonnage(0) = FicIni.Parametre("ErreurNomPersonnage1")
    ErreurNomPersonnage(1) = FicIni.Parametre("ErreurNomPersonnage2")
    ErreurNomPersonnage(2) = FicIni.Parametre("ErreurNomPersonnage3")
    FrameRacePersonnage.Caption = FicIni.Parametre("CadreRacePersonnage")
    FrameSexePersonnage.Caption = FicIni.Parametre("CadreSexePersonnage")
    OptionSexe(0).Caption = FicIni.Parametre("BoutonSexePersonnage1")
    OptionSexe(1).Caption = FicIni.Parametre("BoutonSexePersonnage2")
    LblRestictionSexe.Caption = FicIni.Parametre("ErreurSexePersonnage")
    FrameValidationPersonnage.Caption = FicIni.Parametre("CadreValidationPersonnage")
    FrameGenerationPersonnage.Caption = FicIni.Parametre("CadreGenerationPersonnage")
    CmdGenererPersos(0).Caption = FicIni.Parametre("BoutonGenerationPersonnage1")
    CmdGenererPersos(1).Caption = FicIni.Parametre("BoutonGenerationPersonnage2")
    CmdGenererPersos(2).Caption = FicIni.Parametre("BoutonGenerationPersonnage3")
    
    FrameChargerPersonnage.Caption = FicIni.Parametre("CadreSelectionnerPersonnage")
    ImgLocker(0).ToolTipText = FicIni.Parametre("ImageCadenatFerme")
    ImgLocker(1).ToolTipText = FicIni.Parametre("ImageCadenatOuvert")
    LblIA.Caption = FicIni.Parametre("EtiquetteIA")
    LblIA.ToolTipText = FicIni.Parametre("EtiquetteIAInfoBulle")
    Label22 = Parametres.EtiquetteNivPlus
    Label19 = Parametres.EtiquetteNiveau & " :"
    Label34 = Parametres.EtiquettePeuple & " :"
    Label31 = Parametres.EtiquetteRace & " :"
    Label17 = Parametres.EtiquetteSexe & " :"
    Label10 = Parametres.EtiquetteVie & " :"
    Label11 = Parametres.EtiquetteEnergie & " :"
    Label13 = Parametres.EtiquetteMagie & " :"
    Label12 = Parametres.EtiquetteMoral & " :"
    Label14 = Parametres.EtiquetteAttaque & " :"
    Label15 = Parametres.EtiquetteDefense & " :"
    TabStripFichiers.Tabs(1).Caption = FicIni.Parametre("BandeApparencePersonnage")
    TabStripFichiers.Tabs(2).Caption = FicIni.Parametre("BandeAIPersonnage")
    CmdApparence.Caption = FicIni.Parametre("BoutonPersonnaliserApparencePersonnage")
    OptionApparence(0).Caption = FicIni.Parametre("ChoixPersonnaliserApparencePersonnage")
    OptionApparence(1).Caption = FicIni.Parametre("ChoixFixerApparencePersonnage")
    OptionIAScript(0).Caption = FicIni.Parametre("ChoixIA1")
    OptionIAScript(1).Caption = FicIni.Parametre("ChoixIA2")
    OptionIAScript(2).Caption = FicIni.Parametre("ChoixIA3")
    TabStripCompBest.Tabs(1).Caption = FicIni.Parametre("BandeCompetencesPersonnage")
    TabStripCompBest.Tabs(2).Caption = FicIni.Parametre("BandeBestiairePersonnage")
    TabStripCompBest.Tabs(3).Caption = FicIni.Parametre("BandeTresorsPersonnage")
    
    CmdRetour.Caption = FicIni.Parametre("BoutonRetour")
    CmdCommencer.Caption = FicIni.Parametre("BoutonJouer")
    
    FicIni.fichier = FicIni.chemin & FichierINI
    'Chargement des informations de la langue.
    FicIni2.fichier = FicIni2.chemin & Langues.Dossier & FichierINI
    
    'Charge les campagnes.
    Actualiser_Liste_Campagnes
    FicIni.Section = "Campagnes"
    If Val(FicIni.Parametre("DerniereCampagne")) >= 0 And _
       Val(FicIni.Parametre("DerniereCampagne")) < LstCampagnes.ListCount Then
        LstCampagnes.ListIndex = Val(FicIni.Parametre("DerniereCampagne"))
    End If
    If Val(FicIni.Parametre("DernierScenario")) >= 0 And _
       Val(FicIni.Parametre("DernierScenario")) < LstCampagnesScenarios.ListCount Then
        LstCampagnesScenarios.ListIndex = Val(FicIni.Parametre("DernierScenario"))
    End If
    
    'Paramètres des fiefs.
    FicIni.Section = "Fiefs"
    SliderFiefs.Min = FicIni.Parametre("MinFiefs")
    SliderFiefs.Value = FicIni.Parametre("ValFiefs")
    SliderFiefs.Max = FicIni.Parametre("MaxFiefs")
    
    SliderVillageois(0).Min = FicIni.Parametre("MinVillageoisFief")
    SliderVillageois(0).Max = FicIni.Parametre("MaxVillageoisFief")
    
    Me.Top = 0
    Me.Left = 0
    For i = 0 To SliderFiefs.Max - 1
        If i > 0 Then
            Load LblNumeroFief(i)
            
            LblNumeroFief(i).Top = LblNumeroFief(i - 1).Top + HauteurLignesFiefs
            
            LblNumeroFief(i).Visible = True
            Load SliderVillageois(i)
            SliderVillageois(i).Top = SliderVillageois(i - 1).Top + HauteurLignesFiefs
            SliderVillageois(i).Visible = True
            Load LblVillageois(i)
            LblVillageois(i).Top = LblVillageois(i - 1).Top + HauteurLignesFiefs
            LblVillageois(i).Visible = True
            Load ComboPeuple(i)
            ComboPeuple(i).Top = ComboPeuple(i - 1).Top + HauteurLignesFiefs
            ComboPeuple(i).Visible = True
            Load ComboEquipe(i)
            ComboEquipe(i).Top = ComboEquipe(i - 1).Top + HauteurLignesFiefs
            ComboEquipe(i).Visible = True
        End If
        LblNumeroFief(i) = i + 1
        For j = 0 To Parametres.NombrePeuples - 1
            ComboPeuple(i).AddItem Parametres.Peuples_Nom(j)
        Next j
        ComboPeuple(i).AddItem EtiquetteAleatoire
        SliderVillageois(i).Value = FicIni.Parametre("ValVillageoisFief" & i + 1)
        If FicIni.Parametre("PeupleFief" & i + 1) <= ComboPeuple(i).ListCount Then
            ComboPeuple(i).ListIndex = FicIni.Parametre("PeupleFief" & i + 1) - 1
        Else
            ComboPeuple(i).ListIndex = ComboPeuple(i).ListCount - 1
        End If
    Next i
    ComboEquipe(0).Enabled = False
    
    For i = 0 To SliderFiefs.Max - 1
        ComboEquipe(i).Clear
        For j = 1 To SliderFiefs.Max
            ComboEquipe(i).AddItem j
        Next j
        ComboEquipe(i).ListIndex = FicIni.Parametre("EquipeFief" & i + 1) - 1
    Next i
    For i = 1 To SliderFiefs.Max
        ComboEquipePerso.AddItem i
    Next i
    ComboEquipePerso.ListIndex = 0
    CheckEquilibreFiefs.Value = FicIni.Parametre("EquilibreFiefs")
    
    FicIni.Section = "Monde"
    FicIni2.Section = "Monde"
    
    i = Val(FicIni.Parametre("TypePartie")) + 1
    If i < 1 Or i > TabStripMonde2.Tabs.Count Then i = 1
    TabStripMonde.Tabs(i).Selected = True
    TabStripMonde2.Tabs(i).Selected = True
    OptionCarte(Val(FicIni.Parametre("TypeCarte")) - 1).Value = True
    'Charge la dernière carte jouée.
    If OptionCarte(1).Enabled Then 'Seulement si l'on a au moins une carte chargée.
        If Val(FicIni.Parametre("DerniereCarte")) > LstCartes.ListCount Then
            LstCartes.ListIndex = 0
        Else
            LstCartes.ListIndex = Val(FicIni.Parametre("DerniereCarte")) - 1
        End If
        LstCartes_Click
        OptionCarte_Click Val(FicIni.Parametre("TypeCarte")) - 1
    Else
        OptionCarte(0).Value = True
        OptionCarte_Click 0
    End If
    DernierScenario = Val(FicIni.Parametre("DernierScenario"))
    DernierMonde = Val(FicIni.Parametre("DernierMonde"))
    DernierePartie = Val(FicIni.Parametre("DernierePartie"))
    TabStripMonde_Click
    
    'Paramètres des combos box.
    'Combo taille.
    ReDim TailleLargeur(FicIni.Parametre("NombreTailles") - 1)
    ReDim TailleHauteur(UBound(TailleLargeur()))
    ReDim TailleRessources(UBound(TailleLargeur()))
    ReDim TailleFiefs(UBound(TailleLargeur()))
    For i = 0 To UBound(TailleLargeur())
        ComboTaille.AddItem FicIni2.Parametre("Taille" & i + 1 & "Nom")
        TailleLargeur(i) = Left(FicIni.Parametre("Taille" & i + 1 & "Dimensions"), InStr(FicIni.Parametre("Taille" & i + 1 & "Dimensions"), "*") - 1)
        TailleHauteur(i) = Right(FicIni.Parametre("Taille" & i + 1 & "Dimensions"), Len(FicIni.Parametre("Taille" & i + 1 & "Dimensions")) - InStr(FicIni.Parametre("Taille" & i + 1 & "Dimensions"), "*"))
        TailleRessources(i) = FicIni.Parametre("Taille" & i + 1 & "Ressources")
        TailleFiefs(i) = FicIni.Parametre("Taille" & i + 1 & "Fiefs")
    Next i
    ComboTaille.ListIndex = FicIni.Parametre("TailleChoisie") - 1
    'Combo terrain.
    For i = 0 To Monde.NombreTypesMondes - 1
        ComboTerrain.AddItem Monde.Param_MondeNom(i)
    Next i
    ComboTerrain.ListIndex = FicIni.Parametre("TerrainChoisi") - 1
    'Combo ressources.
    ReDim RessourcesMultiplicateur(FicIni.Parametre("NombreRessources") - 1)
    For i = 0 To FicIni.Parametre("NombreRessources") - 1
        ComboRessources.AddItem FicIni2.Parametre("Ressources" & i + 1 & "Nom")
        RessourcesMultiplicateur(i) = Val(FicIni.Parametre("Ressources" & i + 1 & "Multiplicateur"))
    Next i
    ComboRessources.ListIndex = FicIni.Parametre("RessourceChoisie") - 1
    'Combo des époques.
    For i = 0 To Parametres.NombreEpoques - 1
        ComboEpoque.AddItem Parametres.Epoque_Nom(i)
    Next i
    ComboEpoque.ListIndex = FicIni.Parametre("EpoqueChoisie") - 1
    For i = 0 To Parametres.NombreRessourcesDepart - 1
        ComboRessourcesDepart.AddItem Parametres.RessourcesDeparts_Nom(i)
    Next i
    ComboRessourcesDepart.ListIndex = FicIni.Parametre("RessourcesDepartsChoisie") - 1
    'Combo des vitesses d'époques.
    For i = 0 To Parametres.NombreVitesseEpoques - 1
        ComboVitesseEpoque.AddItem Parametres.VitesseEpoque_Nom(i)
    Next i
    ComboVitesseEpoque.ListIndex = FicIni.Parametre("VitesseEpoqueChoisie") - 1
    'Combo des difficultés.
    j = FicIni.Parametre("NombreDifficultes") - 1
    ReDim DifficulteBonusXP(j)
    ReDim DifficulteCoefVieChateaux(j)
    For i = 0 To j
        ComboDifficulte.AddItem FicIni2.Parametre("Difficulte" & i + 1 & "Nom")
        DifficulteBonusXP(i) = FicIni.Parametre("Difficulte" & i + 1 & "BonusXP")
        DifficulteCoefVieChateaux(i) = Val(FicIni.Parametre("Difficulte" & i + 1 & "CoefVieChateaux"))
    Next i
    ComboDifficulte.ListIndex = FicIni.Parametre("DifficulteChoisie") - 1
    
    'Paramètres de la carte.
    SliderLargeurCarte.Min = FicIni.Parametre("MinLargeur")
    SliderHauteurCarte.Min = FicIni.Parametre("MinHauteur")
    SliderNombreArbres.Min = FicIni.Parametre("MinQuantitéRessources")
    SliderLargeurCarte.Max = FicIni.Parametre("MaxLargeur")
    SliderHauteurCarte.Max = FicIni.Parametre("MaxHauteur")
    SliderNombreArbres.Max = FicIni.Parametre("MaxQuantitéRessources")
    SliderLargeurCarte.Value = FicIni.Parametre("ValLargeur")
    SliderHauteurCarte.Value = FicIni.Parametre("ValHauteur")
    SliderNombreArbres.Value = FicIni.Parametre("ValQuantitéRessources")
    
    'Combo des resurrections.
    'Nombre de resurrections.
    FicIni.Section = "Resurrections_Nombre"
    FicIni2.Section = "Resurrections_Nombre"
    ReDim ResurrectionsValeur(FicIni.Parametre("NombreResurrections") - 1)
    For i = 0 To UBound(ResurrectionsValeur())
        ComboResurrections(0).AddItem FicIni2.Parametre("Resurrection" & i + 1 & "Nom")
        ResurrectionsValeur(i) = Val(FicIni.Parametre("Resurrection" & i + 1 & "Valeur"))
    Next i
    ComboResurrections(0).ListIndex = FicIni.Parametre("ResurrectionChoisie") - 1
    'Vitesse de resurrection.
    FicIni.Section = "Resurrections_Vitesse"
    FicIni2.Section = "Resurrections_Vitesse"
    ReDim ResurrectionsCoefs(FicIni.Parametre("NombreResurrections") - 1)
    For i = 0 To UBound(ResurrectionsCoefs())
        ComboResurrections(1).AddItem FicIni2.Parametre("Resurrection" & i + 1 & "Nom")
        ResurrectionsCoefs(i) = Val(FicIni.Parametre("Resurrection" & i + 1 & "Valeur"))
    Next i
    ComboResurrections(1).ListIndex = FicIni.Parametre("ResurrectionChoisie") - 1
    
    'Paramètres des armées.
    For i = 0 To FrameArmee.Count - 1
        FicIni.Section = "Armée_" & i + 1
        SliderNombreSoldats(i).Min = Val(FicIni.Parametre("MinSoldats"))
        SliderNombreSoldats(i).Value = Val(FicIni.Parametre("ValSoldats"))
        SliderNombreSoldats(i).Max = Val(FicIni.Parametre("MaxSoldats"))
        For j = 0 To Parametres.NombrePeuples - 1
            ComboArmee(i).AddItem Parametres.Peuples_Nom(j)
        Next j
        ComboArmee(i).AddItem EtiquetteAleatoire
        
        If FicIni.Parametre("Peuple") <= ComboArmee(i).ListCount Then
            ComboArmee(i).ListIndex = FicIni.Parametre("Peuple") - 1
        Else
            ComboArmee(i).ListIndex = ComboArmee(i).ListCount - 1
        End If
        'Equipement des armées.
        LstObjets(i).Clear
        For j = 0 To Parametres.NombreObjets - 1
            If Not Parametres.Objet_Unique(j) Then
                If Parametres.Objet_PrixVente(j) > 0 Then
                    LstObjets(i).AddItem Parametres.Objet_Nom(j)
                End If
            End If
        Next j
        CheckEquipementArmee(i).Value = FicIni.Parametre("Equipement")
        If CheckEquipementArmee(i).Value = 1 Then
            For j = 0 To LstObjets(i).ListCount - 1
               LstObjets(i).Selected(j) = Val(FicIni.Parametre("Objet" & j))
            Next j
        End If
        CheckEquipementArmee_Click Int(i)
    Next i
    'Ligne concernant l'armée du joueur.
    'ComboArmee(0).ListIndex = 0
    'ComboArmee(0).Enabled = False
    
    SliderHauteurCarte_Scroll
    SliderLargeurCarte_Scroll
    SliderNombreArbres_Scroll
    SliderFiefs_Scroll
    For i = 0 To SliderFiefs.Max - 1
        SliderVillageois_Scroll (i)
    Next i
    SliderNombreSoldats_Scroll 0
    SliderNombreSoldats_Scroll 1
        
    ReDim Persos(0)
    Set Persos(0) = New ClsJeuPerso
    
    Persos(0).Joueur = ComReseau.PlayerID
    Persos(0).NumeroFief = 0
    ReDim Maisons(0)
    Set Maisons(0) = New ClsJeuBatiment
    NoPerso = 0
    
    Persos(NoPerso).Init NoPerso, Parametres
    FilePerso.Filename = App.Path & CheminSavPerso
    FilePerso.Pattern = "*" & ExtensionFichiersPerso
    TxtNomPerso_Change
    
    Charger_Dossier_IASripts
    FileIA.Filename = App.Path & "\" & CheminIA
    
    Actualiser_Liste_Persos
    
    FicIni.Section = "Personnages"
    'Charge les images des cadenats de protection des personnages.
    ImgLocker(0).Picture = LoadPicture(App.Path & "\Images\Curseurs\Locker0.gif")
    ImgLocker(1).Picture = LoadPicture(App.Path & "\Images\Curseurs\LockerC.gif")
    
    'Affiche ou non toutes les compétences.
    MDIFrmMain.MenuPersonnageCompetencesAfficherToutes.Checked = -CDbl(FicIni.Parametre("CompetencesAfficherToutes"))
    MDIFrmMain.MenuPersonnageCompetencesTri_Click FicIni.Parametre("CompetencesTri")
    MDIFrmMain.MenuPersonnageCompetencesAffichage_Click FicIni.Parametre("CompetencesAffichage")
    If LstPerso.ListCount > 0 Then
        For i = 0 To Val(FicIni.Parametre("NombrePersonnages")) - 1
            If i < LstPerso.ListCount Then
                LstPerso.Selected(i) = Val(FicIni.Parametre("Personnage" & i))
            End If
        Next i
        
        If Val(FicIni.Parametre("DernierPersonnage")) < LstPerso.ListCount Then
            'FilePerso.ListIndex = Val(FicIni.Parametre("DernierPersonnage"))
            LstPerso.ListIndex = Val(FicIni.Parametre("DernierPersonnage"))
            LstPerso.Selected(LstPerso.ListIndex) = False
            LstPerso.Selected(LstPerso.ListIndex) = True
        Else
            LstPerso.ListIndex = 0
            LstPerso.Selected(0) = False
            LstPerso.Selected(0) = True
        End If
        LstPerso.Selected(LstPerso.ListIndex) = True
        TabStripPersonnage.Tabs(2).Selected = True
    Else
        TabStripPersonnage.Tabs(1).Selected = True
    End If
    TabStripFichiers.Tabs(FicIni.Parametre("FichiersSelection") + 1).Selected = True
    TabStripCompBest.Tabs(FicIni.Parametre("CompetencesSelection") + 1).Selected = True
    If Val(FicIni.Parametre("PersonnageJoueur1")) < ComboNomJoueur(0).ListCount Then
        ComboNomJoueur(0).ListIndex = Val(FicIni.Parametre("PersonnageJoueur1"))
    End If
    ComboNomJoueur_Click 0
    
    Actualiser_Joueurs
    
    If Val(FicIni.Parametre("PersonnageJoueur2")) < ComboNomJoueur(1).ListCount Then
        ComboNomJoueur(1).ListIndex = Val(FicIni.Parametre("PersonnageJoueur2"))
    End If
    
    'Chargement de la liste des races.
    For i = 0 To Parametres.NombreRaces - 1
        If Parametres.JouerToutesLesRaces Or Parametres.Race_Jouable(i) Then
            LstRaces.AddItem Parametres.Race_Nom(i)
        End If
    Next i
    LstRaces.ListIndex = 0
    
    'Chargement des informations du mode 2 joueurs.
    CheckJouera2.Value = -CDbl(Jeu.Definir_Mode2Joueurs)
    CheckMaintenirEcransSepares.Value = -CDbl(Jeu.Definir_Mode2JoueursEcranSepares)
    
    'Chargement des informations réseaux.
    With ComReseau
    OptionJouerEnLigne(-CDbl(.Definir_JouerEnLigne)) = True
    TxtNomJoueur = .NomJoueur
    TxtNomSession = .NomSession
    If .NombreJoueursMax = .NombreJoueursMin Then
        SliderNombreJoueurs.Visible = False
        LblNombreJoueurs.Left = SliderNombreJoueurs.Left
    Else
        SliderNombreJoueurs.Min = .NombreJoueursMin
        SliderNombreJoueurs.Max = .NombreJoueursMax
    End If
    SliderNombreJoueurs.Value = .NombreJoueurs
    SliderNombreJoueurs_Scroll
    
    'LblAdresseIP.ToolTipText = recup_ip(0).IPAdr & vbCrLf & recup_ip(1).IPAdr & vbCrLf & recup_ip(2).IPAdr & vbCrLf & recup_ip(3).IPAdr
    For j = 0 To 3
        If recup_ip(j).IPAdr <> "127.0.0.1" And _
           recup_ip(j).IPAdr <> "0.0.0.0" And _
           Left(recup_ip(j).IPAdr, 8) <> "192.168." Then
            LblAdresseINet.Caption = recup_ip(j).IPAdr
        End If
        If Left(recup_ip(j).IPAdr, 8) = "192.168." Then
            LblAdresseIP.Caption = recup_ip(j).IPAdr
        End If
    Next j
    'Affiche les services providers.
    TxtNumeroPort = ComReseau.Definir_NumeroPort
    TxtNumeroIPServeur = ComReseau.Definir_NumeroIPServeur
    
    For i = 1 To .lngNumConnexions
        LstServices.AddItem .ServiceProvider_Nom(i)
    Next i
    'If LstServices.ListIndex < .ServiceProvider - 1 Then
    '    LstServices.ListIndex = 0
    'Else
    LstServices.ListIndex = LstServices.ListCount - 1
    'End If
    End With
    Actualiser_Boutons_Reseau
    If ComReseau.Serveur Then
        TabStripMonde.Tabs("CreerServeur").Selected = True
    ElseIf ComReseau.Client Then
         TabStripMonde.Tabs("RejoindrePartie").Selected = True
    End If
    
    ChargerPerso = True
    EnregistrerPerso = True
    Actualiser_Infos_Perso
    
    Exit Sub
Erreur:
    MsgBox "Une erreur s'est produit pendant le chargement de :" & FicIni.DernierChargement & "(" & FicIni.fichier & ").", vbCritical, "Erreur"
    End
End Sub

Private Sub Charger_Dossier_IASripts()
    On Error GoTo Erreur
    FileIAScript.Filename = App.Path & "\" & CheminIAScript
    FichiersIAScripts = True
    Exit Sub
Erreur:
    FichiersIAScripts = False
End Sub

Private Sub ComboTaille_Click()
    If ComboRessources.ListIndex >= 0 And _
       SliderVillageois().Count >= TailleFiefs(0) Then
        SliderLargeurCarte.Value = TailleLargeur(ComboTaille.ListIndex)
        SliderHauteurCarte.Value = TailleHauteur(ComboTaille.ListIndex)
        SliderNombreArbres.Value = CLng(TailleRessources(ComboTaille.ListIndex) * RessourcesMultiplicateur(ComboRessources.ListIndex))
        If SliderFiefs.Enabled Then SliderFiefs.Value = TailleFiefs(ComboTaille.ListIndex)
        SliderHauteurCarte_Scroll
        SliderLargeurCarte_Scroll
        SliderNombreArbres_Scroll
        If TabStripMonde.Tabs("NouveauMonde").Selected Then
            SliderFiefs_Scroll
        End If
    End If
End Sub

Private Sub ComboRessources_Click()
    ComboTaille_Click
End Sub

Private Sub FileMonde_Click()
    On Error Resume Next
    Dim Temp As String
    'LblDescriptionMonde.Caption = ""
    Open FileMonde.Path & "\" & Left$(FileMonde.List(FileMonde.ListIndex), InStr(FileMonde.List(FileMonde.ListIndex), ".") - 1) & ".txt" For Input As #1
    Line Input #1, Temp
    LblDescriptionMonde.Caption = Temp
    If LblDescriptionMonde.Caption <> "" Then
        While Not EOF(1)
            Line Input #1, Temp
            LblDescriptionMonde.Caption = LblDescriptionMonde.Caption & Chr(13) & Temp
        Wend
    End If
    Close #1
    ImgApercu.Picture = LoadPicture("")
    ImgApercu.Picture = LoadPicture(FileMonde.Path & "\" & Left$(FileMonde.List(FileMonde.ListIndex), InStr(FileMonde.List(FileMonde.ListIndex), ".") - 1) & ".bmp")
    Actualiser_Bouton_Commencer
End Sub

Public Sub CmdCommencer_Click()
    'On lance la partie.
    Partie.EnCours = 1
    If TabStripMonde.Tabs("Mode2Joueurs").Selected Or TabStripMonde.Tabs("CreerServeur").Selected Then
        TabStripMonde.Tabs(DernierePartie + 1).Selected = True
    End If
    TimerApparence.Enabled = False
    Unload FrmInfoRaces
    Me.Hide
    MDIFrmMain.Hide
    
    Jeu.Campagne = TabStripMonde.Tabs("Campagnes").Selected
    FrmMoteur2D.Show
    DoEvents
    FrmMoteur2D.MoteurJeu
    
    MDIFrmMain.MousePointer = 11
    MDIFrmMain.Caption = App.Title & " - déchargement..."
    Me.Show
    Unload FrmMoteur2D
    DoEvents
    Unload Me
    'Form_Unload 0
    'Form_Load
    'MDIFrmMain.WindowState = 2
    Load Me
    MDIFrmMain.Show
    MDIFrmMain.Caption = App.Title
    MDIFrmMain.MousePointer = 0
    
    If Definir_AfficherStastistiquesEnQuittant Then
        Afficher_Stastistiques
    End If
End Sub

Public Sub CmdCreerPerso_Click()
    Dim i As Long
    Persos(NoPerso).Init NoPerso, Parametres
    Persos(NoPerso).Nom = UCase(Left(TxtNomPerso, 1)) & Right(TxtNomPerso, Len(TxtNomPerso) - 1)

    Persos(NoPerso).Changer_Race_Personnage Parametres.Race_Indice(LstRaces.List(LstRaces.ListIndex)), Parametres
    
    Persos(NoPerso).Feminin = OptionSexe(1).Value
    Persos(NoPerso).ChoisirIA = False
    Persos(NoPerso).FichierIA = FichierIADefaut
    'If Persos(Noperso).Nombre_ApparenceBase > 0 Then
    '    CmdApparence_Click
    '    FrmApparence.CmdAleatoire_Click
    '    FrmApparence.CmdQuitter_Click
    'End If
    AffApparence.Tirer_Apparence_Au_Hasard Persos(NoPerso), Parametres
    
    Enregistrer_Sauvegarde_Perso Persos(NoPerso), True
    'Par défaut, l'utilisation automatique des batiments est coché.
    Enregistrer_Sauvegarde_Interface_Defaut Persos(NoPerso)
    
    Actualiser_Liste_Persos
    'Selectionne le nouveau personnage créé.
    For i = 0 To LstPerso.ListCount - 1
        If LstPerso.List(i) = Persos(NoPerso).Nom Then
            LstPerso.Selected(i) = True
            i = LstPerso.ListCount
        End If
    Next i
    'Active les talents (si possible)
    TabStripCompBest.Tabs(3).Selected = True
    For i = 0 To LstTalents.ListCount - 1
        If Parametres.Objet_PrixTalents(Parametres.Race_Objet(Persos(NoPerso).Race, i)) <= 0 Then
            LstTalents.Selected(i) = True
        End If
    Next i
    'LstPerso_Click
    TxtNomPerso = ""
    TabStripPersonnage.Tabs(2).Selected = True
    CmdSupprimerPerso.Enabled = True
End Sub

Public Sub CmdSupprimerPerso_Click()
    Dim i As Long
    If Not LblNombrePersos.Visible Then
        If MsgBox("Etes-vous sûr de vouloir effacer " & LstPerso.List(LstPerso.ListIndex) & " ?", vbOKCancel + vbQuestion, "Confirmation de la suppression du personnage") = vbOK Then
            Supprimer_Sauvegarde_Perso LstPerso.List(LstPerso.ListIndex)
            Actualiser_Liste_Persos
            If LstPerso.ListCount > 0 Then
                LstPerso.Selected(0) = True
                LstPerso.ListIndex = 0
            Else
                CmdSupprimerPerso.Enabled = False
            End If
            LstPerso_Click
            ComboNomJoueur_Click 0
        End If
    Else
        If MsgBox("Etes-vous sûr de vouloir effacer les personnages sélectionés  ?", vbOKCancel + vbQuestion, "Confirmation de la suppression du personnage") = vbOK Then
            For i = 0 To LstPerso.ListCount - 1
                If LstPerso.Selected(i) Then
                    Supprimer_Sauvegarde_Perso LstPerso.List(i)
                End If
            Next i
            Actualiser_Liste_Persos
            If LstPerso.ListCount > 0 Then
                LstPerso.Selected(0) = True
                LstPerso.ListIndex = 0
            Else
                CmdSupprimerPerso.Enabled = False
            End If
            LstPerso_Click
            ComboNomJoueur_Click 0
        End If
    End If
End Sub

Private Sub FileApparence_Click()
    If FileApparence.ListCount > 0 And _
       FileApparence.Filename <> "" Then
        'PicApparence.Picture = LoadPicture(App.Path & AffPerso.CheminHeros(Persos(NoPerso).Feminin) & FileApparence.FileName)
        'Persos(NoPerso).ChoisirApparence = True
        Persos(NoPerso).FichierApparence = FileApparence.Filename
        If OptionApparence(1).Value Then
            ImgApparence.Picture = LoadPicture(FileApparence.Path & "\" & FileApparence.Filename)
        End If
    End If
    If EnregistrerPerso Then
        Enregistrer_Sauvegarde_Perso Persos(NoPerso)
    End If
End Sub

Private Sub FileIA_Click()
    If FileIA.ListCount > 0 And _
       FileIA.Filename <> "" Then
        Persos(NoPerso).FichierIA = FileIA.Filename
    End If
    If EnregistrerPerso Then
        Enregistrer_Sauvegarde_Perso Persos(NoPerso)
    End If
End Sub

Private Sub LstMonde_Click()
    FileMonde.ListIndex = LstMonde.ListIndex
    If TabStripMonde.Tabs("ChargerScenario").Selected Then
        DernierScenario = LstMonde.ListIndex
    Else
        DernierMonde = LstMonde.ListIndex
    End If
End Sub

Public Sub LstPerso_Click()
    Dim i As Integer
    EnregistrerPerso = False
    MDIFrmMain.MousePointer = 11
    Actualiser_Infos_Perso
    NombrePersonnages = 0
    'Vérifie si le joueur a sélectionné plus de personnages que de villageois dans le fief n°1.
    For i = 0 To LstPerso.ListCount - 1
        If LstPerso.Selected(i) Then
            NombrePersonnages = NombrePersonnages + 1
            'If LstPerso.List(i) <> LblNom Then
            '    LblNomJoueur(1) = LstPerso.List(i)
            'End If
        End If
    Next i
    'LblNomJoueur(0) = LblNom
    If NombrePersonnages > 1 Then
        LblNombrePersos.Visible = True
        LblNombrePersos.Caption = NombrePersonnages & " " & Parametres.EtiquettePersonnages
    Else
        LblNombrePersos.Visible = False
        'LblNombrePersos.Caption = ""
    End If
    'If j = 2 Then
    '    LblSelectionner2Personnages.Visible = False
    '    CheckJouera2.Enabled = True
    'Else
    '    LblSelectionner2Personnages.Visible = True
    '    CheckJouera2.Enabled = False
    'End If
    'If j > Val(LblVillageois(0)) Then
    '    LblVillageois(0).Caption = j
    'End If
    'If j > Val(LblNombreSoldats(0)) Then
    '    LblNombreSoldats(0).Caption = j
    'End If
    Actualiser_Joueurs
    Actualiser_Bouton_Commencer
    MDIFrmMain.MousePointer = 0
    EnregistrerPerso = True
End Sub

Private Sub LstRaces_Click()
    Dim i As Integer
    'Indice de la race sélectionnée.
    i = Parametres.Race_Indice(LstRaces.List(LstRaces.ListIndex))
    With Parametres
    TxtRaceDecription = .Race_Description(i)
    If .Race_RestrictionSexe(i, True) Or .Race_RestrictionSexe(i, False) Then
        OptionSexe(0).Enabled = False
        OptionSexe(1).Enabled = False
        LblRestictionSexe.Visible = True
        If .Race_RestrictionSexe(i, True) Then
            OptionSexe(0).Value = True
        ElseIf .Race_RestrictionSexe(i, False) Then
            OptionSexe(1).Value = True
        End If
    Else
        OptionSexe(0).Enabled = True
        OptionSexe(1).Enabled = True
        LblRestictionSexe.Visible = False
    End If
    End With
    CmdInfoRace_Click
End Sub

Private Sub LstRaces_DblClick()
    LstRaces_Click
    CmdInfoRace_Click
End Sub

Private Sub LstServices_Click()
    ComReseau.Definir_ServiceProvider = LstServices.ListIndex + 1
End Sub

Public Sub OptionApparence_Click(Index As Integer)
    Dim i As Long
    On Error GoTo Erreur
    OptionApparence(Index).Value = True
    FileApparence.Enabled = OptionApparence(1).Value
    CmdApparence.Enabled = OptionApparence(0).Value
    If Index = 1 Then
        'On coche le choix de l'apparence.
        'FileApparence.Visible = True
        'ImgApparence.Visible = True
        Persos(NoPerso).ChoisirApparence = True
        'CmdApparence.Enabled = False
        ImgApparence.Picture = LoadPicture(FileApparence.Path & "\" & FileApparence.Filename)
    Else
        'FileApparence.Visible = False
        'ImgApparence.Visible = True
        Persos(NoPerso).ChoisirApparence = False
        'CmdApparence.Enabled = True
        ImgApparence.Picture = LoadPicture(App.Path & CheminSavPerso & Persos(NoPerso).Nom & ".bmp")
        'If ChargerPerso Then CmdApparence_Click
    End If
    If EnregistrerPerso Then
        Enregistrer_Sauvegarde_Perso Persos(NoPerso)
    End If
    Exit Sub
Erreur: 'On n'arrive pas à ouvrir l'image.
    'CmdApparence_Click
    'FrmApparence.CmdQuitter_Click
    'ImgApparence.Picture = LoadPicture(App.Path & CheminSavPerso & Persos(Noperso).Nom & ".bmp")
    ImgApparence.Picture = Nothing
    If ChargerPerso Then
        Enregistrer_Sauvegarde_Perso Persos(NoPerso)
    End If
End Sub

Private Sub OptionCarte_Click(Index As Integer)
    Select Case Index
    Case 0: FrameCarteGeneree.Visible = True
            FrameCarteCharger.Visible = False
            SliderFiefs.Enabled = True
            LblFiefs.BackColor = &H80000005
    Case 1: FrameCarteGeneree.Visible = False
            FrameCarteCharger.Visible = True
            LstCartes_Click
    End Select
End Sub

Private Sub OptionIAScript_Click(Index As Integer)
    'OptionIAScript(Index).Value = True
    OptionIAScript(Index).Value = True
    FileIA.Visible = OptionIAScript(1).Value
    FileIAScript.Visible = OptionIAScript(2).Value 'And FichiersIAScripts
    Persos(NoPerso).ChoisirIA = OptionIAScript(1).Value Or OptionIAScript(2).Value
    Persos(NoPerso).ChoisirIAScript = OptionIAScript(2).Value
    If ChargerPerso Then
        Enregistrer_Sauvegarde_Perso Persos(NoPerso)
    End If
End Sub

Private Sub SliderFiefs_Scroll()
    Dim i As Long
    LblFiefs = SliderFiefs.Value
    For i = 0 To SliderFiefs.Max - 1
        If i < SliderFiefs.Value Then
            LblNumeroFief(i).Visible = True
            SliderVillageois(i).Visible = True
            LblVillageois(i).Visible = True
            ComboPeuple(i).Visible = True
            ComboEquipe(i).Visible = True
        Else
            LblNumeroFief(i).Visible = False
            SliderVillageois(i).Visible = False
            LblVillageois(i).Visible = False
            ComboPeuple(i).Visible = False
            ComboEquipe(i).Visible = False
        End If
        SliderVillageois_Scroll (i)
    Next i
End Sub

Private Sub SliderHauteurCarte_Scroll()
    LblHauteurCarte = SliderHauteurCarte.Value
End Sub

Private Sub SliderLargeurCarte_Scroll()
    LblLargeurCarte = SliderLargeurCarte.Value
End Sub

Private Sub SliderNombreArbres_Scroll()
    LblNombreArbres = SliderNombreArbres.Value
End Sub

Private Sub SliderNombreJoueurs_Scroll()
    LblNombreJoueurs = SliderNombreJoueurs.Value - 1
    ComReseau.Definir_NombreJoueurs = SliderNombreJoueurs.Value
End Sub

Private Sub SliderNombreSoldats_Scroll(Index As Integer)
    LblNombreSoldats(Index) = SliderNombreSoldats(Index).Value
End Sub

Private Sub SliderVillageois_Scroll(Index As Integer)
    Dim i As Long
    If CheckEquilibreFiefs.Value = 1 Then
        For i = 0 To SliderFiefs.Max - 1
            LblVillageois(i) = SliderVillageois(Index).Value
            SliderVillageois(i).Value = SliderVillageois(Index).Value
        Next i
    Else
        LblVillageois(Index) = SliderVillageois(Index).Value
    End If
    'Calcul le total de villageois.
    LblTotalVillageois = 0
    For i = 0 To SliderFiefs.Value - 1
        LblTotalVillageois = Val(LblTotalVillageois) + Val(LblVillageois(i))
    Next i
End Sub

Private Sub TabStripCompBest_Click()
    If TabStripCompBest.Tabs(1).Selected Then
        Actualiser_Infos_Perso_Competences
        LstCompetences.Visible = True
    Else
        LstCompetences.Visible = False
    End If
    If TabStripCompBest.Tabs(2).Selected Then
        Actualiser_Infos_Perso_Bestiaire
        LstBestiaire.Visible = True
    Else
        LstBestiaire.Visible = False
    End If
    If TabStripCompBest.Tabs(3).Selected Then
        Actualiser_Infos_Perso_Talents
        LstPointsTalents.Visible = True
        LstTalents.Visible = True
    Else
        LstPointsTalents.Visible = False
        LstTalents.Visible = False
    End If
End Sub

Private Sub TabStripFichiers_Click()
    ChargerPerso = False
    FileApparence.Visible = TabStripFichiers.Tabs(1).Selected
    OptionApparence(0).Visible = TabStripFichiers.Tabs(1).Selected
    OptionApparence(1).Visible = TabStripFichiers.Tabs(1).Selected
    CmdApparence.Visible = TabStripFichiers.Tabs(1).Selected
    'FileIA.Visible = TabStripFichiers.Tabs(2).Selected
    'CheckIAScript.Visible = TabStripFichiers.Tabs(2).Selected
    If TabStripFichiers.Tabs(2).Selected Then
        'OptionIAScript(0).Enabled = Persos(NoPerso).Protege
        'OptionIAScript(0).Visible = True
        OptionIAScript(1).Visible = True
        OptionIAScript(2).Visible = FichiersIAScripts
        If Persos(NoPerso).ChoisirIAScript And FichiersIAScripts Then
            OptionIAScript(2).Value = True
            'OptionIAScript_Click 2
        ElseIf Persos(NoPerso).ChoisirIA Then
            OptionIAScript(1).Value = True
            'OptionIAScript_Click 0
        Else
            OptionIAScript(0).Value = True
            'OptionIAScript_Click 1
        End If
    Else
        Actualiser_Infos_Apparences
        'OptionIAScript(0).Visible = False
        OptionIAScript(1).Visible = False
        OptionIAScript(2).Visible = False
        FileIA.Visible = False
        FileIAScript.Visible = False
    End If
    'CheckApparence_Click
    ChargerPerso = True
End Sub

Private Sub TabStripMonde_Click()
    Dim i As Integer
    If Not ComReseau.Definir_Actif Then
        If TabStripMonde.Tabs("CreerServeur").Selected Or _
           TabStripMonde.Tabs("RejoindrePartie").Selected Then
            TabStripMonde.Tabs("Campagnes").Selected = True
        End If
    End If
    For i = 0 To TabStripMonde.Tabs.Count - 1
        If TabStripMonde.Tabs(i + 1).Selected Then
            Exit For
        End If
    Next i
    FrameCampagnes.Visible = TabStripMonde.Tabs("Campagnes").Selected
    FrameMonde.Visible = TabStripMonde.Tabs("NouveauMonde").Selected Or TabStripMonde.Tabs("NouvelleBataille").Selected
    FrameFiefs.Visible = TabStripMonde.Tabs("NouveauMonde").Selected
    FrameArmees.Visible = TabStripMonde.Tabs("NouvelleBataille").Selected
    If TabStripMonde.Tabs("NouveauMonde").Selected Then
        OptionCarte(1).Enabled = True
    ElseIf TabStripMonde.Tabs("NouvelleBataille").Selected Then
        OptionCarte(0).Value = True
        OptionCarte_Click (0)
        OptionCarte(1).Enabled = False
    End If
    
    If TabStripMonde.Tabs("ChargerScenario").Selected Or TabStripMonde.Tabs("ChargerMonde").Selected Then
        FrameChargerMonde.Visible = True
        MDIFrmMain.MenuMondeRenommer.Enabled = True
        MDIFrmMain.MenuMondeCommentaires.Enabled = True
        MDIFrmMain.MenuMondeEditer.Enabled = True
        MDIFrmMain.MenuMondeSupprimer.Enabled = True
        MDIFrmMain.MenuMondeEditer.Visible = TabStripMonde.Tabs("ChargerScenario").Selected
        Actualiser_Liste_Mondes
    Else
        FrameChargerMonde.Visible = False
        'MDIFrmMain.MenuMonde.Enabled = False
        MDIFrmMain.MenuMondeRenommer.Enabled = False
        MDIFrmMain.MenuMondeCommentaires.Enabled = False
        MDIFrmMain.MenuMondeEditer.Enabled = False
        MDIFrmMain.MenuMondeSupprimer.Enabled = False
    End If
    FrameMode2Joueurs.Visible = TabStripMonde.Tabs("Mode2Joueurs").Selected
    FrameServeur.Visible = TabStripMonde.Tabs("CreerServeur").Selected
    FrameClient.Visible = TabStripMonde.Tabs("RejoindrePartie").Selected
    FrameReseau.Visible = TabStripMonde.Tabs("CreerServeur").Selected Or TabStripMonde.Tabs("RejoindrePartie").Selected
    FrameInfoReseau.Visible = TabStripMonde.Tabs("CreerServeur").Selected Or TabStripMonde.Tabs("RejoindrePartie").Selected
    
    If i > 0 And i < 5 Then
        DernierePartie = i
        Partie.IndiceMode = i
        If TabStripMonde.Tabs("ChargerScenario").Selected Or TabStripMonde.Tabs("ChargerMonde").Selected And LstMonde.ListIndex >= 0 Then
            Partie.FichierSauvegarde = LstMonde.List(LstMonde.ListIndex)
        Else
            Partie.FichierSauvegarde = ""
        End If
    End If
    
    'FrameServeur.Visible = False
    'FrameClient.Visible = False
    'FrameReseau.Visible = False
    'FrameFrameInfoReseau.Visible = False
    Actualiser_Bouton_Commencer
End Sub

Private Sub TabStripMonde2_Click()
    Dim i As Integer
    For i = 1 To TabStripMonde2.Tabs.Count
        If TabStripMonde2.Tabs(i).Selected Then
            TabStripMonde.Tabs(i).Selected = True
        End If
    Next i
    'TabStripMonde_Click
End Sub

Private Sub TabStripPersonnage_Click()
    FrameNouveauPersonnage.Visible = TabStripPersonnage.Tabs(1).Selected
    FrameChargerPersonnage.Visible = TabStripPersonnage.Tabs(2).Selected And FilePerso.ListCount > 0
    If TabStripPersonnage.Tabs(2).Selected Then
        Unload FrmInfoRaces
    End If
End Sub

Private Sub TimerApparence_Timer()
    Dim Temp As String
    Select Case CompteTimer ' Mod 4
    Case 2:
        ImgApparence.Left = ImgApparence.Left + 360
    Case 3:
        ImgApparence.Left = ImgApparence.Left + 360
        CompteTimer = -1
    Case Else
        ImgApparence.Left = ImgApparence.Left - 360
    End Select
    CompteTimer = CompteTimer + 1
    
    Musique.Jouer_Musique 0, True
    
    'Fait clignoter les messages en rouge.
    If CompteTimer = 0 Then
        LblErreurNom.Visible = Not LblErreurNom.Visible
        LblSelectionner2Personnages.Visible = Not LblSelectionner2Personnages.Visible
        If CmdCreerPerso.Enabled Then
            If CmdCreerPerso.BackColor = &H8000000F Then
                CmdCreerPerso.BackColor = &HFF00&
            Else
                CmdCreerPerso.BackColor = &H8000000F
            End If
        ElseIf CmdCreerPerso.BackColor = &HFF00& Then
            CmdCreerPerso.BackColor = &H8000000F
        End If
        'Fait clignoter le bouton "Jouer".
'        If CmdCommencer.Enabled Then
'            If CmdCommencer.BackColor = &H8000000F Then
'                CmdCommencer.BackColor = &HFF00&
'            Else
'                CmdCommencer.BackColor = &H8000000F
'            End If
'        ElseIf CmdCommencer.BackColor = &HFF00& Then
'            CmdCommencer.BackColor = &H8000000F
'        End If
        'Fait clignoter les boutons de connexion réseau.
        If TabStripMonde.Tabs("CreerServeur").Selected Then
            If Not ComReseau.Connecte Then
                If CmdCreerServeur.BackColor <> &H8000000F Then
                    CmdCreerServeur.BackColor = &H8000000F
                Else
                    CmdCreerServeur.BackColor = &HFF00&
                End If
            End If
        End If
        If TabStripMonde.Tabs("RejoindrePartie").Selected Then
            If LstParties.ListCount = 0 And Not ComReseau.Connecte Then
                If CmdRafraichirParties.BackColor <> &H8000000F Then
                    CmdRafraichirParties.BackColor = &H8000000F
                Else
                    CmdRafraichirParties.BackColor = &HFF00&
                End If
            ElseIf CmdConnecter.Enabled Then
                If CmdConnecter.BackColor <> &H8000000F Then
                    CmdConnecter.BackColor = &H8000000F
                Else
                    CmdConnecter.BackColor = &HFF00&
                End If
            End If
        End If
    End If
    
    'Raffraichit le réseau.
    If ComReseau.Connecte Then
        If ComReseau.Client Then
            ComMessages.Infos_Partie ComReseau.PlayerID
            CmdCommencer.Enabled = Partie.EnCours = 2
        End If
        ComReseau.Message_Lire True
        'If Temp <> "" Then
        '    LstMessages.AddItem Temp
        'End If
        Partie.IndiceMode = DernierePartie
        LblDetailPartie = Partie.Definir_Info_Partie
        CmdRafraichirJoueurs_Click
    Else
        LblDetailPartie = ""
    End If
End Sub

Private Sub Actualiser_Liste_Persos()
    Dim i As Integer
    FilePerso.Refresh
    LstPerso.Clear
    
    For i = 0 To FilePerso.ListCount - 1
        LstPerso.AddItem Left(FilePerso.List(i), InStr(FilePerso.List(i), ".") - 1)
    Next i
            
    If LstPerso.ListCount > 0 Then
        TabStripPersonnage.Tabs(2).Selected = True
        MDIFrmMain.MenuPersonnageExporter.Enabled = True
        MDIFrmMain.MenuPersonnageSelectionner.Enabled = True
        MDIFrmMain.MenuPersonnageSupprimer.Enabled = True
    Else
        TabStripPersonnage.Tabs(1).Selected = True
        MDIFrmMain.MenuPersonnageExporter.Enabled = False
        MDIFrmMain.MenuPersonnageSelectionner.Enabled = False
        MDIFrmMain.MenuPersonnageSupprimer.Enabled = False
    End If
    Actualiser_Bouton_Commencer
End Sub

Private Sub Actualiser_Joueurs()
    Dim i As Long
    Dim Temp1 As Integer, Temp2 As Integer
    Temp1 = ComboNomJoueur(0).ListIndex
    Temp2 = ComboNomJoueur(1).ListIndex
    
    ComboNomJoueur(0).Clear
    ComboNomJoueur(1).Clear
    For i = 0 To LstPerso.ListCount - 1
        'If LstPerso.Selected(i) Then
            ComboNomJoueur(0).AddItem Left(FilePerso.List(i), InStr(FilePerso.List(i), ".") - 1)
            ComboNomJoueur(1).AddItem Left(FilePerso.List(i), InStr(FilePerso.List(i), ".") - 1)
        'End If
    Next i
    If ComboNomJoueur(0).ListCount > 0 Then
        ComboNomJoueur(0).ListIndex = IIf(Temp1 > -1 And Temp1 < ComboNomJoueur(0).ListCount, Temp1, 0)
        ComboNomJoueur(1).ListIndex = IIf(Temp2 > -1 And Temp2 < ComboNomJoueur(1).ListCount, Temp2, 0)
        If ComboNomJoueur(0).ListCount > 1 Then
            If ComboNomJoueur(0).ListIndex = ComboNomJoueur(1).ListIndex Then
                If ComboNomJoueur(0).ListIndex + 1 < ComboNomJoueur(0).ListCount Then
                    ComboNomJoueur(1).ListIndex = ComboNomJoueur(0).ListIndex + 1
                ElseIf ComboNomJoueur(0).ListIndex - 1 < ComboNomJoueur(0).ListCount Then
                    ComboNomJoueur(1).ListIndex = ComboNomJoueur(0).ListIndex - 1
                End If
            End If
        End If
    End If
End Sub

Private Sub Actualiser_Infos_Perso()
    Dim i As Integer
    Dim j As Long
    Dim Niv As Double
'    Dim TempChargerPerso As Boolean
'    TempChargerPerso = ChargerPerso
'    ChargerPerso = False
    If ChargerPerso Then
    If LstPerso.ListCount > 0 Then
        If LstPerso.Selected(LstPerso.ListIndex) Then
            Persos(NoPerso).Init NoPerso, Parametres
            Charger_Sauvegarde_Perso LstPerso.List(LstPerso.ListIndex), Persos(NoPerso)
            If Persos(NoPerso).Protege Then
                ImgLocker(0).Visible = False
                ImgLocker(1).Visible = True
            Else
                ImgLocker(0).Visible = True
                ImgLocker(1).Visible = False
            End If
            LblIA.Visible = Persos(NoPerso).ChoisirIA
            LblNom = Persos(NoPerso).Nom
            'FramePersonnage = Persos(NoPerso).Nom
            LblNom = LblNom
            LblMetier = Persos(NoPerso).Metier(Parametres)
            'LblNom.ToolTipText = LblNom.Caption
            'LblMetier = Persos(NoPerso).Metier(Parametres)
            Niv = Persos(NoPerso).Niveau
            'LblNiveau = Persos(NoPerso).Metier(Parametres) & "Niveau " & Int(Niv)
            'LblNivPlus = Right((Format(Niv, "0.00") - Int(Niv)), 2) & " %"
            If Str(Niv) = Str(Int(Niv) + 1) Then
                LblNiveau = Int(Niv) + 1
                LblNivPlus = "0 %"
            Else
                LblNiveau = Int(Niv)
                LblNivPlus = Int((Format(Niv, "0.00") - Int(Niv)) * 100) & " %"
            End If
            ComboEquipePerso.ListIndex = Persos(NoPerso).NumeroFief
            If Persos(NoPerso).Feminin Then
                LblSexe = Parametres.EtiquetteFeminin
            Else
                LblSexe = Parametres.EtiquetteMasculin
            End If
            LblRace = Parametres.Race_Nom(Persos(NoPerso).Race, Persos(NoPerso).Feminin)
            Select Case MDIFrmMain.AffichageCompetences
            Case 0:
                LblVie = Persos(NoPerso).Definir_MaxVie  'MaxVie
                LblEnergie = Persos(NoPerso).Definir_MaxEnergie   'MaxEnergie
                LblMagie = Persos(NoPerso).Definir_MaxMagie  'MaxMagie
                LblMoral = Persos(NoPerso).Definir_MaxMoral  'MaxMoral
                LblAttaque = Persos(NoPerso).Definir_MaxAttaque 'MaxAttaque
                LblDefense = Persos(NoPerso).Definir_MaxDefense  'MaxDefense
            Case 1:
                LblVie = Persos(NoPerso).NivVie + 1
                LblEnergie = Persos(NoPerso).NivEnergie + 1
                LblMagie = Persos(NoPerso).NivMagie + 1
                LblMoral = Persos(NoPerso).NivMoral + 1
                LblAttaque = Persos(NoPerso).NivAttaque + 1
                LblDefense = Persos(NoPerso).NivDefense + 1
            Case 2:
                LblVie = Persos(NoPerso).ExpVie
                LblEnergie = Persos(NoPerso).ExpEnergie
                LblMagie = Persos(NoPerso).ExpMagie
                LblMoral = Persos(NoPerso).ExpMoral
                LblAttaque = Int(Persos(NoPerso).ExpAttaque / 20)
                LblDefense = Int(Persos(NoPerso).ExpDefense / 20)
            End Select
            
            TabStripCompBest_Click
            
            'Apparences.
            If Parametres.Race_CheminAbsolu(Persos(NoPerso).Race) Then
                FileApparence.Filename = App.Path & Parametres.Race_CheminApparence(Persos(NoPerso).Race, Persos(NoPerso).Feminin)
            Else
                FileApparence.Filename = App.Path & AffPerso.CheminPerso & Parametres.Race_CheminApparence(Persos(NoPerso).Race, Persos(NoPerso).Feminin)
            End If
            Actualiser_Infos_Apparences
            If Persos(NoPerso).FichierApparence <> "" Then
                For i = 0 To FileApparence.ListCount - 1
                    If FileApparence.List(i) = Persos(NoPerso).FichierApparence Then
                        FileApparence.ListIndex = i
                        Exit For
                    End If
                Next i
            End If
            If FileApparence.ListIndex < 0 Then
                FileApparence.ListIndex = 0
            End If
            FileApparence_Click
            
            'IAs
            FileIA.ListIndex = -1
            If Persos(NoPerso).FichierIA <> "" Then
                For i = 0 To FileIA.ListCount - 1
                    If FileIA.List(i) = Persos(NoPerso).FichierIA Then
                        FileIA.ListIndex = i
                        Exit For
                    End If
                Next i
            End If
            If FileIA.ListIndex < 0 Then
                FileIA.ListIndex = 0
            End If
            
            TabStripFichiers_Click
            
            'FileIAScript.Filename = App.Path & "\" & CheminIAScript
            'FileIAScript.ListIndex = -1
            If Persos(NoPerso).FichierIAScript <> "" Then
                For i = 0 To FileIAScript.ListCount - 1
                    If FileIAScript.List(i) = Persos(NoPerso).FichierIAScript Then
                        FileIAScript.ListIndex = i
                        Exit For
                    End If
                Next i
            End If
            If FileIAScript.ListIndex < 0 Or FileIAScript.ListIndex >= FileIAScript.ListCount Then
                FileIAScript.ListIndex = 0
            End If
        End If
    Else
        OptionApparence(0).Enabled = False
        OptionApparence(1).Enabled = False
    End If
    End If
'    ChargerPerso = TempChargerPerso
End Sub

Private Sub Actualiser_Infos_Apparences()
    Dim i As Integer
    CmdApparence.Enabled = Persos(NoPerso).Nombre_ApparenceBase > 0
    If Parametres.Race_DefinirApparence(Persos(NoPerso).Race) Then
        OptionApparence_Click 1
        OptionApparence(0).Enabled = False
        OptionApparence(1).Enabled = False
    Else
        If Persos(NoPerso).ChoisirApparence Then
            OptionApparence_Click 1
        Else
            OptionApparence_Click 0
        End If
        OptionApparence(0).Enabled = True
        OptionApparence(1).Enabled = True
    End If
End Sub

Private Sub Actualiser_Infos_Perso_Competences()
    'liste des compétences.
    Dim i As Integer, j As Integer
    Dim Temp As Boolean
    Dim Competence As ClsJeuCompetences
    Set Competence = New ClsJeuCompetences
    Dim TabComp As Collection
    Set TabComp = New Collection
    With MDIFrmMain.MenuPersonnageCompetencesAfficherToutes
    LstCompetences.Clear
    For i = 0 To Persos(NoPerso).Nombre_CompetencesSpeciales - 1
        If .Checked Or Persos(NoPerso).Niveau_CompetenceSpeciales(i) > 0 Then
            Competence.Nom = Parametres.Speciales_NomCompetences(i)
            Select Case MDIFrmMain.AffichageCompetences
            Case 0:
                Competence.valeur = Persos(NoPerso).Definir_Niveau_CompetenceSpeciales(i)
            Case 1:
                Competence.valeur = Persos(NoPerso).Niveau_CompetenceSpeciales(i) + 1
            Case 2:
                Competence.valeur = Persos(NoPerso).Definir_ExpSpeciales(i)
            End Select
            TabComp.Add Competence
            Set Competence = Nothing
            Set Competence = New ClsJeuCompetences
        End If
    Next i
    For i = 0 To Persos(NoPerso).Nombre_CompetencesRessources - 1
        If .Checked Or Persos(NoPerso).Niveau_CompetenceRessources(i) > 0 Then
            Competence.Nom = Parametres.Ressources_NomCompetence(i)
            Select Case MDIFrmMain.AffichageCompetences
            Case 0:
                Competence.valeur = Persos(NoPerso).Definir_Niveau_CompetenceRessources(i) + Parametres.PersosMaxRessources 'Persos(Noperso).Niveau_CompetenceRessources(i) + Parametres.PersosMaxRessources
            Case 1:
                Competence.valeur = Persos(NoPerso).Niveau_CompetenceRessources(i) + 1
            Case 2:
                Competence.valeur = Persos(NoPerso).Definir_ExpRessources(i)
            End Select
            TabComp.Add Competence
            Set Competence = Nothing
            Set Competence = New ClsJeuCompetences
        End If
    Next i
    For i = 0 To Persos(NoPerso).Nombre_CompetencesServices - 1
        If .Checked Or Persos(NoPerso).Niveau_CompetenceServices(i) > 0 Then
            Competence.Nom = Parametres.Service_NomCompetence(i + 1)
            Select Case MDIFrmMain.AffichageCompetences
            Case 0:
                Competence.valeur = Persos(NoPerso).Definir_Niveau_CompetenceServices(i) 'Persos(Noperso).Definir_Carac_CompetenceServices(i)
            Case 1:
                Competence.valeur = Persos(NoPerso).Niveau_CompetenceServices(i) + 1
            Case 2:
                Competence.valeur = Persos(NoPerso).Definir_ExpServices(i)
            End Select
            TabComp.Add Competence
            Set Competence = Nothing
            Set Competence = New ClsJeuCompetences
        End If
    Next i
    For i = 1 To Persos(NoPerso).Nombre_CompetencesObjets - 1
        j = Parametres.CompetenceObjet_NoListe(i)
        If .Checked Or Persos(NoPerso).Niveau_CompetenceObjets(j) > 0 Then
            Competence.Nom = Parametres.CompetenceObjet_Nom(j)
            Select Case MDIFrmMain.AffichageCompetences
            Case 0:
                Competence.valeur = Persos(NoPerso).Definir_Niveau_CompetenceObjets(j) 'Persos(Noperso).Definir_Carac_CompetenceObjets(i)
            Case 1:
                Competence.valeur = Persos(NoPerso).Niveau_CompetenceObjets(j) + 1
            Case 2:
                Competence.valeur = Persos(NoPerso).Definir_ExpObjets(j)
            End Select
            TabComp.Add Competence
            Set Competence = Nothing
            Set Competence = New ClsJeuCompetences
        End If
    Next i
    Select Case MDIFrmMain.TriCompetences
    Case 1:
        'On trie les compétences selon leur nom.
        Do
            Temp = True 'Si temp passe à faux, on fait un tri supplémentaire.
            For i = 2 To TabComp.Count
                If TabComp(i - 1).Nom > TabComp(i).Nom Then
                    Competence.valeur = TabComp(i).valeur
                    Competence.Nom = TabComp(i).Nom
                    TabComp(i).valeur = TabComp(i - 1).valeur
                    TabComp(i).Nom = TabComp(i - 1).Nom
                    TabComp(i - 1).valeur = Competence.valeur
                    TabComp(i - 1).Nom = Competence.Nom
                    Temp = False
                End If
            Next i
        Loop Until Temp
    Case 2:
        'On trie les compétences selon leur score.
        Do
            Temp = True 'Si temp passe à faux, on fait un tri supplémentaire.
            For i = 2 To TabComp.Count
                If TabComp(i - 1).valeur < TabComp(i).valeur Then
                    Competence.valeur = TabComp(i).valeur
                    Competence.Nom = TabComp(i).Nom
                    TabComp(i).valeur = TabComp(i - 1).valeur
                    TabComp(i).Nom = TabComp(i - 1).Nom
                    TabComp(i - 1).valeur = Competence.valeur
                    TabComp(i - 1).Nom = Competence.Nom
                    Temp = False
                End If
            Next i
        Loop Until Temp
    End Select
    For i = 1 To TabComp.Count
        LstCompetences.AddItem TabComp(i).Nom & " : " & TabComp(i).valeur
    Next i
    If LstCompetences.ListCount > 0 Then
        LstCompetences.AddItem "----------------------------------------------------"
        LstCompetences.AddItem UCase(Parametres.EtiquetteMinimum) & " : " & Persos(NoPerso).Definir_Min_Competence(MDIFrmMain.AffichageCompetences)
        LstCompetences.AddItem UCase(Parametres.EtiquetteMoyenne) & " : " & Format(Persos(NoPerso).Definir_Moy_Competence(MDIFrmMain.AffichageCompetences), "0.00")
        LstCompetences.AddItem UCase(Parametres.EtiquetteMaximum) & " : " & Persos(NoPerso).Definir_Max_Competence(MDIFrmMain.AffichageCompetences)
    End If
    End With
End Sub

Private Sub Actualiser_Infos_Perso_Bestiaire()
    'Liste du bestiaire.
    Dim i As Integer
    With Parametres
    LstBestiaire.Clear
    If Persos(NoPerso).Definir_Bestiaire_Total > 0 Then
        For i = 0 To Persos(NoPerso).Nombre_de_lignes_du_bestiaire - 1
            If Persos(NoPerso).Definir_Bestiaire(.Bestiaire_Ordre_Alpha(i)) > 0 Then
            'If Persos(Noperso).Definir_Bestiaire(i) > 0 Then
                'If i = 0 Then
                    'Nombre d'humains.
                '    LstBestiaire.AddItem "Autres personnages : " & Persos(Noperso).Definir_Bestiaire(i)
                'Else
                    'If Persos(Noperso).Definir_Bestiaire(i) > 1 Then
                        'LstBestiaire.AddItem Parametres.Race_Nom(i, , Persos(Noperso).Definir_Bestiaire(i) > 1) & " : " & Persos(Noperso).Definir_Bestiaire(i)
            LstBestiaire.AddItem Parametres.Race_Nom(.Bestiaire_Ordre_Alpha(i), , Persos(NoPerso).Definir_Bestiaire(.Bestiaire_Ordre_Alpha(i)) > 1) & " : " & Persos(NoPerso).Definir_Bestiaire(.Bestiaire_Ordre_Alpha(i))
            'LstBestiaire.AddItem Parametres.Race_Nom(i, , Persos(Noperso).Definir_Bestiaire(i) > 1) & " : " & Persos(Noperso).Definir_Bestiaire(i)
                    'Else
                    '    LstBestiaire.AddItem Parametres.Monstre_Nom(i - 1) & " : " & Persos(Noperso).Definir_Bestiaire(i)
                    'End If
                'End If
            End If
        Next i
        LstBestiaire.AddItem "----------------------------------------------------"
        LstBestiaire.AddItem UCase(Parametres.EtiquetteTotal) & " : " & Persos(NoPerso).Definir_Bestiaire_Total
    End If
    End With
End Sub

Private Sub Actualiser_Infos_Perso_Talents()
    'Liste des talents.
    Actualiser_Talents_Liste
    Actualiser_Talents_Score
End Sub

Public Sub Actualiser_Liste_Campagnes()
    Dim Temp As String
    
    LstCampagnes.Clear
    
    Temp = Dir(App.Path & CheminCampagnes & Langues.Dossier, vbDirectory)
    While Temp <> ""
        If InStr(Temp, ".") = 0 Then
            LstCampagnes.AddItem Temp
        End If
        Temp = Dir()
    Wend
    
    If LstCampagnes.ListCount > 0 Then
        LstCampagnes.ListIndex = 0
    End If
End Sub

Public Sub Actualiser_Liste_Mondes()
    Dim i As Long
   
    LblDescriptionMonde.Caption = ""
    ImgApercu.Picture = LoadPicture("")
    
    'Paramètres charger monde.
    If TabStripMonde.Tabs("ChargerScenario").Selected Then
        FileMonde.Filename = App.Path & CheminScenario & Langues.Dossier
        FileMonde.Pattern = "*" & ExtensionFichiersScenario
    Else
        FileMonde.Filename = App.Path & CheminSavMonde
        FileMonde.Pattern = "*" & ExtensionFichiersMonde
    End If
    
    FileMonde.Refresh
        
    LstMonde.Clear
    For i = 0 To FileMonde.ListCount - 1
        LstMonde.AddItem Left$(FileMonde.List(i), InStr(FileMonde.List(i), ".") - 1)
    Next i
    If FileMonde.ListCount > 0 Then
        'LstMonde.ListIndex = 0
        If TabStripMonde.Tabs("ChargerScenario").Selected Then
            If DernierScenario < LstMonde.ListCount Then
                LstMonde.ListIndex = DernierScenario
            End If
        Else
            If DernierMonde < LstMonde.ListCount Then
                LstMonde.ListIndex = DernierMonde
            End If
        End If
        If LstMonde.ListIndex < 0 Then
            LstMonde.ListIndex = 0
        End If
    Else
        LstMonde.ListIndex = -1
    End If
    
    Actualiser_Bouton_Commencer
End Sub

Public Sub Actualiser_Bouton_Commencer()
    Dim Actif As Boolean
    If (TabStripMonde.Tabs("CreerServeur").Selected Or TabStripMonde.Tabs("RejoindrePartie").Selected) Then
        Actif = ComReseau.Connecte
    Else
        If TabStripMonde.Tabs("Campagnes").Selected Then
            Actif = LstCampagnes.ListIndex >= 0 And LstCampagnesScenarios.ListIndex >= 0
        ElseIf TabStripMonde.Tabs("ChargerScenario").Selected Or TabStripMonde.Tabs("ChargerMonde").Selected Then
            Actif = FileMonde.ListCount > 0
        Else
            Actif = True
        End If
    End If
    Actif = Actif And FilePerso.ListCount > 0 'Il faut avoir sélectionné au moins un joueur.
    CmdCommencer.Enabled = Actif 'And Not TabStripMonde.Tabs("Mode2Joueurs").Selected 'And Not TabStripMonde.Tabs("Stastistiques").Selected
    MDIFrmMain.MenuPartieCommencer.Enabled = Actif
End Sub

Public Sub Actualiser_Boutons_Reseau()
    CmdCreerServeur.Enabled = Not ComReseau.Connecte
    CmdFermerServeur.Enabled = ComReseau.Serveur
    
    CmdRafraichirParties.Enabled = Not ComReseau.Connecte
    CmdConnecter.Enabled = Not ComReseau.Connecte And LstParties.ListIndex > -1
    CmdDeconnecter.Enabled = ComReseau.Connecte
    
    CmdRafraichirJoueurs.Enabled = ComReseau.Connecte
    CmdEnvoyerMessage.Enabled = ComReseau.Connecte
    
    If ComReseau.Connecte Then
        LblReseau.Visible = True
        If ComReseau.Serveur Then
            LblReseau.Caption = EtiquetteReseauServeur
        Else
            LblReseau.Caption = EtiquetteReseauClient
        End If
    Else
        LblReseau.Visible = False
        LblDetailPartie.Caption = ""
    End If
    
    TxtNomSession.Enabled = Not ComReseau.Connecte
    TxtNomJoueur.Enabled = Not ComReseau.Connecte
    'CheckOuverturePort.Enabled = Not ComReseau.Connecte
    TxtNumeroPort.Enabled = Not ComReseau.Connecte
    OptionJouerEnLigne(0).Enabled = Not ComReseau.Connecte
    OptionJouerEnLigne(1).Enabled = Not ComReseau.Connecte
    LstServices.Enabled = Not ComReseau.Connecte
    Actualiser_Bouton_Commencer
End Sub

Private Sub TxtMessage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'Entrée
        If TxtMessage.Text <> "" And CmdEnvoyerMessage.Enabled Then
            CmdEnvoyerMessage_Click
        End If
    End If
End Sub

Private Sub TxtNomJoueur_Change()
    ComReseau.Definir_NomJoueur = TxtNomJoueur
End Sub

Private Sub TxtNomPerso_DblClick()
    TxtNomPerso = Noms.Tirer_Nom_Aleatoire(OptionSexe(1).Value)
End Sub

Private Sub TxtNumeroIPServeur_Change()
    ComReseau.Definir_NumeroIPServeur = TxtNumeroIPServeur
End Sub

Private Sub TxtNumeroPort_Change()
    ComReseau.Definir_NumeroPort = TxtNumeroPort
End Sub

Private Sub TxtNomPerso_Change()
    Dim i As Long
    Dim Temp As Boolean
    LblErreurNom.Caption = ""
    If TxtNomPerso = "" Then
        LblErreurNom.Caption = ErreurNomPersonnage(0)
        CmdCreerPerso.Enabled = False
    ElseIf InStr(TxtNomPerso, "\") > 0 Or _
           InStr(TxtNomPerso, "/") > 0 Or _
           InStr(TxtNomPerso, ":") > 0 Or _
           InStr(TxtNomPerso, "*") > 0 Or _
           InStr(TxtNomPerso, "?") > 0 Or _
           InStr(TxtNomPerso, Chr(34)) > 0 Or _
           InStr(TxtNomPerso, "<") > 0 Or _
           InStr(TxtNomPerso, ">") > 0 Or _
           InStr(TxtNomPerso, "|") > 0 Or _
           InStr(TxtNomPerso, ".") > 0 Then
            LblErreurNom.Caption = ErreurNomPersonnage(1)
            CmdCreerPerso.Enabled = False
    Else
        Temp = True
        For i = 0 To LstPerso.ListCount
            If LCase(LstPerso.List(i)) = LCase(TxtNomPerso) Then
                'Le personnage éxiste déjà.
                Temp = False
                LblErreurNom.Caption = ErreurNomPersonnage(2)
                i = LstPerso.ListCount
            End If
        Next i
        CmdCreerPerso.Enabled = Temp
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, j As Long
    
    MDIFrmMain.Caption = App.Title & " - déchargement..."
    
    FicIni.fichier = FicIni.chemin & FichierINI
    
    FicIni.Section = "Campagnes"
    FicIni.Parametre("DerniereCampagne") = LstCampagnes.ListIndex
    FicIni.Parametre("DernierScenario") = LstCampagnesScenarios.ListIndex
    
    FicIni.Section = "Personnages"
    For i = 0 To TabStripFichiers.Tabs.Count - 1
        If TabStripFichiers.Tabs(i + 1).Selected Then
            FicIni.Parametre("FichiersSelection") = i
        End If
    Next i
    For i = 0 To TabStripCompBest.Tabs.Count - 1
        If TabStripCompBest.Tabs(i + 1).Selected Then
            FicIni.Parametre("CompetencesSelection") = i
        End If
    Next i
    'Affiche ou non toutes les compétences.
    FicIni.Parametre("CompetencesAfficherToutes") = -CDbl(MDIFrmMain.MenuPersonnageCompetencesAfficherToutes.Checked)
    
    FicIni.Parametre("CompetencesTri") = MDIFrmMain.TriCompetences
    FicIni.Parametre("CompetencesAffichage") = MDIFrmMain.AffichageCompetences
    'Retient la liste des personnages.
    FicIni.Parametre("NombrePersonnages") = LstPerso.ListCount
    For i = 0 To LstPerso.ListCount - 1
        FicIni.Parametre("Personnage" & i) = IIf(LstPerso.Selected(i), "1", "0")
    Next i
    If LstPerso.ListIndex >= 0 Then
        FicIni.Parametre("DernierPersonnage") = LstPerso.ListIndex
    End If
    FicIni.Parametre("PersonnageJoueur1") = ComboNomJoueur(0).ListIndex
    FicIni.Parametre("PersonnageJoueur2") = ComboNomJoueur(1).ListIndex
    
    FicIni.Section = "Monde"
    For i = 0 To TabStripMonde.Tabs.Count - 1
        If TabStripMonde.Tabs(i + 1).Selected Then
            FicIni.Parametre("TypePartie") = i
            Exit For
        End If
    Next i
    'Suavegarde les données des cartes.
    If OptionCarte(0).Value Then
        FicIni.Parametre("TypeCarte") = 1
    Else
        FicIni.Parametre("TypeCarte") = 2
    End If
    FicIni.Parametre("DerniereCarte") = LstCartes.ListIndex + 1
    FicIni.Parametre("DernierePartie") = DernierePartie
    FicIni.Parametre("DernierScenario") = DernierScenario
    FicIni.Parametre("DernierMonde") = DernierMonde
    
    FicIni.Parametre("TailleChoisie") = ComboTaille.ListIndex + 1
    FicIni.Parametre("TerrainChoisi") = ComboTerrain.ListIndex + 1
    FicIni.Parametre("RessourceChoisie") = ComboRessources.ListIndex + 1
    FicIni.Parametre("EpoqueChoisie") = ComboEpoque.ListIndex + 1
    FicIni.Parametre("RessourcesDepartsChoisie") = ComboRessourcesDepart.ListIndex + 1
    FicIni.Parametre("VitesseEpoqueChoisie") = ComboVitesseEpoque.ListIndex + 1
    
    FicIni.Parametre("ValLargeur") = SliderLargeurCarte.Value
    FicIni.Parametre("ValHauteur") = SliderHauteurCarte.Value
    FicIni.Parametre("ValQuantitéRessources") = SliderNombreArbres.Value
    
    FicIni.Parametre("DifficulteChoisie") = ComboDifficulte.ListIndex + 1
    FicIni.Section = "Resurrections_Nombre"
    FicIni.Parametre("ResurrectionChoisie") = ComboResurrections(0).ListIndex + 1
    FicIni.Section = "Resurrections_Vitesse"
    FicIni.Parametre("ResurrectionChoisie") = ComboResurrections(1).ListIndex + 1
    
    FicIni.Section = "Fiefs"
    FicIni.Parametre("ValFiefs") = SliderFiefs.Value
    FicIni.Parametre("EquilibreFiefs") = CheckEquilibreFiefs.Value
    
    For i = 0 To SliderFiefs.Max - 1
        FicIni.Parametre("ValVillageoisFief" & i + 1) = SliderVillageois(i).Value
        FicIni.Parametre("PeupleFief" & i + 1) = ComboPeuple(i).ListIndex + 1
        FicIni.Parametre("EquipeFief" & i + 1) = ComboEquipe(i).ListIndex + 1
    Next i
    
    'Paramètres des armées.
    For i = 0 To FrameArmee.Count - 1
        FicIni.Section = "Armée_" & i + 1
        FicIni.Parametre("MinSoldats") = SliderNombreSoldats(i).Min
        FicIni.Parametre("ValSoldats") = SliderNombreSoldats(i).Value
        FicIni.Parametre("MaxSoldats") = SliderNombreSoldats(i).Max
        FicIni.Parametre("Peuple") = ComboArmee(i).ListIndex + 1
        FicIni.Parametre("Equipement") = CheckEquipementArmee(i).Value
        If CheckEquipementArmee(i).Value = 1 Then
            For j = 0 To LstObjets(i).ListCount - 1
                FicIni.Parametre("Objet" & j) = -CDbl(LstObjets(i).Selected(j))
            Next j
        End If
    Next i
End Sub

Private Sub TxtNomSession_Change()
    ComReseau.Definir_NomSession = TxtNomSession
End Sub

Private Sub LstTalents_Click()
    'Calcul les points de talents et le nombre de telents restants.
    Dim ScoreTalents As Integer
    Dim i As Integer, j As Integer
    Dim Talents() As Boolean
    ScoreTalents = Int(Persos(NoPerso).Niveau)
    
    If Parametres.Race_NombreObjets(Persos(NoPerso).Race) > 0 Then
        ReDim Talents(Parametres.Race_NombreObjets(Persos(NoPerso).Race) - 1)
        For i = 0 To Parametres.Race_NombreObjets(Persos(NoPerso).Race) - 1
            j = Parametres.Race_Objet(Persos(NoPerso).Race, i)
            Talents(i) = LstTalents.Selected(i)
        Next i
        If EnregistrerPerso Then
            Enregistrer_Sauvegarde_Tresor Persos(NoPerso), Talents()
        End If
        Actualiser_Talents_Score
    End If
    Persos(NoPerso).Objet_Equipes_Vider_Tout
    Persos(NoPerso).Objet_Inventaire_Vider_Tout
    'Actualiser_Infos_Perso
End Sub

Private Sub Actualiser_Talents_Score()
    Dim i As Integer
    Dim TalentsRestant As Integer
    Dim MessageAfficher As Boolean 'Enregistre s'il l'on a déjà affiche un message ou pas.
    TalentsRestant = Persos(NoPerso).Defnir_Score_Tresor
    
    'Calcul le nombre de points restants.
    If Parametres.Race_NombreObjets(Persos(NoPerso).Race) > 0 Then
        For i = 0 To Parametres.Race_NombreObjets(Persos(NoPerso).Race) - 1
            If LstTalents.Selected(i) Then
                If TalentsRestant < Parametres.Objet_PrixTalents(Parametres.Race_Objet(Persos(NoPerso).Race, i)) Then
                    If Not MessageAfficher Then
                        MsgBox Parametres.MessageFortuneManquante & Parametres.Objet_Nom(Parametres.Race_Objet(Persos(NoPerso).Race, i)), vbExclamation, Persos(NoPerso).Nom
                        MessageAfficher = True
                    End If
                    LstTalents.Selected(i) = False
                Else
                    TalentsRestant = TalentsRestant - _
                                     Parametres.Objet_PrixTalents(Parametres.Race_Objet(Persos(NoPerso).Race, i))
                End If
            End If
        Next i
    End If
    
     'Affiche les points de talents en en-tête.
    LstPointsTalents.Clear
    LstPointsTalents.AddItem Parametres.EtiquetteFortuneTotale & " : " & Persos(NoPerso).Defnir_Score_Tresor
    LstPointsTalents.AddItem Parametres.EtiquetteFortuneRestante & " : " & TalentsRestant
End Sub

Private Sub Actualiser_Talents_Liste()
    Dim i As Integer
    Dim Talents() As Boolean
    LstTalents.Clear
    Charger_Tableau_Tresor Persos(NoPerso), Talents
    If Parametres.Race_NombreObjets(Persos(NoPerso).Race) > 0 Then
        For i = 0 To Parametres.Race_NombreObjets(Persos(NoPerso).Race) - 1
            LstTalents.AddItem Parametres.Objet_Nom(Parametres.Race_Objet(Persos(NoPerso).Race, i)) & _
                               IIf(Parametres.Objet_PrixTalents(Parametres.Race_Objet(Persos(NoPerso).Race, i)) <= 0, "", _
                               " (" & Parametres.Objet_PrixTalents(Parametres.Race_Objet(Persos(NoPerso).Race, i)) & ")")
        Next i
        For i = 0 To Parametres.Race_NombreObjets(Persos(NoPerso).Race) - 1
            LstTalents.Selected(i) = Talents(i)
        Next i
    End If
End Sub

Public Property Get BonusXP() As Long
    BonusXP = DifficulteBonusXP(ComboDifficulte.ListIndex)
End Property
Public Property Get CoefVieChateaux() As Single
    CoefVieChateaux = DifficulteCoefVieChateaux(ComboDifficulte.ListIndex)
End Property
Public Property Get Definir_Resurrections_Nombre() As Long
    Definir_Resurrections_Nombre = ResurrectionsValeur(ComboResurrections(0).ListIndex)
End Property
Public Property Get Definir_Resurrections_Vitesse() As Single
    Definir_Resurrections_Vitesse = ResurrectionsCoefs(ComboResurrections(1).ListIndex)
End Property
