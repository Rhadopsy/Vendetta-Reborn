VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsJeuCartes"
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

Const FichierINI = "Cartes"
Const SectionINI = "Cartes"
Const SectionINI2 = "Terrain_"
Const SectionINI3 = "Ressource_"

Public Chemin As String

Private ExtentionImage As String
Private ExtentionTexte As String

Private NomImageApercu As String
Private NomImageTerrain As String
Private NomImageRessources  As String
Private NomImageChateaux  As String
Public NomImageCarte As String
Private NomFichierDescription  As String
Private NomFichierInformations As String
Private CreerFichierInformations As Boolean

Private NombreTypesTerrains As Integer
Private TerrainsCouleur() As RGBValue
Private NombreTypesRessources As Integer
Private RessourcesCouleur() As RGBValue

Private NombreCartes As Integer
Private NomCartes() As String 'Mémorise la liste des cartes.

Public DerniereCarte As String 'Mémorise la dernière carte sélectionnée.
Public Chargee As Boolean 'Vrai, on joue avec une carte chargée.

Private Type RGBValue
    Red As Byte
    Green As Byte
    Blue As Byte
End Type
Private Type FILEHEADER 'Infos Sur le Fichier
   ImgType As Integer
    DiskSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    Offset As Long
End Type
Private Type INFOHEADER 'Info sur l'image
    InfoSize As Long
    Width As Long
    Height As Long
    NbPlanes As Integer
    BitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Dim h1 As FILEHEADER
Dim h2 As INFOHEADER

Private Sub Class_Initialize()
    Dim i As Integer
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    Chemin = FicIni.Parametre("Chemin")

    ExtentionImage = FicIni.Parametre("ExtentionImage")
    ExtentionTexte = FicIni.Parametre("ExtentionTexte")

    NomImageApercu = FicIni.Parametre("NomImageApercu")
    NomImageTerrain = FicIni.Parametre("NomImageTerrain")
    NomImageRessources = FicIni.Parametre("NomImageRessources")
    NomImageChateaux = FicIni.Parametre("NomImageChateaux")
    NomImageCarte = FicIni.Parametre("NomImageCarte")
    NomFichierDescription = FicIni.Parametre("NomFichierDescription")
    NomFichierInformations = FicIni.Parametre("NomFichierInformations")
    CreerFichierInformations = FicIni.Parametre("CreerFichierInformations")

    NombreTypesTerrains = Val(FicIni.Parametre("NombreTypesTerrains"))
    ReDim TerrainsCouleur(NombreTypesTerrains - 1)
    For i = 0 To NombreTypesTerrains - 1
        FicIni.Section = SectionINI2 & i + 1
        TerrainsCouleur(i).Red = Val(FicIni.Parametre("R"))
        TerrainsCouleur(i).Green = Val(FicIni.Parametre("V"))
        TerrainsCouleur(i).Blue = Val(FicIni.Parametre("B"))
    Next i
    
    NombreTypesRessources = Val(FicIni.Parametre("NombreTypesRessources"))
    ReDim RessourcesCouleur(NombreTypesRessources - 1)
    For i = 0 To NombreTypesRessources - 1
        FicIni.Section = SectionINI3 & i + 1
        RessourcesCouleur(i).Red = Val(FicIni.Parametre("R"))
        RessourcesCouleur(i).Green = Val(FicIni.Parametre("V"))
        RessourcesCouleur(i).Blue = Val(FicIni.Parametre("B"))
    Next i
    Charger_Cartes
End Sub

Private Sub Charger_Cartes()
    'Charge toutes les cartes en mémoire.
    Dim Temp As String
    NombreCartes = 0
    ReDim NomCartes(0)
    Temp = Dir(App.Path & Chemin, vbDirectory)
    While Temp <> ""
        If InStr(Temp, ".") = 0 Then
            ReDim Preserve NomCartes(NombreCartes)
            NomCartes(NombreCartes) = Temp
            NombreCartes = NombreCartes + 1
        End If
        Temp = Dir()
    Wend
End Sub

Public Sub Charger_Dans_Liste(ByRef Liste As ListBox)
    'Charge la liste des images dans une listbox.
    Dim i  As Integer
    Liste.Clear
    For i = 0 To NombreCartes - 1
        Liste.AddItem NomCartes(i)
    Next i
End Sub

Public Sub Charger_Apercu(ByVal Carte As String, ByRef Apercu As Image)
    On Error GoTo Erreur1
    DerniereCarte = Carte
    'Charge la carte entrée en paramètre dans une image.
    Apercu.Picture = LoadPicture(App.Path & Chemin & Carte & "\" & NomImageApercu & ExtentionImage)
    Redimensionner_Apercu Apercu
    Exit Sub
Erreur1:
    On Error GoTo Erreur2
    Apercu.Picture = LoadPicture(App.Path & Chemin & Carte & "\" & NomImageTerrain & ExtentionImage)
    Redimensionner_Apercu Apercu
    Exit Sub
Erreur2:
End Sub

Private Sub Redimensionner_Apercu(ByRef Apercu As Image)
    Dim H As Long 'Hauteur de l'image.
    Dim l As Long 'Largeur de l'image.
    H = Definir_Hauteur
    l = Definir_Largeur
    If l > H Then
        Apercu.Height = 1815 * H / l
        Apercu.Width = 1815
        Apercu.Top = 280 + ((1815 - (1815 * H / l)) / 2)
        Apercu.Left = 2040
    ElseIf H > l Then
        Apercu.Height = 1815
        Apercu.Width = 1815 * l / H
        Apercu.Top = 280
        Apercu.Left = 2040 + ((1815 - (1815 * l / H)) / 2)
    Else
        Apercu.Height = 1815
        Apercu.Width = 1815
        Apercu.Top = 280
        Apercu.Left = 2040
    End If
End Sub

Public Sub Charger_Description(ByVal Carte As String, ByRef Description As Label)
    'Charge la description de la carte entrée en paramètre dans un label.
    On Error GoTo Erreur1
    Dim Temp As String
    Dim Fichier As String
    Description.Caption = ""
    Fichier = App.Path & Chemin & Carte & "\" & NomFichierDescription & Langues.Suffixe & ExtentionTexte
    Open Fichier For Input As #1
    Do
        Line Input #1, Temp
        Description.Caption = Description.Caption & Chr(10) & Chr(13)
        Description.Caption = Description.Caption & Temp
    Loop Until EOF(1)
    Close #1
    Exit Sub
Erreur1:
    Charger_Description2 Carte, Description
End Sub

Public Sub Charger_Description2(ByVal Carte As String, ByRef Description As Label)
    Dim Temp As String
    Dim Fichier As String
    On Error GoTo Erreur
    'Cherche un fichier description sans le suffixe de langue.
    Fichier = App.Path & Chemin & Carte & "\" & NomFichierDescription & ExtentionTexte
    Open Fichier For Input As #1
    Do
        Line Input #1, Temp
        Description.Caption = Description.Caption & Chr(10) & Chr(13)
        Description.Caption = Description.Caption & Temp
    Loop Until EOF(1)
    Close #1
    Exit Sub
Erreur: 'Aucun fichier description trouvé.
End Sub

Public Sub Charger_Monde(ByRef Monde As ClsJeuMonde)
    On Error GoTo Erreur
    Dim StrFileBuffer
    Dim Couleur As RGBValue: Dim Temp As String * 1: Dim Temp2 As Integer
    Dim PaletteCouleur(255) As RGBValue
    Dim r As Byte: Dim g As Byte: Dim B As Byte: Dim a As Byte 'Variables temporaires des couleurs chargées à partir du fichier.
    Dim i As Long
    Dim X As Long: Dim Y As Long
    Dim Largeur As Long: Dim Hauteur As Long
    Open App.Path & Chemin & DerniereCarte & "\" & NomImageTerrain & ExtentionImage For Binary Access Read As #1
    Get #1, , h1
    Get #1, , h2
    Largeur = h2.Width  'en Black & White
    Hauteur = h2.Height  'taille de l'image sortante
    If h2.Width Mod 4 <> 0 Then GoTo ErreurDimension 'Marche po dans ce cas là, g po compris pourquoi
    If h2.biCompression <> 0 Then GoTo ErreurCompression

    Monde.Largeur = Largeur
    Monde.Hauteur = Hauteur
    Monde.Calculer_Diagonale
    Monde.NombreCases = (Largeur + 1) * (Hauteur + 1)
    
    If h2.BitCount = 24 Then 'Si c une image 16Millions de couleurs, il n'y a po de palette
        Get #1, (Val(h1.Offset)), Temp
        StrFileBuffer = Input(LOF(1), #1)  'ICI il n'y a po de paleete: l'image est enregistrée en BGR et toujours à l'envers
        Get #1, (Val(h1.Offset)), Temp
        X = 0: Y = h2.Height - 1
        'i = Monde.Largeur * Monde.Hauteur - 1
        For i = 1 To (Len(StrFileBuffer) / 3)
            Get #1, , B:  Get #1, , g: Get #1, , r
            Couleur.Red = r
            Couleur.Green = g
            Couleur.Blue = B
            Monde.Param_TerrainCategorie(Y * Monde.Largeur + X) = Type_Terrain(Couleur.Red, Couleur.Green, Couleur.Blue)
            X = X + 1
            If X Mod h2.Width = 0 Then
                Y = Y - 1 'si "x vaut la largeur" de l'image on diminue y
                X = 0
                If Y < 0 Then Exit For
            End If
        Next i
    Else
        For i = 0 To 255
            Get #1, , B: Get #1, , g 'Palette codée en BGR + RESERVED
            Get #1, , r: Get #1, , a
            PaletteCouleur(i).Red = r 'on indexe la couleur coreespond au i ème élément de la palette
            PaletteCouleur(i).Green = g
            PaletteCouleur(i).Blue = B
        Next i
        X = 0: Y = h2.Height - 1 'l'image est enregistrée à l'envers: de bas en haut et de droite à gauche
        Get #1, (Val(h1.Offset)), Temp 'on met le pointeur du fichier au début de la définition de l'image
        StrFileBuffer = Input(LOF(1), #1) 'on met dans le buffer l'image
        Get #1, (Val(h1.Offset) - 1), Temp 'on replace le pointeur au début de l'image
        
        'i = Monde.Largeur * Monde.Hauteur - 1
        For i = 1 To Len(StrFileBuffer)
            Temp2 = Asc(Mid(StrFileBuffer, i, 1))   'on prend le xème caractère de l'image
            'HEUREUSEMENT LA PALLETE EST CLASsée par ordre CROISSANT
            Couleur.Red = PaletteCouleur(Temp2).Red
            Couleur.Green = PaletteCouleur(Temp2).Green
            Couleur.Blue = PaletteCouleur(Temp2).Blue
            Monde.Param_TerrainCategorie(Y * Monde.Largeur + X) = Type_Terrain(Couleur.Red, Couleur.Green, Couleur.Blue)
            X = X + 1
            If X Mod h2.Width = 0 Then
                Y = Y - 1 'si "x vaut la largeur" de l'image on diminue y
                X = 0
                If Y < 0 Then Exit For
            End If
        Next i
    End If
    Close #1
    Monde.Reformer_Terrain
    Exit Sub
ErreurDimension:
    MsgBox "Les dimensions de cette carte sont invalides." & Chr(10) & Chr(13) & "Les dimensions doivent être des multiples de 4.", vbCritical, "Erreur"
    End
ErreurCompression:
    MsgBox "L'image de la carte est compressée." & Chr(10) & Chr(13) & "Impossible de la lire.", vbCritical, "Erreur"
    End
Erreur:
    If Definir_ImageChateaux_Existe Then
        MsgBox "Une erreur s'est produite avec le fichier" & Chr(10) & Chr(13) & _
        App.Path & Chemin & DerniereCarte & "\" & NomImageTerrain & ExtentionImage, vbCritical, "Erreur"
    Else
        MsgBox "Le fichier : " & _
        App.Path & Chemin & DerniereCarte & "\" & NomImageTerrain & ExtentionImage & " est manquant.", vbCritical, "Erreur"
    End If
    Editer_Repertoire App.Path & Chemin & DerniereCarte
End Sub

Public Sub Charger_Ressources(ByVal Monde As ClsJeuMonde, ByRef Ressources() As ClsJeuRessources)
    Dim StrFileBuffer
    Dim Couleur As RGBValue: Dim Temp As String * 1: Dim Temp2 As Integer
    Dim PaletteCouleur(255) As RGBValue
    Dim r As Byte: Dim g As Byte: Dim B As Byte: Dim a As Byte 'Variables temporaires des couleurs chargées à partir du fichier.
    Dim i As Long
    Dim X As Long: Dim Y As Long

    Dim NbRessources As Long 'Compte le nombre de ressources placées au fur et à mesure.
    Dim NbRessourcesFixe As Boolean 'Si vrai, on connaît le nombre de ressources exacte.
    
    Dim FicIniCarte As ClsFicINI
'    If NbRessources = 0 Then
    If Definir_ImageRessources_Existe Then
        Set FicIniCarte = New ClsFicINI
        FicIniCarte.Fichier = Chemin & DerniereCarte & "\" & NomFichierInformations
        FicIniCarte.Section = "Ressources"
        'Si le fichier informations.txt indique les ressources, on ne va pas lire le fichier image.
        If FicIniCarte.Parametre("NombreRessources") <> "" Then
            NbRessources = Val(FicIniCarte.Parametre("NombreRessources"))
            NbRessourcesFixe = True
            ReDim Ressources(NbRessources - 1)
            For i = 0 To NbRessources - 1
                Set Ressources(i) = New ClsJeuRessources
    '            If FicIniCarte.Parametre("Ressource(" & i + 1 & ")_PositionX") = "" Or FicIniCarte.Parametre("Ressource(" & i + 1 & ")_PositionY") = "" Then
    '                NbRessources = 0
    '                Exit For
    '            Else
    '                Ressources(i).PositionX = Val(FicIniCarte.Parametre("Ressource(" & i + 1 & ")_PositionX"))
    '                Ressources(i).PositionY = Val(FicIniCarte.Parametre("Ressource(" & i + 1 & ")_PositionY"))
    '            End If
            Next i
            NbRessources = 0
        End If
        Open App.Path & Chemin & DerniereCarte & "\" & NomImageRessources & ExtentionImage For Binary Access Read As #1
        Get #1, , h1
        Get #1, , h2
        If h2.Width Mod 4 <> 0 Then GoTo ErreurDimension 'Marche po dans ce cas là, g po compris pourquoi
        If h2.biCompression <> 0 Then GoTo ErreurCompression
        
        If h2.BitCount = 24 Then 'Si c une image 16Millions de couleurs, il n'y a po de palette
            Get #1, (Val(h1.Offset)), Temp
            StrFileBuffer = Input(LOF(1), #1)  'ICI il n'y a po de paleete: l'image est enregistrée en BGR et toujours à l'envers
            Get #1, (Val(h1.Offset)), Temp
            X = 0: Y = h2.Height
            For i = 1 To (Len(StrFileBuffer) / 3)
                Get #1, , B:  Get #1, , g: Get #1, , r
                Couleur.Red = r
                Couleur.Green = g
                Couleur.Blue = B
                
                If Couleur.Red <> 0 Or Couleur.Green <> 0 Or Couleur.Blue <> 0 Then
                    If Not NbRessourcesFixe Then
                        ReDim Preserve Ressources(NbRessources)
                        Set Ressources(NbRessources) = New ClsJeuRessources
                    End If
                    Ressources(NbRessources).PositionX = X * Monde.LargeurCase '- AffRessources.Largeur
                    Ressources(NbRessources).PositionY = Y * Monde.HauteurCase - AffRessources.Hauteur
                    Ressources(NbRessources).Definir_Type = Type_Ressource(Couleur.Red, Couleur.Green, Couleur.Blue)
                    NbRessources = NbRessources + 1
                End If
                
                X = X + 1
                If X Mod h2.Width = 0 Then
                    Y = Y - 1
                    X = 0
                End If
            Next i
        Else
            For i = 0 To 255
                Get #1, , B: Get #1, , g 'Palette codée en BGR + RESERVED
                Get #1, , r: Get #1, , a
                PaletteCouleur(i).Red = r 'on indexe la couleur coreespond au i ème élément de la palette
                PaletteCouleur(i).Green = g
                PaletteCouleur(i).Blue = B
            Next i
            X = 0 'l'image est enregistrée à l'envers: de bas en haut et de droite à gauche
            Y = h2.Height
            Get #1, (Val(h1.Offset)), Temp 'on met le pointeur du fichier au début de la définition de l'image
            StrFileBuffer = Input(LOF(1), #1) 'on met dans le buffer l'image
            Get #1, (Val(h1.Offset) - 1), Temp 'on replace le pointeur au début de l'image
            
            For i = 1 To Len(StrFileBuffer)
                Temp2 = Asc(Mid(StrFileBuffer, i, 1)) 'on prend le xème caractère de l'image
                'HEUREUSEMENT LA PALLETE EST CLASsée par ordre CROISSANT
                Couleur.Red = PaletteCouleur(Temp2).Red
                Couleur.Green = PaletteCouleur(Temp2).Green
                Couleur.Blue = PaletteCouleur(Temp2).Blue
                If Couleur.Red <> 0 Or Couleur.Green <> 0 Or Couleur.Blue <> 0 Then
                    If Not NbRessourcesFixe Then
                        ReDim Preserve Ressources(NbRessources)
                        Set Ressources(NbRessources) = New ClsJeuRessources
                    End If
                    Ressources(NbRessources).PositionX = X * Monde.LargeurCase '- AffRessources.Largeur
                    Ressources(NbRessources).PositionY = Y * Monde.HauteurCase - AffRessources.Hauteur
                    Ressources(NbRessources).Definir_Type = Type_Ressource(Couleur.Red, Couleur.Green, Couleur.Blue)
                    NbRessources = NbRessources + 1
                End If
                X = X + 1
                If X Mod h2.Width = 0 Then
                    Y = Y - 1 'si "x vaut la largeur" de l'image on diminue y
                    X = 0
                End If
            Next i
        End If
        Close #1
        If Not NbRessourcesFixe Then
            If CreerFichierInformations Then
                FicIniCarte.Parametre("NombreRessources") = NbRessources
            End If
        End If
'            For i = 0 To NbRessources - 1
'                FicIniCarte.Parametre("Ressource(" & i + 1 & ")_PositionX") = Ressources(i).PositionX
'                FicIniCarte.Parametre("Ressource(" & i + 1 & ")_PositionY") = Ressources(i).PositionY
'            Next i
    Else
        'S'il n'y a pas de fichier bitmap pour les ressources, on les tire au hasard.
        'ReDim Ressources(CLng(Monde.Largeur * Monde.Hauteur / 328))
        ReDim Ressources(CLng(Monde.NombreCasesConstructibles / 300))
        For i = 0 To UBound(Ressources())
            Set Ressources(i) = New ClsJeuRessources
        Next i
        Placer_Ressources
    End If
'    End If
    Exit Sub
ErreurDimension:
    MsgBox "Les dimensions de cette carte sont invalides." & Chr(10) & Chr(13) & "Les dimensions doivent être des multiples de 4.", vbCritical, "Erreur"
    End
ErreurCompression:
    MsgBox "L'image de la carte est compressée." & Chr(10) & Chr(13) & "Impossible de la lire.", vbCritical, "Erreur"
    End
End Sub

Public Sub Charger_Chateaux(ByVal Monde As ClsJeuMonde, ByRef Chateaux() As ClsJeuBatiment)
    On Error GoTo Erreur
    Dim StrFileBuffer
    Dim Couleur As RGBValue: Dim Temp As String * 1: Dim Temp2 As Integer
    Dim PaletteCouleur(255) As RGBValue
    Dim r As Byte: Dim g As Byte: Dim B As Byte: Dim a As Byte 'Variables temporaires des couleurs chargées à partir du fichier.
    Dim i As Long
    Dim X As Long: Dim Y As Long

    Dim NbChateaux As Long 'Compte le nombre de chateaux placés au fur et à mesure.
    Dim FicIniCarte As ClsFicINI
    Set FicIniCarte = New ClsFicINI
    FicIniCarte.Fichier = Chemin & DerniereCarte & "\" & NomFichierInformations
    FicIniCarte.Section = "Chateaux"
    If Definir_FichierInformations_Existe Then
        'Si le fichier informations.txt indique le nombre de chateaux existants, on ne va pas lire le fichier image.
        NbChateaux = Val(FicIniCarte.Parametre("NombreChateaux"))
        ReDim Chateaux(NbChateaux - 1)
        For i = 0 To NbChateaux - 1
            Set Chateaux(i) = New ClsJeuBatiment
            If FicIniCarte.Parametre("Chateau(" & i + 1 & ")_PositionX") = "" Or FicIniCarte.Parametre("Chateau(" & i + 1 & ")_PositionY") = "" Then
                NbChateaux = 0
                Exit For
            Else
                Chateaux(i).PositionX = Val(FicIniCarte.Parametre("Chateau(" & i + 1 & ")_PositionX"))
                Chateaux(i).PositionY = Val(FicIniCarte.Parametre("Chateau(" & i + 1 & ")_PositionY"))
            End If
        Next i
    End If
    If NbChateaux = 0 Then
        Open App.Path & Chemin & DerniereCarte & "\" & NomImageChateaux & ExtentionImage For Binary Access Read As #1
        Get #1, , h1
        Get #1, , h2
        If h2.Width Mod 4 <> 0 Then GoTo ErreurDimension 'Marche po dans ce cas là, g po compris pourquoi
        If h2.biCompression <> 0 Then GoTo ErreurCompression
        
        If h2.BitCount = 24 Then 'Si c une image 16Millions de couleurs, il n'y a po de palette
            Get #1, (Val(h1.Offset)), Temp
            StrFileBuffer = Input(LOF(1), #1)  'ICI il n'y a po de paleete: l'image est enregistrée en BGR et toujours à l'envers
            Get #1, (Val(h1.Offset)), Temp
            X = 0: Y = h2.Height
            For i = 1 To (Len(StrFileBuffer) / 3)
                Get #1, , B:  Get #1, , g: Get #1, , r
                Couleur.Red = r
                Couleur.Green = g
                Couleur.Blue = B
                If Couleur.Red = 255 And Couleur.Green = 255 And Couleur.Blue = 255 Then
                    ReDim Preserve Chateaux(NbChateaux)
                    Set Chateaux(NbChateaux) = New ClsJeuBatiment
                    Chateaux(NbChateaux).PositionX = X * Monde.LargeurCase - Parametres.Peuples_ChateauLargeur(0) / 2
                    Chateaux(NbChateaux).PositionY = Y * Monde.HauteurCase - Parametres.Peuples_ChateauHauteur(0) / 2
                    NbChateaux = NbChateaux + 1
                End If
                X = X + 1
                If X Mod h2.Width = 0 Then
                    Y = Y - 1
                    X = 0
                End If
            Next i
        Else
            For i = 0 To 255
                Get #1, , B: Get #1, , g 'Palette codée en BGR + RESERVED
                Get #1, , r: Get #1, , a
                PaletteCouleur(i).Red = r 'on indexe la couleur coreespond au i ème élément de la palette
                PaletteCouleur(i).Green = g
                PaletteCouleur(i).Blue = B
            Next i
            X = 0 'l'image est enregistrée à l'envers: de bas en haut et de droite à gauche
            Y = h2.Height
            Get #1, (Val(h1.Offset)), Temp 'on met le pointeur du fichier au début de la définition de l'image
            StrFileBuffer = Input(LOF(1), #1) 'on met dans le buffer l'image
            Get #1, (Val(h1.Offset) - 1), Temp 'on replace le pointeur au début de l'image
            
            For i = 1 To Len(StrFileBuffer)
                Temp2 = Asc(Mid(StrFileBuffer, i, 1)) 'on prend le xème caractère de l'image
                'HEUREUSEMENT LA PALLETE EST CLASsée par ordre CROISSANT
                Couleur.Red = PaletteCouleur(Temp2).Red
                Couleur.Green = PaletteCouleur(Temp2).Green
                Couleur.Blue = PaletteCouleur(Temp2).Blue
                If Couleur.Red = 255 And Couleur.Green = 255 And Couleur.Blue = 255 Then
                    ReDim Preserve Chateaux(NbChateaux)
                    Set Chateaux(NbChateaux) = New ClsJeuBatiment
                    Chateaux(NbChateaux).PositionX = X * Monde.LargeurCase - Parametres.Peuples_ChateauLargeur(0) / 2
                    Chateaux(NbChateaux).PositionY = Y * Monde.HauteurCase - Parametres.Peuples_ChateauHauteur(0) / 2
                    NbChateaux = NbChateaux + 1
                End If
                X = X + 1
                If X Mod h2.Width = 0 Then
                    Y = Y - 1 'si "x vaut la largeur" de l'image on diminue y
                    X = 0
                End If
            Next i
        End If
        Close #1
        If CreerFichierInformations Then
            For i = 0 To NbChateaux - 1
                FicIniCarte.Parametre("Chateau(" & i + 1 & ")_PositionX") = Chateaux(i).PositionX
                FicIniCarte.Parametre("Chateau(" & i + 1 & ")_PositionY") = Chateaux(i).PositionY
            Next i
        End If
    End If
    Echanger_Emplacement_Chateaux
    Exit Sub
ErreurDimension:
    MsgBox "Les dimensions de cette carte sont invalides." & Chr(10) & Chr(13) & "Les dimensions doivent être des multiples de 4.", vbCritical, "Erreur"
    End
ErreurCompression:
    MsgBox "L'image de la carte est compressée." & Chr(10) & Chr(13) & "Impossible de la lire.", vbCritical, "Erreur"
    End
Erreur:
    If Definir_ImageChateaux_Existe Then
        MsgBox "Une erreur s'est produite avec le fichier" & Chr(10) & Chr(13) & _
        App.Path & Chemin & DerniereCarte & "\" & NomImageChateaux & ExtentionImage, vbCritical, "Erreur"
    Else
        MsgBox "Le fichier : " & _
        App.Path & Chemin & DerniereCarte & "\" & NomImageChateaux & ExtentionImage & " est manquant.", vbCritical, "Erreur"
    End If
    Editer_Repertoire App.Path & Chemin & DerniereCarte
End Sub

Private Function Echanger_Emplacement_Chateaux()
    'Place le chateau du joueur de manière aléatoire sur la carte.
    Dim i As Integer
    Dim j As Integer
    Dim TempChateauxX As Collection
    Dim TempChateauxY As Collection
    Set TempChateauxX = New Collection
    Set TempChateauxY = New Collection
    For i = 0 To UBound(Chateaux())
        TempChateauxX.Add Chateaux(i).PositionX
        TempChateauxY.Add Chateaux(i).PositionY
    Next i
    For i = 0 To UBound(Chateaux())
        j = Int(Rnd * TempChateauxX.Count) + 1
        Chateaux(i).PositionX = TempChateauxX(j)
        Chateaux(i).PositionY = TempChateauxY(j)
        TempChateauxX.Remove j
        TempChateauxY.Remove j
    Next i
End Function

Public Property Get Definir_Largeur() As Long
    Open App.Path & Chemin & DerniereCarte & "\" & NomImageTerrain & ExtentionImage For Binary Access Read As #1
    Get #1, , h1
    Get #1, , h2
    Definir_Largeur = h2.Width
    Close #1
End Property

Public Property Get Definir_Hauteur() As Long
    Open App.Path & Chemin & DerniereCarte & "\" & NomImageTerrain & ExtentionImage For Binary Access Read As #1
    Get #1, , h1
    Get #1, , h2
    Definir_Hauteur = h2.Height
    Close #1
End Property

Public Property Get Definir_Nombre_Chateaux() As Integer
    On Error GoTo Erreur
    'Renvoie le nombre de chateaux dans la carte.
    Dim StrFileBuffer
    Dim Couleur As RGBValue: Dim Temp As String * 1
    Dim PaletteCouleur(255) As RGBValue
    Dim r As Byte: Dim g As Byte: Dim B As Byte: Dim a As Byte 'Variables temporaires des couleurs chargées à partir du fichier.
    Dim i As Long
    Dim X As Long: Dim Y As Long

    Dim NbChateaux As Long 'Compte le nombre de chateaux placés au fur et à mesure.
    
    Dim FicIniCarte As ClsFicINI
    Set FicIniCarte = New ClsFicINI
    FicIniCarte.Fichier = Chemin & DerniereCarte & "\" & NomFichierInformations
    FicIniCarte.Section = "Chateaux"
    If Definir_FichierInformations_Existe Then
        'Si le fichier informations.txt indique le nombre de chateaux existants, on ne va pas lire le fichier image.
        NbChateaux = Val(FicIniCarte.Parametre("NombreChateaux"))
    End If
    If NbChateaux = 0 Then
        Open App.Path & Chemin & DerniereCarte & "\" & NomImageChateaux & ExtentionImage For Binary Access Read As #1
        Get #1, , h1
        Get #1, , h2
        If h2.Width Mod 4 <> 0 Then GoTo ErreurDimension 'Marche po dans ce cas là, g po compris pourquoi
        If h2.biCompression <> 0 Then GoTo ErreurCompression
        
        If h2.BitCount = 24 Then 'Si c une image 16Millions de couleurs, il n'y a po de palette
            Get #1, (Val(h1.Offset)), Temp
            StrFileBuffer = Input(LOF(1), #1)  'ICI il n'y a po de paleete: l'image est enregistrée en BGR et toujours à l'envers
            Get #1, (Val(h1.Offset)), Temp
            For i = 1 To (Len(StrFileBuffer) / 3)
                Get #1, , B:  Get #1, , g: Get #1, , r
                Couleur.Red = r
                Couleur.Green = g
                Couleur.Blue = B
                If Couleur.Red = 255 And Couleur.Green = 255 And Couleur.Blue = 255 Then
                    NbChateaux = NbChateaux + 1
                End If
            Next i
        Else
            For i = 0 To 255
                Get #1, , B: Get #1, , g 'Palette codée en BGR + RESERVED
                Get #1, , r: Get #1, , a
                PaletteCouleur(i).Red = r 'on indexe la couleur coreespond au i ème élément de la palette
                PaletteCouleur(i).Green = g
                PaletteCouleur(i).Blue = B
            Next i
            Get #1, (Val(h1.Offset)), Temp 'on met le pointeur du fichier au début de la définition de l'image
            StrFileBuffer = Input(LOF(1), #1) 'on met dans le buffer l'image
            Get #1, (Val(h1.Offset) - 1), Temp 'on replace le pointeur au début de l'image
            
            For i = 1 To Len(StrFileBuffer)
                Temp = Asc(Mid(StrFileBuffer, i, 1)) 'on prend le xème caractère de l'image
                'HEUREUSEMENT LA PALLETE EST CLASsée par ordre CROISSANT
                Couleur.Red = PaletteCouleur(Temp).Red
                Couleur.Green = PaletteCouleur(Temp).Green
                Couleur.Blue = PaletteCouleur(Temp).Blue
                If Couleur.Red = 255 And Couleur.Green = 255 And Couleur.Blue = 255 Then
                    NbChateaux = NbChateaux + 1
                End If
            Next i
        End If
        Close #1
        If CreerFichierInformations Then
            FicIniCarte.Parametre("NombreChateaux") = NbChateaux 'Enregistre le nombre de chateaux dans un fichier temporaire.
        End If
    End If
    Definir_Nombre_Chateaux = NbChateaux
    Exit Property
ErreurDimension:
    MsgBox "Les dimensions de cette carte sont invalides." & Chr(10) & Chr(13) & "Les dimensions doivent être des multiples de 4.", vbCritical, "Erreur"
    End
ErreurCompression:
    MsgBox "L'image de la carte est compressée." & Chr(10) & Chr(13) & "Impossible de la lire.", vbCritical, "Erreur"
    End
Erreur:
    If Definir_ImageChateaux_Existe Then
        MsgBox "Une erreur s'est produite avec le fichier" & Chr(10) & Chr(13) & _
        App.Path & Chemin & DerniereCarte & "\" & NomImageChateaux & ExtentionImage, vbCritical, "Erreur"
    Else
        MsgBox "Le fichier : " & _
        App.Path & Chemin & DerniereCarte & "\" & NomImageChateaux & ExtentionImage & " est manquant.", vbCritical, "Erreur"
    End If
    Editer_Repertoire App.Path & Chemin & DerniereCarte
End Property

Private Function Type_Terrain(ByVal r As Byte, ByVal v As Byte, ByVal B As Byte) As Integer
    'Renvoie le type de terrain en fonction de la couleur du pixel entré en paramètre.
    Dim i As Integer
    For i = 0 To NombreTypesTerrains - 1
       If r = TerrainsCouleur(i).Red And _
          v = TerrainsCouleur(i).Green And _
          B = TerrainsCouleur(i).Blue Then
           Type_Terrain = i
           Exit For
       End If
    Next i
End Function

Private Function Type_Ressource(ByVal r As Byte, ByVal v As Byte, ByVal B As Byte) As Integer
    'Renvoie le type de resssource en fonction de la couleur du pixel entré en paramètre.
    Dim i As Integer
    If r = 255 And v = 255 And B = 255 Then
        i = Rnd * NombreTypesRessources
        If i >= NombreTypesRessources Then i = 0
        Type_Ressource = i
    Else
        For i = 0 To NombreTypesRessources - 1
           If r = RessourcesCouleur(i).Red And _
              v = RessourcesCouleur(i).Green And _
              B = RessourcesCouleur(i).Blue Then
               Type_Ressource = i
               Exit For
           End If
        Next i
    End If
End Function

Public Property Get Definir_ImageCarte_Chemin() As String
    Definir_ImageCarte_Chemin = Chemin & DerniereCarte & "\" & NomImageCarte
End Property
Public Property Get Definir_ImageCarte_Existe() As Boolean
    On Error GoTo Erreur
    'Dim Fichier As String
    'Fichier = App.Path & Definir_ImageCarte_Chemin & ExtentionImage
    'Open Fichier For Input As #1
    'Close #1
    Dim l As Long
    l = FileLen(App.Path & Definir_ImageCarte_Chemin & ExtentionImage)
    Definir_ImageCarte_Existe = l > 0
    Exit Property
Erreur:
    Definir_ImageCarte_Existe = False
End Property

Public Property Get Definir_ImageRessources_Existe() As Boolean
    On Error GoTo Erreur
    Dim l As Long
    l = FileLen(App.Path & Chemin & DerniereCarte & "\" & NomImageRessources & ExtentionImage)
    Definir_ImageRessources_Existe = l > 0
    Exit Property
Erreur:
    Definir_ImageRessources_Existe = False
End Property

Public Property Get Definir_ImageChateaux_Existe() As Boolean
    On Error GoTo Erreur
    Dim l As Long
    l = FileLen(App.Path & Chemin & DerniereCarte & "\" & NomImageChateaux & ExtentionImage)
    Definir_ImageChateaux_Existe = l > 0
    Exit Property
Erreur:
    Definir_ImageChateaux_Existe = False
End Property

Public Property Get Definir_FichierInformations_Existe() As Boolean
    On Error GoTo Erreur
    Dim l As Long
    l = FileLen(App.Path & Chemin & DerniereCarte & "\" & NomFichierInformations & FicIni.ExtensionFichierINI)
    Definir_FichierInformations_Existe = l > 0
    Exit Property
Erreur:
    Definir_FichierInformations_Existe = False
End Property

'Propriétés
Public Property Get Definir_CreerFichierInformations() As Boolean
    Definir_CreerFichierInformations = CreerFichierInformations
End Property
Public Property Let Definir_CreerFichierInformations(ByVal Valeur As Boolean)
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    CreerFichierInformations = Valeur
    FicIni.Parametre("CreerFichierInformations") = -CDbl(Valeur)
End Property
