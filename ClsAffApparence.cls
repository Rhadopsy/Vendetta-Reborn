VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAffApparence"
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

Const FichierINI = "Apparences"
Const SectionINI = "Apparences"

Private Chemin As String

'Public NombreCategoriesBase As Integer
'Private DossierCategorie() As String
'Private NombreImageCategorieF() As Integer
'Private NombreImageCategorieM() As Integer
'Private SousDossierFeminin As String
'Private SousDossierMasculin As String
'Private DossierRaces As String
'Private DossierObjets As String

Private NombreApparences As Long
Private NombreDossiers As Integer
Private DossiersApparences() As String
Private DossiersApparencesNombreImages() As Integer

Private Sprites() As ClsAffSprite
Private SpritesNom() As String
Private SpritesCategorie() As String

Public Sub Init1()
    Dim i As Integer, j As Integer
    Dim Temp As String
        
    FicIni.Fichier = FicIni.Chemin & FichierINI
    FicIni.Section = SectionINI
    
    Chemin = FicIni.Parametre("Chemin")
    'NombreCategoriesBase = Val(FicIni.Parametre("NombreCategoriesBase"))
    'ReDim DossierCategorie(NombreCategoriesBase - 1)
    'ReDim NombreImageCategorieF(NombreCategoriesBase - 1)
    'ReDim NombreImageCategorieM(NombreCategoriesBase - 1)
    'For i = 0 To NombreCategoriesBase - 1
    '    DossierCategorie(i) = FicIni.Parametre("DossierCategorie" & i + 1)
    'Next i
    'SousDossierFeminin = FicIni.Parametre("SousDossierFeminin")
    'SousDossierMasculin = FicIni.Parametre("SousDossierMasculin")
    'DossierRaces = FicIni.Parametre("DossierRaces")
    'DossierObjets = FicIni.Parametre("DossierObjets")
   
    ReDim Sprites(0)
    Set Sprites(0) = New ClsAffSprite
    ReDim SpritesNom(0)
    ReDim SpritesCategorie(0)
    
    'Recherche les dossiers où il y a les images des apparences.
    ReDim DossiersApparences(0)
    Temp = Dir(App.Path & Sprites(0).CheminImage & Chemin, vbDirectory)
    While Temp <> ""
        If InStr(Temp, ".") = 0 Then
            ReDim Preserve DossiersApparences(UBound(DossiersApparences()) + 1)
            DossiersApparences(UBound(DossiersApparences()) - 1) = Temp
        End If
        Temp = Dir()
    Wend
    ReDim Preserve DossiersApparences(UBound(DossiersApparences()) - 1)
    'Recherche les sous répertoires.
    For i = 0 To UBound(DossiersApparences())
        Temp = Dir(App.Path & Sprites(0).CheminImage & Chemin & DossiersApparences(i) & "\", vbDirectory)
        While Temp <> ""
            If InStr(Temp, ".") = 0 Then
                ReDim Preserve DossiersApparences(UBound(DossiersApparences()) + 1)
                DossiersApparences(UBound(DossiersApparences())) = DossiersApparences(i) & "\" & Temp
            End If
            Temp = Dir()
        Wend
    Next i
    NombreDossiers = UBound(DossiersApparences()) + 1
End Sub

Public Sub Init2()
    Dim i As Integer, j As Integer
    Dim Temp As String
    ReDim DossiersApparencesNombreImages(NombreDossiers - 1)
    
    NombreApparences = 0
    For j = 0 To UBound(DossiersApparences())
        If Left(DossiersApparences(j), 1) = "." Then
            Temp = Dir(App.Path & Right(DossiersApparences(j), Len(DossiersApparences(j)) - 1) & "\", vbNormal)
        Else
            Temp = Dir(App.Path & Sprites(0).CheminImage & Chemin & DossiersApparences(j) & "\", vbNormal)
        End If
        While Temp <> ""
            NombreApparences = NombreApparences + 1
            DossiersApparencesNombreImages(j) = DossiersApparencesNombreImages(j) + 1
            Temp = Dir()
        Wend
    Next j
    ReDim SpritesNom(NombreApparences - 1)
    ReDim SpritesCategorie(NombreApparences - 1)
    i = 0
    For j = 0 To UBound(DossiersApparences())
        If Left(DossiersApparences(j), 1) = "." Then
            Temp = Dir(App.Path & Right(DossiersApparences(j), Len(DossiersApparences(j)) - 1) & "\", vbNormal)
        Else
            Temp = Dir(App.Path & Sprites(0).CheminImage & Chemin & DossiersApparences(j) & "\", vbNormal)
        End If
        While Temp <> ""
            SpritesNom(i) = Temp
            'Enregistre dans quelle catégorie se trouve l'image.
            SpritesCategorie(i) = DossiersApparences(j)
            i = i + 1
            Temp = Dir()
        Wend
    Next j
End Sub

Public Sub Charger(ByRef dd As DirectDraw7)
    Dim i As Long, j As Long
    Dim Temp As String
    
    ReDim Sprites(NombreApparences - 1)
    
    'Chargement des apparences.
    Set Sprites(0) = New ClsAffSprite
    For j = 0 To UBound(DossiersApparences())
        If Left(DossiersApparences(j), 1) = "." Then
            Temp = Dir(App.Path & Right(DossiersApparences(j), Len(DossiersApparences(j)) - 1) & "\", vbNormal)
        Else
            Temp = Dir(App.Path & Sprites(0).CheminImage & Chemin & DossiersApparences(j) & "\", vbNormal)
        End If
        While Temp <> ""
            Set Sprites(i) = New ClsAffSprite
            If Left(DossiersApparences(j), 1) = "." Then
                Sprites(i).Charger dd, Right(DossiersApparences(j), Len(DossiersApparences(j)) - 1) & "\" & _
                                       Left$(Temp, InStr(Temp, ".") - 1), True
            Else
                Sprites(i).Charger dd, Chemin & DossiersApparences(j) & "\" & _
                                       Left$(Temp, InStr(Temp, ".") - 1)
            End If
            i = i + 1
            Temp = Dir()
        Wend
    Next j
End Sub

Public Sub Tirer_Apparence_Au_Hasard(ByRef Perso As ClsJeuPerso, _
                                     ByVal Parametres As ClsJeuParam)
    Dim i As Long, j As Long, k As Long
    Dim Race As Integer
    Dim Dossier As String
    
    With Parametres
    'Perso.Nombre_ApparenceBase = NombreCategoriesBase
    Race = Perso.Race
    If .Race_DefinirNombreCategoriesApparences(Race) > 0 Then
        Perso.Nombre_ApparenceBase = .Race_DefinirNombreCategoriesApparences(Race)
        For i = 0 To Perso.Nombre_ApparenceBase - 1
            'Choisit la tête, le corps, les cheveux et les habits de base du personnage.
            Dossier = .Race_DefinirDossiersCategories(Race, i, Perso.Feminin)
            For k = 0 To NombreDossiers - 1
                If DossiersApparences(k) = Dossier Then
                    j = DossiersApparencesNombreImages(k)
                End If
            Next k
            'If Perso.Feminin Then
            '    j = NombreImageCategorieF(i)
            'Else
            '    j = NombreImageCategorieM(i)
            'End If
            
            j = Int(Rnd * j)
            
            For k = 0 To NombreApparences - 1
                'If SpritesCategorie(k) = i * 2 + IIf(Perso.Feminin, 1, 0) Then
                If SpritesCategorie(k) = Dossier Then
                    If j = 0 Then
                        Perso.Definir_FichierApparenceBase(i) = SpritesNom(k)
                        k = NombreApparences
                    End If
                    j = j - 1
                End If
            Next k
        Next i
    End If
    End With
End Sub

Public Sub Definir_Indices(ByRef Perso As ClsJeuPerso)
    'Cette fonction affecte les indices aux fichiers d'apparences choisies.
    Dim i As Integer, j As Integer
    Dim Temp As Boolean
    If Parametres.Race_DefinirNombreCategoriesApparences(Perso.Race) > 0 Then
        For i = 0 To Perso.Nombre_ApparenceBase - 1
            Temp = False
            For j = 0 To NombreApparences - 1
                If Perso.Definir_FichierApparenceBase(i) = SpritesNom(j) Then
                    Perso.Definir_ApparenceBase(i) = j
                    j = NombreApparences
                    Temp = True
                End If
            Next j
            If Not Temp Then 'Tire une apparence au hasard si aucune apparence n'a été trouvée.
                Tirer_Apparence_Au_Hasard Perso, Parametres
                Definir_Indices Perso
            End If
        Next i
    End If
End Sub

Public Sub Definir_Indices_Generaux(ByRef Parametres As ClsJeuParam)
    'Cette fonction affecte les indices aux fichiers d'apparences choisies.
    Dim i As Integer, j As Integer
    With Parametres
    For i = 0 To .NombreObjets - 1
        If .Objet_ChangerApparence(i) Then
            For j = 0 To NombreApparences - 1
                If .Objet_FichierApparence(i, False) = SpritesNom(j) Then
                    .Objet_NumeroApparence(i, False) = j
                End If
                If .Objet_FichierApparence(i, True) = SpritesNom(j) Then
                    .Objet_NumeroApparence(i, True) = j
                End If
            Next j
        End If
    Next i
    End With
End Sub

Public Sub Afficher(ByVal X As Long, _
                    ByVal Y As Long, _
                    ByRef NumeroPas As Integer, _
                    ByVal Direction As Integer, _
                    ByVal LargeurImage As Integer, _
                    ByVal HauteurImage As Integer, _
                    ByVal Perso As ClsJeuPerso, _
                    ByRef SurfBack As DirectDrawSurface7, _
                    Optional ByVal ChoisirApparence As Boolean, _
                    Optional ByVal Translucide As Boolean, _
                    Optional ByVal DecalageX As Integer = 0, _
                    Optional ByVal DecalageY As Integer = 0)
    'Pour la direction :
    '0 = immobile '1 = haut '2 = droite '3 = bas '4 = gauche
    Dim i As Integer
    Dim Temp As Boolean
    
    With Perso
    If .Monture > 0 Then
        Temp = True 'Passe à faux si l'on affiche la monture en avant plan.
        'Affiche la monture si elle est en avant plan.
        If .Definir_ApparenceAvantPlan(Direction) Then
            Afficher_Element .Definir_ApparenceObjet(.Monture), X + .ApparenceDecalageXTemp, Y + .ApparenceDecalageXTemp, NumeroPas, Direction, .Definir_ApparenceLargeur, .Definir_ApparenceHauteur, SurfBack, Translucide, DecalageX, DecalageY
            Temp = False
        End If
        For i = 0 To .Nombre_ApparenceBase - 1
            If DecalageX = 0 And DecalageY = 0 Then
                Afficher_Element .Definir_ApparenceBase(i), X + .Definir_ApparenceDecalageX(Direction), Y + .Definir_ApparenceDecalageY(Direction), 1, Direction, LargeurImage, HauteurImage, SurfBack, Translucide, DecalageX, DecalageY
            Else
                Afficher_Element .Definir_ApparenceBase(i), X + .Definir_ApparenceDecalageX(Direction) + DecalageX, Y + .Definir_ApparenceDecalageY(Direction) + DecalageY, 1, Direction, LargeurImage, HauteurImage, SurfBack, Translucide, DecalageX + .Definir_ApparenceDecalageX(Direction), DecalageY + .Definir_ApparenceDecalageY(Direction)
            End If
        Next i
        If .ApparenceObjetAfficher Then
            For i = 1 To .Nombre_ApparenceObjet
                If i = .Monture Then 'Cas où l'on a affaire à l'affichage de la monture.
                    If Temp Then
                        Afficher_Element .Definir_ApparenceObjet(i), X + .ApparenceDecalageXTemp, Y + .ApparenceDecalageXTemp, NumeroPas, Direction, .Definir_ApparenceLargeur, .Definir_ApparenceHauteur, SurfBack, Translucide, DecalageX, DecalageY
                    End If
                Else
                    If DecalageX = 0 And DecalageY = 0 Then
                        Afficher_Element .Definir_ApparenceObjet(i), X + .Definir_ApparenceDecalageX(Direction), Y + .Definir_ApparenceDecalageY(Direction), 1, Direction, LargeurImage, HauteurImage, SurfBack, Translucide, DecalageX, DecalageY
                    Else
                        Afficher_Element .Definir_ApparenceObjet(i), X + .Definir_ApparenceDecalageX(Direction) + DecalageX, Y + .Definir_ApparenceDecalageY(Direction) + DecalageY, 1, Direction, LargeurImage, HauteurImage, SurfBack, Translucide, DecalageX + .Definir_ApparenceDecalageX(Direction), DecalageY + .Definir_ApparenceDecalageY(Direction)
                    End If
                End If
            Next i
        End If
    Else
        For i = 0 To .Nombre_ApparenceBase - 1
            Afficher_Element .Definir_ApparenceBase(i), X, Y, NumeroPas, Direction, LargeurImage, HauteurImage, SurfBack, Translucide, DecalageX, DecalageY
        Next i
        If .ApparenceObjetAfficher Then
            For i = 1 To .Nombre_ApparenceObjet
                Afficher_Element .Definir_ApparenceObjet(i), X, Y, NumeroPas, Direction, LargeurImage, HauteurImage, SurfBack, Translucide, DecalageX, DecalageY
            Next i
        End If
    End If
    End With
End Sub

Public Sub Afficher_Element(ByVal Apparence As Integer, _
                            ByVal X As Long, _
                            ByVal Y As Long, _
                            ByRef NumeroPas As Integer, _
                            ByVal Direction As Integer, _
                            ByVal LargeurImage As Integer, _
                            ByVal HauteurImage As Integer, _
                            ByRef SurfBack As DirectDrawSurface7, _
                            Optional ByVal Translucide As Boolean, _
                            Optional ByVal DecalageX As Integer = 0, _
                            Optional ByVal DecalageY As Integer = 0)
    With Sprites(Apparence)
        
        If Direction > 0 Then
            If NumeroPas = 3 Then
                .Selection (Direction - 1) * HauteurImage, _
                            Direction * HauteurImage, _
                            LargeurImage, _
                            2 * LargeurImage
            Else
                .Selection (Direction - 1) * HauteurImage, _
                            Direction * HauteurImage, _
                            NumeroPas * LargeurImage, _
                            (NumeroPas + 1) * LargeurImage
            End If
        Else
            .Selection HauteurImage * 2, _
                        HauteurImage * 3, _
                        LargeurImage, _
                        2 * LargeurImage
        End If
        
        'On affiche le sprite sur l'arrière plan.
        .Afficher X, Y, SurfBack, True, Translucide, DecalageX, DecalageY
        
        End With
End Sub

Public Property Get Definir_Chemin(ByVal Categorie As Integer, _
                                   ByVal Race As Integer, _
                                   ByVal Feminin As Boolean, _
                                   ByVal Parametres As ClsJeuParam) As String
    Dim Temp As String
    Temp = Parametres.Race_DefinirDossiersCategories(Race, Categorie, Feminin)
    If Left(Temp, 1) = "." Then
        'Chemin absolu.
        Definir_Chemin = App.Path & Right(Temp, Len(Temp) - 1)
    Else
        'Chemin relatif.
        Definir_Chemin = App.Path & _
                         Sprites(0).CheminImage & _
                         Chemin & _
                         Temp
    End If
    'DossierCategorie(Categorie) & "\" & IIf(Feminin, SousDossierFeminin, SousDossierMasculin)
End Property

Public Sub Ajouter_Dossier(ByVal Dossier As String)
    Dim Temp As String
    Dim i As Integer
    Dim DossierExistant As Boolean
    DossierExistant = False 'Si passe à vrai, le dossier existe déjà et il n'est pas besoin de l'ajouter.
    
    For i = 0 To NombreDossiers - 1
        If DossiersApparences(i) = Dossier Then
            DossierExistant = True
            Exit For
        End If
    Next i
    
    If Not DossierExistant Then
        NombreDossiers = NombreDossiers + 1
        ReDim Preserve DossiersApparences(NombreDossiers - 1)
        ReDim Preserve DossiersApparencesNombreImages(NombreDossiers - 1)
        DossiersApparences(NombreDossiers - 1) = Dossier
'        i = NombreApparences - 1
'
'        Temp = Dir(App.Path & Dossier & "\", vbNormal)
'        While Temp <> ""
'            NombreApparences = NombreApparences + 1
'            DossiersApparencesNombreImages(NombreDossiers - 1) = DossiersApparencesNombreImages(NombreDossiers - 1) + 1
'            Temp = Dir()
'        Wend
'        ReDim Preserve SpritesNom(NombreApparences - 1)
'        ReDim Preserve SpritesCategorie(NombreApparences - 1)
'        Temp = Dir(App.Path & DossiersApparences(NombreDossiers - 1) & "\", vbNormal)
'        While Temp <> ""
'            SpritesNom(i) = Temp
'            'Enregistre dans quelle catégorie se trouve l'image.
'            SpritesCategorie(i) = DossiersApparences(NombreDossiers - 1)
'            i = i + 1
'            Temp = Dir()
'        Wend
    End If
End Sub
