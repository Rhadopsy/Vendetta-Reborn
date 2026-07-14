VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsJeuMonde"
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

Const FichierINI = "Terrain"
Const SectionINI1 = "Terrains"
Const SectionINI2 = "Mondes"

'Largeur d'une case du monde en pixels.
Public LargeurCase As Integer
Public HauteurCase As Integer

Public TerrainOptimise As Boolean
Public TerrainOptimiseDepart As Boolean

Public TerrainAnime As Boolean
Public NombreEtapesAnimation As Integer

Private NombreTypesTerrains As Integer
Private Type ParamTerrain
    Nom As String
    Effet As Integer
    CoefVitesse As Single
    Animation As Boolean
    Inconstructible As Boolean
    ImgPersoDecalageX As Integer
    ImgPersoDecalageY As Integer
End Type
Private Terrains() As ParamTerrain

Public NombreTypesMondes As Integer
Private Type ParamMonde
    Nom As String
    Relief As Single
    SurfaceTerrain() As Integer
    Route As Integer
End Type
Private Mondes() As ParamMonde

Private TerrainApparence() As Long
Private TerrainCategorie() As Long
Public TypeTerrain As Integer
'Dimensions en case.
Public Largeur As Long
Public Hauteur As Long
Public Diagonale As Long

Private Sub Class_Initialize()
    Dim i As Integer, j As Integer

    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI1
    
    TerrainOptimise = Val(FicIni.Parametre("TerrainOptimise"))
    TerrainOptimiseDepart = TerrainOptimise
    
    LargeurCase = Val(FicIni.Parametre("LargeurCase"))
    HauteurCase = Val(FicIni.Parametre("HauteurCase"))
    TerrainAnime = Val(FicIni.Parametre("TerrainAnime"))
    NombreEtapesAnimation = Val(FicIni.Parametre("NombreEtapesAnimation"))
    'Chargement des informations sur les terrains.
    NombreTypesTerrains = Val(FicIni.Parametre("NombreTypesTerrains"))
    ReDim Terrains(NombreTypesTerrains)
    For i = 0 To NombreTypesTerrains
        FicIni.Section = "Terrain_" & i
        Terrains(i).Nom = FicIni.Parametre("Nom")
        Terrains(i).Effet = Val(FicIni.Parametre("Effet")) - 1
        Terrains(i).CoefVitesse = Val(FicIni.Parametre("CoefVitesse"))
        Terrains(i).Animation = Val(FicIni.Parametre("Animation"))
        Terrains(i).Inconstructible = Val(FicIni.Parametre("Inconstructible"))
        Terrains(i).ImgPersoDecalageX = Val(FicIni.Parametre("ImgPersoDecalageX"))
        Terrains(i).ImgPersoDecalageY = Val(FicIni.Parametre("ImgPersoDecalageY"))
    Next i
    
    'Chargement des informations sur les mondes.
    FicIni.Section = SectionINI2
    NombreTypesMondes = Val(FicIni.Parametre("NombreTypesMondes"))
    ReDim Mondes(NombreTypesMondes - 1)
    For i = 0 To NombreTypesMondes - 1
        FicIni.Section = "Monde_" & i + 1
        Mondes(i).Nom = FicIni.Parametre("Nom")
        Mondes(i).Relief = Val(FicIni.Parametre("Relief"))
        ReDim Mondes(i).SurfaceTerrain(NombreTypesTerrains)
        For j = 0 To NombreTypesTerrains
            Mondes(i).SurfaceTerrain(j) = Val(FicIni.Parametre("SurfaceTerrain(" & j & ")"))
        Next j
        Mondes(i).Route = Val(FicIni.Parametre("Route"))
    Next i
    
    ReDim TerrainApparence(0)
    ReDim TerrainCategorie(0)
End Sub

Public Sub Generer_Monde(Optional ByVal TypeMonde As Integer)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim Relief As Single
    
    Erase TerrainApparence()
    Erase TerrainCategorie()
    ReDim TerrainApparence((Largeur + 1) * (Hauteur + 1))
    ReDim TerrainCategorie((Largeur + 1) * (Hauteur + 1))
    
    TypeTerrain = TypeMonde
    Relief = 1 - Mondes(TypeMonde).Relief
    
    For i = 0 To Largeur
        For j = 0 To Hauteur
            If Rnd > Relief Then
                TerrainCategorie(k) = ((NombreTypesTerrains - 1) * Rnd) + 1
                If TerrainCategorie(k) > NombreTypesTerrains Then
                    TerrainCategorie(k) = 0
                ElseIf Mondes(TypeMonde).SurfaceTerrain(TerrainCategorie(k)) = 0 Then
                    TerrainCategorie(k) = 0
                End If
            End If
            k = k + 1
        Next j
    Next i
    For i = 0 To NombreTypesTerrains
        For j = 0 To Mondes(TypeMonde).SurfaceTerrain(i)
            Faire_Pousser i
            If Mondes(TypeMonde).SurfaceTerrain(i) > 0 Then _
                FrmMoteur2D.Afficher_Image_Chargement dd, 30, _
                Parametres.Etiquette_Chargement(1) & _
                Format(((i - 1) / NombreTypesTerrains + j / NombreTypesTerrains / Mondes(TypeMonde).SurfaceTerrain(i)) * 100, "00") & " %"
                'Format(i / NombreTypesTerrains * 100, "00") & " % - " & Format(j / Mondes(TypeMonde).SurfaceTerrain(i) * 100, "00") & " %"
        Next j
    Next i
    Reformer_Terrain
End Sub

Public Sub Calculer_Diagonale()
    Diagonale = Sqr((Largeur * LargeurCase) ^ 2 + (Hauteur * HauteurCase) ^ 2)
End Sub

Private Sub Faire_Pousser(ByVal TypeTerrain As Integer)
    'On Local Error Resume Next
    Dim j As Long
    Dim Chance As Long
    For j = 0 To UBound(TerrainApparence())
        If TerrainApparence(j) = 0 Then
            If Definir_Terrain(j - Largeur) = TypeTerrain Then Chance = Chance + 15
            If Definir_Terrain(j + Largeur) = TypeTerrain Then Chance = Chance + 15
            If Definir_Terrain(j - 1) = TypeTerrain Then Chance = Chance + 15
            If Definir_Terrain(j + 1) = TypeTerrain Then Chance = Chance + 15
            If Chance > 0 Then
                If Definir_Terrain(j - Largeur - 1) = TypeTerrain Then Chance = Chance + 10
                If Definir_Terrain(j - Largeur + 1) = TypeTerrain Then Chance = Chance + 10
                If Definir_Terrain(j + Largeur - 1) = TypeTerrain Then Chance = Chance + 10
                If Definir_Terrain(j + Largeur + 1) = TypeTerrain Then Chance = Chance + 10
            End If
            If Rnd * 100 < Chance Then
                TerrainCategorie(j) = TypeTerrain
            End If
            Chance = 0
        End If
    Next j
    'Erreur_Inscrire "Floraison terrain : Ok" & " (" & i & "/" & Quantite & ")"
End Sub

Public Sub Contruire_Routes()
    On Error Resume Next
    Dim i As Integer, j As Integer
    Dim X As Long, Y As Long  'Coordonnées de la route que l'on trace.
    Dim IndiceRoute As Integer
    Dim RouteTraceeX1 As Boolean
    Dim RouteTraceeY1 As Boolean
    Dim RouteTraceeX2 As Boolean
    Dim RouteTraceeY2 As Boolean
    IndiceRoute = Mondes(TypeTerrain).Route
    If IndiceRoute > 0 Then
    For i = 0 To UBound(Fiefs()) - 1
        For j = i + 1 To UBound(Fiefs())
            If i <> j Then
               'Chateaux(i).Visible And _
               Chateaux(j).Visible Then
                'Trace une route du chateau_i vers chateau_j
                X = Chateaux(i).PositionX + Chateaux(i).Largeur / 2: Y = Chateaux(i).PositionY + Chateaux(i).Hauteur / 2
                RouteTraceeX1 = False: RouteTraceeY1 = False
                RouteTraceeX2 = False: RouteTraceeY2 = False
                Do
                    TerrainCategorie(TerrainIndice(X, Y)) = IndiceRoute
                    If Rnd > 0.8 Then
                        If Rnd > 0.5 Then
                            If Rnd > 0.5 Then
                                X = X + LargeurCase
                                If X > Largeur * LargeurCase Then X = Largeur * LargeurCase
                            Else
                                X = X - LargeurCase
                                If X < 0 Then X = 0
                            End If
                        Else
                            If Rnd > 0.5 Then
                                Y = Y + HauteurCase
                                'If y > Hauteur * HauteurCase Then y = Hauteur * HauteurCase
                            Else
                                Y = Y - HauteurCase
                                If Y < 0 Then Y = 0
                            End If
                        End If
                        TerrainCategorie(TerrainIndice(X, Y)) = IndiceRoute
                    End If
                    If Not RouteTraceeX1 Then
                        If X > Chateaux(j).PositionX + Chateaux(j).Largeur / 2 Then
                            X = X - LargeurCase
                        Else
                            RouteTraceeX1 = True
                        End If
                    ElseIf Not RouteTraceeX2 Then
                        If X < Chateaux(j).PositionX + Chateaux(j).Largeur / 2 Then
                            X = X + LargeurCase
                        Else
                            RouteTraceeX2 = True
                        End If
                    ElseIf Not RouteTraceeY1 Then
                        If Y > Chateaux(j).PositionY + Chateaux(j).Hauteur / 2 Then
                            Y = Y - HauteurCase
                        Else
                            RouteTraceeY1 = True
                        End If
                    ElseIf Y < Chateaux(j).PositionY + Chateaux(j).Hauteur / 2 Then
                            Y = Y + LargeurCase
                    Else
                            RouteTraceeY2 = True
                    End If
                Loop Until RouteTraceeY2
                'Loop Until Est_Superpose(x, y, Chateaux(j).PositionX, Chateaux(j).PositionY, , , Chateaux(j).Largeur, Chateaux(j).Hauteur)
            End If
        Next j
    Next i
    End If
    Reformer_Terrain
End Sub

Public Sub Contruire_Routes2()
    On Error Resume Next
    Dim i As Integer, j As Integer
    Dim k As Long, l As Long
    Dim IndiceRoute As Integer
    Dim Perso As ClsJeuPerso
    Set Perso = New ClsJeuPerso
    Perso.Init 0, Parametres
    Perso.Calculer_Longueur_Pas
    IndiceRoute = Mondes(TypeTerrain).Route
    If IndiceRoute > 0 Then
    With Perso
    For i = 0 To UBound(Fiefs()) ' - 1
        j = i + 1
        If j > UBound(Fiefs()) Then j = 0
        'For j = i + 1 To UBound(Fiefs())
            'If i <> j Then
                'Trace une route du chateau_i vers chateau_j
                'Un personnage virtuel trace le chemin sous ses pas.
                .PositionX = Chateaux(i).PositionX + Chateaux(i).Largeur / 2
                .PositionY = Chateaux(i).PositionY + Chateaux(i).Hauteur / 2
                .Aller_A Chateaux(j).PositionX + Chateaux(j).Largeur / 2, _
                         Chateaux(j).PositionY + Chateaux(j).Hauteur / 2
                Do
                    TerrainCategorie(TerrainIndice(.PositionX, .PositionY)) = IndiceRoute
                    TerrainCategorie(TerrainIndice(.PositionX + LargeurCase / 2, .PositionY + HauteurCase / 2)) = IndiceRoute
                    TerrainCategorie(TerrainIndice(.PositionX + LargeurCase, .PositionY + HauteurCase)) = IndiceRoute
                    'TerrainCategorie(TerrainIndice(.PositionX + LargeurCase * 1.5, .PositionY + HauteurCase * 1.5)) = IndiceRoute
                    'TerrainCategorie(TerrainIndice(.PositionX + Monde.LargeurCase * 2, .PositionY + Monde.HauteurCase * 2)) = IndiceRoute
                    If Rnd > 0.8 Then
                        If Rnd > 0.5 Then
                            If Rnd > 0.5 Then
                                .PositionX = .PositionX + LargeurCase / 2
                                If .PositionX > Largeur * LargeurCase Then .PositionX = Largeur * LargeurCase
                            Else
                                .PositionX = .PositionX - LargeurCase / 2
                                If .PositionX < 0 Then .PositionX = 0
                            End If
                        Else
                            If Rnd > 0.5 Then
                                .PositionY = .PositionY + HauteurCase / 2
                                If .PositionY > Hauteur * HauteurCase Then .PositionY = Hauteur * HauteurCase
                            Else
                                .PositionY = .PositionY - HauteurCase / 2
                                If .PositionY < 0 Then .PositionY = 0
                            End If
                        End If
                    End If
                Loop Until Not .Marcher2
            'End If
        'Next j
        'Pave autour du chateau.
        For k = Chateaux(i).PositionX - LargeurCase To Chateaux(i).PositionX + Chateaux(i).Largeur + LargeurCase Step LargeurCase
            For l = Chateaux(i).PositionY To Chateaux(i).PositionY + Chateaux(i).Hauteur Step HauteurCase
                TerrainCategorie(TerrainIndice(k, l)) = IndiceRoute
            Next l
        Next k
    Next i
    End With
    End If
    Reformer_Terrain
End Sub

Public Sub Reformer_Terrain()
    'Cette fonction ajuste les images de chaque case de terrain en fonction des cases adjacentes.
    'On Error Resume Next
    Dim i As Long
    For i = 0 To UBound(TerrainCategorie())
        'Herbe.
        If Definir_Terrain(i) > 0 Then
            If Definir_Terrain(i - Largeur) > 0 And _
               Definir_Terrain(i + Largeur) > 0 And _
               Definir_Terrain(i - 1) > 0 And _
               Definir_Terrain(i + 1) > 0 Then
               TerrainApparence(i) = 1 'Terrain plein.
            ElseIf Definir_Terrain(i - Largeur) = 0 And _
                   Definir_Terrain(i + Largeur) > 0 And _
                   Definir_Terrain(i - 1) = 0 And _
                   Definir_Terrain(i + 1) = 0 Then
                   TerrainApparence(i) = 2 'Pointe haut.
            ElseIf Definir_Terrain(i - Largeur) > 0 And _
                   Definir_Terrain(i + Largeur) = 0 And _
                   Definir_Terrain(i - 1) = 0 And _
                   Definir_Terrain(i + 1) = 0 Then
                   TerrainApparence(i) = 3  'Pointe bas.
            ElseIf Definir_Terrain(i - Largeur) = 0 And _
                   Definir_Terrain(i + Largeur) = 0 And _
                   Definir_Terrain(i - 1) = 0 And _
                   Definir_Terrain(i + 1) > 0 Then
                   TerrainApparence(i) = 4  'Pointe gauche.
            ElseIf Definir_Terrain(i - Largeur) = 0 And _
                   Definir_Terrain(i + Largeur) = 0 And _
                   Definir_Terrain(i - 1) > 0 And _
                   Definir_Terrain(i + 1) = 0 Then
                   TerrainApparence(i) = 5  'Pointe droite.
            ElseIf Definir_Terrain(i - Largeur) > 0 And _
                   Definir_Terrain(i + Largeur) > 0 And _
                   Definir_Terrain(i - 1) = 0 And _
                   Definir_Terrain(i + 1) = 0 Then
                   TerrainApparence(i) = 6 'Pont vertical.
            ElseIf Definir_Terrain(i - Largeur) = 0 And _
                   Definir_Terrain(i + Largeur) = 0 And _
                   Definir_Terrain(i - 1) > 0 And _
                   Definir_Terrain(i + 1) > 0 Then
                   TerrainApparence(i) = 7  'Pont horizontal.
            ElseIf Definir_Terrain(i - Largeur) = 0 And _
                   Definir_Terrain(i + Largeur) > 0 And _
                   Definir_Terrain(i - 1) > 0 And _
                   Definir_Terrain(i + 1) > 0 Then
                   TerrainApparence(i) = 8  'Creu en haut.
            ElseIf Definir_Terrain(i - Largeur) > 0 And _
                   Definir_Terrain(i + Largeur) = 0 And _
                   Definir_Terrain(i - 1) > 0 And _
                   Definir_Terrain(i + 1) > 0 Then
                   TerrainApparence(i) = 9  'Creu en bas.
            ElseIf Definir_Terrain(i - Largeur) > 0 And _
                   Definir_Terrain(i + Largeur) > 0 And _
                   Definir_Terrain(i - 1) = 0 And _
                   Definir_Terrain(i + 1) > 0 Then
                   TerrainApparence(i) = 10 'Creu à gauche.
            ElseIf Definir_Terrain(i - Largeur) > 0 And _
                   Definir_Terrain(i + Largeur) > 0 And _
                   Definir_Terrain(i - 1) > 0 And _
                   Definir_Terrain(i + 1) = 0 Then
                   TerrainApparence(i) = 11  'Creu à droite.
            ElseIf Definir_Terrain(i - Largeur) = 0 And _
                   Definir_Terrain(i + Largeur) > 0 And _
                   Definir_Terrain(i - 1) = 0 And _
                   Definir_Terrain(i + 1) > 0 Then
                   TerrainApparence(i) = 12  'Coin en haut à gauche.
            ElseIf Definir_Terrain(i - Largeur) = 0 And _
                   Definir_Terrain(i + Largeur) > 0 And _
                   Definir_Terrain(i - 1) > 0 And _
                   Definir_Terrain(i + 1) = 0 Then
                   TerrainApparence(i) = 13  'Coin en haut à droite.
            ElseIf Definir_Terrain(i - Largeur) > 0 And _
                   Definir_Terrain(i + Largeur) = 0 And _
                   Definir_Terrain(i - 1) = 0 And _
                   Definir_Terrain(i + 1) > 0 Then
                   TerrainApparence(i) = 14  'Coin en bas à gauche.
            ElseIf Definir_Terrain(i - Largeur) > 0 And _
                   Definir_Terrain(i + Largeur) = 0 And _
                   Definir_Terrain(i - 1) > 0 And _
                   Definir_Terrain(i + 1) = 0 Then
                   TerrainApparence(i) = 15  'Coin en bas à droite.
            Else: TerrainCategorie(i) = 0
            End If
        End If
    Next i
    Calculer_Diagonale
End Sub

Public Function Constructible(ByVal X As Long, _
                              ByVal Y As Long, _
                              Optional ByVal l As Integer, _
                              Optional ByVal H As Integer) As Boolean
    'Renvoie vrai, si l'on peut construire sur cet emplacement.
    'X, Y : Coordonnées.
    'L : Largeur, H : Hauteur.
    Dim i As Integer, j As Integer
    Dim Temp As Long
    If l = 0 And H = 0 Then
        'Cas d'un point.
        Temp = TerrainIndice(X, Y)
        If Temp > 0 And Temp < UBound(TerrainCategorie()) Then
            Constructible = Not Terrains(TerrainCategorie(Temp)).Inconstructible
        Else
            Constructible = False
        End If
    Else 'Cas d'une surface.
        Constructible = True
        For i = 0 To l Step LargeurCase
            For j = 0 To H Step HauteurCase
                If Not Constructible(X + i, Y + j) Then
                    Constructible = False
                    i = l
                    j = H
                End If
            Next j
        Next i
    End If
End Function

Public Function CoefVitesse(ByRef Perso As ClsJeuPerso, _
                            ByVal X As Long, _
                            ByVal Y As Long) As Single
    'Renvoie le coefficient qui influence la vitesse du personnage sur cet emplacement.
    'X, Y : Coordonnées.
    Dim Temp As Long
    Temp = TerrainIndice(X, Y)
    If Temp > 0 And Temp < UBound(TerrainCategorie()) Then
        CoefVitesse = Terrains(TerrainCategorie(Temp)).CoefVitesse
        If Terrains(TerrainCategorie(Temp)).Effet > -1 Then
            If Perso.Definir_Resistance(Parametres.Objet_AttaqueResistance(Terrains(TerrainCategorie(Temp)).Effet)) > Parametres.ResistanceMinimum Then
                Perso.EffetTemp_Ajouter Terrains(TerrainCategorie(Temp)).Effet, Parametres.Objet_Duree(Terrains(TerrainCategorie(Temp)).Effet), Parametres.Objet_Duree(Terrains(TerrainCategorie(Temp)).Effet)
            End If
        End If
    Else
        CoefVitesse = 1
    End If
End Function
Public Function CoefVitesseSimple(ByVal X As Long, _
                                  ByVal Y As Long) As Single
    'Renvoie le coefficient qui influence la vitesse du personnage sur cet emplacement.
    'X, Y : Coordonnées.
    Dim Temp As Long
    Temp = TerrainIndice(X, Y)
    If Temp > 0 And Temp < UBound(TerrainCategorie()) Then
        CoefVitesseSimple = Terrains(TerrainCategorie(Temp)).CoefVitesse
    Else
        CoefVitesseSimple = 1
    End If
End Function
Public Function ImgDecalageX(ByVal X As Long, _
                             ByVal Y As Long) As Integer
    'Modifie l'image du personnage sur cet emplacement.
    'X, Y : Coordonnées.
    Dim Temp As Long
    Temp = TerrainIndice(X, Y)
    If Temp > 0 And Temp < UBound(TerrainCategorie()) Then
        ImgDecalageX = Terrains(TerrainCategorie(Temp)).ImgPersoDecalageX
    End If
End Function
Public Function ImgDecalageY(ByVal X As Long, _
                             ByVal Y As Long) As Integer
    'Modifie l'image du personnage sur cet emplacement.
    'X, Y : Coordonnées.
    Dim Temp As Long
    Temp = TerrainIndice(X, Y)
    If Temp > 0 And Temp < UBound(TerrainCategorie()) Then
        ImgDecalageY = Terrains(TerrainCategorie(Temp)).ImgPersoDecalageY
    End If
End Function

Public Function TerrainIndice(ByVal X As Long, _
                              ByVal Y As Long) As Long
    'Renvoie l'indice du terrain sur cet emplacement.
    'X, Y : Coordonnées.
    TerrainIndice = CLng(X / LargeurCase) + CLng(Y / HauteurCase) * Largeur
End Function


Private Property Get Definir_Terrain(ByVal Index As Long) As Long
    If Index >= 0 And Index <= UBound(TerrainApparence()) Then
        Definir_Terrain = TerrainCategorie(Index)
    Else
        Definir_Terrain = 0
    End If
End Property

Public Property Get Param_TerrainApparence(ByVal Index As Long) As Integer
    Param_TerrainApparence = TerrainApparence(Index)
End Property
Public Property Let Param_TerrainApparence(ByVal Index As Long, ByVal Valeur As Integer)
    If Index > UBound(TerrainApparence()) Then
        ReDim Preserve TerrainApparence(Index)
    End If
    TerrainApparence(Index) = Valeur
End Property

Public Property Get Param_TerrainCategorie(ByVal Index As Long) As Integer
    Param_TerrainCategorie = TerrainCategorie(Index)
End Property
Public Property Let Param_TerrainCategorie(ByVal Index As Long, ByVal Valeur As Integer)
    If Index > UBound(TerrainCategorie()) Then
        ReDim Preserve TerrainCategorie(Index)
    End If
    TerrainCategorie(Index) = Valeur
End Property
Public Property Get Param_TerrainAnimation(ByVal Index As Long) As Integer
    Param_TerrainAnimation = Terrains(TerrainCategorie(Index)).Animation
End Property

Public Function Longueur_Diagonale() As Long
    Longueur_Diagonale = Sqr((Largeur * LargeurCase) ^ 2 + _
                             (Hauteur * HauteurCase) ^ 2)
End Function
    
Public Property Get NombreCases() As Long
    NombreCases = UBound(TerrainApparence()) + 1
End Property
Public Property Let NombreCases(ByVal Valeur As Long)
    ReDim Preserve TerrainApparence(Valeur)
    ReDim Preserve TerrainCategorie(Valeur)
End Property
Public Property Get NombreCasesConstructibles() As Long
    Dim i As Long
    Dim Res As Long
    For i = 0 To UBound(TerrainApparence())
        If Not Terrains(TerrainCategorie(i)).Inconstructible Then
            Res = Res + 1
        End If
    Next i
    NombreCasesConstructibles = Res
End Property

' ************************************************************************************
' *                             Propriétes des mondes                                *
' ************************************************************************************

Public Property Get Param_MondeNom(ByVal Index As Long) As String
    Param_MondeNom = Mondes(Index).Nom
End Property

Public Property Get Definir_TerrainOptimise() As Boolean
    Definir_TerrainOptimise = TerrainOptimiseDepart
End Property
Public Property Let Definir_TerrainOptimise(ByVal Valeur As Boolean)
    TerrainOptimise = Valeur
    TerrainOptimiseDepart = Valeur
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI1
    FicIni.Parametre("TerrainOptimise") = -CBool(Valeur)
End Property

Public Property Get Definir_TerrainAnime() As Boolean
    Definir_TerrainAnime = TerrainAnime
End Property
Public Property Let Definir_TerrainAnime(ByVal Valeur As Boolean)
    TerrainAnime = Valeur
    FicIni.Fichier = FicIni.Chemin & Langues.Dossier & FichierINI
    FicIni.Section = SectionINI1
    FicIni.Parametre("TerrainAnime") = -CBool(Valeur)
End Property

