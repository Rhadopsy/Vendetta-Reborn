Attribute VB_Name = "ModGenererMonde"
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

'Public Sub Placer_Chateaux()
'    Dim i As Long
'    Dim j As Long
'    Dim k As Long
'    Dim DistanceMin As Long
'    Dim EmplacementOk As Boolean
'
'    For j = 0 To UBound(Fiefs())
'        For DistanceMin = CLng((Monde.Largeur * AffTerrain.Largeur + _
'                          Monde.Hauteur * AffTerrain.Hauteur) / 2) To 0 _
'                          Step CLng(-(Parametres.ChateauxLargeur + Parametres.ChateauxHauteur))  ' / 2)
'            For i = 0 To 1000 'Nombre de tentatives.
'                Chateaux(j).PositionX = CLng(RND * (Monde.Largeur * AffTerrain.Largeur - Parametres.ChateauxLargeur))
'                Chateaux(j).PositionY = CLng(RND * (Monde.Hauteur * AffTerrain.Hauteur - Parametres.ChateauxHauteur))
'                If ObstacleChateau(j, Chateaux(j).PositionX, Chateaux(j).PositionY, Parametres.ChateauxLargeur, Parametres.ChateauxHauteur) Then
'                    EmplacementOk = True
'                    For k = 0 To UBound(Fiefs())
'                        If DistanceMin > Sqr((Chateaux(k).PositionX - Chateaux(j).PositionX) ^ 2 + _
'                                             (Chateaux(k).PositionY - Chateaux(j).PositionY) ^ 2) Or _
'                           Chateaux(j).PositionX < Screen.Width / 15 / 6 Or _
'                           Chateaux(j).PositionY < Screen.Width / 15 / 6 Or _
'                           Chateaux(j).PositionX > Monde.Largeur * AffTerrain.Largeur - Screen.Width / 15 / 6 - Parametres.ChateauxLargeur Or _
'                           Chateaux(j).PositionY > Monde.Hauteur * AffTerrain.Hauteur - Screen.Width / 15 / 6 - Parametres.ChateauxHauteur Then
'                            EmplacementOk = False
'                            k = UBound(Fiefs())
'                            'j = UBound(Fiefs())
'                        End If
'                    Next k
'                    If EmplacementOk Then
'                        i = 100000
'                        DistanceMin = 0
'                    End If
'                End If
'            Next i
'        Next DistanceMin
'    Next j
'
'
'    For j = 0 To UBound(Chateaux())
'        Chateaux(j).MaxVie = Parametres.ChateauxMaxVie
'        Chateaux(j).Definir_Vie = Chateaux(j).MaxVie
'        'Debug.Print Chateaux(j).PositionX & ", " & Chateaux(j).PositionY
'    Next j
'End Sub

Public Sub Tirer_Monstre_Aleatoire(ByVal NumeroPerso As Long, _
                                   ByVal Peuple As Integer, _
                                   Optional ByVal Grade As Integer = -1, _
                                   Optional ByVal SansEquipement As Boolean)
    Dim i As Integer
    Dim Temp As Integer
    Dim Race As Integer
    Dim FichierIA As String
    
    With Parametres
    If .Peuples_NombreMonstres(Peuple) > 0 Then
        'Tire un monstre au hasard
        If Grade = -1 Then
            For i = 0 To .Peuples_NombreMonstres(Peuple) - 1
                If Rnd < .Peuples_MonstreChance(Peuple, i) Then
                    Race = .Peuples_MonstreIndice(Peuple, i)
                    FichierIA = .Peuples_MonstreFichierIA(Peuple, i)
                    i = .Peuples_NombreMonstres(Peuple)
                End If
            Next i
            If Race = 0 Then
                Race = .Peuples_MonstreIndice(Peuple, .Peuples_NombreMonstres(Peuple) - 1)
                FichierIA = .Peuples_MonstreFichierIA(Peuple, .Peuples_NombreMonstres(Peuple) - 1)
            End If
        Else
            Race = .Peuples_MonstreIndice(Peuple, Grade)
            FichierIA = .Peuples_MonstreFichierIA(Peuple, Grade)
        End If
    Else 'Tire au hasard une race jouable.
        If Rnd < .ChancesRaceAleatoire Then
            Temp = Int(Rnd * (.NombreRacesJouables)) '- 1))
            For i = 1 To .NombreRaces - 1
                If .Race_Jouable(i) Then
                    Temp = Temp - 1
                    If Temp <= 0 Then
                        Race = i
                        i = .NombreRaces
                    End If
                End If
            Next i
        Else
            Race = 0
        End If
    End If
    Persos(NumeroPerso).Changer_Race_Personnage Race, Parametres, SansEquipement
    If FichierIA <> "" Then 'Attribue le fichierIA définit dans le peuple.
        Persos(NumeroPerso).FichierIA = FichierIA
        Persos(NumeroPerso).ChoisirIA = True
    End If
    End With
End Sub


Public Sub Placer_Chateaux(Optional ByVal PositionAleatoire As Boolean = True)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    If PositionAleatoire Then
        Dim PlaceChateaux As ClsPlaceObject, PlaceResult As Collection, PlaceOfCastel As ClsMatrixPoint
        Set PlaceChateaux = New ClsPlaceObject
        PlaceChateaux.SetBoardSize Monde.Hauteur - 9, Monde.Largeur - 9
        If UBound(Fiefs()) > 0 Then PlaceChateaux.SetNumberOfObject UBound(Fiefs()) + 1
        PlaceChateaux.SetPlacedObjectAsWall True
        PlaceChateaux.CalculateObjectsPlace
        Set PlaceResult = PlaceChateaux.GetLastObjectCollection
        k = 0
        For Each PlaceOfCastel In PlaceResult
            Chateaux(k).PositionX = CLng(Int(Monde.LargeurCase / 2 - Rnd * Monde.LargeurCase) + PlaceOfCastel.ColumnPos * Monde.LargeurCase)
            Chateaux(k).PositionY = CLng(Int(Monde.HauteurCase / 2 - Rnd * Monde.HauteurCase) + PlaceOfCastel.LinePos * Monde.HauteurCase)
            k = k + 1
        Next
    End If
    
    For i = 0 To UBound(Chateaux())
        Chateaux(i).MaxVie = Parametres.Peuples_ChateauMaxVie(Fiefs(i).Definir_TypePeuple)
        Chateaux(i).Largeur = Parametres.Peuples_ChateauLargeur(Fiefs(i).Definir_TypePeuple)
        Chateaux(i).Hauteur = Parametres.Peuples_ChateauHauteur(Fiefs(i).Definir_TypePeuple)
        Chateaux(i).EntreePositionX = Parametres.Peuples_ChateauEntreePositionX(Fiefs(i).Definir_TypePeuple)
        Chateaux(i).EntreePositionY = Parametres.Peuples_ChateauEntreePositionY(Fiefs(i).Definir_TypePeuple)
            For j = 0 To Fiefs(i).Epoque
                Chateaux(i).MaxVie = Chateaux(i).MaxVie + Parametres.Epoque_BonusVieChateau(j)
            Next j
        Chateaux(i).Definir_Vie = Chateaux(i).MaxVie
    Next i
End Sub

Public Sub Placer_Ressources()
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim Z As Integer
    For i = 0 To UBound(Ressources())
        'Place les arbres.
        If ComReseau.Connecte Then
            X = Rnd * (Monde.Largeur * Monde.LargeurCase - AffRessources.Largeur)
            Y = Rnd * (Monde.Hauteur * Monde.HauteurCase - AffRessources.Hauteur)
        Else
            Do
                X = Rnd * (Monde.Largeur * Monde.LargeurCase - AffRessources.Largeur)
                Y = Rnd * (Monde.Hauteur * Monde.HauteurCase - AffRessources.Hauteur)
            Loop Until Obstacle(0, X, Y, AffRessources.Largeur, AffRessources.Hauteur) = 0
        End If
        Ressources(i).PositionX = X
        Ressources(i).PositionY = Y
        Ressources(i).Definir_Type_Aleatoire Parametres
    Next i
End Sub

Public Sub Placer_Joueurs()
    Dim i As Long
    Dim j As Long
    'On place le joueur.
    For i = 0 To UBound(Persos())
        'Persos(i).PositionX = Chateaux(Persos(i).NumeroFief).PositionX + CLng(RND * 128)
        'Persos(i).PositionY = Chateaux(Persos(i).NumeroFief).PositionY + CLng(RND * 128)
        Persos(i).PositionX = Chateaux(Persos(i).NumeroFief).PositionX + Chateaux(Persos(i).NumeroFief).Largeur / 2 - Persos(i).Largeur / 2 + Chateaux(Persos(i).NumeroFief).EntreePositionX
        Persos(i).PositionY = Chateaux(Persos(i).NumeroFief).PositionY + Chateaux(Persos(i).NumeroFief).Hauteur / 2 - Persos(i).Hauteur / 2 + Chateaux(Persos(i).NumeroFief).EntreePositionY
        Persos(i).DansSonChateau = True
        Persos(i).DirectionX = Persos(i).PositionX
        Persos(i).DirectionY = Persos(i).PositionY
        'If i <> Noperso Then
        '    If Persos(i).IndiceMonstre = 0 Then
        '        If RND > 0.5 Then
        '            Persos(i).Feminin = True
        '        End If
        '        Persos(i).Changer_Race_Personnage 0, Parametres, True
        '    End If
        'Else
        '    Persos(i).ControleJoueur = True
        'End If
    Next i
End Sub

Public Sub Perso_Charger_Ressources_Depart(ByVal Perso As ClsJeuPerso)
    Dim i As Integer
    If Partie.IndiceMode = 1 Then
        If Not Parametres.Race_DepartSansRessources(Perso.Race) Then
            For i = 0 To Parametres.NombreRessources - 1
                If Parametres.Ressources_Epoque(i) <= Fiefs(Perso.NumeroFief).Epoque And _
                   Parametres.Ressources_NecessairePourConstruire(i) Then
                    Perso.Definir_ressources(i) = Parametres.RessourcesDeparts_PersosRessourcesDepart(Partie.RessourcesDeDepart)
                End If
                Perso.Definir_Argent = Parametres.RessourcesDeparts_PersosArgentDepart(Partie.RessourcesDeDepart)
            Next i
        End If
    End If
End Sub

Private Function PlaceItem(longueurBegin As Long, longueurEnd As Long, largeurBegin As Long, largeurEnd As Long, nbItem As Long) As Boolean()
    Dim longueur As Integer, Largeur As Integer, nbChateau As Integer, placeIdeal As Integer
    Dim ItemPlacer() As Boolean
    ReDim ItemPlacer(longueurBegin To longueurEnd, largeurBegin To largeurEnd) As Boolean
    nbChateau = nbItem
    Dim tableauCht() As Integer
    Dim distCht() As Integer
    ReDim distCht(1 To nbChateau, 1 To nbChateau) As Integer
    tableauCht = Placement_init(longueurBegin, longueurEnd, largeurBegin, largeurEnd, nbChateau, ItemPlacer)
    DeplaceCht tableauCht, longueurBegin, longueurEnd, largeurBegin, largeurEnd, nbChateau, ItemPlacer
    PlaceItem = ItemPlacer
End Function

Private Function Placement_init(longueurBegin As Long, longueurEnd As Long, largeurBegin As Long, largeurEnd As Long, nbChateau As Integer, ByRef PlacerItem() As Boolean) As Integer()
    Dim i As Long
    Dim lig As Integer, col As Integer
    ReDim placeChateau(1 To nbChateau, 1 To 2) As Integer
    For i = 1 To nbChateau
        lig = Int(Rnd * (longueurEnd - longueurBegin) + longueurBegin)
        col = Int(Rnd * (largeurEnd - largeurBegin) + largeurBegin)
        While PlacerItem(lig, col)
            lig = Int(Rnd * (longueurEnd - longueurBegin) + longueurBegin)
            col = Int(Rnd * (largeurEnd - largeurBegin) + largeurBegin)
        Wend
        PlacerItem(lig, col) = True
        placeChateau(i, 1) = lig
        placeChateau(i, 2) = col
    Next i
    Placement_init = placeChateau
End Function

Private Sub DeplaceCht(PosCht() As Integer, longueurBegin As Long, longueurEnd As Long, largeurBegin As Long, largeurEnd As Long, nbChateau As Integer, ByRef PlacerItem() As Boolean)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim m As Long
    Dim tmp_lig As Long
    Dim tmp_col As Long
    Dim nbCaseMin As Integer
    Dim oldPos() As Integer
    ReDim oldPos(1 To nbChateau, 2) As Integer
    nbCaseMin = Int((longueurEnd + largeurEnd - longueurBegin - largeurBegin) / (nbChateau - 1)) + 5
    'FrmMoteur2D.debuggeur.Imprime nbCaseMin
    Dim Faux As Boolean, Valeur As Integer
    Dim pos(-1 To 1, -1 To 1) As Integer
    Faux = True
    While i < 5000 And Faux
        Faux = False
        For j = 1 To UBound(PosCht, 1)
            pos(-1, -1) = 0: pos(-1, 0) = 0: pos(-1, 1) = 0
            pos(0, -1) = 0: pos(0, 0) = 0: pos(0, 1) = 0
            pos(1, -1) = 0: pos(1, 0) = 0: pos(1, 1) = 0
            For k = -1 To 1
                For l = -1 To 1
                    For m = 1 To UBound(PosCht, 1)
                        If m <> j Then
                            If (Abs(PosCht(j, 1) + k - PosCht(m, 1)) + Abs(PosCht(j, 2) + l - PosCht(m, 2))) < nbCaseMin Then
                                pos(k, l) = pos(k, l) + 1
                            End If
                        End If
                    Next m
                Next l
            Next k
            Valeur = nbChateau
            If pos(0, 0) > 0 Then
                Faux = True
                For k = -1 To 1
                    For l = -1 To 1
                        If pos(k, l) <= Valeur And PosCht(j, 1) + k > longueurBegin And PosCht(j, 1) + k <= longueurEnd _
                        And PosCht(j, 2) + l > largeurBegin And PosCht(j, 2) + l <= largeurEnd Then
                            If Not PlacerItem(PosCht(j, 1) + k, PosCht(j, 2) + l) And ((pos(k, l) = Valeur And Rnd < 0.5) Or pos(k, l) <= Valeur) Then
                                Valeur = pos(k, l)
                                tmp_lig = PosCht(j, 1) + k
                                tmp_col = PosCht(j, 2) + l
                            End If
                        End If
                    Next l
                Next k
                If Not PlacerItem(tmp_lig, tmp_col) Then
                    PlacerItem(PosCht(j, 1), PosCht(j, 2)) = False
                    If Not (oldPos(j, 1) = PosCht(j, 1) + k And oldPos(j, 2) = PosCht(j, 2) + 1) Then
                        oldPos(j, 1) = PosCht(j, 1)
                        oldPos(j, 2) = PosCht(j, 2)
                        tmp_lig = Int(Rnd * (longueurEnd - longueurBegin) + longueurBegin)
                        tmp_col = Int(Rnd * (largeurEnd - largeurBegin) + largeurBegin)
                        While PlacerItem(tmp_lig, tmp_col)
                            tmp_lig = Int(Rnd * (longueurEnd - longueurBegin) + longueurBegin)
                            tmp_col = Int(Rnd * (largeurEnd - largeurBegin) + largeurBegin)
                        Wend
                        PosCht(j, 1) = tmp_lig
                        PosCht(j, 2) = tmp_col
                    Else
                        oldPos(j, 1) = PosCht(j, 1)
                        oldPos(j, 2) = PosCht(j, 2)
                        PosCht(j, 1) = tmp_lig
                        PosCht(j, 2) = tmp_col
                    End If
                    PlacerItem(PosCht(j, 1), PosCht(j, 2)) = True
                End If
            End If
        Next j
        i = i + 1
    Wend
End Sub

