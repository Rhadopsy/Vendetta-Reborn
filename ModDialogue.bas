Attribute VB_Name = "ModDialogue"
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

Const ChanceRepondreAide = 3 '4
Public Const ChanceRepondreOrdre = 0.1
Public Const IntervalSoldatsX = 24
Public Const IntervalSoldatsY = 24

Public Sub Appeler_De_lAide(ByVal IndicePerso As Long)
    'If Persos(IndicePerso).Race = 0 Then
    Dim i As Long
    Dim j As Long
    Dim Pne As Long
    Pne = Persos(IndicePerso).NumeroEquipe
    For i = 0 To UBound(Persos())
        With Persos(i)
        If .NumeroEquipe = Pne Then
        If .IA Then
        If .IA_Protecteur And .NombreSoldats > 4 Then
            'Les protecteurs envoients leurs soldats pour aider leur alliés.
            Select Case Persos(IndicePerso).Action
            Case 300:
                If Persos(IndicePerso).IA Then
                    Persos(IndicePerso).Definir_Commentaires = ""
                    Persos(IndicePerso).Definir_Commentaires = Commentaires.Message(6, Persos(IndicePerso))
                End If
                Ordonner_Chargez i, 0, Persos(IndicePerso).IndicePerso, , True
                '.Attaquer_Perso Persos(IndicePerso).IndicePerso
                .Definir_Commentaires = Commentaires.Message(10, Persos(i))
            Case 301:
                If Persos(IndicePerso).IA Then
                    Persos(IndicePerso).Definir_Commentaires = ""
                    Persos(IndicePerso).Definir_Commentaires = Commentaires.Message(7, Persos(IndicePerso))
                End If
                Ordonner_Chargez i, 1, Persos(IndicePerso).IndiceMaison, , True
                '.Attaquer_Maison Persos(IndicePerso).IndiceMaison
                .Definir_Commentaires = Commentaires.Message(10, Persos(i))
            Case 302:
                If Persos(IndicePerso).IA Then
                    Persos(IndicePerso).Definir_Commentaires = ""
                    Persos(IndicePerso).Definir_Commentaires = Commentaires.Message(8, Persos(IndicePerso))
                End If
                Ordonner_Chargez i, 2, Persos(IndicePerso).IndiceChateau, , True
                '.Attaquer_Chateau Persos(IndicePerso).IndiceChateau
                .Definir_Commentaires = Commentaires.Message(10, Persos(i))
            Case Else:
                For j = 0 To UBound(Persos())
                    If Persos(j).Action = 300 And Persos(j).IndicePerso = IndicePerso Or _
                       Persos(j).Action = 301 And Persos(j).IndiceMaison = IndicePerso And _
                       Rnd < 0.5 Then
                        If Persos(IndicePerso).IA Then
                            Persos(IndicePerso).Definir_Commentaires = ""
                            Persos(IndicePerso).Definir_Commentaires = Commentaires.Message(9, Persos(IndicePerso))
                        End If
                        '.Attaquer_Perso Persos(j).Numero
                        Ordonner_Chargez i, 0, Persos(j).Numero, , True
                        .Definir_Commentaires = Commentaires.Message(10, Persos(i))
                    j = UBound(Persos())
                    End If
                Next j
            End Select
        ElseIf Not .Invocation And Not .IA_Pacifique And Not .IA_Egoiste Then
        If .Action < 102 And .Vitalite = 1 Then
        If Rnd < ChanceRepondreAide / (Fiefs(Persos(IndicePerso).NumeroFief).NombreCitoyens + 1) Then
        'If Persos(i).IA And Not Persos(i).Invocation And _
           Persos(i).Action < 300 And _
           Persos(i).Vie > Persos(i).Definir_MaxVie * 0.5 And _
           Persos(i).Vitalite = 1 And _
           Persos(i).NumeroEquipe = Pne And _
           RND < ChanceRepondreAide / (Fiefs(Persos(IndicePerso).NumeroFief).NombreCitoyens + 1) Then
            Select Case Persos(IndicePerso).Action
            Case 300:
                If Persos(IndicePerso).IA Then
                    Persos(IndicePerso).Definir_Commentaires = ""
                    Persos(IndicePerso).Definir_Commentaires = Commentaires.Message(6, Persos(IndicePerso))
                End If
                .Attaquer_Perso Persos(IndicePerso).IndicePerso
                .Definir_Commentaires = Commentaires.Message(10, Persos(i))
            Case 301:
                If Persos(IndicePerso).IA Then
                    Persos(IndicePerso).Definir_Commentaires = ""
                    Persos(IndicePerso).Definir_Commentaires = Commentaires.Message(7, Persos(IndicePerso))
                End If
                .Attaquer_Maison Persos(IndicePerso).IndiceMaison
                .Definir_Commentaires = Commentaires.Message(10, Persos(i))
            Case 302:
                If Persos(IndicePerso).IA Then
                    Persos(IndicePerso).Definir_Commentaires = ""
                    Persos(IndicePerso).Definir_Commentaires = Commentaires.Message(8, Persos(IndicePerso))
                End If
                .Attaquer_Chateau Persos(IndicePerso).IndiceChateau
                .Definir_Commentaires = Commentaires.Message(10, Persos(i))
            Case Else:
                For j = 0 To UBound(Persos())
                    If Persos(j).Action = 300 And Persos(j).IndicePerso = IndicePerso Or _
                       Persos(j).Action = 301 And Persos(j).IndiceMaison = IndicePerso And _
                       Rnd < 0.5 Then
                        If Persos(IndicePerso).IA Then
                            Persos(IndicePerso).Definir_Commentaires = ""
                            Persos(IndicePerso).Definir_Commentaires = Commentaires.Message(9, Persos(IndicePerso))
                        End If
                        .Attaquer_Perso Persos(j).Numero
                        .Definir_Commentaires = Commentaires.Message(10, Persos(i))
                    j = UBound(Persos())
                    End If
                Next j
            End Select
        End If
        End If
        End If
        End If
        End If
        End With
    Next i
    'End If
End Sub

Public Sub Ordonner_Chargez(ByVal IndicePerso As Long, _
                            Optional ByVal TypeCible As Integer = -1, _
                            Optional ByVal IndiceCible As Long = -1, _
                            Optional ByVal Ctrl As Boolean, _
                            Optional ByVal Shift As Boolean, _
                            Optional ByVal Alt As Boolean, _
                            Optional ByVal SansCommentaires As Boolean)
    'Type Cible (0:Personnage, 1:Maison, 2:Chateau)
    Dim i As Long, j As Long
    Dim EnDanger As Boolean
    Dim IndiceEnnemi As Long
    Dim DistanceEnnemi As Long
    Dim NombreEnvoyes As Long
    Dim NombreMax As Long
    If TypeCible = -1 Then
        Select Case Persos(IndicePerso).Action
        Case 300:
            TypeCible = 0
            If IndiceCible = -1 Then IndiceCible = Persos(IndicePerso).IndicePerso
        Case 301:
            TypeCible = 1
            If IndiceCible = -1 Then IndiceCible = Persos(IndicePerso).IndiceMaison
        Case 302:
            TypeCible = 2
            If IndiceCible = -1 Then IndiceCible = Persos(IndicePerso).IndiceChateau
        End Select
    End If
    If Shift Then
        If Ctrl Then
            NombreMax = 10
        Else
            NombreMax = 5
        End If
    Else
        If Ctrl Then
            NombreMax = 1
        Else
            NombreMax = Persos(IndicePerso).NombreSoldats
        End If
    End If
    
    'Recherche l'indice d'un ennemi qui attaque le chef ou sa maison.
'    IndiceEnnemi = -1
'    DistanceEnnemi = Monde.Longueur_Diagonale
'    For i = 0 To UBound(Persos())
'        With Persos(i)
'        If .Action = 300 And .IndicePerso = IndicePerso Or _
'           .Action = 301 And .IndiceMaison = IndicePerso Then
'            'Regarde si cet adversaire est plus proche que le précédent.
'            If IndiceEnnemi >= 0 Then
'                If DistanceEnnemi > Sqr((Persos(IndicePerso).PositionX - .PositionX) ^ 2 + _
'                                        (Persos(IndicePerso).PositionY - .PositionY) ^ 2) Then
'                    DistanceEnnemi = Sqr((Persos(IndicePerso).PositionX - .PositionX) ^ 2 + _
'                                         (Persos(IndicePerso).PositionY - .PositionY) ^ 2)
'                    IndiceEnnemi = i
'                End If
'            Else
'                IndiceEnnemi = i
'                EnDanger = True
'            End If
'        End If
'        End With
'    Next i
    'If TypeCible = -1 And Not EnDanger Then
    '    Ordonner_Formation IndicePerso
    'End If
    
    For i = 0 To UBound(Persos())
        With Persos(i)
        If .NumeroChef = IndicePerso And _
           .Vivant And _
           Not .EnCombat And _
           Not .Incommandable And .Attitude = 0 And .IA_Ordre = 0 And _
           .ChoisirIA = False And .Action <> 2 And .Action <> 3 Then    'Les Serviteurs obéissent à leur chef.
            Select Case TypeCible
            Case 0:
                .Attaquer_Perso IndiceCible
                'Cmd_Perso_AttaquerPerso i, IndiceCible
                .IA_Ordre = 1
                'Cmd_Perso_DefinirIAOrdre i, 1
                .IA_Cible = IndiceCible
                'Cmd_Perso_DefinirIACible i, IndiceCible
                If Rnd < ChanceRepondreOrdre And Not SansCommentaires Then
                    .Definir_Commentaires = Commentaires.Message(31, Persos(i))
                    'Cmd_Perso_DefinirCommentaires i, Commentaires.Message(31, Persos(i))
                End If
            Case 1:
                .Attaquer_Maison IndiceCible
                'Cmd_Perso_AttaquerMaison i, IndiceCible
                .IA_Ordre = 2
                'Cmd_Perso_DefinirIAOrdre i, 2
                .IA_Cible = IndiceCible
                'Cmd_Perso_DefinirIACible i, IndiceCible
                If Rnd < ChanceRepondreOrdre And Not SansCommentaires Then
                    .Definir_Commentaires = Commentaires.Message(31, Persos(i))
                    'Cmd_Perso_DefinirCommentaires i, Commentaires.Message(31, Persos(i))
                End If
            Case 2:
                .Attaquer_Chateau IndiceCible
                'Cmd_Perso_AttaquerChateau i, IndiceCible
                .IA_Ordre = 3
                'Cmd_Perso_DefinirIAOrdre i, 3
                .IA_Cible = IndiceCible
                'Cmd_Perso_DefinirIACible i, IndiceCible
                If Rnd < ChanceRepondreOrdre And Not SansCommentaires Then
                    .Definir_Commentaires = Commentaires.Message(31, Persos(i))
                    'Cmd_Perso_DefinirCommentaires i, Commentaires.Message(31, Persos(i))
                End If
            Case Else:
                'If EnDanger Then
                    'Fonce sur l'ennemi le plus proche.
                '    .Attaquer_Perso Persos(IndiceEnnemi).Numero
                '    If RND < ChanceRepondreOrdre Then .Definir_Commentaires = Commentaires.Message(31, Persos(i))
                'Else
                    'Si le chef n'est pas en danger, l'unité revient en formation.
                    .Aller_A Persos(.NumeroChef).DirectionX + .ChefDecalageX, _
                             Persos(.NumeroChef).DirectionY + .ChefDecalageY
                    'Cmd_Perso_Deplacer i, Persos(.NumeroChef).DirectionX + .ChefDecalageX, _
                                          Persos(.NumeroChef).DirectionY + .ChefDecalageY
                    .ChefPositionX = .PositionX
                    .ChefPositionY = .PositionY
                    'Cmd_Perso_DefinirChefPosition i, .PositionX, .PositionY
                    If Rnd < ChanceRepondreOrdre And Not SansCommentaires Then
                        .Definir_Commentaires = Commentaires.Message(31, Persos(i))
                        'Cmd_Perso_DefinirCommentaires i, Commentaires.Message(31, Persos(i))
                    End If
                'End If
            End Select
            NombreEnvoyes = NombreEnvoyes + 1
            If NombreEnvoyes >= NombreMax Then
                i = UBound(Persos())
            End If
        End If
        End With
    Next i
    
    With Persos(IndicePerso)
    If NombreEnvoyes > 0 And Not SansCommentaires Then
        Select Case TypeCible
        Case 0:
            If NombreEnvoyes = 1 Then
                .Definir_Commentaires = ""
                'Cmd_Perso_DefinirCommentaires IndicePerso, ""
                .Definir_Commentaires = Commentaires.Message(25, Persos(IndicePerso))
                'Cmd_Perso_DefinirCommentaires IndicePerso, Commentaires.Message(25, Persos(IndicePerso))
            Else
                .Definir_Commentaires = ""
                'Cmd_Perso_DefinirCommentaires IndicePerso, ""
                .Definir_Commentaires = Commentaires.Message(26, Persos(IndicePerso))
                'Cmd_Perso_DefinirCommentaires IndicePerso, Commentaires.Message(26, Persos(IndicePerso))
            End If
        Case 1:
            If NombreEnvoyes = 1 Then
                .Definir_Commentaires = ""
                'Cmd_Perso_DefinirCommentaires IndicePerso, ""
                .Definir_Commentaires = Commentaires.Message(27, Persos(IndicePerso))
                'Cmd_Perso_DefinirCommentaires IndicePerso, Commentaires.Message(27, Persos(IndicePerso))
            Else
                .Definir_Commentaires = ""
                'Cmd_Perso_DefinirCommentaires IndicePerso, ""
                .Definir_Commentaires = Commentaires.Message(28, Persos(IndicePerso))
                'Cmd_Perso_DefinirCommentaires IndicePerso, Commentaires.Message(28, Persos(IndicePerso))
            End If
        Case 2:
            If NombreEnvoyes = 1 Then
                .Definir_Commentaires = ""
                'Cmd_Perso_DefinirCommentaires IndicePerso, ""
                .Definir_Commentaires = Commentaires.Message(29, Persos(IndicePerso))
                'Cmd_Perso_DefinirCommentaires IndicePerso, Commentaires.Message(29, Persos(IndicePerso))
            Else
                .Definir_Commentaires = ""
                'Cmd_Perso_DefinirCommentaires IndicePerso, ""
                .Definir_Commentaires = Commentaires.Message(30, Persos(IndicePerso))
                'Cmd_Perso_DefinirCommentaires IndicePerso, Commentaires.Message(30, Persos(IndicePerso))
            End If
        Case Else:
            'If EnDanger Then
                '.Definir_Commentaires = ""
                '.Definir_Commentaires = Commentaires.Message(24, Persos(IndicePerso))
                'Cmd_Perso_DefinirCommentaires IndicePerso, Commentaires.Message(24, Persos(IndicePerso))
            'Else
                'If .EnDeplacement Then
                    .Definir_Commentaires = ""
                    'Cmd_Perso_DefinirCommentaires IndicePerso, ""
                    .Definir_Commentaires = Commentaires.Message(23, Persos(IndicePerso))
                    'Cmd_Perso_DefinirCommentaires IndicePerso, Commentaires.Message(23, Persos(IndicePerso))
                'Else
                    '.Definir_Commentaires = ""
                    '.Definir_Commentaires = Commentaires.Message(22, Persos(IndicePerso))
                    'Cmd_Perso_DefinirCommentaires IndicePerso, Commentaires.Message(22, Persos(IndicePerso))
                'End If
            'End If
        End Select
    'Else
        '.Definir_Commentaires = ""
        '.Definir_Commentaires = Commentaires.Message(32, Persos(IndicePerso))
        'Cmd_Perso_DefinirCommentaires IndicePerso, Commentaires.Message(32, Persos(IndicePerso))
    End If
    End With
End Sub

Public Sub Ordonner_Arret(ByVal IndicePerso As Long, _
                          Optional ByVal ArmeeSelectionnee As Boolean, _
                          Optional ByVal SoldatsSelectionnes As Collection)
    Dim i As Long
    Dim Commentaire As Boolean
    If ArmeeSelectionnee Then
        For i = 1 To SoldatsSelectionnes.Count
            With Persos(SoldatsSelectionnes.Item(i))
            '.Arreter
            Cmd_Perso_Arreter SoldatsSelectionnes.Item(i)
            '.IA_Ordre = 0
            Cmd_Perso_DefinirIAOrdre SoldatsSelectionnes.Item(i), 0
            End With
        Next i
        Commentaire = True
    Else
        For i = 0 To UBound(Persos())
            If Persos(i).NumeroChef = IndicePerso And _
               Persos(i).Vivant And _
               Not Persos(i).Incommandable And Persos(i).ChoisirIA = False Then
                'Persos(i).Arreter
                Cmd_Perso_Arreter i
                Cmd_Perso_DefinirIAOrdre i, 0
                Commentaire = True
            End If
        Next i
    End If
    'Le personnage fait un commentaire, s'il avait au moins l'un de ses suivants qui obéit.
    If Commentaire Then
        'Persos(IndicePerso).Definir_Commentaires = ""
        'Persos(IndicePerso).Definir_Commentaires = Commentaires.Message(21, Persos(IndicePerso))
        Cmd_Perso_DefinirCommentaires Persos(IndicePerso).Numero, ""
        Cmd_Perso_DefinirCommentaires Persos(IndicePerso).Numero, Commentaires.Message(21, Persos(IndicePerso))
    End If
End Sub

Public Sub Ordonner_Formation(ByVal IndicePerso As Long, _
                              Optional ByVal ChangerFormation As Boolean, _
                              Optional ByVal SansCommentaires As Boolean)
    'Type de formations:
    '0 :Bataillon en haut.
    '1 :Bataillon en bas.
    '2 :Bataillon à gauche.
    '3 :Bataillon à droite.
    Dim i As Long, j As Long, X As Long, Y As Long
    Dim Rangee As Long 'Nombre d'hommes par rangée.
    
    With Persos(IndicePerso)
    If ChangerFormation Then .Changer_Formation
    'Rangee = Int(Sqr(Persos(IndicePerso).NombreSoldats)) * 1.5
    Rangee = Int(Sqr(.NombreSoldats)) * 1.5
    If Rangee = 0 Then Rangee = 1
    Select Case .Formation
        Case 0: Rangee = Rangee / 1.5
                X = IntervalSoldatsX * Rangee / 2 - (IntervalSoldatsX / 2)
                Y = IntervalSoldatsY * Rangee / 2 + (IntervalSoldatsY / 2)
        Case 1: X = IntervalSoldatsX * Rangee / 2 - (IntervalSoldatsX / 2)
                Y = 0
        Case 2: X = IntervalSoldatsX * Rangee / 2 - (IntervalSoldatsX / 2)
                Y = 0
        Case 3: X = 0
                Y = IntervalSoldatsY * Rangee / 2 - (IntervalSoldatsY / 2)
        Case 4: X = 0
                Y = IntervalSoldatsY * Rangee / 2 - (IntervalSoldatsY / 2)
    End Select
    End With
    For i = 0 To UBound(Persos())
        With Persos(i)
        If .NumeroChef = IndicePerso And _
           .Vivant And _
           .IA_Ordre = 0 And _
           .Attitude = 0 And _
           Not .Incommandable And .ChoisirIA = False Then
            Select Case Persos(IndicePerso).Formation
            Case 0, 1:
                If j Mod Rangee = 0 Then
                    Y = Y - IntervalSoldatsY
                    X = X - IntervalSoldatsX * (Rangee - 1)
                Else
                    X = X + IntervalSoldatsX
                End If
            Case 2:
                If j Mod Rangee = 0 Then
                    Y = Y + IntervalSoldatsY
                    X = X - IntervalSoldatsX * (Rangee - 1)
                Else
                    X = X + IntervalSoldatsX
                End If
            Case 3:
                If j Mod Rangee = 0 Then
                    X = X - IntervalSoldatsX
                    Y = Y - IntervalSoldatsY * (Rangee - 1)
                Else
                    Y = Y + IntervalSoldatsY
                End If
            Case 4:
                If j Mod Rangee = 0 Then
                    X = X + IntervalSoldatsX
                    Y = Y - IntervalSoldatsY * (Rangee - 1)
                Else
                    Y = Y + IntervalSoldatsY
                End If
            End Select
            'Si l'invocation est superposée au maître, on le décale.
            If X = 0 And Y = 0 Then
                i = i - 1
            Else
                .ChefDecalageX = X
                .ChefDecalageY = Y
                'Va a son emplacement.
                If Not Persos(.NumeroChef).EnCombat Then
                    .Aller_A Persos(.NumeroChef).DirectionX + .ChefDecalageX, _
                                     Persos(.NumeroChef).DirectionY + .ChefDecalageY
                    .ChefPositionX = .DirectionX
                    .ChefPositionY = .DirectionY
                End If
            End If
            j = j + 1
        End If
        End With
    Next i
    
    If Not SansCommentaires Then
        With Persos(IndicePerso)
        If .EnDeplacement Then
            .Definir_Commentaires = Commentaires.Message(23, Persos(IndicePerso))
            'Cmd_Perso_DefinirCommentaires Persos(IndicePerso).Numero, Commentaires.Message(23, Persos(IndicePerso))
        Else
            .Definir_Commentaires = Commentaires.Message(22, Persos(IndicePerso))
            'Cmd_Perso_DefinirCommentaires Persos(IndicePerso).Numero, Commentaires.Message(22, Persos(IndicePerso))
        End If
        End With
    End If
End Sub

'Private Function Nombre_Soldats(ByVal IndicePerso As Long, _
'                                Optional ByVal Attitude As Integer) As Long
'    Dim i As Long, Res As Long
'    For i = 0 To UBound(Persos())
'        With Persos(i)
'        If .NumeroChef = IndicePerso And _
'           .Vivant And _
'           .Attitude = Attitude And _
'           Not .Incommandable Then
'            Res = Res + 1
'        End If
'        End With
'    Next i
'    Nombre_Soldats = Res
'End Function

Public Sub Ordonner_Activer_IA(ByVal IndicePerso As Long, _
                               Optional ByVal IA As Boolean = True)
    Dim i As Long
    If Persos(IndicePerso).NombreSoldats > 0 Then
        For i = 0 To UBound(Persos())
            With Persos(i)
            If Persos(i).NumeroChef = IndicePerso And _
               Persos(i).Vivant Then
                If .IA <> IA Then Cmd_Perso_Definir_IA i, IA
                '.IA = IA
            End If
            End With
        Next i
    End If
End Sub

Public Sub Ordonner_Deplacement(ByVal Perso As ClsJeuPerso, _
                                ByVal Soldats As Collection, _
                                ByVal X As Long, _
                                ByVal Y As Long)
    Dim i As Long
    For i = 1 To Soldats.Count
        With Persos(Soldats.Item(i))
        '.Aller_A X + .ChefDecalageX, Y + .ChefDecalageY
        Cmd_Perso_Deplacer Soldats.Item(i), X + .ChefDecalageX, Y + .ChefDecalageY
        '.ChefPositionX = X + .ChefDecalageX
        '.ChefPositionY = Y + .ChefDecalageY
        Cmd_Perso_DefinirChefPosition Soldats.Item(i), X + .ChefDecalageX, Y + .ChefDecalageY
        'Persos(Soldats.Item(i)).IA_Ordre = 0
        Cmd_Perso_DefinirIAOrdre Soldats.Item(i), 0
        If Rnd < ChanceRepondreOrdre And Soldats.Item(i) <> Perso.Numero Then
            'Persos(Soldats.Item(i)).Definir_Commentaires = _
                Commentaires.Message(31, Persos(Soldats.Item(i)))
            Cmd_Perso_DefinirCommentaires Soldats.Item(i), Commentaires.Message(31, Persos(Soldats.Item(i)))
        End If
        End With
    Next i
    'Perso.Definir_Commentaires = Commentaires.Message(23, Perso)
    Cmd_Perso_DefinirCommentaires Perso.Numero, Commentaires.Message(23, Perso)
End Sub

Public Sub Ordonner_Attaquer_Perso(ByVal Perso As ClsJeuPerso, _
                                   ByVal Soldats As Collection, _
                                   ByVal Cible As Long)
    Dim i As Long
    If Cible >= 0 Then
        For i = 1 To Soldats.Count
            'Persos(Soldats.Item(i)).Attaquer_Perso Cible
            Cmd_Perso_AttaquerPerso Soldats.Item(i), Cible
            'Persos(Soldats.Item(i)).IA_Ordre = 1
            Cmd_Perso_DefinirIAOrdre Soldats.Item(i), 1
            'Persos(Soldats.Item(i)).IA_Cible = Cible
            Cmd_Perso_DefinirIACible Soldats.Item(i), Cible
            If Rnd < ChanceRepondreOrdre And Soldats.Item(i) <> Perso.Numero Then
                'Persos(Soldats.Item(i)).Definir_Commentaires = _
                    Commentaires.Message(31, Persos(Soldats.Item(i)))
                Cmd_Perso_DefinirCommentaires Soldats.Item(i), Commentaires.Message(31, Persos(Soldats.Item(i)))
            End If
        Next i
        If Soldats.Count = 1 Then
            'Perso.Definir_Commentaires = Commentaires.Message(25, Perso)
            Cmd_Perso_DefinirCommentaires Perso.Numero, Commentaires.Message(25, Perso)
        Else
            'Perso.Definir_Commentaires = Commentaires.Message(26, Perso)
            Cmd_Perso_DefinirCommentaires Perso.Numero, Commentaires.Message(26, Perso)
        End If
    End If
End Sub

Public Sub Ordonner_Attaquer_Maison(ByVal Perso As ClsJeuPerso, _
                                    ByVal Soldats As Collection, _
                                    ByVal Cible As Long)
    Dim i As Long
    If Cible >= 0 Then
        For i = 1 To Soldats.Count
            'Persos(Soldats.Item(i)).Attaquer_Maison Cible
            Cmd_Perso_AttaquerMaison Soldats.Item(i), Cible
            'Persos(Soldats.Item(i)).IA_Ordre = 2
            Cmd_Perso_DefinirIAOrdre Soldats.Item(i), 2
            'Persos(Soldats.Item(i)).IA_Cible = Cible
            Cmd_Perso_DefinirIACible Soldats.Item(i), Cible
            If Rnd < ChanceRepondreOrdre And Soldats.Item(i) <> Perso.Numero Then
                'Persos(Soldats.Item(i)).Definir_Commentaires = _
                    Commentaires.Message(31, Persos(Soldats.Item(i)))
                Cmd_Perso_DefinirCommentaires Soldats.Item(i), Commentaires.Message(31, Persos(Soldats.Item(i)))
            End If
        Next i
        If Soldats.Count = 1 Then
            'Perso.Definir_Commentaires = Commentaires.Message(27, Perso)
            Cmd_Perso_DefinirCommentaires Perso.Numero, Commentaires.Message(27, Perso)
        Else
            'Perso.Definir_Commentaires = Commentaires.Message(28, Perso)
            Cmd_Perso_DefinirCommentaires Perso.Numero, Commentaires.Message(28, Perso)
        End If
    End If
End Sub

Public Sub Ordonner_Attaquer_Chateau(ByVal Perso As ClsJeuPerso, _
                                     ByVal Soldats As Collection, _
                                     ByVal Cible As Long)
    Dim i As Long
    If Cible >= 0 Then
        For i = 1 To Soldats.Count
            'Persos(Soldats.Item(i)).Attaquer_Chateau Cible
            Cmd_Perso_AttaquerChateau Soldats.Item(i), Cible
            'Persos(Soldats.Item(i)).IA_Ordre = 3
            Cmd_Perso_DefinirIAOrdre Soldats.Item(i), 3
            'Persos(Soldats.Item(i)).IA_Cible = Cible
            Cmd_Perso_DefinirIACible Soldats.Item(i), Cible
            If Rnd < ChanceRepondreOrdre And Soldats.Item(i) <> Perso.Numero Then
                'Persos(Soldats.Item(i)).Definir_Commentaires = _
                    Commentaires.Message(31, Persos(Soldats.Item(i)))
                Cmd_Perso_DefinirCommentaires Soldats.Item(i), Commentaires.Message(31, Persos(Soldats.Item(i)))
            End If
        Next i
        If Soldats.Count = 1 Then
            'Perso.Definir_Commentaires = Commentaires.Message(29, Perso)
            Cmd_Perso_DefinirCommentaires Perso.Numero, Commentaires.Message(29, Perso)
        Else
            'Perso.Definir_Commentaires = Commentaires.Message(30, Perso)
            Cmd_Perso_DefinirCommentaires Perso.Numero, Commentaires.Message(30, Perso)
        End If
    End If
End Sub

Public Sub Ordonner_Garder_Perso(ByVal Perso As ClsJeuPerso, _
                                 ByVal Soldats As Collection, _
                                 ByVal Cible As Long)
    Dim i As Long
    If Cible >= 0 Then
        For i = 1 To Soldats.Count
            With Persos(Soldats.Item(i))
            Cmd_Perso_DefinirIAOrdre Soldats.Item(i), 4
            Cmd_Perso_DefinirIACible Soldats.Item(i), Cible
            Cmd_Perso_Deplacer Soldats.Item(i), _
                               Persos(Cible).PositionX + Persos(Cible).Largeur / 2 - .Largeur / 2, _
                               Persos(Cible).PositionY + Persos(Cible).Largeur / 2 - .Hauteur / 2
            End With
        Next i
        If Soldats.Count = 1 Then
            Cmd_Perso_DefinirCommentaires Perso.Numero, Commentaires.Message(38, Perso)
            'Perso.Definir_Commentaires = Commentaires.Message(38, Perso)
        Else
            Cmd_Perso_DefinirCommentaires Perso.Numero, Commentaires.Message(39, Perso)
            'Perso.Definir_Commentaires = Commentaires.Message(39, Perso)
        End If
    End If
End Sub

'Public Sub Ordonner_Garder_Maison(ByVal Perso As ClsJeuPerso, _
'                                  ByVal Soldats As Collection, _
'                                  ByVal Cible As Long)
'    Dim i As Long
'    If Cible >= 0 Then
'        For i = 1 To Soldats.Count
'            If Soldats.Item(i) = Perso.Numero Then
'                With Perso
'                If .Numero = Cible Then
'                    .Rentrer_Maison Maisons(Cible)
'                Else
'                    .Aller_Maison Maisons(Cible)
'                End If
'                End With
'            Else
'                With Persos(Soldats.Item(i))
'                .IA_Ordre = 1
'                .IndiceMaison = Cible
'                .Aller_A Maisons(Cible).PositionX + Maisons(Cible).Largeur / 2 - Persos(Soldats.Item(i)).Largeur / 2, _
'                         Maisons(Cible).PositionY + Maisons(Cible).Largeur / 2 - Persos(Soldats.Item(i)).Hauteur / 2
'                'IA_Garder_Batiment Soldats.Item(i), False
'                End With
'            End If
'        Next i
'        If Soldats.Count = 1 Then
'            If Cible = Perso.Numero Then
'                Perso.Definir_Commentaires = Commentaires.Message(36, Perso)
'            Else
'                Perso.Definir_Commentaires = Commentaires.Message(38, Perso)
'            End If
'        Else
'            If Cible = Perso.Numero Then
'                Perso.Definir_Commentaires = Commentaires.Message(37, Perso)
'            Else
'                Perso.Definir_Commentaires = Commentaires.Message(39, Perso)
'            End If
'        End If
'    End If
'End Sub

'Public Sub Ordonner_Garder_Chateau(ByVal Perso As ClsJeuPerso, _
'                                   ByVal Soldats As Collection, _
'                                   ByVal Cible As Long)
'    Dim i As Long
'    If Cible >= 0 Then
'        For i = 1 To Soldats.Count
'            If Soldats.Item(i) = Perso.Numero Then
'                If Perso.NumeroFief = Cible Then
'                    Perso.Rentrer_Chateau Chateaux(Cible)
'                Else
'                    Perso.Aller_Chateau Chateaux(Cible), Cible
'                End If
'            Else
'                With Persos(Soldats.Item(i))
'                Persos(Soldats.Item(i)).IA_Ordre = 2
'                Persos(Soldats.Item(i)).IndiceChateau = Cible
'                .Aller_A Chateaux(Cible).PositionX + Chateaux(Cible).Largeur / 2 - Persos(Soldats.Item(i)).Largeur / 2, _
'                         Chateaux(Cible).PositionY + Chateaux(Cible).Largeur / 2 - Persos(Soldats.Item(i)).Hauteur / 2
'                End With
'                'IA_Garder_Batiment i, True
'            End If
'        Next i
'        If Soldats.Count = 1 Then
'            If Cible = Perso.NumeroFief Then
'                Perso.Definir_Commentaires = Commentaires.Message(40, Perso)
'            Else
'                Perso.Definir_Commentaires = Commentaires.Message(42, Perso)
'            End If
'        Else
'            If Cible = Perso.NumeroFief Then
'                Perso.Definir_Commentaires = Commentaires.Message(41, Perso)
'            Else
'                Perso.Definir_Commentaires = Commentaires.Message(43, Perso)
'            End If
'        End If
'    End If
'End Sub

