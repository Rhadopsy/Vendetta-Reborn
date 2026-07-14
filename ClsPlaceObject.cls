VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPlaceObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class is aimed to place object in a matrix with suffisant distance between them."
' This class is aimed to place object in a matrix with suffisant distance between them.
' Copyright (C) 2005  LEVEQUE Franck
'
' This library is free software; you can redistribute it and/or
' modify it under the terms of the GNU Lesser General Public
' License as published by the Free Software Foundation; either
' version 2.1 of the License, or (at your option) any later version.
'
' This library is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' Lesser General Public License for more details.
'
' You should have received a copy of the GNU Lesser General Public
' License along with this library; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

Private NbLine As Long
Attribute NbLine.VB_VarDescription = "Line size of matrix"
Private NbColumn As Long
Attribute NbColumn.VB_VarDescription = "ColumnSize of matrix"
Private CalculatedObjects As Collection
Private PredefinedObject As Collection
Private NbOfObjectToCalculate As Long
Private IsPredefinedObjectAWall As Boolean
Public DebugCycleNumber As Long
Public MinSpaceBetweenTwoObject As Long

Public Function GetLastObjectCollection() As Collection
Attribute GetLastObjectCollection.VB_Description = "Return the last object collection calculated."
    Set GetLastObjectCollection = CalculatedObjects
End Function

Public Function CalculateObjectsPlace() As Collection
Attribute CalculateObjectsPlace.VB_Description = "Calculate objets collection according to the class parameters ."
        InitOjectPlace
        MoveObjects
End Function

Public Sub SetPlacedObjectAsWall(ByVal ObjectsAreWall As Boolean)
Attribute SetPlacedObjectAsWall.VB_Description = "This methods determines if predefined object are treated as wall (objects just needs to avoid to be on the ""wall"") or not (predefined object doesn't move but takes part in calculation of other objects place)"
    IsPredefinedObjectAWall = ObjectsAreWall
End Sub

Public Sub SetNumberOfObject(ByVal ObjectToplace As Long)
Attribute SetNumberOfObject.VB_Description = "This method determine the number of object to place on the matrix."
    NbOfObjectToCalculate = ObjectToplace
    If NbColumn > 0 And NbLine > 0 Then
            MinSpaceBetweenTwoObject = Int((NbLine + NbColumn) / (Sqr(NbOfObjectToCalculate - 2) + 1)) - (Int(NbOfObjectToCalculate / Sqr(NbOfObjectToCalculate)) + 2)
    End If
End Sub

Public Sub SetPlacedObjectCollection(PlacedObject As Collection)
Attribute SetPlacedObjectCollection.VB_Description = "This method determines the objects already placed in the matrix (Predefined)"
    Set PredefinedObject = PlacedObject
End Sub

Public Sub SetBoardSize(ByVal LineNumber As Long, ByVal ColumnNumber As Long)
Attribute SetBoardSize.VB_Description = "This methods determine the matrix size"
    NbColumn = ColumnNumber
    NbLine = LineNumber
    If NbOfObjectToCalculate > 0 Then
        MinSpaceBetweenTwoObject = Int((NbLine + NbColumn) / (Sqr(NbOfObjectToCalculate - 2) + 1)) - (Int(NbOfObjectToCalculate / Sqr(NbOfObjectToCalculate)) + 2)
    End If
End Sub
Private Function GetBoxSize() As BoxSize
    Dim CurrentProblemSize As Long
    CurrentProblemSize = NbOfObjectToCalculate + PredefinedObject.Count
    GetBoxSize.LineSize = NbLine
    GetBoxSize.ColumnSize = NbColumn
    While CurrentProblemSize > 1
        CurrentProblemSize = Int(CurrentProblemSize / 2 + 0.5)
        If GetBoxSize.LineSize > GetBoxSize.ColumnSize Then
            GetBoxSize.LineSize = Int(GetBoxSize.LineSize / 2)
        Else
            GetBoxSize.ColumnSize = Int(GetBoxSize.ColumnSize / 2)
        End If
    Wend
End Function

Private Sub InitOjectPlace()
    Dim RemainingObjectToplace As Long, c As ClsMatrixPoint
    Dim Box As BoxSize, CurPlace As BoxPlace, BoxIsOccupied As Boolean
    Box = GetBoxSize()
    'Debug.Print "BoxSize(" & Box.LineSize & "," & Box.ColumnSize & ")"
    CurPlace.TopLine = 1
    CurPlace.LeftColumn = 1
    RemainingObjectToplace = NbOfObjectToCalculate
    Set CalculatedObjects = New Collection
    While RemainingObjectToplace > 0
        BoxIsOccupied = False
        For Each c In PredefinedObject
            If c.LinePos >= CurPlace.TopLine And c.LinePos <= CurPlace.TopLine + Box.LineSize _
            And c.ColumnPos >= CurPlace.LeftColumn And c.ColumnPos <= CurPlace.LeftColumn + Box.ColumnSize Then
                BoxIsOccupied = True
                Exit For
            End If
        Next
        If Not BoxIsOccupied Then
            For Each c In CalculatedObjects
                If c.LinePos >= CurPlace.TopLine And c.LinePos <= CurPlace.TopLine + Box.LineSize _
                And c.ColumnPos >= CurPlace.LeftColumn And c.ColumnPos <= CurPlace.LeftColumn + Box.ColumnSize Then
                    BoxIsOccupied = True
                    Exit For
                End If
            Next
        End If
        If Not BoxIsOccupied Then
            RemainingObjectToplace = RemainingObjectToplace - 1
            Set c = New ClsMatrixPoint
            c.LinePos = Int(Rnd * Box.LineSize + CurPlace.TopLine)
            c.ColumnPos = Int(Rnd * Box.ColumnSize + CurPlace.LeftColumn)
            CalculatedObjects.Add c
        End If
        If Not (CurPlace.LeftColumn + (2 * Box.ColumnSize - 1) > NbColumn) Then
            CurPlace.LeftColumn = CurPlace.LeftColumn + Box.ColumnSize
        Else
            CurPlace.LeftColumn = 1
            CurPlace.TopLine = CurPlace.TopLine + Box.LineSize
        End If
    Wend
End Sub

Private Function MoveObjects(Optional ByVal DebugCycle As Long = -1) As Boolean
i = 0
Dim nbCaseMin As Integer
Dim TmpValCase As Long, tmp_lig As Long, tmp_col As Long
nbCaseMin = MinSpaceBetweenTwoObject
Dim Faux As Boolean, valeur As Long, j As ClsMatrixPoint, m As ClsMatrixPoint
Dim pos(-1 To 1, -1 To 1) As Long
Faux = True
While i < 10000 And Faux
    Faux = False
    For Each j In CalculatedObjects
        ' Initialisation des alentours de l'objet
        For k = -1 To 1
            For l = -1 To 1
                pos(k, l) = NbColumn * NbLine
            Next
        Next
        
        ' Vérification des alentours
        For k = -1 To 1
            For l = -1 To 1
                For Each m In CalculatedObjects
                    If m.ColumnPos <> j.ColumnPos Or m.LinePos <> j.LinePos Then
                        TmpValCase = Abs(j.LinePos + k - m.LinePos) + Abs(j.ColumnPos + l - m.ColumnPos)
                        If TmpValCase < nbCaseMin And TmpValCase < pos(k, l) Then
                            pos(k, l) = TmpValCase
                        End If
                    End If
                Next
            Next
        Next
        ' Initialisation de la migration
        If pos(0, 0) < nbCaseMin Then
            Faux = True
            valeur = pos(0, 0)
            tmp_lig = j.LinePos
            tmp_col = j.ColumnPos
            For k = -1 To 1
                For l = -1 To 1
                    If pos(k, l) >= valeur And pos(k, l) > 0 _
                    And j.LinePos + k > 0 And j.LinePos + k <= NbLine _
                    And j.ColumnPos + l > 0 And j.ColumnPos + l <= NbColumn Then
                        If (pos(k, l) = valeur And Rnd < 0.5) Or pos(k, l) > valeur Then
                            valeur = pos(k, l)
                            tmp_lig = j.LinePos + k
                            tmp_col = j.ColumnPos + l
                        End If
                    End If
                Next
            Next
            j.LinePos = tmp_lig
            j.ColumnPos = tmp_col
        End If
    Next
    i = i + 1
    If DebugCycle > -1 Then
        MoveObjects = Faux
        Faux = False
    End If
Wend
End Function
Public Sub initDebugMode()
    DebugCycleNumber = 0
End Sub
Public Function DebugPlaceObject() As Boolean
    Select Case DebugCycleNumber
        Case 0:
            InitOjectPlace
            DebugPlaceObject = True
        Case Else:
            DebugPlaceObject = MoveObjects(DebugCycleNumber)
    End Select
    DebugCycleNumber = DebugCycleNumber + 1
End Function

Private Sub Class_Initialize()
    Set CalculatedObjects = New Collection
    Set PredefinedObject = New Collection
End Sub
