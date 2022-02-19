VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'// ----------------------------------------------------------------------------------------
'//                              COPYRIGHT NOTICE
'// ----------------------------------------------------------------------------------------
'//
'// The Source Code Store LLC
'// ACTIVEGANTT SCHEDULER COMPONENT FOR VISUAL BASIC 6
'// ACTIVEX COMPONENT
'// Copyright (c) 2002-2004 The Source Code Store LLC
'//
'// All Rights Reserved. No parts of this file may be reproduced or transmitted in any
'// form or by any means without the written permission of the author.
'// ----------------------------------------------------------------------------------------
Option Explicit

'// Parent Control Pointer
Private mp_oControl As ActiveGanttVBCtl
'// Object Member Variables
Private mp_oCollection As clsCollectionBase
Private mp_lTopOffset As Long

'// ---------------------------------------------------------------------------------------------------------------------
'// Initialization & Destruction
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oCollection = New clsCollectionBase
    mp_oCollection.Initialize Value, "Row"
End Sub

Private Sub Class_Terminate()
    Set mp_oCollection = Nothing
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Properties
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Count() As Long
    Count = mp_oCollection.m_lCount
End Property

Public Property Get Item(ByVal Index As String) As clsRow
    Set Item = mp_oCollection.m_oItem(Index, 50228, 50229, 50230, 50231)
End Property

Friend Property Get oCollection() As clsCollectionBase
    Set oCollection = mp_oCollection
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub Add(ByVal Key As String, Optional ByVal Caption As String = "", Optional ByVal MergeCells As Boolean = False, Optional ByVal Container As Boolean = True, Optional ByVal StyleIndex As String = "0")
    mp_oCollection.AddMode = True
    Dim oRow As New clsRow
    Dim lIndex As Long
    oRow.Initialize mp_oControl
    oRow.Key = Key
    oRow.Caption = Caption
    oRow.MergeCells = MergeCells
    oRow.Container = Container
    oRow.StyleIndex = StyleIndex
    mp_oCollection.m_Add oRow, Key, 50232, 50233, True, 50386
    For lIndex = 1 To mp_oControl.Columns.Count
        Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_oControl.Rows.Count)
        oRow.Cells.Add
    Next lIndex
    mp_oControl.f_oVScrollBar.Max = Count
End Sub

Public Sub Clear()
    mp_oControl.SelectedRowIndex = 0
    mp_oControl.Tasks.Clear
    mp_oControl.Milestones.Clear
    mp_oCollection.m_Clear
    mp_oControl.f_oVScrollBar.Max = Count
End Sub

Public Sub Remove(ByVal Index As String)
    Dim sRIndex As String
    Dim sRKey As String
    mp_oCollection.m_GetKeyAndIndex Index, sRKey, sRIndex
    mp_oControl.Tasks.oCollection.m_CollectionRemoveWhere "RowKey", sRKey, sRIndex
    mp_oControl.Milestones.oCollection.m_CollectionRemoveWhere "RowKey", sRKey, sRIndex
    mp_oControl.SelectedRowIndex = 0
    mp_oCollection.m_Remove Index, 50234, 50235, 50236, 50237
    mp_oControl.f_oVScrollBar.Max = Count
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Custom Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub MoveRow(ByVal OriginRowIndex As Long, ByVal DestRowIndex As Long)
    If OriginRowIndex < 1 Or OriginRowIndex > Count Then
        Exit Sub
    End If
    If DestRowIndex < 1 Or DestRowIndex > Count Then
        Exit Sub
    End If
    If DestRowIndex = OriginRowIndex Then
        Exit Sub
    End If
    mp_oCollection.m_lCopyAndMoveItems OriginRowIndex, DestRowIndex
End Sub

Public Sub SortRows(ByVal PropertyName As String, ByVal Descending As Boolean, ByVal SortType As E_SORTTYPE, Optional ByVal StartIndex As Long = -1, Optional ByVal EndIndex As Long = -1)
    If StartIndex = -1 Then
        StartIndex = 1
    End If
    If EndIndex = -1 Then
        EndIndex = Count
    End If
    If Count = 0 Then Exit Sub
    If StartIndex < 1 Or StartIndex > Count Then
        Exit Sub
    End If
    If EndIndex < 1 Or EndIndex > Count Then
        Exit Sub
    End If
    If EndIndex = StartIndex Then
        Exit Sub
    End If
    mp_oCollection.m_Sort PropertyName, Descending, SortType, StartIndex, EndIndex
End Sub

Public Sub SortCells(ByVal CellIndex As Long, ByVal PropertyName As String, ByVal Descending As Boolean, ByVal SortType As E_SORTTYPE, Optional ByVal StartIndex As Long = -1, Optional ByVal EndIndex As Long = -1)
    If StartIndex = -1 Then
        StartIndex = 1
    End If
    If EndIndex = -1 Then
        EndIndex = Count
    End If
    If Count = 0 Then Exit Sub
    If StartIndex < 1 Or StartIndex > Count Then
        Exit Sub
    End If
    If EndIndex < 1 Or EndIndex > Count Then
        Exit Sub
    End If
    If EndIndex = StartIndex Then
        Exit Sub
    End If
    If CellIndex < 1 Or CellIndex > mp_oControl.Columns.Count Then
        Exit Sub
    End If
    mp_oCollection.m_SortCells CellIndex, PropertyName, Descending, SortType, StartIndex, EndIndex
End Sub

Friend Sub ClearCells()
    Dim lIndex As Long
    Dim oRow As clsRow
    For lIndex = 1 To mp_oCollection.m_lCount
        Set oRow = mp_oCollection.m_oReturnArrayElement(lIndex)
        oRow.Cells.Clear
    Next lIndex
End Sub

Friend Function Height() As Long
    Dim lBuffer As Long
    Dim lIndex As Long
    Dim oRow As clsRow
    If Count = 0 Then
        Height = 0
        Exit Function
    End If
    For lIndex = 1 To Count
        Set oRow = mp_oCollection.m_oReturnArrayElement(lIndex)
        lBuffer = lBuffer + oRow.Height
    Next lIndex
    Height = lBuffer
End Function

Friend Function CalculateHeight(ByVal StartIndex As Long, ByVal EndIndex As Long) As Long
    Dim lBuffer As Long
    Dim lIndex As Long
    Dim oRow As clsRow
    If StartIndex = 0 Then
        CalculateHeight = 0
        Exit Function
    End If
    For lIndex = StartIndex To EndIndex
        Set oRow = mp_oCollection.m_oReturnArrayElement(lIndex)
        lBuffer = lBuffer + oRow.Height
    Next lIndex
    CalculateHeight = lBuffer
End Function

Friend Function CalculateRows(ByVal StartIndex As Long, ByVal Height As Long) As Long
    Dim lBuffer As Long
    Dim lIndex As Long
    Dim oRow As clsRow
    Dim lRows As Long
    lRows = 1
    If StartIndex = 0 Then
        CalculateRows = lRows
        Exit Function
    End If
    For lIndex = StartIndex To Count
        Set oRow = mp_oCollection.m_oReturnArrayElement(lIndex)
        lBuffer = lBuffer + oRow.Height
        If lBuffer > Height Then
            Exit For
        End If
        lRows = lRows + 1
    Next lIndex
    CalculateRows = lRows
End Function

'// ---------------------------------------------------------------------------------------------------------------------
'// Position
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Position()
    Dim oRow As clsRow
    Dim lRowIndex As Long
    Dim lBottomBuff As Long
    Dim oClientArea As clsClientArea
    Set oClientArea = mp_oControl.CurrentViewObject.ClientArea
    If Count = 0 Then
        oClientArea.f_LastVisibleRow = 0
        mp_lTopOffset = mp_oControl.CurrentViewObject.ClientArea.Top
        Exit Sub
    Else
        mp_lTopOffset = 0
    End If
    For lRowIndex = 1 To (oClientArea.FirstVisibleRow - 1)
        Set oRow = mp_oCollection.m_oReturnArrayElement(lRowIndex)
        oRow.f_lTop = -200
        oRow.f_lBottom = -200 + oRow.Height
        oRow.f_bVisible = False
    Next lRowIndex
    For lRowIndex = oClientArea.FirstVisibleRow To Count
        Set oRow = mp_oCollection.m_oReturnArrayElement(lRowIndex)
        oRow.f_bVisible = True
        If lRowIndex = oClientArea.FirstVisibleRow Then
            oRow.f_lTop = mp_oControl.CurrentViewObject.ClientArea.Top
        Else
            oRow.f_lTop = lBottomBuff + 1
        End If
        oRow.f_lBottom = oRow.Top + oRow.Height
        lBottomBuff = oRow.Bottom
        mp_lTopOffset = oRow.Bottom
        oClientArea.f_LastVisibleRow = lRowIndex
        If mp_lTopOffset >= mp_oControl.CurrentViewObject.ClientArea.Bottom Then
            Exit For
        End If
    Next lRowIndex
    For lRowIndex = (oClientArea.LastVisibleRow + 1) To Count
        Set oRow = mp_oCollection.m_oReturnArrayElement(lRowIndex)
        oRow.f_lTop = mp_oControl.CurrentViewObject.ClientArea.Bottom + 200
        oRow.f_lBottom = mp_oControl.CurrentViewObject.ClientArea.Bottom + 200 + oRow.Height
        oRow.f_bVisible = False
    Next lRowIndex
End Sub

Friend Property Get TopOffset() As Long
    TopOffset = mp_lTopOffset
End Property

Friend Property Let TopOffset(ByVal Value As Long)
    mp_lTopOffset = Value
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Draw
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Draw()
    Dim lCellIndex As Long
    Dim lRowIndex As Long
    Dim oRow As clsRow
    Dim oColumn As clsColumn
    Dim oCell As clsCell
    Dim bDraw As Boolean
    mp_oControl.GrphLib.ClipRegion mp_oControl.mt_LeftMargin, mp_oControl.CurrentViewObject.ClientArea.Top, mp_oControl.Splitter.Left, mp_oControl.mt_TableBottom, False
    bDraw = False
    mp_oControl.FireDraw E_EVENTTARGET.EVT_TABLE, bDraw, 0, 0, mp_oControl.GrphLib.mp_lHdc
    If bDraw = True Then
        Exit Sub
    End If
    If Count = 0 Then
        Exit Sub
    End If
    For lRowIndex = mp_oControl.CurrentViewObject.ClientArea.FirstVisibleRow To mp_oControl.CurrentViewObject.ClientArea.LastVisibleRow
        Set oRow = mp_oCollection.m_oReturnArrayElement(lRowIndex)
        If oRow.Visible = True Then
            If oRow.MergeCells = True Then
                If mp_oControl.GrphLib.RectangleIntersection(oRow.Left, oRow.Top, oRow.Right, oRow.Bottom) Then
                    mp_oControl.GrphLib.ClipRegion oRow.Left, oRow.Top, oRow.Right, oRow.Bottom, True
                    bDraw = False
                    mp_oControl.FireDraw E_EVENTTARGET.EVT_ROW, bDraw, lRowIndex, 0, mp_oControl.GrphLib.mp_lHdc
                    If bDraw = False Then
                        mp_oControl.mp_DrawItem oRow.Left, oRow.Right, oRow.Top, oRow.Bottom, oRow.StyleIndex, oRow.Caption, (lRowIndex = mp_oControl.SelectedRowIndex), oRow.Picture, 0, 0, Nothing
                    End If
                End If
            Else
                For lCellIndex = 1 To mp_oControl.Columns.Count
                    Set oCell = oRow.Cells.oCollection.m_oReturnArrayElement(lCellIndex)
                    Set oColumn = mp_oControl.Columns.oCollection.m_oReturnArrayElement(lCellIndex)
                    If oColumn.Visible = True Then
                        If mp_oControl.GrphLib.RectangleIntersection(oCell.LeftTrim, oCell.Top, oCell.RightTrim, oCell.Bottom) Then
                            mp_oControl.GrphLib.ClipRegion oCell.LeftTrim, oCell.Top, oCell.RightTrim, oCell.Bottom, True
                            bDraw = False
                            mp_oControl.FireDraw E_EVENTTARGET.EVT_CELL, bDraw, lRowIndex, lCellIndex, mp_oControl.GrphLib.mp_lHdc
                            If bDraw = False Then
                                mp_oControl.mp_DrawItem oCell.Left, oCell.Right - 1, oCell.Top, oCell.Bottom, oCell.StyleIndex, oCell.Caption, (lRowIndex = mp_oControl.SelectedRowIndex And lCellIndex = mp_oControl.SelectedCellIndex), oCell.Picture, oCell.LeftTrim, oCell.RightTrim, Nothing
                            End If
                        End If
                    End If
                Next lCellIndex
            End If
        End If
    Next lRowIndex
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim lIndex As Long
    Dim oRow As clsRow
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Rows"
    oXML.InitializeWriter
    For lIndex = 1 To Count
        Set oRow = mp_oCollection.m_oReturnArrayElement(lIndex)
        oXML.WriteObject oRow.GetXML
    Next lIndex
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim lIndex As Long
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Rows"
    oXML.SetXML sXML
    oXML.InitializeReader
    mp_oCollection.m_Clear
    For lIndex = 1 To oXML.ReadCollectionCount
        Dim oRow As New clsRow
        oRow.Initialize mp_oControl
        oRow.SetXML oXML.ReadCollectionObject(lIndex)
        mp_oCollection.AddMode = True
        mp_oCollection.m_Add oRow, oRow.Key, 50000, 50000
        Set oRow = Nothing
    Next lIndex
    mp_oControl.f_oVScrollBar.Max = Count
End Sub






