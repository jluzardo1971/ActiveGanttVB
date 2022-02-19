VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouseEvents"
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
    Private mp_lX As Long
    Private mp_lY As Long
    Private mp_oButton As Long
    Private mp_yMoveType As E_OPERATION
    Private mp_yClickType As E_OPERATION
    Private mp_yDragType As E_OPERATION
    Private mp_bCompleteDragOperation As Boolean
    Private mp_bMilestoneMove As Boolean
    Private mp_bTaskMove As Boolean
    Private mp_dtAddEnd As Date
    Private mp_dtAddStart As Date
    Private mp_lDeltaToLeft As Long
    Private mp_lDeltaToRight As Long
    Private mp_lDragOptions As Long
    Private mp_lDragColumnIndex As Long
    Private mp_lDragRowIndex As Long
    Private mp_lOrgDragRowIndex As Long
    Private mp_lFixedDeltaDates As Long
    Private mp_lFixedX As Long
    Private mp_lRowDraggingDestinationIndex As Long
    Private mp_lColumnDraggingDestinationIndex As Long
    Private mp_lColumnToDragIndex As Long
    Private mp_lRowToDragIndex As Long
    Private mp_lVisGIIndex As Long
    Private mp_lVisMilestoneIndex As Long
    Private mp_lVisColumnIndex As Long
    Private mp_lVisRowIndex As Long
    Private mp_lVisCellIndex As Long
    Private mp_sAddRowKey As String
    Private mp_oToolTip As Label
    Private dtNull As Date

    Private Enum E_MOUSEEVENTS
        MouseMoveNoDrag = 0
        MouseDown = 1
        MouseMove = 2
        MouseUp = 3
        MouseClick = 4
        MouseDblClick = 5
    End Enum

    Private Enum E_OPERATION
        EO_NONE = 0

        EO_TASKADDITION = 1
        EO_TASKMOVEMENT = 2
        EO_TASKSTRETCHLEFT = 3
        EO_TASKSTRETCHRIGHT = 4
        EO_TASKSELECTION = 5

        EO_MILESTONEADDITION = 6
        EO_MILESTONEMOVEMENT = 7
        EO_MILESTONESELECTION = 8

        EO_ROWSIZING = 9
        EO_ROWMOVEMENT = 10
        EO_ROWSELECTION = 11

        EO_COLUMNSIZING = 12
        EO_COLUMNMOVEMENT = 13
        EO_COLUMNSELECTION = 14

        EO_CLIENTAREA = 15

        EO_TIMELINEMOVEMENT = 16
        EO_TIMELINESELECTION = 17

        EO_SPLITTERMOVEMENT = 18

    End Enum

    Private Enum E_CURSORTYPE
        CT_NORMAL = 0
        CT_SIZETASK = 1
        CT_MOVETASK = 2
        CT_MOVEMILESTONE = 3
        CT_CLIENTAREA = 4
        CT_MOVESPLITTER = 5
        CT_ROWHEIGHT = 6
        CT_COLUMNWIDTH = 7
        CT_MOVEROW = 8
        CT_MOVECOLUMN = 11
        CT_SCROLLTIMELINE = 9
        CT_NODROP = 10
    End Enum
    
    Private Enum E_FOCUSTYPE
        FCT_NORMAL = 0
        FCT_KEEPLEFTRIGHTBOUNDS = 1
        FCT_ADD = 2
        FCT_VERTICALSPLITTER = 3
    End Enum

    '// ---------------------------------------------------------------------------------------------------------------------
    '// Construction/Destruction & Initialization
    '// ---------------------------------------------------------------------------------------------------------------------

    Public Sub Initialize(ByRef Value As ActiveGanttVBCtl)
        '// Parent Control Pointer
        Set mp_oControl = Value
        '// Object Member Variables
        Set mp_oToolTip = mp_oControl.f_ToolTip
        dtNull = CDate(0)
        mp_yMoveType = E_OPERATION.EO_NONE
        mp_yClickType = E_OPERATION.EO_NONE
        mp_yDragType = E_OPERATION.EO_NONE
        mp_bCompleteDragOperation = False
        mp_bMilestoneMove = False
        mp_bTaskMove = False
    End Sub

    Private Sub Class_Terminate()
        Set mp_oControl = Nothing
    End Sub

    '// ---------------------------------------------------------------------------------------------------------------------
    '// Methods
    '// ---------------------------------------------------------------------------------------------------------------------

    Public Sub HideToolTip()
        mp_oToolTip.Visible = False
    End Sub

    Private Sub mp_MouseMovements(ByVal v_yOpType As E_OPERATION, ByVal v_yMouseEvent As E_MOUSEEVENTS, ByVal X As Long, ByVal Y As Long)
        Dim yObjectType As E_EVENTTARGET
        Dim lObjectIndex As Long
        Dim lParentObjectIndex As Long
        Dim oRow As clsRow
        yObjectType = E_EVENTTARGET.EVT_NONE
        lObjectIndex = 0
        lParentObjectIndex = 0
        Select Case v_yOpType
            Case E_OPERATION.EO_TASKADDITION
                yObjectType = E_EVENTTARGET.EVT_CLIENTAREA
                mp_DragAdd v_yMouseEvent, X, Y
            Case E_OPERATION.EO_TASKMOVEMENT
                yObjectType = E_EVENTTARGET.EVT_TASK
                lObjectIndex = mp_oControl.SelectedTaskIndex
                If v_yMouseEvent = E_MOUSEEVENTS.MouseMoveNoDrag Then
                    mp_DragMoveTask E_MOUSEEVENTS.MouseMoveNoDrag, X, Y, mp_iOverSelectedTask(X, Y)
                Else
                    mp_DragMoveTask v_yMouseEvent, X, Y, E_TASKMOVETYPE.GMT_MOVE
                End If
            Case E_OPERATION.EO_TASKSTRETCHLEFT
                mp_DragMoveTask v_yMouseEvent, X, Y, E_TASKMOVETYPE.GMT_STRETCHLEFT
            Case E_OPERATION.EO_TASKSTRETCHRIGHT
                mp_DragMoveTask v_yMouseEvent, X, Y, E_TASKMOVETYPE.GMT_STRETCHRIGHT
            Case E_OPERATION.EO_TASKSELECTION
                yObjectType = E_EVENTTARGET.EVT_TASK
                lObjectIndex = mp_lVisGIIndex
                mp_DragSelectingTask v_yMouseEvent, X, Y
            Case E_OPERATION.EO_MILESTONEADDITION
                yObjectType = E_EVENTTARGET.EVT_CLIENTAREA
                mp_DragAdd v_yMouseEvent, X, Y
            Case E_OPERATION.EO_MILESTONEMOVEMENT
                yObjectType = E_EVENTTARGET.EVT_MILESTONE
                lObjectIndex = mp_oControl.SelectedMilestoneIndex
                mp_DragMoveMilestone v_yMouseEvent, X, Y
            Case E_OPERATION.EO_MILESTONESELECTION
                yObjectType = E_EVENTTARGET.EVT_MILESTONE
                lObjectIndex = mp_lVisMilestoneIndex
                mp_DragSelectingMilestone v_yMouseEvent, X, Y
            Case E_OPERATION.EO_ROWSIZING
                mp_DragRowHeight v_yMouseEvent, X, Y
            Case E_OPERATION.EO_ROWMOVEMENT
                Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_oControl.SelectedRowIndex)
                If oRow.MergeCells = True Then
                    yObjectType = E_EVENTTARGET.EVT_ROW
                    lObjectIndex = mp_oControl.SelectedRowIndex
                Else
                    yObjectType = E_EVENTTARGET.EVT_CELL
                    lObjectIndex = mp_oControl.SelectedCellIndex
                    lParentObjectIndex = mp_oControl.SelectedRowIndex
                End If
                mp_DragMoveRow v_yMouseEvent, X, Y
            Case E_OPERATION.EO_ROWSELECTION
                Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_lVisRowIndex)
                If oRow.MergeCells = True Then
                    yObjectType = E_EVENTTARGET.EVT_ROW
                    lObjectIndex = mp_lVisRowIndex
                Else
                    yObjectType = E_EVENTTARGET.EVT_CELL
                    lObjectIndex = mp_lVisCellIndex
                    lParentObjectIndex = mp_lVisRowIndex
                End If
                mp_DragSelectingRow v_yMouseEvent, X, Y
            Case E_OPERATION.EO_COLUMNSIZING
                mp_DragColumnWidth v_yMouseEvent, X, Y
            Case E_OPERATION.EO_COLUMNMOVEMENT
                yObjectType = E_EVENTTARGET.EVT_COLUMN
                lObjectIndex = mp_oControl.SelectedColumnIndex
                mp_DragMoveColumn v_yMouseEvent, X, Y
            Case E_OPERATION.EO_COLUMNSELECTION
                yObjectType = E_EVENTTARGET.EVT_COLUMN
                lObjectIndex = mp_lVisColumnIndex
                mp_DragSelectingColumn v_yMouseEvent, X, Y
            Case E_OPERATION.EO_CLIENTAREA
                yObjectType = E_EVENTTARGET.EVT_CLIENTAREA
                Select Case v_yMouseEvent
                    Case E_MOUSEEVENTS.MouseMoveNoDrag
                    Case E_MOUSEEVENTS.MouseDown
                        mp_yDragType = E_OPERATION.EO_CLIENTAREA
                    Case E_MOUSEEVENTS.MouseMove
                    Case E_MOUSEEVENTS.MouseUp
                End Select
            Case E_OPERATION.EO_TIMELINEMOVEMENT
                yObjectType = E_EVENTTARGET.EVT_TIMELINE
                mp_DragTimeLine v_yMouseEvent, X, Y
            Case E_OPERATION.EO_TIMELINESELECTION
                yObjectType = E_EVENTTARGET.EVT_TIMELINE
                Select Case v_yMouseEvent
                    Case E_MOUSEEVENTS.MouseMoveNoDrag
                    Case E_MOUSEEVENTS.MouseDown
                        mp_yDragType = E_OPERATION.EO_TIMELINESELECTION
                    Case E_MOUSEEVENTS.MouseMove
                    Case E_MOUSEEVENTS.MouseUp
                End Select
            Case E_OPERATION.EO_SPLITTERMOVEMENT
                mp_DragMoveSplitter v_yMouseEvent, X, Y
            Case E_OPERATION.EO_NONE
                Select Case v_yMouseEvent
                    Case E_MOUSEEVENTS.MouseMoveNoDrag
                    Case E_MOUSEEVENTS.MouseDown
                        mp_yDragType = E_OPERATION.EO_NONE
                    Case E_MOUSEEVENTS.MouseMove
                    Case E_MOUSEEVENTS.MouseUp
                End Select
        End Select

        Select Case v_yMouseEvent
            Case E_MOUSEEVENTS.MouseClick
                mp_oControl.FireControlClick yObjectType, lObjectIndex, lParentObjectIndex, X, Y, mp_oButton
            Case E_MOUSEEVENTS.MouseDblClick
                mp_oControl.FireControlDblClick yObjectType, lObjectIndex, lParentObjectIndex, X, Y, mp_oButton
            Case E_MOUSEEVENTS.MouseDown
                mp_oControl.FireControlMouseDown yObjectType, lObjectIndex, lParentObjectIndex, X, Y, mp_oButton
            Case E_MOUSEEVENTS.MouseMove
                mp_oControl.FireControlMouseMove yObjectType, lObjectIndex, lParentObjectIndex, X, Y, mp_oButton
            Case E_MOUSEEVENTS.MouseUp
                mp_oControl.FireControlMouseUp yObjectType, lObjectIndex, lParentObjectIndex, X, Y, mp_oButton
        End Select

    End Sub

    Public Sub Click()
        mp_MouseMovements mp_yClickType, E_MOUSEEVENTS.MouseClick, mp_lX, mp_lY
    End Sub

    Public Sub DblClick()
        mp_MouseMovements mp_yClickType, E_MOUSEEVENTS.MouseDblClick, mp_lX, mp_lY
    End Sub

    Public Sub m_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        mp_oButton = Button
        mp_MouseMovements mp_yMoveType, E_MOUSEEVENTS.MouseDown, X, Y
        mp_yClickType = mp_yMoveType
        mp_yMoveType = E_OPERATION.EO_NONE
    End Sub

    Public Sub m_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Dim oRow As clsRow
        If mp_yDragType <> E_OPERATION.EO_NONE Then
            mp_MouseMovements mp_yDragType, E_MOUSEEVENTS.MouseMove, X, Y
        Else
            Dim yOpType As E_OPERATION
            yOpType = E_OPERATION.EO_NONE
            If Not (X > mp_oControl.Splitter.Right And X < mp_oControl.mt_RightMargin And Y > mp_oControl.CurrentViewObject.ClientArea.Top And Y < mp_oControl.Rows.TopOffset) Then
                'RaiseEvent CursorPositionChanged("", E_TYPE.TP_NONE, 0, False)
            End If
            If mp_bOverSplitter(X, Y) = True Then
                yOpType = E_OPERATION.EO_SPLITTERMOVEMENT
            ElseIf mp_bOverEmptySpace(Y) = True Then
                mp_MovingOverNothing
            ElseIf mp_bOverTimeLine(X, Y) = True Then
                yOpType = E_OPERATION.EO_TIMELINEMOVEMENT
            ElseIf mp_bOverColumnDivision(X, Y) = True Then
                yOpType = E_OPERATION.EO_COLUMNSIZING
            ElseIf mp_bOverSelectedColumn(X, Y) = True Then
                yOpType = E_OPERATION.EO_COLUMNMOVEMENT
            ElseIf mp_bOverColumn(X, Y) = True Then
                yOpType = E_OPERATION.EO_COLUMNSELECTION
            ElseIf mp_bOverRowDivision(X, Y) = True Then
                yOpType = E_OPERATION.EO_ROWSIZING
            ElseIf mp_bOverSelectedRow(X, Y) = True Then
                yOpType = E_OPERATION.EO_ROWMOVEMENT
            ElseIf mp_bOverRow(X, Y) = True Then
                yOpType = E_OPERATION.EO_ROWSELECTION
            ElseIf mp_iOverSelectedTask(X, Y) <> E_TASKMOVETYPE.GMT_NONE Then
                yOpType = E_OPERATION.EO_TASKMOVEMENT
            ElseIf mp_bOverTask(X, Y) = True Then
                yOpType = E_OPERATION.EO_TASKSELECTION
            ElseIf mp_bOverSelectedMilestone(X, Y) = True Then
                yOpType = E_OPERATION.EO_MILESTONEMOVEMENT
            ElseIf mp_bOverMilestone(X, Y) = True Then
                yOpType = E_OPERATION.EO_MILESTONESELECTION
            ElseIf mp_bOverClientArea(X, Y) = True Then
                yOpType = E_OPERATION.EO_TASKADDITION
            End If
            mp_MouseMovements yOpType, E_MOUSEEVENTS.MouseMoveNoDrag, X, Y
        End If
    End Sub

    Public Sub m_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        mp_oButton = Button
        mp_lX = X
        mp_lY = Y
        mp_MouseMovements mp_yDragType, E_MOUSEEVENTS.MouseUp, X, Y
        'RaiseEvent CursorPositionChanged("", E_TYPE.TP_NONE, 0, False)
        mp_oToolTip.Visible = False
        Dim bDragRowHeight As Boolean
        If mp_yDragType = E_OPERATION.EO_ROWSIZING Then
            bDragRowHeight = True
        End If
        mp_yDragType = E_OPERATION.EO_NONE
        mp_SetCursor E_CURSORTYPE.CT_NORMAL
        mp_oControl.GrphLib.f_FocusLeft = 0
        mp_oControl.GrphLib.f_FocusTop = 0
        mp_oControl.GrphLib.f_FocusRight = 0
        mp_oControl.GrphLib.f_FocusBottom = 0
        mp_oControl.Redraw
        If bDragRowHeight = True Then
            'RaiseEvent Scroll(1)
        End If
    End Sub

    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '//Drag Operations
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub mp_DragSelectingColumn(ByVal v_yMouseOperation As E_MOUSEEVENTS, ByRef X As Long, ByRef Y As Long)
        Select Case v_yMouseOperation
            Case E_MOUSEEVENTS.MouseMoveNoDrag
                mp_oToolTip.Visible = False
                mp_yMoveType = E_OPERATION.EO_COLUMNSELECTION
                mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEEVENTS.MouseDown
                If mp_bOverColumn(X, Y) = False Then
                    mp_oControl.mp_ErrorReport 50000, "Unexpected Error", "mp_DragSelectingHeading"
                End If
                mp_yDragType = E_OPERATION.EO_COLUMNSELECTION
            Case E_MOUSEEVENTS.MouseMove
            Case E_MOUSEEVENTS.MouseUp
                mp_oControl.SelectedColumnIndex = mp_lVisColumnIndex
                mp_oControl.FireObjectSelected E_EVENTTARGET.EVT_COLUMN, mp_oControl.SelectedColumnIndex, 0
        End Select
    End Sub

    Private Sub mp_DragSelectingRow(ByVal v_yMouseOperation As E_MOUSEEVENTS, ByRef X As Long, ByRef Y As Long)
        Dim oRow As clsRow
        Select Case v_yMouseOperation
            Case E_MOUSEEVENTS.MouseMoveNoDrag
                mp_oToolTip.Visible = False
                mp_yMoveType = E_OPERATION.EO_ROWSELECTION
                mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEEVENTS.MouseDown
                If mp_bOverRow(X, Y) = False Then
                    mp_oControl.mp_ErrorReport 50000, "Unexpected Error", "mp_DragSelectingRow"
                End If
                mp_yDragType = E_OPERATION.EO_ROWSELECTION
            Case E_MOUSEEVENTS.MouseMove
            Case E_MOUSEEVENTS.MouseUp
                mp_oControl.SelectedRowIndex = mp_lVisRowIndex
                mp_oControl.SelectedCellIndex = mp_lVisCellIndex
                Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_oControl.SelectedRowIndex)
                If oRow.MergeCells = True Then
                    mp_oControl.FireObjectSelected E_EVENTTARGET.EVT_ROW, mp_oControl.SelectedRowIndex, 0
                Else
                    mp_oControl.FireObjectSelected E_EVENTTARGET.EVT_CELL, mp_oControl.SelectedCellIndex, mp_oControl.SelectedRowIndex
                End If
        End Select
    End Sub

    Private Sub mp_DragSelectingMilestone(ByVal v_yMouseOperation As E_MOUSEEVENTS, ByRef X As Long, ByRef Y As Long)
        Dim oMilestone As clsMilestone
        Select Case v_yMouseOperation
            Case E_MOUSEEVENTS.MouseMoveNoDrag
                ToolTipCaption E_TOOLTIPACTION.TA_OVERMILESTONE, mp_lVisMilestoneIndex, X, Y, dtNull, dtNull
                mp_yMoveType = E_OPERATION.EO_MILESTONESELECTION
                mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEEVENTS.MouseDown
                mp_yDragType = E_OPERATION.EO_MILESTONESELECTION
            Case E_MOUSEEVENTS.MouseMove
            Case E_MOUSEEVENTS.MouseUp
                mp_oControl.SelectedMilestoneIndex = mp_lVisMilestoneIndex
                If mp_oControl.CurrentViewObject.ClientArea.Grid.SnapToGrid = True And mp_oControl.CurrentViewObject.ClientArea.Grid.SnapToGridOnSelection = True Then
                    Set oMilestone = mp_oControl.Milestones.oCollection.m_oReturnArrayElement(mp_oControl.SelectedMilestoneIndex)
                    oMilestone.MilestoneDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.f_Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.f_Factor, oMilestone.MilestoneDate)
                End If
                mp_oControl.FireObjectSelected E_EVENTTARGET.EVT_MILESTONE, mp_oControl.SelectedMilestoneIndex, 0
        End Select
    End Sub

    Private Sub mp_DragSelectingTask(ByVal v_yMouseOperation As E_MOUSEEVENTS, ByRef X As Long, ByRef Y As Long)
        Dim oTask As clsTask
        Select Case v_yMouseOperation
            Case E_MOUSEEVENTS.MouseMoveNoDrag
                ToolTipCaption E_TOOLTIPACTION.TA_OVERTASK, mp_lVisGIIndex, X, Y, dtNull, dtNull
                mp_yMoveType = E_OPERATION.EO_TASKSELECTION
                mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEEVENTS.MouseDown
                mp_yDragType = E_OPERATION.EO_TASKSELECTION
            Case E_MOUSEEVENTS.MouseMove
            Case E_MOUSEEVENTS.MouseUp
                mp_oControl.SelectedTaskIndex = mp_lVisGIIndex
                If mp_oControl.CurrentViewObject.ClientArea.Grid.SnapToGrid = True And mp_oControl.CurrentViewObject.ClientArea.Grid.SnapToGridOnSelection = True Then
                    Set oTask = mp_oControl.Tasks.oCollection.m_oReturnArrayElement(mp_oControl.SelectedTaskIndex)
                    oTask.StartDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.f_Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.f_Factor, oTask.StartDate)
                    oTask.EndDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.f_Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.f_Factor, oTask.EndDate)
                End If
                mp_oControl.FireObjectSelected E_EVENTTARGET.EVT_TASK, mp_oControl.SelectedTaskIndex, 0
        End Select
    End Sub

    Private Sub mp_DragMoveColumn(ByVal v_yMouseOperation As E_MOUSEEVENTS, ByRef X As Long, ByRef Y As Long)
        Dim oColumn As clsColumn
        Dim oRow As clsRow
        Dim lIndex As Long
        Dim bCancel As Boolean
        Select Case v_yMouseOperation
            Case E_MOUSEEVENTS.MouseMoveNoDrag
                If mp_oControl.AllowColumnSwap = True Then
                    mp_yMoveType = E_OPERATION.EO_COLUMNMOVEMENT
                    mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
                Else
                    mp_yMoveType = E_OPERATION.EO_COLUMNSELECTION
                    mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
                End If
            Case E_MOUSEEVENTS.MouseDown
                If mp_oControl.AllowColumnSwap = True Then
                    bCancel = False
                    mp_oControl.FireBeginObjectMove E_EVENTTARGET.EVT_COLUMN, mp_oControl.SelectedColumnIndex, bCancel
                    If mp_oButton = vbLeftButton And bCancel = False Then
                        mp_SetCursor (E_CURSORTYPE.CT_MOVECOLUMN)
                        mp_yDragType = E_OPERATION.EO_COLUMNMOVEMENT
                        mp_lColumnToDragIndex = mp_lReturnCellIndexByPosition(X)
                    End If
                    If bCancel = True Then
                        CancelUIOperations
                    End If
                Else
                    mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
                    mp_yDragType = E_OPERATION.EO_COLUMNSELECTION
                End If
            Case E_MOUSEEVENTS.MouseMove
                mp_oControl.GrphLib.EraseReversibleFrames
                mp_DynamicColumnMove (X)
                bCancel = False
                mp_lDragColumnIndex = mp_lReturnCellIndexByPosition(X)
                mp_oControl.FireObjectMove E_EVENTTARGET.EVT_COLUMN, mp_lDragColumnIndex, bCancel
                If bCancel = False Then
                    If mp_lDragColumnIndex = -1 Then
                        mp_SetCursor (E_CURSORTYPE.CT_NODROP)
                        Exit Sub
                    Else
                        mp_SetCursor (E_CURSORTYPE.CT_MOVECOLUMN)
                    End If
                    Set oColumn = mp_oControl.Columns.oCollection.m_oReturnArrayElement(mp_lDragColumnIndex)
                    mp_DrawMovingReversibleFrame oColumn.LeftTrim, oColumn.Top, oColumn.RightTrim, oColumn.Bottom, E_FOCUSTYPE.FCT_NORMAL
                End If
                If bCancel = True Then
                    CancelUIOperations
                End If
            Case E_MOUSEEVENTS.MouseUp
                mp_oControl.GrphLib.EraseReversibleFrames
                bCancel = False
                mp_lColumnDraggingDestinationIndex = mp_lReturnCellIndexByPosition(X)
                mp_oControl.FireEndObjectMove E_EVENTTARGET.EVT_COLUMN, mp_lColumnDraggingDestinationIndex, bCancel
                If mp_bCompleteDragOperation = True And bCancel = False Then
                    If mp_lColumnToDragIndex = mp_lColumnDraggingDestinationIndex Then
                        Exit Sub
                    End If
                    mp_oControl.SelectedColumnIndex = mp_oControl.Columns.oCollection.m_lCopyAndMoveItems(mp_lColumnToDragIndex, mp_lColumnDraggingDestinationIndex)
                    For lIndex = 1 To mp_oControl.Rows.Count
                        Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex)
                        oRow.Cells.oCollection.m_lCopyAndMoveItems mp_lColumnToDragIndex, mp_lColumnDraggingDestinationIndex
                    Next lIndex
                    mp_oControl.FireCompleteObjectMove E_EVENTTARGET.EVT_COLUMN, mp_oControl.SelectedColumnIndex
                End If
                mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
                If bCancel = True Then
                    CancelUIOperations
                End If
        End Select
    End Sub

    Private Sub mp_DragAdd(ByVal v_yMouseOperation As E_MOUSEEVENTS, ByRef X As Long, ByRef Y As Long)
        Dim oRow As clsRow
        Select Case v_yMouseOperation
            Case E_MOUSEEVENTS.MouseMoveNoDrag
                ToolTipCaption E_TOOLTIPACTION.TA_CLIENTAREA, 0, X, Y, dtNull, dtNull
                If mp_oControl.AllowAdd = True Then
                    Select Case mp_oControl.AddMode
                        Case E_ADDMODE.AT_TASKADD
                            mp_yMoveType = E_OPERATION.EO_TASKADDITION
                            mp_SetCursor (E_CURSORTYPE.CT_CLIENTAREA)
                        Case E_ADDMODE.AT_MILESTONEADD
                            mp_yMoveType = E_OPERATION.EO_MILESTONEADDITION
                            mp_SetCursor (E_CURSORTYPE.CT_CLIENTAREA)
                            mp_lDragRowIndex = mp_lReturnRowIndexByPosition(Y)
                            If mp_lDragRowIndex = 0 Then
                                Exit Sub
                            End If
                            Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_lDragRowIndex)
                            If (mp_oControl.mp_bDetectConflict(mp_oControl.MathLib.GetDateFromXCoordinate(X), mp_oControl.MathLib.CurrentDateTime, oRow.Key, 0, mp_oControl.CurrentLayer, E_TYPE.TP_MILESTONE) = True Or oRow.Container = False) And mp_oControl.CurrentViewObject.ClientArea.DetectConflicts = True Then
                                Exit Sub
                            End If
                            ToolTipCaption E_TOOLTIPACTION.TA_CLIENTAREA, 0, X, Y, dtNull, dtNull
                    End Select
                Else
                    mp_yMoveType = E_OPERATION.EO_CLIENTAREA
                    mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
                End If
            Case E_MOUSEEVENTS.MouseDown
                If (mp_oButton = vbLeftButton) Then
                    X = mp_fSnapX(X)
                    mp_dtAddStart = mp_oControl.MathLib.GetDateFromXCoordinate(X)
                    mp_lDragRowIndex = mp_lReturnRowIndexByPosition(Y)
                    If (mp_lDragRowIndex = 0) Then
                        mp_bCompleteDragOperation = False
                        Exit Sub
                    End If
                    Select Case mp_oControl.AddMode
                        Case E_ADDMODE.AT_TASKADD
                            mp_yDragType = E_OPERATION.EO_TASKADDITION
                        Case E_ADDMODE.AT_MILESTONEADD
                            mp_yDragType = E_OPERATION.EO_MILESTONEADDITION
                    End Select
                Else
                    mp_yDragType = E_OPERATION.EO_CLIENTAREA
                End If
            Case E_MOUSEEVENTS.MouseMove
                Select Case mp_oControl.AddMode
                    Case E_ADDMODE.AT_TASKADD
                        Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_lDragRowIndex)
                        X = mp_fSnapX(X)
                        mp_oControl.GrphLib.EraseReversibleFrames
                        If (mp_oControl.mp_bDetectConflict(mp_dtAddStart, mp_oControl.MathLib.GetDateFromXCoordinate(X), oRow.Key, 0, mp_oControl.CurrentLayer) = True Or oRow.Container = False) And mp_oControl.CurrentViewObject.ClientArea.DetectConflicts = True Then
                            mp_SetCursor (E_CURSORTYPE.CT_NODROP)
                        Else
                            mp_SetCursor (E_CURSORTYPE.CT_CLIENTAREA)
                        End If
                        mp_DynamicTimeLineMove (X)
                        ToolTipCaption E_TOOLTIPACTION.TA_TASKADD, mp_oControl.SelectedTaskIndex, X, Y, mp_dtAddStart, mp_oControl.MathLib.GetDateFromXCoordinate(X)
                        mp_DrawMovingReversibleFrame mp_oControl.MathLib.GetXCoordinateFromDate(mp_dtAddStart), oRow.Top, X, oRow.Bottom, E_FOCUSTYPE.FCT_ADD
                    Case E_ADDMODE.AT_MILESTONEADD
                End Select
            Case E_MOUSEEVENTS.MouseUp
                mp_oControl.GrphLib.EraseReversibleFrames
                If mp_bCompleteDragOperation = True Then
                    Select Case mp_oControl.AddMode
                        Case E_ADDMODE.AT_TASKADD
                            Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_lDragRowIndex)
                            X = mp_fSnapX(X)
                            mp_dtAddEnd = mp_oControl.MathLib.GetDateFromXCoordinate(X)
                            If mp_dtAddEnd = mp_dtAddStart Then
                                mp_oControl.SelectedTaskIndex = 0
                                Exit Sub
                            End If
                            If mp_dtAddEnd < mp_dtAddStart Then
                                mp_oControl.Tasks.Add "", oRow.Key, mp_dtAddEnd, mp_dtAddStart, "", "0", mp_oControl.CurrentLayer
                            Else
                                mp_oControl.Tasks.Add "", oRow.Key, mp_dtAddStart, mp_dtAddEnd, "", "0", mp_oControl.CurrentLayer
                            End If
                            mp_oControl.SelectedTaskIndex = mp_oControl.Tasks.Count
                            mp_oControl.FireAdded E_EVENTTARGET.EVT_TASK, mp_oControl.Tasks.Count
                        Case E_ADDMODE.AT_MILESTONEADD
                            Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_lDragRowIndex)
                            mp_oControl.Milestones.Add "", oRow.Key, mp_dtAddStart, "", "0", mp_oControl.CurrentLayer
                            mp_oControl.SelectedMilestoneIndex = mp_oControl.Milestones.Count
                            mp_oControl.FireAdded E_EVENTTARGET.EVT_MILESTONE, mp_oControl.Milestones.Count
                    End Select
                End If
        End Select
    End Sub

    Private Sub mp_DragMoveSplitter(ByVal v_yMouseOperation As E_MOUSEEVENTS, ByRef X As Long, ByRef Y As Long)
        Select Case v_yMouseOperation
            Case E_MOUSEEVENTS.MouseMoveNoDrag
                mp_oToolTip.Visible = False
                If mp_oControl.AllowSplitterMove = True Then
                    mp_yMoveType = E_OPERATION.EO_SPLITTERMOVEMENT
                    mp_SetCursor (E_CURSORTYPE.CT_MOVESPLITTER)
                    mp_lDragOptions = 0
                Else
                    mp_MovingOverNothing
                End If
            Case E_MOUSEEVENTS.MouseDown
                mp_yDragType = E_OPERATION.EO_SPLITTERMOVEMENT
            Case E_MOUSEEVENTS.MouseMove
                mp_oControl.GrphLib.EraseReversibleFrames
                mp_DrawMovingReversibleFrame X, 0, X + 2, 0, E_FOCUSTYPE.FCT_VERTICALSPLITTER
            Case E_MOUSEEVENTS.MouseUp
                If X > (mp_oControl.GrphLib.Width - 10) Then
                    X = mp_oControl.GrphLib.Width - 10
                End If
                If X < 10 Then
                    X = 10
                End If
                If X > mp_oControl.Columns.Width Then
                    X = mp_oControl.Columns.Width
                    mp_oControl.f_oHScrollBar1.Value = 0
                End If
                mp_oControl.Splitter.Position = X
                mp_oControl.GrphLib.EraseReversibleFrames
        End Select
    End Sub

    Private Sub mp_DragTimeLine(ByVal v_yMouseOperation As E_MOUSEEVENTS, ByRef X As Long, ByRef Y As Long)
        Select Case v_yMouseOperation
            Case E_MOUSEEVENTS.MouseMoveNoDrag
                mp_oToolTip.Visible = False
                If mp_oControl.AllowTimeLineScroll = True Then
                    mp_yMoveType = E_OPERATION.EO_TIMELINEMOVEMENT
                    mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
                Else
                    mp_yMoveType = E_OPERATION.EO_TIMELINESELECTION
                    mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
                End If
            Case E_MOUSEEVENTS.MouseDown
                If (mp_oButton = vbLeftButton) Then
                    mp_SetCursor (E_CURSORTYPE.CT_SCROLLTIMELINE)
                    mp_yDragType = E_OPERATION.EO_TIMELINEMOVEMENT
                    mp_lFixedX = X
                Else
                    mp_yDragType = E_OPERATION.EO_TIMELINESELECTION
                End If
            Case E_MOUSEEVENTS.MouseMove
            Case E_MOUSEEVENTS.MouseUp
                If mp_oControl.f_oHScrollBar2.Enabled = False Then
                    mp_oControl.CurrentViewObject.TimeLine.f_StartDate = mp_oControl.MathLib.DateTimeAdd(mp_oControl.CurrentViewObject.f_Interval, (mp_lFixedX - X) * mp_oControl.CurrentViewObject.f_Factor, mp_oControl.CurrentViewObject.TimeLine.StartDate)
                End If
                mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
        End Select
    End Sub

    Private Sub mp_DragColumnWidth(ByVal v_yMouseOperation As E_MOUSEEVENTS, ByRef X As Long, ByRef Y As Long)
        Dim oColumn As clsColumn
        Dim bCancel As Boolean
        Select Case v_yMouseOperation
            Case E_MOUSEEVENTS.MouseMoveNoDrag
                If mp_oControl.AllowColumnSize = True Then
                    mp_yMoveType = E_OPERATION.EO_COLUMNSIZING
                    mp_SetCursor (E_CURSORTYPE.CT_COLUMNWIDTH)
                    mp_oToolTip.Visible = False
                    mp_lDragOptions = 0
                Else
                    mp_MovingOverNothing
                End If
            Case E_MOUSEEVENTS.MouseDown
                bCancel = False
                mp_oControl.FireBeginObjectSize E_EVENTTARGET.EVT_COLUMN, mp_lDragColumnIndex, bCancel
                If bCancel = False Then
                    mp_yDragType = E_OPERATION.EO_COLUMNSIZING
                Else
                    CancelUIOperations
                End If
            Case E_MOUSEEVENTS.MouseMove
                bCancel = False
                mp_oControl.GrphLib.EraseReversibleFrames
                mp_oControl.FireObjectSize E_EVENTTARGET.EVT_COLUMN, mp_lDragColumnIndex, bCancel
                If bCancel = False Then
                    mp_DrawMovingReversibleFrame X, 0, X + 2, mp_oControl.GrphLib.Height, E_FOCUSTYPE.FCT_NORMAL
                Else
                    CancelUIOperations
                End If
            Case E_MOUSEEVENTS.MouseUp
                bCancel = False
                mp_oControl.GrphLib.EraseReversibleFrames
                mp_oControl.FireEndObjectSize E_EVENTTARGET.EVT_COLUMN, mp_lDragColumnIndex, bCancel
                If bCancel = False Then
                    If X < mp_oControl.mt_BorderThickness Then
                        X = mp_oControl.mt_BorderThickness
                    End If
                    If X > mp_oControl.Splitter.Position Then
                        mp_oControl.Splitter.Position = X
                    End If
                    Set oColumn = mp_oControl.Columns.oCollection.m_oReturnArrayElement(mp_lDragColumnIndex)
                    oColumn.Width = oColumn.Width + (X - oColumn.Right)
                    If oColumn.Width < mp_oControl.MinColumnWidth Then
                        oColumn.Width = mp_oControl.MinColumnWidth
                    End If
                    If mp_oControl.Splitter.Position > mp_oControl.Columns.Width Then
                        mp_oControl.Splitter.Position = mp_oControl.Columns.Width
                        mp_oControl.f_oHScrollBar1.Value = 0
                    End If
                    mp_oControl.FireCompleteObjectSize E_EVENTTARGET.EVT_COLUMN, mp_lDragColumnIndex
                Else
                    CancelUIOperations
                End If
        End Select
    End Sub

    Private Sub mp_DragMoveRow(ByVal v_yMouseOperation As E_MOUSEEVENTS, ByRef X As Long, ByRef Y As Long)
        Dim oRow As clsRow
        Dim bCancel As Boolean
        Select Case v_yMouseOperation
            Case E_MOUSEEVENTS.MouseMoveNoDrag
                If mp_oControl.AllowRowSwap = True Then
                    mp_yMoveType = E_OPERATION.EO_ROWMOVEMENT
                    mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
                Else
                    mp_yMoveType = E_OPERATION.EO_ROWSELECTION
                    mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
                End If
            Case E_MOUSEEVENTS.MouseDown
                If mp_oControl.AllowRowSwap = True Then
                    bCancel = False
                    mp_oControl.FireBeginObjectMove E_EVENTTARGET.EVT_ROW, mp_oControl.SelectedRowIndex, bCancel
                    If mp_oButton = vbLeftButton And bCancel = False Then
                        mp_SetCursor (E_CURSORTYPE.CT_MOVEROW)
                        mp_yDragType = E_OPERATION.EO_ROWMOVEMENT
                        mp_lRowToDragIndex = mp_lReturnRowIndexByPosition(Y)
                    End If
                    If bCancel = True Then
                        CancelUIOperations
                    End If
                Else
                    mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
                    mp_yDragType = E_OPERATION.EO_ROWSELECTION
                End If
            Case E_MOUSEEVENTS.MouseMove
                mp_oControl.GrphLib.EraseReversibleFrames
                mp_DynamicRowMove (Y)
                bCancel = False
                mp_lDragRowIndex = mp_lReturnRowIndexByPosition(Y)
                mp_oControl.FireObjectMove E_EVENTTARGET.EVT_ROW, mp_lDragRowIndex, bCancel
                If bCancel = False Then
                    If mp_lDragRowIndex = 0 Then
                        mp_SetCursor (E_CURSORTYPE.CT_NODROP)
                        Exit Sub
                    Else
                        mp_SetCursor (E_CURSORTYPE.CT_MOVEROW)
                    End If
                    Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_lDragRowIndex)
                    mp_DrawMovingReversibleFrame oRow.Left, oRow.Top, oRow.Right, oRow.Bottom, E_FOCUSTYPE.FCT_NORMAL
                End If
                If bCancel = True Then
                    CancelUIOperations
                End If
            Case E_MOUSEEVENTS.MouseUp
                mp_oControl.GrphLib.EraseReversibleFrames
                bCancel = False
                mp_lRowDraggingDestinationIndex = mp_lReturnRowIndexByPosition(Y)
                mp_oControl.FireEndObjectMove E_EVENTTARGET.EVT_ROW, mp_lRowDraggingDestinationIndex, bCancel
                If mp_bCompleteDragOperation = True And bCancel = False Then
                    If mp_lRowToDragIndex = mp_lRowDraggingDestinationIndex Then
                        Exit Sub
                    End If
                    mp_oControl.SelectedRowIndex = mp_oControl.Rows.oCollection.m_lCopyAndMoveItems(mp_lRowToDragIndex, mp_lRowDraggingDestinationIndex)
                    mp_oControl.FireCompleteObjectMove E_EVENTTARGET.EVT_ROW, mp_oControl.SelectedRowIndex
                End If
                mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
                If bCancel = True Then
                    CancelUIOperations
                End If
        End Select
    End Sub

    Private Sub mp_DragRowHeight(ByVal v_yMouseOperation As E_MOUSEEVENTS, ByRef X As Long, ByRef Y As Long)
        Dim oRow As clsRow
        Dim bCancel As Boolean
        Select Case v_yMouseOperation
            Case E_MOUSEEVENTS.MouseMoveNoDrag
                If mp_oControl.AllowRowSize = True Then
                    mp_yMoveType = E_OPERATION.EO_ROWSIZING
                    mp_SetCursor (E_CURSORTYPE.CT_ROWHEIGHT)
                    mp_oToolTip.Visible = False
                Else
                    mp_MovingOverNothing
                End If
            Case E_MOUSEEVENTS.MouseDown
                bCancel = False
                mp_oControl.FireBeginObjectSize E_EVENTTARGET.EVT_ROW, mp_lDragRowIndex, bCancel
                If bCancel = False Then
                    mp_yDragType = E_OPERATION.EO_ROWSIZING
                    mp_lDragOptions = Y
                Else
                    CancelUIOperations
                End If
            Case E_MOUSEEVENTS.MouseMove
                bCancel = False
                mp_oControl.GrphLib.EraseReversibleFrames
                mp_oControl.FireObjectSize E_EVENTTARGET.EVT_ROW, mp_lDragRowIndex, bCancel
                If bCancel = False Then
                    mp_DrawMovingReversibleFrame 0, Y, mp_oControl.GrphLib.Width, Y + 2, E_FOCUSTYPE.FCT_NORMAL
                Else
                    CancelUIOperations
                End If
            Case E_MOUSEEVENTS.MouseUp
                bCancel = False
                mp_oControl.GrphLib.EraseReversibleFrames
                mp_oControl.FireEndObjectSize E_EVENTTARGET.EVT_ROW, mp_lDragRowIndex, bCancel
                If bCancel = False Then
                    Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_lDragRowIndex)
                    If (Y - mp_lDragOptions + oRow.Height) < mp_oControl.MinRowHeight Then
                        oRow.Height = mp_oControl.MinRowHeight
                    Else
                        oRow.Height = Y - mp_lDragOptions + oRow.Height
                    End If
                    mp_oControl.FireCompleteObjectSize E_EVENTTARGET.EVT_ROW, mp_lDragRowIndex
                Else
                    CancelUIOperations
                End If
        End Select
    End Sub

    Private Sub mp_DragMoveMilestone(ByVal v_yMouseOperation As E_MOUSEEVENTS, ByRef X As Long, ByRef Y As Long)
        Dim oMilestone As clsMilestone
        Dim oRow As clsRow
        Dim bCancel As Boolean
        Select Case v_yMouseOperation
            Case E_MOUSEEVENTS.MouseMoveNoDrag
                ToolTipCaption E_TOOLTIPACTION.TA_OVERMILESTONE, mp_oControl.SelectedMilestoneIndex, X, Y, dtNull, dtNull
                Set oMilestone = mp_oControl.Milestones.oCollection.m_oReturnArrayElement(mp_oControl.SelectedMilestoneIndex)
                If mp_oControl.AllowEdit = True Then
                    If oMilestone.AllowedMovement <> E_MOVEMENTTYPE.MT_MOVEMENTDISABLED Then
                        mp_yMoveType = E_OPERATION.EO_MILESTONEMOVEMENT
                        mp_SetCursor (E_CURSORTYPE.CT_MOVEMILESTONE)
                    Else
                        mp_yMoveType = E_OPERATION.EO_MILESTONESELECTION
                        mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
                    End If
                Else
                    mp_yMoveType = E_OPERATION.EO_MILESTONESELECTION
                    mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
                End If
            Case E_MOUSEEVENTS.MouseDown
                mp_bMilestoneMove = False
                bCancel = False
                If mp_oButton = vbLeftButton Then
                    If mp_oControl.SelectedMilestoneIndex = 0 Then
                        Exit Sub
                    End If
                    X = mp_fSnapX(X)
                    mp_lFixedX = X
                    Set oMilestone = mp_oControl.Milestones.oCollection.m_oReturnArrayElement(mp_oControl.SelectedMilestoneIndex)
                    mp_lOrgDragRowIndex = mp_lReturnRowIndexByPosition(Y)
                    mp_oControl.FireBeginObjectMove E_EVENTTARGET.EVT_MILESTONE, mp_oControl.SelectedMilestoneIndex, bCancel
                    If bCancel = False Then
                        mp_yDragType = E_OPERATION.EO_MILESTONEMOVEMENT
                        mp_dtAddStart = oMilestone.MilestoneDate
                    Else
                        CancelUIOperations
                    End If
                End If
            Case E_MOUSEEVENTS.MouseMove
                mp_bMilestoneMove = True
                bCancel = False
                X = mp_fSnapX(X)
                mp_lDragRowIndex = mp_lReturnRowIndexByPosition(Y)
                mp_oControl.GrphLib.EraseReversibleFrames
                If bCancel = False Then
                    mp_DynamicRowMove (Y)
                    ToolTipCaption E_TOOLTIPACTION.TA_MOVEMILESTONE, mp_oControl.SelectedMilestoneIndex, X, Y, mp_oControl.MathLib.GetDateFromXCoordinate(X), dtNull
                    If X < 0 Or X > mp_oControl.GrphLib.Width Or Y < 0 Or Y > mp_oControl.GrphLib.Height Then
                        'If mp_yOLEDragMode = E_DROPMODE.DRP_MANUAL Then
                        '    'UserControl.OLEDrag()
                        '    mp_oControl.GrphLib.DrawReversibleFrameEx()
                        '    Exit Sub
                        'Else
                        mp_SetCursor (E_CURSORTYPE.CT_NODROP)
                        mp_oControl.GrphLib.DrawReversibleFrameEx
                        'End If
                    End If
                    If mp_lDragRowIndex = 0 Then
                        mp_SetCursor (E_CURSORTYPE.CT_NODROP)
                        mp_oControl.GrphLib.DrawReversibleFrameEx
                        Exit Sub
                    End If
                    Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_lDragRowIndex)
                    Set oMilestone = mp_oControl.Milestones.oCollection.m_oReturnArrayElement(mp_oControl.SelectedMilestoneIndex)
                    If (mp_lDragRowIndex <> mp_lOrgDragRowIndex And oMilestone.AllowedMovement = E_MOVEMENTTYPE.MT_RESTRICTEDTOROW) Or (oRow.Container = False) Or (mp_oControl.mp_bDetectConflict(mp_oControl.MathLib.GetDateFromXCoordinate(X), mp_oControl.MathLib.CurrentDateTime, oRow.Key, mp_oControl.SelectedMilestoneIndex, oMilestone.LayerIndex, E_TYPE.TP_MILESTONE) = True And mp_oControl.CurrentViewObject.ClientArea.DetectConflicts = True) Then
                        mp_SetCursor (E_CURSORTYPE.CT_NODROP)
                        mp_oControl.GrphLib.DrawReversibleFrameEx
                        Exit Sub
                    End If
                    mp_SetCursor (E_CURSORTYPE.CT_MOVEMILESTONE)
                    mp_oControl.FireObjectMove E_EVENTTARGET.EVT_MILESTONE, mp_oControl.SelectedMilestoneIndex, bCancel
                    mp_DynamicTimeLineMove (X)
                    mp_DrawMovingReversibleFrame X - mp_oControl.CurrentViewObject.ClientArea.MilestoneSelectionOffset, oRow.Top, X + mp_oControl.CurrentViewObject.ClientArea.MilestoneSelectionOffset, oRow.Bottom, E_FOCUSTYPE.FCT_KEEPLEFTRIGHTBOUNDS
                Else
                    CancelUIOperations
                End If
            Case E_MOUSEEVENTS.MouseUp
                bCancel = False
                mp_oControl.GrphLib.EraseReversibleFrames
                If mp_bCompleteDragOperation = True And mp_bMilestoneMove = True Then
                    X = mp_fSnapX(X)
                    Set oMilestone = mp_oControl.Milestones.oCollection.m_oReturnArrayElement(mp_oControl.SelectedMilestoneIndex)
                    Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_lDragRowIndex)
                    mp_oControl.FireEndObjectMove E_EVENTTARGET.EVT_MILESTONE, mp_oControl.SelectedMilestoneIndex, bCancel
                    If bCancel = False Then
                        oMilestone.MilestoneDate = mp_oControl.MathLib.GetDateFromXCoordinate(X)
                        oMilestone.RowKey = oRow.Key
                        If oMilestone.MilestoneDate <> mp_dtAddStart Then
                            mp_oControl.FireCompleteObjectMove E_EVENTTARGET.EVT_MILESTONE, mp_oControl.SelectedMilestoneIndex
                        End If
                    Else
                        CancelUIOperations
                    End If
                End If
        End Select
    End Sub

    Private Sub mp_DragMoveTask(ByVal v_yMouseOperation As E_MOUSEEVENTS, ByRef X As Long, ByRef Y As Long, ByVal v_lMoveOp As E_TASKMOVETYPE)
        Dim oTask As clsTask
        Dim oRow As clsRow
        Dim bCancel As Boolean
        Select Case v_yMouseOperation
            Case E_MOUSEEVENTS.MouseMoveNoDrag
                ToolTipCaption E_TOOLTIPACTION.TA_OVERTASK, mp_oControl.SelectedTaskIndex, X, Y, dtNull, dtNull
                Set oTask = mp_oControl.Tasks.oCollection.m_oReturnArrayElement(mp_oControl.SelectedTaskIndex)
                If mp_oControl.AllowEdit = True Then
                    Select Case v_lMoveOp
                        Case E_TASKMOVETYPE.GMT_STRETCHLEFT
                            If oTask.AllowStretchLeft = True Then
                                mp_yMoveType = E_OPERATION.EO_TASKSTRETCHLEFT
                                mp_SetCursor (E_CURSORTYPE.CT_SIZETASK)
                                mp_lDragOptions = 1
                            Else
                                mp_yMoveType = E_OPERATION.EO_TASKSELECTION
                                mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
                            End If
                        Case E_TASKMOVETYPE.GMT_STRETCHRIGHT
                            If oTask.AllowStretchRight = True Then
                                mp_yMoveType = E_OPERATION.EO_TASKSTRETCHRIGHT
                                mp_SetCursor (E_CURSORTYPE.CT_SIZETASK)
                                mp_lDragOptions = 2
                            Else
                                mp_yMoveType = E_OPERATION.EO_TASKSELECTION
                                mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
                            End If
                        Case E_TASKMOVETYPE.GMT_MOVE
                            If oTask.AllowedMovement <> E_MOVEMENTTYPE.MT_MOVEMENTDISABLED Then
                                mp_yMoveType = E_OPERATION.EO_TASKMOVEMENT
                                mp_SetCursor (E_CURSORTYPE.CT_MOVETASK)
                            Else
                                mp_yMoveType = E_OPERATION.EO_TASKSELECTION
                                mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
                            End If
                    End Select
                Else
                    mp_yMoveType = E_OPERATION.EO_TASKSELECTION
                    mp_SetCursor (E_CURSORTYPE.CT_NORMAL)
                End If
            Case E_MOUSEEVENTS.MouseDown
                mp_bTaskMove = False
                bCancel = False
                mp_lDragRowIndex = mp_lReturnRowIndexByPosition(Y)
                If mp_oButton = vbLeftButton Then
                    If mp_oControl.SelectedTaskIndex = 0 Then
                        Exit Sub
                    End If
                    Set oTask = mp_oControl.Tasks.oCollection.m_oReturnArrayElement(mp_oControl.SelectedTaskIndex)
                    X = mp_fSnapX(X)
                    mp_lOrgDragRowIndex = mp_lReturnRowIndexByPosition(Y)
                    Select Case v_lMoveOp
                        Case E_TASKMOVETYPE.GMT_STRETCHLEFT
                            mp_oControl.FireBeginObjectSize E_EVENTTARGET.EVT_TASK, mp_oControl.SelectedTaskIndex, bCancel
                            If bCancel = False Then
                                mp_yDragType = E_OPERATION.EO_TASKSTRETCHLEFT
                                mp_dtAddStart = oTask.StartDate
                                ToolTipCaption E_TOOLTIPACTION.TA_STRETCHTASK, mp_oControl.SelectedTaskIndex, X, Y, dtNull, dtNull
                            Else
                                CancelUIOperations
                            End If
                        Case E_TASKMOVETYPE.GMT_STRETCHRIGHT
                            mp_oControl.FireBeginObjectSize E_EVENTTARGET.EVT_TASK, mp_oControl.SelectedTaskIndex, bCancel
                            If bCancel = False Then
                                mp_yDragType = E_OPERATION.EO_TASKSTRETCHRIGHT
                                mp_dtAddEnd = oTask.EndDate

                                ToolTipCaption E_TOOLTIPACTION.TA_STRETCHTASK, mp_oControl.SelectedTaskIndex, X, Y, dtNull, dtNull
                            Else
                                CancelUIOperations
                            End If
                        Case E_TASKMOVETYPE.GMT_MOVE
                            mp_oControl.FireBeginObjectMove E_EVENTTARGET.EVT_TASK, mp_oControl.SelectedTaskIndex, bCancel
                            If bCancel = False Then
                                mp_yDragType = E_OPERATION.EO_TASKMOVEMENT
                                mp_lFixedX = X
                                mp_lDeltaToLeft = X - mp_oControl.MathLib.GetXCoordinateFromDate(oTask.StartDate)
                                mp_lDeltaToRight = mp_oControl.MathLib.GetXCoordinateFromDate(oTask.EndDate) - X
                                mp_dtAddStart = oTask.StartDate
                                mp_dtAddEnd = oTask.EndDate
                                mp_sAddRowKey = oTask.RowKey
                                mp_lFixedDeltaDates = mp_oControl.MathLib.DateTimeDiff("s", mp_dtAddStart, mp_dtAddEnd)
                            Else
                                CancelUIOperations
                            End If
                    End Select
                End If
            Case E_MOUSEEVENTS.MouseMove
                mp_bTaskMove = True
                bCancel = False
                X = mp_fSnapX(X)
                mp_lDragRowIndex = mp_lReturnRowIndexByPosition(Y)
                mp_oControl.GrphLib.EraseReversibleFrames
                Select Case v_lMoveOp
                    Case E_TASKMOVETYPE.GMT_STRETCHLEFT
                        Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_lOrgDragRowIndex)
                        Set oTask = mp_oControl.Tasks.oCollection.m_oReturnArrayElement(mp_oControl.SelectedTaskIndex)
                        mp_dtAddEnd = oTask.EndDate
                        mp_oControl.FireObjectSize E_EVENTTARGET.EVT_TASK, mp_oControl.SelectedTaskIndex, bCancel
                        If bCancel = False Then
                            If (mp_lDragRowIndex <> mp_lOrgDragRowIndex) Or (mp_oControl.mp_bDetectConflict(mp_oControl.MathLib.GetDateFromXCoordinate(X), mp_dtAddEnd, oRow.Key, mp_oControl.SelectedTaskIndex, oTask.LayerIndex) = True And mp_oControl.CurrentViewObject.ClientArea.DetectConflicts = True) Then
                                mp_SetCursor (E_CURSORTYPE.CT_NODROP)
                            Else
                                mp_SetCursor (E_CURSORTYPE.CT_SIZETASK)
                            End If
                            mp_DynamicTimeLineMove (X)
                            ToolTipCaption E_TOOLTIPACTION.TA_STRETCHTASK, mp_oControl.SelectedTaskIndex, X, Y, dtNull, dtNull
                            mp_DrawMovingReversibleFrame X, oRow.Top, mp_oControl.MathLib.GetXCoordinateFromDate(mp_dtAddEnd), oRow.Bottom, E_FOCUSTYPE.FCT_KEEPLEFTRIGHTBOUNDS
                        Else
                            CancelUIOperations
                        End If
                    Case E_TASKMOVETYPE.GMT_STRETCHRIGHT
                        Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_lOrgDragRowIndex)
                        Set oTask = mp_oControl.Tasks.oCollection.m_oReturnArrayElement(mp_oControl.SelectedTaskIndex)
                        mp_dtAddStart = oTask.StartDate
                        mp_oControl.FireObjectSize E_EVENTTARGET.EVT_TASK, mp_oControl.SelectedTaskIndex, bCancel
                        If bCancel = False Then
                            If (mp_lDragRowIndex <> mp_lOrgDragRowIndex) Or (mp_oControl.mp_bDetectConflict(mp_dtAddStart, mp_oControl.MathLib.GetDateFromXCoordinate(X), oRow.Key, mp_oControl.SelectedTaskIndex, oTask.LayerIndex) = True And mp_oControl.CurrentViewObject.ClientArea.DetectConflicts = True) Then
                                mp_SetCursor (E_CURSORTYPE.CT_NODROP)
                            Else
                                mp_SetCursor (E_CURSORTYPE.CT_SIZETASK)
                            End If
                            mp_DynamicTimeLineMove (X)
                            ToolTipCaption E_TOOLTIPACTION.TA_STRETCHTASK, mp_oControl.SelectedTaskIndex, X, Y, dtNull, dtNull
                            mp_DrawMovingReversibleFrame mp_oControl.MathLib.GetXCoordinateFromDate(mp_dtAddStart), oRow.Top, X, oRow.Bottom, E_FOCUSTYPE.FCT_KEEPLEFTRIGHTBOUNDS
                        Else
                            CancelUIOperations
                        End If
                    Case E_TASKMOVETYPE.GMT_MOVE
                        mp_oControl.FireObjectMove E_EVENTTARGET.EVT_TASK, mp_oControl.SelectedTaskIndex, bCancel
                        If bCancel = False Then
                            mp_DynamicRowMove (Y)
                            ToolTipCaption E_TOOLTIPACTION.TA_MOVETASK, mp_oControl.SelectedTaskIndex, X, Y, mp_oControl.MathLib.GetDateFromXCoordinate(X - mp_lDeltaToLeft), mp_oControl.MathLib.GetDateFromXCoordinate(X + mp_lDeltaToRight)
                            If X < 0 Or X > mp_oControl.GrphLib.Width Or Y < 0 Or Y > mp_oControl.GrphLib.Height Then
                                'If mp_yOLEDragMode = E_DROPMODE.DRP_MANUAL Then
                                '    'UserControl.OLEDrag
                                '    mp_oControl.GrphLib.DrawReversibleFrameEx
                                '    Exit Sub
                                'Else
                                mp_SetCursor (E_CURSORTYPE.CT_NODROP)
                                mp_oControl.GrphLib.DrawReversibleFrameEx
                                'Exit Sub
                                'End If
                            End If
                            If mp_lDragRowIndex = 0 Then
                                mp_SetCursor (E_CURSORTYPE.CT_NODROP)
                                mp_oControl.GrphLib.DrawReversibleFrameEx
                                Exit Sub
                            End If
                            Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_lDragRowIndex)
                            Set oTask = mp_oControl.Tasks.oCollection.m_oReturnArrayElement(mp_oControl.SelectedTaskIndex)
                            If (mp_lDragRowIndex <> mp_lOrgDragRowIndex And oTask.AllowedMovement = E_MOVEMENTTYPE.MT_RESTRICTEDTOROW) Or (oRow.Container = False) Or (mp_oControl.mp_bDetectConflict(mp_oControl.MathLib.GetDateFromXCoordinate(X - mp_lDeltaToLeft), mp_oControl.MathLib.GetDateFromXCoordinate(X + mp_lDeltaToRight), oRow.Key, mp_oControl.SelectedTaskIndex, oTask.LayerIndex) = True And mp_oControl.CurrentViewObject.ClientArea.DetectConflicts = True) Then
                                mp_SetCursor (E_CURSORTYPE.CT_NODROP)
                                mp_oControl.GrphLib.DrawReversibleFrameEx
                                Exit Sub
                            End If
                            mp_SetCursor (E_CURSORTYPE.CT_MOVETASK)
                            mp_DynamicTimeLineMove (X)
                            mp_DrawMovingReversibleFrame X - mp_lDeltaToLeft, oRow.Top, X + mp_lDeltaToRight, oRow.Bottom, E_FOCUSTYPE.FCT_KEEPLEFTRIGHTBOUNDS
                        Else
                            CancelUIOperations
                        End If
                End Select
            Case E_MOUSEEVENTS.MouseUp
                bCancel = False
                mp_oControl.GrphLib.EraseReversibleFrames
                If mp_bCompleteDragOperation = True And mp_bTaskMove = True Then
                    X = mp_fSnapX(X)
                    Set oTask = mp_oControl.Tasks.oCollection.m_oReturnArrayElement(mp_oControl.SelectedTaskIndex)
                    Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_lDragRowIndex)
                    Select Case v_lMoveOp
                        Case E_TASKMOVETYPE.GMT_STRETCHLEFT
                            If mp_dtAddEnd < mp_oControl.MathLib.GetDateFromXCoordinate(X) Then Exit Sub
                            mp_oControl.FireEndObjectSize E_EVENTTARGET.EVT_TASK, mp_oControl.SelectedTaskIndex, bCancel
                            If bCancel = False Then
                                oTask.StartDate = mp_oControl.MathLib.GetDateFromXCoordinate(X)
                                oTask.EndDate = mp_dtAddEnd
                                If mp_oControl.CurrentViewObject.ClientArea.Grid.SnapToGrid = True Then
                                    oTask.StartDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.f_Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.f_Factor, oTask.StartDate)
                                    oTask.EndDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.f_Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.f_Factor, oTask.EndDate)
                                End If
                                If oTask.StartDate <> mp_dtAddStart Then
                                    mp_oControl.FireCompleteObjectSize E_EVENTTARGET.EVT_TASK, mp_oControl.SelectedTaskIndex
                                End If
                            Else
                                CancelUIOperations
                            End If
                        Case E_TASKMOVETYPE.GMT_STRETCHRIGHT
                            If mp_dtAddStart > mp_oControl.MathLib.GetDateFromXCoordinate(X) Then Exit Sub
                            mp_oControl.FireEndObjectSize E_EVENTTARGET.EVT_TASK, mp_oControl.SelectedTaskIndex, bCancel
                            If bCancel = False Then
                                oTask.StartDate = mp_dtAddStart
                                oTask.EndDate = mp_oControl.MathLib.GetDateFromXCoordinate(X)
                                If mp_oControl.CurrentViewObject.ClientArea.Grid.SnapToGrid = True Then
                                    oTask.StartDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.f_Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.f_Factor, oTask.StartDate)
                                    oTask.EndDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.f_Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.f_Factor, oTask.EndDate)
                                End If
                                If oTask.EndDate <> mp_dtAddEnd Then
                                    mp_oControl.FireCompleteObjectSize E_EVENTTARGET.EVT_TASK, mp_oControl.SelectedTaskIndex
                                End If
                            Else
                                CancelUIOperations
                            End If
                        Case E_TASKMOVETYPE.GMT_MOVE
                            mp_oControl.FireEndObjectMove E_EVENTTARGET.EVT_TASK, mp_oControl.SelectedTaskIndex, bCancel
                            If bCancel = False Then
                                oTask.StartDate = mp_oControl.MathLib.GetDateFromXCoordinate(X - mp_lDeltaToLeft)
                                oTask.EndDate = mp_oControl.MathLib.GetDateFromXCoordinate(X + mp_lDeltaToRight)
                                If mp_oControl.CurrentViewObject.ClientArea.Grid.SnapToGrid = True Then
                                    oTask.StartDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.f_Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.f_Factor, oTask.StartDate)
                                    oTask.EndDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.f_Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.f_Factor, oTask.EndDate)
                                End If
                                If mp_lFixedDeltaDates <> mp_oControl.MathLib.DateTimeDiff("s", mp_dtAddStart, mp_dtAddEnd) Then
                                    mp_oControl.mp_ErrorReport 50169, "Task duration is inconsistent", "ActiveGanttVBCtl.mp_DragMoveTask"
                                End If
                                oTask.RowKey = oRow.Key
                                If oTask.StartDate <> mp_dtAddStart Or oTask.EndDate <> mp_dtAddEnd Or oTask.RowKey <> mp_sAddRowKey Then
                                    mp_oControl.FireCompleteObjectMove E_EVENTTARGET.EVT_TASK, mp_oControl.SelectedTaskIndex
                                End If
                            Else
                                CancelUIOperations
                            End If
                    End Select
                End If
        End Select
    End Sub

    Private Sub mp_MovingOverNothing()
        mp_yMoveType = E_OPERATION.EO_NONE
        'mp_oControl.Cursor = Cursors.Default
        'RaiseEvent CursorPositionChanged("", E_TYPE.TP_NONE, 0, False)
    End Sub

    Private Sub mp_DynamicTimeLineMove(ByVal v_X As Long)
        If v_X > mp_oControl.CurrentViewObject.TimeLine.f_lEnd Then
            If mp_oControl.f_oHScrollBar2.Enabled = False Then
                mp_oControl.CurrentViewObject.TimeLine.f_StartDate = mp_oControl.MathLib.DateTimeAdd(mp_oControl.CurrentViewObject.f_ScrollInterval, mp_oControl.CurrentViewObject.f_Factor, mp_oControl.CurrentViewObject.TimeLine.StartDate)
            Else
                If mp_oControl.f_oHScrollBar2.Value < mp_oControl.f_oHScrollBar2.Max Then
                    mp_oControl.f_oHScrollBar2.Value = mp_oControl.f_oHScrollBar2.Value + 1
                End If
            End If
            mp_oControl.Redraw
        End If
        If v_X < mp_oControl.CurrentViewObject.TimeLine.f_lStart Then
            If mp_oControl.f_oHScrollBar2.Enabled = False Then
                mp_oControl.CurrentViewObject.TimeLine.f_StartDate = mp_oControl.MathLib.DateTimeAdd(mp_oControl.CurrentViewObject.f_ScrollInterval, -mp_oControl.CurrentViewObject.f_Factor, mp_oControl.CurrentViewObject.TimeLine.StartDate)
            Else
                If mp_oControl.f_oHScrollBar2.Value > 0 Then
                    mp_oControl.f_oHScrollBar2.Value = mp_oControl.f_oHScrollBar2.Value - 1
                End If
            End If
            mp_oControl.Redraw
        End If
    End Sub

    Private Sub mp_DynamicRowMove(ByVal v_Y As Long)
        If v_Y < mp_oControl.CurrentViewObject.TimeLine.Bottom Then
            If mp_oControl.CurrentViewObject.ClientArea.FirstVisibleRow > 1 Then
                mp_oControl.CurrentViewObject.ClientArea.FirstVisibleRow = mp_oControl.CurrentViewObject.ClientArea.FirstVisibleRow - 1
                mp_oControl.f_oVScrollBar.Value = mp_oControl.f_oVScrollBar.Value - 1
                mp_oControl.Redraw
                Exit Sub
            End If
        End If
        If v_Y > mp_oControl.CurrentViewObject.ClientArea.Bottom Then
            If mp_oControl.f_oVScrollBar.Value < mp_oControl.f_oVScrollBar.Max Then
                mp_oControl.CurrentViewObject.ClientArea.FirstVisibleRow = mp_oControl.CurrentViewObject.ClientArea.FirstVisibleRow + 1
                mp_oControl.f_oVScrollBar.Value = mp_oControl.f_oVScrollBar.Value + 1
                mp_oControl.Redraw
            End If
        End If
    End Sub

    Private Sub mp_DynamicColumnMove(ByVal v_X As Long)
        If v_X < mp_oControl.mt_LeftMargin Then
            If mp_oControl.f_oHScrollBar1.Value > 20 Then
                mp_oControl.f_oHScrollBar1.Value = mp_oControl.f_oHScrollBar1.Value - 20
            Else
                mp_oControl.f_oHScrollBar1.Value = 0
            End If
            mp_oControl.Redraw
            Exit Sub
        End If
        If v_X > mp_oControl.Splitter.Left Then
            If mp_oControl.f_oHScrollBar1.Value < (mp_oControl.f_oHScrollBar1.Max - 20) Then
                mp_oControl.f_oHScrollBar1.Value = mp_oControl.f_oHScrollBar1.Value + 20
            Else
                mp_oControl.f_oHScrollBar1.Value = mp_oControl.f_oHScrollBar1.Max
            End If
            mp_oControl.Redraw
            Exit Sub
        End If
    End Sub

    Private Sub mp_DrawMovingReversibleFrame(ByVal v_X1 As Long, ByVal v_Y1 As Long, ByVal v_X2 As Long, ByVal v_Y2 As Long, ByVal v_yFocusType As E_FOCUSTYPE)
        mp_oControl.GrphLib.f_FocusLeft = v_X1
        mp_oControl.GrphLib.f_FocusTop = v_Y1
        mp_oControl.GrphLib.f_FocusRight = v_X2
        mp_oControl.GrphLib.f_FocusBottom = v_Y2
        Select Case v_yFocusType
            Case E_FOCUSTYPE.FCT_NORMAL
            Case E_FOCUSTYPE.FCT_KEEPLEFTRIGHTBOUNDS
                If mp_oControl.GrphLib.f_FocusLeft < mp_oControl.CurrentViewObject.TimeLine.f_lStart Then
                    mp_oControl.GrphLib.f_FocusLeft = mp_oControl.CurrentViewObject.TimeLine.f_lStart
                End If
                If mp_oControl.GrphLib.f_FocusRight > mp_oControl.CurrentViewObject.TimeLine.f_lEnd Then
                    mp_oControl.GrphLib.f_FocusRight = mp_oControl.CurrentViewObject.TimeLine.f_lEnd
                End If
            Case E_FOCUSTYPE.FCT_ADD
                If mp_oControl.GrphLib.f_FocusLeft < mp_oControl.CurrentViewObject.TimeLine.f_lStart Then
                    mp_oControl.GrphLib.f_FocusLeft = mp_oControl.CurrentViewObject.TimeLine.f_lStart
                End If
                If mp_oControl.GrphLib.f_FocusRight < mp_oControl.GrphLib.f_FocusLeft Then
                    mp_oControl.GrphLib.f_FocusRight = mp_oControl.GrphLib.f_FocusLeft
                    mp_oControl.GrphLib.f_FocusLeft = v_X2
                End If

            Case E_FOCUSTYPE.FCT_VERTICALSPLITTER
                If mp_oControl.GrphLib.f_FocusLeft >= mp_oControl.Splitter.Right Then
                    mp_oControl.GrphLib.f_FocusBottom = mp_oControl.CurrentViewObject.ClientArea.Bottom
                Else
                    mp_oControl.GrphLib.f_FocusBottom = mp_oControl.mt_TableBottom
                End If
        End Select
        mp_oControl.GrphLib.DrawReversibleFrameEx
    End Sub

Private Function mp_bOverSplitter(ByVal X As Single, ByVal Y As Single) As Boolean
    If mp_oControl.Columns.Count = 0 Then
        mp_bOverSplitter = False
        Exit Function
    End If
    If X >= (mp_oControl.Splitter.Right - 6) And X <= mp_oControl.Splitter.Right And Y < mp_oControl.GrphLib.Height() Then
        mp_bOverSplitter = True
    Else
        mp_bOverSplitter = False
    End If
End Function

Private Function mp_bOverEmptySpace(ByVal Y As Single) As Boolean
    If Y > mp_oControl.Rows.TopOffset Then
        mp_bOverEmptySpace = True
    Else
        mp_bOverEmptySpace = False
    End If
End Function

Private Function mp_bOverTimeLine(ByVal X As Single, ByVal Y As Single) As Boolean
    If X >= mp_oControl.CurrentViewObject.TimeLine.f_lStart And X <= mp_oControl.CurrentViewObject.TimeLine.f_lEnd And Y <= mp_oControl.CurrentViewObject.TimeLine.Bottom And Y >= mp_oControl.CurrentViewObject.TimeLine.Top Then
        mp_bOverTimeLine = True
    Else
        mp_bOverTimeLine = False
    End If
End Function

Private Function mp_bOverClientArea(ByVal X As Single, ByVal Y As Single) As Boolean
    If X >= mp_oControl.CurrentViewObject.TimeLine.f_lStart And X <= mp_oControl.CurrentViewObject.TimeLine.f_lEnd And Y >= mp_oControl.CurrentViewObject.ClientArea.Top Then
        mp_bOverClientArea = True
    Else
        mp_bOverClientArea = False
    End If
End Function

Private Function mp_bOverColumnDivision(ByVal X As Single, ByVal Y As Single) As Boolean
    Dim lIndex As Long
    Dim oColumn As clsColumn
    If Not (X <= mp_oControl.Splitter.Left And Y <= mp_oControl.CurrentViewObject.TimeLine.Bottom) Then
        mp_bOverColumnDivision = False
        Exit Function
    End If
    For lIndex = 1 To mp_oControl.Columns.Count
        Set oColumn = mp_oControl.Columns.oCollection.m_oReturnArrayElement(lIndex)
        If X >= (oColumn.Right - 3) And X <= oColumn.Right Then
            mp_lDragColumnIndex = lIndex
            mp_bOverColumnDivision = True
            Exit Function
        End If
    Next lIndex
    mp_bOverColumnDivision = False
End Function

Private Function mp_bOverSelectedColumn(ByVal X As Single, ByVal Y As Single) As Boolean
    Dim oColumn As clsColumn
    If mp_oControl.SelectedColumnIndex = 0 Or mp_oControl.Columns.Count = 0 Then
        mp_bOverSelectedColumn = False
        Exit Function
    End If
    Set oColumn = mp_oControl.Columns.oCollection.m_oReturnArrayElement(mp_oControl.SelectedColumnIndex)
    If X >= oColumn.Left And X <= oColumn.Right And Y >= oColumn.Top And Y <= oColumn.Bottom Then
        mp_bOverSelectedColumn = True
    Else
        mp_bOverSelectedColumn = False
    End If
End Function

Private Function mp_bOverColumn(ByVal X As Single, ByVal Y As Single) As Boolean
    Dim oColumn As clsColumn
    Dim lIndex As Long
    If Not (X <= mp_oControl.Splitter.Left And Y <= mp_oControl.CurrentViewObject.TimeLine.Bottom) Then
        mp_lVisColumnIndex = 0
        mp_bOverColumn = False
        Exit Function
    End If
    For lIndex = 1 To mp_oControl.Columns.Count
        Set oColumn = mp_oControl.Columns.oCollection.m_oReturnArrayElement(lIndex)
        If oColumn.Visible = True Then
            If X >= oColumn.Left And X <= oColumn.Right And Y >= oColumn.Top And Y <= oColumn.Bottom Then
                mp_lVisColumnIndex = lIndex
                mp_bOverColumn = True
                Exit Function
            End If
        End If
    Next lIndex
    mp_lVisColumnIndex = 0
    mp_bOverColumn = False
End Function

Private Function mp_bOverRowDivision(ByVal X As Single, ByVal Y As Single) As Boolean
    Dim lIndex As Long
    Dim oRow As clsRow
    If X > mp_oControl.Splitter.Left Or mp_oControl.Rows.Count = 0 Then
        mp_bOverRowDivision = False
        Exit Function
    End If
    For lIndex = mp_oControl.CurrentViewObject.ClientArea.FirstVisibleRow To mp_oControl.CurrentViewObject.ClientArea.LastVisibleRow
        Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex)
        If Y >= oRow.Bottom And Y <= oRow.Bottom + 5 Then
            mp_lDragRowIndex = lIndex
            mp_bOverRowDivision = True
            Exit Function
        End If
    Next lIndex
    mp_bOverRowDivision = False
End Function

Private Function mp_bOverSelectedRow(ByVal X As Single, ByVal Y As Single) As Boolean
    Dim oRow As clsRow
    If mp_oControl.SelectedRowIndex = 0 Or mp_oControl.Rows.Count = 0 Then
        mp_bOverSelectedRow = False
        Exit Function
    End If
    Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_oControl.SelectedRowIndex)
    If oRow.MergeCells = True Then
        If X >= oRow.Left And X <= oRow.Right And Y >= oRow.Top And Y <= oRow.Bottom Then
            mp_bOverSelectedRow = True
        Else
            mp_bOverSelectedRow = False
        End If
    Else
        If X >= oRow.Left And X <= oRow.Right And Y >= oRow.Top And Y <= oRow.Bottom Then
            If mp_oControl.SelectedCellIndex = mp_lReturnCellIndexByPosition(X) Then
                mp_bOverSelectedRow = True
            Else
                mp_bOverSelectedRow = False
            End If
        Else
            mp_bOverSelectedRow = False
        End If
    End If
End Function

Private Function mp_bOverRow(ByVal X As Single, ByVal Y As Single) As Boolean
    Dim oRow As clsRow
    Dim lIndex As Long
    If Not (X <= mp_oControl.CurrentViewObject.TimeLine.f_lStart And Y > mp_oControl.CurrentViewObject.TimeLine.Bottom) Then
        mp_lVisRowIndex = 0
        mp_lVisCellIndex = 0
        mp_bOverRow = False
        Exit Function
    End If
    For lIndex = 1 To mp_oControl.Rows.Count
        Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex)
        If oRow.Visible = True Then
            If X >= oRow.Left And X <= oRow.Right And Y >= oRow.Top And Y <= oRow.Bottom Then
                mp_lVisRowIndex = lIndex
                If oRow.MergeCells = False Then
                    mp_lVisCellIndex = mp_lReturnCellIndexByPosition(X)
                Else
                    mp_lVisCellIndex = 0
                End If
                mp_bOverRow = True
                Exit Function
            End If
        End If
    Next lIndex
    mp_lVisRowIndex = 0
    mp_lVisCellIndex = 0
    mp_bOverRow = False
End Function

Private Function mp_iOverSelectedTask(ByVal X As Single, ByVal Y As Single) As E_TASKMOVETYPE
    Dim oSelectedTask As clsTask
    If X < mp_oControl.CurrentViewObject.TimeLine.f_lStart Then
        mp_iOverSelectedTask = E_TASKMOVETYPE.GMT_NONE
        Exit Function
    End If
    If X > mp_oControl.CurrentViewObject.TimeLine.f_lEnd Then
        mp_iOverSelectedTask = E_TASKMOVETYPE.GMT_NONE
        Exit Function
    End If
    If mp_lVisGIIndex = 0 Or mp_oControl.SelectedTaskIndex = 0 Then
        mp_iOverSelectedTask = E_TASKMOVETYPE.GMT_NONE
        Exit Function
    End If
    If mp_lVisGIIndex <> mp_oControl.SelectedTaskIndex Then
        mp_iOverSelectedTask = E_TASKMOVETYPE.GMT_NONE
        Exit Function
    End If
    Set oSelectedTask = mp_oControl.Tasks.oCollection.m_oReturnArrayElement(mp_oControl.SelectedTaskIndex)
    If X >= oSelectedTask.LeftTrim And X <= oSelectedTask.RightTrim And Y >= oSelectedTask.Top And Y <= oSelectedTask.Bottom And mp_bInCurrentLayer(oSelectedTask.LayerIndex) Then
        If X >= oSelectedTask.Left And X <= oSelectedTask.Left + 2 Then
            If oSelectedTask.f_bLeftVisible = True Then
                mp_iOverSelectedTask = E_TASKMOVETYPE.GMT_STRETCHLEFT
            Else
                mp_iOverSelectedTask = E_TASKMOVETYPE.GMT_MOVE
            End If
            Exit Function
        End If
        If X >= oSelectedTask.Right - 2 And X <= oSelectedTask.Right Then
            If oSelectedTask.f_bRightVisible = True Then
                mp_iOverSelectedTask = E_TASKMOVETYPE.GMT_STRETCHRIGHT
            Else
                mp_iOverSelectedTask = E_TASKMOVETYPE.GMT_MOVE
            End If
            Exit Function
        End If
        mp_iOverSelectedTask = E_TASKMOVETYPE.GMT_MOVE
        Exit Function
    End If
    mp_iOverSelectedTask = E_TASKMOVETYPE.GMT_NONE
End Function

Private Function mp_bOverTask(ByVal X As Single, ByVal Y As Single) As Boolean
    Dim oTask As clsTask
    Dim lIndex As Long
    For lIndex = 1 To mp_oControl.Tasks.Count
        Set oTask = mp_oControl.Tasks.oCollection.m_oReturnArrayElement(lIndex)
        If oTask.Visible = True And mp_bInCurrentLayer(oTask.LayerIndex) Then
            If X >= oTask.LeftTrim And X <= oTask.RightTrim And Y >= oTask.Top And Y <= oTask.Bottom Then
                mp_lVisGIIndex = lIndex
                mp_bOverTask = True
                Exit Function
            End If
        End If
    Next lIndex
    mp_lVisGIIndex = 0
    mp_bOverTask = False
End Function

Private Function mp_bOverMilestone(ByVal X As Single, ByVal Y As Single) As Boolean
    Dim oMilestone As clsMilestone
    Dim lIndex As Long
    For lIndex = 1 To mp_oControl.Milestones.Count
        Set oMilestone = mp_oControl.Milestones.oCollection.m_oReturnArrayElement(lIndex)
        If oMilestone.Visible = True And mp_bInCurrentLayer(oMilestone.LayerIndex) Then
            If X >= oMilestone.LeftTrim And X <= oMilestone.RightTrim And Y >= oMilestone.Top And Y <= oMilestone.Bottom Then
                mp_lVisMilestoneIndex = lIndex
                mp_bOverMilestone = True
                Exit Function
            End If
        End If
    Next lIndex
    mp_lVisMilestoneIndex = 0
    mp_bOverMilestone = False
End Function

Private Function mp_bOverSelectedMilestone(ByVal X As Single, ByVal Y As Single) As Boolean
    Dim oMilestone As clsMilestone
    If X < mp_oControl.CurrentViewObject.TimeLine.f_lStart Then Exit Function
    If X > mp_oControl.CurrentViewObject.TimeLine.f_lEnd Then Exit Function
    If mp_lVisMilestoneIndex = 0 Or mp_oControl.SelectedMilestoneIndex = 0 Then Exit Function
    If mp_lVisMilestoneIndex <> mp_oControl.SelectedMilestoneIndex Then Exit Function
    Set oMilestone = mp_oControl.Milestones.oCollection.m_oReturnArrayElement(mp_oControl.SelectedMilestoneIndex)
    If X >= oMilestone.LeftTrim And X <= oMilestone.RightTrim And Y >= oMilestone.Top And Y <= oMilestone.Bottom And mp_bInCurrentLayer(oMilestone.LayerIndex) Then
        mp_bOverSelectedMilestone = True
        Exit Function
    End If
End Function

Private Function mp_bInCurrentLayer(ByVal sLayer As String) As Boolean
    If (mp_oControl.EnableObjects = E_ENABLEOBJECTS.EO_ALLLAYERS) Then
        mp_bInCurrentLayer = True
        Exit Function
    Else
        Dim lLayerIndex As Long
        Dim lCurrentLayerIndex As Long
        lLayerIndex = mp_oControl.Layers.oCollection.m_lReturnIndex(sLayer, True)
        lCurrentLayerIndex = mp_oControl.Layers.oCollection.m_lReturnIndex(mp_oControl.CurrentLayer, True)
        If (lLayerIndex <> lCurrentLayerIndex) Then
            mp_bInCurrentLayer = False
            Exit Function
        Else
            mp_bInCurrentLayer = True
            Exit Function
        End If
    End If
End Function

Private Function mp_fSnapX(ByVal X As Single) As Single
    Dim dtDate As Date
    If mp_oControl.CurrentViewObject.ClientArea.Grid.SnapToGrid = False Then
        mp_fSnapX = X
        Exit Function
    End If
    dtDate = mp_oControl.MathLib.GetDateFromXCoordinate(X)
    dtDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.f_Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.f_Factor, dtDate)
    mp_fSnapX = mp_oControl.MathLib.GetXCoordinateFromDate(dtDate)
End Function

Friend Function mp_lReturnRowIndexByPosition(ByVal Y As Long) As Long
    Dim oRow As clsRow
    Dim lVisRowIndex As Long
    If (mp_oControl.Rows.Count = 0) Then
        mp_lReturnRowIndexByPosition = 0
        Exit Function
    End If
    For lVisRowIndex = mp_oControl.CurrentViewObject.ClientArea.FirstVisibleRow To mp_oControl.CurrentViewObject.ClientArea.LastVisibleRow
        Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(lVisRowIndex)
        If Y >= oRow.Top And Y <= oRow.Bottom And oRow.Visible = True Then
            mp_lReturnRowIndexByPosition = lVisRowIndex
            Exit Function
        End If
    Next lVisRowIndex
End Function

Friend Function mp_lReturnVisibleRowIndexByKey(ByVal Key As String) As Long
    Dim lVisRowIndex As Long
    lVisRowIndex = mp_oControl.Rows.oCollection.m_lFindIndexByKey(Key)
    If lVisRowIndex >= mp_oControl.CurrentViewObject.ClientArea.FirstVisibleRow And lVisRowIndex <= mp_oControl.CurrentViewObject.ClientArea.LastVisibleRow Then
        mp_lReturnVisibleRowIndexByKey = lVisRowIndex
    Else
        mp_lReturnVisibleRowIndexByKey = 0
    End If
End Function

Friend Function mp_lReturnCellIndexByPosition(ByVal X As Long) As Long
    Dim oColumn As clsColumn
    Dim lIndex As Long
    For lIndex = 1 To mp_oControl.Columns.Count
        Set oColumn = mp_oControl.Columns.oCollection.m_oReturnArrayElement(lIndex)
        If X > oColumn.Left And X < oColumn.Right Then
            mp_lReturnCellIndexByPosition = lIndex
            Exit Function
        End If
    Next lIndex
    mp_lReturnCellIndexByPosition = -1
End Function

Private Sub mp_SetCursor(ByVal v_iCursorType As E_CURSORTYPE)
    mp_bCompleteDragOperation = True
    mp_oControl.f_SetCursor v_iCursorType
End Sub

Public Sub CancelUIOperations()
    mp_yDragType = E_OPERATION.EO_NONE
    mp_yMoveType = E_OPERATION.EO_NONE
    mp_yClickType = E_OPERATION.EO_NONE
End Sub

    Public Sub ToolTipCaption(ByVal ToolTipAction As E_TOOLTIPACTION, ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal StartDate As Date, ByVal EndDate As Date)
        Dim sCaption As String
        Dim bDisplay As Boolean
        bDisplay = True
        Dim oTask As clsTask
        Dim oMilestone As clsMilestone
        Select Case ToolTipAction
            Case E_TOOLTIPACTION.TA_CLIENTAREA
                sCaption = mp_oControl.StrLib.StrFormat(mp_oControl.MathLib.GetDateFromXCoordinate(X), mp_oControl.CurrentViewObject.ClientArea.ToolTipFormat)
            Case E_TOOLTIPACTION.TA_TASKADD
                sCaption = mp_oControl.StrLib.StrFormat(StartDate, mp_oControl.CurrentViewObject.ClientArea.ToolTipFormat) & " - " & mp_oControl.StrLib.StrFormat(EndDate, mp_oControl.CurrentViewObject.ClientArea.ToolTipFormat)
            Case E_TOOLTIPACTION.TA_STRETCHTASK
                sCaption = mp_oControl.StrLib.StrFormat(mp_oControl.MathLib.GetDateFromXCoordinate(X), mp_oControl.CurrentViewObject.ClientArea.ToolTipFormat)
            Case E_TOOLTIPACTION.TA_MOVETASK
                sCaption = mp_oControl.StrLib.StrFormat(StartDate, mp_oControl.CurrentViewObject.ClientArea.ToolTipFormat) & " - " & mp_oControl.StrLib.StrFormat(EndDate, mp_oControl.CurrentViewObject.ClientArea.ToolTipFormat)
            Case E_TOOLTIPACTION.TA_OVERTASK
                Set oTask = mp_oControl.Tasks.oCollection.m_oReturnArrayElementKey(Index)
                sCaption = mp_oControl.StrLib.StrFormat(oTask.StartDate, mp_oControl.CurrentViewObject.ClientArea.ToolTipFormat) & " - " & mp_oControl.StrLib.StrFormat(oTask.EndDate, mp_oControl.CurrentViewObject.ClientArea.ToolTipFormat)
            Case E_TOOLTIPACTION.TA_MOVEMILESTONE
                sCaption = mp_oControl.StrLib.StrFormat(StartDate, mp_oControl.CurrentViewObject.ClientArea.ToolTipFormat)
            Case E_TOOLTIPACTION.TA_OVERMILESTONE
                Set oMilestone = mp_oControl.Milestones.oCollection.m_oReturnArrayElementKey(Index)
                sCaption = mp_oControl.StrLib.StrFormat(oMilestone.MilestoneDate, mp_oControl.CurrentViewObject.ClientArea.ToolTipFormat)
        End Select
        If mp_oControl.CurrentViewObject.ClientArea.ToolTipsVisible = True Then
            mp_oControl.FireToolTip ToolTipAction, sCaption, StartDate, EndDate, Index, bDisplay
        End If
        'RaiseEvent CursorPositionChanged(sCaption, E_TYPE.TP_TASK, 0, False)
        If mp_oControl.CurrentViewObject.ClientArea.ToolTipsVisible = True And bDisplay = True Then
            mp_oToolTip.Caption = sCaption
            mp_oToolTip.Top = Y - 20
            If X + mp_oToolTip.Width > mp_oControl.CurrentViewObject.TimeLine.f_lEnd Then
                mp_oToolTip.Left = mp_oControl.CurrentViewObject.TimeLine.f_lEnd - mp_oToolTip.Width
            Else
                mp_oToolTip.Left = X
            End If
            mp_oToolTip.Visible = True
        Else
            mp_oToolTip.Visible = False
        End If
    End Sub

