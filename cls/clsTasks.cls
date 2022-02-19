VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTasks"
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

'// ---------------------------------------------------------------------------------------------------------------------
'// Initialization & Destruction
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oCollection = New clsCollectionBase
    mp_oCollection.Initialize Value, "Task"
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

Public Property Get Item(ByVal Index As String) As clsTask
    Set Item = mp_oCollection.m_oItem(Index, 50215, 50216, 50217, 50218)
End Property

Friend Property Get oCollection() As clsCollectionBase
    Set oCollection = mp_oCollection
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub Add(ByVal Caption As String, ByVal RowKey As String, ByVal StartDate As Date, ByVal EndDate As Date, Optional ByVal Key As String = "", Optional ByVal StyleIndex As String = "0", Optional ByVal LayerIndex As String = "0")
    mp_oCollection.AddMode = True
    Dim oTask As New clsTask
    Key = mp_oControl.StrLib.StrTrim(Key)
    Caption = mp_oControl.StrLib.StrTrim(Caption)
    RowKey = mp_oControl.StrLib.StrTrim(RowKey)
    oTask.Initialize mp_oControl
    oTask.Caption = Caption
    oTask.RowKey = RowKey
    oTask.StartDate = StartDate
    oTask.EndDate = EndDate
    oTask.Key = Key
    oTask.StyleIndex = StyleIndex
    oTask.LayerIndex = LayerIndex
    mp_oCollection.m_Add oTask, Key, 50213, 50214
    Set oTask = Nothing
End Sub

Public Sub Clear()
    Dim lIndex As Long
    Dim oMilestone As clsMilestone
    mp_oControl.Percentages.Clear
    For lIndex = 1 To mp_oControl.Milestones.Count
        Set oMilestone = mp_oControl.Milestones.oCollection.m_oReturnArrayElement(lIndex)
        oMilestone.Predecessors.oCollection.m_CollectionRemoveWhereNot "TaskKey", ""
    Next lIndex
    mp_oControl.SelectedTaskIndex = 0
    mp_oCollection.m_Clear
End Sub

Public Sub Remove(ByVal Index As String)
    Dim lIndex As Long
    Dim sRIndex As String
    Dim sRKey As String
    Dim oTask As clsTask
    Dim oMilestone As clsMilestone
    mp_oCollection.m_GetKeyAndIndex Index, sRKey, sRIndex
    mp_oControl.Percentages.oCollection.m_CollectionRemoveWhere "TaskKey", sRKey, sRIndex
    For lIndex = 1 To mp_oControl.Tasks.Count
        Set oTask = mp_oControl.Tasks.oCollection.m_oReturnArrayElement(lIndex)
        oTask.Predecessors.oCollection.m_CollectionRemoveWhere "TaskKey", sRKey, sRIndex
    Next lIndex
    For lIndex = 1 To mp_oControl.Milestones.Count
        Set oMilestone = mp_oControl.Milestones.oCollection.m_oReturnArrayElement(lIndex)
        oMilestone.Predecessors.oCollection.m_CollectionRemoveWhere "TaskKey", sRKey, sRIndex
    Next lIndex
    mp_oControl.SelectedTaskIndex = 0
    mp_oCollection.m_Remove Index, 50219, 50220, 50221, 50222
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Custom Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub Sort(ByVal PropertyName As String, ByVal Descending As Boolean, ByVal SortType As E_SORTTYPE, Optional ByVal StartIndex As Long = -1, Optional ByVal EndIndex As Long = -1)
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

'// ---------------------------------------------------------------------------------------------------------------------
'// Position
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Position()
    Dim oTask As clsTask
    Dim lIndex As Long
    Dim lVisRowIndex As Long
    Dim oRow As clsRow
    Dim lTop As Long
    Dim lBottom As Long
    For lIndex = 1 To Count
        Set oTask = mp_oCollection.m_oReturnArrayElement(lIndex)
        lVisRowIndex = mp_oControl.f_MouseEvents.mp_lReturnVisibleRowIndexByKey(oTask.RowKey) '// Will return 0 if row is not visible
        If mp_oControl.Layers.FItem(oTask.LayerIndex).Visible = True Then
            Select Case oTask.TaskAttributes
                Case E_TASKATTRIBUTES.TA_INDEFINITESTART
                    oTask.StartDate = mp_oControl.CurrentViewObject.TimeLine.StartDate
                Case E_TASKATTRIBUTES.TA_INDEFINITEEND
                    oTask.EndDate = mp_oControl.CurrentViewObject.TimeLine.EndDate
                Case E_TASKATTRIBUTES.TA_INDEFINITESTARTANDEND
                    oTask.StartDate = mp_oControl.CurrentViewObject.TimeLine.StartDate
                    oTask.EndDate = mp_oControl.CurrentViewObject.TimeLine.EndDate
            End Select
            If (lVisRowIndex <> 0) And (((oTask.StartDate >= mp_oControl.CurrentViewObject.TimeLine.StartDate And oTask.StartDate <= mp_oControl.CurrentViewObject.TimeLine.EndDate) Or (oTask.EndDate >= mp_oControl.CurrentViewObject.TimeLine.StartDate And oTask.EndDate <= mp_oControl.CurrentViewObject.TimeLine.EndDate)) Or (oTask.StartDate < mp_oControl.CurrentViewObject.TimeLine.StartDate And oTask.EndDate > mp_oControl.CurrentViewObject.TimeLine.EndDate)) Then
                oTask.f_bVisible = True
            Else
                oTask.f_bVisible = False
            End If
            oTask.f_lLeft = mp_oControl.MathLib.GetXCoordinateFromDate(oTask.StartDate)
            If oTask.Left < -100 Then
                oTask.f_lLeft = -100
            End If
            oTask.f_lRight = mp_oControl.MathLib.GetXCoordinateFromDate(oTask.EndDate)
            If oTask.Right > (mp_oControl.mt_RightMargin + 100) Then
                oTask.f_lRight = (mp_oControl.mt_RightMargin + 100)
            End If
            Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElementKey(oTask.RowKey)
            lTop = oRow.Top
            lBottom = oRow.Bottom - 1
            If (oRow.Height <= -1) Then
                oTask.f_bVisible = False
                oTask.f_lTop = lTop
                oTask.f_lBottom = lTop
            Else
                If mp_oControl.mp_bPositionItem(lTop, lBottom, mp_oControl.Styles.FItem(oTask.StyleIndex)) = True Then
                    oTask.f_lTop = lTop
                    oTask.f_lBottom = lBottom
                Else
                    oTask.f_bVisible = False
                End If
                If (oTask.Right - oTask.Left) < 0 Then
                    oTask.f_bVisible = False
                End If
            End If
        Else
            oTask.f_bVisible = False
        End If
    Next lIndex
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Draw
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Draw()
    Dim lIndex As Long
    Dim oTask As clsTask
    Dim bDraw As Boolean
    If Count = 0 Then
        Exit Sub
    End If
    For lIndex = 1 To Count
        Set oTask = mp_oCollection.m_oReturnArrayElement(lIndex)
        If oTask.Visible = True Then
            If mp_oControl.GrphLib.RectangleIntersection(oTask.LeftTrim, oTask.Top, oTask.RightTrim, oTask.Bottom) Then
                mp_oControl.GrphLib.ClipRegion oTask.LeftTrim, oTask.Top, oTask.RightTrim, oTask.Bottom, True
                bDraw = False
                mp_oControl.FireDraw E_EVENTTARGET.EVT_TASK, bDraw, lIndex, 0, mp_oControl.GrphLib.mp_lHdc
                If bDraw = False Then
                    mp_oControl.mp_DrawItem oTask.Left, oTask.Right, oTask.Top, oTask.Bottom, oTask.StyleIndex, oTask.Caption, (lIndex = mp_oControl.SelectedTaskIndex), oTask.Picture, oTask.LeftTrim, oTask.RightTrim, Nothing
                End If
                oTask.Predecessors.Draw
            End If
        End If
    Next lIndex
    
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim lIndex As Long
    Dim oTask As clsTask
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Tasks"
    oXML.InitializeWriter
    For lIndex = 1 To Count
        Set oTask = mp_oCollection.m_oReturnArrayElement(lIndex)
        oXML.WriteObject oTask.GetXML
    Next lIndex
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim lIndex As Long
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Tasks"
    oXML.SetXML sXML
    oXML.InitializeReader
    mp_oCollection.m_Clear
    For lIndex = 1 To oXML.ReadCollectionCount
        Dim oTask As New clsTask
        oTask.Initialize mp_oControl
        oTask.SetXML oXML.ReadCollectionObject(lIndex)
        mp_oCollection.AddMode = True
        mp_oCollection.m_Add oTask, oTask.Key, 50000, 50000
        Set oTask = Nothing
    Next lIndex
End Sub
