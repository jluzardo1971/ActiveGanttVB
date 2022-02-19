VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMilestones"
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
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oCollection = New clsCollectionBase
    mp_oCollection.Initialize Value, "Milestone"
    
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

Public Property Get Item(ByVal Index As String) As clsMilestone
    Set Item = mp_oCollection.m_oItem(Index, 50345, 50346, 50347, 50348)
End Property

Friend Property Get oCollection() As clsCollectionBase
    Set oCollection = mp_oCollection
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub Add(ByVal Caption As String, ByVal RowKey As String, ByVal MilestoneDate As Date, Optional ByVal Key As String = "", Optional ByVal StyleIndex As String = "0", Optional ByVal LayerIndex As String = "0")
    mp_oCollection.AddMode = True
    Dim oMilestone As New clsMilestone
    oMilestone.Initialize mp_oControl
    oMilestone.Caption = Caption
    oMilestone.RowKey = RowKey
    oMilestone.MilestoneDate = MilestoneDate
    oMilestone.Key = Key
    oMilestone.StyleIndex = StyleIndex
    oMilestone.LayerIndex = LayerIndex
    mp_oCollection.m_Add oMilestone, Key, 50335, 50336
    Set oMilestone = Nothing
End Sub

Public Sub Clear()
    Dim lIndex As Long
    Dim oTask As clsTask
    For lIndex = 1 To mp_oControl.Tasks.Count
        Set oTask = mp_oControl.Tasks.oCollection.m_oReturnArrayElement(lIndex)
        oTask.Predecessors.oCollection.m_CollectionRemoveWhereNot "MilestoneKey", ""
    Next lIndex
    mp_oControl.SelectedMilestoneIndex = 0
    mp_oCollection.m_Clear
End Sub

Public Sub Remove(ByVal Index As String)
    Dim lIndex As Long
    Dim oTask As clsTask
    Dim oMilestone As clsMilestone
    Dim sRIndex As String
    Dim sRKey As String
    mp_oCollection.m_GetKeyAndIndex Index, sRKey, sRIndex
    For lIndex = 1 To mp_oControl.Tasks.Count
        Set oTask = mp_oControl.Tasks.oCollection.m_oReturnArrayElement(lIndex)
        oTask.Predecessors.oCollection.m_CollectionRemoveWhere "MilestoneKey", sRKey, sRIndex
    Next lIndex
    For lIndex = 1 To mp_oControl.Milestones.Count
        Set oMilestone = mp_oControl.Milestones.oCollection.m_oReturnArrayElement(lIndex)
        oMilestone.Predecessors.oCollection.m_CollectionRemoveWhere "MilestoneKey", sRKey, sRIndex
    Next lIndex
    mp_oControl.SelectedMilestoneIndex = 0
    mp_oCollection.m_Remove Index, 50365, 50366, 50367, 50368
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
    Dim oMilestone As clsMilestone
    Dim lIndex As Long
    Dim lVisRowIndex As Long
    Dim oRow As clsRow
    Dim lTop As Long
    Dim lBottom As Long
    For lIndex = 1 To Count
        Set oMilestone = mp_oCollection.m_oReturnArrayElement(lIndex)
        lVisRowIndex = mp_oControl.f_MouseEvents.mp_lReturnVisibleRowIndexByKey(oMilestone.RowKey) '// Will return 0 if row is not visible
        If mp_oControl.Layers.FItem(oMilestone.LayerIndex).Visible = True Then
            If (lVisRowIndex <> 0) And (oMilestone.MilestoneDate >= mp_oControl.CurrentViewObject.TimeLine.StartDate And oMilestone.MilestoneDate <= mp_oControl.CurrentViewObject.TimeLine.EndDate) Then
                oMilestone.f_bVisible = True
            Else
                oMilestone.f_bVisible = False
            End If
            oMilestone.f_lLeft = mp_oControl.MathLib.GetXCoordinateFromDate(oMilestone.MilestoneDate) - mp_oControl.CurrentViewObject.ClientArea.MilestoneSelectionOffset
            If oMilestone.Left < -100 Then
                oMilestone.f_lLeft = -100
            End If
            oMilestone.f_lRight = mp_oControl.MathLib.GetXCoordinateFromDate(oMilestone.MilestoneDate) + mp_oControl.CurrentViewObject.ClientArea.MilestoneSelectionOffset
            If oMilestone.Right > (mp_oControl.mt_RightMargin + 100) Then
                oMilestone.f_lRight = (mp_oControl.mt_RightMargin + 100)
            End If
            Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElementKey(oMilestone.RowKey)
            lTop = oRow.Top
            lBottom = oRow.Bottom - 1
            If (oRow.Height <= -1) Then
                oMilestone.f_bVisible = False
                oMilestone.f_lTop = lTop
                oMilestone.f_lBottom = lTop
            Else
                If mp_oControl.mp_bPositionItem(lTop, lBottom, mp_oControl.Styles.FItem(oMilestone.StyleIndex)) = True Then
                    oMilestone.f_lTop = lTop
                    oMilestone.f_lBottom = lBottom
                Else
                    oMilestone.f_bVisible = False
                End If
                If (oMilestone.Right - oMilestone.Left) < 0 Then
                    oMilestone.f_bVisible = False
                End If
            End If
        Else
            oMilestone.f_bVisible = False
        End If
    Next lIndex
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Draw
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Draw()
    Dim lIndex As Long
    Dim oMilestone As clsMilestone
    Dim bDraw As Boolean
    If Count = 0 Then
        Exit Sub
    End If
    For lIndex = 1 To Count
        Set oMilestone = mp_oCollection.m_oReturnArrayElement(lIndex)
        If oMilestone.Visible = True Then
            If mp_oControl.GrphLib.RectangleIntersection(oMilestone.LeftTrim, oMilestone.Top, oMilestone.RightTrim, oMilestone.Bottom) Then
                mp_oControl.GrphLib.ClipRegion oMilestone.LeftTrim, oMilestone.Top, oMilestone.RightTrim, oMilestone.Bottom, True
                bDraw = False
                mp_oControl.FireDraw E_EVENTTARGET.EVT_MILESTONE, bDraw, lIndex, 0, mp_oControl.GrphLib.mp_lHdc
                If bDraw = False Then
                    mp_oControl.mp_DrawItemI oMilestone, oMilestone.StyleIndex
                End If
                oMilestone.Predecessors.Draw
            End If
        End If
    Next lIndex
    
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim lIndex As Long
    Dim oMilestone As clsMilestone
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Milestones"
    oXML.InitializeWriter
    For lIndex = 1 To Count
        Set oMilestone = mp_oCollection.m_oReturnArrayElement(lIndex)
        oXML.WriteObject oMilestone.GetXML
    Next lIndex
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim lIndex As Long
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Milestones"
    oXML.SetXML sXML
    oXML.InitializeReader
    mp_oCollection.m_Clear
    For lIndex = 1 To oXML.ReadCollectionCount
        Dim oMilestone As New clsMilestone
        oMilestone.Initialize mp_oControl
        oMilestone.SetXML oXML.ReadCollectionObject(lIndex)
        mp_oCollection.AddMode = True
        mp_oCollection.m_Add oMilestone, oMilestone.Key, 50000, 50000
        Set oMilestone = Nothing
    Next lIndex
End Sub
