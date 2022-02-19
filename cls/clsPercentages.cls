VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPercentages"
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
    mp_oCollection.Initialize Value, "Percentage"
    
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

Public Property Get Item(ByVal Index As String) As clsPercentage
    Set Item = mp_oCollection.m_oItem(Index, 50353, 50354, 50355, 50356)
End Property

Friend Property Get oCollection() As clsCollectionBase
    Set oCollection = mp_oCollection
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub Add(ByVal TaskKey As String, ByVal PercentageGroupIndex As String, ByVal Percent As Single, Optional ByVal Key As String = "")
    mp_oCollection.AddMode = True
    Dim oPercentage As New clsPercentage
    oPercentage.Initialize mp_oControl
    oPercentage.TaskKey = TaskKey
    oPercentage.Percent = Percent
    oPercentage.PercentageGroupIndex = PercentageGroupIndex
    mp_oCollection.m_Add oPercentage, Key, 50339, 50340
    Set oPercentage = Nothing
End Sub

Public Sub Clear()
    mp_oCollection.m_Clear
End Sub

Public Sub Remove(ByVal Index As String)
    mp_oCollection.m_Remove Index, 50373, 50374, 50375, 50376
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Position
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Position()
    Dim oPercentage As clsPercentage
    Dim oTask As clsTask
    Dim oPercentageGroup As clsPercentageGroup
    Dim lIndex As Long
    Dim oRow As clsRow
    Dim oStyle As clsStyle
    Dim lTop As Long
    Dim lBottom As Long
    For lIndex = 1 To Count
        Set oPercentage = mp_oCollection.m_oReturnArrayElement(lIndex)
        Set oTask = mp_oControl.Tasks.oCollection.m_oReturnArrayElementKey(oPercentage.TaskKey)
        If oTask.Visible = True Then
            Set oPercentageGroup = mp_oControl.PercentageGroups.oCollection.m_oReturnArrayElementKey(oPercentage.PercentageGroupIndex)
            If oPercentageGroup.Visible = True Then
                oPercentage.f_bVisible = True
                Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElementKey(oTask.RowKey)
                Set oStyle = mp_oControl.Styles.FItem(oPercentageGroup.StyleIndex)
                lTop = oRow.Top
                lBottom = oRow.Bottom
                oPercentage.f_lLeft = mp_oControl.MathLib.GetXCoordinateFromDate(oTask.StartDate)
                If oPercentage.Left < -100 Then
                    oPercentage.f_lLeft = -100
                End If
                oPercentage.f_lRight = mp_oControl.MathLib.GetXCoordinateFromDate(oTask.EndDate)
                oPercentage.f_lRight = oPercentage.Left + ((oPercentage.Right - oPercentage.Left) * oPercentage.Percent)
                If oPercentage.Right > (mp_oControl.mt_RightMargin + 100) Then
                    oPercentage.f_lRight = (mp_oControl.mt_RightMargin + 100)
                End If
                If mp_oControl.mp_bPositionItem(lTop, lBottom, mp_oControl.Styles.FItem(oPercentageGroup.StyleIndex)) = False Then
                    oPercentage.f_bVisible = False
                End If
                oPercentage.f_lTop = lTop
                oPercentage.f_lBottom = lBottom
                If (oPercentage.Right - oPercentage.Left) < 0 Then
                    oPercentage.f_bVisible = False
                End If
            Else
                oPercentage.f_bVisible = False
            End If
        Else
            oPercentage.f_bVisible = False
        End If
    Next
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Draw
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Draw()
    Dim lIndex As Long
    Dim oPercentage As clsPercentage
    Dim oTask As clsTask
    Dim oPercentageGroup As clsPercentageGroup
    Dim oRow As clsRow
    Dim oStyle As clsStyle
    Dim lTop As Long
    Dim lBottom As Long
    Dim bDraw As Boolean
    If Count = 0 Then
        Exit Sub
    End If
    If Count = 0 Then
        Exit Sub
    End If
    For lIndex = 1 To Count
        Set oPercentage = mp_oCollection.m_oReturnArrayElement(lIndex)
        If oPercentage.Visible = True Then
            If mp_oControl.GrphLib.RectangleIntersection(oPercentage.LeftTrim, oPercentage.Top, oPercentage.RightTrim, oPercentage.Bottom) Then
                mp_oControl.GrphLib.ClipRegion oPercentage.LeftTrim, oPercentage.Top, oPercentage.RightTrim, oPercentage.Bottom, True
                Set oTask = mp_oControl.Tasks.oCollection.m_oReturnArrayElementKey(oPercentage.TaskKey)
                Set oPercentageGroup = mp_oControl.PercentageGroups.oCollection.m_oReturnArrayElementKey(oPercentage.PercentageGroupIndex)
                Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElementKey(oTask.RowKey)
                Set oStyle = mp_oControl.Styles.FItem(oPercentageGroup.StyleIndex)
                bDraw = False
                mp_oControl.FireDraw E_EVENTTARGET.EVT_PERCENTAGE, bDraw, lIndex, 0, mp_oControl.GrphLib.mp_lHdc
                If bDraw = False Then
                    mp_oControl.mp_DrawItem oPercentage.Left, oPercentage.Right, oPercentage.Top, oPercentage.Bottom, oPercentageGroup.StyleIndex, mp_oControl.StrLib.StrFormat(oPercentage.Percent, oPercentageGroup.Format), False, Nothing, oPercentage.LeftTrim, oPercentage.RightTrim, Nothing
                End If
            End If
        End If
    Next lIndex
End Sub


'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim lIndex As Long
    Dim oPercentage As clsPercentage
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Percentages"
    oXML.InitializeWriter
    For lIndex = 1 To Count
        Set oPercentage = mp_oCollection.m_oReturnArrayElement(lIndex)
        oXML.WriteObject oPercentage.GetXML
    Next lIndex
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim lIndex As Long
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Percentages"
    oXML.SetXML sXML
    oXML.InitializeReader
    mp_oCollection.m_Clear
    For lIndex = 1 To oXML.ReadCollectionCount
        Dim oPercentage As New clsPercentage
        oPercentage.Initialize mp_oControl
        oPercentage.SetXML oXML.ReadCollectionObject(lIndex)
        mp_oCollection.AddMode = True
        mp_oCollection.m_Add oPercentage, oPercentage.Key, 50000, 50000
        Set oPercentage = Nothing
    Next lIndex
End Sub
