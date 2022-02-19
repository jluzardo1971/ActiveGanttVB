VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClientArea"
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
Private mp_oTimeLine As clsTimeLine
Private mp_bDetectConflicts As Boolean
Private mp_lMilestoneSelectionOffset As Long
Private mp_lLastVisibleRow As Long
Private mp_oGrid As clsGrid
Private mp_sToolTipFormat As String
Private mp_bToolTipsVisible As Boolean


'// ---------------------------------------------------------------------------------------------------------------------
'// Initialization & Destruction
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl, ByRef oTimeLine As clsTimeLine)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oTimeLine = oTimeLine
    mp_bDetectConflicts = True
    mp_lMilestoneSelectionOffset = 5
    mp_lLastVisibleRow = 0
    Set mp_oGrid = New clsGrid
    mp_oGrid.Initialize mp_oControl, mp_oTimeLine
    mp_sToolTipFormat = "ddddd"
    mp_bToolTipsVisible = True
End Sub

Private Sub Class_Terminate()
    Set mp_oTimeLine = Nothing
    Set mp_oGrid = Nothing
    Set mp_oControl = Nothing
    
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Object Member Variables
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get DetectConflicts() As Boolean
    DetectConflicts = mp_bDetectConflicts
End Property

Public Property Let DetectConflicts(ByVal Value As Boolean)
    mp_bDetectConflicts = Value
End Property

Public Property Get MilestoneSelectionOffset() As Long
    MilestoneSelectionOffset = mp_lMilestoneSelectionOffset
End Property

Public Property Let MilestoneSelectionOffset(ByVal Value As Long)
    mp_lMilestoneSelectionOffset = Value
End Property

Public Property Get FirstVisibleRow() As Long
    If mp_oControl.Rows.Count = 0 Then
        FirstVisibleRow = 0
        Exit Sub
    End If
    If (mp_oControl.f_oVScrollBar.Value > mp_oControl.Rows.Count) Then
        mp_oControl.f_oVScrollBar.Value = mp_oControl.Rows.Count
    End If
    If (mp_oControl.f_oVScrollBar.Max <> mp_oControl.Rows.Count) Then
        mp_oControl.f_oVScrollBar.Max = mp_oControl.Rows.Count
    End If
    FirstVisibleRow = mp_oControl.f_oVScrollBar.Value
End Property

Public Property Let FirstVisibleRow(ByVal Value As Long)
    If mp_oControl.Rows.Count = 0 Then
        Exit Property
    End If
    If Value < 1 Then
        Value = 1
    End If
    If Value > mp_oControl.Rows.Count Then
        Value = mp_oControl.Rows.Count
    End If
    mp_oControl.f_oVScrollBar.Value = Value
End Property

Public Property Get LastVisibleRow() As Long
    LastVisibleRow = mp_lLastVisibleRow
End Property

Friend Property Let f_LastVisibleRow(ByVal Value As Long)
    mp_lLastVisibleRow = Value
End Property

Public Property Get Grid() As clsGrid
    Set Grid = mp_oGrid
End Property

Public Property Get ToolTipFormat() As String
    ToolTipFormat = mp_sToolTipFormat
End Property

Public Property Let ToolTipFormat(ByVal Value As String)
    mp_sToolTipFormat = Value
End Property

Public Property Get ToolTipsVisible() As Boolean
    ToolTipsVisible = mp_bToolTipsVisible
End Property

Public Property Let ToolTipsVisible(ByVal Value As Boolean)
    mp_bToolTipsVisible = Value
End Property

Public Property Get Top() As Long
    If (mp_oTimeLine.Height = 0) Then
        Top = mp_oControl.mt_BorderThickness
    Else
        Top = mp_oTimeLine.Bottom + 1
    End If
End Property

Public Property Get Bottom() As Long
    If mp_oControl.f_oHScrollBar2.State = 3 Then
        Bottom = mp_oControl.GrphLib.Height - mp_oControl.mt_BorderThickness - 1 - mp_oControl.f_oHScrollBar2.Height
    Else
        Bottom = mp_oControl.GrphLib.Height - mp_oControl.mt_BorderThickness - 1
    End If
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Draw
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Draw()
    Dim lRowIndex As Long
    Dim oRow As clsRow
    If mp_oControl.Rows.Count = 0 Then
        Exit Sub
    End If
    mp_oControl.GrphLib.ClipRegion mp_oControl.Splitter.Right, mp_oControl.CurrentViewObject.ClientArea.Top, mp_oControl.mt_RightMargin, mp_oControl.CurrentViewObject.ClientArea.Bottom, True
    For lRowIndex = FirstVisibleRow To mp_lLastVisibleRow
        Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(lRowIndex)
        If oRow.ClientAreaStyleIndex <> "" Then
            mp_oControl.mp_DrawItem mp_oControl.Splitter.Right, mp_oControl.mt_RightMargin, oRow.Top, oRow.Bottom, oRow.ClientAreaStyleIndex, "", False, Nothing, 0, 0, Nothing
        End If
        If mp_oGrid.HorizontalLines = True Then
            mp_oControl.GrphLib.DrawLine mp_oControl.Splitter.Right, oRow.Bottom, mp_oControl.mt_RightMargin, oRow.Bottom, GRE_LINETYPE.LT_NORMAL, mp_oGrid.Color, GRE_LINEDRAWSTYLE.LDS_SOLID
        End If
    Next lRowIndex
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "ClientArea"
    oXML.InitializeWriter
    oXML.WriteProperty "DetectConflicts", mp_bDetectConflicts
    oXML.WriteProperty "MilestoneSelectionOffset", mp_lMilestoneSelectionOffset
    oXML.WriteProperty "FirstVisibleRow", FirstVisibleRow
    oXML.WriteProperty "ToolTipFormat", mp_sToolTipFormat
    oXML.WriteProperty "ToolTipsVisible", mp_bToolTipsVisible
    oXML.WriteObject mp_oGrid.GetXML
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "ClientArea"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadProperty "DetectConflicts", mp_bDetectConflicts
    oXML.ReadProperty "MilestoneSelectionOffset", mp_lMilestoneSelectionOffset
    Dim lFirstVisibleRow As Long
    oXML.ReadProperty "FirstVisibleRow", lFirstVisibleRow
    FirstVisibleRow = lFirstVisibleRow
    oXML.ReadProperty "ToolTipFormat", mp_sToolTipFormat
    oXML.ReadProperty "ToolTipsVisible", mp_bToolTipsVisible
    mp_oGrid.SetXML oXML.ReadObject("Grid")
End Sub



