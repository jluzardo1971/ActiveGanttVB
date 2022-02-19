VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrid"
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
Private mp_bHorizontalLines As Boolean
Private mp_bVerticalLines As Boolean
Private mp_bSnapToGrid As Boolean
Private mp_bSnapToGridOnSelection As Boolean
Private mp_clrColor As OLE_COLOR
Private mp_sInterval As String
Private mp_lFactor As Long
Private mp_oTimeLine As clsTimeLine

'// ---------------------------------------------------------------------------------------------------------------------
'// Initialization & Destruction
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl, ByRef oTimeLine As clsTimeLine)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oTimeLine = oTimeLine
    mp_bHorizontalLines = True
    mp_bVerticalLines = False
    mp_bSnapToGrid = False
    mp_bSnapToGridOnSelection = True
    mp_clrColor = GRE_COLORS.CLR_BUTTONFACE
    mp_sInterval = "n"
    mp_lFactor = 15
End Sub

Private Sub Class_Terminate()
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Object Member Variables
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get HorizontalLines() As Boolean
    HorizontalLines = mp_bHorizontalLines
End Property

Public Property Let HorizontalLines(ByVal Value As Boolean)
    mp_bHorizontalLines = Value
End Property

Public Property Get VerticalLines() As Boolean
    VerticalLines = mp_bVerticalLines
End Property

Public Property Let VerticalLines(ByVal Value As Boolean)
    mp_bVerticalLines = Value
End Property

Public Property Get SnapToGrid() As Boolean
    SnapToGrid = mp_bSnapToGrid
End Property

Public Property Let SnapToGrid(ByVal Value As Boolean)
    mp_bSnapToGrid = Value
End Property

Public Property Get SnapToGridOnSelection() As Boolean
    SnapToGridOnSelection = mp_bSnapToGridOnSelection
End Property

Public Property Let SnapToGridOnSelection(ByVal Value As Boolean)
    mp_bSnapToGridOnSelection = Value
End Property

Public Property Get Color() As OLE_COLOR
    Color = mp_clrColor
End Property

Public Property Let Color(ByVal Value As OLE_COLOR)
    mp_clrColor = mp_oControl.GrphLib.ConvertColor(Value)
End Property

Public Property Get Interval() As String
    Interval = mp_lFactor & mp_sInterval
End Property

Public Property Let Interval(ByVal Value As String)
    mp_oControl.mp_ProcessInterval Value, mp_sInterval, mp_lFactor
End Property

Friend Property Get f_Interval() As String
    f_Interval = mp_sInterval
End Property

Friend Property Get f_Factor() As Long
    f_Factor = mp_lFactor
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Draw
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Draw()
    Dim dtBuff As Date
    If mp_bVerticalLines = False Then
        Exit Sub
    End If
    If mp_oControl.MathLib.GetXCoordinateFromDate(mp_oControl.MathLib.DateTimeAdd(mp_sInterval, mp_lFactor, mp_oTimeLine.StartDate)) - mp_oControl.MathLib.GetXCoordinateFromDate(mp_oTimeLine.StartDate) < 5 Then
        Exit Sub
    End If
    mp_oControl.GrphLib.ClipRegion mp_oTimeLine.f_lStart, mp_oControl.CurrentViewObject.ClientArea.Top, mp_oTimeLine.f_lEnd, mp_oControl.CurrentViewObject.ClientArea.Bottom, True
    dtBuff = mp_oControl.MathLib.RoundDate(mp_sInterval, mp_lFactor, mp_oTimeLine.StartDate)
    If mp_oControl.MathLib.GetXCoordinateFromDate(dtBuff) >= mp_oTimeLine.f_lStart Then
        mp_PaintVerticalGridLine mp_oControl.MathLib.GetXCoordinateFromDate(dtBuff), GRE_LINEDRAWSTYLE.LDS_SOLID
    End If
    Do While dtBuff < mp_oTimeLine.EndDate
        dtBuff = mp_oControl.MathLib.DateTimeAdd(mp_sInterval, mp_lFactor, dtBuff)
        mp_PaintVerticalGridLine mp_oControl.MathLib.GetXCoordinateFromDate(dtBuff), GRE_LINEDRAWSTYLE.LDS_SOLID
    Loop
End Sub

Private Sub mp_PaintVerticalGridLine(ByVal fXCoordinate As Long, ByVal v_lDrawStyle As GRE_LINEDRAWSTYLE)
    mp_oControl.GrphLib.DrawLine fXCoordinate, mp_oControl.CurrentViewObject.ClientArea.Top, fXCoordinate, mp_oControl.Rows.TopOffset, GRE_LINETYPE.LT_NORMAL, mp_clrColor, v_lDrawStyle
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Grid"
    oXML.InitializeWriter
    oXML.WriteProperty "HorizontalLines", mp_bHorizontalLines
    oXML.WriteProperty "VerticalLines", mp_bVerticalLines
    oXML.WriteProperty "SnapToGrid", mp_bSnapToGrid
    oXML.WriteProperty "SnapToGridOnSelection", mp_bSnapToGridOnSelection
    oXML.WriteProperty "Color", mp_clrColor
    oXML.WriteProperty "Interval", mp_sInterval
    oXML.WriteProperty "Factor", mp_lFactor
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Grid"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadProperty "HorizontalLines", mp_bHorizontalLines
    oXML.ReadProperty "VerticalLines", mp_bVerticalLines
    oXML.ReadProperty "SnapToGrid", mp_bSnapToGrid
    oXML.ReadProperty "SnapToGridOnSelection", mp_bSnapToGridOnSelection
    oXML.ReadPropertyColor "Color", mp_clrColor
    oXML.ReadProperty "Interval", mp_sInterval
    oXML.ReadProperty "Factor", mp_lFactor
End Sub


