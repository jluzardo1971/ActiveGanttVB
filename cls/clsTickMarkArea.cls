VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTickMarkArea"
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
Private mp_sInterval As String
Private mp_lFactor As Long
Private mp_oFont As StdFont
Private mp_lHeight As Long
Private mp_lBigTickMarkHeight As Long
Private mp_lMediumTickMarkHeight As Long
Private mp_lSmallTickMarkHeight As Long
Private mp_bVisible As Boolean
Private mp_lTextOffset As Long
Private mp_oTickMarks As clsTickMarks
Private mp_oTimeLine As clsTimeLine




'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl, ByRef oTimeLine As clsTimeLine, ByVal bInit As Boolean)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oTimeLine = oTimeLine
    If bInit = False Then
        Set mp_oFont = New StdFont
        mp_lHeight = 23
        mp_sInterval = "n"
        mp_lFactor = 5
        mp_lBigTickMarkHeight = 12
        mp_lMediumTickMarkHeight = 9
        mp_lSmallTickMarkHeight = 7
        mp_bVisible = True
        mp_lTextOffset = 13
    Else
        Set mp_oFont = mp_oControl.DefaultValues.TickMarkArea.Font
        mp_lHeight = mp_oControl.DefaultValues.TickMarkArea.Height
        mp_sInterval = mp_oControl.DefaultValues.TickMarkArea.f_Interval
        mp_lFactor = mp_oControl.DefaultValues.TickMarkArea.f_Factor
        mp_lBigTickMarkHeight = mp_oControl.DefaultValues.TickMarkArea.BigTickMarkHeight
        mp_lMediumTickMarkHeight = mp_oControl.DefaultValues.TickMarkArea.MediumTickMarkHeight
        mp_lSmallTickMarkHeight = mp_oControl.DefaultValues.TickMarkArea.SmallTickMarkHeight
        mp_bVisible = mp_oControl.DefaultValues.TickMarkArea.Visible
        mp_lTextOffset = mp_oControl.DefaultValues.TickMarkArea.TextOffset
    End If
    Set mp_oTickMarks = New clsTickMarks
    mp_oTickMarks.Initialize mp_oControl
End Sub

Private Sub Class_Terminate()
    Set mp_oTickMarks = Nothing
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Object Member Variables
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get TickMarks() As clsTickMarks
    Set TickMarks = mp_oTickMarks
End Property

Public Property Get Interval() As String
    Interval = mp_lFactor & mp_sInterval
End Property

Public Property Let Interval(ByVal Value As String)
    mp_oControl.mp_ProcessInterval Value, mp_sInterval, mp_lFactor
End Property

Friend Property Get f_Factor() As Long
    f_Factor = mp_lFactor
End Property

Friend Property Get f_Interval() As String
    f_Interval = mp_sInterval
End Property

Public Property Get Font() As Font
    Set Font = mp_oFont
End Property

Public Property Set Font(ByVal Value As Font)
    Set mp_oFont = Value
End Property

Public Property Get Height() As Long
    Height = mp_lHeight
End Property

Public Property Let Height(ByVal Value As Long)
    mp_lHeight = Value
End Property

Public Property Get BigTickMarkHeight() As Long
    BigTickMarkHeight = mp_lBigTickMarkHeight
End Property

Public Property Let BigTickMarkHeight(ByVal Value As Long)
    mp_lBigTickMarkHeight = Value
End Property

Public Property Get MediumTickMarkHeight() As Long
    MediumTickMarkHeight = mp_lMediumTickMarkHeight
End Property

Public Property Let MediumTickMarkHeight(ByVal Value As Long)
    mp_lMediumTickMarkHeight = Value
End Property

Public Property Get SmallTickMarkHeight() As Long
    SmallTickMarkHeight = mp_lSmallTickMarkHeight
End Property

Public Property Let SmallTickMarkHeight(ByVal Value As Long)
    mp_lSmallTickMarkHeight = Value
End Property

Public Property Get Visible() As Boolean
    Visible = mp_bVisible
End Property

Public Property Let Visible(ByVal Value As Boolean)
    mp_bVisible = Value
End Property

Public Property Get TextOffset() As Long
    TextOffset = mp_lTextOffset
End Property

Public Property Let TextOffset(ByVal Value As Long)
    mp_lTextOffset = Value
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Draw
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Draw()
    Dim dtStart As Date
    Dim dtTimeLineBuffer As Date
    Dim lTimeLineBuffer As Long
    Dim fXCoordinate As Long
    Dim oTickMark As clsTickMark
    Dim lIndex As Long
    Dim lValue As Long
    dtStart = mp_oControl.MathLib.RoundDate(mp_sInterval, mp_lFactor, mp_oTimeLine.StartDate)
    lTimeLineBuffer = mp_oControl.MathLib.GetXCoordinateFromDate(dtStart)
    dtTimeLineBuffer = mp_oControl.MathLib.GetDateFromXCoordinate(lTimeLineBuffer - 100)
    dtTimeLineBuffer = mp_oControl.MathLib.RoundDate(mp_sInterval, mp_lFactor, dtTimeLineBuffer)
    Do While (fXCoordinate < mp_oControl.GrphLib.Width + 100)
        fXCoordinate = mp_oControl.MathLib.GetXCoordinateFromDate(dtTimeLineBuffer)
        For lIndex = 1 To mp_oTickMarks.Count
            Select Case mp_sInterval
                Case "s"
                    lValue = mp_oControl.MathLib.GetSecond(dtTimeLineBuffer)
                Case "n"
                    lValue = mp_oControl.MathLib.GetMinute(dtTimeLineBuffer)
                Case "h"
                    lValue = mp_oControl.MathLib.GetHour(dtTimeLineBuffer)
                Case "d"
                    lValue = mp_oControl.MathLib.GetDay(dtTimeLineBuffer)
                Case "w"
                    lValue = mp_oControl.MathLib.GetDay(dtTimeLineBuffer)
                Case "y"
                    lValue = mp_oControl.MathLib.GetDay(dtTimeLineBuffer)
                Case "ww"
                    lValue = mp_oControl.MathLib.GetWeek(dtTimeLineBuffer)
                Case "m"
                    lValue = mp_oControl.MathLib.GetMonth(dtTimeLineBuffer)
                Case "q"
                    lValue = mp_oControl.MathLib.GetQuarter(dtTimeLineBuffer)
                Case "yyyy"
                    lValue = mp_oControl.MathLib.GetYear(dtTimeLineBuffer)
            End Select
            Set oTickMark = mp_oTickMarks.oCollection.m_oReturnArrayElement(lIndex)
            If (oTickMark.Interval = lValue And oTickMark.Modulus = False) Then
                Select Case oTickMark.TickMarkType
                    Case E_TICKMARKTYPES.TLT_BIG
                        PaintTickMark fXCoordinate, mp_lBigTickMarkHeight
                    Case E_TICKMARKTYPES.TLT_MEDIUM
                        PaintTickMark fXCoordinate, mp_lMediumTickMarkHeight
                    Case E_TICKMARKTYPES.TLT_SMALL
                        PaintTickMark fXCoordinate, mp_lSmallTickMarkHeight
                End Select
                If oTickMark.DisplayCaption = True Then
                    PaintCaption fXCoordinate, dtTimeLineBuffer, oTickMark.CaptionFormat
                End If
            ElseIf (oTickMark.Modulus = True) Then
                If (oTickMark.Interval > 0) Then
                    If (lValue Mod oTickMark.Interval = 0) Then
                        Select Case oTickMark.TickMarkType
                            Case E_TICKMARKTYPES.TLT_BIG
                                PaintTickMark fXCoordinate, mp_lBigTickMarkHeight
                            Case E_TICKMARKTYPES.TLT_MEDIUM
                                PaintTickMark fXCoordinate, mp_lMediumTickMarkHeight
                            Case E_TICKMARKTYPES.TLT_SMALL
                                PaintTickMark fXCoordinate, mp_lSmallTickMarkHeight
                        End Select
                        If oTickMark.DisplayCaption = True Then
                            PaintCaption fXCoordinate, dtTimeLineBuffer, oTickMark.CaptionFormat
                        End If
                    End If
                End If
            End If
        Next lIndex
        dtTimeLineBuffer = mp_oControl.MathLib.DateTimeAdd(mp_sInterval, mp_lFactor, dtTimeLineBuffer)
    Loop
End Sub

Private Sub PaintTickMark(ByVal fXCoordinate As Long, ByVal v_lTickMarkHeight As Long)
    mp_oControl.GrphLib.DrawLine fXCoordinate, mp_oTimeLine.Bottom - v_lTickMarkHeight, fXCoordinate, mp_oTimeLine.Bottom, GRE_LINETYPE.LT_NORMAL, mp_oTimeLine.ForeColor, GRE_LINEDRAWSTYLE.LDS_SOLID
End Sub

Private Sub PaintCaption(ByVal fXCoordinate As Single, ByVal dtTimeLineBuffer As Date, ByVal sFormat As String)
    Dim sDateBuff As String
    Dim lLeft As Long
    Dim lTop As Long
    Dim lRight As Long
    Dim lBottom As Long
    Dim lStringWidth As Long
    Dim lStringHeight As Long
    sDateBuff = mp_oControl.StrLib.StrFormat(dtTimeLineBuffer, sFormat)
    lStringWidth = mp_oControl.mp_lStrWidth(sDateBuff, mp_oFont)
    lStringHeight = mp_oControl.mp_lStrHeight(sDateBuff, mp_oFont)
    lLeft = fXCoordinate - (lStringWidth / 2) - 10
    lTop = mp_oTimeLine.Bottom - mp_lTextOffset - lStringHeight
    lRight = fXCoordinate + (lStringWidth / 2) + 10
    lBottom = lTop + lStringHeight
    mp_oControl.GrphLib.TextOutEx lLeft, lTop, lRight, lBottom, sDateBuff, GRE_HORIZONTALALIGNMENT.HAL_CENTER, GRE_VERTICALALIGNMENT.VAL_CENTER, mp_oTimeLine.ForeColor, mp_oFont
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TickMarkArea"
    oXML.InitializeWriter
    oXML.WriteProperty "Interval", mp_sInterval
    oXML.WriteProperty "Factor", mp_lFactor
    oXML.WritePropertyFont "Font", mp_oFont
    oXML.WriteProperty "Height", mp_lHeight
    oXML.WriteProperty "BigTickMarkHeight", mp_lBigTickMarkHeight
    oXML.WriteProperty "MediumTickMarkHeight", mp_lMediumTickMarkHeight
    oXML.WriteProperty "SmallTickMarkHeight", mp_lSmallTickMarkHeight
    oXML.WriteProperty "Visible", mp_bVisible
    oXML.WriteProperty "TextOffset", mp_lTextOffset
    oXML.WriteObject mp_oTickMarks.GetXML
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TickMarkArea"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadProperty "Interval", mp_sInterval
    oXML.ReadProperty "Factor", mp_lFactor
    oXML.ReadPropertyFont "Font", mp_oFont
    oXML.ReadProperty "Height", mp_lHeight
    oXML.ReadProperty "BigTickMarkHeight", mp_lBigTickMarkHeight
    oXML.ReadProperty "MediumTickMarkHeight", mp_lMediumTickMarkHeight
    oXML.ReadProperty "SmallTickMarkHeight", mp_lSmallTickMarkHeight
    oXML.ReadProperty "Visible", mp_bVisible
    oXML.ReadProperty "TextOffset", mp_lTextOffset
    mp_oTickMarks.SetXML oXML.ReadObject("TickMarks")
End Sub
