VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTier"
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
Private mp_bVisible As Boolean
Private mp_lFactor As Long
Private mp_lHeight As Long
Private mp_sInterval As String
Private mp_sTag As String
Private mp_yTierType As E_TIERTYPE
Private mp_oFont As StdFont
Private mp_yTierPosition As E_TIERPOSITION
Private mp_sTierPosition As String
Private mp_oTierArea As clsTierArea

'// ---------------------------------------------------------------------------------------------------------------------
'// Initialization & Destruction
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl, ByRef oTierArea As clsTierArea, ByVal yTierPosition As E_TIERPOSITION, ByVal bInit As Boolean)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oTierArea = oTierArea
    mp_yTierPosition = yTierPosition
    Select Case mp_yTierPosition
        Case E_TIERPOSITION.SP_UPPER
            mp_sTierPosition = "UpperTier"
        Case E_TIERPOSITION.SP_MIDDLE
            mp_sTierPosition = "MiddleTier"
        Case E_TIERPOSITION.SP_LOWER
            mp_sTierPosition = "LowerTier"
    End Select
    If bInit = False Then
        mp_bVisible = True
        mp_lFactor = 1
        mp_lHeight = 14
        mp_sTag = ""
        mp_sInterval = "w"
        mp_yTierType = E_TIERTYPE.ST_DAYOFWEEK
        Set mp_oFont = New StdFont
    Else
        Select Case mp_yTierPosition
            Case E_TIERPOSITION.SP_UPPER
                mp_bVisible = mp_oControl.DefaultValues.UpperTier.Visible
                mp_sInterval = mp_oControl.DefaultValues.UpperTier.f_Interval
                mp_lFactor = mp_oControl.DefaultValues.UpperTier.f_Factor
                mp_yTierType = mp_oControl.DefaultValues.UpperTier.TierType
                mp_lHeight = mp_oControl.DefaultValues.UpperTier.Height
                Set mp_oFont = mp_oControl.DefaultValues.UpperTier.Font
            Case E_TIERPOSITION.SP_MIDDLE
                mp_bVisible = mp_oControl.DefaultValues.MiddleTier.Visible
                mp_sInterval = mp_oControl.DefaultValues.MiddleTier.f_Interval
                mp_lFactor = mp_oControl.DefaultValues.MiddleTier.f_Factor
                mp_yTierType = mp_oControl.DefaultValues.MiddleTier.TierType
                mp_lHeight = mp_oControl.DefaultValues.MiddleTier.Height
                Set mp_oFont = mp_oControl.DefaultValues.MiddleTier.Font
            Case E_TIERPOSITION.SP_LOWER
                mp_bVisible = mp_oControl.DefaultValues.LowerTier.Visible
                mp_sInterval = mp_oControl.DefaultValues.LowerTier.f_Interval
                mp_lFactor = mp_oControl.DefaultValues.LowerTier.f_Factor
                mp_yTierType = mp_oControl.DefaultValues.LowerTier.TierType
                mp_lHeight = mp_oControl.DefaultValues.LowerTier.Height
                Set mp_oFont = mp_oControl.DefaultValues.LowerTier.Font
        End Select
    End If
End Sub

Private Sub Class_Terminate()
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Object Member Variables
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Visible() As Boolean
    Visible = mp_bVisible
End Property

Public Property Let Visible(ByVal Value As Boolean)
    mp_bVisible = Value
End Property

Public Property Get Tag() As String
    Tag = mp_sTag
End Property

Public Property Let Tag(ByVal Value As String)
    mp_sTag = Value
End Property

Public Property Get Interval() As String
    Interval = mp_lFactor & mp_sInterval
End Property

Public Property Let Interval(ByVal Value As String)
    mp_oControl.mp_ProcessInterval Value, mp_sInterval, mp_lFactor
    mp_yTierType = E_TIERTYPE.ST_CUSTOM
End Property

Friend Property Get f_Interval() As String
    f_Interval = mp_sInterval
End Property

Friend Property Get f_Factor() As Long
    f_Factor = mp_lFactor
End Property

Public Property Get TierType() As E_TIERTYPE
    TierType = mp_yTierType
End Property

Public Property Let TierType(ByVal Value As E_TIERTYPE)
    mp_yTierType = Value
    Select Case mp_yTierType
        Case E_TIERTYPE.ST_YEAR
            mp_sInterval = "yyyy"
            mp_lFactor = 1
        Case E_TIERTYPE.ST_QUARTER
            mp_sInterval = "q"
            mp_lFactor = 1
        Case E_TIERTYPE.ST_MONTH
            mp_sInterval = "m"
            mp_lFactor = 1
        Case E_TIERTYPE.ST_WEEK
            mp_sInterval = "ww"
            mp_lFactor = 1
        Case E_TIERTYPE.ST_DAYOFWEEK
            mp_sInterval = "w"
            mp_lFactor = 1
        Case E_TIERTYPE.ST_DAY
            mp_sInterval = "d"
            mp_lFactor = 1
        Case E_TIERTYPE.ST_DAYOFYEAR
            mp_sInterval = "y"
            mp_lFactor = 1
        Case E_TIERTYPE.ST_HOUR
            mp_sInterval = "h"
            mp_lFactor = 1
        Case E_TIERTYPE.ST_MINUTE
            mp_sInterval = "n"
            mp_lFactor = 1
    End Select
End Property

Public Property Get Height() As Long
    Height = mp_lHeight
End Property

Public Property Let Height(ByVal Value As Long)
    mp_lHeight = Value
End Property

Public Property Get Font() As Font
    Set Font = mp_oFont
End Property

Public Property Set Font(ByVal Value As Font)
    Set mp_oFont = Value
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Position
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Position()
    Dim dtStart As Date
    Dim dtEnd As Date
    Dim lTop As Integer
    Dim lBottom As Integer
    Dim lTierHeight As Long
    If (mp_bVisible = False) Then
        Exit Sub
    End If
    lTierHeight = Height
    lTop = mp_oTierArea.TimeLine.TiersTickMarksPosition(mp_sTierPosition)
    lBottom = lTop + lTierHeight - 2
    If (mp_oControl.MathLib.GetXCoordinateFromDate(mp_oControl.MathLib.DateTimeAdd(mp_sInterval, mp_lFactor, mp_oTierArea.TimeLine.StartDate)) - mp_oControl.MathLib.GetXCoordinateFromDate(mp_oTierArea.TimeLine.StartDate) > 5) Then
        dtEnd = mp_oControl.MathLib.RoundDate(mp_sInterval, mp_lFactor, mp_oTierArea.TimeLine.StartDate)
        If (mp_oControl.MathLib.GetXCoordinateFromDate(dtEnd) >= mp_oTierArea.TimeLine.f_lStart) Then
            dtStart = mp_oControl.MathLib.DateTimeAdd(mp_sInterval, -mp_lFactor, dtEnd)
            dtStart = mp_oControl.MathLib.RoundDate(mp_sInterval, mp_lFactor, dtStart)
            Draw dtStart, dtEnd, lTop, lBottom
        End If
        Do While (dtEnd < mp_oTierArea.TimeLine.EndDate)
            dtStart = dtEnd
            dtEnd = mp_oControl.MathLib.DateTimeAdd(mp_sInterval, mp_lFactor, dtEnd)
            dtStart = mp_oControl.MathLib.RoundDate(mp_sInterval, mp_lFactor, dtStart)
            dtEnd = mp_oControl.MathLib.RoundDate(mp_sInterval, mp_lFactor, dtEnd)
            Draw dtStart, dtEnd, lTop, lBottom
        Loop
    End If
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Draw
'// ---------------------------------------------------------------------------------------------------------------------

Private Sub Draw(ByVal dtStart As Date, ByVal dtEnd As Date, ByVal lTop As Long, ByVal lBottom As Long)
    Dim lStart As Long
    Dim lEnd As Long
    Dim lStartTrim As Long
    Dim lEndTrim As Long
    Dim sStyleIndex As String
    
    Dim sCaption As String
    Dim lCaptionWidth As Long
    
    Dim clrForeColor As OLE_COLOR
    Dim clrBackColor As OLE_COLOR
    
    lStart = mp_oControl.MathLib.GetXCoordinateFromDate(dtStart)
    lEnd = mp_oControl.MathLib.GetXCoordinateFromDate(dtEnd)
    If (lStart < mp_oTierArea.TimeLine.f_lStart) Then
        lStartTrim = mp_oTierArea.TimeLine.f_lStart
    Else
        lStartTrim = lStart
    End If
    If (lEnd > mp_oTierArea.TimeLine.f_lEnd) Then
        lEndTrim = mp_oTierArea.TimeLine.f_lEnd
    Else
        lEndTrim = lEnd
    End If
    sStyleIndex = ""
    sCaption = ""
    If (mp_yTierType = E_TIERTYPE.ST_CUSTOM) Then
        mp_oControl.FireCustomTierDraw mp_yTierPosition, dtStart, dtEnd, lStart, lEnd, lTop, lBottom, lStartTrim, lEndTrim, mp_oControl.GrphLib.mp_lHdc, sCaption, sStyleIndex
        If (sStyleIndex <> "") Then
            mp_oControl.mp_DrawItem lStart, lEnd, lTop, lBottom, sStyleIndex, sCaption, False, Nothing, lStartTrim, lEndTrim, Nothing
        End If
    Else
        If (mp_sInterval = "yyyy") Then
            clrForeColor = mp_oTierArea.TierAppearance.YearInterval.Item(mp_oControl.MathLib.GetLastDigit(mp_oControl.MathLib.GetYear(dtStart)) + 1).ForeColor
            clrBackColor = mp_oTierArea.TierAppearance.YearInterval.Item(mp_oControl.MathLib.GetLastDigit(mp_oControl.MathLib.GetYear(dtStart)) + 1).BackColor
            sCaption = mp_oControl.StrLib.StrFormat(dtStart, mp_oTierArea.TierFormat.YearIntervalFormat)
        ElseIf (mp_sInterval = "q") Then
            clrForeColor = mp_oTierArea.TierAppearance.QuarterInterval.Item(mp_oControl.MathLib.GetQuarter(dtStart)).ForeColor
            clrBackColor = mp_oTierArea.TierAppearance.QuarterInterval.Item(mp_oControl.MathLib.GetQuarter(dtStart)).BackColor
            sCaption = mp_oControl.StrLib.StrFormat(dtStart, mp_oTierArea.TierFormat.QuarterIntervalFormat)
        ElseIf (mp_sInterval = "m") Then
            clrForeColor = mp_oTierArea.TierAppearance.MonthInterval.Item(mp_oControl.MathLib.GetMonth(dtStart)).ForeColor
            clrBackColor = mp_oTierArea.TierAppearance.MonthInterval.Item(mp_oControl.MathLib.GetMonth(dtStart)).BackColor
            sCaption = mp_oControl.StrLib.StrFormat(dtStart, mp_oTierArea.TierFormat.MonthIntervalFormat)
        ElseIf (mp_sInterval = "ww") Then
            clrForeColor = mp_oTierArea.TierAppearance.WeekInterval.Item(mp_oControl.MathLib.GetLastDigit(mp_oControl.MathLib.GetWeek(dtStart)) + 1).ForeColor
            clrBackColor = mp_oTierArea.TierAppearance.WeekInterval.Item(mp_oControl.MathLib.GetLastDigit(mp_oControl.MathLib.GetWeek(dtStart)) + 1).BackColor
            sCaption = mp_oControl.StrLib.StrFormat(dtStart, mp_oTierArea.TierFormat.WeekIntervalFormat)
        ElseIf (mp_sInterval = "w") Then
            clrBackColor = mp_oTierArea.TierAppearance.DayOfWeekInterval.Item(mp_oControl.MathLib.GetDayOfWeek(dtStart)).BackColor
            clrForeColor = mp_oTierArea.TierAppearance.DayOfWeekInterval.Item(mp_oControl.MathLib.GetDayOfWeek(dtStart)).ForeColor
            sCaption = mp_oControl.StrLib.StrFormat(dtStart, mp_oTierArea.TierFormat.DayOfWeekIntervalFormat)
        ElseIf (mp_sInterval = "d") Then
            clrBackColor = mp_oTierArea.TierAppearance.DayInterval.Item(mp_oControl.MathLib.GetDay(dtStart)).BackColor
            clrForeColor = mp_oTierArea.TierAppearance.DayInterval.Item(mp_oControl.MathLib.GetDay(dtStart)).ForeColor
            sCaption = mp_oControl.StrLib.StrFormat(dtStart, mp_oTierArea.TierFormat.DayIntervalFormat)
        ElseIf (mp_sInterval = "y") Then
            clrBackColor = mp_oTierArea.TierAppearance.DayOfYearInterval.Item(mp_oControl.MathLib.GetDayOfYear(dtStart)).BackColor
            clrForeColor = mp_oTierArea.TierAppearance.DayOfYearInterval.Item(mp_oControl.MathLib.GetDayOfYear(dtStart)).ForeColor
            sCaption = mp_oControl.StrLib.StrFormat(dtStart, mp_oTierArea.TierFormat.DayOfYearIntervalFormat)
        ElseIf (mp_sInterval = "h") Then
            clrBackColor = mp_oTierArea.TierAppearance.HourInterval.Item(mp_oControl.MathLib.GetLastDigit(mp_oControl.MathLib.GetHour(dtStart)) + 1).BackColor
            clrForeColor = mp_oTierArea.TierAppearance.HourInterval.Item(mp_oControl.MathLib.GetLastDigit(mp_oControl.MathLib.GetHour(dtStart)) + 1).ForeColor
            sCaption = mp_oControl.StrLib.StrFormat(dtStart, mp_oTierArea.TierFormat.HourIntervalFormat)
        ElseIf (mp_sInterval = "n") Then
            clrBackColor = mp_oTierArea.TierAppearance.MinuteInterval.Item(mp_oControl.MathLib.GetLastDigit(mp_oControl.MathLib.GetMinute(dtStart)) + 1).BackColor
            clrForeColor = mp_oTierArea.TierAppearance.MinuteInterval.Item(mp_oControl.MathLib.GetLastDigit(mp_oControl.MathLib.GetMinute(dtStart)) + 1).ForeColor
            sCaption = mp_oControl.StrLib.StrFormat(dtStart, mp_oTierArea.TierFormat.MinuteIntervalFormat)
        End If
        mp_oControl.GrphLib.DrawLine lStart, lTop, lEnd, lBottom, GRE_LINETYPE.LT_FILLED, clrBackColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        If lEnd > mp_oTierArea.TimeLine.f_lEnd Then
            lEnd = mp_oTierArea.TimeLine.f_lEnd
        End If
        lCaptionWidth = mp_oControl.mp_lStrWidth(sCaption, mp_oFont)
        If (lEnd - lStart) > lCaptionWidth Then
            mp_oControl.FireTierCaptionDraw sCaption, dtStart, mp_yTierPosition
            mp_oControl.GrphLib.TextOutEx lStartTrim, lTop, lEndTrim, lBottom, sCaption, GRE_HORIZONTALALIGNMENT.HAL_CENTER, GRE_VERTICALALIGNMENT.VAL_CENTER, clrForeColor, mp_oFont
        End If
    End If
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, mp_sTierPosition
    oXML.InitializeWriter
    oXML.WriteProperty "Visible", mp_bVisible
    oXML.WriteProperty "Factor", mp_lFactor
    oXML.WriteProperty "Interval", mp_sInterval
    oXML.WriteProperty "Tag", mp_sTag
    oXML.WriteProperty "TierType", mp_yTierType
    oXML.WriteProperty "yTierPosition", mp_yTierPosition
    oXML.WriteProperty "sTierPosition", mp_sTierPosition
    oXML.WritePropertyFont "Font", mp_oFont
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, mp_sTierPosition
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadProperty "Visible", mp_bVisible
    oXML.ReadProperty "Factor", mp_lFactor
    oXML.ReadProperty "Interval", mp_sInterval
    oXML.ReadProperty "Tag", mp_sTag
    oXML.ReadProperty "TierType", mp_yTierType
    oXML.ReadProperty "yTierPosition", mp_yTierPosition
    oXML.ReadProperty "sTierPosition", mp_sTierPosition
    oXML.ReadPropertyFont "Font", mp_oFont
End Sub

