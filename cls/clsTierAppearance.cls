VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTierAppearance"
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
Private mp_oMinuteInterval As clsTierColors
Private mp_oHourInterval As clsTierColors
Private mp_oDayInterval As clsTierColors
Private mp_oDayOfWeekInterval As clsTierColors
Private mp_oDayOfYearInterval As clsTierColors
Private mp_oWeekInterval As clsTierColors
Private mp_oMonthInterval As clsTierColors
Private mp_oQuarterInterval As clsTierColors
Private mp_oYearInterval As clsTierColors

'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl, ByVal bInit As Boolean)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Dim lIndex As Long
    If bInit = False Then
        Set mp_oMinuteInterval = New clsTierColors
        mp_oMinuteInterval.Initialize mp_oControl, E_TIERTYPE.ST_MINUTE
        mp_oMinuteInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oMinuteInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oMinuteInterval.Add GRE_COLORS.CLR_CYAN, GRE_COLORS.CLR_BLACK
        mp_oMinuteInterval.Add GRE_COLORS.CLR_PALETURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oMinuteInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oMinuteInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oMinuteInterval.Add GRE_COLORS.CLR_CYAN, GRE_COLORS.CLR_BLACK
        mp_oMinuteInterval.Add GRE_COLORS.CLR_PALETURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oMinuteInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oMinuteInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK
        
        Set mp_oHourInterval = New clsTierColors
        mp_oHourInterval.Initialize mp_oControl, E_TIERTYPE.ST_HOUR
        mp_oHourInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oHourInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oHourInterval.Add GRE_COLORS.CLR_CYAN, GRE_COLORS.CLR_BLACK
        mp_oHourInterval.Add GRE_COLORS.CLR_PALETURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oHourInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oHourInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oHourInterval.Add GRE_COLORS.CLR_CYAN, GRE_COLORS.CLR_BLACK
        mp_oHourInterval.Add GRE_COLORS.CLR_PALETURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oHourInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oHourInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK
        
        
        Set mp_oDayInterval = New clsTierColors
        mp_oDayInterval.Initialize mp_oControl, E_TIERTYPE.ST_DAY
        mp_oDayInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oDayInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oDayInterval.Add GRE_COLORS.CLR_CYAN, GRE_COLORS.CLR_BLACK
        mp_oDayInterval.Add GRE_COLORS.CLR_PALETURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oDayInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oDayInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oDayInterval.Add GRE_COLORS.CLR_CYAN, GRE_COLORS.CLR_BLACK
        mp_oDayInterval.Add GRE_COLORS.CLR_PALETURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oDayInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oDayInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK
        
        
        Set mp_oDayOfWeekInterval = New clsTierColors
        mp_oDayOfWeekInterval.Initialize mp_oControl, E_TIERTYPE.ST_DAYOFWEEK
        mp_oDayOfWeekInterval.Add GRE_COLORS.CLR_CORNFLOWERBLUE, GRE_COLORS.CLR_BLACK, "Sunday"
        mp_oDayOfWeekInterval.Add GRE_COLORS.CLR_MEDIUMSLATEBLUE, GRE_COLORS.CLR_BLACK, "Monday"
        mp_oDayOfWeekInterval.Add GRE_COLORS.CLR_SLATEBLUE, GRE_COLORS.CLR_WHITE, "Tuesday"
        mp_oDayOfWeekInterval.Add GRE_COLORS.CLR_ROYALBLUE, GRE_COLORS.CLR_WHITE, "Wednesday"
        mp_oDayOfWeekInterval.Add GRE_COLORS.CLR_SKYBLUE, GRE_COLORS.CLR_WHITE, "Thursday"
        mp_oDayOfWeekInterval.Add GRE_COLORS.CLR_DEEPSKYBLUE, GRE_COLORS.CLR_BLACK, "Friday"
        mp_oDayOfWeekInterval.Add GRE_COLORS.CLR_DODGERBLUE, GRE_COLORS.CLR_BLACK, "Saturday"
        
        Set mp_oDayOfYearInterval = New clsTierColors
        mp_oDayOfYearInterval.Initialize mp_oControl, E_TIERTYPE.ST_DAYOFYEAR
        mp_oDayOfYearInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oDayOfYearInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oDayOfYearInterval.Add GRE_COLORS.CLR_CYAN, GRE_COLORS.CLR_BLACK
        mp_oDayOfYearInterval.Add GRE_COLORS.CLR_PALETURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oDayOfYearInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oDayOfYearInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oDayOfYearInterval.Add GRE_COLORS.CLR_CYAN, GRE_COLORS.CLR_BLACK
        mp_oDayOfYearInterval.Add GRE_COLORS.CLR_PALETURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oDayOfYearInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oDayOfYearInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK
        
        Set mp_oWeekInterval = New clsTierColors
        mp_oWeekInterval.Initialize mp_oControl, E_TIERTYPE.ST_WEEK
        mp_oWeekInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oWeekInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oWeekInterval.Add GRE_COLORS.CLR_CYAN, GRE_COLORS.CLR_BLACK
        mp_oWeekInterval.Add GRE_COLORS.CLR_PALETURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oWeekInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oWeekInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oWeekInterval.Add GRE_COLORS.CLR_CYAN, GRE_COLORS.CLR_BLACK
        mp_oWeekInterval.Add GRE_COLORS.CLR_PALETURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oWeekInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oWeekInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK
        
        
        Set mp_oMonthInterval = New clsTierColors
        mp_oMonthInterval.Initialize mp_oControl, E_TIERTYPE.ST_MONTH
        mp_oMonthInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK, "January"
        mp_oMonthInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK, "February"
        mp_oMonthInterval.Add GRE_COLORS.CLR_CYAN, GRE_COLORS.CLR_BLACK, "March"
        mp_oMonthInterval.Add GRE_COLORS.CLR_PALETURQUOISE, GRE_COLORS.CLR_BLACK, "April"
        mp_oMonthInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK, "May"
        mp_oMonthInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK, "June"
        mp_oMonthInterval.Add GRE_COLORS.CLR_CYAN, GRE_COLORS.CLR_BLACK, "July"
        mp_oMonthInterval.Add GRE_COLORS.CLR_PALETURQUOISE, GRE_COLORS.CLR_BLACK, "August"
        mp_oMonthInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK, "September"
        mp_oMonthInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK, "October"
        mp_oMonthInterval.Add GRE_COLORS.CLR_CYAN, GRE_COLORS.CLR_BLACK, "November"
        mp_oMonthInterval.Add GRE_COLORS.CLR_PALETURQUOISE, GRE_COLORS.CLR_BLACK, "December"
        
        Set mp_oQuarterInterval = New clsTierColors
        mp_oQuarterInterval.Initialize mp_oControl, E_TIERTYPE.ST_QUARTER
        mp_oQuarterInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oQuarterInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oQuarterInterval.Add GRE_COLORS.CLR_CYAN, GRE_COLORS.CLR_BLACK
        mp_oQuarterInterval.Add GRE_COLORS.CLR_PALETURQUOISE, GRE_COLORS.CLR_BLACK
        
        Set mp_oYearInterval = New clsTierColors
        mp_oYearInterval.Initialize mp_oControl, E_TIERTYPE.ST_YEAR
        mp_oYearInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oYearInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oYearInterval.Add GRE_COLORS.CLR_CYAN, GRE_COLORS.CLR_BLACK
        mp_oYearInterval.Add GRE_COLORS.CLR_PALETURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oYearInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oYearInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oYearInterval.Add GRE_COLORS.CLR_CYAN, GRE_COLORS.CLR_BLACK
        mp_oYearInterval.Add GRE_COLORS.CLR_PALETURQUOISE, GRE_COLORS.CLR_BLACK
        mp_oYearInterval.Add GRE_COLORS.CLR_CADETBLUE, GRE_COLORS.CLR_BLACK
        mp_oYearInterval.Add GRE_COLORS.CLR_DARKTURQUOISE, GRE_COLORS.CLR_BLACK

    Else
        Set mp_oMinuteInterval = New clsTierColors
        mp_oMinuteInterval.Initialize mp_oControl, E_TIERTYPE.ST_MINUTE
        For lIndex = 1 To mp_oControl.DefaultValues.TierAppearance.MinuteInterval.Count
            mp_oMinuteInterval.Add mp_oControl.DefaultValues.TierAppearance.MinuteInterval.Item(lIndex).BackColor, mp_oControl.DefaultValues.TierAppearance.MinuteInterval.Item(lIndex).ForeColor, mp_oControl.DefaultValues.TierAppearance.MinuteInterval.Item(lIndex).Key
        Next lIndex
        Set mp_oHourInterval = New clsTierColors
        mp_oHourInterval.Initialize mp_oControl, E_TIERTYPE.ST_HOUR
        For lIndex = 1 To mp_oControl.DefaultValues.TierAppearance.HourInterval.Count
            mp_oHourInterval.Add mp_oControl.DefaultValues.TierAppearance.HourInterval.Item(lIndex).BackColor, mp_oControl.DefaultValues.TierAppearance.HourInterval.Item(lIndex).ForeColor, mp_oControl.DefaultValues.TierAppearance.HourInterval.Item(lIndex).Key
        Next lIndex
        Set mp_oDayInterval = New clsTierColors
        mp_oDayInterval.Initialize mp_oControl, E_TIERTYPE.ST_DAY
        For lIndex = 1 To mp_oControl.DefaultValues.TierAppearance.DayInterval.Count
            mp_oDayInterval.Add mp_oControl.DefaultValues.TierAppearance.DayInterval.Item(lIndex).BackColor, mp_oControl.DefaultValues.TierAppearance.DayInterval.Item(lIndex).ForeColor, mp_oControl.DefaultValues.TierAppearance.DayInterval.Item(lIndex).Key
        Next lIndex
        Set mp_oDayOfWeekInterval = New clsTierColors
        mp_oDayOfWeekInterval.Initialize mp_oControl, E_TIERTYPE.ST_DAYOFWEEK
        For lIndex = 1 To mp_oControl.DefaultValues.TierAppearance.DayOfWeekInterval.Count
            mp_oDayOfWeekInterval.Add mp_oControl.DefaultValues.TierAppearance.DayOfWeekInterval.Item(lIndex).BackColor, mp_oControl.DefaultValues.TierAppearance.DayOfWeekInterval.Item(lIndex).ForeColor, mp_oControl.DefaultValues.TierAppearance.DayOfWeekInterval.Item(lIndex).Key
        Next lIndex
        Set mp_oDayOfYearInterval = New clsTierColors
        mp_oDayOfYearInterval.Initialize mp_oControl, E_TIERTYPE.ST_DAYOFYEAR
        For lIndex = 1 To mp_oControl.DefaultValues.TierAppearance.DayOfYearInterval.Count
            mp_oDayOfYearInterval.Add mp_oControl.DefaultValues.TierAppearance.DayOfYearInterval.Item(lIndex).BackColor, mp_oControl.DefaultValues.TierAppearance.DayOfYearInterval.Item(lIndex).ForeColor, mp_oControl.DefaultValues.TierAppearance.DayOfYearInterval.Item(lIndex).Key
        Next lIndex
        Set mp_oWeekInterval = New clsTierColors
        mp_oWeekInterval.Initialize mp_oControl, E_TIERTYPE.ST_WEEK
        For lIndex = 1 To mp_oControl.DefaultValues.TierAppearance.WeekInterval.Count
            mp_oWeekInterval.Add mp_oControl.DefaultValues.TierAppearance.WeekInterval.Item(lIndex).BackColor, mp_oControl.DefaultValues.TierAppearance.WeekInterval.Item(lIndex).ForeColor, mp_oControl.DefaultValues.TierAppearance.WeekInterval.Item(lIndex).Key
        Next lIndex
        Set mp_oMonthInterval = New clsTierColors
        mp_oMonthInterval.Initialize mp_oControl, E_TIERTYPE.ST_MONTH
        For lIndex = 1 To mp_oControl.DefaultValues.TierAppearance.MonthInterval.Count
            mp_oMonthInterval.Add mp_oControl.DefaultValues.TierAppearance.MonthInterval.Item(lIndex).BackColor, mp_oControl.DefaultValues.TierAppearance.MonthInterval.Item(lIndex).ForeColor, mp_oControl.DefaultValues.TierAppearance.MonthInterval.Item(lIndex).Key
        Next lIndex
        Set mp_oQuarterInterval = New clsTierColors
        mp_oQuarterInterval.Initialize mp_oControl, E_TIERTYPE.ST_QUARTER
        For lIndex = 1 To mp_oControl.DefaultValues.TierAppearance.QuarterInterval.Count
            mp_oQuarterInterval.Add mp_oControl.DefaultValues.TierAppearance.QuarterInterval.Item(lIndex).BackColor, mp_oControl.DefaultValues.TierAppearance.QuarterInterval.Item(lIndex).ForeColor, mp_oControl.DefaultValues.TierAppearance.QuarterInterval.Item(lIndex).Key
        Next lIndex
        Set mp_oYearInterval = New clsTierColors
        mp_oYearInterval.Initialize mp_oControl, E_TIERTYPE.ST_YEAR
        For lIndex = 1 To mp_oControl.DefaultValues.TierAppearance.YearInterval.Count
            mp_oYearInterval.Add mp_oControl.DefaultValues.TierAppearance.YearInterval.Item(lIndex).BackColor, mp_oControl.DefaultValues.TierAppearance.YearInterval.Item(lIndex).ForeColor, mp_oControl.DefaultValues.TierAppearance.YearInterval.Item(lIndex).Key
        Next lIndex
    End If
End Sub

Private Sub Class_Terminate()
    Set mp_oControl = Nothing
    Set mp_oMinuteInterval = Nothing
    Set mp_oHourInterval = Nothing
    Set mp_oDayInterval = Nothing
    Set mp_oDayOfWeekInterval = Nothing
    Set mp_oDayOfYearInterval = Nothing
    Set mp_oWeekInterval = Nothing
    Set mp_oMonthInterval = Nothing
    Set mp_oQuarterInterval = Nothing
    Set mp_oYearInterval = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Object Member Variables
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get MinuteInterval() As clsTierColors
    Set MinuteInterval = mp_oMinuteInterval
End Property

Public Property Get HourInterval() As clsTierColors
    Set HourInterval = mp_oHourInterval
End Property

Public Property Get DayInterval() As clsTierColors
    Set DayInterval = mp_oDayInterval
End Property

Public Property Get DayOfWeekInterval() As clsTierColors
    Set DayOfWeekInterval = mp_oDayOfWeekInterval
End Property

Public Property Get DayOfYearInterval() As clsTierColors
    Set DayOfYearInterval = mp_oDayOfYearInterval
End Property

Public Property Get WeekInterval() As clsTierColors
    Set WeekInterval = mp_oWeekInterval
End Property

Public Property Get MonthInterval() As clsTierColors
    Set MonthInterval = mp_oMonthInterval
End Property

Public Property Get QuarterInterval() As clsTierColors
    Set QuarterInterval = mp_oQuarterInterval
End Property

Public Property Get YearInterval() As clsTierColors
    Set YearInterval = mp_oYearInterval
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TierAppearance"
    oXML.InitializeWriter
    oXML.WriteObject mp_oMinuteInterval.GetXML
    oXML.WriteObject mp_oHourInterval.GetXML
    oXML.WriteObject mp_oDayInterval.GetXML
    oXML.WriteObject mp_oDayOfWeekInterval.GetXML
    oXML.WriteObject mp_oDayOfYearInterval.GetXML
    oXML.WriteObject mp_oWeekInterval.GetXML
    oXML.WriteObject mp_oMonthInterval.GetXML
    oXML.WriteObject mp_oQuarterInterval.GetXML
    oXML.WriteObject mp_oYearInterval.GetXML
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TierAppearance"
    oXML.SetXML sXML
    oXML.InitializeReader
    mp_oMinuteInterval.SetXML oXML.ReadObject("TierColorsMinuteInterval")
    mp_oHourInterval.SetXML oXML.ReadObject("TierColorsHourInterval")
    mp_oDayInterval.SetXML oXML.ReadObject("TierColorsDayInterval")
    mp_oDayOfWeekInterval.SetXML oXML.ReadObject("TierColorsDayOfWeekInterval")
    mp_oDayOfYearInterval.SetXML oXML.ReadObject("TierColorsDayOfYearInterval")
    mp_oWeekInterval.SetXML oXML.ReadObject("TierColorsWeekInterval")
    mp_oMonthInterval.SetXML oXML.ReadObject("TierColorsMonthInterval")
    mp_oQuarterInterval.SetXML oXML.ReadObject("TierColorsQuarterInterval")
    mp_oYearInterval.SetXML oXML.ReadObject("TierColorsYearInterval")
End Sub
