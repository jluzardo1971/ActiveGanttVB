VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTierFormat"
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
Private mp_sMinuteIntervalFormat As String
Private mp_sHourIntervalFormat As String
Private mp_sDayIntervalFormat As String
Private mp_sDayOfWeekIntervalFormat As String
Private mp_sDayOfYearIntervalFormat As String
Private mp_sWeekIntervalFormat As String
Private mp_sMonthIntervalFormat As String
Private mp_sQuarterIntervalFormat As String
Private mp_sYearIntervalFormat As String

'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl, ByVal bInit As Boolean)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    If bInit = False Then
        mp_sMinuteIntervalFormat = "nn"
        mp_sHourIntervalFormat = "hh:nn"
        mp_sDayIntervalFormat = "d"
        mp_sDayOfWeekIntervalFormat = "dddd d"
        mp_sDayOfYearIntervalFormat = "y"
        mp_sWeekIntervalFormat = "ww"
        mp_sMonthIntervalFormat = "mmmm yyyy"
        mp_sQuarterIntervalFormat = "q""Q"" yyyy"
        mp_sYearIntervalFormat = "yyyy"
    Else
        mp_sMinuteIntervalFormat = mp_oControl.DefaultValues.TierFormat.MinuteIntervalFormat
        mp_sHourIntervalFormat = mp_oControl.DefaultValues.TierFormat.HourIntervalFormat
        mp_sDayIntervalFormat = mp_oControl.DefaultValues.TierFormat.DayIntervalFormat
        mp_sDayOfWeekIntervalFormat = mp_oControl.DefaultValues.TierFormat.DayOfWeekIntervalFormat
        mp_sDayOfYearIntervalFormat = mp_oControl.DefaultValues.TierFormat.DayOfYearIntervalFormat
        mp_sWeekIntervalFormat = mp_oControl.DefaultValues.TierFormat.WeekIntervalFormat
        mp_sMonthIntervalFormat = mp_oControl.DefaultValues.TierFormat.MonthIntervalFormat
        mp_sQuarterIntervalFormat = mp_oControl.DefaultValues.TierFormat.QuarterIntervalFormat
        mp_sYearIntervalFormat = mp_oControl.DefaultValues.TierFormat.YearIntervalFormat
    End If
End Sub

Private Sub Class_Terminate()
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Object Member Variables
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get MinuteIntervalFormat() As String
    MinuteIntervalFormat = mp_sMinuteIntervalFormat
End Property

Public Property Let MinuteIntervalFormat(ByVal Value As String)
    mp_sMinuteIntervalFormat = Value
End Property

Public Property Get HourIntervalFormat() As String
    HourIntervalFormat = mp_sHourIntervalFormat
End Property

Public Property Let HourIntervalFormat(ByVal Value As String)
    mp_sHourIntervalFormat = Value
End Property

Public Property Get DayIntervalFormat() As String
    DayIntervalFormat = mp_sDayIntervalFormat
End Property

Public Property Let DayIntervalFormat(ByVal Value As String)
    mp_sDayIntervalFormat = Value
End Property

Public Property Get DayOfWeekIntervalFormat() As String
    DayOfWeekIntervalFormat = mp_sDayOfWeekIntervalFormat
End Property

Public Property Let DayOfWeekIntervalFormat(ByVal Value As String)
    mp_sDayOfWeekIntervalFormat = Value
End Property

Public Property Get DayOfYearIntervalFormat() As String
    DayOfYearIntervalFormat = mp_sDayOfYearIntervalFormat
End Property

Public Property Let DayOfYearIntervalFormat(ByVal Value As String)
    mp_sDayOfYearIntervalFormat = Value
End Property


Public Property Get WeekIntervalFormat() As String
    WeekIntervalFormat = mp_sWeekIntervalFormat
End Property

Public Property Let WeekIntervalFormat(ByVal Value As String)
    mp_sWeekIntervalFormat = Value
End Property

Public Property Get MonthIntervalFormat() As String
    MonthIntervalFormat = mp_sMonthIntervalFormat
End Property

Public Property Let MonthIntervalFormat(ByVal Value As String)
    mp_sMonthIntervalFormat = Value
End Property

Public Property Get QuarterIntervalFormat() As String
    QuarterIntervalFormat = mp_sQuarterIntervalFormat
End Property

Public Property Let QuarterIntervalFormat(ByVal Value As String)
    mp_sQuarterIntervalFormat = Value
End Property

Public Property Get YearIntervalFormat() As String
    YearIntervalFormat = mp_sYearIntervalFormat
End Property

Public Property Let YearIntervalFormat(ByVal Value As String)
    mp_sYearIntervalFormat = Value
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TierFormat"
    oXML.InitializeWriter
    oXML.WriteProperty "MinuteIntervalFormat", mp_sMinuteIntervalFormat
    oXML.WriteProperty "HourIntervalFormat", mp_sHourIntervalFormat
    oXML.WriteProperty "DayIntervalFormat", mp_sDayIntervalFormat
    oXML.WriteProperty "DayOfWeekIntervalFormat", mp_sDayOfWeekIntervalFormat
    oXML.WriteProperty "DayOfYearIntervalFormat", mp_sDayOfYearIntervalFormat
    oXML.WriteProperty "WeekIntervalFormat", mp_sWeekIntervalFormat
    oXML.WriteProperty "MonthIntervalFormat", mp_sMonthIntervalFormat
    oXML.WriteProperty "QuarterIntervalFormat", mp_sQuarterIntervalFormat
    oXML.WriteProperty "YearIntervalFormat", mp_sYearIntervalFormat
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TierFormat"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadProperty "MinuteIntervalFormat", mp_sMinuteIntervalFormat
    oXML.ReadProperty "HourIntervalFormat", mp_sHourIntervalFormat
    oXML.ReadProperty "DayIntervalFormat", mp_sDayIntervalFormat
    oXML.ReadProperty "DayOfWeekIntervalFormat", mp_sDayOfWeekIntervalFormat
    oXML.ReadProperty "DayOfYearIntervalFormat", mp_sDayOfYearIntervalFormat
    oXML.ReadProperty "WeekIntervalFormat", mp_sWeekIntervalFormat
    oXML.ReadProperty "MonthIntervalFormat", mp_sMonthIntervalFormat
    oXML.ReadProperty "QuarterIntervalFormat", mp_sQuarterIntervalFormat
    oXML.ReadProperty "YearIntervalFormat", mp_sYearIntervalFormat
End Sub
