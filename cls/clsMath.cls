VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMath"
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

'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
End Sub

Private Sub Class_Terminate()
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Date Algebra Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Function DateTimeAdd(ByVal Interval As String, ByVal Number As Long, ByVal dtDate As Date) As Date
    If Interval = "w" Or Interval = "y" Then
        Interval = "d"
    End If
    DateTimeAdd = DateAdd(Interval, Number, dtDate)
End Function

Public Function DateTimeDiff(ByVal Interval As String, ByVal dtDate1 As Date, ByVal dtDate2 As Date) As Long
On Error GoTo ErrorHandler
    If Interval = "w" Or Interval = "y" Then
        Interval = "d"
    End If
    DateTimeDiff = DateDiff(Interval, dtDate1, dtDate2)
Exit Function

ErrorHandler:
    DateTimeDiff = 0
End Function

Public Function NewDateTime(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Hour As Integer, ByVal Minute As Integer, ByVal Second As Integer) As Date
    NewDateTime = DateSerial(Year, Month, Day) + TimeSerial(Hour, Minute, Second)
End Function

Public Function CurrentDateTime() As Date
    CurrentDateTime = Now()
End Function

Public Function GetXCoordinateFromDate(ByVal dtCoordinate As Date) As Long
    GetXCoordinateFromDate = (DateTimeDiff(mp_oControl.CurrentViewObject.f_Interval, mp_oControl.CurrentViewObject.TimeLine.StartDate, dtCoordinate) / mp_oControl.CurrentViewObject.f_Factor) + mp_oControl.CurrentViewObject.TimeLine.f_lStart
End Function

Public Function GetDateFromXCoordinate(ByVal XCoordinate As Long) As Date
    GetDateFromXCoordinate = DateTimeAdd(mp_oControl.CurrentViewObject.f_Interval, (XCoordinate - mp_oControl.CurrentViewObject.TimeLine.f_lStart) * mp_oControl.CurrentViewObject.f_Factor, mp_oControl.CurrentViewObject.TimeLine.StartDate)
End Function

'// ---------------------------------------------------------------------------------------------------------------------
'// DatePart Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetSecond(ByVal dtDate As Date) As Integer
    GetSecond = Second(dtDate)
End Function

Public Function GetHour(ByVal dtDate As Date) As Integer
    GetHour = Hour(dtDate)
End Function

Public Function GetMinute(ByVal dtDate As Date) As Integer
    GetMinute = Minute(dtDate)
End Function

Public Function GetDay(ByVal dtDate As Date) As Integer
    GetDay = Day(dtDate)
End Function

Public Function GetDayOfWeek(ByVal dtDate As Date) As Integer
    GetDayOfWeek = Weekday(dtDate)
End Function

Public Function GetDayOfYear(ByVal dtDate As Date) As Integer
    GetDayOfYear = DatePart("y", dtDate)
End Function

Public Function GetWeek(ByVal dtDate As Date) As Integer
    GetWeek = DatePart("ww", dtDate)
End Function

Public Function GetMonth(ByVal dtDate As Date) As Integer
    GetMonth = Month(dtDate)
End Function

Public Function GetQuarter(ByVal dtDate As Date) As Integer
    Select Case Month(dtDate)
        Case 1 To 3
            GetQuarter = 1
        Case 4 To 6
            GetQuarter = 2
        Case 7 To 9
            GetQuarter = 3
        Case 10 To 12
            GetQuarter = 4
    End Select
End Function

Public Function GetYear(ByVal dtDate As Date) As Integer
    GetYear = Year(dtDate)
End Function

Public Function GetLastDigit(ByVal Number As Long) As Long
    Dim sReturn As String
    sReturn = CStr(Number)
    GetLastDigit = CLng(Right$(sReturn, 1))
End Function

'// ---------------------------------------------------------------------------------------------------------------------
'// Round Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Function RoundDate(ByVal Interval As String, ByVal Number As Long, ByVal dtDate As Date) As Date
    Dim lBuffer As Long
    Dim lBuffer2 As Long
    If (Interval = "s") Then
        lBuffer = GetSecond(dtDate)
        lBuffer2 = Round(lBuffer, Number)
        RoundDate = DateTimeAdd("s", lBuffer2 - lBuffer, dtDate)
    ElseIf (Interval = "n") Then
        Select Case Number
            Case 1
                dtDate = RoundDate("s", 60, dtDate)
                lBuffer = GetSecond(dtDate)
                lBuffer2 = Round(lBuffer, 60)
                RoundDate = DateTimeAdd("s", lBuffer2 - lBuffer, dtDate)
            Case Else
                dtDate = RoundDate("n", 1, dtDate)
                lBuffer = GetMinute(dtDate)
                lBuffer2 = Round(lBuffer, Number)
                RoundDate = DateTimeAdd("n", lBuffer2 - lBuffer, dtDate)
        End Select
    ElseIf (Interval = "h") Then
        Select Case Number
            Case 1
                dtDate = RoundDate("n", 1, dtDate)
                lBuffer = GetMinute(dtDate)
                lBuffer2 = Round(lBuffer, 60)
                RoundDate = DateTimeAdd("n", lBuffer2 - lBuffer, dtDate)
            Case Else
                dtDate = RoundDate("h", 1, dtDate)
                lBuffer = GetHour(dtDate)
                lBuffer2 = Round(lBuffer, Number)
                RoundDate = DateTimeAdd("h", lBuffer2 - lBuffer, dtDate)
        End Select
    ElseIf (Interval = "d") Then
        Select Case Number
            Case 1
                dtDate = RoundDate("h", 1, dtDate)
                lBuffer = GetHour(dtDate)
                lBuffer2 = Round(lBuffer, 24)
                RoundDate = DateTimeAdd("h", lBuffer2 - lBuffer, dtDate)
            Case Else
                dtDate = RoundDate("d", 1, dtDate)
                lBuffer = GetDay(dtDate)
                lBuffer2 = Round(lBuffer, Number)
                RoundDate = DateTimeAdd("d", lBuffer2 - lBuffer, dtDate)
        End Select
    ElseIf (Interval = "ww") Then
        Select Case Number
            Case 1
                dtDate = RoundDate("d", 1, dtDate)
                lBuffer = GetDayOfWeek(dtDate)
                If lBuffer <= 3 Then
                    dtDate = DateTimeAdd("d", -(lBuffer - 1), dtDate)
                ElseIf lBuffer >= 4 Then
                    dtDate = DateTimeAdd("d", 8 - lBuffer, dtDate)
                End If
                RoundDate = dtDate
            Case Else
                dtDate = RoundDate("ww", 1, dtDate)
                lBuffer = GetDay(dtDate)
                lBuffer2 = Round(lBuffer, Number)
                RoundDate = DateTimeAdd("ww", lBuffer2 - lBuffer, dtDate)
        End Select
    ElseIf (Interval = "m") Then
        Select Case Number
            Case 1
                Dim dtNextMonth As Date
                dtDate = RoundDate("d", 1, dtDate)
                lBuffer = GetDay(dtDate)
                If lBuffer = 1 Then
                    RoundDate = dtDate
                ElseIf lBuffer >= 15 Then
                    dtNextMonth = DateTimeAdd("m", 1, dtDate)
                    RoundDate = NewDateTime(Year(dtNextMonth), Month(dtNextMonth), 1, 0, 0, 0)
                Else
                    RoundDate = NewDateTime(Year(dtDate), Month(dtDate), 1, 0, 0, 0)
                End If
            Case Else
                dtDate = RoundDate("m", 1, dtDate)
                lBuffer = GetMonth(dtDate)
                lBuffer2 = Round(lBuffer - 1, Number) + 1
                RoundDate = DateTimeAdd("m", lBuffer2 - lBuffer, dtDate)
        End Select
    ElseIf (Interval = "q") Then
        dtDate = RoundDate("d", 1, dtDate)
        RoundDate = RoundDate("m", 3, dtDate)
    ElseIf (Interval = "yyyy") Then
        Select Case Number
            Case 1
                dtDate = RoundDate("m", 1, dtDate)
                lBuffer = GetMonth(dtDate)
                lBuffer2 = Round(lBuffer, 11) + 1
                If lBuffer = 1 Then
                    RoundDate = NewDateTime(Year(dtDate), 1, 1, 0, 0, 0)
                ElseIf lBuffer = 12 Then
                    RoundDate = NewDateTime(Year(dtDate) + 1, 1, 1, 0, 0, 0)
                End If
            Case Else
                dtDate = RoundDate("yyyy", 1, dtDate)
                lBuffer = GetYear(dtDate)
                lBuffer2 = Round(lBuffer, Number)
                RoundDate = DateTimeAdd("yyyy", lBuffer2 - lBuffer, dtDate)
        End Select
    Else
        RoundDate = dtDate
    End If
End Function

Public Function RoundDouble(ByVal dParam As Double) As Long
    RoundDouble = dParam
End Function

Friend Function Round(ByVal v_lNumberToRound As Long, ByVal v_lRoundTo As Long) As Long
    Dim lRoundToHalf As Long
    Dim lMultiplier As Long
    Do While v_lNumberToRound > v_lRoundTo
        v_lNumberToRound = v_lNumberToRound - v_lRoundTo
        lMultiplier = lMultiplier + 1
    Loop
    lRoundToHalf = Abs(Int(-(v_lRoundTo / 2)))
    If v_lNumberToRound >= lRoundToHalf Then
        v_lNumberToRound = v_lRoundTo
    Else
        v_lNumberToRound = 0
    End If
    Round = (v_lRoundTo * lMultiplier) + v_lNumberToRound
End Function

Friend Function lAbs(ByVal Number As Long) As Long
    lAbs = Abs(Number)
End Function




