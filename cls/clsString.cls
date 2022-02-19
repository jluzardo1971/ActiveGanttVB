VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsString"
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

Private Const LOCALE_STIME = &H1E        '  time separator
Private Const LOCALE_SDATE = &H1D        '  date separator
Private Const LOCALE_S1159 = &H28        '  AM designator
Private Const LOCALE_S2359 = &H29        '  PM designator
Private Const LOCALE_SSHORTDATE = &H1F        '  short date format string
Private Const LOCALE_SLONGDATE = &H20        '  long date format string
Private Const LOCALE_STIMEFORMAT = &H1003      '  time format string

Private Const LOCALE_SABBREVDAYNAME1 = &H31        '  abbreviated name for Monday
Private Const LOCALE_SABBREVDAYNAME2 = &H32        '  abbreviated name for Tuesday
Private Const LOCALE_SABBREVDAYNAME3 = &H33        '  abbreviated name for Wednesday
Private Const LOCALE_SABBREVDAYNAME4 = &H34        '  abbreviated name for Thursday
Private Const LOCALE_SABBREVDAYNAME5 = &H35        '  abbreviated name for Friday
Private Const LOCALE_SABBREVDAYNAME6 = &H36        '  abbreviated name for Saturday
Private Const LOCALE_SABBREVDAYNAME7 = &H37        '  abbreviated name for Sunday
Private Const LOCALE_SDAYNAME1 = &H2A        '  long name for Monday
Private Const LOCALE_SDAYNAME2 = &H2B        '  long name for Tuesday
Private Const LOCALE_SDAYNAME3 = &H2C        '  long name for Wednesday
Private Const LOCALE_SDAYNAME4 = &H2D        '  long name for Thursday
Private Const LOCALE_SDAYNAME5 = &H2E        '  long name for Friday
Private Const LOCALE_SDAYNAME6 = &H2F        '  long name for Saturday
Private Const LOCALE_SDAYNAME7 = &H30        '  long name for Sunday
Private Const LOCALE_SABBREVMONTHNAME1 = &H44        '  abbreviated name for January
Private Const LOCALE_SABBREVMONTHNAME2 = &H45        '  abbreviated name for February
Private Const LOCALE_SABBREVMONTHNAME3 = &H46        '  abbreviated name for March
Private Const LOCALE_SABBREVMONTHNAME4 = &H47        '  abbreviated name for April
Private Const LOCALE_SABBREVMONTHNAME5 = &H48        '  abbreviated name for May
Private Const LOCALE_SABBREVMONTHNAME6 = &H49        '  abbreviated name for June
Private Const LOCALE_SABBREVMONTHNAME7 = &H4A        '  abbreviated name for July
Private Const LOCALE_SABBREVMONTHNAME8 = &H4B        '  abbreviated name for August
Private Const LOCALE_SABBREVMONTHNAME9 = &H4C        '  abbreviated name for September
Private Const LOCALE_SABBREVMONTHNAME10 = &H4D        '  abbreviated name for October
Private Const LOCALE_SABBREVMONTHNAME11 = &H4E        '  abbreviated name for November
Private Const LOCALE_SABBREVMONTHNAME12 = &H4F        '  abbreviated name for December
Private Const LOCALE_SMONTHNAME1 = &H38        '  long name for January
Private Const LOCALE_SMONTHNAME2 = &H39        '  long name for February
Private Const LOCALE_SMONTHNAME3 = &H3A        '  long name for March
Private Const LOCALE_SMONTHNAME4 = &H3B        '  long name for April
Private Const LOCALE_SMONTHNAME5 = &H3C        '  long name for May
Private Const LOCALE_SMONTHNAME6 = &H3D        '  long name for June
Private Const LOCALE_SMONTHNAME7 = &H3E        '  long name for July
Private Const LOCALE_SMONTHNAME8 = &H3F        '  long name for August
Private Const LOCALE_SMONTHNAME9 = &H40        '  long name for September
Private Const LOCALE_SMONTHNAME10 = &H41        '  long name for October
Private Const LOCALE_SMONTHNAME11 = &H42        '  long name for November
Private Const LOCALE_SMONTHNAME12 = &H43        '  long name for December

Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

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
    '//
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Function StrFormat(ByVal Expression As Variant, ByVal sFormat As String) As String
    If TypeName(Expression) = "Date" Then
        StrFormat = mp_FormatString(Expression, sFormat)
    Else
        StrFormat = Format(Expression, sFormat)
    End If
End Function

Public Function StrLeft(ByVal Expression As String, ByVal Length As Long) As String
    StrLeft = Left(Expression, Length)
End Function

Public Function StrRight(ByVal Expression As String, ByVal Length As Long) As String
    StrRight = Right(Expression, Length)
End Function

Public Function StrMid(ByVal Expression As String, ByVal Start As Long, ByVal Length As Long) As String
    StrMid = mid(Expression, Start, Length)
End Function

Public Function StrLowerCase(ByVal Expression As String) As String
   StrLowerCase = LCase(Expression)
End Function

Public Function StrUpperCase(ByVal Expression As String) As String
   StrUpperCase = UCase(Expression)
End Function

Public Function StrIsNumeric(ByVal Expression As String) As Boolean
    StrIsNumeric = IsNumeric(Expression)
End Function

Public Function StrCLng(ByVal Expression As String) As Long
    StrCLng = CLng(Expression)
End Function

Public Function StrCStr(ByVal Expression As Long) As String
    StrCStr = CStr(Expression)
End Function

Public Function StrTrim(ByVal Expression As String) As String
    StrTrim = Trim(Expression)
End Function

Public Function StrReplace(ByVal sExpression As String, ByVal sFind As String, ByVal sReplace As String) As String
    StrReplace = Replace(sExpression, sFind, sReplace)
End Function

Public Function StrLen(ByVal Expression As String) As Long
    StrLen = Len(Expression)
End Function

'// ---------------------------------------------------------------------------------------------------------------------
'// Private Date Format Methods
'// ---------------------------------------------------------------------------------------------------------------------

Private Function mp_FormatString(ByVal dtExpression As Date, ByVal sFormat As String) As String
    Dim oResultsCollection As Collection
    Dim iPosition As Long
    Dim sChar As String
    Dim sResult As String
    Dim lIndex As Long
    Dim sBuff As String
    Dim bTwelveHourClock As Boolean
    sResult = ""
    Set oResultsCollection = New Collection
    iPosition = 1
    Do While StrLen(sFormat) > 0
        sChar = StrMid(sFormat, iPosition, 1)
        iPosition = mp_FmtBreakDown(sChar, sFormat, oResultsCollection)
        sFormat = StrRight(sFormat, StrLen(sFormat) - iPosition)
        iPosition = 1
    Loop
    bTwelveHourClock = mp_FmtPreProcess(oResultsCollection)
    For lIndex = 1 To oResultsCollection.Count
        sBuff = oResultsCollection.Item(lIndex)
        sResult = sResult & mp_FmtProcess(sBuff, dtExpression, bTwelveHourClock)
    Next lIndex
    mp_FormatString = sResult
End Function

Private Function mp_FmtBreakDown(ByVal sChar As String, ByVal sFormat As String, ByRef oResultsCollection As Collection) As Long
    Select Case sChar
        Case ":"
            oResultsCollection.Add sChar
            mp_FmtBreakDown = 1
            Exit Function
        Case "/"
            oResultsCollection.Add sChar
            mp_FmtBreakDown = 1
            Exit Function
        Case "c"
            oResultsCollection.Add sChar
            mp_FmtBreakDown = 1
            Exit Function
        Case "d"
            mp_FmtBreakDown = mp_FmtRepeatingStrings(sChar, sFormat, 6, oResultsCollection)
            Exit Function
        Case "a"
            If mp_FmtForwardSearch(sFormat, "am/pm") = True Then
                oResultsCollection.Add "am/pm"
                mp_FmtBreakDown = 5
                Exit Function
            ElseIf mp_FmtForwardSearch(sFormat, "a/p") = True Then
                oResultsCollection.Add "a/p"
                mp_FmtBreakDown = 3
                Exit Function
            Else
                oResultsCollection.Add sChar
                mp_FmtBreakDown = 1
                Exit Function
            End If
        Case "A"
            If mp_FmtForwardSearch(sFormat, "AM/PM") = True Then
                oResultsCollection.Add "AM/PM"
                mp_FmtBreakDown = 5
                Exit Function
            ElseIf mp_FmtForwardSearch(sFormat, "A/P") = True Then
                oResultsCollection.Add "A/P"
                mp_FmtBreakDown = 3
                Exit Function
            ElseIf mp_FmtForwardSearch(sFormat, "AMPM") = True Then
                oResultsCollection.Add "AMPM"
                mp_FmtBreakDown = 4
                Exit Function
            Else
                oResultsCollection.Add sChar
                mp_FmtBreakDown = 1
                Exit Function
            End If
        Case "w"
            mp_FmtBreakDown = mp_FmtRepeatingStrings(sChar, sFormat, 2, oResultsCollection)
            Exit Function
        Case "m"
            mp_FmtBreakDown = mp_FmtRepeatingStrings(sChar, sFormat, 4, oResultsCollection)
            Exit Function
        Case "q"
            oResultsCollection.Add sChar
            mp_FmtBreakDown = 1
            Exit Function
        Case "y"
            mp_FmtBreakDown = mp_FmtRepeatingStrings(sChar, sFormat, 4, oResultsCollection)
            Exit Function
        Case "h"
            oResultsCollection.Add sChar
            mp_FmtBreakDown = 1
            Exit Function
        Case "H"
            If mp_FmtForwardSearch(sFormat, "Hh") = True Then
                oResultsCollection.Add "Hh"
                mp_FmtBreakDown = 2
                Exit Function
            Else
                oResultsCollection.Add sChar
                mp_FmtBreakDown = 1
                Exit Function
            End If
        Case "N"
            If mp_FmtForwardSearch(sFormat, "Nn") = True Then
                oResultsCollection.Add "Nn"
                mp_FmtBreakDown = 2
                Exit Function
            Else
                oResultsCollection.Add sChar
                mp_FmtBreakDown = 1
                Exit Function
            End If
        Case "S"
            If mp_FmtForwardSearch(sFormat, "Ss") = True Then
                oResultsCollection.Add "Ss"
                mp_FmtBreakDown = 2
                Exit Function
            Else
                oResultsCollection.Add sChar
                mp_FmtBreakDown = 1
                Exit Function
            End If
        Case "t"
            If mp_FmtForwardSearch(sFormat, "ttttt") = True Then
                oResultsCollection.Add "ttttt"
                mp_FmtBreakDown = 5
                Exit Function
            Else
                oResultsCollection.Add sChar
                mp_FmtBreakDown = 1
                Exit Function
            End If
        Case """"
            mp_FmtBreakDown = mp_FmtStringLiteral(sFormat, oResultsCollection)
            Exit Function
        Case "'"
            mp_FmtBreakDown = mp_FmtStringLiteral(sFormat, oResultsCollection)
            Exit Function
        Case Else
            oResultsCollection.Add sChar
            mp_FmtBreakDown = 1
            Exit Function
    End Select
End Function

Private Function mp_FmtRepeatingStrings(ByVal sChar As String, ByVal sFormat As String, ByVal lMax As Long, ByRef oResultsCollection As Collection) As Long
    Dim lPosition As Long
    Dim sCharBuff As String
    Dim sBuff As String
    lPosition = 1
    sBuff = sChar
    Do While ((lPosition) < StrLen(sFormat)) And ((lPosition + 1) <= lMax)
        sCharBuff = StrMid(sFormat, lPosition + 1, 1)
        If sCharBuff = sChar Then
            lPosition = lPosition + 1
            sBuff = sBuff & sChar
        Else
            Exit Do
        End If
    Loop
    oResultsCollection.Add sBuff
    mp_FmtRepeatingStrings = StrLen(sBuff)
End Function

Private Function mp_FmtForwardSearch(ByVal sFormat As String, ByVal SearchFor As String) As Boolean
    If StrLen(SearchFor) > StrLen(sFormat) Then
        mp_FmtForwardSearch = False
        Exit Function
    End If
    If StrLeft(sFormat, StrLen(SearchFor)) = SearchFor Then
        mp_FmtForwardSearch = True
    Else
        mp_FmtForwardSearch = False
    End If
End Function

Private Function mp_FmtStringLiteral(ByVal sFormat As String, ByRef oResultsCollection As Collection) As Long
    Dim lPosition As Long
    Dim sCharBuff As String
    Dim bFound As Boolean
    bFound = False
    lPosition = 2
    Do While (lPosition < StrLen(sFormat))
        sCharBuff = StrMid(sFormat, lPosition, 1)
        If sCharBuff = """" Or sCharBuff = "'" Then
            bFound = True
            Exit Do
        Else
            lPosition = lPosition + 1
        End If
    Loop
    If bFound = True Then
        oResultsCollection.Add "{strl}" & StrMid(sFormat, 2, lPosition - 2)
    Else
        oResultsCollection.Add "{strl}" & StrMid(sFormat, 2, lPosition - 1)
    End If
    mp_FmtStringLiteral = lPosition
End Function

Private Function mp_FmtPreProcess(ByRef oResultsCollection As Collection) As Boolean
    Dim lIndex As Long
    Dim sParam As String
    For lIndex = 1 To oResultsCollection.Count
        sParam = oResultsCollection.Item(lIndex)
        If (sParam = "AM/PM") Then
            mp_FmtPreProcess = True
            Exit Function
        ElseIf (sParam = "am/pm") Then
            mp_FmtPreProcess = True
            Exit Function
        ElseIf (sParam = "A/P") Then
            mp_FmtPreProcess = True
            Exit Function
        ElseIf (sParam = "a/p") Then
            mp_FmtPreProcess = True
            Exit Function
        ElseIf (sParam = "AMPM") Then
            mp_FmtPreProcess = True
            Exit Function
        End If
    Next lIndex
    mp_FmtPreProcess = False
    Exit Function
End Function

Private Function mp_FmtProcess(ByVal sParam As String, ByVal dtDate As Date, ByVal bTwelveHourClock As Boolean) As String
    Dim sReturn As String
    Dim lReturn As Long
    If (sParam = ":") Then
        mp_FmtProcess = GetLocaleInfoEx(LOCALE_STIME)
        Exit Function
    ElseIf (sParam = "/") Then
        mp_FmtProcess = GetLocaleInfoEx(LOCALE_SDATE)
        Exit Function
    ElseIf (sParam = "c") Then
        mp_FmtProcess = mp_FormatString(dtDate, "ddddd ttttt")
        Exit Function
    ElseIf (sParam = "d") Then
        sReturn = StrCStr(mp_oControl.MathLib.GetDay(dtDate))
        mp_FmtProcess = sReturn
        Exit Function
    ElseIf (sParam = "dd") Then
        sReturn = StrCStr(mp_oControl.MathLib.GetDay(dtDate))
        If (StrLen(sReturn) = 1) Then
            sReturn = "0" & sReturn
        End If
        mp_FmtProcess = sReturn
        Exit Function
    ElseIf (sParam = "ddd") Then
        lReturn = mp_oControl.MathLib.GetDayOfWeek(dtDate)
        Select Case lReturn
            Case 1
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVDAYNAME7)
                Exit Function
            Case 2
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVDAYNAME1)
                Exit Function
            Case 3
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVDAYNAME2)
                Exit Function
            Case 4
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVDAYNAME3)
                Exit Function
            Case 5
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVDAYNAME4)
                Exit Function
            Case 6
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVDAYNAME5)
                Exit Function
            Case 7
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVDAYNAME6)
                Exit Function
        End Select
    ElseIf (sParam = "dddd") Then
        lReturn = mp_oControl.MathLib.GetDayOfWeek(dtDate)
        Select Case lReturn
            Case 1
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SDAYNAME7)
                Exit Function
            Case 2
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SDAYNAME1)
                Exit Function
            Case 3
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SDAYNAME2)
                Exit Function
            Case 4
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SDAYNAME3)
                Exit Function
            Case 5
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SDAYNAME4)
                Exit Function
            Case 6
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SDAYNAME5)
                Exit Function
            Case 7
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SDAYNAME6)
                Exit Function
        End Select
    ElseIf (sParam = "ddddd") Then
        sReturn = GetLocaleInfoEx(LOCALE_SSHORTDATE)
        sReturn = StrReplace(sReturn, "M", "m")
        mp_FmtProcess = mp_FormatString(dtDate, sReturn)
        Exit Function
    ElseIf (sParam = "dddddd") Then
        sReturn = GetLocaleInfoEx(LOCALE_SLONGDATE)
        sReturn = StrReplace(sReturn, "M", "m")
        mp_FmtProcess = mp_FormatString(dtDate, sReturn)
        Exit Function
    ElseIf (sParam = "w") Then
        sReturn = StrCStr(mp_oControl.MathLib.GetDayOfWeek(dtDate))
        mp_FmtProcess = sReturn
        Exit Function
    ElseIf (sParam = "ww") Then
        sReturn = StrCStr(mp_oControl.MathLib.GetWeek(dtDate))
        mp_FmtProcess = sReturn
        Exit Function
    ElseIf (sParam = "m") Then
        sReturn = StrCStr(mp_oControl.MathLib.GetMonth(dtDate))
        mp_FmtProcess = sReturn
        Exit Function
    ElseIf (sParam = "mm") Then
        sReturn = StrCStr(mp_oControl.MathLib.GetMonth(dtDate))
        If (StrLen(sReturn) = 1) Then
            sReturn = "0" & sReturn
        End If
        mp_FmtProcess = sReturn
        Exit Function
    ElseIf (sParam = "mmm") Then
        lReturn = mp_oControl.MathLib.GetMonth(dtDate)
        Select Case lReturn
            Case 1
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME1)
                Exit Function
            Case 2
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME2)
                Exit Function
            Case 3
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME3)
                Exit Function
            Case 4
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME4)
                Exit Function
            Case 5
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME5)
                Exit Function
            Case 6
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME6)
                Exit Function
            Case 7
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME7)
                Exit Function
            Case 8
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME8)
                Exit Function
            Case 9
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME9)
                Exit Function
            Case 10
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME10)
                Exit Function
            Case 11
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME11)
                Exit Function
            Case 12
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SABBREVMONTHNAME12)
                Exit Function
        End Select
    ElseIf (sParam = "mmmm") Then
        lReturn = mp_oControl.MathLib.GetMonth(dtDate)
        Select Case lReturn
            Case 1
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SMONTHNAME1)
                Exit Function
            Case 2
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SMONTHNAME2)
                Exit Function
            Case 3
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SMONTHNAME3)
                Exit Function
            Case 4
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SMONTHNAME4)
                Exit Function
            Case 5
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SMONTHNAME5)
                Exit Function
            Case 6
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SMONTHNAME6)
                Exit Function
            Case 7
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SMONTHNAME7)
                Exit Function
            Case 8
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SMONTHNAME8)
                Exit Function
            Case 9
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SMONTHNAME9)
                Exit Function
            Case 10
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SMONTHNAME10)
                Exit Function
            Case 11
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SMONTHNAME11)
                Exit Function
            Case 12
                mp_FmtProcess = GetLocaleInfoEx(LOCALE_SMONTHNAME12)
                Exit Function
        End Select
    ElseIf (sParam = "q") Then
        sReturn = StrCStr(mp_oControl.MathLib.GetQuarter(dtDate))
        mp_FmtProcess = sReturn
        Exit Function
    ElseIf (sParam = "y") Then
        sReturn = StrCStr(mp_oControl.MathLib.GetDayOfYear(dtDate))
        mp_FmtProcess = sReturn
        Exit Function
    ElseIf (sParam = "yy") Then
        sReturn = StrCStr(mp_oControl.MathLib.GetYear(dtDate))
        sReturn = StrRight(sReturn, 2)
        mp_FmtProcess = sReturn
        Exit Function
    ElseIf (sParam = "yyyy") Then
        sReturn = StrCStr(mp_oControl.MathLib.GetYear(dtDate))
        mp_FmtProcess = sReturn
        Exit Function
    ElseIf (sParam = "h") Then
        lReturn = mp_oControl.MathLib.GetHour(dtDate)
        If (bTwelveHourClock = True) Then
            lReturn = lReturn
            If (lReturn > 12) Then
                lReturn = lReturn - 12
            End If
            If (lReturn = 0) Then
                lReturn = 12
            End If
        End If
        sReturn = StrCStr(lReturn)
        mp_FmtProcess = sReturn
        Exit Function
    ElseIf (sParam = "Hh") Then
        lReturn = mp_oControl.MathLib.GetHour(dtDate)
        If (bTwelveHourClock = True) Then
            lReturn = lReturn
            If (lReturn > 12) Then
                lReturn = lReturn - 12
            End If
            If (lReturn = 0) Then
                lReturn = 12
            End If
        End If
        sReturn = StrCStr(lReturn)
        If (StrLen(sReturn) = 1) Then
            sReturn = "0" & sReturn
        End If
        mp_FmtProcess = sReturn
        Exit Function
    ElseIf (sParam = "N") Then
        sReturn = StrCStr(mp_oControl.MathLib.GetMinute(dtDate))
        mp_FmtProcess = sReturn
        Exit Function
    ElseIf (sParam = "Nn") Then
        sReturn = StrCStr(mp_oControl.MathLib.GetMinute(dtDate))
        If (StrLen(sReturn) = 1) Then
            sReturn = "0" & sReturn
        End If
        mp_FmtProcess = sReturn
        Exit Function
    ElseIf (sParam = "S") Then
        sReturn = StrCStr(mp_oControl.MathLib.GetSecond(dtDate))
        mp_FmtProcess = sReturn
        Exit Function
    ElseIf (sParam = "Ss") Then
        sReturn = StrCStr(mp_oControl.MathLib.GetSecond(dtDate))
        If (StrLen(sReturn) = 1) Then
            sReturn = "0" & sReturn
        End If
        mp_FmtProcess = sReturn
        Exit Function
    ElseIf (sParam = "ttttt") Then
        sReturn = GetLocaleInfoEx(LOCALE_STIMEFORMAT)
        sReturn = StrReplace(sReturn, "mm", "Nn")
        sReturn = StrReplace(sReturn, "m", "N")
        sReturn = StrReplace(sReturn, "ss", "Ss")
        sReturn = StrReplace(sReturn, "s", "S")
        sReturn = StrReplace(sReturn, "SS", "Ss")
        sReturn = StrReplace(sReturn, "H", "h")
        sReturn = StrReplace(sReturn, "HH", "Hh")
        sReturn = StrReplace(sReturn, "hh", "Hh")
        sReturn = StrReplace(sReturn, "tt", "AMPM")
        mp_FmtProcess = mp_FormatString(dtDate, sReturn)
        Exit Function
    ElseIf (sParam = "AM/PM") Then
        If mp_oControl.MathLib.GetHour(dtDate) < 12 Then
            mp_FmtProcess = "AM"
            Exit Function
        Else
            mp_FmtProcess = "PM"
            Exit Function
        End If
    ElseIf (sParam = "am/pm") Then
        If mp_oControl.MathLib.GetHour(dtDate) < 12 Then
            mp_FmtProcess = "am"
            Exit Function
        Else
            mp_FmtProcess = "pm"
            Exit Function
        End If
    ElseIf (sParam = "A/P") Then
        If mp_oControl.MathLib.GetHour(dtDate) < 12 Then
            mp_FmtProcess = "A"
            Exit Function
        Else
            mp_FmtProcess = "P"
            Exit Function
        End If
    ElseIf (sParam = "a/p") Then
        If mp_oControl.MathLib.GetHour(dtDate) < 12 Then
            mp_FmtProcess = "a"
            Exit Function
        Else
            mp_FmtProcess = "p"
            Exit Function
        End If
    ElseIf (sParam = "AMPM") Then
        If mp_oControl.MathLib.GetHour(dtDate) < 12 Then
            mp_FmtProcess = GetLocaleInfoEx(LOCALE_S1159)
            Exit Function
        Else
            mp_FmtProcess = GetLocaleInfoEx(LOCALE_S2359)
            Exit Function
        End If
    ElseIf (StrLeft(sParam, 6) = "{strl}") Then
        sReturn = StrRight(sParam, StrLen(sParam) - 6)
        mp_FmtProcess = sReturn
        Exit Function
    End If
    sReturn = sParam
    mp_FmtProcess = sReturn
End Function

Private Function GetLocaleInfoEx(ByVal LCType As Long) As String
    Dim Symbol As String
    Dim iRet1 As Long
    Dim iRet2 As Long
    Dim lpLCDataVar As String
    Dim Pos As Integer
    Dim Locale As Long
    Locale = GetUserDefaultLCID()
    iRet1 = GetLocaleInfo(Locale, LCType, lpLCDataVar, 0)
    Symbol = String$(iRet1, 0)
    iRet2 = GetLocaleInfo(Locale, LCType, Symbol, iRet1)
    Pos = InStr(Symbol, Chr$(0))
    If Pos > 0 Then Symbol = StrLeft(Symbol, Pos - 1)
    GetLocaleInfoEx = Symbol
End Function



