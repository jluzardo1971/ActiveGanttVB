VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrinter"
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
Private mp_dtPrintAreaStartDate As Date
Private mp_dtPrintAreaEndDate As Date
Private mp_dtPrintStartDateBuffer As Date
Private mp_oView As clsView


'// ---------------------------------------------------------------------------------------------------------------------
'// Initialization & Destruction
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub InitializeClass(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
End Sub

Private Sub Class_Terminate()
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Object Member Variables
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub Initialize(ByVal StartDate As Date, ByVal EndDate As Date, Optional ByVal ControlHeight As Long = -1)
    Const CorrectionFactor = 5
    
    Set mp_oView = mp_oControl.CurrentViewObject
    mp_dtPrintAreaStartDate = StartDate
    mp_dtPrintAreaEndDate = EndDate
    mp_oControl.GrphLib.CustomWidth = mp_oControl.MathLib.DateTimeDiff(mp_oView.f_Interval, StartDate, EndDate) / mp_oView.f_Factor + CorrectionFactor
    mp_oControl.GrphLib.CustomWidth = mp_oControl.GrphLib.CustomWidth + mp_oControl.Splitter.Right
    If ControlHeight = -1 Then
        mp_oControl.GrphLib.CustomHeight = mp_oControl.Rows.Height + (mp_oControl.Rows.Count * 1) + mp_oControl.CurrentViewObject.ClientArea.Top + mp_oControl.mt_BorderThickness
        If mp_oControl.GrphLib.CustomHeight < mp_oControl.f_Height Then
            mp_oControl.GrphLib.CustomHeight = mp_oControl.f_Height
        End If
    Else
        mp_oControl.GrphLib.CustomHeight = mp_oControl.f_Height
    End If
    If mp_oControl.f_oHScrollBar2.Enabled = False Then
        mp_dtPrintStartDateBuffer = mp_oView.TimeLine.StartDate
        mp_oView.TimeLine.f_StartDate = mp_dtPrintAreaStartDate
    Else
        mp_dtPrintStartDateBuffer = mp_oView.TimeLine.ScrollBar.StartDate
        mp_oView.TimeLine.ScrollBar.StartDate = mp_dtPrintAreaStartDate
    End If
End Sub

Public Sub Terminate()
    If mp_oControl.f_oHScrollBar2.Enabled = False Then
        mp_oView.TimeLine.f_StartDate = mp_dtPrintStartDateBuffer
    Else
        mp_oView.TimeLine.ScrollBar.StartDate = mp_dtPrintStartDateBuffer
    End If
End Sub

Public Sub PrintControl(ByVal DestHdc As Long, ByVal XOrigin As Long, ByVal YOrigin As Long, ByVal XOriginExtents As Long, ByVal YOriginExtents As Long, ByVal MarginX As Long, ByVal MarginY As Long, ByVal DestScale As Long)
    mp_oControl.GrphLib.CustomPrinting = True
    mp_oControl.GrphLib.StartPrintControl DestHdc, XOrigin, YOrigin, XOriginExtents, YOriginExtents, MarginX, MarginY, DestScale
    mp_oControl.mp_PositionScrollBars
    mp_oControl.f_Draw
    mp_oControl.GrphLib.CustomPrinting = False
    mp_oControl.mp_PositionScrollBars
    mp_oControl.GrphLib.EndPrintControl
End Sub

Public Property Get PrintAreaEndDate() As Date
    PrintAreaEndDate = mp_dtPrintAreaEndDate
End Property

Public Property Get PrintAreaHeight() As Long
    PrintAreaHeight = mp_oControl.GrphLib.CustomHeight
End Property

Public Property Get PrintAreaStartDate() As Date
    PrintAreaStartDate = mp_dtPrintAreaStartDate
End Property

Public Property Get PrintAreaWidth() As Long
    PrintAreaWidth = mp_oControl.GrphLib.CustomWidth
End Property
