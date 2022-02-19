VERSION 5.00
Begin VB.Form fPrintDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print"
   ClientHeight    =   4680
   ClientLeft      =   1890
   ClientTop       =   1935
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraOrientation 
      Caption         =   "Orientation"
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   2640
      Width           =   6375
      Begin VB.ComboBox cboOrientation 
         Height          =   315
         ItemData        =   "fPrintDialog.frx":0000
         Left            =   120
         List            =   "fPrintDialog.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame fraForm 
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   6375
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview..."
         Height          =   375
         Left            =   4560
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdSchedulePrinterSettings 
         Caption         =   "Schedule Printer Settings..."
         Height          =   375
         Left            =   2040
         TabIndex        =   19
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame fraPageRange 
      Caption         =   "Page Range"
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2895
      Begin VB.TextBox txtEndPage 
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtStartPage 
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblTo 
         Caption         =   "To:"
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblFrom 
         Caption         =   "From:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblNumberOfPages 
         Caption         =   "Total Pages:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame fraPage 
      Caption         =   "Page "
      Height          =   1455
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   3255
      Begin VB.TextBox txtScale 
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtPageHeight 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtPageWidth 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblPercentage 
         Caption         =   "%"
         Height          =   255
         Left            =   2400
         TabIndex        =   23
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblScale 
         Caption         =   "Scale:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblPixelsH 
         Caption         =   "Pixels"
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblPageHeight 
         Caption         =   "Height:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblPixelsW 
         Caption         =   "Pixels"
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblPageWidth 
         Caption         =   "Width:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraPrinter 
      Caption         =   "Printer"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.ComboBox cboPrinterName 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label lblPrinterName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "fPrintDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private mp_lColumns As Long
Private mp_lRows As Long
Private mp_lPageNumber As Long
Private mp_lXMargin As Long
Private mp_lYMargin As Long
Public mp_oControl As ActiveGanttVBCtl
Public m_lXPreviewMargin As Long
Public m_lYPreviewMargin As Long

Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long


Public Property Get StartPage() As Long
    If Trim$(txtStartPage.Text) = "" Then
        StartPage = 1
        Exit Property
    End If
    StartPage = txtStartPage.Text
End Property

Public Property Get EndPage() As Long
    If Trim$(txtEndPage.Text) = "" Then
        EndPage = 1
        Exit Property
    End If
    EndPage = txtEndPage.Text
End Property

Public Property Get PageLength() As Long
    If Trim$(txtPageWidth.Text) = "" Then
        PageLength = 0
        Exit Property
    End If
    PageLength = txtPageWidth.Text
End Property

Public Property Get PageHeight() As Long
    If Trim$(txtPageHeight.Text) = "" Then
        PageHeight = 0
        Exit Property
    End If
    PageHeight = txtPageHeight.Text
End Property

Public Property Get PagesInXDirection() As Long
    If PageLength = 0 Then
        PagesInXDirection = 0
        Exit Property
    End If
    PagesInXDirection = Abs(Int(-(mp_oControl.Printer.PrintAreaWidth / PageLength)))
End Property

Public Property Get PagesInYDirection() As Long
    If PageHeight = 0 Then
        PagesInYDirection = 0
        Exit Property
    End If
    PagesInYDirection = Abs(Int(-(mp_oControl.Printer.PrintAreaHeight / PageHeight)))
End Property

Public Property Get TotalPages() As Long
    TotalPages = PagesInXDirection * PagesInYDirection
End Property

Public Property Get ScheduleScale() As Long
    If Trim$(txtScale.Text) = "" Then
        ScheduleScale = 100
        Exit Property
    End If
    ScheduleScale = txtScale.Text
End Property

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mp_oControl.CurrentViewObject.ClientArea.FirstVisibleRow = 1
    If Printers.Count = 0 Then
        MsgBox "Before attempting to print a printer must be installed via the control panel", , "No printer installed"
        Exit Sub
    End If
    Dim X As Printer
    For Each X In Printers
       If X.DeviceName = cboPrinterName.Text Then
          Set Printer = X
          Exit For
       End If
    Next
    Dim i As Long
    For i = StartPage To EndPage
        Printer.Orientation = cboOrientation.ListIndex + 1
        Printer.Print
        PrintControl Printer.hdc, i, True
        Printer.NewPage
    Next i
    Printer.EndDoc
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    mp_oControl.CurrentViewObject.ClientArea.FirstVisibleRow = 1
    Set fPrintPreview.mp_Parent = Me
    fPrintPreview.Show 1, Me
End Sub

Private Sub cmdSchedulePrinterSettings_Click()
    Set fSchedulePrintSettings.mp_Parent = Me
    fSchedulePrintSettings.Show 1, Me
End Sub

Private Sub Form_Load()
    Dim X As Printer
    Dim i As Long
    If Printers.Count = 0 Then
        MsgBox "Before attempting to print a printer must be installed via the control panel", , "No printer installed"
        Exit Sub
    End If
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    For Each X In Printers
       cboPrinterName.AddItem X.DeviceName
       If Printer.DeviceName = X.DeviceName Then
        cboPrinterName.ListIndex = i
       End If
       i = i + 1
    Next
    cboOrientation.ListIndex = 0
    mp_lXMargin = 50
    mp_lYMargin = 50
    
    m_lXPreviewMargin = 100
    m_lYPreviewMargin = 100
    
    mp_oControl.Printer.Initialize mp_oControl.CurrentViewObject.TimeLine.StartDate, mp_oControl.CurrentViewObject.TimeLine.EndDate
    txtPageHeight.Text = 700
    txtPageWidth.Text = 500
    txtScale.Text = 100
    txtStartPage.Text = 1
    txtEndPage.Text = TotalPages
End Sub

Private Sub txtPageHeight_Change()
    lblNumberOfPages.Caption = "Total Pages: " & TotalPages
End Sub

Private Sub txtPageWidth_Change()
    lblNumberOfPages.Caption = "Total Pages: " & TotalPages
End Sub

Public Property Get XMargin() As Long
    XMargin = mp_lXMargin
End Property

Public Property Let XMargin(ByVal vNewValue As Long)
    mp_lXMargin = vNewValue
End Property

Public Property Get YMargin() As Long
    YMargin = mp_lYMargin
End Property

Public Property Let YMargin(ByVal vNewValue As Long)
    mp_lYMargin = vNewValue
End Property

Public Sub PrintControl(ByVal lHdc As Long, ByVal lPageNumber As Long, ByVal ToPrinter As Boolean, Optional ByVal lPhysicalOffsetX As Long = 0, Optional ByVal lPhysicalOffsetY As Long = 0, Optional ByVal lHScrollPos As Long = 0, Optional ByVal lVScrollPos As Long = 0)
    Dim lRow As Long
    Dim lColumn As Long
    Dim lPageLength As Long
    Dim lPageHeight As Long
    Dim iPixelsY As Long
    iPixelsY = GetDeviceCaps(lHdc, LOGPIXELSY)
    lRow = Abs(Int(-(lPageNumber / PagesInXDirection)))
    lColumn = lPageNumber - ((lRow - 1) * PagesInXDirection)
    If ((lColumn - 1) * PageLength) + PageLength > mp_oControl.Printer.PrintAreaWidth Then
        lPageLength = mp_oControl.Printer.PrintAreaWidth - ((lColumn - 1) * PageLength)
    Else
        lPageLength = PageLength
    End If
    If ((lRow - 1) * PageHeight) + PageHeight > mp_oControl.Printer.PrintAreaHeight Then
        lPageHeight = mp_oControl.Printer.PrintAreaHeight - ((lRow - 1) * PageHeight)
    Else
        lPageHeight = PageHeight
    End If
    Dim lXMargin As Long
    Dim lYMargin As Long
    If ToPrinter = True Then
        Dim iScreenPixelsY As Long
        iScreenPixelsY = GetDeviceCaps(Me.hdc, LOGPIXELSY)
        lXMargin = XMargin * (100 / ScheduleScale)
        lYMargin = YMargin * (100 / ScheduleScale)
        mp_oControl.Printer.PrintControl lHdc, (lColumn - 1) * PageLength, (lRow - 1) * PageHeight, lPageLength, lPageHeight, lXMargin, lYMargin, iPixelsY / iScreenPixelsY * ScheduleScale
    Else
        lXMargin = (XMargin + m_lXPreviewMargin + lPhysicalOffsetX - lHScrollPos) * 100 / ScheduleScale
        lYMargin = (YMargin + m_lYPreviewMargin + lPhysicalOffsetY - lVScrollPos) * 100 / ScheduleScale
        mp_oControl.Printer.PrintControl lHdc, (lColumn - 1) * PageLength, (lRow - 1) * PageHeight, lPageLength, lPageHeight, lXMargin, lYMargin, ScheduleScale
    End If
End Sub
