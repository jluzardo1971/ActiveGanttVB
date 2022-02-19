VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{404F18CB-304C-4658-92D0-119074ED7C75}#1.0#0"; "ActiveGanttVB2.ocx"
Begin VB.Form fProjectManagement01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Source Code Store - ActiveGantt Scheduler Control - Project Management Example"
   ClientHeight    =   7965
   ClientLeft      =   2355
   ClientTop       =   1275
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame fraForm 
      Height          =   7635
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      Begin ActiveGanttVB.ActiveGanttVBCtl ActiveGanttVBCtl1 
         Height          =   7215
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   12726
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList imglstColumns 
         Left            =   840
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   27
         ImageHeight     =   20
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement01.frx":0000
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   17
         ImageHeight     =   17
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement01.frx":06E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement01.frx":0AAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement01.frx":0E74
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement01.frx":123C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement01.frx":1604
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement01.frx":19CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement01.frx":1D94
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement01.frx":215C
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement01.frx":2524
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement01.frx":28EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement01.frx":2CB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement01.frx":307C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuPrint 
         Caption         =   "Print..."
      End
      Begin VB.Menu mnuLine010 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "fProjectManagement01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ----------------------------------------------------------------------------------------
'//                              COPYRIGHT NOTICE
'// ----------------------------------------------------------------------------------------
'//
'// The Source Code Store ACTIVEGANTT COMPONENT FOR VISUAL BASIC 6.0. VERSION 2
'// Copyright (c) 2002-2004, Julio Luzardo.
'//
'// All Rights Reserved. No parts of this file may be reproduced or transmitted in any
'// form or by any means without the written permission of the author.
'// ----------------------------------------------------------------------------------------

Option Explicit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private mp_lRowIndex As Long
Private mp_lCellIndex As Long
Private mp_bDoNotAllowTextEdit As Boolean

Private Function NewDate(ByVal Month As Long, ByVal Day As Long, ByVal Year As Long) As Date
    NewDate = DateSerial(Year, Month, Day)
End Function

Private Sub ActiveGanttVBCtl1_ControlRedrawn()
    txtEdit.Visible = False
End Sub

Private Sub ActiveGanttVBCtl1_CellClick(ByVal RowIndex As Long, ByVal CellIndex As Long, ByVal X As Single, ByVal Y As Single, ByVal Button As Integer)
    If mp_bDoNotAllowTextEdit = True Then
        mp_bDoNotAllowTextEdit = False
        Exit Sub
    End If
    If RowIndex = mp_lRowIndex And CellIndex = mp_lCellIndex Then
        txtEdit.Left = (ActiveGanttVBCtl1.Rows.Item(RowIndex).Cells.Item(3).LeftTrim * Screen.TwipsPerPixelX)
        txtEdit.Top = (ActiveGanttVBCtl1.Rows.Item(RowIndex).Cells.Item(3).Top * Screen.TwipsPerPixelY)
        txtEdit.Height = ((ActiveGanttVBCtl1.Rows.Item(RowIndex).Cells.Item(3).Bottom - ActiveGanttVBCtl1.Rows.Item(RowIndex).Cells.Item(3).Top) * Screen.TwipsPerPixelX)
        txtEdit.Width = ((ActiveGanttVBCtl1.Rows.Item(RowIndex).Cells.Item(3).RightTrim - ActiveGanttVBCtl1.Rows.Item(RowIndex).Cells.Item(3).LeftTrim) * Screen.TwipsPerPixelY)
        txtEdit.Text = ActiveGanttVBCtl1.Rows.Item(RowIndex).Cells.Item(3).Caption
        txtEdit.BackColor = ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Rows.Item(RowIndex).Cells.Item(3).StyleIndex).BackColor
        Set txtEdit.Font = ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Rows.Item(RowIndex).Cells.Item(3).StyleIndex).Font
        txtEdit.Visible = True
        txtEdit.SetFocus
        txtEdit.SelStart = 0
        txtEdit.SelLength = Len(txtEdit.Text)
    End If
End Sub

Private Sub ActiveGanttVBCtl1_CellSelected(ByVal RowIndex As Long, ByVal CellIndex As Long)
    mp_bDoNotAllowTextEdit = True
    If CellIndex = 3 Then
        mp_lRowIndex = RowIndex
        mp_lCellIndex = CellIndex
    Else
        mp_lRowIndex = 0
        mp_lCellIndex = 0
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = "The Source Code Store - ActiveGantt Scheduler Control Version " & ActiveGanttVBCtl1.Version & " - Project Management Example"
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
       Dim lIndex As Integer
        With ActiveGanttVBCtl1
            .AutomaticRedraw = False
            .DefaultValues.RowHeight = 20
            .TimeBlockBehaviour = TBB_CONTROLEXTENTS
            With .Styles
                .Add "Title"
                .Item("Title").CaptionAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT
                .Item("Title").CaptionXMargin = 5
                .Item("Title").Font.Name = "Microsoft Sans Serif"
                .Item("Title").Font.Size = 8
                .Item("Title").Font.Bold = True
                .Add "Task"
                .Item("Task").CaptionAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT
                .Item("Task").CaptionXMargin = 25
                .Item("Task").Font.Name = "Microsoft Sans Serif"
                .Item("Task").Font.Size = 7
                .Item("Task").BorderStyle = E_STYLEBORDER.SBR_NONE
                .Item("Task").Appearance = E_STYLEAPPEARANCE.SA_CELL
                .Item("Task").BackColor = RGB(255, 255, 255)
                .Item("Task").BorderColor = &H8000000F
                .Add "Columns"
                .Item("Columns").CaptionAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM
                .Item("Columns").CaptionYMargin = 5
                .Item("Columns").Font.Name = "Microsoft Sans Serif"
                .Item("Columns").Font.Size = 8
                .Item("Columns").Font.Bold = True
                Set .Item("Columns").ImageList = imglstColumns
                .Item("Columns").PictureAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM
                .Add "Tasks1"
                .Item("Tasks1").Appearance = E_STYLEAPPEARANCE.SA_GRAPHICAL
                .Item("Tasks1").StartPictureIndex = 1
                .Item("Tasks1").MiddlePictureIndex = 2
                .Item("Tasks1").EndPictureIndex = 3
                Set .Item("Tasks1").ImageList = ImageList1
                .Add "Tasks2"
                .Item("Tasks2").Appearance = E_STYLEAPPEARANCE.SA_GRAPHICAL
                .Item("Tasks2").StartPictureIndex = 4
                .Item("Tasks2").MiddlePictureIndex = 5
                .Item("Tasks2").EndPictureIndex = 6
                Set .Item("Tasks2").ImageList = ImageList1
                .Add "Tasks3"
                .Item("Tasks3").Appearance = E_STYLEAPPEARANCE.SA_GRAPHICAL
                .Item("Tasks3").StartPictureIndex = 7
                .Item("Tasks3").MiddlePictureIndex = 8
                .Item("Tasks3").EndPictureIndex = 9
                Set .Item("Tasks3").ImageList = ImageList1
                .Add "Tasks4"
                .Item("Tasks4").Appearance = E_STYLEAPPEARANCE.SA_GRAPHICAL
                .Item("Tasks4").StartPictureIndex = 10
                .Item("Tasks4").MiddlePictureIndex = 11
                .Item("Tasks4").EndPictureIndex = 12
                Set .Item("Tasks4").ImageList = ImageList1
                .Add "TimeBlocks"
                .Item("TimeBlocks").Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .Item("TimeBlocks").BackColor = RGB(195, 222, 210)
                .Item("TimeBlocks").BackgroundPattern = GRE_BACKGROUNDPATTERN.FP_SOLID
            End With
            With .Columns
                .Add ""
                .Item(1).Caption = "ID"
                .Item(1).Width = 25
                .Item(1).StyleIndex = "Columns"
                .Add ""
                .Item(2).Width = 30
                .Item(2).StyleIndex = 3
                .Item(2).PictureIndex = 1
                .Add "Task Name"
                .Item(3).Width = 200
                .Item(3).StyleIndex = "Columns"
            End With
            With .Rows
                .Add "K1", "1"
                .Item("K1").Cells.Item(3).Caption = "Legacy Systems Integration"
                .Item("K1").Cells.Item(3).StyleIndex = "Title"
                .Add "K2", "2"
                .Item("K2").Cells.Item(3).Caption = "Implement CGS"
                .Item("K2").Cells.Item(3).StyleIndex = "Title"
                .Item("K2").Container = False
                .Add "K3", "3"
                .Item("K3").Cells.Item(3).Caption = "Implement Customers/Contact/Leads"
                .Add "K4", "4"
                .Item("K4").Cells.Item(3).Caption = "Analyse Strategy for Completing CGS Implementation"
                .Add "K5", "5"
                .Item("K5").Cells.Item(3).Caption = "Implement Customer Financials"
                .Add "K6", "6"
                .Item("K6").Cells.Item(3).Caption = "Implement CRM"
                .Add "K7", "7"
                .Item("K7").Cells.Item(3).Caption = "Implement Financials"
                .Add "K8", "8"
                .Item("K8").Cells.Item(3).Caption = "Implement Customer Records"
                .Add "K9", "9"
                .Item("K9").Cells.Item(3).Caption = "Advanced CRM Pilot"
                .Add "K10", "10"
                .Item("K10").Cells.Item(3).Caption = "Implement EDI"
                .Item("K10").Cells.Item(3).StyleIndex = "Title"
                .Item("K10").Container = False
                .Add "K11", "11"
                .Item("K11").Cells.Item(3).Caption = "Electronic Documents"
                .Add "K12", "12"
                .Item("K12").Cells.Item(3).Caption = "Analyse Strategy for Completing EDI Implementation"
                .Add "K13", "13"
                .Item("K13").Cells.Item(3).Caption = "Implement Payroll/401 K"
                .Add "K14", "14"
                .Item("K14").Cells.Item(3).Caption = "Procurement"
                .Add "K15", "15"
                .Item("K15").Cells.Item(3).Caption = "Implement Administration"
                .Add "K16", "16"
                .Item("K16").Cells.Item(3).Caption = "RATTLE Tracking"
                .Add "K17", "17"
                .Item("K17").Cells.Item(3).Caption = "Implement Employee Health and Safety"
                .Add "K18", "18"
                .Item("K18").Cells.Item(3).Caption = "Implement Planning"
                .Add "K19", "19"
                .Item("K19").Cells.Item(3).Caption = "Implement Remaining Records"
                .Add "K20", "20"
                .Item("K20").Cells.Item(3).Caption = "System Environment"
                .Item("K20").Cells.Item(3).StyleIndex = "Title"
                .Item("K20").Container = False
                .Add "K21", "21"
                .Item("K21").Cells.Item(3).Caption = "Initial Management Support System"
                .Add "K22", "22"
                .Item("K22").Cells.Item(3).Caption = "Initial Self Service and Support System"
                .Add "K23", "23"
                .Item("K23").Cells.Item(3).Caption = "Complete Management Support System"
                .Add "K24", "24"
                .Item("K24").Cells.Item(3).Caption = "Complete Self Service and Support System"
                .Add "K25", "25"
                .Item("K25").Cells.Item(3).Caption = "User Support and Documentation"
                .Item("K25").Cells.Item(3).StyleIndex = "Title"
                .Add "K26", "26"
                .Item("K26").Cells.Item(3).Caption = "Systems Infrastructure"
                .Item("K26").Cells.Item(3).StyleIndex = "Title"
                For lIndex = 1 To .Count
                    .Item("K" & lIndex).Cells.Item(1).Caption = lIndex
                    If .Item("K" & lIndex).Cells.Item(3).StyleIndex = "0" Then
                        .Item("K" & lIndex).Cells.Item(3).StyleIndex = "Task"
                    End If
                Next lIndex
            End With
            With .TimeBlocks
                .Add NewDate(1, 1, 2003), NewDate(4, 1, 2003), , "TimeBlocks"
                .Add NewDate(4, 15, 2003), NewDate(8, 1, 2003), , "TimeBlocks"
            End With
            With .Tasks
                .Add "", "K1", NewDate(1, 1, 2003), NewDate(4, 1, 2003), , "Tasks3"
                .Add "", "K1", NewDate(4, 15, 2003), NewDate(8, 1, 2003), , "Tasks4"
                .Add "", "K3", NewDate(1, 1, 2003), NewDate(2, 1, 2003), , "Tasks1"
                .Add "", "K4", NewDate(1, 1, 2003), NewDate(2, 15, 2003), , "Tasks1"
                .Add "", "K5", NewDate(2, 15, 2003), NewDate(7, 1, 2003), , "Tasks1"
                .Add "", "K6", NewDate(2, 15, 2003), NewDate(7, 1, 2003), , "Tasks1"
                .Add "", "K7", NewDate(2, 15, 2003), NewDate(7, 1, 2003), , "Tasks1"
                .Add "", "K8", NewDate(2, 15, 2003), NewDate(7, 1, 2003), , "Tasks1"
                .Add "", "K9", NewDate(5, 1, 2003), NewDate(7, 1, 2003), , "Tasks1"
                .Add "", "K11", NewDate(2, 1, 2003), NewDate(3, 11, 2003), , "Tasks1"
                .Add "", "K12", NewDate(2, 15, 2003), NewDate(3, 20, 2003), , "Tasks1"
                .Add "", "K13", NewDate(3, 20, 2003), NewDate(7, 15, 2003), , "Tasks1"
                .Add "", "K14", NewDate(3, 20, 2003), NewDate(7, 5, 2003), , "Tasks1"
                .Add "", "K15", NewDate(3, 20, 2003), NewDate(6, 25, 2003), , "Tasks1"
                .Add "", "K16", NewDate(3, 20, 2003), NewDate(5, 1, 2003), , "Tasks1"
                .Add "", "K17", NewDate(3, 20, 2003), NewDate(5, 1, 2003), , "Tasks1"
                .Add "", "K18", NewDate(3, 20, 2003), NewDate(5, 1, 2003), , "Tasks1"
                .Add "", "K19", NewDate(3, 20, 2003), NewDate(5, 1, 2003), , "Tasks1"
                .Add "", "K16", NewDate(5, 7, 2003), NewDate(7, 15, 2003), , "Tasks2"
                .Add "", "K17", NewDate(5, 8, 2003), NewDate(7, 15, 2003), , "Tasks2"
                .Add "", "K18", NewDate(5, 10, 2003), NewDate(7, 15, 2003), , "Tasks2"
                .Add "", "K19", NewDate(5, 15, 2003), NewDate(7, 15, 2003), , "Tasks2"
                .Add "", "K21", NewDate(12, 1, 2002), NewDate(4, 1, 2003), , "Tasks1"
                .Add "", "K22", NewDate(12, 1, 2002), NewDate(3, 20, 2003), , "Tasks1"
                .Add "", "K23", NewDate(5, 1, 2003), NewDate(8, 1, 2003), , "Tasks1"
                .Add "", "K24", NewDate(6, 1, 2003), NewDate(8, 1, 2003), , "Tasks1"
                .Add "", "K25", NewDate(1, 1, 2003), NewDate(8, 1, 2003), , "Tasks3"
                .Add "", "K26", NewDate(1, 1, 2003), NewDate(7, 5, 2003), , "Tasks3"
            End With
            Dim oTickMarks As clsTickMarks
            .Views.Add "12h", "1d", ST_QUARTER, ST_WEEK, ST_MONTH
            .Views.Item("1").TierFormat.QuarterIntervalFormat = "q""Q"" yyyy"
            .Views.Item("1").TierFormat.MonthIntervalFormat = "mmm"
            Set oTickMarks = .Views.Item("1").TickMarks
            oTickMarks.Add 1, TLT_BIG, True, "d", False
            oTickMarks.Add 10, TLT_BIG, True, "d", False
            oTickMarks.Add 20, TLT_BIG, True, "d", False
            oTickMarks.Add 5, TLT_SMALL, False, "", False
            oTickMarks.Add 15, TLT_SMALL, False, "", False
            oTickMarks.Add 25, TLT_SMALL, False, "", False
            .SplitterPosition = 255
            .AutomaticRedraw = True
            .CurrentView = "1"
            .CurrentViewObject.TimeLine.Position (NewDate(12, 1, 2002))   '// This method redraws the screen
            SetParent txtEdit.hwnd, ActiveGanttVBCtl1.hwnd
            .Redraw
        End With
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuPrint_Click()
    Set fPrintDialog.mp_oControl = ActiveGanttVBCtl1
    fPrintDialog.Show 1, Me
    ActiveGanttVBCtl1.Printer.Terminate
    ActiveGanttVBCtl1.Redraw
End Sub

Private Sub txtEdit_LostFocus()
    ActiveGanttVBCtl1.Rows.Item(mp_lRowIndex).Cells.Item(3).Caption = txtEdit.Text
    ActiveGanttVBCtl1.Redraw
    txtEdit.Visible = False
End Sub










































































































