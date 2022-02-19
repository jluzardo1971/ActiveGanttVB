VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CCC1C7D1-F592-4261-9A57-CA48F869B175}#1.0#0"; "ActiveGanttVB2.ocx"
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
   Begin VB.Frame fraForm 
      Height          =   7635
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      Begin ActiveGanttVB.ActiveGanttVBCtl ActiveGanttVBCtl1 
         Height          =   7215
         Left            =   120
         TabIndex        =   1
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
'// The Source Code Store LLC
'// ACTIVEGANTT SCHEDULER COMPONENT FOR VISUAL BASIC 6
'// ACTIVEX COMPONENT
'// Copyright (c) 2002-2004 The Source Code Store LLC
'//
'// All Rights Reserved. No parts of this file may be reproduced or transmitted in any
'// form or by any means without the written permission of the author.
'// ----------------------------------------------------------------------------------------
Option Explicit

Private Function NewDate(ByVal Month As Long, ByVal Day As Long, ByVal Year As Long) As Date
    NewDate = DateSerial(Year, Month, Day)
End Function

Private Sub Form_Load()
    Dim lIndex As Integer
    Me.Caption = "The Source Code Store - ActiveGantt Scheduler Control Version " & ActiveGanttVBCtl1.Version & " - Project Management Example"
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2

    ActiveGanttVBCtl1.DefaultValues.RowHeight = 20
    ActiveGanttVBCtl1.TimeBlockBehaviour = TBB_CONTROLEXTENTS
    
    ActiveGanttVBCtl1.Styles.Add "Title"
    ActiveGanttVBCtl1.Styles.Item("Title").CaptionAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT
    ActiveGanttVBCtl1.Styles.Item("Title").CaptionXMargin = 5
    ActiveGanttVBCtl1.Styles.Item("Title").Font.Name = "Microsoft Sans Serif"
    ActiveGanttVBCtl1.Styles.Item("Title").Font.Size = 8
    ActiveGanttVBCtl1.Styles.Item("Title").Font.Bold = True
    
    ActiveGanttVBCtl1.Styles.Add "Task"
    ActiveGanttVBCtl1.Styles.Item("Task").CaptionAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT
    ActiveGanttVBCtl1.Styles.Item("Task").CaptionXMargin = 25
    ActiveGanttVBCtl1.Styles.Item("Task").Font.Name = "Microsoft Sans Serif"
    ActiveGanttVBCtl1.Styles.Item("Task").Font.Size = 7
    ActiveGanttVBCtl1.Styles.Item("Task").BorderStyle = E_STYLEBORDER.SBR_NONE
    ActiveGanttVBCtl1.Styles.Item("Task").Appearance = E_STYLEAPPEARANCE.SA_CELL
    ActiveGanttVBCtl1.Styles.Item("Task").BackColor = RGB(255, 255, 255)
    ActiveGanttVBCtl1.Styles.Item("Task").BorderColor = &H8000000F

    ActiveGanttVBCtl1.Styles.Add "Columns"
    ActiveGanttVBCtl1.Styles.Item("Columns").CaptionAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM
    ActiveGanttVBCtl1.Styles.Item("Columns").CaptionYMargin = 5
    ActiveGanttVBCtl1.Styles.Item("Columns").Font.Name = "Microsoft Sans Serif"
    ActiveGanttVBCtl1.Styles.Item("Columns").Font.Size = 8
    ActiveGanttVBCtl1.Styles.Item("Columns").Font.Bold = True
    ActiveGanttVBCtl1.Styles.Item("Columns").PictureAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM

    ActiveGanttVBCtl1.Styles.Add "TimeBlocks"
    ActiveGanttVBCtl1.Styles.Item("TimeBlocks").Appearance = E_STYLEAPPEARANCE.SA_FLAT
    ActiveGanttVBCtl1.Styles.Item("TimeBlocks").BackColor = RGB(195, 222, 210)
    ActiveGanttVBCtl1.Styles.Item("TimeBlocks").BackgroundPattern = GRE_BACKGROUNDPATTERN.FP_SOLID

    AddTaskStyle "Tasks1", 1, 2, 3
    AddTaskStyle "Tasks2", 4, 5, 6
    AddTaskStyle "Tasks3", 7, 8, 9
    AddTaskStyle "Tasks4", 10, 11, 12
    


    ActiveGanttVBCtl1.Columns.Add ""
    ActiveGanttVBCtl1.Columns.Item(1).Caption = "ID"
    ActiveGanttVBCtl1.Columns.Item(1).Width = 25
    ActiveGanttVBCtl1.Columns.Item(1).StyleIndex = "Columns"
    
    ActiveGanttVBCtl1.Columns.Add ""
    ActiveGanttVBCtl1.Columns.Item(2).Width = 30
    ActiveGanttVBCtl1.Columns.Item(2).StyleIndex = 3
    Set ActiveGanttVBCtl1.Columns.Item(2).Picture = imglstColumns.ListImages.Item(1).Picture
    
    ActiveGanttVBCtl1.Columns.Add "Task Name"
    ActiveGanttVBCtl1.Columns.Item(3).Width = 200
    ActiveGanttVBCtl1.Columns.Item(3).StyleIndex = "Columns"

    AddRowTitle "K1", "Legacy Systems Integration", True
    AddRowTitle "K2", "Implement CGS", False
    AddRowItem "K3", "Implement Customers/Contact/Leads"
    AddRowItem "K4", "Analyse Strategy for Completing CGS Implementation"
    AddRowItem "K5", "Implement Customer Financials"
    AddRowItem "K6", "Implement CRM"
    AddRowItem "K7", "Implement Financials"
    AddRowItem "K8", "Implement Customer Records"
    AddRowItem "K9", "Advanced CRM Pilot"
    AddRowTitle "K10", "Implement EDI", False
    AddRowItem "K11", "Electronic Documents"
    AddRowItem "K12", "Analyse Strategy for Completing EDI Implementation"
    AddRowItem "K13", "Implement Payroll/401 K"
    AddRowItem "K14", "Procurement"
    AddRowItem "K15", "Implement Administration"
    AddRowItem "K16", "RATTLE Tracking"
    AddRowItem "K17", "Implement Employee Health and Safety"
    AddRowItem "K18", "Implement Planning"
    AddRowItem "K19", "Implement Remaining Records"
    AddRowTitle "K20", "System Environment", False
    AddRowItem "K21", "Initial Management Support System"
    AddRowItem "K22", "Initial Self Service and Support System"
    AddRowItem "K23", "Complete Management Support System"
    AddRowItem "K24", "Complete Self Service and Support System"
    AddRowTitle "K25", "User Support and Documentation", True
    AddRowTitle "K26", "Systems Infrastructure", True

'    ActiveGanttVBCtl1.TimeBlocks.Add NewDate(1, 1, 2003), NewDate(4, 1, 2003), , "TimeBlocks"
'    ActiveGanttVBCtl1.TimeBlocks.Add NewDate(4, 15, 2003), NewDate(8, 1, 2003), , "TimeBlocks"
    ActiveGanttVBCtl1.Tasks.Add "", "K1", NewDate(1, 1, 2003), NewDate(4, 1, 2003), , "Tasks3"
    ActiveGanttVBCtl1.Tasks.Add "", "K1", NewDate(4, 15, 2003), NewDate(8, 1, 2003), , "Tasks4"
    ActiveGanttVBCtl1.Tasks.Add "", "K3", NewDate(1, 1, 2003), NewDate(2, 1, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K4", NewDate(1, 1, 2003), NewDate(2, 15, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K5", NewDate(2, 15, 2003), NewDate(7, 1, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K6", NewDate(2, 15, 2003), NewDate(7, 1, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K7", NewDate(2, 15, 2003), NewDate(7, 1, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K8", NewDate(2, 15, 2003), NewDate(7, 1, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K9", NewDate(5, 1, 2003), NewDate(7, 1, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K11", NewDate(2, 1, 2003), NewDate(3, 11, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K12", NewDate(2, 15, 2003), NewDate(3, 20, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K13", NewDate(3, 20, 2003), NewDate(7, 15, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K14", NewDate(3, 20, 2003), NewDate(7, 5, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K15", NewDate(3, 20, 2003), NewDate(6, 25, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K16", NewDate(3, 20, 2003), NewDate(5, 1, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K17", NewDate(3, 20, 2003), NewDate(5, 1, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K18", NewDate(3, 20, 2003), NewDate(5, 1, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K19", NewDate(3, 20, 2003), NewDate(5, 1, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K16", NewDate(5, 7, 2003), NewDate(7, 15, 2003), , "Tasks2"
    ActiveGanttVBCtl1.Tasks.Add "", "K17", NewDate(5, 8, 2003), NewDate(7, 15, 2003), , "Tasks2"
    ActiveGanttVBCtl1.Tasks.Add "", "K18", NewDate(5, 10, 2003), NewDate(7, 15, 2003), , "Tasks2"
    ActiveGanttVBCtl1.Tasks.Add "", "K19", NewDate(5, 15, 2003), NewDate(7, 15, 2003), , "Tasks2"
    ActiveGanttVBCtl1.Tasks.Add "", "K21", NewDate(12, 1, 2002), NewDate(4, 1, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K22", NewDate(12, 1, 2002), NewDate(3, 20, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K23", NewDate(5, 1, 2003), NewDate(8, 1, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K24", NewDate(6, 1, 2003), NewDate(8, 1, 2003), , "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K25", NewDate(1, 1, 2003), NewDate(8, 1, 2003), , "Tasks3"
    ActiveGanttVBCtl1.Tasks.Add "", "K26", NewDate(1, 1, 2003), NewDate(7, 5, 2003), , "Tasks3"


    Dim oTickMarks As clsTickMarks
    ActiveGanttVBCtl1.Views.Add "12h", "1d", ST_QUARTER, ST_WEEK, ST_MONTH
    ActiveGanttVBCtl1.Views.Item("1").TimeLine.TierArea.TierFormat.QuarterIntervalFormat = "q""Q"" yyyy"
    ActiveGanttVBCtl1.Views.Item("1").TimeLine.TierArea.TierFormat.MonthIntervalFormat = "mmm"
    Set oTickMarks = ActiveGanttVBCtl1.Views.Item("1").TimeLine.TickMarkArea.TickMarks
    oTickMarks.Add 1, TLT_BIG, True, "d", False
    oTickMarks.Add 10, TLT_BIG, True, "d", False
    oTickMarks.Add 20, TLT_BIG, True, "d", False
    oTickMarks.Add 5, TLT_SMALL, False, "", False
    oTickMarks.Add 15, TLT_SMALL, False, "", False
    oTickMarks.Add 25, TLT_SMALL, False, "", False
    ActiveGanttVBCtl1.Splitter.Position = 255
    ActiveGanttVBCtl1.CurrentView = "1"
    ActiveGanttVBCtl1.CurrentViewObject.TimeLine.Position (NewDate(12, 1, 2002))
    
    ActiveGanttVBCtl1.CurrentViewObject.TimeLine.TierArea.MiddleTier.Visible = True
    ActiveGanttVBCtl1.CurrentViewObject.ClientArea.Grid.Interval = "1m"
    ActiveGanttVBCtl1.CurrentViewObject.ClientArea.Grid.VerticalLines = True
    
    ActiveGanttVBCtl1.CurrentViewObject.TimeLine.ProgressLine.LineType = TLMT_USER
    ActiveGanttVBCtl1.CurrentViewObject.TimeLine.ProgressLine.Position = NewDate(3, 13, 2003)
    ActiveGanttVBCtl1.CurrentViewObject.TimeLine.ProgressLine.Length = TLMA_BOTH
    
    
    ActiveGanttVBCtl1.CurrentViewObject.TimeLine.ScrollBar.StartDate = NewDate(12, 1, 2002)
    ActiveGanttVBCtl1.CurrentViewObject.TimeLine.ScrollBar.Interval = "1d"
    ActiveGanttVBCtl1.CurrentViewObject.TimeLine.ScrollBar.SmallChange = 1
    ActiveGanttVBCtl1.CurrentViewObject.TimeLine.ScrollBar.LargeChange = 10
    ActiveGanttVBCtl1.CurrentViewObject.TimeLine.ScrollBar.Max = 100
    ActiveGanttVBCtl1.CurrentViewObject.TimeLine.ScrollBar.Enabled = True
    
    
    

    ActiveGanttVBCtl1.WriteXML "c:\test.xml"
    ActiveGanttVBCtl1.Redraw
End Sub

Private Sub AddRowTitle(ByVal Key As String, ByVal Caption As String, ByVal Container As Boolean)
    ActiveGanttVBCtl1.Rows.Add Key, Key
    ActiveGanttVBCtl1.Rows.Item(Key).Cells.Item(3).Caption = Caption
    ActiveGanttVBCtl1.Rows.Item(Key).Cells.Item(3).StyleIndex = "Title"
    If Container = False Then
        ActiveGanttVBCtl1.Rows.Item(Key).Container = False
        ActiveGanttVBCtl1.Rows.Item(Key).ClientAreaStyleIndex = "0"
    Else
        ActiveGanttVBCtl1.Rows.Item(Key).Container = True
    End If
    ActiveGanttVBCtl1.Rows.Item(Key).Cells.Item(1).Caption = Replace(Key, "K", "")
End Sub

Private Sub AddRowItem(ByVal Key As String, ByVal Caption As String)
    ActiveGanttVBCtl1.Rows.Add Key, Key
    ActiveGanttVBCtl1.Rows.Item(Key).Cells.Item(3).Caption = Caption
    ActiveGanttVBCtl1.Rows.Item(Key).Cells.Item(3).StyleIndex = "Task"
    ActiveGanttVBCtl1.Rows.Item(Key).Cells.Item(1).Caption = Replace(Key, "K", "")
End Sub

Private Sub AddTaskStyle(ByVal Key As String, StartPicture As Long, MiddlePicture As Long, EndPicture As Long)
    ActiveGanttVBCtl1.Styles.Add Key
    ActiveGanttVBCtl1.Styles.Item(Key).Appearance = E_STYLEAPPEARANCE.SA_GRAPHICAL
    Set ActiveGanttVBCtl1.Styles.Item(Key).TaskStyle.StartPicture = ImageList1.ListImages.Item(StartPicture).Picture
    Set ActiveGanttVBCtl1.Styles.Item(Key).TaskStyle.MiddlePicture = ImageList1.ListImages.Item(MiddlePicture).Picture
    Set ActiveGanttVBCtl1.Styles.Item(Key).TaskStyle.EndPicture = ImageList1.ListImages.Item(EndPicture).Picture
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
