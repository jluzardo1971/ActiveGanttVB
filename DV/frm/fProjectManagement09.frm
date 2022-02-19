VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CCC1C7D1-F592-4261-9A57-CA48F869B175}#1.0#0"; "ActiveGanttVB2.ocx"
Begin VB.Form fProjectManagement09 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7635
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
   End
   Begin MSComctlLib.ImageList imglstColumns 
      Left            =   12000
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
            Picture         =   "fProjectManagement09.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fProjectManagement09"
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
    Me.Caption = "The Source Code Store - ActiveGantt Scheduler Control Version " & ActiveGanttVBCtl1.Version & " - Project Management Example"
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Dim lIndex As Long
    ActiveGanttVBCtl1.DefaultValues.RowHeight = 20
    ActiveGanttVBCtl1.AllowRowSize = False
    ActiveGanttVBCtl1.AddMode = AT_TASKADD
    ActiveGanttVBCtl1.Styles.Add "Title"
    ActiveGanttVBCtl1.Styles.Item("Title").Appearance = E_STYLEAPPEARANCE.SA_RAISED
    ActiveGanttVBCtl1.Styles.Item("Title").CaptionAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT
    ActiveGanttVBCtl1.Styles.Item("Title").CaptionXMargin = 5
    ActiveGanttVBCtl1.Styles.Item("Title").Font.Bold = True
    ActiveGanttVBCtl1.Styles.Add "TimeLineTiers"
    ActiveGanttVBCtl1.Styles.Item("TimeLineTiers").Font.Size = 7
    ActiveGanttVBCtl1.Styles.Item("TimeLineTiers").Font.Bold = True
    ActiveGanttVBCtl1.Styles.Add "Task"
    ActiveGanttVBCtl1.Styles.Item("Task").CaptionAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT
    ActiveGanttVBCtl1.Styles.Item("Task").CaptionXMargin = 25
    ActiveGanttVBCtl1.Styles.Item("Task").Font.Name = "Microsoft Sans Serif"
    ActiveGanttVBCtl1.Styles.Item("Task").Font.Size = 7
    ActiveGanttVBCtl1.Styles.Item("Task").BorderStyle = E_STYLEBORDER.SBR_NONE
    ActiveGanttVBCtl1.Styles.Item("Task").Appearance = E_STYLEAPPEARANCE.SA_CELL
    ActiveGanttVBCtl1.Styles.Item("Task").BackColor = &H80000005
    'ActiveGanttVBCtl1.Styles.Item("Task").BorderColor = ActiveGanttVBCtl1.GridLinesColor
    ActiveGanttVBCtl1.Styles.Add "Columns"
    ActiveGanttVBCtl1.Styles.Item("Columns").Appearance = E_STYLEAPPEARANCE.SA_RAISED
    ActiveGanttVBCtl1.Styles.Item("Columns").CaptionAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM
    ActiveGanttVBCtl1.Styles.Item("Columns").CaptionYMargin = 5
    ActiveGanttVBCtl1.Styles.Item("Columns").Font.Bold = True
    ActiveGanttVBCtl1.Styles.Item("Columns").PictureAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM
    ActiveGanttVBCtl1.Styles.Add "Cells"
    ActiveGanttVBCtl1.Styles.Item("Cells").Appearance = E_STYLEAPPEARANCE.SA_RAISED
    ActiveGanttVBCtl1.Styles.Add "Milestones1"
    
    ActiveGanttVBCtl1.Styles.Item("Milestones1").Appearance = E_STYLEAPPEARANCE.SA_FLAT
    ActiveGanttVBCtl1.Styles.Item("Milestones1").MilestoneStyle.ShapeIndex = 3
    ActiveGanttVBCtl1.Styles.Item("Milestones1").Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
    ActiveGanttVBCtl1.Styles.Item("Milestones1").OffsetTop = 5
    ActiveGanttVBCtl1.Styles.Item("Milestones1").OffsetBottom = 10
    
    ActiveGanttVBCtl1.Styles.Add "Predecessors1"
    
    ActiveGanttVBCtl1.Styles.Item("Predecessors1").ForeColor = RGB(0, 0, 255)
    ActiveGanttVBCtl1.Styles.Item("Predecessors1").PredecessorStyle.Style = GRE_CONNLINESTYLE.PDS_NORMAL
    
    
    AddTaskStyle "Tasks1", RGB(0, 0, 255), RGB(0, 0, 255)
    AddTaskStyle "Tasks2", RGB(0, 255, 0), RGB(0, 255, 0)
    AddTaskStyle "Tasks3", RGB(255, 0, 0), RGB(255, 0, 0)
    AddTaskStyle "Tasks4", RGB(0, 0, 0), RGB(0, 0, 0)
    
    ActiveGanttVBCtl1.Columns.Add ""
    ActiveGanttVBCtl1.Columns.Item(1).Caption = "ID"
    ActiveGanttVBCtl1.Columns.Item(1).Width = 25
    ActiveGanttVBCtl1.Columns.Item(1).StyleIndex = "Columns"
    ActiveGanttVBCtl1.Columns.Add ""
    ActiveGanttVBCtl1.Columns.Item(2).Width = 30
    ActiveGanttVBCtl1.Columns.Item(2).StyleIndex = "Columns"
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

    ActiveGanttVBCtl1.Milestones.Add "", "K4", NewDate(4, 15, 2003), "Mil1", "Milestones1"
    ActiveGanttVBCtl1.Milestones.Item("Mil1").Predecessors.Add "ASCGS", OT_TASK
    ActiveGanttVBCtl1.Milestones.Add "", "K15", NewDate(3, 8, 2003), "Mil2", "Milestones1"
    ActiveGanttVBCtl1.Milestones.Item("Mil2").Predecessors.Add "ASCEDII", OT_TASK

    ActiveGanttVBCtl1.Tasks.Add "", "K1", NewDate(1, 1, 2003), NewDate(4, 1, 2003), "LSI1", "Tasks3"
    ActiveGanttVBCtl1.Tasks.Add "", "K1", NewDate(4, 15, 2003), NewDate(8, 1, 2003), "LSI2", "Tasks4"
    ActiveGanttVBCtl1.Tasks.Item("LSI2").Predecessors.Add "LSI1", OT_TASK
    ActiveGanttVBCtl1.Tasks.Add "", "K3", NewDate(1, 1, 2003), NewDate(2, 1, 2003), "ICCL", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K4", NewDate(2, 1, 2003), NewDate(4, 1, 2003), "ASCGS", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Item("ASCGS").Predecessors.Add "ICCL", OT_TASK
    ActiveGanttVBCtl1.Tasks.Add "", "K5", NewDate(1, 1, 2003), NewDate(1, 15, 2003), "ICF", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K6", NewDate(1, 15, 2003), NewDate(2, 15, 2003), "ICRM", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Item("ICRM").Predecessors.Add "ICF", OT_TASK, , "Predecessors1"
    ActiveGanttVBCtl1.Tasks.Add "", "K7", NewDate(2, 15, 2003), NewDate(4, 1, 2003), "IF", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Item("IF").Predecessors.Add "ICRM", OT_TASK, , "Predecessors1"
    ActiveGanttVBCtl1.Tasks.Add "", "K8", NewDate(4, 15, 2003), NewDate(5, 1, 2003), "ICR", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Item("ICR").Predecessors.Add "IF", OT_TASK, , "Predecessors1"
    ActiveGanttVBCtl1.Tasks.Add "", "K9", NewDate(5, 1, 2003), NewDate(8, 1, 2003), "ACRMP", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Item("ACRMP").Predecessors.Add "ICR", OT_TASK, , "Predecessors1"
    ActiveGanttVBCtl1.Tasks.Add "", "K11", NewDate(1, 1, 2003), NewDate(2, 11, 2003), "ED", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K12", NewDate(2, 15, 2003), NewDate(3, 1, 2003), "ASCEDII", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Item("ASCEDII").Predecessors.Add "ED", OT_TASK
    ActiveGanttVBCtl1.Tasks.Add "", "K13", NewDate(3, 11, 2003), NewDate(7, 15, 2003), "IP401K", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Item("IP401K").Predecessors.Add "ASCEDII", OT_TASK
    ActiveGanttVBCtl1.Tasks.Add "", "K14", NewDate(3, 20, 2003), NewDate(7, 5, 2003), "PROC", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Item("PROC").Predecessors.Add "ASCEDII", OT_TASK
    ActiveGanttVBCtl1.Tasks.Add "", "K15", NewDate(3, 25, 2003), NewDate(6, 25, 2003), "IA", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Item("IA").Predecessors.Add "ASCEDII", OT_TASK
    
    ActiveGanttVBCtl1.Tasks.Add "", "K16", NewDate(3, 20, 2003), NewDate(5, 1, 2003), "RATTTRCK", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Item("RATTTRCK").Predecessors.Add "Mil2", E_OBJECTTYPE.OT_MILESTONE
    ActiveGanttVBCtl1.Tasks.Add "", "K17", NewDate(3, 20, 2003), NewDate(5, 1, 2003), "IEHS", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Item("IEHS").Predecessors.Add "Mil2", E_OBJECTTYPE.OT_MILESTONE
    ActiveGanttVBCtl1.Tasks.Add "", "K18", NewDate(3, 20, 2003), NewDate(5, 1, 2003), "IP", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Item("IP").Predecessors.Add "Mil2", E_OBJECTTYPE.OT_MILESTONE
    ActiveGanttVBCtl1.Tasks.Add "", "K19", NewDate(3, 20, 2003), NewDate(5, 1, 2003), "IPRR", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Item("IPRR").Predecessors.Add "Mil2", E_OBJECTTYPE.OT_MILESTONE
    
    ActiveGanttVBCtl1.Tasks.Add "", "K16", NewDate(5, 7, 2003), NewDate(7, 15, 2003), "RATTTRCK2", "Tasks2"
    ActiveGanttVBCtl1.Tasks.Item("RATTTRCK2").Predecessors.Add "RATTTRCK", OT_TASK, "Predecessor1"
    ActiveGanttVBCtl1.Tasks.Add "", "K17", NewDate(5, 8, 2003), NewDate(7, 15, 2003), "IEHS2", "Tasks2"
    ActiveGanttVBCtl1.Tasks.Item("IEHS2").Predecessors.Add "IEHS", OT_TASK, "Predecessor2"
    ActiveGanttVBCtl1.Tasks.Add "", "K18", NewDate(5, 10, 2003), NewDate(7, 15, 2003), "IP2", "Tasks2"
    ActiveGanttVBCtl1.Tasks.Item("IP2").Predecessors.Add "IP", OT_TASK, "Predecessor3"
    ActiveGanttVBCtl1.Tasks.Add "", "K19", NewDate(5, 15, 2003), NewDate(7, 15, 2003), "IPRR2", "Tasks2"
    ActiveGanttVBCtl1.Tasks.Item("IPRR2").Predecessors.Add "IPRR", OT_TASK, "Predecessor4"
    ActiveGanttVBCtl1.Tasks.Add "", "K21", NewDate(12, 1, 2002), NewDate(4, 1, 2003), "IMSS", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K22", NewDate(12, 1, 2002), NewDate(3, 20, 2003), "ISSSS", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Add "", "K23", NewDate(5, 1, 2003), NewDate(8, 1, 2003), "CMSS", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Item("CMSS").Predecessors.Add "IMSS", OT_TASK, "Predecessor5"
    ActiveGanttVBCtl1.Tasks.Add "", "K24", NewDate(6, 1, 2003), NewDate(8, 1, 2003), "CSSSS", "Tasks1"
    ActiveGanttVBCtl1.Tasks.Item("CSSSS").Predecessors.Add "ISSSS", OT_TASK, "Predecessor6"
    ActiveGanttVBCtl1.Tasks.Add "", "K25", NewDate(1, 1, 2003), NewDate(8, 1, 2003), , "Tasks3"
    ActiveGanttVBCtl1.Tasks.Add "", "K26", NewDate(1, 1, 2003), NewDate(7, 5, 2003), , "Tasks3"

    ActiveGanttVBCtl1.Views.Add "12h", "1m", ST_CUSTOM, ST_CUSTOM, ST_CUSTOM
    ActiveGanttVBCtl1.Views.Item("1").TimeLine.TierArea.UpperTier.Interval = "1q"
    ActiveGanttVBCtl1.Views.Item("1").TimeLine.TierArea.UpperTier.Height = 17
    ActiveGanttVBCtl1.Views.Item("1").TimeLine.TierArea.LowerTier.Interval = "1m"
    ActiveGanttVBCtl1.Views.Item("1").TimeLine.TierArea.LowerTier.Height = 17
    ActiveGanttVBCtl1.Views.Item("1").TimeLine.TickMarkArea.Visible = False
    ActiveGanttVBCtl1.CurrentView = "1"
    ActiveGanttVBCtl1.Splitter.Position = 255
    ActiveGanttVBCtl1.CurrentViewObject.TimeLine.Position NewDate(12, 1, 2002)
    ActiveGanttVBCtl1.Redraw
    
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

Private Sub ActiveGanttVBCtl1_CustomTierDraw(ByVal Position As ActiveGanttVB.E_TIERPOSITION, ByVal StartDate As Date, ByVal EndDate As Date, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal LeftTrim As Long, ByVal RightTrim As Long, ByVal lHdc As Long, Caption As String, StyleIndex As String)
    If Position = SP_LOWER Then
        StyleIndex = "TimeLineTiers"
        Caption = Format(StartDate, "mmm")
    ElseIf Position = SP_UPPER Then
        StyleIndex = "TimeLineTiers"
        If Month(StartDate) >= 1 And Month(StartDate) <= 3 Then
            Caption = Year(StartDate) & " Q1"
        ElseIf Month(StartDate) >= 4 And Month(StartDate) <= 6 Then
            Caption = Year(StartDate) & " Q2"
        ElseIf Month(StartDate) >= 7 And Month(StartDate) <= 9 Then
            Caption = Year(StartDate) & " Q3"
        ElseIf Month(StartDate) >= 10 And Month(StartDate) <= 12 Then
            Caption = Year(StartDate) & " Q4"
        End If
    End If
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

Private Sub AddTaskStyle(ByVal Key As String, ByVal BackColor As OLE_COLOR, ByVal BorderColor As OLE_COLOR)
    ActiveGanttVBCtl1.Styles.Add Key
    ActiveGanttVBCtl1.Styles.Item(Key).Appearance = E_STYLEAPPEARANCE.SA_FLAT
    ActiveGanttVBCtl1.Styles.Item(Key).BackgroundPattern = GRE_BACKGROUNDPATTERN.FP_DOWNWARDDIAGONAL
    ActiveGanttVBCtl1.Styles.Item(Key).Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
    ActiveGanttVBCtl1.Styles.Item(Key).BackColor = BackColor
    ActiveGanttVBCtl1.Styles.Item(Key).BorderColor = BorderColor
    ActiveGanttVBCtl1.Styles.Item(Key).BorderStyle = E_STYLEBORDER.SBR_SINGLE
    ActiveGanttVBCtl1.Styles.Item(Key).OffsetTop = 5
    ActiveGanttVBCtl1.Styles.Item(Key).OffsetBottom = 10
    ActiveGanttVBCtl1.Styles.Item(Key).SelectionRectangleVisible = True
    ActiveGanttVBCtl1.Styles.Item(Key).SelectionRectangleOffsetTop = 0
    ActiveGanttVBCtl1.Styles.Item(Key).SelectionRectangleOffsetLeft = 0
    ActiveGanttVBCtl1.Styles.Item(Key).SelectionRectangleOffsetRight = 0
    ActiveGanttVBCtl1.Styles.Item(Key).SelectionRectangleOffsetBottom = 0
End Sub

Private Sub ActiveGanttVBCtl1_ToolTip(ToolTipAction As ActiveGanttVB.E_TOOLTIPACTION, Caption As String, ByVal StartDate As Date, ByVal EndDate As Date, ByVal Index As Long, DisplayToolTip As Boolean)
    Select Case ToolTipAction
        Case E_TOOLTIPACTION.TA_OVERTASK
            Caption = ActiveGanttVBCtl1.Rows.Item(ActiveGanttVBCtl1.Tasks.Item(Index).RowKey).Cells.Item(3).Caption
        Case E_TOOLTIPACTION.TA_OVERMILESTONE
            Caption = ActiveGanttVBCtl1.Rows.Item(ActiveGanttVBCtl1.Milestones.Item(Index).RowKey).Cells.Item(3).Caption
    End Select
End Sub
