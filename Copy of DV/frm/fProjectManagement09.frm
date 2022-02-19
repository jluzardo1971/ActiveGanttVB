VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{404F18CB-304C-4658-92D0-119074ED7C75}#1.0#0"; "ActiveGanttVB2.ocx"
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
'// The Source Code Store ACTIVEGANTT COMPONENT FOR VISUAL BASIC 6.0. VERSION 2
'// Copyright (c) 2002-2004, Julio Luzardo.
'//
'// All Rights Reserved. No parts of this file may be reproduced or transmitted in any
'// form or by any means without the written permission of the author.
'// ----------------------------------------------------------------------------------------

Option Explicit

Private mp_lRowIndex As Long
Private mp_lCellIndex As Long
Private mp_lSelectedTask As Long
Private mp_lSelectedMilestone As Long

Private Function NewDate(ByVal Month As Long, ByVal Day As Long, ByVal Year As Long) As Date
    NewDate = DateSerial(Year, Month, Day)
End Function


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

Private Sub Form_Load()
    Me.Caption = "The Source Code Store - ActiveGantt Scheduler Control Version " & ActiveGanttVBCtl1.Version & " - Project Management Example"
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Dim lIndex As Long
    With ActiveGanttVBCtl1
        .AutomaticRedraw = False
        .DefaultValues.RowHeight = 20
        .AllowRowSize = False
        With .Styles
            .Add "Title"
            .Item("Title").Appearance = E_STYLEAPPEARANCE.SA_RAISED
            .Item("Title").CaptionAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT
            .Item("Title").CaptionXMargin = 5
            .Item("Title").Font.Bold = True
            .Add "TimeLineTiers"
            .Item("TimeLineTiers").Font.Size = 7
            .Item("TimeLineTiers").Font.Bold = True
            .Add "Task"
            .Item("Task").CaptionAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT
            .Item("Task").CaptionXMargin = 25
            .Item("Task").Font.Name = "Microsoft Sans Serif"
            .Item("Task").Font.Size = 7
            .Item("Task").BorderStyle = E_STYLEBORDER.SBR_NONE
            .Item("Task").Appearance = E_STYLEAPPEARANCE.SA_CELL
            .Item("Task").BackColor = &H80000005
            '.Item("Task").BorderColor = ActiveGanttVBCtl1.GridLinesColor
            .Add "Columns"
            Set .Item("Columns").ImageList = imglstColumns
            .Item("Columns").Appearance = E_STYLEAPPEARANCE.SA_RAISED
            .Item("Columns").CaptionAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM
            .Item("Columns").CaptionYMargin = 5
            .Item("Columns").Font.Bold = True
            .Item("Columns").PictureAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM
            .Add "Cells"
            .Item("Cells").Appearance = E_STYLEAPPEARANCE.SA_RAISED
            .Add "Milestones1"
            With .Item("Milestones1")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .StartShapeIndex = 3
                .Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
                .OffsetTop = 5
                .OffsetBottom = 10
            End With
            .Add "Predecessors1"
            With .Item("Predecessors1")
                .ForeColor = RGB(0, 0, 255)
                .PredecessorStyle = GRE_CONNLINESTYLE.PDS_NORMAL
            End With
            .Add "Tasks1"
            With .Item("Tasks1")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .BackgroundPattern = GRE_BACKGROUNDPATTERN.FP_DOWNWARDDIAGONAL
                .Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
                .BackColor = RGB(0, 0, 255)
                .BorderColor = RGB(0, 0, 255)
                .BorderStyle = E_STYLEBORDER.SBR_SINGLE
                .OffsetTop = 5
                .OffsetBottom = 10
                .SelectionRectangleVisible = True
                .SelectionRectangleOffsetTop = 0
                .SelectionRectangleOffsetLeft = 0
                .SelectionRectangleOffsetRight = 0
                .SelectionRectangleOffsetBottom = 0
            End With
            .Add "Tasks2"
            With .Item("Tasks2")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .BackgroundPattern = GRE_BACKGROUNDPATTERN.FP_DOWNWARDDIAGONAL
                .Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
                .BackColor = RGB(0, 255, 0)
                .BorderColor = RGB(0, 255, 0)
                .BorderStyle = E_STYLEBORDER.SBR_SINGLE
                .OffsetTop = 5
                .OffsetBottom = 10
                .SelectionRectangleVisible = True
                .SelectionRectangleOffsetTop = 0
                .SelectionRectangleOffsetLeft = 0
                .SelectionRectangleOffsetRight = 0
                .SelectionRectangleOffsetBottom = 0
            End With
            .Add "Tasks3"
            With .Item("Tasks3")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .BackgroundPattern = GRE_BACKGROUNDPATTERN.FP_DOWNWARDDIAGONAL
                .Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
                .BackColor = RGB(255, 0, 0)
                .BorderColor = RGB(255, 0, 0)
                .BorderStyle = E_STYLEBORDER.SBR_SINGLE
                .OffsetTop = 5
                .OffsetBottom = 10
                .SelectionRectangleVisible = True
                .SelectionRectangleOffsetTop = 0
                .SelectionRectangleOffsetLeft = 0
                .SelectionRectangleOffsetRight = 0
                .SelectionRectangleOffsetBottom = 0
                .StartShapeIndex = 1
                .EndShapeIndex = 2
            End With
            .Add "Tasks4"
            With .Item("Tasks4")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .BackgroundPattern = GRE_BACKGROUNDPATTERN.FP_DOWNWARDDIAGONAL
                .Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
                .BackColor = RGB(0, 0, 0)
                .BorderColor = RGB(0, 0, 0)
                .BorderStyle = E_STYLEBORDER.SBR_SINGLE
                .OffsetTop = 5
                .OffsetBottom = 10
                .SelectionRectangleVisible = True
                .SelectionRectangleOffsetTop = 0
                .SelectionRectangleOffsetLeft = 0
                .SelectionRectangleOffsetRight = 0
                .SelectionRectangleOffsetBottom = 0
            End With
        End With
        With .Columns
            .Add ""
            .Item(1).Caption = "ID"
            .Item(1).Width = 25
            .Item(1).StyleIndex = "Columns"
            .Add ""
            .Item(2).Width = 30
            .Item(2).StyleIndex = "Columns"
            .Item(2).PictureIndex = 1
            .Add "Task Name"
            .Item(3).Width = 200
            .Item(3).StyleIndex = "Columns"
        End With
        With .Rows
            .Add "K1"
            .Item("K1").Cells.Item(3).Caption = "Legacy Systems Integration"
            .Item("K1").Cells.Item(3).StyleIndex = "Title"
            .Add "K2"
            .Item("K2").Cells.Item(3).Caption = "Implement CGS"
            .Item("K2").Cells.Item(3).StyleIndex = "Title"
            .Item("K2").Container = False
            .Add "K3"
            .Item("K3").Cells.Item(3).Caption = "Implement Customers/Contact/Leads"
            .Add "K4"
            .Item("K4").Cells.Item(3).Caption = "Analyse Strategy for Completing CGS Implementation"
            .Add "K5"
            .Item("K5").Cells.Item(3).Caption = "Implement Customer Financials"
            .Add "K6"
            .Item("K6").Cells.Item(3).Caption = "Implement CRM"
            .Add "K7"
            .Item("K7").Cells.Item(3).Caption = "Implement Financials"
            .Add "K8"
            .Item("K8").Cells.Item(3).Caption = "Implement Customer Records"
            .Add "K9"
            .Item("K9").Cells.Item(3).Caption = "Advanced CRM Pilot"
            .Add "K10"
            .Item("K10").Cells.Item(3).Caption = "Implement EDI"
            .Item("K10").Cells.Item(3).StyleIndex = "Title"
            .Item("K10").Container = False
            .Add "K11"
            .Item("K11").Cells.Item(3).Caption = "Electronic Documents"
            .Add "K12"
            .Item("K12").Cells.Item(3).Caption = "Analyse Strategy for Completing EDI Implementation"
            .Add "K13"
            .Item("K13").Cells.Item(3).Caption = "Implement Payroll/401 K"
            .Add "K14"
            .Item("K14").Cells.Item(3).Caption = "Procurement"
            .Add "K15"
            .Item("K15").Cells.Item(3).Caption = "Implement Administration"
            .Add "K16"
            .Item("K16").Cells.Item(3).Caption = "RATTLE Tracking"
            .Add "K17"
            .Item("K17").Cells.Item(3).Caption = "Implement Employee Health and Safety"
            .Add "K18"
            .Item("K18").Cells.Item(3).Caption = "Implement Planning"
            .Add "K19"
            .Item("K19").Cells.Item(3).Caption = "Implement Remaining Records"
            .Add "K20"
            .Item("K20").Cells.Item(3).Caption = "System Environment"
            .Item("K20").Cells.Item(3).StyleIndex = "Title"
            .Item("K20").Container = False
            .Add "K21"
            .Item("K21").Cells.Item(3).Caption = "Initial Management Support System"
            .Add "K22"
            .Item("K22").Cells.Item(3).Caption = "Initial Self Service and Support System"
            .Add "K23"
            .Item("K23").Cells.Item(3).Caption = "Complete Management Support System"
            .Add "K24"
            .Item("K24").Cells.Item(3).Caption = "Complete Self Service and Support System"
            .Add "K25"
            .Item("K25").Cells.Item(3).Caption = "User Support and Documentation"
            .Item("K25").Cells.Item(3).StyleIndex = "Title"
            .Add "K26"
            .Item("K26").Cells.Item(3).Caption = "Systems Infrastructure"
            .Item("K26").Cells.Item(3).StyleIndex = "Title"
            For lIndex = 1 To .Count
                .Item("K" & lIndex).Cells.Item(1).Caption = lIndex
                .Item("K" & lIndex).Cells.Item(1).StyleIndex = "Cells"
                .Item("K" & lIndex).Cells.Item(2).StyleIndex = "Cells"
                If .Item("K" & lIndex).Cells.Item(3).StyleIndex = "0" Then
                    .Item("K" & lIndex).Cells.Item(3).StyleIndex = "Task"
                End If
            Next lIndex
        End With
        With .Milestones
            .Add "", "K4", NewDate(4, 15, 2003), "Mil1", "Milestones1"
            .Item("Mil1").Predecessors.Add "ASCGS", OT_Task
            .Add "", "K15", NewDate(3, 8, 2003), "Mil2", "Milestones1"
            .Item("Mil2").Predecessors.Add "ASCEDII", OT_Task
        End With
        With .Tasks
            .Add "", "K1", NewDate(1, 1, 2003), NewDate(4, 1, 2003), "LSI1", "Tasks3"
            .Add "", "K1", NewDate(4, 15, 2003), NewDate(8, 1, 2003), "LSI2", "Tasks4"
            .Item("LSI2").Predecessors.Add "LSI1", OT_Task
            .Add "", "K3", NewDate(1, 1, 2003), NewDate(2, 1, 2003), "ICCL", "Tasks1"
            .Add "", "K4", NewDate(2, 1, 2003), NewDate(4, 1, 2003), "ASCGS", "Tasks1"
            .Item("ASCGS").Predecessors.Add "ICCL", OT_Task
            .Add "", "K5", NewDate(1, 1, 2003), NewDate(1, 15, 2003), "ICF", "Tasks1"
            .Add "", "K6", NewDate(1, 15, 2003), NewDate(2, 15, 2003), "ICRM", "Tasks1"
            .Item("ICRM").Predecessors.Add "ICF", OT_Task, , "Predecessors1"
            .Add "", "K7", NewDate(2, 15, 2003), NewDate(4, 1, 2003), "IF", "Tasks1"
            .Item("IF").Predecessors.Add "ICRM", OT_Task, , "Predecessors1"
            .Add "", "K8", NewDate(4, 15, 2003), NewDate(5, 1, 2003), "ICR", "Tasks1"
            .Item("ICR").Predecessors.Add "IF", OT_Task, , "Predecessors1"
            .Add "", "K9", NewDate(5, 1, 2003), NewDate(8, 1, 2003), "ACRMP", "Tasks1"
            .Item("ACRMP").Predecessors.Add "ICR", OT_Task, , "Predecessors1"
            .Add "", "K11", NewDate(1, 1, 2003), NewDate(2, 11, 2003), "ED", "Tasks1"
            .Add "", "K12", NewDate(2, 15, 2003), NewDate(3, 1, 2003), "ASCEDII", "Tasks1"
            .Item("ASCEDII").Predecessors.Add "ED", OT_Task
            .Add "", "K13", NewDate(3, 11, 2003), NewDate(7, 15, 2003), "IP401K", "Tasks1"
            .Item("IP401K").Predecessors.Add "ASCEDII", OT_Task
            .Add "", "K14", NewDate(3, 20, 2003), NewDate(7, 5, 2003), "PROC", "Tasks1"
            .Item("PROC").Predecessors.Add "ASCEDII", OT_Task
            .Add "", "K15", NewDate(3, 25, 2003), NewDate(6, 25, 2003), "IA", "Tasks1"
            .Item("IA").Predecessors.Add "ASCEDII", OT_Task
            
            .Add "", "K16", NewDate(3, 20, 2003), NewDate(5, 1, 2003), "RATTTRCK", "Tasks1"
            .Item("RATTTRCK").Predecessors.Add "Mil2", E_OBJECTTYPE.OT_MILESTONE
            .Add "", "K17", NewDate(3, 20, 2003), NewDate(5, 1, 2003), "IEHS", "Tasks1"
            .Item("IEHS").Predecessors.Add "Mil2", E_OBJECTTYPE.OT_MILESTONE
            .Add "", "K18", NewDate(3, 20, 2003), NewDate(5, 1, 2003), "IP", "Tasks1"
            .Item("IP").Predecessors.Add "Mil2", E_OBJECTTYPE.OT_MILESTONE
            .Add "", "K19", NewDate(3, 20, 2003), NewDate(5, 1, 2003), "IPRR", "Tasks1"
            .Item("IPRR").Predecessors.Add "Mil2", E_OBJECTTYPE.OT_MILESTONE
            
            .Add "", "K16", NewDate(5, 7, 2003), NewDate(7, 15, 2003), "RATTTRCK2", "Tasks2"
            .Item("RATTTRCK2").Predecessors.Add "RATTTRCK", OT_Task, "Predecessor1"
            .Add "", "K17", NewDate(5, 8, 2003), NewDate(7, 15, 2003), "IEHS2", "Tasks2"
            .Item("IEHS2").Predecessors.Add "IEHS", OT_Task, "Predecessor2"
            .Add "", "K18", NewDate(5, 10, 2003), NewDate(7, 15, 2003), "IP2", "Tasks2"
            .Item("IP2").Predecessors.Add "IP", OT_Task, "Predecessor3"
            .Add "", "K19", NewDate(5, 15, 2003), NewDate(7, 15, 2003), "IPRR2", "Tasks2"
            .Item("IPRR2").Predecessors.Add "IPRR", OT_Task, "Predecessor4"
            .Add "", "K21", NewDate(12, 1, 2002), NewDate(4, 1, 2003), "IMSS", "Tasks1"
            .Add "", "K22", NewDate(12, 1, 2002), NewDate(3, 20, 2003), "ISSSS", "Tasks1"
            .Add "", "K23", NewDate(5, 1, 2003), NewDate(8, 1, 2003), "CMSS", "Tasks1"
            .Item("CMSS").Predecessors.Add "IMSS", OT_Task, "Predecessor5"
            .Add "", "K24", NewDate(6, 1, 2003), NewDate(8, 1, 2003), "CSSSS", "Tasks1"
            .Item("CSSSS").Predecessors.Add "ISSSS", OT_Task, "Predecessor6"
            .Add "", "K25", NewDate(1, 1, 2003), NewDate(8, 1, 2003), , "Tasks3"
            .Add "", "K26", NewDate(1, 1, 2003), NewDate(7, 5, 2003), , "Tasks3"
        End With
        .Views.Add "12h", "1m", ST_CUSTOM, ST_CUSTOM, ST_CUSTOM
        .Views.Item("1").UpperTier.Interval = "1q"
        .Views.Item("1").UpperTier.Height = 17
        .Views.Item("1").LowerTier.Interval = "1m"
        .Views.Item("1").LowerTier.Height = 17
        .Views.Item("1").TickMarkArea.Visible = False
        .CurrentView = "1"
        .SplitterPosition = 255
        .AutomaticRedraw = True
        .CurrentViewObject.TimeLine.Position NewDate(12, 1, 2002)
        .Redraw
    End With
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuDelete_Click()
    If mp_lSelectedTask <> 0 Then
        ActiveGanttVBCtl1.Tasks.Remove mp_lSelectedTask
    ElseIf mp_lSelectedMilestone <> 0 Then
        ActiveGanttVBCtl1.Milestones.Remove mp_lSelectedMilestone
    End If
End Sub

Private Sub mnuPrint_Click()
    Set fPrintDialog.mp_oControl = ActiveGanttVBCtl1
    fPrintDialog.Show 1, Me
    ActiveGanttVBCtl1.Printer.Terminate
    ActiveGanttVBCtl1.Redraw
End Sub

Private Sub ActiveGanttVBCtl1_ToolTipHoveringOverTask(Caption As String, ByVal Index As Long, DisplayToolTip As Boolean)
    '// Display Custom Tooltip
    Caption = ActiveGanttVBCtl1.Rows.Item(ActiveGanttVBCtl1.Tasks.Item(Index).RowKey).Cells.Item(3).Caption
End Sub

Private Sub ActiveGanttVBCtl1_ToolTipHoveringOverMilestone(Caption As String, ByVal Index As Long, DisplayToolTip As Boolean)
    '// Display Custom Tooltip
    Caption = ActiveGanttVBCtl1.Rows.Item(ActiveGanttVBCtl1.Milestones.Item(Index).RowKey).Cells.Item(3).Caption
End Sub

Private Sub ActiveGanttVBCtl1_ToolTipMovingMilestone(Caption As String, ByVal MilestoneDate As Date, DisplayToolTip As Boolean)
    '// Display Custom Tooltip
    Caption = "Milestone is Being Moved"
End Sub

Private Sub ActiveGanttVBCtl1_ToolTipMovingTask(Caption As String, ByVal StartDate As Date, ByVal EndDate As Date, DisplayToolTip As Boolean)
    '// Display Custom Tooltip
    Caption = "Task is Being Moved"
End Sub

Private Sub ActiveGanttVBCtl1_ToolTipStretchingTask(Caption As String, ByVal NewDate As Date, DisplayToolTip As Boolean)
    '// Display Custom Tooltip
    Caption = "Task is Being Stretched"
End Sub

Private Sub ActiveGanttVBCtl1_ToolTipAddingTask(Caption As String, ByVal StartDate As Date, ByVal EndDate As Date, DisplayToolTip As Boolean)
    '// Display Custom Tooltip
    Caption = "Task is Being Addded"
End Sub

Private Sub ActiveGanttVBCtl1_ToolTipCanAddOrStretching(Caption As String, DisplayToolTip As Boolean)
    '// Cursor is over client area display nothing
    DisplayToolTip = False
End Sub
