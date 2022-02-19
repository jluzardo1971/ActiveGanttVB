VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CCC1C7D1-F592-4261-9A57-CA48F869B175}#1.0#0"; "ActiveGanttVB2.ocx"
Object = "{A4F6894D-7F88-4359-BACE-61F7DE949168}#1.0#0"; "XTTreeviewVB.ocx"
Begin VB.Form fProjectManagement03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      Begin XTTreeviewVB.XTTreeviewVBCtl XTTreeviewVBCtl1 
         Height          =   7335
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   12938
      End
      Begin ActiveGanttVB.ActiveGanttVBCtl ActiveGanttVBCtl1 
         Height          =   7335
         Left            =   3840
         TabIndex        =   1
         Top             =   360
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   12938
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
   Begin MSComctlLib.ImageList imgTreeView 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fProjectManagement03.frx":0000
            Key             =   "FolderClosed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fProjectManagement03.frx":0354
            Key             =   "FolderOpen"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fProjectManagement03.frx":06A8
            Key             =   "ActiveGantt"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fProjectManagement03.frx":09FC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fProjectManagement03"
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
    
    
    XTTreeviewVBCtl1.Styles.Add "Columns"
    XTTreeviewVBCtl1.Styles.Item("Columns").Appearance = E_STYLEAPPEARANCE.SA_RAISED
    XTTreeviewVBCtl1.Styles.Item("Columns").CaptionAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM
    XTTreeviewVBCtl1.Styles.Item("Columns").CaptionYMargin = 5
    XTTreeviewVBCtl1.Styles.Item("Columns").Font.Bold = True
    
    XTTreeviewVBCtl1.Header.Caption = "Task Name"
    XTTreeviewVBCtl1.Header.StyleIndex = "Columns"
    
    XTTreeviewVBCtl1.DefaultValues.NodeHeight = 20
    
    AddTitleNode "K1", "Legacy Systems Integration"
    
    AddTitleNode "K2", "Implement CGS"
    AddNode "K2", "K3", "Implement Customers/Contact/Leads"
    AddNode "K2", "K4", "Analyse Strategy for Completing CGS Implementation"
    AddNode "K2", "K5", "Implement Customer Financials"
    AddNode "K2", "K6", "Implement CRM"
    AddNode "K2", "K7", "Implement Financials"
    AddNode "K2", "K8", "Implement Customer Records"
    AddNode "K2", "K9", "Advanced CRM Pilot"
    
    AddTitleNode "K10", "Implement EDI"
    AddNode "K10", "K11", "Electronic Documents"
    AddNode "K10", "K12", "Analyse Strategy for Completing EDI Implementation"
    AddNode "K10", "K13", "Implement Payroll/401 K"
    AddNode "K10", "K14", "Procurement"
    AddNode "K10", "K15", "Implement Administration"
    AddNode "K10", "K16", "RATTLE Tracking"
    AddNode "K10", "K17", "Implement Employee Health and Safety"
    AddNode "K10", "K18", "Implement Planning"
    AddNode "K10", "K19", "Implement Remaining Records"
    
    AddTitleNode "K20", "System Environment"
    AddNode "K20", "K21", "Initial Management Support System"
    AddNode "K20", "K22", "Initial Self Service and Support System"
    AddNode "K20", "K23", "Complete Management Support System"
    AddNode "K20", "K24", "Complete Self Service and Support System"
    
    AddTitleNode "K25", "User Support and Documentation"

    AddTitleNode "K26", "Systems Infrastructure"

    XTTreeviewVBCtl1.Pictures = True
    XTTreeviewVBCtl1.Checkboxes = True
    XTTreeviewVBCtl1.FullRowSelect = True
    XTTreeviewVBCtl1.HorizontalLines = True
    
    ActiveGanttVBCtl1.ScrollBarsVisible = False
    ActiveGanttVBCtl1.DefaultValues.RowHeight = 20
    ActiveGanttVBCtl1.AllowRowSize = False
    ActiveGanttVBCtl1.AddMode = AT_TASKADD
    
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
    '.Item("Task").BorderColor = ActiveGanttVBCtl1.GridLinesColor
    
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

    
    ActiveGanttVBCtl1.Rows.Add "K1"
    ActiveGanttVBCtl1.Rows.Add "K2"
    ActiveGanttVBCtl1.Rows.Add "K3"
    ActiveGanttVBCtl1.Rows.Add "K4"
    ActiveGanttVBCtl1.Rows.Add "K5"
    ActiveGanttVBCtl1.Rows.Add "K6"
    ActiveGanttVBCtl1.Rows.Add "K7"
    ActiveGanttVBCtl1.Rows.Add "K8"
    ActiveGanttVBCtl1.Rows.Add "K9"
    ActiveGanttVBCtl1.Rows.Add "K10"
    ActiveGanttVBCtl1.Rows.Add "K11"
    ActiveGanttVBCtl1.Rows.Add "K12"
    ActiveGanttVBCtl1.Rows.Add "K13"
    ActiveGanttVBCtl1.Rows.Add "K14"
    ActiveGanttVBCtl1.Rows.Add "K15"
    ActiveGanttVBCtl1.Rows.Add "K16"
    ActiveGanttVBCtl1.Rows.Add "K17"
    ActiveGanttVBCtl1.Rows.Add "K18"
    ActiveGanttVBCtl1.Rows.Add "K19"
    ActiveGanttVBCtl1.Rows.Add "K20"
    ActiveGanttVBCtl1.Rows.Add "K21"
    ActiveGanttVBCtl1.Rows.Add "K22"
    ActiveGanttVBCtl1.Rows.Add "K23"
    ActiveGanttVBCtl1.Rows.Add "K24"
    ActiveGanttVBCtl1.Rows.Add "K25"
    ActiveGanttVBCtl1.Rows.Add "K26"

    
    ActiveGanttVBCtl1.Milestones.Add "", "K4", NewDate(4, 15, 2003), "Mil1", "Milestones1"
    ActiveGanttVBCtl1.Milestones.Item("Mil1").Predecessors.Add "ASCGS", OT_TASK
    ActiveGanttVBCtl1.Milestones.Add "", "K15", NewDate(3, 8, 2003), "Mil2", "Milestones1"
    ActiveGanttVBCtl1.Milestones.Item("Mil2").Predecessors.Add "ASCEDII", OT_TASK
    
    ActiveGanttVBCtl1.Tasks.Add "", "K1", NewDate(1, 1, 2003), NewDate(4, 1, 2003), "LSI1", "Tasks3"
    ActiveGanttVBCtl1.Tasks.Add "", "K1", NewDate(4, 15, 2003), NewDate(8, 1, 2003), "LSI2", "Tasks3"
    ActiveGanttVBCtl1.Tasks.Item("LSI2").Predecessors.Add "LSI1", OT_TASK
    
    ActiveGanttVBCtl1.Tasks.Add "", "K2", NewDate(1, 1, 2003), NewDate(8, 1, 2003), "", "Tasks3"
    
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
    
    ActiveGanttVBCtl1.Tasks.Add "", "K10", NewDate(1, 1, 2003), NewDate(7, 15, 2003), "", "Tasks3"
    
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
    
    ActiveGanttVBCtl1.Tasks.Add "", "K20", NewDate(12, 1, 2002), NewDate(8, 1, 2003), "", "Tasks3"
    
    
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
    
    XTTreeviewVBCtl1.Header.Height = ActiveGanttVBCtl1.CurrentViewObject.TimeLine.Height
    
    XTTreeviewVBCtl1.Redraw
    ActiveGanttVBCtl1.Redraw
    
    XTTreeviewVBCtl1.WriteXML "C:\XTTest1.xml"
    
    
End Sub

Private Sub AddTitleNode(ByVal sKey As String, ByVal sCaption As String)
    XTTreeviewVBCtl1.Nodes.Add "", RS_CHILD, sKey, sCaption
    Set XTTreeviewVBCtl1.Nodes.Item(sKey).Image = imgTreeView.ListImages.Item(1).Picture
    Set XTTreeviewVBCtl1.Nodes.Item(sKey).ExpandedImage = imgTreeView.ListImages.Item(2).Picture
    XTTreeviewVBCtl1.Nodes.Item(sKey).Font.Bold = True
End Sub

Private Sub AddNode(ByVal sParentKey As String, ByVal sKey As String, ByVal sCaption As String)
    XTTreeviewVBCtl1.Nodes.Add sParentKey, RS_CHILD, sKey, sCaption
    Set XTTreeviewVBCtl1.Nodes.Item(sKey).Image = imgTreeView.ListImages.Item(4).Picture
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


Private Sub XTTreeviewVBCtl1_ControlRedrawn()
    If ActiveGanttVBCtl1.Rows.Count = 0 Then
        Exit Sub
    End If
    ActiveGanttVBCtl1.CurrentViewObject.ClientArea.FirstVisibleRow = XTTreeviewVBCtl1.FirstVisibleNode
    Dim lIndex As Long
    Dim oNode As clsNode
    For lIndex = XTTreeviewVBCtl1.FirstVisibleNode To XTTreeviewVBCtl1.LastVisibleNode
        Set oNode = XTTreeviewVBCtl1.Nodes.Item(lIndex)
        If oNode.Visible = True Then
            ActiveGanttVBCtl1.Rows.Item(oNode.Key).Height = 20
        Else
            ActiveGanttVBCtl1.Rows.Item(oNode.Key).Height = -1
        End If
    Next lIndex
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

