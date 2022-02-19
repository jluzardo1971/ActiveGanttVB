VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CCC1C7D1-F592-4261-9A57-CA48F869B175}#1.0#0"; "ActiveGanttVB2.ocx"
Begin VB.Form fProjectManagement08A 
   Caption         =   "Form1"
   ClientHeight    =   7560
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraForm 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   255
         Left            =   9480
         TabIndex        =   9
         Top             =   7080
         Width           =   1695
      End
      Begin VB.TextBox txtProjectTitle 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   9855
      End
      Begin VB.CommandButton cmdChangeTimeLine 
         Caption         =   "Change TimeLine"
         Height          =   255
         Left            =   5160
         TabIndex        =   1
         Top             =   7080
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   255
         Left            =   3600
         TabIndex        =   2
         Top             =   7080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   51511297
         CurrentDate     =   38271
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   7080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   51511297
         CurrentDate     =   38271
      End
      Begin MSComctlLib.ImageList imglstColumns 
         Left            =   480
         Top             =   840
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
               Picture         =   "fProjectManagement08A.frx":0000
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   720
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
               Picture         =   "fProjectManagement08A.frx":06E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement08A.frx":0AAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement08A.frx":0E74
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement08A.frx":123C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement08A.frx":1604
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement08A.frx":19CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement08A.frx":1D94
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement08A.frx":215C
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement08A.frx":2524
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement08A.frx":28EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement08A.frx":2CB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fProjectManagement08A.frx":307C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin ActiveGanttVB.ActiveGanttVBCtl ActiveGanttVBCtl1 
         Height          =   6375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   11245
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
      Begin VB.Label Label3 
         Caption         =   "Project Title:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "StartDate:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   7080
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "EndDate:"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   7080
         Width           =   855
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuAddRows 
         Caption         =   "Add Rows"
      End
   End
   Begin VB.Menu mnuTaskStyles 
      Caption         =   "Task Styles"
      Begin VB.Menu mnuTaskStyle1 
         Caption         =   "TaskStyle1"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuTaskStyle2 
         Caption         =   "TaskStyle2"
      End
      Begin VB.Menu mnuTaskStyle3 
         Caption         =   "TaskStyle3"
      End
      Begin VB.Menu mnuTaskStyle4 
         Caption         =   "TaskStyle4"
      End
   End
End
Attribute VB_Name = "fProjectManagement08A"
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

Public bAddNew As Boolean
Public lProjectID As Long

Public sRowCaption As String
Public bRowContainer As Boolean
Public bRowTitle As Boolean
Public bRowOK As Boolean

Private Function NewDate(ByVal Month As Long, ByVal Day As Long, ByVal Year As Long) As Date
    NewDate = DateSerial(Year, Month, Day)
End Function

Private Sub ActiveGanttVBCtl1_ObjectAdded(ByVal EventTarget As ActiveGanttVB.E_EVENTTARGET, ByVal ObjectIndex As Long)
    If EventTarget = EVT_TASK Then
        If mnuTaskStyle1.Checked = True Then
            ActiveGanttVBCtl1.Tasks.Item(ObjectIndex).StyleIndex = "Tasks1"
        ElseIf mnuTaskStyle2.Checked = True Then
            ActiveGanttVBCtl1.Tasks.Item(ObjectIndex).StyleIndex = "Tasks2"
        ElseIf mnuTaskStyle3.Checked = True Then
            ActiveGanttVBCtl1.Tasks.Item(ObjectIndex).StyleIndex = "Tasks3"
        ElseIf mnuTaskStyle4.Checked = True Then
            ActiveGanttVBCtl1.Tasks.Item(ObjectIndex).StyleIndex = "Tasks4"
        End If
    End If
End Sub

Private Sub cmdChangeTimeLine_Click()
    If dtpEndDate.Value < dtpStartDate.Value Or dtpEndDate.Value = dtpStartDate.Value Then
        MsgBox "EndDate must be greater than StartDate", vbOKOnly, "Error"
        Exit Sub
    End If
    '// Interval setting(The smaller the interval the greater the accuracy, but setting and interval that is too small
    '// Will generate an overflow)
    Dim sInterval As String
    sInterval = mp_ChooseInterval
    '// Calculate the length of the TimeLine in pixels
    Dim lStart As Long
    Dim lEnd As Long
    Dim lTimeLineLength As Long
    lStart = ActiveGanttVBCtl1.MathLib.GetXCoordinateFromDate(ActiveGanttVBCtl1.CurrentViewObject.TimeLine.StartDate)
    lEnd = ActiveGanttVBCtl1.MathLib.GetXCoordinateFromDate(ActiveGanttVBCtl1.CurrentViewObject.TimeLine.EndDate)
    lTimeLineLength = lEnd - lStart
    '// Calculate the difference in minutes between the desired dates
    Dim lDifference As Long
    lDifference = ActiveGanttVBCtl1.MathLib.DateTimeDiff(sInterval, dtpStartDate.Value, dtpEndDate.Value)
    '// Calculate the factor
    Dim lFactor As Long
    lFactor = lDifference / lTimeLineLength
    '// Set the view's new interval setting
    ActiveGanttVBCtl1.CurrentViewObject.Interval = lFactor & sInterval
    ActiveGanttVBCtl1.CurrentViewObject.TimeLine.Position dtpStartDate.Value
    ActiveGanttVBCtl1.Redraw
End Sub

Private Function mp_ChooseInterval() As String
    If mp_IsValidInterval("s") = True Then
        mp_ChooseInterval = "s"
        Exit Function
    End If
    If mp_IsValidInterval("n") = True Then
        mp_ChooseInterval = "n"
        Exit Function
    End If
    If mp_IsValidInterval("h") = True Then
        mp_ChooseInterval = "h"
        Exit Function
    End If
    If mp_IsValidInterval("d") = True Then
        mp_ChooseInterval = "d"
        Exit Function
    End If
    If mp_IsValidInterval("m") = True Then
        mp_ChooseInterval = "m"
        Exit Function
    End If
    If mp_IsValidInterval("yyyy") = True Then
        mp_ChooseInterval = "yyyy"
        Exit Function
    End If
End Function

Private Function mp_IsValidInterval(ByVal sInterval As String) As Boolean
    Dim lDummy As Long
    lDummy = ActiveGanttVBCtl1.MathLib.DateTimeDiff(sInterval, dtpStartDate.Value, dtpEndDate.Value)
    If lDummy <> 0 Then
        mp_IsValidInterval = True
    Else
        '// DateTimeDiff Function returns 0 when it encounters an overflow
        mp_IsValidInterval = False
    End If
End Function

Private Sub cmdUpdate_Click()
    Dim tb_Projects As New ADODB.Recordset
    Dim ProjectRepositoryMDB As New ADODB.Connection
  
    ProjectRepositoryMDB.Open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "\frm\ProjectRepository.mdb;UID=;PWD="
    tb_Projects.CursorType = adOpenKeyset
    tb_Projects.LockType = adLockOptimistic
    
    If bAddNew = True Then
        tb_Projects.Open "SELECT * FROM tb_Projects ORDER BY lProjectID", ProjectRepositoryMDB
        tb_Projects.AddNew
    Else
        tb_Projects.Open "SELECT * FROM tb_Projects WHERE lProjectID =" & lProjectID, ProjectRepositoryMDB
    End If
    tb_Projects!sProjectDescription = txtProjectTitle.Text
    tb_Projects!sRowsXML = ActiveGanttVBCtl1.Rows.GetXML()
    tb_Projects!sViewsXML = ActiveGanttVBCtl1.Views.GetXML()
    tb_Projects!sTasksXML = ActiveGanttVBCtl1.Tasks.GetXML()
    tb_Projects!dtStartDate = ActiveGanttVBCtl1.CurrentViewObject.TimeLine.StartDate
    tb_Projects!dtEndDate = ActiveGanttVBCtl1.CurrentViewObject.TimeLine.EndDate
    tb_Projects.Update
    
    tb_Projects.Close
    ProjectRepositoryMDB.Close

End Sub

Private Sub Form_Load()
    Me.Caption = "The Source Code Store - ActiveGantt Scheduler Control Version " & ActiveGanttVBCtl1.Version & " - Project Management Example"
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    dtpStartDate.Value = NewDate(12, 1, 2002)
    dtpEndDate.Value = NewDate(8, 1, 2003)
    
    
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
    
    If bAddNew = True Then
        ActiveGanttVBCtl1.Views.Add "12h", "1m", ST_CUSTOM, ST_CUSTOM, ST_CUSTOM
        ActiveGanttVBCtl1.Views.Item("1").TimeLine.TierArea.UpperTier.Interval = "1q"
        ActiveGanttVBCtl1.Views.Item("1").TimeLine.TierArea.UpperTier.Height = 17
        ActiveGanttVBCtl1.Views.Item("1").TimeLine.TierArea.LowerTier.Interval = "1m"
        ActiveGanttVBCtl1.Views.Item("1").TimeLine.TierArea.LowerTier.Height = 17
        ActiveGanttVBCtl1.Views.Item("1").TimeLine.TickMarkArea.Visible = False
        ActiveGanttVBCtl1.CurrentView = "1"
        ActiveGanttVBCtl1.Splitter.Position = 255
        ActiveGanttVBCtl1.CurrentViewObject.TimeLine.Position NewDate(12, 1, 2002)
    Else
        Dim tb_Projects As New ADODB.Recordset
        Dim ProjectRepositoryMDB As New ADODB.Connection
        
        ProjectRepositoryMDB.Open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "\frm\ProjectRepository.mdb;UID=;PWD="
        tb_Projects.CursorType = adOpenKeyset
        tb_Projects.LockType = adLockReadOnly
        
        tb_Projects.Open "SELECT * FROM tb_Projects WHERE lProjectID =" & lProjectID, ProjectRepositoryMDB
    
        txtProjectTitle.Text = tb_Projects!sProjectDescription
        ActiveGanttVBCtl1.Rows.SetXML tb_Projects!sRowsXML
        ActiveGanttVBCtl1.Tasks.SetXML tb_Projects!sTasksXML
        ActiveGanttVBCtl1.Views.SetXML tb_Projects!sViewsXML
        
        tb_Projects.Close
        ProjectRepositoryMDB.Close
        ActiveGanttVBCtl1.CurrentView = "1"
        ActiveGanttVBCtl1.Splitter.Position = 255

    End If
    
    dtpStartDate = ActiveGanttVBCtl1.CurrentViewObject.TimeLine.StartDate
    dtpEndDate = ActiveGanttVBCtl1.CurrentViewObject.TimeLine.EndDate

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


Private Sub mnuAddRows_Click()
    sRowCaption = ""
    bRowContainer = True
    bRowTitle = False
    fProjectManagement08B.Show 1, Me
    If bRowOK = True Then
        If bRowTitle = True Then
            AddRowTitle "K" & (ActiveGanttVBCtl1.Rows.Count + 1), sRowCaption, bRowContainer
        Else
            AddRowItem "K" & (ActiveGanttVBCtl1.Rows.Count + 1), sRowCaption
        End If
    End If
    ActiveGanttVBCtl1.Redraw
End Sub

Private Sub mnuTaskStyle1_Click()
    mnuTaskStyle1.Checked = True
    mnuTaskStyle2.Checked = False
    mnuTaskStyle3.Checked = False
    mnuTaskStyle4.Checked = False
End Sub

Private Sub mnuTaskStyle2_Click()
    mnuTaskStyle1.Checked = False
    mnuTaskStyle2.Checked = True
    mnuTaskStyle3.Checked = False
    mnuTaskStyle4.Checked = False
End Sub

Private Sub mnuTaskStyle3_Click()
    mnuTaskStyle1.Checked = False
    mnuTaskStyle2.Checked = False
    mnuTaskStyle3.Checked = True
    mnuTaskStyle4.Checked = False
End Sub

Private Sub mnuTaskStyle4_Click()
    mnuTaskStyle1.Checked = False
    mnuTaskStyle2.Checked = False
    mnuTaskStyle3.Checked = False
    mnuTaskStyle4.Checked = True
End Sub
