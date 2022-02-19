VERSION 5.00
Object = "{404F18CB-304C-4658-92D0-119074ED7C75}#1.0#0"; "ActiveGanttVB2.ocx"
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
      Begin ActiveGanttVB.ActiveGanttVBCtl ActiveGanttVBCtl1 
         Height          =   7455
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   13150
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
End
Attribute VB_Name = "fProjectManagement03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTTYPE
  X As Long
  Y As Long
End Type

Private Const PS_DOT = 2
Private Const PS_SOLID = 0

Private Declare Function SelectObject Lib "GDI32.DLL" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTTYPE, ByVal nCount As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private Function NewDate(ByVal Month As Long, ByVal Day As Long, ByVal Year As Long) As Date
    NewDate = DateSerial(Year, Month, Day)
End Function

Private Sub ActiveGanttVBCtl1_CellDraw(CustomDraw As Boolean, ByVal RowIndex As Long, ByVal CellIndex As Long, ByVal lHdc As Long)
    If ActiveGanttVBCtl1.Rows.Item(RowIndex).Height > -1 Then
        CustomDraw = True
        Dim sCaption As String
        Dim sKey As String
        Dim sTag As String
        Dim lTextX As Single
        Dim lTextY As Single
        Dim oFont As New StdFont
        oFont.Name = "Arial"
        oFont.Size = 8
        sCaption = ActiveGanttVBCtl1.Rows.Item(RowIndex).Caption
        sKey = ActiveGanttVBCtl1.Rows.Item(RowIndex).Key
        sTag = ActiveGanttVBCtl1.Rows.Item(RowIndex).Tag
        lTextX = ActiveGanttVBCtl1.Columns.Item(1).Left + (Len(sKey) * 5)
        lTextY = ActiveGanttVBCtl1.Rows.Item(RowIndex).Top + 1
        If bHasChildren(RowIndex) = True Then
            If sTag = "+" Then
                DrawRectangle lHdc, RGB(0, 0, 0), lTextX - 10, ActiveGanttVBCtl1.Rows.Item(RowIndex).Top + 3, 8, 8
                DrawLine lHdc, RGB(0, 0, 0), lTextX - 8, ActiveGanttVBCtl1.Rows.Item(RowIndex).Top + 7, lTextX - 3, _
                ActiveGanttVBCtl1.Rows.Item(RowIndex).Top + 7
                DrawLine lHdc, RGB(0, 0, 0), lTextX - 6, ActiveGanttVBCtl1.Rows.Item(RowIndex).Top + 5, lTextX - 6, _
                ActiveGanttVBCtl1.Rows.Item(RowIndex).Top + 10
            ElseIf sTag = "-" Then
                DrawRectangle lHdc, RGB(0, 0, 0), lTextX - 10, ActiveGanttVBCtl1.Rows.Item(RowIndex).Top + 3, 8, 8
                DrawLine lHdc, RGB(0, 0, 0), lTextX - 8, ActiveGanttVBCtl1.Rows.Item(RowIndex).Top + 7, lTextX - 3, ActiveGanttVBCtl1.Rows.Item(RowIndex).Top + 7
            End If
        End If
        If bIsChild(sKey) = True And sTag = "" Then
            Dim Y As Single
            Dim X1 As Single
            X1 = ActiveGanttVBCtl1.Columns.Item(1).Left + (Len(sKey) * 5)
            Y = CSng(ActiveGanttVBCtl1.Rows.Item(RowIndex).Top + ((ActiveGanttVBCtl1.Rows.Item(RowIndex).Bottom - _
            ActiveGanttVBCtl1.Rows.Item(RowIndex).Top) / 2))
            DrawLine lHdc, RGB(0, 0, 0), X1, Y, X1 - 6, Y
        End If
        DrawString sCaption, lHdc, oFont, RGB(0, 0, 0), lTextX + 3, lTextY
    Else
        CustomDraw = False
    End If
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


Private Sub Form_Load()
    Me.Caption = "The Source Code Store - ActiveGantt Scheduler Control Version " & ActiveGanttVBCtl1.Version & " - Project Management Example"
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Dim lIndex As Long
    With ActiveGanttVBCtl1
        .AutomaticRedraw = False
        .DefaultValues.RowHeight = 20
        .AllowRowSize = False
        .AddMode = AT_TASKADD
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
'            .Add "Columns"
'            Set .Item("Columns").ImageList = imglstColumns
'            .Item("Columns").Appearance = E_STYLEAPPEARANCE.SA_RAISED
'            .Item("Columns").CaptionAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM
'            .Item("Columns").CaptionYMargin = 5
'            .Item("Columns").Font.Bold = True
'            .Item("Columns").PictureAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM
            
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
            .Add "", 30
            .Add "", 255
        End With
        With .Rows
        
        
'    ActiveGanttVBCtl1.DefaultValues.RowHeight = 20
'    ActiveGanttVBCtl1.Columns.Add ""
'    ActiveGanttVBCtl1.Rows.Add "K010", "Row K010", True
'    ActiveGanttVBCtl1.Rows.Item("K010").Tag = "-"
'    ActiveGanttVBCtl1.Rows.Add "K010010", "Row K010010", True
'    ActiveGanttVBCtl1.Rows.Item("K010010").Tag = "-"
'    ActiveGanttVBCtl1.Rows.Add "K010010010", "Row K010010010", True
'    ActiveGanttVBCtl1.Rows.Add "K010020", "Row K010020", True
'    ActiveGanttVBCtl1.Rows.Add "K010030", "Row K010030", True
'    ActiveGanttVBCtl1.Rows.Item("K010030").Tag = "-"
'    ActiveGanttVBCtl1.Rows.Add "K010030010", "Row K010030010", True
'    ActiveGanttVBCtl1.Rows.Add "K010030020", "Row K010030020", True
'    ActiveGanttVBCtl1.Rows.Add "K020", "Row K020", True
'    ActiveGanttVBCtl1.Rows.Item("K020").Tag = "-"
'    ActiveGanttVBCtl1.Rows.Add "K020010", "Row K020010", True
'    ActiveGanttVBCtl1.Rows.Add "K020020", "Row K020020", True
'    ActiveGanttVBCtl1.Rows.Add "K020030", "Row K020030", True
'    ActiveGanttVBCtl1.Rows.Add "K030", "Row K030", True
'    ActiveGanttVBCtl1.Rows.Item("K030").Tag = "-"
'    ActiveGanttVBCtl1.Rows.Add "K030010", "Row K030010", True
'    ActiveGanttVBCtl1.Rows.Add "K030020", "Row K030020", True
'    ActiveGanttVBCtl1.Rows.Add "K030030", "Row K030030", True
'    ActiveGanttVBCtl1.Rows.Add "K030040", "Row K030040", True
        
        
            .Add "K010", "Legacy Systems Integration", False
            .Add "K020", "Implement CGS", False
            .Item("K020").Tag = "-"
            .Add "K020010", "", False
            .Item("K020010").Cells.Item(2).Caption = "Implement Customers/Contact/Leads"
            .Item("K020010").Cells.Item(2).StyleIndex = "ThirdColumn"
            
            .Add "K020020", "Analyse Strategy for Completing CGS Implementation", False
            .Add "K020030", "Implement Customer Financials", False
            .Add "K020040", "Implement CRM", False
            .Add "K020050", "Implement Financials", False
            .Add "K020060", "Implement Customer Records", False
            .Add "K020070", "Advanced CRM Pilot", False
            .Add "K030", "Implement EDI", False
            .Item("K030").Tag = "-"
            .Add "K030010", "Electronic Documents", False
            .Add "K030020", "Analyse Strategy for Completing EDI Implementation", False
            .Add "K030030", "Implement Payroll/401 K", False
            .Add "K030040", "Procurement", False
            .Add "K030050", "Implement Administration", False
            .Add "K030060", "RATTLE Tracking", False
            .Add "K030070", "Implement Employee Health and Safety", False
            .Add "K030080", "Implement Planning", False
            .Add "K030090", "Implement Remaining Records", False
            
            

'            .Item("K10").Cells.Item(3).Caption = "Implement EDI"
'            .Item("K10").Cells.Item(3).StyleIndex = "Title"
'            .Item("K10").Container = False
'            .Add "K11"
'            .Item("K11").Cells.Item(3).Caption = "Electronic Documents"
'            .Add "K12"
'            .Item("K12").Cells.Item(3).Caption = "Analyse Strategy for Completing EDI Implementation"
'            .Add "K13"
'            .Item("K13").Cells.Item(3).Caption = "Implement Payroll/401 K"
'            .Add "K14"
'            .Item("K14").Cells.Item(3).Caption = "Procurement"
'            .Add "K15"
'            .Item("K15").Cells.Item(3).Caption = "Implement Administration"
'            .Add "K16"
'            .Item("K16").Cells.Item(3).Caption = "RATTLE Tracking"
'            .Add "K17"
'            .Item("K17").Cells.Item(3).Caption = "Implement Employee Health and Safety"
'            .Add "K18"
'            .Item("K18").Cells.Item(3).Caption = "Implement Planning"
'            .Add "K19"
'            .Item("K19").Cells.Item(3).Caption = "Implement Remaining Records"
'            .Add "K20"
'            .Item("K20").Cells.Item(3).Caption = "System Environment"
'            .Item("K20").Cells.Item(3).StyleIndex = "Title"
'            .Item("K20").Container = False
'            .Add "K21"
'            .Item("K21").Cells.Item(3).Caption = "Initial Management Support System"
'            .Add "K22"
'            .Item("K22").Cells.Item(3).Caption = "Initial Self Service and Support System"
'            .Add "K23"
'            .Item("K23").Cells.Item(3).Caption = "Complete Management Support System"
'            .Add "K24"
'            .Item("K24").Cells.Item(3).Caption = "Complete Self Service and Support System"
'            .Add "K25"
'            .Item("K25").Cells.Item(3).Caption = "User Support and Documentation"
'            .Item("K25").Cells.Item(3).StyleIndex = "Title"
'            .Add "K26"
'            .Item("K26").Cells.Item(3).Caption = "Systems Infrastructure"
'            .Item("K26").Cells.Item(3).StyleIndex = "Title"
'            For lIndex = 1 To .Count
'                .Item("K" & lIndex).Cells.Item(1).Caption = lIndex
'                .Item("K" & lIndex).Cells.Item(1).StyleIndex = "Cells"
'                .Item("K" & lIndex).Cells.Item(2).StyleIndex = "Cells"
'                If .Item("K" & lIndex).Cells.Item(3).StyleIndex = "0" Then
'                    .Item("K" & lIndex).Cells.Item(3).StyleIndex = "Task"
'                End If
'            Next lIndex
        End With
'        With .Milestones
'            .Add "", "K4", NewDate(4, 15, 2003), "Mil1", "Milestones1"
'            .Item("Mil1").Predecessors.Add "ASCGS", OT_Task
'            .Add "", "K15", NewDate(3, 8, 2003), "Mil2", "Milestones1"
'            .Item("Mil2").Predecessors.Add "ASCEDII", OT_Task
'        End With
        With .Tasks
            .Add "", "K010", NewDate(1, 1, 2003), NewDate(4, 1, 2003), "LSI1", "Tasks3"
            .Add "", "K010", NewDate(4, 15, 2003), NewDate(8, 1, 2003), "LSI2", "Tasks4"
            .Item("LSI2").Predecessors.Add "LSI1", OT_Task
            .Add "", "K020010", NewDate(1, 1, 2003), NewDate(2, 1, 2003), "ICCL", "Tasks1"
            .Add "", "K020020", NewDate(2, 1, 2003), NewDate(4, 1, 2003), "ASCGS", "Tasks1"
            .Item("ASCGS").Predecessors.Add "ICCL", OT_Task
            .Add "", "K020030", NewDate(1, 1, 2003), NewDate(1, 15, 2003), "ICF", "Tasks1"
            .Add "", "K020040", NewDate(1, 15, 2003), NewDate(2, 15, 2003), "ICRM", "Tasks1"
            .Item("ICRM").Predecessors.Add "ICF", OT_Task, , "Predecessors1"
            .Add "", "K020050", NewDate(2, 15, 2003), NewDate(4, 1, 2003), "IF", "Tasks1"
            .Item("IF").Predecessors.Add "ICRM", OT_Task, , "Predecessors1"
            .Add "", "K020060", NewDate(4, 15, 2003), NewDate(5, 1, 2003), "ICR", "Tasks1"
            .Item("ICR").Predecessors.Add "IF", OT_Task, , "Predecessors1"
            .Add "", "K020070", NewDate(5, 1, 2003), NewDate(8, 1, 2003), "ACRMP", "Tasks1"
            .Item("ACRMP").Predecessors.Add "ICR", OT_Task, , "Predecessors1"
'            .Add "", "K11", NewDate(1, 1, 2003), NewDate(2, 11, 2003), "ED", "Tasks1"
'            .Add "", "K12", NewDate(2, 15, 2003), NewDate(3, 1, 2003), "ASCEDII", "Tasks1"
'            .Item("ASCEDII").Predecessors.Add "ED", OT_Task
'            .Add "", "K13", NewDate(3, 11, 2003), NewDate(7, 15, 2003), "IP401K", "Tasks1"
'            .Item("IP401K").Predecessors.Add "ASCEDII", OT_Task
'            .Add "", "K14", NewDate(3, 20, 2003), NewDate(7, 5, 2003), "PROC", "Tasks1"
'            .Item("PROC").Predecessors.Add "ASCEDII", OT_Task
'            .Add "", "K15", NewDate(3, 25, 2003), NewDate(6, 25, 2003), "IA", "Tasks1"
'            .Item("IA").Predecessors.Add "ASCEDII", OT_Task
'
'            .Add "", "K16", NewDate(3, 20, 2003), NewDate(5, 1, 2003), "RATTTRCK", "Tasks1"
'            .Item("RATTTRCK").Predecessors.Add "Mil2", E_OBJECTTYPE.OT_MILESTONE
'            .Add "", "K17", NewDate(3, 20, 2003), NewDate(5, 1, 2003), "IEHS", "Tasks1"
'            .Item("IEHS").Predecessors.Add "Mil2", E_OBJECTTYPE.OT_MILESTONE
'            .Add "", "K18", NewDate(3, 20, 2003), NewDate(5, 1, 2003), "IP", "Tasks1"
'            .Item("IP").Predecessors.Add "Mil2", E_OBJECTTYPE.OT_MILESTONE
'            .Add "", "K19", NewDate(3, 20, 2003), NewDate(5, 1, 2003), "IPRR", "Tasks1"
'            .Item("IPRR").Predecessors.Add "Mil2", E_OBJECTTYPE.OT_MILESTONE
'
'            .Add "", "K16", NewDate(5, 7, 2003), NewDate(7, 15, 2003), "RATTTRCK2", "Tasks2"
'            .Item("RATTTRCK2").Predecessors.Add "RATTTRCK", OT_Task, "Predecessor1"
'            .Add "", "K17", NewDate(5, 8, 2003), NewDate(7, 15, 2003), "IEHS2", "Tasks2"
'            .Item("IEHS2").Predecessors.Add "IEHS", OT_Task, "Predecessor2"
'            .Add "", "K18", NewDate(5, 10, 2003), NewDate(7, 15, 2003), "IP2", "Tasks2"
'            .Item("IP2").Predecessors.Add "IP", OT_Task, "Predecessor3"
'            .Add "", "K19", NewDate(5, 15, 2003), NewDate(7, 15, 2003), "IPRR2", "Tasks2"
'            .Item("IPRR2").Predecessors.Add "IPRR", OT_Task, "Predecessor4"
'            .Add "", "K21", NewDate(12, 1, 2002), NewDate(4, 1, 2003), "IMSS", "Tasks1"
'            .Add "", "K22", NewDate(12, 1, 2002), NewDate(3, 20, 2003), "ISSSS", "Tasks1"
'            .Add "", "K23", NewDate(5, 1, 2003), NewDate(8, 1, 2003), "CMSS", "Tasks1"
'            .Item("CMSS").Predecessors.Add "IMSS", OT_Task, "Predecessor5"
'            .Add "", "K24", NewDate(6, 1, 2003), NewDate(8, 1, 2003), "CSSSS", "Tasks1"
'            .Item("CSSSS").Predecessors.Add "ISSSS", OT_Task, "Predecessor6"
'            .Add "", "K25", NewDate(1, 1, 2003), NewDate(8, 1, 2003), , "Tasks3"
'            .Add "", "K26", NewDate(1, 1, 2003), NewDate(7, 5, 2003), , "Tasks3"
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

Private Function IsSibling(ByVal sMasterKey As String, ByVal sKey As String) As Boolean
    Dim sSiblingID As String
    If Len(sMasterKey) = 4 Then
        sSiblingID = "K"
    Else
        sSiblingID = Left(sMasterKey, Len(sMasterKey) - 3)
    End If
    If (Len(sMasterKey) = Len(sKey)) Then
        If (Left(sKey, Len(sSiblingID)) = sSiblingID) Then
            IsSibling = True
        Else
            IsSibling = False
        End If
    Else
        IsSibling = False
    End If
End Function

Private Function bHasChildren(ByVal Index As Integer) As Boolean
    Dim sMasterKey As String
    Dim sKey As String
    If Index >= ActiveGanttVBCtl1.Rows.Count Then
        bHasChildren = False
        Exit Function
    End If
    sMasterKey = ActiveGanttVBCtl1.Rows.Item(Index).Key
    Index = Index + 1
    sKey = ActiveGanttVBCtl1.Rows.Item(Index).Key
    If Len(sKey) < Len(sMasterKey) Then
        bHasChildren = False
        Exit Function
    End If
    If Left(sKey, Len(sMasterKey)) = sMasterKey Then
        bHasChildren = True
        Exit Function
    End If
End Function
    
Private Function bIsChild(ByVal Key As String) As Boolean
    If Len(Key) = 4 Then
        bIsChild = False
    Else
        bIsChild = True
    End If
End Function
    
Private Sub HideChildren(ByVal sKey As String)
    Dim i As Integer
    Dim sChildKey As String
    For i = 1 To ActiveGanttVBCtl1.Rows.Count
        sChildKey = ActiveGanttVBCtl1.Rows.Item(i).Key
        If Len(sChildKey) > Len(sKey) Then
            If (Left(sChildKey, Len(sKey)) = sKey) And (Len(sChildKey) > _
            Len(sKey)) Then
                ActiveGanttVBCtl1.Rows.Item(i).Height = -1
            End If
        End If
    Next i
End Sub
    
Private Sub ShowChildren(ByVal sKey As String)
    Dim i As Integer
    Dim sChildKey As String
    For i = 1 To ActiveGanttVBCtl1.Rows.Count
        sChildKey = ActiveGanttVBCtl1.Rows.Item(i).Key
        If Len(sChildKey) > Len(sKey) Then
            If (Left(sChildKey, Len(sKey)) = sKey) And (Len(sChildKey) > _
            Len(sKey)) Then
                ActiveGanttVBCtl1.Rows.Item(i).Height = 20
                If ActiveGanttVBCtl1.Rows.Item(i).Tag = "+" Then
                    ActiveGanttVBCtl1.Rows.Item(i).Tag = "-"
                End If
            End If
        End If
    Next i
End Sub
    
Private Sub ActiveGanttVBCtl1_RowClick(ByVal Index As Long, ByVal X As Single, ByVal Y As Single, ByVal Button As Integer)
    Dim sTag As String
    If bHasChildren(Index) = True Then
        sTag = ActiveGanttVBCtl1.Rows.Item(Index).Tag
        If sTag = "+" Then
            ActiveGanttVBCtl1.Rows.Item(Index).Tag = "-"
            ShowChildren (ActiveGanttVBCtl1.Rows.Item(Index).Key)
        ElseIf sTag = "-" Then
            ActiveGanttVBCtl1.Rows.Item(Index).Tag = "+"
            HideChildren (ActiveGanttVBCtl1.Rows.Item(Index).Key)
        End If
        ActiveGanttVBCtl1.Redraw
    End If
End Sub

Private Sub ActiveGanttVBCtl1_TableDraw(ByVal lHdc As Long)
    Dim i As Integer
    Dim k As Integer
    Dim sMasterKey As String
    Dim sKey As String
    Dim lTextX As Single
    Dim lTop As Single
    Dim lBottom As Single
    For i = 1 To ActiveGanttVBCtl1.Rows.Count
        sMasterKey = ActiveGanttVBCtl1.Rows.Item(i).Key
        If bHasChildren(i) = True And _
        ActiveGanttVBCtl1.Rows.Item(i).Height > -1 And _
        ActiveGanttVBCtl1.Rows.Item(i).Tag = "-" Then
            lTextX = ActiveGanttVBCtl1.Columns.Item(1).Left + _
            ((Len(sMasterKey) + 3) * 5) - 6
            lTop = CSng(ActiveGanttVBCtl1.Rows.Item(i).Top + _
            ((ActiveGanttVBCtl1.Rows.Item(i).Bottom - _
            ActiveGanttVBCtl1.Rows.Item(i).Top) / 2)) + 5
            If bHasChildren(i + 1) = False Then
                lBottom = CSng(ActiveGanttVBCtl1.Rows.Item(i + 1).Top + _
                ((ActiveGanttVBCtl1.Rows.Item(i + 1).Bottom - _
                ActiveGanttVBCtl1.Rows.Item(i + 1).Top) / 2))
            Else
                lBottom = CSng(ActiveGanttVBCtl1.Rows.Item(i + 1).Top) + 4
            End If
            DrawLine lHdc, RGB(0, 0, 0), lTextX, lTop, lTextX, lBottom
        End If
        For k = i + 1 To ActiveGanttVBCtl1.Rows.Count
            sKey = ActiveGanttVBCtl1.Rows.Item(k).Key
            If IsSibling(sMasterKey, sKey) = True Then
                If ActiveGanttVBCtl1.Rows.Item(i).Height > -1 And _
                ActiveGanttVBCtl1.Rows.Item(k).Height > -1 Then
                    lTextX = ActiveGanttVBCtl1.Columns.Item(1).Left + _
                    (Len(sKey) * 5) - 6
                    If (ActiveGanttVBCtl1.Rows.Item(i).Tag = "") Then
                        lTop = CSng(ActiveGanttVBCtl1.Rows.Item(i).Top + _
                        ((ActiveGanttVBCtl1.Rows.Item(i).Bottom - _
                        ActiveGanttVBCtl1.Rows.Item(i).Top) / 2))
                    Else
                        lTop = CSng(ActiveGanttVBCtl1.Rows.Item(i).Top + _
                        ((ActiveGanttVBCtl1.Rows.Item(i).Bottom - _
                        ActiveGanttVBCtl1.Rows.Item(i).Top) / 2)) + 1
                    End If
                    If (ActiveGanttVBCtl1.Rows.Item(k).Tag = "") Then
                        lBottom = CSng(ActiveGanttVBCtl1.Rows.Item(k).Top + _
                        ((ActiveGanttVBCtl1.Rows.Item(k).Bottom - _
                        ActiveGanttVBCtl1.Rows.Item(k).Top) / 2)) + 1
                    Else
                        lBottom = CSng(ActiveGanttVBCtl1.Rows.Item(k).Top) + 3
                    End If
                    DrawLine lHdc, RGB(0, 0, 0), lTextX, lTop, lTextX, lBottom
                End If
                Exit For
            End If
        Next k
    Next i
End Sub

Private Sub DrawRectangle(ByVal hdc As Long, ByVal lColor As OLE_COLOR, ByVal v_X1 As Long, ByVal v_Y1 As Long, ByVal Width As Long, ByVal Height As Long)
    Dim hPen As Long
    Dim HoldPen As Long
    Dim Points() As POINTTYPE
    hPen = CreatePen(PS_SOLID, 1, lColor)
    HoldPen = SelectObject(hdc, hPen)
    ReDim Points(4)
    Points(0).X = v_X1
    Points(0).Y = v_Y1
    Points(1).X = v_X1 + Width
    Points(1).Y = v_Y1
    Points(2).X = v_X1 + Width
    Points(2).Y = v_Y1 + Height
    Points(3).X = v_X1
    Points(3).Y = v_Y1 + Height
    Points(4).X = v_X1
    Points(4).Y = v_Y1
    Polyline hdc, Points(0), 5
    SelectObject hdc, HoldPen
    DeleteObject hPen
End Sub

Private Sub DrawLine(ByVal hdc As Long, ByVal lColor As OLE_COLOR, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
    Dim hPen As Long
    Dim HoldPen As Long
    Dim Points() As POINTTYPE
    hPen = CreatePen(PS_SOLID, 1, lColor)
    HoldPen = SelectObject(hdc, hPen)
    ReDim Points(1)
    Points(0).X = X1
    Points(0).Y = Y1
    Points(1).X = X2
    Points(1).Y = Y2
    Polyline hdc, Points(0), 2
    SelectObject hdc, HoldPen
    DeleteObject hPen
End Sub

Private Sub DrawString(ByVal sCaption As String, ByVal lHdc As Long, ByRef oFont As StdFont, ByVal lColor As OLE_COLOR, ByVal X1 As Long, ByVal Y1 As Long)
    Dim holdFont As Long
    Dim FontI As IFont
    Set FontI = oFont
    holdFont = SelectObject(lHdc, FontI.hFont)
    SetTextColor lHdc, lColor
    TextOut lHdc, X1, Y1, sCaption, Len(sCaption)
    SelectObject lHdc, holdFont
End Sub


