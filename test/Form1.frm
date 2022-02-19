VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7860
   ClientLeft      =   1410
   ClientTop       =   495
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   10035
   Begin VB.PictureBox ActiveGanttVBCtl1 
      Height          =   3855
      Left            =   240
      ScaleHeight     =   3795
      ScaleWidth      =   6195
      TabIndex        =   15
      Top             =   120
      Width           =   6255
   End
   Begin VB.CommandButton cmdTestForm 
      Caption         =   "Test Form"
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   7080
      Width           =   2415
   End
   Begin VB.Frame fraClassesTest 
      Caption         =   "Classes"
      Height          =   2535
      Left            =   6840
      TabIndex        =   11
      Top             =   2160
      Width           =   3135
      Begin VB.TextBox txtRandomClassTesting 
         Height          =   285
         Left            =   2160
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdRandomClassTesting 
         Caption         =   "Random Class Testing"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.TextBox txtRandomFastTesting 
      Height          =   285
      Left            =   3240
      TabIndex        =   10
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton cmdRandomFastTesting 
      Caption         =   "Random Fast Testing"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox txtDisplayMessage 
      Height          =   2055
      Left            =   6480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   5640
      Width           =   3495
   End
   Begin VB.CommandButton cmdUserInterfaceTesting 
      Caption         =   "UserInterface Testing"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   6120
      Width           =   2295
   End
   Begin VB.TextBox txtCrazyTesting 
      Height          =   285
      Left            =   3240
      TabIndex        =   6
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton cmdCrazyTesting 
      Caption         =   "Crazy Testing"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   5760
      Width           =   2295
   End
   Begin VB.TextBox txtTasksCount 
      Height          =   285
      Left            =   9000
      TabIndex        =   3
      Text            =   "0"
      Top             =   240
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   120
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   17
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":03C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":078C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B52
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   4440
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
            Picture         =   "Form1.frx":0F18
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A70
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E38
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2200
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":25C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2990
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2D58
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3120
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":34E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":38B0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtRandomTesting 
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdRandomTesting 
      Caption         =   "Random Testing"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton cmdCustomTesting 
      Caption         =   "Custom Testing"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label lblTasksCount 
      Caption         =   "Tasks Count:"
      Height          =   255
      Left            =   7200
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mp_lRowIndex As Long
Private mp_bSeed As Boolean
Private mp_bDisplayMessage As Boolean
Private mp_lStressFactor As Long


Private Sub cmdCrazyTesting_Click()
    mp_bSeed = True
    mp_bDisplayMessage = False
    ActiveGanttVBCtl1.ErrorReports = RE_RAISE
    Dim lIterationNumber As Long
    For lIterationNumber = 0 To 200000
        txtCrazyTesting.Text = lIterationNumber
        txtCrazyTesting.Refresh
        mp_RunTest
    Next lIterationNumber
End Sub

Private Sub cmdCustomTesting_Click()
    ActiveGanttVBCtl1.ErrorReports = RE_RAISE
    ActiveGanttVBCtl1.Rows.Add "K1"
    ActiveGanttVBCtl1.Rows.Add "K1"
End Sub

Private Sub cmdRandomClassTesting_Click()
    Dim fStart As Single
    fStart = Timer
    mp_bSeed = False
    mp_bDisplayMessage = False
    ActiveGanttVBCtl1.ErrorReports = RE_RAISE
    Dim lIterationNumber As Long
    For lIterationNumber = 0 To 125
        txtRandomClassTesting.Text = lIterationNumber
        txtRandomClassTesting.Refresh
        If txtRandomClassTesting.Text = "71" Then
        End If
        If lIterationNumber > 124 Then
          Debug.Print
        End If
        mp_RunTest
    Next lIterationNumber
    txtDisplayMessage.Text = Timer - fStart
End Sub

Private Sub cmdRandomFastTesting_Click()
    mp_bSeed = False
    mp_bDisplayMessage = False
    ActiveGanttVBCtl1.ErrorReports = RE_RAISE
    ActiveGanttVBCtl1.AutomaticRedraw = False
    Dim lIterationNumber As Long
    For lIterationNumber = 0 To 200000
        txtRandomFastTesting.Text = lIterationNumber
        txtRandomFastTesting.Refresh
        If txtRandomFastTesting.Text = "7399" Then
            Debug.Print
        End If
        mp_RunTest
    Next lIterationNumber
End Sub

Private Sub cmdRandomTesting_Click()
    mp_bSeed = False
    mp_bDisplayMessage = False
    ActiveGanttVBCtl1.ErrorReports = RE_RAISE
    Dim lIterationNumber As Long
    For lIterationNumber = 0 To 200000
        txtRandomTesting.Text = lIterationNumber
        txtRandomTesting.Refresh
        If txtRandomTesting.Text = "7399" Then
            Debug.Print
        End If
        mp_RunTest
    Next lIterationNumber
End Sub

Public Function RndLong(ByVal lLowerBound As Long, ByVal lUpperBound As Long) As Long
    If mp_bSeed = True Then
        Randomize
    End If
    RndLong = Int((lUpperBound - lLowerBound + 1) * Rnd + lLowerBound)
End Function

Public Function RndAbsoluteImage() As String
    Select Case RndLong(1, 12)
        Case 1
            RndAbsoluteImage = App.Path & "\images\proj1start.bmp"
        Case 2
            RndAbsoluteImage = App.Path & "\images\proj1Middle.bmp"
        Case 3
            RndAbsoluteImage = App.Path & "\images\proj1End.bmp"
        Case 4
            RndAbsoluteImage = App.Path & "\images\proj2start.bmp"
        Case 5
            RndAbsoluteImage = App.Path & "\images\proj2Middle.bmp"
        Case 6
            RndAbsoluteImage = App.Path & "\images\proj2End.bmp"
        Case 7
            RndAbsoluteImage = App.Path & "\images\proj3start.bmp"
        Case 8
            RndAbsoluteImage = App.Path & "\images\proj3Middle.bmp"
        Case 9
            RndAbsoluteImage = App.Path & "\images\proj3End.bmp"
        Case 10
            RndAbsoluteImage = App.Path & "\images\proj4start.bmp"
        Case 11
            RndAbsoluteImage = App.Path & "\images\proj4Middle.bmp"
        Case 12
            RndAbsoluteImage = App.Path & "\images\proj4End.bmp"
    End Select
End Function

Public Function RndBool() As Boolean
    If mp_bSeed = True Then
        Randomize
    End If
    Dim LIndex As Long
    LIndex = Int((1 - 0 + 1) * Rnd + 0)
    If LIndex = 0 Then
        RndBool = False
    Else
        RndBool = True
    End If
End Function

Public Function RndString(Optional minLength As Long = 0) As String
    Dim lStringLength As Long
    Dim sString As String
    lStringLength = RndLong(minLength, 150)
    Do While Len(sString) < lStringLength
        sString = sString & Chr$(RndLong(33, 126))
    Loop
    RndString = sString
End Function

Public Function RndStyle() As String
    If ActiveGanttVBCtl1.Styles.Count <> 0 Then
        If RndBool = True Then
            RndStyle = ActiveGanttVBCtl1.Styles.Item(RndLong(1, ActiveGanttVBCtl1.Styles.Count)).Key
        Else
            RndStyle = RndLong(1, ActiveGanttVBCtl1.Styles.Count)
        End If
    Else
        RndStyle = "0"
    End If
End Function

Public Function RndObjectKey(ByRef oCollection As Object) As String
    Dim LIndex As Long
    If oCollection.Count <> 0 Then
        LIndex = RndLong(1, oCollection.Count)
        If RndBool = True Then
            If oCollection.Item(LIndex).Key <> "" Then
                RndObjectKey = oCollection.Item(LIndex).Key
            Else
                RndObjectKey = LIndex
            End If
        Else
            RndObjectKey = LIndex
        End If
    Else
        RndObjectKey = ""
    End If
End Function

Public Function RndRowKey() As String
    If ActiveGanttVBCtl1.Rows.Count <> 0 Then
        RndRowKey = ActiveGanttVBCtl1.Rows.Item(RndLong(1, ActiveGanttVBCtl1.Rows.Count)).Key
    Else
        RndRowKey = ""
    End If
End Function

Public Function RndTaskKey() As String
    If ActiveGanttVBCtl1.Tasks.Count <> 0 Then
        RndTaskKey = ActiveGanttVBCtl1.Tasks.Item(RndLong(1, ActiveGanttVBCtl1.Tasks.Count)).Key
    Else
        RndTaskKey = ""
    End If
End Function

Public Function RndPercentageGroupKey() As String
    If ActiveGanttVBCtl1.PercentageGroups.Count <> 0 Then
        RndPercentageGroupKey = ActiveGanttVBCtl1.PercentageGroups.Item(RndLong(1, ActiveGanttVBCtl1.PercentageGroups.Count)).Key
    Else
        RndPercentageGroupKey = ""
    End If
End Function

Public Function RndDateInterval() As String
    Dim lParam As Long
    lParam = RndLong(1, 10)
    Select Case lParam
        Case 1
            RndDateInterval = "yyyy"
        Case 2
            RndDateInterval = "q"
        Case 3
            RndDateInterval = "m"
        Case 4
            RndDateInterval = "y"
        Case 5
            RndDateInterval = "d"
        Case 6
            RndDateInterval = "w"
        Case 7
            RndDateInterval = "ww"
        Case 8
            RndDateInterval = "h"
        Case 9
            RndDateInterval = "n"
        Case 10
            RndDateInterval = "s"
    End Select
End Function

Public Function RndLongInterval(ByVal sInterval As String) As Long
    Select Case sInterval
        Case "yyyy"
            RndLongInterval = RndLong(1, 5)
        Case "q"
            RndLongInterval = RndLong(1, 20)
        Case "m"
            RndLongInterval = RndLong(1, 60)
        Case "y"
            RndLongInterval = RndLong(1, 900)
        Case "d"
            RndLongInterval = RndLong(1, 900)
        Case "w"
            RndLongInterval = RndLong(1, 900)
        Case "ww"
            RndLongInterval = RndLong(1, 200)
        Case "h"
            RndLongInterval = RndLong(1, 2000)
        Case "n"
            RndLongInterval = RndLong(1, 5000)
        Case "s"
            RndLongInterval = RndLong(1, 10000)
    End Select
End Function

Private Function RndColor() As Long
    RndColor = RGB(RndLong(0, 255), RndLong(0, 255), RndLong(0, 255))
End Function

Private Sub mp_RunTest()
    Dim lTestIndex As Long
    lTestIndex = RndLong(1, 24)
    Select Case lTestIndex
        Case 1
            mp_AddRow
        Case 2
            mp_RemoveRow
        Case 3
            mp_ClearRows
        Case 4
            mp_AddColumn
        Case 5
            mp_RemoveColumn
        Case 6
            mp_ClearColumns
        Case 7
            mp_AddTask
        Case 8
            mp_RemoveTask
        Case 9
            mp_ClearTasks
        Case 10
            mp_AddStyle
        Case 11
            mp_RemoveStyle
        Case 12
            mp_ClearStyles
        Case 13
            mp_SetCellCaption
        Case 14
            mp_SetCellStyleIndex
        Case 15
            mp_ZoomFactor
        Case 16
            mp_SplitterPosition
        Case 17
            mp_HorizontalScrollBar
        Case 18
            mp_PositionTimeLine
        Case 19
            mp_AllowFixedColumnSize
        Case 20
            mp_AllowTaskAdd
        Case 21
            mp_AllowTaskEdit
        Case 22
            mp_AllowColumnSize
        Case 23
            mp_AllowRowSize
        Case 24
            mp_AllowRowSwap
    End Select
End Sub

Private Sub mp_AddRow()
    Dim lIterationIndex As Long
    Dim sPicture As String
    Dim sRndStyle As String
    sPicture = ""
    sRndStyle = "0"
    sRndStyle = RndStyle
    If sRndStyle <> "0" Then
        If Not (ActiveGanttVBCtl1.Styles.Item(sRndStyle).ImageList Is Nothing) Then
            sPicture = RndLong(1, 4)
        End If
    Else
        If RndBool = True Then
            sPicture = RndAbsoluteImage
        End If
    End If
    For lIterationIndex = 0 To mp_lStressFactor
        mp_lRowIndex = mp_lRowIndex + 1
        ActiveGanttVBCtl1.Rows.Add "K" & mp_lRowIndex, "Row: " & mp_lRowIndex, RndBool, RndBool, sRndStyle, sPicture
    Next lIterationIndex
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "Rows Added"
    End If
End Sub

Private Sub mp_RemoveRow()
    If ActiveGanttVBCtl1.Rows.Count <> 0 Then
        If RndBool = True Then
            ActiveGanttVBCtl1.Rows.Remove RndLong(1, ActiveGanttVBCtl1.Rows.Count)
        Else
            ActiveGanttVBCtl1.Rows.Remove RndRowKey
        End If
    End If
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "Row Removed"
    End If
End Sub

Private Sub mp_ClearRows()
    Dim LIndex As Long
    LIndex = RndLong(0, 10)
    If LIndex > 7 Then
        ActiveGanttVBCtl1.Rows.Clear
        If mp_bDisplayMessage = True Then
            txtDisplayMessage.Text = "Rows Cleared"
        End If
    Else
        If mp_bDisplayMessage = True Then
            txtDisplayMessage.Text = "Tried to Clear Rows"
        End If
    End If
End Sub

Private Sub mp_AddColumn()
    Dim sPicture As String
    Dim sRndStyle As String
    sPicture = ""
    sRndStyle = "0"
    sRndStyle = RndStyle
    If sRndStyle <> "0" Then
        If Not (ActiveGanttVBCtl1.Styles.Item(sRndStyle).ImageList Is Nothing) Then
            sPicture = RndLong(1, 4)
        End If
    Else
        If RndBool = True Then
            sPicture = RndAbsoluteImage
        End If
    End If
    ActiveGanttVBCtl1.Columns.Add RndString, RndLong(0, 115), sRndStyle, sPicture
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "Column Added"
    End If
End Sub

Private Sub mp_RemoveColumn()
    If ActiveGanttVBCtl1.Columns.Count <> 1 Then
        ActiveGanttVBCtl1.Columns.Remove RndLong(1, ActiveGanttVBCtl1.Columns.Count)
    End If
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "Columns Removed"
    End If
End Sub

Private Sub mp_ClearColumns()
    ActiveGanttVBCtl1.Columns.Clear
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "Columns Cleared"
    End If
End Sub

Private Sub mp_AddTask()
    Dim lTestIterator As Long
    Dim sPicture As String
    Dim sRndStyle As String
    Dim sDateInterval As String
    Dim sRowKey
    Dim dtStart As Date
    Dim dtEnd As Date
    For lTestIterator = 0 To mp_lStressFactor
        sPicture = ""
        sRndStyle = "0"
        sRndStyle = RndStyle
        If sRndStyle <> "0" Then
            If Not (ActiveGanttVBCtl1.Styles.Item(sRndStyle).ImageList Is Nothing) Then
                sPicture = RndLong(1, 4)
            End If
        Else
            If RndBool = True Then
                sPicture = RndAbsoluteImage
            End If
        End If
        sDateInterval = RndDateInterval
        If RndBool = True Then
            dtStart = DateAdd(sDateInterval, RndLongInterval(sDateInterval), Now)
        Else
            dtStart = DateAdd(sDateInterval, -RndLongInterval(sDateInterval), Now)
        End If
        sDateInterval = RndDateInterval
        dtEnd = DateAdd(sDateInterval, RndLongInterval(sDateInterval), dtStart)
        sRowKey = RndRowKey
        If sRowKey <> "" Then
            If ActiveGanttVBCtl1.InConflict(dtStart, dtEnd, sRowKey) = False Then
                ActiveGanttVBCtl1.Tasks.Add RndString, sRowKey, dtStart, dtEnd, RndString(7), sRndStyle, sPicture
            End If
        End If
    Next lTestIterator
    If ActiveGanttVBCtl1.Tasks.Count > CLng(txtTasksCount.Text) Then
        txtTasksCount.Text = ActiveGanttVBCtl1.Tasks.Count
        txtTasksCount.Refresh
    End If
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "Tasks Added"
    End If
End Sub

Private Sub mp_ClearTasks()
    ActiveGanttVBCtl1.Tasks.Clear
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "Tasks Cleared"
    End If
End Sub

Private Sub mp_AddStyle()
    Dim lIterator As Long
    Dim lAppearance As Long
    For lIterator = 0 To mp_lStressFactor
        ActiveGanttVBCtl1.Styles.Add RndString(10)
        lAppearance = RndLong(0, 4)
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).Appearance = lAppearance
        If lAppearance = 3 Then
            Dim lGanttType As Long
            lGanttType = RndLong(1, 4)
            Set ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).ImageList = ImageList1
            Select Case lGanttType
                Case 1
                    ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).StartPictureIndex = 1
                    ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).MiddlePictureIndex = 2
                    ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).EndPictureIndex = 3
                Case 2
                    ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).StartPictureIndex = 4
                    ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).MiddlePictureIndex = 5
                    ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).EndPictureIndex = 6
                Case 3
                    ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).StartPictureIndex = 7
                    ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).MiddlePictureIndex = 8
                    ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).EndPictureIndex = 9
                Case 4
                    ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).StartPictureIndex = 10
                    ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).MiddlePictureIndex = 11
                    ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).EndPictureIndex = 12
            End Select
        Else
            If RndBool = True Then
                Set ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).ImageList = ImageList2
            End If
        End If
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).CaptionAlignmentHorizontal = RndLong(1, 3)
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).CaptionAlignmentVertical = RndLong(1, 3)
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).PictureAlignmentHorizontal = RndLong(1, 3)
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).PictureAlignmentVertical = RndLong(1, 3)
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).CaptionXMargin = RndLong(0, 15)
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).CaptionYMargin = RndLong(0, 15)
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).PictureXMargin = RndLong(0, 15)
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).PictureYMargin = RndLong(0, 15)
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).CaptionVisible = RndBool
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).ClipCaption = RndBool
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).UseMask = RndBool
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).SelectionRectangleOffsetLeft = RndLong(0, 15)
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).SelectionRectangleOffsetTop = RndLong(0, 15)
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).SelectionRectangleOffsetRight = RndLong(0, 15)
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).SelectionRectangleOffsetBottom = RndLong(0, 15)
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).BackColor = RndColor
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).BorderColor = RndColor
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).ForeColor = RndColor
        ActiveGanttVBCtl1.Styles.Item(ActiveGanttVBCtl1.Styles.Count).BorderStyle = RndLong(0, 1)
    Next lIterator
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "Styles Added"
    End If
End Sub

Private Sub mp_RemoveStyle()
    If ActiveGanttVBCtl1.Styles.Count <> 0 Then
        If RndBool = True Then
            ActiveGanttVBCtl1.Styles.Remove RndLong(1, ActiveGanttVBCtl1.Styles.Count)
        Else
            ActiveGanttVBCtl1.Styles.Remove RndStyle
        End If
    End If
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "Style Removed"
    End If
End Sub

Private Sub mp_ClearStyles()
    Dim LIndex As Long
    LIndex = RndLong(0, 10)
    If LIndex > 7 Then
        ActiveGanttVBCtl1.Styles.Clear
        If mp_bDisplayMessage = True Then
            txtDisplayMessage.Text = "Styles Cleared"
        End If
    Else
        If mp_bDisplayMessage = True Then
            txtDisplayMessage.Text = "Tried to clear Styles"
        End If
    End If
End Sub

Private Sub mp_SetCellCaption()
    If ActiveGanttVBCtl1.Rows.Count <> 0 Then
        ActiveGanttVBCtl1.Rows.Item(RndLong(1, ActiveGanttVBCtl1.Rows.Count)).Cell(RndLong(1, ActiveGanttVBCtl1.Columns.Count)).Caption = RndString
    End If
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "CellCaption Changed"
    End If
End Sub

Private Sub mp_SetCellStyleIndex()
    If ActiveGanttVBCtl1.Rows.Count <> 0 Then
        ActiveGanttVBCtl1.Rows.Item(RndLong(1, ActiveGanttVBCtl1.Rows.Count)).Cell(RndLong(1, ActiveGanttVBCtl1.Columns.Count)).StyleIndex = RndStyle
    End If
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "CellStyleIndex Changed"
    End If
End Sub

Private Sub mp_ZoomFactor()
    ActiveGanttVBCtl1.ZoomFactor = RndLong(0, 13)
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "ZoomFactor Changed: " & ActiveGanttVBCtl1.ZoomFactor
    End If
End Sub

Private Sub mp_SplitterPosition()
    ActiveGanttVBCtl1.SplitterPosition = RndLong(0, 10000)
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "SplitterPosition Changed: " & ActiveGanttVBCtl1.SplitterPosition
    End If
End Sub

Private Sub cmdTestForm_Click()
    fTestForm.Show 1, Me
End Sub

Private Sub cmdUserInterfaceTesting_Click()
    Screen.MousePointer = 11
    mp_bSeed = True
    mp_bDisplayMessage = True
    ActiveGanttVBCtl1.ErrorReports = RE_RAISE
    mp_RunTest
    Screen.MousePointer = 0
End Sub

Private Sub mp_HorizontalScrollBar()
    Dim lMax As Long
    Dim lSmall As Long
    Dim lLarge As Long
    Dim sInterval As String
    If ActiveGanttVBCtl1.HorizontalScrollBarEnabled = False Then
        lMax = RndLong(1000, 10000)
        lSmall = RndLong(1, lMax / 10)
        lLarge = RndLong(lSmall, lMax / 5)
        sInterval = RndDateInterval
        ActiveGanttVBCtl1.HorizontalScrollBarInterval = RndDateInterval
        ActiveGanttVBCtl1.HorizontalScrollBarMax = lMax
        ActiveGanttVBCtl1.HorizontalScrollBarSmallChange = lSmall
        ActiveGanttVBCtl1.HorizontalScrollBarLargeChange = lLarge
        ActiveGanttVBCtl1.HorizontalScrollBarStart = DateAdd(sInterval, RndLongInterval(sInterval), Now)
        ActiveGanttVBCtl1.HorizontalScrollBarEnabled = True
        If mp_bDisplayMessage = True Then
            txtDisplayMessage.Text = "Horizontal Scroll bar enabled: "
        End If
    Else
        ActiveGanttVBCtl1.HorizontalScrollBarEnabled = False
        If mp_bDisplayMessage = True Then
            txtDisplayMessage.Text = "Horizontal Scroll bar disabled: "
        End If
    End If
End Sub

Private Sub mp_PositionTimeLine()
    Dim sInterval As String
    Dim dtDate As Date
    sInterval = RndDateInterval
    dtDate = DateAdd(sInterval, RndLongInterval(sInterval), Now)
    ActiveGanttVBCtl1.PositionTimeLine dtDate
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "PositionTimeLine: " & dtDate
    End If
End Sub

'        Case 19 'AllowFixedColumnSize
'        Case 20 'AllowTaskAdd
'        Case 21 'AllowTaskEdit
'        Case 22 'AllowColumnSize
'        Case 23 'AllowRowSize
'        Case 24 'AllowRowSwap

Private Sub mp_AllowFixedColumnSize()
    ActiveGanttVBCtl1.AllowFixedColumnSize = Not ActiveGanttVBCtl1.AllowFixedColumnSize
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "AllowFixedColumnSize: " & ActiveGanttVBCtl1.AllowFixedColumnSize
    End If
End Sub

Private Sub mp_AllowTaskAdd()
    ActiveGanttVBCtl1.AllowAdd = Not ActiveGanttVBCtl1.AllowAdd
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "AllowTaskAdd: " & ActiveGanttVBCtl1.AllowAdd
    End If
End Sub

Private Sub mp_AllowTaskEdit()
    ActiveGanttVBCtl1.AllowEdit = Not ActiveGanttVBCtl1.AllowEdit
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "AllowTaskEdit: " & ActiveGanttVBCtl1.AllowEdit
    End If
End Sub

Private Sub mp_AllowColumnSize()
    ActiveGanttVBCtl1.AllowColumnSize = Not ActiveGanttVBCtl1.AllowColumnSize
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "AllowColumnSize: " & ActiveGanttVBCtl1.AllowColumnSize
    End If
End Sub

Private Sub mp_AllowRowSwap()
    ActiveGanttVBCtl1.AllowRowSwap = Not ActiveGanttVBCtl1.AllowRowSwap
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "AllowRowSwap: " & ActiveGanttVBCtl1.AllowRowSwap
    End If
End Sub

Private Sub mp_AllowRowSize()
    ActiveGanttVBCtl1.AllowRowSize = Not ActiveGanttVBCtl1.AllowRowSize
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "AllowRowSize: " & ActiveGanttVBCtl1.AllowRowSize
    End If
End Sub


Private Sub mp_AddPercentageGroup()
    ActiveGanttVBCtl1.PercentageGroups.Add RndString, RndBool, RndStyle
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "PercentageGroup Added"
    End If
End Sub

Private Sub mp_RemovePercentageGroup()
    If ActiveGanttVBCtl1.PercentageGroups.Count <> 0 Then
        ActiveGanttVBCtl1.PercentageGroups.Remove RndObjectKey(ActiveGanttVBCtl1.PercentageGroups)
    End If
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "PercentageGroup Removed"
    End If
End Sub

Private Sub mp_ClearPercentageGroups()
    ActiveGanttVBCtl1.PercentageGroups.Clear
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "PercentageGroups Cleared"
    End If
End Sub

Private Sub mp_AddLayer()
    ActiveGanttVBCtl1.Layers.Add RndString, True
End Sub

Private Sub mp_RemoveLayer()
    If ActiveGanttVBCtl1.Layers.Count <> 0 Then
        ActiveGanttVBCtl1.Layers.Remove RndObjectKey(ActiveGanttVBCtl1.Layers)
    End If
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "Layer Removed"
    End If
End Sub

Private Sub mp_ClearLayers()
    ActiveGanttVBCtl1.Layers.Clear
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "Layers Cleared"
    End If
End Sub

Private Sub mp_AddPercentage()
    If ActiveGanttVBCtl1.PercentageGroups.Count <> 0 And ActiveGanttVBCtl1.Tasks.Count <> 0 Then
        ActiveGanttVBCtl1.Percentages.Add RndTaskKey, RndPercentageGroupKey, RndLong(0, 100) / 100, RndString
    End If
End Sub

Private Sub mp_RemovePercentage()
    If ActiveGanttVBCtl1.Percentages.Count <> 0 Then
        ActiveGanttVBCtl1.Percentages.Remove RndObjectKey(ActiveGanttVBCtl1.Percentages)
    End If
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "Percentage Removed"
    End If
End Sub

Private Sub mp_ClearPercentages()
    ActiveGanttVBCtl1.Percentages.Clear
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "Percentages Cleared"
    End If
End Sub

Private Sub mp_AddMilestone()
    Dim sDateInterval As String
    sDateInterval = RndDateInterval
    Dim sPicture As String
    Dim sRndStyle As String
    sPicture = ""
    sRndStyle = "0"
    sRndStyle = RndStyle
    If sRndStyle <> "0" Then
        If Not (ActiveGanttVBCtl1.Styles.Item(sRndStyle).ImageList Is Nothing) Then
            sPicture = RndLong(1, 4)
        End If
    Else
        If RndBool = True Then
            sPicture = RndAbsoluteImage
        End If
    End If
    ActiveGanttVBCtl1.Milestones.Add RndString, RndRowKey, DateAdd(sDateInterval, RndLongInterval(sDateInterval), Now), RndString, RndStyle, sPicture
End Sub

Private Sub mp_RemoveMilestone()
    If ActiveGanttVBCtl1.Milestones.Count <> 0 Then
        ActiveGanttVBCtl1.Milestones.Remove RndObjectKey(ActiveGanttVBCtl1.Milestones)
    End If
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "Milestone Removed"
    End If
End Sub

Private Sub mp_ClearMilestones()
    ActiveGanttVBCtl1.Milestones.Clear
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "Milestones Cleared"
    End If
End Sub

Private Sub mp_ClassTest()
    Dim lTestIndex As Long
    lTestIndex = RndLong(1, 26)
    Select Case lTestIndex
        Case 1
            mp_AddRow
        Case 2
            mp_RemoveRow
        Case 3
            mp_ClearRows
        Case 4
            mp_AddColumn
        Case 5
            mp_RemoveColumn
        Case 6
            mp_ClearColumns
        Case 7
            mp_AddTask
        Case 8
            mp_RemoveTask
        Case 9
            mp_ClearTasks
        Case 10
            mp_AddStyle
        Case 11
            mp_RemoveStyle
        Case 12
            mp_ClearStyles
        Case 13
            mp_SetCellCaption
        Case 14
            mp_SetCellStyleIndex
        Case 15
            mp_AddPercentageGroup
        Case 16
            mp_RemovePercentageGroup
        Case 17
            mp_ClearPercentageGroups
        Case 18
            mp_AddLayer
        Case 19
            mp_RemoveLayer
        Case 20
            mp_ClearLayers
        Case 21
            mp_AddPercentage
        Case 22
            mp_RemovePercentage
        Case 23
            mp_ClearPercentages
        Case 24
            mp_AddMilestone
        Case 25
            mp_RemoveMilestone
        Case 26
            mp_ClearMilestones
    End Select
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    mp_lStressFactor = 10
End Sub

Private Sub mp_RemoveTask()
    If ActiveGanttVBCtl1.Tasks.Count <> 0 Then
        ActiveGanttVBCtl1.Tasks.Remove RndObjectKey(ActiveGanttVBCtl1.Tasks)
    End If
    If mp_bDisplayMessage = True Then
        txtDisplayMessage.Text = "Task Removed"
    End If
End Sub
