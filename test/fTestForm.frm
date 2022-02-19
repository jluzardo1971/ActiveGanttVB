VERSION 5.00
Object = "{B20AA969-D46A-41DA-AEBB-0E5015A3B668}#1.0#0"; "ActiveGanttVB2.ocx"
Begin VB.Form fTestForm 
   Caption         =   "Test Form"
   ClientHeight    =   7860
   ClientLeft      =   2850
   ClientTop       =   795
   ClientWidth     =   11130
   LinkTopic       =   "Form2"
   ScaleHeight     =   7860
   ScaleWidth      =   11130
   Begin ActiveGanttVB.ActiveGanttVBCtl ActiveGanttVBCtl1 
      Height          =   4455
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7858
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty UpperTierFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty LowerTierFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty NotchFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TimeLineMarkerDate=   38145.6669560185
      HorizontalScrollBarStart=   38145.6669560185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   1680
      TabIndex        =   0
      Top             =   5400
      Width           =   2655
   End
End
Attribute VB_Name = "fTestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim sStart As Single
    Dim sEnd As Single
    sStart = Timer
    Dim L As Long
    ActiveGanttVBCtl1.Styles.Add "Milestones"
    With ActiveGanttVBCtl1.Styles.Item("Milestones")
        .Appearance = SA_FLAT
        .StartShapeIndex = 1
        .Placement = PLC_OFFSETPLACEMENT
        .BackgroundPattern = FP_DARK
        .OffsetTop = 10
        .OffsetBottom = 10
    End With
    ActiveGanttVBCtl1.AutomaticRedraw = False
    For L = 0 To 1000
        ActiveGanttVBCtl1.Rows.Add "K" & L, "Row: " & L, True
        ActiveGanttVBCtl1.Tasks.Add "Gantt: " & L, "K" & L, Now(), DateAdd("h", 1, Now()), "G" & L
        ActiveGanttVBCtl1.Milestones.Add "", "K" & L, DateAdd("h", 4, Now()), "M" & L, "Milestones"
    Next L
    ActiveGanttVBCtl1.Redraw
    sEnd = Timer
    MsgBox sEnd - sStart
    Debug.Print ActiveGanttVBCtl1.Rows.Item("K34").Caption
    Debug.Print ActiveGanttVBCtl1.Tasks.Item("G34").Caption
End Sub

