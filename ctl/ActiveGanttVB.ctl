VERSION 5.00
Begin VB.UserControl ActiveGanttVBCtl 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10440
   KeyPreview      =   -1  'True
   PropertyPages   =   "ActiveGanttVB.ctx":0000
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   696
   ToolboxBitmap   =   "ActiveGanttVB.ctx":0014
   Begin VB.PictureBox picFocus 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1095
      ScaleWidth      =   1095
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.PictureBox picDragMoveCur 
      Height          =   735
      Left            =   3240
      Picture         =   "ActiveGanttVB.ctx":0328
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picSizeRowCur 
      Height          =   495
      Left            =   2520
      Picture         =   "ActiveGanttVB.ctx":047A
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picScrollTimeLineCur 
      Height          =   495
      Left            =   1920
      Picture         =   "ActiveGanttVB.ctx":05CC
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picSizeColumnCur 
      Height          =   495
      Left            =   1320
      Picture         =   "ActiveGanttVB.ctx":08D6
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.HScrollBar hsbHorizontal2 
      Height          =   255
      Left            =   2640
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1695
   End
   Begin VB.HScrollBar hsbHorizontal1 
      Height          =   255
      LargeChange     =   20
      Left            =   960
      MousePointer    =   1  'Arrow
      SmallChange     =   5
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1335
   End
   Begin VB.VScrollBar vsbVertical 
      Height          =   4335
      LargeChange     =   5
      Left            =   4680
      Max             =   1
      Min             =   1
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Value           =   1
      Width           =   255
   End
   Begin VB.Label mp_oToolTip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
End
Attribute VB_Name = "ActiveGanttVBCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
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

'//  ----------------------------------------------------------------------------------------
'//  Graphics Library Public Enumerations
'//  ----------------------------------------------------------------------------------------

Public Enum GRE_ARROWDIRECTION
    AWD_UP = 0
    AWD_DOWN = 1
    AWD_LEFT = 2
    AWD_RIGHT = 3
End Enum

Public Type GRE_POINTTYPE
    X As Long
    Y As Long
End Type

Public Enum GRE_LINEDRAWSTYLE
    LDS_SOLID = 0
    LDS_DASH = 1
    LDS_DOT = 2
    LDS_DASHDOT = 3
    LDS_DASHDOTDOT = 4
End Enum

Public Enum GRE_EDGETYPE
    ET_SUNKEN = 1
    ET_RAISED = 2
End Enum

Public Enum GRE_BUTTONSTYLE
    BT_NORMALWINDOWS = 0
    BT_LIGHTWEIGHT = 1
End Enum

Public Enum GRE_COLORS
    CLR_BLACK = 0 '// RGB(0, 0, 0)
    CLR_WHITE = 16777215 '// RGB(255, 255, 255)
    CLR_RED = 255 '// RGB(255, 0, 0)
    CLR_DARKGREY = 8421504 '// RGB(128, 128, 128) 'Gray
    CLR_VERYDARKGREY = 4210752 '// RGB(64, 64, 64) 'ControlDarkDark
    CLR_VERYLIGHTGREY = 15461355 '// RGB(235, 235, 235) 'WhiteSmoke
    CLR_ALMOSTBLACK = 4342082 '// RGB(66, 66, 65)
    CLR_BUTTONFACE = 13160660 '// RGB(15, 0, 0)
    
    CLR_CORNFLOWERBLUE = 15570276 '// RGB(100, 149, 237)
    CLR_MEDIUMSLATEBLUE = 15624315 '// RGB(123, 104, 238)
    CLR_SLATEBLUE = 13458026 '// RGB(106, 90, 205)
    CLR_ROYALBLUE = 14772545 '// RGB(65, 105, 225)
    CLR_SKYBLUE = 15453831 '// RGB(135, 206, 235)
    CLR_DEEPSKYBLUE = 16760576 '// RGB(0, 191, 255)
    CLR_DODGERBLUE = 16748574 '// RGB(30, 144, 255)
    
    CLR_CADETBLUE = 10526303 '// RGB(95, 158, 160)
    CLR_DARKTURQUOISE = 13749760 '// RGB(0, 206, 209)
    CLR_CYAN = 16776960 '// RGB(0, 255, 255)
    CLR_PALETURQUOISE = 15658671 '// RGB(175, 238, 238)
End Enum

Public Enum GRE_LINETYPE
    LT_NORMAL = 0
    LT_BORDER = 1
    LT_FILLED = 2
End Enum

Public Enum GRE_BACKGROUNDPATTERN
    FP_SOLID = 0
    FP_TRANSPARENT = 1
    FP_HORIZONTALLINE = 2
    FP_VERTICALLINE = 3
    FP_UPWARDDIAGONAL = 4
    FP_DOWNWARDDIAGONAL = 5
    FP_CROSS = 6
    FP_DIAGONALCROSS = 7
    FP_LIGHT = 8
    FP_MEDIUM = 9
    FP_DARK = 10
    FP_GRADIENT = 11
End Enum

Public Enum GRE_FIGURETYPE
    FT_NONE = 0
    FT_PROJECTUP = 1
    FT_PROJECTDOWN = 2
    FT_DIAMOND = 3
    FT_CIRCLEDIAMOND = 4
    FT_TRIANGLEUP = 5
    FT_TRIANGLEDOWN = 6
    FT_TRIANGLERIGHT = 7
    FT_TRIANGLELEFT = 8
    FT_CIRCLETRIANGLEUP = 9
    FT_CIRCLETRIANGLEDOWN = 10
    FT_ARROWUP = 11
    FT_ARROWDOWN = 12
    FT_CIRCLEARROWUP = 13
    FT_CIRCLEARROWDOWN = 14
    FT_SMALLPROJECTUP = 15
    FT_SMALLPROJECTDOWN = 16
    FT_RECTANGLE = 17
    FT_SQUARE = 18
    FT_CIRCLE = 19
End Enum

Public Enum GRE_DRAWTEXTCONSTANTS
    DT_TOP = &H0
    DT_LEFT = &H0
    DT_CENTER = &H1
    DT_RIGHT = &H2
    DT_VCENTER = &H4
    DT_BOTTOM = &H8
    DT_WORDBREAK = &H10
    DT_SINGLELINE = &H20
    DT_EXPANDTABS = &H40
    DT_TABSTOP = &H80
    DT_NOCLIP = &H100
    DT_EXTERNALLEADING = &H200
    DT_NOPREFIX = &H800
    DT_INTERNAL = &H1000
    DT_EDITCONTROL = &H2000
    DT_PATH_ELLIPSIS = &H4000
    DT_END_ELLIPSIS = &H8000
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
End Enum

Public Enum GRE_CONNLINESTYLE
    PDS_NORMAL = 0
    PDS_ENHANCED = 1
    PDS_STRAIGHTLINES = 3
End Enum

Public Enum GRE_ARROWHEADS
    AH_END = 0
    AH_START = 1
    AH_BOTH = 2
    AH_NONE = 3
End Enum

Public Enum GRE_VERTICALALIGNMENT
    VAL_TOP = 1
    VAL_CENTER = 2
    VAL_BOTTOM = 3
End Enum

Public Enum GRE_HORIZONTALALIGNMENT
    HAL_LEFT = 1
    HAL_CENTER = 2
    HAL_RIGHT = 3
End Enum

Public Enum GRE_GRADIENTFILLMODE
    GDT_HORIZONTAL = 0
    GDT_VERTICAL = 1
End Enum

'//  ----------------------------------------------------------------------------------------
'//  ActiveGantt Library Public Enumerations
'//  ----------------------------------------------------------------------------------------

    Public Enum E_EVENTTARGET
        EVT_NONE = 0
        EVT_TASK = 1
        EVT_PERCENTAGE = 2
        EVT_MILESTONE = 3
        EVT_ROW = 4
        EVT_CELL = 5
        EVT_COLUMN = 6
        EVT_CLIENTAREA = 7
        EVT_EMPTYAREA = 8
        EVT_TABLE = 9
        EVT_TIMELINE = 10
        EVT_TIMEBLOCK = 11
    End Enum


    Public Enum E_TOOLTIPACTION
        TA_CLIENTAREA = 1
        TA_TASKADD = 2
        TA_STRETCHTASK = 3
        TA_MOVETASK = 4
        TA_OVERTASK = 5
        TA_MOVEMILESTONE = 6
        TA_OVERMILESTONE = 7
    End Enum


Public Enum E_TIERPOSITION
    SP_UPPER = 0
    SP_LOWER = 1
    SP_MIDDLE = 2
End Enum

Public Enum E_TIERTYPE
    ST_DAYOFWEEK = 1
    ST_MONTH = 2
    ST_QUARTER = 3
    ST_YEAR = 4
    ST_WEEK = 5
    ST_CUSTOM = 6
    ST_DAY = 7
    ST_DAYOFYEAR = 8
    ST_HOUR = 9
    ST_MINUTE = 10
End Enum

Public Enum E_SORTTYPE
    ES_STRING = 0
    ES_NUMERIC = 1
    ES_DATE = 2
End Enum

Public Enum E_ADDMODE
    AT_TASKADD = 0
    AT_MILESTONEADD = 1
End Enum

Public Enum E_EDITMODE
    ET_TASKMILESTONE = 0
    ET_PERCENTAGE = 1
End Enum

Public Enum E_OBJECTTYPE
    OT_TASK = 0
    OT_MILESTONE = 1
End Enum

Public Enum E_PLACEMENT
    PLC_ROWEXTENTSPLACEMENT = 0
    PLC_OFFSETPLACEMENT = 1
End Enum

Public Enum E_CAPTIONPLACEMENT
    SCP_OBJECTEXTENTSPLACEMENT = 0
    SCP_OFFSETPLACEMENT = 1
    SCP_EXTERIORPLACEMENT = 2
End Enum

Public Enum E_TASKATTRIBUTES
    TA_DEFINITESTARTANDEND = 0
    TA_INDEFINITESTART = 1
    TA_INDEFINITEEND = 2
    TA_INDEFINITESTARTANDEND = 3
End Enum

Public Enum E_TYPE
    TP_NONE = -1
    TP_TASK = 0
    TP_MILESTONE = 1
End Enum

Public Enum E_STYLEAPPEARANCE
    SA_RAISED = 0
    SA_SUNKEN = 1
    SA_FLAT = 2
    SA_GRAPHICAL = 3
    SA_CELL = 4
End Enum

Public Enum E_SCROLLBEHAVIOUR
    SB_DISABLE = 0
    SB_HIDE = 1
End Enum

Public Enum E_REPORTERRORS
    RE_MSGBOX = 0
    RE_RAISE = 1
    RE_RAISEEVENT = 2
    RE_HIDE = 3
End Enum

Public Enum E_PROGRESSLINELENGTH
    TLMA_TICKMARKAREA = 0
    TLMA_CLIENTAREA = 1
    TLMA_BOTH = 2
    TLMA_NONE = 4
End Enum

Public Enum E_STYLEBORDER
    SBR_NONE = 0
    SBR_SINGLE = 1
End Enum

Public Enum E_BORDERSTYLE
    TLB_NONE = 0
    TLB_SINGLE = 1
    TLB_3D = 2
End Enum

Public Enum E_PROGRESSLINETYPE
    TLMT_SYSTEMTIME = 0
    TLMT_USER = 1
End Enum

Public Enum E_DROPMODE
    DRP_NONE = 0
    DRP_MANUAL = 1
End Enum

Public Enum E_TIMEBLOCKBEHAVIOUR
    TBB_ROWEXTENTS = 0
    TBB_CONTROLEXTENTS = 1
End Enum

Public Enum E_ENABLEOBJECTS
    EO_CURRENTLAYERONLY = 0
    EO_ALLLAYERS = 1
End Enum

Public Enum E_MOVEMENTTYPE
    MT_UNRESTRICTED = 0
    MT_RESTRICTEDTOROW = 1
    MT_MOVEMENTDISABLED = 2
End Enum

Public Enum E_TASKMOVETYPE
    GMT_NONE = -1
    GMT_MOVE = 0
    GMT_STRETCHLEFT = 1
    GMT_STRETCHRIGHT = 2
End Enum

Public Enum E_TICKMARKTYPES
    TLT_BIG = 0
    TLT_MEDIUM = 1
    TLT_SMALL = 2
End Enum

Public Enum E_SCROLLBAR
    SCR_VERTICAL = 0
    SCR_HORIZONTAL1 = 1
    SCR_HORIZONTAL2 = 2
End Enum


'//  ----------------------------------------------------------------------------------------
'//  ActiveGantt Private Enumerations
'//  ----------------------------------------------------------------------------------------



Private Enum E_PROPOPTYPE
    POT_INITPROPERTIES = 0
    POT_READPROPERTIES = 1
    POT_WRITEPROPERTIES = 2
End Enum

Private Enum E_SCROLLSTATE
    SS_CANTDISPLAY = 0
    SS_NOTNEEDED = 1
    SS_NEEDED = 2
    SS_SHOWN = 3
    SS_HIDDEN = 4
End Enum

Private Enum E_DRAWOPTYPE
    DOT_ALL = 0
    DOT_ROWSANDCLIENTAREA = 1
    DOT_TABLEAREA = 2
    DOT_TIMELINEANDCLIENTAREA = 3
End Enum

Private Enum E_CURSORTYPE
    CT_NORMAL = 0
    CT_SIZETASK = 1
    CT_MOVETASK = 2
    CT_MOVEMILESTONE = 3
    CT_CLIENTAREA = 4
    CT_MOVESPLITTER = 5
    CT_ROWHEIGHT = 6
    CT_COLUMNWIDTH = 7
    CT_MOVEROW = 8
    CT_MOVECOLUMN = 11
    CT_SCROLLTIMELINE = 9
    CT_NODROP = 10
End Enum

'//  ----------------------------------------------------------------------------------------
'//  Member Variables
'//  ----------------------------------------------------------------------------------------

Public Rows As clsRows
Attribute Rows.VB_VarHelpID = -1
Public Tasks As clsTasks
Attribute Tasks.VB_VarHelpID = -1
Public Columns As clsColumns
Attribute Columns.VB_VarHelpID = -1
Public Styles As clsStyles
Attribute Styles.VB_VarHelpID = -1
Public Layers As clsLayers
Attribute Layers.VB_VarHelpID = -1
Public Milestones As clsMilestones
Attribute Milestones.VB_VarHelpID = -1
Public Percentages As clsPercentages
Attribute Percentages.VB_VarHelpID = -1
Public PercentageGroups As clsPercentageGroups
Attribute PercentageGroups.VB_VarHelpID = -1
Public TimeBlocks As clsTimeBlocks
Public Views As clsViews
Public DefaultValues As clsDefaultValues
Private mp_oMouseEvents As clsMouseEvents
Private mp_oCurrentView As clsView
Public Splitter As clsSplitter
Public Printer As clsPrinter
Private clsG As clsGraphics
Private clsM As clsMath
Private clsS As clsString

'// Color

'// Font
Private mp_oFont As StdFont

'// Boolean
Private mp_bAllowAdd As Boolean
Private mp_bAllowEdit As Boolean
Private mp_bAllowSplitterMove As Boolean
Private mp_bAllowColumnSize As Boolean
Private mp_bAllowRowSize As Boolean
Private mp_bAllowRowSwap As Boolean
Private mp_bAllowColumnSwap As Boolean
Private mp_bAllowTimeLineScroll As Boolean
Private mp_bScrollBarsVisible As Boolean
Private mp_bFlickerFree As Boolean
Private mp_bPropertiesRead As Boolean

'// Date
Private mp_dtTimeLineEndBuffer As Date
Private mp_dtTimeLineStartBuffer As Date

'// Long Integer
Private mp_lMinRowHeight As Long
Private mp_lMinColumnWidth As Long
Private mp_lFontCharWidth As Long
Private mp_lSelectedTaskIndex As Long
Private mp_lSelectedMilestoneIndex As Long
Private mp_lSelectedColumnIndex As Long
Private mp_lSelectedRowIndex As Long
Private mp_lSelectedCellIndex As Long

'// String
Private mp_sCurrentLayer As String
Private mp_sCurrentView As String
Private mp_sXML As String

'// Short Integer
Private mp_yAddMode As E_ADDMODE
Private mp_yEditMode As E_EDITMODE
Private mp_yScrollBarBehaviour As E_SCROLLBEHAVIOUR
Private mp_yTimeBlockBehaviour As E_TIMEBLOCKBEHAVIOUR
Private mp_yEnableObjects As E_ENABLEOBJECTS
Private mp_yOLEDragMode As E_DROPMODE
Private mp_yErrorReports As E_REPORTERRORS
Private mp_yBorderStyle As E_BORDERSTYLE
Private mp_yDrawOperationType As E_DRAWOPTYPE

'// Event Variables
Private mp_TPCaption As String
Private mp_TPDisplayToolTip As Boolean

Private WithEvents oHScrollBar1 As clsHScrollBarEx
Attribute oHScrollBar1.VB_VarHelpID = -1
Private WithEvents oHScrollBar2 As clsHScrollBarEx
Attribute oHScrollBar2.VB_VarHelpID = -1
Private WithEvents oVScrollBar As clsVScrollBarEx
Attribute oVScrollBar.VB_VarHelpID = -1

'//  ----------------------------------------------------------------------------------------
'//  ActiveGantt Public Events
'//  ----------------------------------------------------------------------------------------

Public Event ControlClick(ByVal EventTarget As E_EVENTTARGET, ByVal ObjectIndex As Long, ByVal ParentObjectIndex As Long, ByVal X As Single, ByVal Y As Single, ByVal Button As Integer)
Public Event ControlDblClick(ByVal EventTarget As E_EVENTTARGET, ByVal ObjectIndex As Long, ByVal ParentObjectIndex As Long, ByVal X As Single, ByVal Y As Single, ByVal Button As Integer)
Public Event ControlMouseDown(ByVal EventTarget As E_EVENTTARGET, ByVal ObjectIndex As Long, ByVal ParentObjectIndex As Long, ByVal X As Single, ByVal Y As Single, ByVal Button As Integer)
Public Event ControlMouseMove(ByVal EventTarget As E_EVENTTARGET, ByVal ObjectIndex As Long, ByVal ParentObjectIndex As Long, ByVal X As Single, ByVal Y As Single, ByVal Button As Integer)
Public Event ControlMouseUp(ByVal EventTarget As E_EVENTTARGET, ByVal ObjectIndex As Long, ByVal ParentObjectIndex As Long, ByVal X As Single, ByVal Y As Single, ByVal Button As Integer)
Public Event ObjectSelected(ByVal EventTarget As E_EVENTTARGET, ByVal ObjectIndex As Long, ByVal ParentObjectIndex As Long)

Public Event Draw(ByVal EventTarget As E_EVENTTARGET, ByRef CustomDraw As Boolean, ByVal ObjectIndex As Long, ByVal ParentObjectIndex As Long, ByVal lHdc As Long)
Public Event PredecessorDraw(ByRef CustomDraw As Boolean, ByVal SourceIndex As Long, ByVal SourceType As E_OBJECTTYPE, ByVal LinkIndex As Long, ByVal LinkType As E_OBJECTTYPE, ByVal lHdc As Long)
Public Event CustomTierDraw(ByVal Position As E_TIERPOSITION, ByVal StartDate As Date, ByVal EndDate As Date, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal LeftTrim As Long, ByVal RightTrim As Long, ByVal lHdc As Long, ByRef Caption As String, ByRef StyleIndex As String)
Public Event TierCaptionDraw(ByRef Caption As String, ByVal dtDate As Date, ByVal Position As E_TIERPOSITION)

Public Event ObjectAdded(ByVal EventTarget As E_EVENTTARGET, ByVal ObjectIndex As Long)

Public Event BeginObjectMove(ByVal EventTarget As E_EVENTTARGET, ByVal Index As Long, ByRef Cancel As Boolean)
Public Event ObjectMove(ByVal EventTarget As E_EVENTTARGET, ByVal Index As Long, ByRef Cancel As Boolean)
Public Event EndObjectMove(ByVal EventTarget As E_EVENTTARGET, ByVal Index As Long, ByRef Cancel As Boolean)
Public Event CompleteObjectMove(ByVal EventTarget As E_EVENTTARGET, ByVal Index As Long)

Public Event BeginObjectSize(ByVal EventTarget As E_EVENTTARGET, ByVal Index As Long, ByRef Cancel As Boolean)
Public Event ObjectSize(ByVal EventTarget As E_EVENTTARGET, ByVal Index As Long, ByRef Cancel As Boolean)
Public Event EndObjectSize(ByVal EventTarget As E_EVENTTARGET, ByVal Index As Long, ByRef Cancel As Boolean)
Public Event CompleteObjectSize(ByVal EventTarget As E_EVENTTARGET, ByVal Index As Long)

Public Event OLETaskStartDrag(ByVal Index As Long, Data As DataObject, ByVal AllowedEffects As Integer)
Public Event OLEMilestoneStartDrag(ByVal Index As Long, Data As DataObject, ByVal AllowedEffects As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEDragOver(Data As DataObject, Effect As Integer, Button As Integer, Shift As Integer, ByVal RowIndex As Long, ByVal DatePosition As Date, State As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Integer, Button As Integer, Shift As Integer, ByVal RowIndex As Long, ByVal DatePosition As Date)

Public Event ActiveGanttError(ByVal Number As Long, ByVal Description As String, ByVal Source As String)
Public Event Scroll(ByVal ScrollBarType As E_SCROLLBAR, ByVal Offset As Integer)
Public Event TimeLineChanged()
Public Event ControlRedrawn()

Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Public Event ToolTip(ByRef ToolTipAction As E_TOOLTIPACTION, ByRef Caption As String, ByVal StartDate As Date, ByVal EndDate As Date, ByVal Index As Long, ByRef DisplayToolTip As Boolean)

Friend Sub FireControlClick(ByVal EventTarget As E_EVENTTARGET, ByVal ObjectIndex As Long, ByVal ParentObjectIndex As Long, ByVal X As Single, ByVal Y As Single, ByVal Button As Integer)
    RaiseEvent ControlClick(EventTarget, ObjectIndex, ParentObjectIndex, X, Y, Button)
End Sub

Friend Sub FireControlDblClick(ByVal EventTarget As E_EVENTTARGET, ByVal ObjectIndex As Long, ByVal ParentObjectIndex As Long, ByVal X As Single, ByVal Y As Single, ByVal Button As Integer)
    RaiseEvent ControlDblClick(EventTarget, ObjectIndex, ParentObjectIndex, X, Y, Button)
End Sub

Friend Sub FireControlMouseDown(ByVal EventTarget As E_EVENTTARGET, ByVal ObjectIndex As Long, ByVal ParentObjectIndex As Long, ByVal X As Single, ByVal Y As Single, ByVal Button As Integer)
    RaiseEvent ControlMouseDown(EventTarget, ObjectIndex, ParentObjectIndex, X, Y, Button)
End Sub

Friend Sub FireControlMouseMove(ByVal EventTarget As E_EVENTTARGET, ByVal ObjectIndex As Long, ByVal ParentObjectIndex As Long, ByVal X As Single, ByVal Y As Single, ByVal Button As Integer)
    RaiseEvent ControlMouseMove(EventTarget, ObjectIndex, ParentObjectIndex, X, Y, Button)
End Sub

Friend Sub FireControlMouseUp(ByVal EventTarget As E_EVENTTARGET, ByVal ObjectIndex As Long, ByVal ParentObjectIndex As Long, ByVal X As Single, ByVal Y As Single, ByVal Button As Integer)
    RaiseEvent ControlMouseUp(EventTarget, ObjectIndex, ParentObjectIndex, X, Y, Button)
End Sub

Friend Sub FireObjectSelected(ByVal EventTarget As E_EVENTTARGET, ByVal ObjectIndex As Long, ByVal ParentObjectIndex As Long)
    RaiseEvent ObjectSelected(EventTarget, ObjectIndex, ParentObjectIndex)
End Sub

Friend Sub FireScroll(ByVal ScrollBarType As E_SCROLLBAR, ByVal Offset As Integer)
    RaiseEvent Scroll(ScrollBarType, Offset)
End Sub

Friend Sub FireCustomTierDraw(ByVal Position As E_TIERPOSITION, ByVal StartDate As Date, ByVal EndDate As Date, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal LeftTrim As Long, ByVal RightTrim As Long, ByVal lHdc As Long, ByRef Caption As String, ByRef StyleIndex As String)
    RaiseEvent CustomTierDraw(Position, StartDate, EndDate, Left, Top, Right, Bottom, LeftTrim, RightTrim, lHdc, Caption, StyleIndex)
End Sub

Friend Sub FireTierCaptionDraw(ByRef Caption As String, ByVal dtDate As Date, ByVal Position As E_TIERPOSITION)
    RaiseEvent TierCaptionDraw(Caption, dtDate, Position)
End Sub

Friend Sub FireDraw(ByVal EventTarget As E_EVENTTARGET, ByRef CustomDraw As Boolean, ByVal ObjectIndex As Long, ByVal ParentObjectIndex As Long, ByVal lHdc As Long)
    RaiseEvent Draw(EventTarget, CustomDraw, ObjectIndex, ParentObjectIndex, lHdc)
End Sub

Friend Sub FirePredecessorDraw(ByRef CustomDraw As Boolean, ByVal SourceIndex As Long, ByVal SourceType As E_OBJECTTYPE, ByVal LinkIndex As Long, ByVal LinkType As E_OBJECTTYPE, ByVal lHdc As Long)
    RaiseEvent PredecessorDraw(CustomDraw, SourceIndex, SourceType, LinkIndex, LinkType, lHdc)
End Sub

Friend Sub FireAdded(ByVal EventTarget As E_EVENTTARGET, ByVal ObjectIndex As Long)
    RaiseEvent ObjectAdded(EventTarget, ObjectIndex)
End Sub


Friend Sub FireBeginObjectMove(ByVal EventTarget As E_EVENTTARGET, ByVal Index As Long, ByRef Cancel As Boolean)
    RaiseEvent BeginObjectMove(EventTarget, Index, Cancel)
End Sub

Friend Sub FireObjectMove(ByVal EventTarget As E_EVENTTARGET, ByVal Index As Long, ByRef Cancel As Boolean)
    RaiseEvent ObjectMove(EventTarget, Index, Cancel)
End Sub

Friend Sub FireEndObjectMove(ByVal EventTarget As E_EVENTTARGET, ByVal Index As Long, ByRef Cancel As Boolean)
    RaiseEvent EndObjectMove(EventTarget, Index, Cancel)
End Sub

Friend Sub FireCompleteObjectMove(ByVal EventTarget As E_EVENTTARGET, ByVal Index As Long)
    RaiseEvent CompleteObjectMove(EventTarget, Index)
End Sub

Friend Sub FireBeginObjectSize(ByVal EventTarget As E_EVENTTARGET, ByVal Index As Long, ByRef Cancel As Boolean)
    RaiseEvent BeginObjectSize(EventTarget, Index, Cancel)
End Sub

Friend Sub FireObjectSize(ByVal EventTarget As E_EVENTTARGET, ByVal Index As Long, ByRef Cancel As Boolean)
    RaiseEvent ObjectSize(EventTarget, Index, Cancel)
End Sub

Friend Sub FireEndObjectSize(ByVal EventTarget As E_EVENTTARGET, ByVal Index As Long, ByRef Cancel As Boolean)
    RaiseEvent EndObjectSize(EventTarget, Index, Cancel)
End Sub

Friend Sub FireCompleteObjectSize(ByVal EventTarget As E_EVENTTARGET, ByVal Index As Long)
    RaiseEvent CompleteObjectSize(EventTarget, Index)
End Sub

Friend Sub FireToolTip(ByRef ToolTipAction As E_TOOLTIPACTION, ByRef Caption As String, ByVal StartDate As Date, ByVal EndDate As Date, ByVal Index As Long, ByRef DisplayToolTip As Boolean)
    RaiseEvent ToolTip(ToolTipAction, Caption, StartDate, EndDate, Index, DisplayToolTip)
End Sub

'//  ----------------------------------------------------------------------------------------
'//  Friend Functions
'//  ----------------------------------------------------------------------------------------

Friend Function f_Hdc()
    f_Hdc = UserControl.hdc
End Function

Public Function ControlHdc() As Long
    ControlHdc = UserControl.hdc
End Function

Friend Sub f_DrawLine(ByVal v_X2 As Long, ByVal v_Y2 As Long)
    UserControl.Line (v_X2, v_Y2)-(v_X2, v_Y2)
End Sub

Friend Function f_Width() As Long
    f_Width = UserControl.Width / Screen.TwipsPerPixelX
End Function

Friend Function f_Height() As Long
    f_Height = UserControl.Height / Screen.TwipsPerPixelY
End Function

Friend Function mp_lStrWidth(ByRef sString As String, ByRef r_oFont As StdFont) As Long
    Set UserControl.Font = r_oFont
    mp_lStrWidth = UserControl.TextWidth(sString) * mp_lFontCharWidth
End Function

Friend Function mp_lStrHeight(ByRef sString As String, ByRef r_oFont As StdFont) As Long
    Set UserControl.Font = r_oFont
    mp_lStrHeight = UserControl.TextHeight(sString)
End Function

Friend Function f_oHScrollBar1() As clsHScrollBarEx
    Set f_oHScrollBar1 = oHScrollBar1
End Function

Friend Function f_oHScrollBar2() As clsHScrollBarEx
    Set f_oHScrollBar2 = oHScrollBar2
End Function

Friend Function f_oVScrollBar() As clsVScrollBarEx
    Set f_oVScrollBar = oVScrollBar
End Function

Friend Function f_ToolTip() As Label
    Set f_ToolTip = mp_oToolTip
End Function

Friend Function f_Parent() As Object
    Set f_Parent = UserControl.Parent
End Function

Friend Sub f_SetCursor(ByVal v_iCursorType As E_CURSORTYPE)
    Select Case v_iCursorType
        Case E_CURSORTYPE.CT_NORMAL
            UserControl.MousePointer = 0
        Case E_CURSORTYPE.CT_SIZETASK
            UserControl.MousePointer = vbSizeWE
        Case E_CURSORTYPE.CT_MOVETASK
            UserControl.MousePointer = vbSizePointer
        Case E_CURSORTYPE.CT_MOVEMILESTONE
            UserControl.MousePointer = vbSizePointer
        Case E_CURSORTYPE.CT_CLIENTAREA
            UserControl.MousePointer = vbCrosshair
        Case E_CURSORTYPE.CT_MOVESPLITTER
            UserControl.MousePointer = vbSizeWE
        Case E_CURSORTYPE.CT_ROWHEIGHT
            Set UserControl.MouseIcon = picSizeRowCur.Picture
            UserControl.MousePointer = 99
        Case E_CURSORTYPE.CT_COLUMNWIDTH
            Set UserControl.MouseIcon = picSizeColumnCur.Picture
            UserControl.MousePointer = 99
        Case E_CURSORTYPE.CT_MOVEROW
            Set UserControl.MouseIcon = picDragMoveCur.Picture
            UserControl.MousePointer = 99
        Case E_CURSORTYPE.CT_MOVECOLUMN
            Set UserControl.MouseIcon = picDragMoveCur.Picture
            UserControl.MousePointer = 99
        Case E_CURSORTYPE.CT_SCROLLTIMELINE
            Set UserControl.MouseIcon = picScrollTimeLineCur.Picture
            UserControl.MousePointer = 99
        Case E_CURSORTYPE.CT_NODROP
            UserControl.MousePointer = vbNoDrop
            'mp_bCompleteDragOperation = False
    End Select
End Sub

Friend Function f_UserMode() As Boolean
    f_UserMode = UserControl.Ambient.UserMode
End Function

Friend Sub f_Draw()
    mp_yDrawOperationType = E_DRAWOPTYPE.DOT_ALL
    mp_Draw
End Sub

'//  ----------------------------------------------------------------------------------------
'//  Library Objects
'//  ----------------------------------------------------------------------------------------

Public Function MathLib() As clsMath
    Set MathLib = clsM
End Function

Public Function StrLib() As clsString
    Set StrLib = clsS
End Function

Friend Function GrphLib() As clsGraphics
    Set GrphLib = clsG
End Function

'//  ----------------------------------------------------------------------------------------
'//  Construction / Destruction & Initialization
'//  ----------------------------------------------------------------------------------------

Private Sub UserControl_Initialize()
    Set oHScrollBar1 = New clsHScrollBarEx
    oHScrollBar1.Initialize Me, hsbHorizontal1
    Set oHScrollBar2 = New clsHScrollBarEx
    oHScrollBar2.Initialize Me, hsbHorizontal2
    Set oVScrollBar = New clsVScrollBarEx
    oVScrollBar.Initialize Me, vsbVertical


    Set clsG = New clsGraphics
    clsG.Initialize Me
    Set clsM = New clsMath
    clsM.Initialize Me
    Set clsS = New clsString
    clsS.Initialize Me
    #If DemoVersion Then
        fAbout.Initialize Me
        fAbout.Show 1
    #End If
    Set DefaultValues = New clsDefaultValues
    DefaultValues.Initialize Me
    Set Rows = New clsRows
    Rows.Initialize Me
    Set Tasks = New clsTasks
    Tasks.Initialize Me
    Set Columns = New clsColumns
    Columns.Initialize Me
    Set Styles = New clsStyles
    Styles.Initialize Me
    Set Layers = New clsLayers
    Layers.Initialize Me
    Set Milestones = New clsMilestones
    Milestones.Initialize Me
    Set Percentages = New clsPercentages
    Percentages.Initialize Me
    Set PercentageGroups = New clsPercentageGroups
    PercentageGroups.Initialize Me
    Set TimeBlocks = New clsTimeBlocks
    TimeBlocks.Initialize Me
    Set Views = New clsViews
    Views.Initialize Me
    Set Printer = New clsPrinter
    Printer.InitializeClass Me
    Set mp_oCurrentView = Views.FItem("0")
    Set Splitter = New clsSplitter
    Splitter.Initialize Me
    Set mp_oMouseEvents = New clsMouseEvents
    mp_oMouseEvents.Initialize Me
    
End Sub

Private Sub UserControl_Terminate()
    Set PercentageGroups = Nothing
    Set Percentages = Nothing
    Set Milestones = Nothing
    Set Layers = Nothing
    Set Styles = Nothing
    Set Columns = Nothing
    Set Tasks = Nothing
    Set Rows = Nothing
End Sub

Private Sub UserControl_InitProperties()
    Dim PropBag As PropertyBag
    DoPropExchange PropBag, E_PROPOPTYPE.POT_INITPROPERTIES
    mp_bPropertiesRead = True
    Me.Redraw
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    DoPropExchange PropBag, E_PROPOPTYPE.POT_READPROPERTIES
    mp_bPropertiesRead = True
    Me.Redraw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    DoPropExchange PropBag, E_PROPOPTYPE.POT_WRITEPROPERTIES
End Sub

Private Sub DoPropExchange(PropBag As PropertyBag, v_yOpType As E_PROPOPTYPE)
    Dim Value As Variant
    PX_Variant PropBag, "FontCharWidth", mp_lFontCharWidth, 1, v_yOpType
    PX_Variant PropBag, "AllowSplitterMove", mp_bAllowSplitterMove, True, v_yOpType
    PX_Variant PropBag, "AllowColumnSize", mp_bAllowColumnSize, True, v_yOpType
    PX_Variant PropBag, "AllowAdd", mp_bAllowAdd, True, v_yOpType
    PX_Variant PropBag, "AllowEdit", mp_bAllowEdit, True, v_yOpType
    PX_Variant PropBag, "AllowRowSize", mp_bAllowRowSize, True, v_yOpType
    PX_Variant PropBag, "AllowRowSwap", mp_bAllowRowSwap, True, v_yOpType
    PX_Variant PropBag, "AllowColumnSwap", mp_bAllowColumnSwap, True, v_yOpType
    PX_Variant PropBag, "AllowTimeLineScroll", mp_bAllowTimeLineScroll, True, v_yOpType
    PX_Variant PropBag, "FlickerFree", mp_bFlickerFree, True, v_yOpType
    PX_Variant PropBag, "ScrollBarsVisible", mp_bScrollBarsVisible, True, v_yOpType
    PX_Variant PropBag, "ScrollBarBehaviour", mp_yScrollBarBehaviour, E_SCROLLBEHAVIOUR.SB_HIDE, v_yOpType
    PX_Variant PropBag, "TimeBlockBehaviour", mp_yTimeBlockBehaviour, E_TIMEBLOCKBEHAVIOUR.TBB_ROWEXTENTS, v_yOpType
    PX_Variant PropBag, "ErrorReports", mp_yErrorReports, E_REPORTERRORS.RE_MSGBOX, v_yOpType
    PX_Variant PropBag, "EnableObjects", mp_yEnableObjects, E_ENABLEOBJECTS.EO_CURRENTLAYERONLY, v_yOpType
    PX_Variant PropBag, "CurrentLayer", mp_sCurrentLayer, "0", v_yOpType
    PX_Variant PropBag, "MinRowHeight", mp_lMinRowHeight, 5, v_yOpType
    PX_Variant PropBag, "MinColumnWidth", mp_lMinColumnWidth, 5, v_yOpType
    PX_Variant PropBag, "BorderStyle", mp_yBorderStyle, E_BORDERSTYLE.TLB_3D, v_yOpType
    PX_Variant PropBag, "AddMode", mp_yAddMode, E_ADDMODE.AT_TASKADD, v_yOpType
    PX_Variant PropBag, "EditMode", mp_yEditMode, E_EDITMODE.ET_TASKMILESTONE, v_yOpType
    PX_Variant PropBag, "OLEDragMode", mp_yOLEDragMode, E_DROPMODE.DRP_NONE, v_yOpType
    Value = UserControl.OLEDropMode
    PX_Variant PropBag, "OLEDropMode", Value, E_DROPMODE.DRP_NONE, v_yOpType
    UserControl.OLEDropMode = Value
    PX_Variant PropBag, "Font", mp_oFont, UserControl.Ambient.Font, v_yOpType, True
    Value = UserControl.BackColor
    PX_ColorEx PropBag, "BackColor", Value, UserControl.BackColor, GRE_COLORS.CLR_WHITE, v_yOpType
    UserControl.BackColor = Value
End Sub

Private Sub PX_Variant(PropBag As PropertyBag, ByVal Name As String, ByRef Value As Variant, ByVal DefaultValue As Variant, ByVal v_yOpType As E_PROPOPTYPE, Optional ByVal bUseSet As Boolean = False)
    Select Case v_yOpType
        Case E_PROPOPTYPE.POT_INITPROPERTIES
            If bUseSet = False Then
                Value = DefaultValue
            Else
                Set Value = DefaultValue
            End If
        Case E_PROPOPTYPE.POT_READPROPERTIES
            If bUseSet = False Then
                Value = PropBag.ReadProperty(Name, DefaultValue)
            Else
                Set Value = PropBag.ReadProperty(Name, DefaultValue)
            End If
        Case E_PROPOPTYPE.POT_WRITEPROPERTIES
            PropBag.WriteProperty Name, Value, DefaultValue
    End Select
End Sub

Private Sub PX_ColorEx(PropBag As PropertyBag, ByVal Name As String, ByRef Value As Variant, ByRef Value2 As OLE_COLOR, ByVal DefaultValue As Variant, ByVal v_yOpType As E_PROPOPTYPE)
    Select Case v_yOpType
        Case E_PROPOPTYPE.POT_INITPROPERTIES
            Value = DefaultValue
            Value2 = clsG.ConvertColor(DefaultValue)
        Case E_PROPOPTYPE.POT_READPROPERTIES
            Value = PropBag.ReadProperty(Name, DefaultValue)
            Value2 = clsG.ConvertColor(Value)
        Case E_PROPOPTYPE.POT_WRITEPROPERTIES
            PropBag.WriteProperty Name, Value, DefaultValue
    End Select
End Sub

    '// ---------------------------------------------------------------------------------------------------------------------
    '// Mouse Events
    '// ---------------------------------------------------------------------------------------------------------------------

Private Sub UserControl_Click()
    mp_oMouseEvents.Click
End Sub

Private Sub UserControl_DblClick()
    mp_oMouseEvents.DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mp_oMouseEvents.m_MouseDown Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mp_oMouseEvents.m_MouseMove Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mp_oMouseEvents.m_MouseUp Button, Shift, X, Y
End Sub

'//  ----------------------------------------------------------------------------------------
'//  ActiveGantt Key Events
'//  ----------------------------------------------------------------------------------------

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'//  ----------------------------------------------------------------------------------------
'//  ActiveGantt Painting And Resizing
'//  ----------------------------------------------------------------------------------------

Private Sub UserControl_Paint()
    If mp_bPropertiesRead = False Then
        Exit Sub
    End If
    picFocus.Left = clsG.Width
    picFocus.Top = clsG.Height
    picFocus.Width = 1
    picFocus.Height = 1
    mp_PositionScrollBars
    If UserControl.Ambient.UserMode = False Then
        mp_DrawDesignMode
    Else
        mp_Draw
    End If
    If mp_oCurrentView.TimeLine.StartDate <> mp_dtTimeLineStartBuffer Or mp_oCurrentView.TimeLine.EndDate <> mp_dtTimeLineEndBuffer Then
        mp_dtTimeLineStartBuffer = mp_oCurrentView.TimeLine.StartDate
        mp_dtTimeLineEndBuffer = mp_oCurrentView.TimeLine.EndDate
        RaiseEvent TimeLineChanged
    End If
    oVScrollBar.Max = Rows.Count
    RaiseEvent ControlRedrawn
End Sub

Private Sub UserControl_Resize()
    UserControl_Paint
End Sub

'//  ----------------------------------------------------------------------------------------
'//  ActiveGantt Scroll Bar Events
'//  ----------------------------------------------------------------------------------------


Private Sub oHScrollBar1_ValueChanged(ByVal Offset As Integer)
    mp_yDrawOperationType = E_DRAWOPTYPE.DOT_TABLEAREA
    clsG.InvalidateRectangle UserControl.hWnd, mt_LeftMargin, 0, Splitter.Left - mt_LeftMargin, clsG.Height
    UserControl_Paint
    RaiseEvent Scroll(E_SCROLLBAR.SCR_HORIZONTAL1, Offset)
End Sub

Private Sub oHScrollBar2_ValueChanged(ByVal Offset As Integer)
    mp_yDrawOperationType = E_DRAWOPTYPE.DOT_TIMELINEANDCLIENTAREA
    clsG.InvalidateRectangle UserControl.hWnd, Splitter.Right, 0, clsG.Width - Splitter.Right, clsG.Height
    UserControl_Paint
    RaiseEvent Scroll(E_SCROLLBAR.SCR_HORIZONTAL2, Offset)
End Sub

Private Sub oVScrollBar_ValueChanged(ByVal Offset As Integer)
    mp_yDrawOperationType = E_DRAWOPTYPE.DOT_ROWSANDCLIENTAREA
    clsG.InvalidateRectangle UserControl.hWnd, 0, mp_oCurrentView.ClientArea.Top, clsG.Width, clsG.Height - mp_oCurrentView.ClientArea.Top
    UserControl_Paint
    RaiseEvent Scroll(E_SCROLLBAR.SCR_VERTICAL, Offset)
End Sub


'//  ----------------------------------------------------------------------------------------
'//  Scrollbar Positioning code
'//  ----------------------------------------------------------------------------------------

Friend Sub mp_PositionScrollBars()
    If UserControl.Ambient.UserMode = False Or clsG.CustomPrinting = True Or mp_bScrollBarsVisible = False Then
        oVScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
        oHScrollBar1.State = E_SCROLLSTATE.SS_CANTDISPLAY
        oHScrollBar2.State = E_SCROLLSTATE.SS_CANTDISPLAY
    Else
        If clsG.Height <= mp_oCurrentView.ClientArea.Top Then
            oVScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
            oHScrollBar1.State = E_SCROLLSTATE.SS_CANTDISPLAY
            oHScrollBar2.State = E_SCROLLSTATE.SS_CANTDISPLAY
        Else
            '// Determine need for oHScrollBar1
            If Columns.Width > Splitter.Position Then '//Leave Splitter.Position
                oHScrollBar1.State = E_SCROLLSTATE.SS_NEEDED
            Else
                oHScrollBar1.State = E_SCROLLSTATE.SS_NOTNEEDED
            End If
            If Splitter.Left < 5 Then
                oHScrollBar1.State = E_SCROLLSTATE.SS_CANTDISPLAY
            End If
            '// Determine need for oHScrollBar2
            If Splitter.Right < clsG.Width - (18 + mp_yBorderStyle) Then
                If mp_oCurrentView.TimeLine.ScrollBar.Enabled = True Then
                    oHScrollBar2.State = E_SCROLLSTATE.SS_NEEDED
                Else
                    oHScrollBar2.State = E_SCROLLSTATE.SS_NOTNEEDED
                End If
            Else
                oHScrollBar2.State = E_SCROLLSTATE.SS_CANTDISPLAY
            End If
            '// Determine need for oVScrollBar
            If ((Rows.Height() + mp_oCurrentView.ClientArea.Top + oHScrollBar1.Height + mp_yBorderStyle) > clsG.Height) Then
                If oHScrollBar2.State = E_SCROLLSTATE.SS_CANTDISPLAY Then
                    oVScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
                Else
                    oVScrollBar.State = E_SCROLLSTATE.SS_NEEDED
                End If
            Else
                If Rows.Count > 0 Then
                    mp_oCurrentView.ClientArea.FirstVisibleRow = 1
                End If
                oVScrollBar.State = E_SCROLLSTATE.SS_NOTNEEDED
            End If
        End If
    End If
    If oVScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
        mp_PositionVerticalScrollBar
    End If
    If oHScrollBar1.State = E_SCROLLSTATE.SS_SHOWN Then
        mp_PositionHorizontal1ScrollBar
    End If
    If oHScrollBar2.State = E_SCROLLSTATE.SS_SHOWN Then
        mp_PositionHorizontal2ScrollBar
    End If
End Sub

Private Sub mp_PositionVerticalScrollBar()
    oVScrollBar.Left = clsG.Width - oVScrollBar.Width - mp_yBorderStyle
    oVScrollBar.Top = mt_TopMargin
    If oHScrollBar2.State = E_SCROLLSTATE.SS_SHOWN Then
        If (clsG.Height - (mp_yBorderStyle * 2) - oHScrollBar1.Height) > 0 Then
            oVScrollBar.Height = clsG.Height - (mp_yBorderStyle * 2) - oHScrollBar1.Height
        End If
    Else
        If (clsG.Height - (mp_yBorderStyle * 2)) > 0 Then
            oVScrollBar.Height = clsG.Height - (mp_yBorderStyle * 2)
        End If
    End If
    oVScrollBar.SmallChange = 1
    If Rows.CalculateHeight(mp_oCurrentView.ClientArea.FirstVisibleRow, Rows.Count) > clsG.Height - mp_oCurrentView.ClientArea.Top - oHScrollBar1.Height - mp_yBorderStyle Then
        oVScrollBar.LargeChange = Rows.CalculateRows(mp_oCurrentView.ClientArea.FirstVisibleRow, clsG.Height - mp_oCurrentView.ClientArea.Top - oHScrollBar1.Height - mp_yBorderStyle)
    End If
End Sub

Private Sub mp_PositionHorizontal1ScrollBar()
    oHScrollBar1.Left = mp_yBorderStyle
    oHScrollBar1.Top = clsG.Height - oHScrollBar1.Height - mp_yBorderStyle
    If Splitter.Left > 0 Then
        oHScrollBar1.Width = Splitter.Left
    End If
    oHScrollBar1.Min = 0
    oHScrollBar1.Max = Columns.Width - Splitter.Position '//Leave Splitter.Position
End Sub

Private Sub mp_PositionHorizontal2ScrollBar()
    oHScrollBar2.Left = Splitter.Right
    oHScrollBar2.Top = clsG.Height - oHScrollBar2.Height - mp_yBorderStyle
    If oVScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
        If clsG.Width - (mp_yBorderStyle) - Splitter.Right - oVScrollBar.Width > 0 Then
            oHScrollBar2.Width = clsG.Width - mp_yBorderStyle - Splitter.Right - oVScrollBar.Width
        End If
    Else
        If clsG.Width - (mp_yBorderStyle) - Splitter.Right > 0 Then
            oHScrollBar2.Width = clsG.Width - mp_yBorderStyle - Splitter.Right
        End If
    End If
End Sub

'//  ----------------------------------------------------------------------------------------
'//  Drawing Functions
'//  ----------------------------------------------------------------------------------------

Private Sub mp_Draw()
    clsG.ResetFocusRectangle
    Select Case mp_yDrawOperationType
        Case E_DRAWOPTYPE.DOT_ALL
            clsG.DrawLine 0, 0, clsG.Width, clsG.Height, GRE_LINETYPE.LT_FILLED, UserControl.BackColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        Case E_DRAWOPTYPE.DOT_ROWSANDCLIENTAREA
            clsG.DrawLine 0, mp_oCurrentView.ClientArea.Top, clsG.Width, clsG.Height, GRE_LINETYPE.LT_FILLED, UserControl.BackColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        Case E_DRAWOPTYPE.DOT_TABLEAREA
            clsG.DrawLine mt_LeftMargin, 0, Splitter.Left, clsG.Height, GRE_LINETYPE.LT_FILLED, UserControl.BackColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        Case E_DRAWOPTYPE.DOT_TIMELINEANDCLIENTAREA
            clsG.DrawLine Splitter.Right, 0, clsG.Width, clsG.Height, GRE_LINETYPE.LT_FILLED, UserControl.BackColor, GRE_LINEDRAWSTYLE.LDS_SOLID
    End Select
    '// Positioning Code
    mp_oCurrentView.TimeLine.Calculate
    Columns.Position
    Rows.Position
    TimeBlocks.Position
    Tasks.Position
    Milestones.Position
    Percentages.Position
    '// Drawing Code
    If mp_yDrawOperationType = E_DRAWOPTYPE.DOT_ALL Or mp_yDrawOperationType = E_DRAWOPTYPE.DOT_TABLEAREA Then
        Columns.Draw
    End If
    If mp_yDrawOperationType = E_DRAWOPTYPE.DOT_ALL Or mp_yDrawOperationType = E_DRAWOPTYPE.DOT_TIMELINEANDCLIENTAREA Then
        mp_oCurrentView.TimeLine.Draw
        mp_oCurrentView.TimeLine.ProgressLine.Draw
    End If
    If mp_yDrawOperationType = E_DRAWOPTYPE.DOT_ALL Or mp_yDrawOperationType = E_DRAWOPTYPE.DOT_ROWSANDCLIENTAREA Or mp_yDrawOperationType = E_DRAWOPTYPE.DOT_TABLEAREA Then
        Rows.Draw
    End If
    If mp_yDrawOperationType = E_DRAWOPTYPE.DOT_ALL Or mp_yDrawOperationType = E_DRAWOPTYPE.DOT_ROWSANDCLIENTAREA Or mp_yDrawOperationType = E_DRAWOPTYPE.DOT_TIMELINEANDCLIENTAREA Then
        TimeBlocks.Draw
        mp_oCurrentView.ClientArea.Draw
        mp_oCurrentView.ClientArea.Grid.Draw
        Tasks.Draw
        Percentages.Draw
        Milestones.Draw
        mp_oCurrentView.TimeLine.ProgressLine.Draw
    End If
    Splitter.Draw
    mp_DrawControlBorder
    mp_DrawDebugMetrics
    '// Demo Version
    #If DemoVersion Then
        If clsG.CustomPrinting = True Then
            Dim oFont As New StdFont
            oFont.Name = "Arial"
            oFont.Size = 14
            oFont.Bold = True
            clsG.TextOutEx Splitter.Right, mp_oCurrentView.ClientArea.Top, mt_RightMargin, mp_oCurrentView.ClientArea.Bottom, "DEMO VERSION", HAL_CENTER, VAL_CENTER, RGB(255, 0, 0), oFont, True
        End If
    #End If
    mp_yDrawOperationType = E_DRAWOPTYPE.DOT_ALL
End Sub

Private Sub mp_DrawDesignMode()
    Dim lLeftBox As Long
    Dim lTop As Long
    Dim lRightBox As Long
    Dim lBottom As Long
    Dim lLeftCA As Long
    Dim lRightCA As Long
    '// DrawTimeLine
    clsG.DrawLine 0, 0, clsG.Width, clsG.Height, GRE_LINETYPE.LT_FILLED, UserControl.BackColor, GRE_LINEDRAWSTYLE.LDS_SOLID
    mp_oCurrentView.TimeLine.Calculate
    mp_oCurrentView.TimeLine.Draw
    clsG.ClearClipRegion
    '// Columns
    lLeftBox = mt_LeftMargin
    lTop = mt_TopMargin
    lRightBox = Splitter.Left
    lBottom = mp_oCurrentView.TimeLine.Bottom
    clsG.DrawEdge lLeftBox, lTop, lRightBox, lBottom, GRE_COLORS.CLR_BUTTONFACE, GRE_BUTTONSTYLE.BT_NORMALWINDOWS, GRE_EDGETYPE.ET_RAISED, True
    clsG.DrawTextEx lLeftBox, lTop, lRightBox, lBottom, "Column", DT_SINGLELINE Or DT_CENTER Or DT_VCENTER, GRE_COLORS.CLR_BLACK, mp_oFont
    '// Rows
    mp_oCurrentView.ClientArea.FirstVisibleRow = 1
    lLeftBox = mt_LeftMargin
    lTop = mp_oCurrentView.ClientArea.Top
    lRightBox = Splitter.Left
    lBottom = mp_oCurrentView.ClientArea.Top + 40
    lLeftCA = Splitter.Right
    lRightCA = mt_RightMargin
    clsG.DrawEdge lLeftBox, lTop, lRightBox, lBottom, GRE_COLORS.CLR_BUTTONFACE, GRE_BUTTONSTYLE.BT_NORMALWINDOWS, GRE_EDGETYPE.ET_RAISED, True
    clsG.DrawTextEx lLeftBox, lTop, lRightBox, lBottom, "Cell", DT_SINGLELINE Or DT_CENTER Or DT_VCENTER, GRE_COLORS.CLR_BLACK, mp_oFont
    If mp_oCurrentView.ClientArea.Grid.HorizontalLines = True Then
        clsG.DrawLine lLeftCA, lBottom, lRightCA, lBottom, GRE_LINETYPE.LT_NORMAL, mp_oCurrentView.ClientArea.Grid.Color, GRE_LINEDRAWSTYLE.LDS_SOLID
    End If
    Rows.TopOffset = mp_oCurrentView.ClientArea.Top + 40
    mp_oCurrentView.ClientArea.f_LastVisibleRow = 1
    '// ADD
    Splitter.Draw
    mp_DrawControlBorder
End Sub

'//  ----------------------------------------------------------------------------------------
'//  ActiveGantt Object Drawing Functions
'//  ----------------------------------------------------------------------------------------



Private Sub mp_DrawDebugMetrics()

End Sub

Private Sub mp_DrawControlBorder()
    clsG.ClipRegion 0, 0, clsG.Width, clsG.Height, True
    Select Case mp_yBorderStyle
        Case E_BORDERSTYLE.TLB_SINGLE
            clsG.DrawLine 0, 0, clsG.Width - 1, clsG.Height - 1, GRE_LINETYPE.LT_BORDER, GRE_COLORS.CLR_BLACK, GRE_LINEDRAWSTYLE.LDS_SOLID
        Case E_BORDERSTYLE.TLB_3D
            clsG.DrawEdge 0, 0, clsG.Width - 1, clsG.Height - 1, GRE_COLORS.CLR_BLACK, GRE_BUTTONSTYLE.BT_NORMALWINDOWS, GRE_EDGETYPE.ET_SUNKEN, False
    End Select
End Sub

'//  ----------------------------------------------------------------------------------------
'//  ActiveGantt Drawing Functions
'//  ----------------------------------------------------------------------------------------

Friend Sub mp_DrawItemI(ByRef oMilestone As clsMilestone, ByVal sStyleIndex As String)
    Dim oStyle As clsStyle
    Dim oMilestoneStyle As clsMilestoneStyle
    If clsS.StrIsNumeric(sStyleIndex) Then
        If CLng(sStyleIndex) < 0 Or CLng(sStyleIndex) > Styles.Count Then
            mp_ErrorReport 50238, "Style object element not found when preparing to draw, invalid index", "ActiveGanttVBCtl.mp_DrawItemI"
            Exit Sub
        End If
    Else
        If Styles.oCollection.m_bDoesKeyExist(sStyleIndex) = False Then
            mp_ErrorReport 50239, "Style object element not found when preparing to draw, invalid key", "ActiveGanttVBCtl.mp_DrawItemI"
            Exit Sub
        End If
    End If
    Set oStyle = Styles.FItem(sStyleIndex)
    Select Case oStyle.Appearance
        Case E_STYLEAPPEARANCE.SA_FLAT, E_STYLEAPPEARANCE.SA_CELL, E_STYLEAPPEARANCE.SA_RAISED, E_STYLEAPPEARANCE.SA_SUNKEN
            Set oMilestoneStyle = oStyle.MilestoneStyle
            clsG.DrawFigure clsM.GetXCoordinateFromDate(oMilestone.MilestoneDate), oMilestone.Top, oMilestone.Bottom - oMilestone.Top, oMilestone.Bottom - oMilestone.Top, oMilestoneStyle.ShapeIndex, oMilestoneStyle.BorderColor, oMilestoneStyle.FillColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        Case E_STYLEAPPEARANCE.SA_GRAPHICAL
            clsG.DrawPicture oMilestone.Picture, oStyle.PictureAlignmentHorizontal, oStyle.PictureAlignmentVertical, oStyle.PictureXMargin, oStyle.PictureYMargin, oMilestone.Left, oMilestone.Right, oMilestone.Top, oMilestone.Bottom, oStyle.UseMask
        Case Else
            mp_ErrorReport 50388, "Invalid Style appearance when preparing to draw Milestone", "ActiveGanttVBCtl.mp_DrawItemI"
            Exit Sub
    End Select
    mp_DrawItemCaption oMilestone.Left, oMilestone.Top, oMilestone.Right, oMilestone.Bottom, oMilestone.LeftTrim, oMilestone.RightTrim, oStyle, oMilestone.Caption
End Sub

Friend Sub mp_DrawItem(ByVal v_lLeft As Long, ByVal v_lRight As Long, ByVal v_lTop As Long, ByVal v_lBottom As Long, ByVal sStyleIndex As String, ByVal sCaption As String, ByVal v_bIsSelected As Boolean, ByRef v_oPicture As StdPicture, ByVal v_lLeftTrim As Long, ByVal v_lRightTrim As Long, ByRef v_oStyle As clsStyle)
    Dim oStyle As clsStyle
    Dim oTaskStyle As clsTaskStyle
    If (v_oStyle Is Nothing) Then
        If clsS.StrIsNumeric(sStyleIndex) Then
            If CLng(sStyleIndex) < 0 Or CLng(sStyleIndex) > Styles.Count Then
                mp_ErrorReport 50238, "Style object element not found when preparing to draw, invalid index", "ActiveGanttVBCtl.mp_DrawItem"
                Exit Sub
            End If
        Else
            If Styles.oCollection.m_bDoesKeyExist(sStyleIndex) = False Then
                mp_ErrorReport 50239, "Style object element not found when preparing to draw, invalid key", "ActiveGanttVBCtl.mp_DrawItem"
                Exit Sub
            End If
        End If
        Set oStyle = Styles.FItem(sStyleIndex)
    Else
        Set oStyle = v_oStyle
    End If
    Set oTaskStyle = oStyle.TaskStyle
    Select Case oStyle.Appearance
        Case E_STYLEAPPEARANCE.SA_RAISED
            clsG.DrawEdge v_lLeft, v_lTop, v_lRight, v_lBottom, oStyle.BackColor, oStyle.ButtonStyle, GRE_EDGETYPE.ET_RAISED, True
        Case E_STYLEAPPEARANCE.SA_SUNKEN
            clsG.DrawEdge v_lLeft, v_lTop, v_lRight, v_lBottom, oStyle.BackColor, oStyle.ButtonStyle, GRE_EDGETYPE.ET_SUNKEN, True
        Case E_STYLEAPPEARANCE.SA_FLAT
            If (oStyle.BackgroundPattern = GRE_BACKGROUNDPATTERN.FP_SOLID) Then
                clsG.DrawLine v_lLeft, v_lTop, v_lRight, v_lBottom, GRE_LINETYPE.LT_FILLED, oStyle.BackColor, GRE_LINEDRAWSTYLE.LDS_SOLID
            ElseIf (oStyle.BackgroundPattern = GRE_BACKGROUNDPATTERN.FP_GRADIENT) Then
                clsG.GradientFill v_lLeft, v_lTop, v_lRight, v_lBottom, oStyle.StartGradientColor, oStyle.EndGradientColor, oStyle.GradientFillMode
            Else
                clsG.DrawHatch v_lLeft, v_lTop, v_lRight, v_lBottom, oStyle.BackColor, oStyle.BackgroundPattern, oStyle.HatchFactor
            End If
            If oStyle.BorderStyle = E_STYLEBORDER.SBR_SINGLE Then
                clsG.DrawLine v_lLeft, v_lTop, v_lRight, v_lBottom, GRE_LINETYPE.LT_BORDER, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID
            End If
            clsG.ClipRegion Splitter.Right, v_lTop, mt_RightMargin, v_lBottom + 1, False
            clsG.DrawFigure v_lRight, v_lTop, v_lBottom - v_lTop, v_lBottom - v_lTop, oTaskStyle.EndShapeIndex, oTaskStyle.EndBorderColor, oTaskStyle.EndFillColor, LDS_SOLID
            clsG.DrawFigure v_lLeft, v_lTop, v_lBottom - v_lTop, v_lBottom - v_lTop, oTaskStyle.StartShapeIndex, oTaskStyle.StartBorderColor, oTaskStyle.StartFillColor, LDS_SOLID
            clsG.RestorePreviousClipRegion
        Case E_STYLEAPPEARANCE.SA_CELL
            clsG.DrawLine v_lLeft, v_lTop, v_lRight, v_lBottom, GRE_LINETYPE.LT_FILLED, oStyle.BackColor, GRE_LINEDRAWSTYLE.LDS_SOLID
            clsG.DrawLine v_lLeft, v_lBottom, v_lRight, v_lBottom, GRE_LINETYPE.LT_NORMAL, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        Case E_STYLEAPPEARANCE.SA_GRAPHICAL
            Dim lPictureHeight As Long
            Dim lPictureWidth As Long
            lPictureHeight = mp_lPXH(oTaskStyle.MiddlePicture.Height)
            lPictureWidth = mp_lPXW(oTaskStyle.MiddlePicture.Width)
            clsG.TilePictureHorizontal oTaskStyle.MiddlePicture.Handle, v_lLeft, v_lTop, v_lRight, v_lBottom, oStyle.UseMask
            '// Exit if the start and end sections don't fit
            If (v_lRight - v_lLeft) > (lPictureWidth * 2) Then
                clsG.PaintPicture oTaskStyle.StartPicture.Handle, v_lLeft, v_lTop, v_lLeft + lPictureWidth, v_lTop + lPictureHeight, 0, 0, oStyle.UseMask
                clsG.PaintPicture oTaskStyle.EndPicture.Handle, v_lRight - lPictureWidth, v_lTop, v_lRight, v_lTop + lPictureHeight, 0, 0, oStyle.UseMask
            End If
    End Select
    If Not (v_oPicture Is Nothing) Then
        clsG.DrawPicture v_oPicture, oStyle.PictureAlignmentHorizontal, oStyle.PictureAlignmentVertical, oStyle.PictureXMargin, oStyle.PictureYMargin, v_lLeft, v_lRight, v_lTop, v_lBottom, oStyle.UseMask
    End If
    mp_DrawItemCaption v_lLeft, v_lTop, v_lRight, v_lBottom, v_lLeftTrim, v_lRightTrim, oStyle, sCaption
    If oStyle.SelectionRectangleVisible = True And v_bIsSelected Then
        clsG.DrawFocusRectangle v_lLeft + oStyle.SelectionRectangleOffsetLeft, v_lTop + oStyle.SelectionRectangleOffsetTop, v_lRight - oStyle.SelectionRectangleOffsetRight, v_lBottom - oStyle.SelectionRectangleOffsetBottom
    End If
End Sub

Friend Sub mp_DrawItemCaption(ByVal v_lLeft As Long, ByVal v_lTop As Long, ByVal v_lRight As Long, ByVal v_lBottom As Long, ByVal v_lLeftTrim As Long, ByVal v_lRightTrim As Long, ByRef oStyle As clsStyle, ByVal sCaption As String)
    Dim lTextLeft As Long
    Dim lTextRight As Long
    Dim lTextTop As Long
    Dim lTextBottom As Long
    If oStyle.CaptionVisible = False Then
        Exit Sub
    End If
    If sCaption = "" Then
        Exit Sub
    End If
    Select Case oStyle.CaptionPlacement
        Case E_CAPTIONPLACEMENT.SCP_OBJECTEXTENTSPLACEMENT
            If (oStyle.DrawCaptionInVisibleArea = False) Then
                lTextLeft = v_lLeft
                lTextRight = v_lRight
            Else
                lTextLeft = v_lLeftTrim
                lTextRight = v_lRightTrim
            End If
            lTextTop = v_lTop
            lTextBottom = v_lBottom
            If oStyle.CaptionAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT Then
                lTextLeft = v_lLeft + oStyle.CaptionXMargin
            End If
            If oStyle.CaptionAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT Then
                lTextRight = v_lRight - oStyle.CaptionXMargin
            End If
            If oStyle.CaptionAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP Then
                lTextTop = v_lTop + oStyle.CaptionYMargin
            End If
            If oStyle.CaptionAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM Then
                lTextBottom = v_lBottom - oStyle.CaptionYMargin
            End If
            clsG.TextOutEx lTextLeft, lTextTop, lTextRight, lTextBottom, sCaption, oStyle.CaptionAlignmentHorizontal, oStyle.CaptionAlignmentVertical, oStyle.ForeColor, oStyle.Font, oStyle.ClipCaption
        Case E_CAPTIONPLACEMENT.SCP_OFFSETPLACEMENT
            clsG.DrawTextEx v_lLeft + oStyle.CaptionOffsetLeft, v_lTop + oStyle.CaptionOffsetTop, v_lRight - oStyle.CaptionOffsetRight, v_lBottom - oStyle.CaptionOffsetBottom, sCaption, oStyle.CaptionFlags, oStyle.ForeColor, oStyle.Font
        Case E_CAPTIONPLACEMENT.SCP_EXTERIORPLACEMENT
            If oStyle.CaptionAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT Then
                lTextLeft = v_lLeft - mp_lStrWidth(sCaption, oStyle.Font) - oStyle.CaptionXMargin
                lTextRight = v_lLeft - oStyle.CaptionXMargin
            End If
            If oStyle.CaptionAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT Then
                lTextLeft = v_lRight + oStyle.CaptionXMargin
                lTextRight = v_lRight + mp_lStrWidth(sCaption, oStyle.Font) + oStyle.CaptionXMargin
            End If
            If oStyle.CaptionAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_CENTER Then
                lTextLeft = v_lLeft
                lTextRight = v_lRight
            End If
            If oStyle.CaptionAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP Then
                lTextTop = v_lTop - mp_lStrHeight(sCaption, oStyle.Font) - oStyle.CaptionYMargin
                lTextBottom = v_lTop - oStyle.CaptionYMargin
            End If
            If oStyle.CaptionAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM Then
                lTextTop = v_lBottom + oStyle.CaptionYMargin
                lTextBottom = v_lBottom + mp_lStrHeight(sCaption, oStyle.Font) + oStyle.CaptionYMargin
            End If
            If oStyle.CaptionAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_CENTER Then
                lTextTop = v_lTop
                lTextBottom = v_lBottom
            End If
            clsG.ClipRegion Splitter.Right, mp_oCurrentView.ClientArea.Top, mt_RightMargin, mp_oCurrentView.ClientArea.Bottom, False
            clsG.TextOutEx lTextLeft, lTextTop, lTextRight, lTextBottom, sCaption, GRE_HORIZONTALALIGNMENT.HAL_CENTER, GRE_VERTICALALIGNMENT.VAL_CENTER, oStyle.ForeColor, oStyle.Font, oStyle.ClipCaption
            clsG.RestorePreviousClipRegion
    End Select
End Sub

Friend Function mp_bPositionItem(ByRef r_lTop As Long, ByRef r_lBottom As Long, ByRef v_oStyle As clsStyle) As Boolean
    If v_oStyle.Placement = PLC_ROWEXTENTSPLACEMENT Or v_oStyle.Appearance = SA_CELL Then
        mp_bPositionItem = True
        Exit Function
    End If
    If v_oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT Then
        Dim lTop As Long
        Dim lBottom As Long
        lTop = r_lTop
        lBottom = r_lBottom
        r_lTop = r_lTop + v_oStyle.OffsetTop
        r_lBottom = r_lTop + v_oStyle.OffsetBottom
        If Not (r_lTop > lTop And r_lTop < lBottom) Then
            mp_bPositionItem = False
            Exit Function
        End If
        mp_bPositionItem = True
    End If
End Function

'//******************************************************************************************
'// DATE RELATED AUXILIARY FUNCTIONS
'//******************************************************************************************

Friend Function mp_bDetectConflict(ByVal StartDate As Date, ByVal EndDate As Date, ByVal RowKey As String, ByVal ExcludeIndex As Long, ByVal LayerIndex As String, Optional ByVal CompareType As E_TYPE = 0) As Boolean
    Dim oTask As clsTask
    Dim oMilestone As clsMilestone
    Dim lIndex As Long
    Dim lLayerIndex As Long
    lLayerIndex = Layers.oCollection.m_lReturnIndex(LayerIndex, True)
    If (lLayerIndex = -1) Then
        mp_ErrorReport 50000, "Invalid Layer Index", "ActiveGanttVBCtl.mp_bDetectConflict"
        mp_bDetectConflict = False
        Exit Function
    End If
    If CompareType = E_TYPE.TP_TASK Then
        For lIndex = 1 To Tasks.Count
            Set oTask = Tasks.oCollection.m_oReturnArrayElement(lIndex)
            If RowKey = oTask.RowKey And ExcludeIndex <> lIndex Then
                If (Layers.oCollection.m_lReturnIndex(oTask.LayerIndex, True) = lLayerIndex) Then
                    '// oTask              S------------------E
                    '// interval                S------------------E
                    If StartDate = oTask.StartDate Or EndDate = oTask.EndDate Then
                        mp_bDetectConflict = True
                        Exit Function
                    End If
                    '// oTask              S------------------E
                    '// interval                             S------------------E
                    If StartDate > oTask.StartDate And StartDate < oTask.EndDate Then
                        mp_bDetectConflict = True
                        Exit Function
                    End If
                    '// oTask              S------------------E
                    '// interval        S------------------E
                    If EndDate > oTask.StartDate And EndDate < oTask.EndDate Then
                        mp_bDetectConflict = True
                        Exit Function
                    End If
                    '// oTask              S------------------E
                    '// interval             S-------------------------E
                    If StartDate < oTask.StartDate And EndDate > oTask.EndDate Then
                        mp_bDetectConflict = True
                        Exit Function
                    End If
                    '// oTask         S--------------------------E
                    '// interval                     S---------E
                    If StartDate > oTask.StartDate And EndDate < oTask.EndDate Then
                        mp_bDetectConflict = True
                        Exit Function
                    End If
                End If
            End If
        Next lIndex
        For lIndex = 1 To Milestones.Count
            Set oMilestone = Milestones.oCollection.m_oReturnArrayElement(lIndex)
            If oMilestone.RowKey = RowKey Then
                If (Layers.oCollection.m_lReturnIndex(oMilestone.LayerIndex, True) = lLayerIndex) Then
                    If StartDate < oMilestone.MilestoneDate And EndDate > oMilestone.MilestoneDate Then
                        mp_bDetectConflict = True
                        Exit Function
                    End If
                    If EndDate < oMilestone.MilestoneDate And StartDate > oMilestone.MilestoneDate Then
                        mp_bDetectConflict = True
                        Exit Function
                    End If
                End If
            End If
        Next lIndex
    Else
        For lIndex = 1 To Milestones.Count
            Set oMilestone = Milestones.oCollection.m_oReturnArrayElement(lIndex)
            If oMilestone.RowKey = RowKey And ExcludeIndex <> lIndex Then
                If (Layers.oCollection.m_lReturnIndex(oMilestone.LayerIndex, True) = lLayerIndex) Then
                    If StartDate = oMilestone.MilestoneDate Then
                        mp_bDetectConflict = True
                        Exit Function
                    End If
                End If
            End If
        Next lIndex
        For lIndex = 1 To Tasks.Count
            Set oTask = Tasks.oCollection.m_oReturnArrayElement(lIndex)
            If oTask.RowKey = RowKey Then
                If (Layers.oCollection.m_lReturnIndex(oTask.LayerIndex, True) = lLayerIndex) Then
                    If oTask.StartDate < StartDate And oTask.EndDate > StartDate Then
                        mp_bDetectConflict = True
                        Exit Function
                    End If
                    If oTask.EndDate < StartDate And oTask.StartDate > StartDate Then
                        mp_bDetectConflict = True
                        Exit Function
                    End If
                End If
            End If
        Next lIndex
    End If
    mp_bDetectConflict = False
End Function

'//******************************************************************************************
'// ERROR REPORTING AUXILIARY FUNCTIONS
'//******************************************************************************************

Friend Sub mp_ErrorReport(ByVal ErrNumber As Long, ByVal ErrDescription As String, ByVal ErrSource As String)
    Select Case mp_yErrorReports
        Case E_REPORTERRORS.RE_MSGBOX
            MsgBox ErrNumber & ": " & ErrDescription, vbOKOnly, ErrSource
        Case E_REPORTERRORS.RE_HIDE
        Case E_REPORTERRORS.RE_RAISE
            If ErrNumber > 50000 Then
                Err.Raise ErrNumber, , ErrDescription, "agvb20.chm", ErrNumber - 50000
            Else
                Err.Raise ErrNumber, , ErrDescription
            End If
        Case E_REPORTERRORS.RE_RAISEEVENT
            RaiseEvent ActiveGanttError(ErrNumber, ErrDescription, ErrSource)
    End Select
End Sub

'//  ----------------------------------------------------------------------------------------
'//  Metrics
'//  ----------------------------------------------------------------------------------------

Friend Property Get mt_BorderThickness() As Integer
    mt_BorderThickness = mp_yBorderStyle
End Property

Friend Property Get mt_TableBottom() As Long
    If oHScrollBar1.State = E_SCROLLSTATE.SS_SHOWN Then
        mt_TableBottom = clsG.Height - mp_yBorderStyle - 1 - oHScrollBar1.Height
    Else
        mt_TableBottom = clsG.Height - mp_yBorderStyle - 1
    End If
End Property

Friend Property Get mt_TopMargin() As Long
    mt_TopMargin = mp_yBorderStyle
End Property

Friend Property Get mt_LeftMargin() As Long
    mt_LeftMargin = mp_yBorderStyle
End Property

Friend Property Get mt_RightMargin() As Long
    If oVScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
        mt_RightMargin = clsG.Width - mp_yBorderStyle - 1 - oVScrollBar.Width
    Else
        mt_RightMargin = clsG.Width - mp_yBorderStyle - 1
    End If
End Property

Friend Property Get mt_BottomMargin() As Long
    mt_BottomMargin = clsG.Height - mp_yBorderStyle - 1
End Property

Private Function mp_lPXW(ByRef Width As Long) As Long
    mp_lPXW = ScaleX(Width, 8, 3)
End Function

Private Function mp_lPXH(ByRef Height As Long) As Long
    mp_lPXH = ScaleY(Height, 8, 3)
End Function


'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Private Sub mp_ReadXML(ByVal Path As String, Optional ByVal bReadFromVariable As Boolean = False)
    Dim oXML As New clsXML
    oXML.Initialize Me, "ActiveGanttCtl"
    Dim sCurrentView As String
    Dim sVersion As String
    If bReadFromVariable = False Then
        oXML.ReadXML Path
    Else
        oXML.SetXML mp_sXML
    End If
    Clear
    oXML.InitializeReader
    oXML.ReadProperty "Version", sVersion

    oXML.ReadProperty "AddMode", mp_yAddMode
    oXML.ReadProperty "AllowAdd", mp_bAllowAdd
    oXML.ReadProperty "AllowEdit", mp_bAllowEdit
    oXML.ReadProperty "AllowSplitterMove", mp_bAllowSplitterMove
    oXML.ReadProperty "AllowColumnSize", mp_bAllowColumnSize
    oXML.ReadProperty "AllowColumnSwap", mp_bAllowColumnSwap
    oXML.ReadProperty "AllowRowSize", mp_bAllowRowSize
    oXML.ReadProperty "AllowRowSwap", mp_bAllowRowSwap
    oXML.ReadProperty "AllowTimeLineScroll", mp_bAllowTimeLineScroll
    
    oXML.ReadProperty "BackColor", UserControl.BackColor
    oXML.ReadProperty "BorderStyle", mp_yBorderStyle
    
    oXML.ReadProperty "CurrentLayer", mp_sCurrentLayer
    oXML.ReadProperty "CurrentView", sCurrentView
    
    
    oXML.ReadProperty "EditMode", mp_yEditMode
    oXML.ReadProperty "EnableObjects", mp_yEnableObjects
    oXML.ReadProperty "ErrorReports", mp_yErrorReports
    
    oXML.ReadProperty "FlickerFree", mp_bFlickerFree
    oXML.ReadPropertyFont "Font", mp_oFont
    oXML.ReadProperty "FontCharWidth", mp_lFontCharWidth
    
    
    
    oXML.ReadProperty "MinColumnWidth", mp_lMinColumnWidth
    oXML.ReadProperty "MinRowHeight", mp_lMinRowHeight
    
    
    oXML.ReadProperty "ScrollBarBehaviour", mp_yScrollBarBehaviour
    oXML.ReadProperty "ScrollBarsVisible", mp_bScrollBarsVisible
    oXML.ReadProperty "SelectedTaskIndex", mp_lSelectedTaskIndex
    oXML.ReadProperty "SelectedMilestoneIndex", mp_lSelectedMilestoneIndex
    oXML.ReadProperty "SelectedColumnIndex", mp_lSelectedColumnIndex
    oXML.ReadProperty "SelectedRowIndex", mp_lSelectedRowIndex
    oXML.ReadProperty "SelectedCellIndex", mp_lSelectedCellIndex
    
    
    oXML.ReadProperty "TimeBlockBehaviour", mp_yTimeBlockBehaviour
    
    
    Rows.SetXML oXML.ReadObject("Rows")
    Columns.SetXML oXML.ReadObject("Columns")
    DefaultValues.SetXML oXML.ReadObject("DefaultValues")
    Styles.SetXML oXML.ReadObject("Styles")
    Milestones.SetXML oXML.ReadObject("Milestones")
    Tasks.SetXML oXML.ReadObject("Tasks")
    Views.SetXML oXML.ReadObject("Views")
    TimeBlocks.SetXML oXML.ReadObject("TimeBlocks")
    Percentages.SetXML oXML.ReadObject("Percentages")
    PercentageGroups.SetXML oXML.ReadObject("PercentageGroups")
    Splitter.SetXML oXML.ReadObject("Splitter")

    CurrentView = sCurrentView
    
    mp_oCurrentView.TimeLine.Position mp_oCurrentView.TimeLine.StartDate
End Sub

Public Sub ReadXML(ByVal Path As String)
    mp_ReadXML Path
End Sub

Private Sub mp_WriteXML(ByVal Path As String, Optional ByVal bWriteToVariable As Boolean = False)
    Dim oXML As New clsXML
    oXML.Initialize Me, "ActiveGanttCtl"
    oXML.InitializeWriter
    oXML.WriteProperty "Version", "AGVB"
    oXML.WriteProperty "AddMode", mp_yAddMode
    oXML.WriteProperty "AllowAdd", mp_bAllowAdd
    oXML.WriteProperty "AllowEdit", mp_bAllowEdit
    oXML.WriteProperty "AllowSplitterMove", mp_bAllowSplitterMove
    oXML.WriteProperty "AllowColumnSize", mp_bAllowColumnSize
    oXML.WriteProperty "AllowColumnSwap", mp_bAllowColumnSwap
    oXML.WriteProperty "AllowRowSize", mp_bAllowRowSize
    oXML.WriteProperty "AllowRowSwap", mp_bAllowRowSwap
    oXML.WriteProperty "AllowTimeLineScroll", mp_bAllowTimeLineScroll
    oXML.WriteProperty "BackColor", UserControl.BackColor
    oXML.WriteProperty "BorderStyle", mp_yBorderStyle
    
    oXML.WriteProperty "CurrentLayer", mp_sCurrentLayer
    oXML.WriteProperty "CurrentView", mp_sCurrentView
    
    oXML.WriteProperty "EditMode", mp_yEditMode
    oXML.WriteProperty "EnableObjects", mp_yEnableObjects
    oXML.WriteProperty "ErrorReports", mp_yErrorReports
    
    oXML.WriteProperty "FlickerFree", mp_bFlickerFree
    oXML.WritePropertyFont "Font", mp_oFont
    oXML.WriteProperty "FontCharWidth", mp_lFontCharWidth
    
    oXML.WriteProperty "MinColumnWidth", mp_lMinColumnWidth
    oXML.WriteProperty "MinRowHeight", mp_lMinRowHeight
    
    oXML.WriteProperty "ScrollBarBehaviour", mp_yScrollBarBehaviour
    oXML.WriteProperty "ScrollBarsVisible", mp_bScrollBarsVisible
    oXML.WriteProperty "SelectedTaskIndex", mp_lSelectedTaskIndex
    oXML.WriteProperty "SelectedMilestoneIndex", mp_lSelectedMilestoneIndex
    oXML.WriteProperty "SelectedColumnIndex", mp_lSelectedColumnIndex
    oXML.WriteProperty "SelectedRowIndex", mp_lSelectedRowIndex
    oXML.WriteProperty "SelectedCellIndex", mp_lSelectedCellIndex
    
    oXML.WriteProperty "TimeBlockBehaviour", mp_yTimeBlockBehaviour
    
    oXML.WriteObject Rows.GetXML
    oXML.WriteObject Columns.GetXML
    oXML.WriteObject DefaultValues.GetXML
    oXML.WriteObject Styles.GetXML
    oXML.WriteObject Milestones.GetXML
    oXML.WriteObject Tasks.GetXML
    oXML.WriteObject Views.GetXML
    oXML.WriteObject TimeBlocks.GetXML
    oXML.WriteObject Percentages.GetXML
    oXML.WriteObject PercentageGroups.GetXML
    oXML.WriteObject Splitter.GetXML
    

    
    If bWriteToVariable = False Then
        oXML.WriteXML Path
    Else
        mp_sXML = oXML.GetXML
    End If
End Sub

Public Sub WriteXML(ByVal Path As String)
    mp_WriteXML Path
End Sub

Friend Property Get f_MouseEvents() As clsMouseEvents
    Set f_MouseEvents = mp_oMouseEvents
End Property

Public Property Get Font() As Font
    Set Font = mp_oFont
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get AllowAdd() As Boolean
    AllowAdd = mp_bAllowAdd
End Property

Public Property Let AllowAdd(ByVal Value As Boolean)
    mp_bAllowAdd = Value
    PropertyChanged "AllowAdd"
End Property

Public Property Get AllowEdit() As Boolean
    AllowEdit = mp_bAllowEdit
End Property

Public Property Let AllowEdit(ByVal Value As Boolean)
    mp_bAllowEdit = Value
    PropertyChanged "AllowEdit"
End Property

Public Property Get AllowSplitterMove() As Boolean
Attribute AllowSplitterMove.VB_Description = "Returns or sets a boolean value that determines whether the user can change the width of the fixed column. "
Attribute AllowSplitterMove.VB_HelpID = 33
    AllowSplitterMove = mp_bAllowSplitterMove
End Property

Public Property Let AllowSplitterMove(ByVal Value As Boolean)
    mp_bAllowSplitterMove = Value
    PropertyChanged "AllowSplitterMove"
End Property

Public Property Get AllowColumnSize() As Boolean
Attribute AllowColumnSize.VB_Description = "Returns or sets a boolean value that determines whether the user can change the width of Column objects. "
Attribute AllowColumnSize.VB_HelpID = 143
    AllowColumnSize = mp_bAllowColumnSize
End Property

Public Property Let AllowColumnSize(ByVal Value As Boolean)
    mp_bAllowColumnSize = Value
    PropertyChanged "AllowColumnSize"
End Property

Public Property Get AllowRowSize() As Boolean
Attribute AllowRowSize.VB_Description = "Returns or sets a boolean value that determines whether the user can change the height of the Row objects. "
Attribute AllowRowSize.VB_HelpID = 36
    AllowRowSize = mp_bAllowRowSize
End Property

Public Property Let AllowRowSize(ByVal Value As Boolean)
    mp_bAllowRowSize = Value
    PropertyChanged "AllowRowSize"
End Property

Public Property Get AllowRowSwap() As Boolean
Attribute AllowRowSwap.VB_Description = "Returns or sets a boolean value that determines whether the user can change the vertical order of the Row objects. "
Attribute AllowRowSwap.VB_HelpID = 37
    AllowRowSwap = mp_bAllowRowSwap
End Property

Public Property Let AllowRowSwap(ByVal Value As Boolean)
    mp_bAllowRowSwap = Value
    PropertyChanged "AllowRowSwap"
End Property

Public Property Get AllowColumnSwap() As Boolean
    AllowColumnSwap = mp_bAllowColumnSwap
End Property

Public Property Let AllowColumnSwap(ByVal Value As Boolean)
    mp_bAllowColumnSwap = Value
    PropertyChanged "AllowColumnSwap"
End Property

Public Property Get AllowTimeLineScroll() As Boolean
Attribute AllowTimeLineScroll.VB_Description = "Returns or sets a Boolean value that determines whether the user can move the time line by dragging it."
Attribute AllowTimeLineScroll.VB_HelpID = 405
    AllowTimeLineScroll = mp_bAllowTimeLineScroll
End Property

Public Property Let AllowTimeLineScroll(ByVal Value As Boolean)
    mp_bAllowTimeLineScroll = Value
    PropertyChanged "AllowTimeLineScroll"
End Property

Public Property Get ScrollBarsVisible() As Boolean
    ScrollBarsVisible = mp_bScrollBarsVisible
End Property

Public Property Let ScrollBarsVisible(ByVal Value As Boolean)
    mp_bScrollBarsVisible = Value
    PropertyChanged "ScrollBarsVisible"
End Property

Public Property Get AddMode() As E_ADDMODE
    AddMode = mp_yAddMode
End Property

Public Property Let AddMode(ByVal Value As E_ADDMODE)
    mp_yAddMode = Value
    PropertyChanged "AddMode"
End Property

Public Property Get EditMode() As E_EDITMODE
    EditMode = mp_yEditMode
End Property

Public Property Let EditMode(ByVal Value As E_EDITMODE)
    mp_yEditMode = Value
    PropertyChanged "EditMode"
End Property

Public Property Get ErrorReports() As E_REPORTERRORS
    ErrorReports = mp_yErrorReports
End Property

Public Property Let ErrorReports(ByVal Value As E_REPORTERRORS)
    mp_yErrorReports = Value
    PropertyChanged "ErrorReports"
End Property

Public Property Get OLEDragMode() As E_DROPMODE
    OLEDragMode = mp_yOLEDragMode
End Property

Public Property Let OLEDragMode(ByVal Value As E_DROPMODE)
    mp_yOLEDragMode = Value
    PropertyChanged "OLEDragMode"
End Property

Public Property Get OLEDropMode() As E_DROPMODE
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal Value As E_DROPMODE)
    UserControl.OLEDropMode = Value
    PropertyChanged "OLEDropMode"
End Property

Public Property Get EnableObjects() As E_ENABLEOBJECTS
    EnableObjects = mp_yEnableObjects
End Property

Public Property Let EnableObjects(ByVal Value As E_ENABLEOBJECTS)
    mp_yEnableObjects = Value
    PropertyChanged "EnableObjects"
End Property

Public Property Get BorderStyle() As E_BORDERSTYLE
    BorderStyle = mp_yBorderStyle
End Property

Public Property Let BorderStyle(ByVal Value As E_BORDERSTYLE)
    mp_yBorderStyle = Value
    PropertyChanged "BorderStyle"
End Property

Public Property Get ScrollBarBehaviour() As E_SCROLLBEHAVIOUR
    ScrollBarBehaviour = mp_yScrollBarBehaviour
End Property

Public Property Let ScrollBarBehaviour(ByVal Value As E_SCROLLBEHAVIOUR)
    mp_yScrollBarBehaviour = Value
    PropertyChanged "ScrollBarBehaviour"
End Property

Public Property Get TimeBlockBehaviour() As E_TIMEBLOCKBEHAVIOUR
    TimeBlockBehaviour = mp_yTimeBlockBehaviour
End Property

Public Property Let TimeBlockBehaviour(ByVal Value As E_TIMEBLOCKBEHAVIOUR)
    mp_yTimeBlockBehaviour = Value
    PropertyChanged "TimeBlockBehaviour"
End Property

Public Property Get CurrentLayer() As String
    CurrentLayer = mp_sCurrentLayer
End Property

Public Property Let CurrentLayer(ByVal Value As String)
    mp_sCurrentLayer = Value
    PropertyChanged "CurrentLayer"
End Property

Public Property Get CurrentView() As String
    CurrentView = mp_sCurrentView
End Property

Public Property Let CurrentView(ByVal Value As String)
    If Value = "" Then
        Value = "0"
    End If
    Set mp_oCurrentView = Views.FItem(Value)
    mp_sCurrentView = Value
    PropertyChanged "CurrentView"
End Property

Public Property Get CurrentViewObject() As clsView
    Set CurrentViewObject = mp_oCurrentView
End Property

Public Property Get MinRowHeight() As Long
    MinRowHeight = mp_lMinRowHeight
End Property

Public Property Let MinRowHeight(ByVal Value As Long)
    mp_lMinRowHeight = Value
    PropertyChanged "MinRowHeight"
End Property

Public Property Get MinColumnWidth() As Long
    MinColumnWidth = mp_lMinColumnWidth
End Property

Public Property Let MinColumnWidth(ByVal Value As Long)
    mp_lMinColumnWidth = Value
    PropertyChanged "MinColumnWidth"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns or sets a long integer value that specifies the background color of the ActiveGantt Schedule Control. "
Attribute BackColor.VB_HelpID = 39
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
    UserControl.BackColor = clsG.ConvertColor(Value)
    PropertyChanged "BackColor"
End Property

Public Property Get SelectedTaskIndex() As Long
    SelectedTaskIndex = mp_lSelectedTaskIndex
End Property

Public Property Let SelectedTaskIndex(ByVal Value As Long)
    If Value < 1 Or Value > Tasks.Count Then Value = Tasks.Count
    mp_lSelectedTaskIndex = Value
End Property

Public Property Get SelectedMilestoneIndex() As Long
    SelectedMilestoneIndex = mp_lSelectedMilestoneIndex
End Property

Public Property Let SelectedMilestoneIndex(ByVal Value As Long)
    If Value < 1 Or Value > Milestones.Count Then Value = Milestones.Count
    mp_lSelectedMilestoneIndex = Value
End Property

Public Property Get SelectedColumnIndex() As Long
    SelectedColumnIndex = mp_lSelectedColumnIndex
End Property

Public Property Let SelectedColumnIndex(ByVal Value As Long)
    If Value < 1 Or Value > Columns.Count Then Value = Columns.Count
    mp_lSelectedColumnIndex = Value
End Property

Public Property Get SelectedRowIndex() As Long
    SelectedRowIndex = mp_lSelectedRowIndex
End Property

Public Property Let SelectedRowIndex(ByVal Value As Long)
    If Value < 1 Or Value > Rows.Count Then Value = Rows.Count
    mp_lSelectedRowIndex = Value
End Property

Public Property Get SelectedCellIndex() As Long
    SelectedCellIndex = mp_lSelectedCellIndex
End Property

Public Property Let SelectedCellIndex(ByVal Value As Long)
    If Value < 1 Or Value > Columns.Count Then Value = Columns.Count
    mp_lSelectedCellIndex = Value
End Property

Friend Sub mp_ProcessInterval(ByVal Value As String, ByRef Interval As String, ByRef Factor As Long)
    Dim i As Integer
    Dim sInterval As String
    Dim lFactor As Long
    sInterval = clsS.StrLowerCase(Value)
    i = 1
    Do While clsS.StrIsNumeric(clsS.StrMid(sInterval, i, 1))
        i = i + 1
    Loop
    lFactor = clsS.StrLeft(sInterval, i - 1)
    sInterval = clsS.StrRight(sInterval, clsS.StrLen(sInterval) - clsS.StrLen(CStr(lFactor)))
    If (Not (sInterval = "s" Or sInterval = "n" Or sInterval = "h" Or sInterval = "d" Or sInterval = "w" Or sInterval = "y" Or sInterval = "ww" Or sInterval = "m" Or sInterval = "q" Or sInterval = "yyyy")) Then
        mp_ErrorReport 50170, "Invalid Interval", "ActiveGanttVBCtl.mp_ProcessInterval"
        Exit Sub
    End If
    Interval = sInterval
    Factor = lFactor
End Sub

Public Sub ClearSelections()
    mp_lSelectedTaskIndex = 0
    mp_lSelectedMilestoneIndex = 0
    mp_lSelectedColumnIndex = 0
    mp_lSelectedRowIndex = 0
    mp_lSelectedCellIndex = 0
End Sub

Public Function InConflict(ByVal StartDate As Date, ByVal EndDate As Date, ByVal RowKey As String, Optional ByVal Layer As Long = 0) As Boolean
Attribute InConflict.VB_Description = "Returns a Boolean value that determines whether a range of dates are in conflict with an existing Task. "
Attribute InConflict.VB_HelpID = 131
    InConflict = mp_bDetectConflict(StartDate, EndDate, RowKey, 0, Layer)
End Function

Public Sub Redraw()
Attribute Redraw.VB_Description = "Forces a repaint of the control. "
Attribute Redraw.VB_HelpID = 67
    mp_yDrawOperationType = E_DRAWOPTYPE.DOT_ALL
    UserControl_Paint
End Sub

Public Function GetDateFromXCoordinate(ByVal X As Long) As Date
    GetDateFromXCoordinate = clsM.GetDateFromXCoordinate(X)
End Function

Public Function GetXCoordinateFromDate(ByVal DatePosition As Date) As Long
    GetXCoordinateFromDate = clsM.GetXCoordinateFromDate(DatePosition)
End Function

Public Sub SaveToBMP(Path As String)
    SavePicture Image, Path
End Sub

'// OLE Drag Source Events *********************************************************************************************

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
'    AllowedEffects = 2
'    If mp_yDragType = E_DRAGTYPE.DragMoveTask Then
'        clsG.f_FocusLeft = 0
'        clsG.f_FocusRight = 0
'        clsG.f_FocusBottom = 0
'        clsG.f_FocusTop = 0
'        RaiseEvent OLETaskStartDrag(mp_lSelectedTaskIndex, Data, AllowedEffects)
'    ElseIf mp_yDragType = E_DRAGTYPE.DragMoveMilestone Then
'        clsG.f_FocusLeft = 0
'        clsG.f_FocusRight = 0
'        clsG.f_FocusBottom = 0
'        clsG.f_FocusTop = 0
'        RaiseEvent OLEMilestoneStartDrag(mp_lSelectedTaskIndex, Data, AllowedEffects)
'    End If
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
'    RaiseEvent OLECompleteDrag(Effect)
'    If mp_yDragType = E_DRAGTYPE.DragMoveTask Then
'        CancelUIOperations
'        Tasks.Remove mp_lSelectedTaskIndex
'        UserControl_Paint
'    ElseIf mp_yDragType = E_DRAGTYPE.DragMoveMilestone Then
'        CancelUIOperations
'        Milestones.Remove mp_lSelectedTaskIndex
'        UserControl_Paint
'    End If
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

'// OLE Drop Destination Events *****************************************************************************************

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
'    Dim lRowIndex As Long
'    If X > Splitter.Right And X < mt_RightMargin Then
'        lRowIndex = mp_lReturnRowIndexByPosition(Y)
'        If lRowIndex <> 0 Then
'            RaiseEvent OLEDragOver(Data, Effect, Button, Shift, lRowIndex, clsM.GetDateFromXCoordinate(X), State)
'        End If
'    End If
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim lRowIndex As Long
'    If X > Splitter.Right And X < mt_RightMargin Then
'        lRowIndex = mp_lReturnRowIndexByPosition(Y)
'        If lRowIndex <> 0 Then
'            RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, lRowIndex, clsM.GetDateFromXCoordinate(X))
'        End If
'    End If
End Sub


Public Sub Clear()
    Rows.Clear
    Styles.Clear
    Layers.Clear
    Columns.Clear
    TimeBlocks.Clear
    PercentageGroups.Clear
    Views.Clear
    Redraw
End Sub

Public Property Get Version() As String
    Version = CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
End Property

Public Property Get FontCharWidth() As Long
    FontCharWidth = mp_lFontCharWidth
End Property

Public Property Let FontCharWidth(ByVal Value As Long)
    mp_lFontCharWidth = Value
    PropertyChanged "FontCharWidth"
End Property

Public Function GetXML() As String
    mp_WriteXML "", True
    GetXML = mp_sXML
End Function

Public Sub SetXML(ByVal sXML As String)
    mp_sXML = sXML
    mp_ReadXML "", True
End Sub

Public Property Get FlickerFree() As Boolean
    FlickerFree = mp_bFlickerFree
End Property

Public Property Let FlickerFree(ByVal Value As Boolean)
    mp_bFlickerFree = Value
    PropertyChanged "FlickerFree"
End Property

Public Property Get TPCaption() As String
    TPCaption = mp_TPCaption
End Property

Public Property Let TPCaption(ByVal Value As String)
    mp_TPCaption = Value
    PropertyChanged "TPCaption"
End Property

Public Property Get TPDisplayTooltip() As Boolean
    TPDisplayTooltip = mp_TPDisplayToolTip
End Property

Public Property Let TPDisplayTooltip(ByVal Value As Boolean)
    mp_TPDisplayToolTip = Value
    PropertyChanged "TPDisplayTooltip"
End Property

Public Sub AboutBox()
    fAbout.Show 1, Me
End Sub



































