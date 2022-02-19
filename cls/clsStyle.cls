VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

'// Parent Control Pointer
Private mp_oControl As ActiveGanttVBCtl
'// clsItemBase
Private mp_sKey As String
Private mp_lIndex As Long
'// Object Member Variables

Private mp_oTaskStyle As clsTaskStyle
Private mp_oMilestoneStyle As clsMilestoneStyle
Private mp_oPredecessorStyle As clsPredecessorStyle


Private mp_bCaptionVisible As Boolean
Private mp_bClipCaption As Boolean
Private mp_bDrawCaptionInVisibleArea As Boolean
Private mp_bSelectionRectangleVisible As Boolean
Private mp_bUseMask As Boolean
Private mp_clrBackColor As OLE_COLOR
Private mp_clrForeColor As OLE_COLOR
Private mp_clrBorderColor As OLE_COLOR

Private mp_clrEndGradientColor As OLE_COLOR
Private mp_clrStartGradientColor As OLE_COLOR
Private mp_Font As New StdFont
Private mp_iHatchFactor As Integer

Private mp_lCaptionFlags As Long
Private mp_lCaptionOffsetBottom As Long
Private mp_lCaptionOffsetLeft As Long
Private mp_lCaptionOffsetRight As Long
Private mp_lCaptionOffsetTop As Long
Private mp_lCaptionXMargin As Long
Private mp_lCaptionYMargin As Long

Private mp_lOffsetBottom As Long
Private mp_lOffsetTop As Long
Private mp_lPictureXMargin As Long
Private mp_lPictureYMargin As Long
Private mp_lSelectionRectangleOffsetBottom As Long
Private mp_lSelectionRectangleOffsetLeft As Long
Private mp_lSelectionRectangleOffsetRight As Long
Private mp_lSelectionRectangleOffsetTop As Long

Private mp_sTag As String
Private mp_yAppearance As E_STYLEAPPEARANCE

Private mp_yBackgroundPattern As GRE_BACKGROUNDPATTERN
Private mp_yBorderStyle As E_STYLEBORDER
Private mp_yButtonStyle As GRE_BUTTONSTYLE
Private mp_yCaptionAlignmentHorizontal As GRE_HORIZONTALALIGNMENT
Private mp_yCaptionAlignmentVertical As GRE_VERTICALALIGNMENT
Private mp_yCaptionPlacement As E_CAPTIONPLACEMENT

Private mp_yGradientFillMode As GRE_GRADIENTFILLMODE
Private mp_yPictureAlignmentHorizontal As GRE_HORIZONTALALIGNMENT
Private mp_yPictureAlignmentVertical As GRE_VERTICALALIGNMENT
Private mp_yPlacement As E_PLACEMENT



'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// clsItemBase
    mp_sKey = ""
    mp_lIndex = 0
    '// Object Member Variables
    Set mp_oTaskStyle = New clsTaskStyle
    mp_oTaskStyle.Initialize mp_oControl
    
    Set mp_oMilestoneStyle = New clsMilestoneStyle
    mp_oMilestoneStyle.Initialize mp_oControl
    
    Set mp_oPredecessorStyle = New clsPredecessorStyle
    mp_oPredecessorStyle.Initialize mp_oControl
    
    
    
    mp_bCaptionVisible = True
    mp_bClipCaption = True
    mp_bDrawCaptionInVisibleArea = False
    mp_bSelectionRectangleVisible = True
    mp_bUseMask = True
    mp_clrBackColor = GRE_COLORS.CLR_BUTTONFACE
    mp_clrForeColor = GRE_COLORS.CLR_BLACK
    mp_clrBorderColor = GRE_COLORS.CLR_BLACK
    mp_clrEndGradientColor = GRE_COLORS.CLR_BLACK
    mp_clrStartGradientColor = GRE_COLORS.CLR_BLACK
    Set mp_Font = New StdFont
    mp_iHatchFactor = 5
    mp_lCaptionFlags = 0
    mp_lCaptionOffsetBottom = 0
    mp_lCaptionOffsetLeft = 0
    mp_lCaptionOffsetRight = 0
    mp_lCaptionOffsetTop = 0
    mp_lCaptionXMargin = 0
    mp_lCaptionYMargin = 0

    mp_lOffsetBottom = 10
    mp_lOffsetTop = 10
    mp_lPictureXMargin = 3
    mp_lPictureYMargin = 3
    mp_lSelectionRectangleOffsetBottom = 3
    mp_lSelectionRectangleOffsetLeft = 3
    mp_lSelectionRectangleOffsetRight = 3
    mp_lSelectionRectangleOffsetTop = 3

    mp_sTag = ""
    mp_yAppearance = E_STYLEAPPEARANCE.SA_RAISED

    mp_yBackgroundPattern = GRE_BACKGROUNDPATTERN.FP_SOLID
    mp_yBorderStyle = E_STYLEBORDER.SBR_NONE
    mp_yButtonStyle = GRE_BUTTONSTYLE.BT_NORMALWINDOWS
    mp_yCaptionAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_CENTER
    mp_yCaptionAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_CENTER
    mp_yCaptionPlacement = E_CAPTIONPLACEMENT.SCP_OBJECTEXTENTSPLACEMENT

    mp_yGradientFillMode = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL

    mp_yPictureAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_CENTER
    mp_yPictureAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_CENTER
    mp_yPlacement = E_PLACEMENT.PLC_ROWEXTENTSPLACEMENT
    

End Sub

Private Sub Class_Terminate()
    Set mp_Font = Nothing
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// clsItemBase
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Key() As String
    Key = mp_sKey
End Property

Public Property Let Key(ByVal Value As String)
    mp_oControl.Styles.oCollection.mp_SetKey mp_sKey, Value, 50000
End Property

Public Property Get Index() As Long
    Index = mp_lIndex
End Property

Public Property Let Index(ByVal Value As Long)
    mp_lIndex = Value
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Object Member Variables
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get TaskStyle() As clsTaskStyle
    Set TaskStyle = mp_oTaskStyle
End Property

Public Property Get MilestoneStyle() As clsMilestoneStyle
    Set MilestoneStyle = mp_oMilestoneStyle
End Property

Public Property Get PredecessorStyle() As clsPredecessorStyle
    Set PredecessorStyle = mp_oPredecessorStyle
End Property


Public Property Get CaptionVisible() As Boolean
    CaptionVisible = mp_bCaptionVisible
End Property

Public Property Let CaptionVisible(ByVal Value As Boolean)
    mp_bCaptionVisible = Value
End Property

Public Property Get ClipCaption() As Boolean
    ClipCaption = mp_bClipCaption
End Property

Public Property Let ClipCaption(ByVal Value As Boolean)
    mp_bClipCaption = Value
End Property

Public Property Get DrawCaptionInVisibleArea() As Boolean
    DrawCaptionInVisibleArea = mp_bDrawCaptionInVisibleArea
End Property

Public Property Let DrawCaptionInVisibleArea(ByVal Value As Boolean)
    mp_bDrawCaptionInVisibleArea = Value
End Property

Public Property Get SelectionRectangleVisible() As Boolean
    SelectionRectangleVisible = mp_bSelectionRectangleVisible
End Property

Public Property Let SelectionRectangleVisible(ByVal Value As Boolean)
    mp_bSelectionRectangleVisible = Value
End Property

Public Property Get UseMask() As Boolean
    UseMask = mp_bUseMask
End Property

Public Property Let UseMask(ByVal Value As Boolean)
    mp_bUseMask = Value
End Property
    
Public Property Get BackColor() As OLE_COLOR
    BackColor = mp_clrBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
    mp_clrBackColor = mp_oControl.GrphLib.ConvertColor(Value)
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = mp_clrBorderColor
End Property

Public Property Let BorderColor(ByVal Value As OLE_COLOR)
    mp_clrBorderColor = mp_oControl.GrphLib.ConvertColor(Value)
End Property

Public Property Get ButtonStyle() As GRE_BUTTONSTYLE
    ButtonStyle = mp_yButtonStyle
End Property

Public Property Let ButtonStyle(ByVal Value As GRE_BUTTONSTYLE)
    mp_yButtonStyle = Value
End Property
    
Public Property Get EndGradientColor() As OLE_COLOR
    EndGradientColor = mp_clrEndGradientColor
End Property

Public Property Let EndGradientColor(ByVal Value As OLE_COLOR)
    mp_clrEndGradientColor = mp_oControl.GrphLib.ConvertColor(Value)
End Property
    
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mp_clrForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
    mp_clrForeColor = mp_oControl.GrphLib.ConvertColor(Value)
End Property

Public Property Get StartGradientColor() As OLE_COLOR
    StartGradientColor = mp_clrStartGradientColor
End Property

Public Property Let StartGradientColor(ByVal Value As OLE_COLOR)
    mp_clrStartGradientColor = mp_oControl.GrphLib.ConvertColor(Value)
End Property
    
Public Property Get Font() As Font
    Set Font = mp_Font
End Property

Public Property Set Font(ByVal Value As Font)
    Set mp_Font = Value
End Property

Public Property Get HatchFactor() As Integer
    HatchFactor = mp_iHatchFactor
End Property

Public Property Let HatchFactor(ByVal Value As Integer)
    mp_iHatchFactor = Value
End Property



Public Property Get CaptionFlags() As Long
    CaptionFlags = mp_lCaptionFlags
End Property

Public Property Let CaptionFlags(ByVal Value As Long)
    mp_lCaptionFlags = Value
End Property

Public Property Get CaptionOffsetBottom() As Long
    CaptionOffsetBottom = mp_lCaptionOffsetBottom
End Property

Public Property Let CaptionOffsetBottom(ByVal Value As Long)
    mp_lCaptionOffsetBottom = Value
End Property

Public Property Get CaptionOffsetLeft() As Long
    CaptionOffsetLeft = mp_lCaptionOffsetLeft
End Property

Public Property Let CaptionOffsetLeft(ByVal Value As Long)
    mp_lCaptionOffsetLeft = Value
End Property

Public Property Get CaptionOffsetRight() As Long
    CaptionOffsetRight = mp_lCaptionOffsetRight
End Property

Public Property Let CaptionOffsetRight(ByVal Value As Long)
    mp_lCaptionOffsetRight = Value
End Property

Public Property Get CaptionOffsetTop() As Long
    CaptionOffsetTop = mp_lCaptionOffsetTop
End Property

Public Property Let CaptionOffsetTop(ByVal Value As Long)
    mp_lCaptionOffsetTop = Value
End Property

Public Property Get CaptionXMargin() As Long
    CaptionXMargin = mp_lCaptionXMargin
End Property

Public Property Let CaptionXMargin(ByVal Value As Long)
    mp_lCaptionXMargin = Value
End Property

Public Property Get CaptionYMargin() As Long
    CaptionYMargin = mp_lCaptionYMargin
End Property

Public Property Let CaptionYMargin(ByVal Value As Long)
    mp_lCaptionYMargin = Value
End Property



Public Property Get OffsetBottom() As Long
    OffsetBottom = mp_lOffsetBottom
End Property

Public Property Let OffsetBottom(ByVal Value As Long)
    mp_lOffsetBottom = Value
End Property

Public Property Get OffsetTop() As Long
    OffsetTop = mp_lOffsetTop
End Property

Public Property Let OffsetTop(ByVal Value As Long)
    mp_lOffsetTop = Value
End Property

Public Property Get PictureXMargin() As Long
    PictureXMargin = mp_lPictureXMargin
End Property

Public Property Let PictureXMargin(ByVal Value As Long)
    mp_lPictureXMargin = Value
End Property

Public Property Get PictureYMargin() As Long
    PictureYMargin = mp_lPictureYMargin
End Property

Public Property Let PictureYMargin(ByVal Value As Long)
    mp_lPictureYMargin = Value
End Property

Public Property Get SelectionRectangleOffsetBottom() As Long
    SelectionRectangleOffsetBottom = mp_lSelectionRectangleOffsetBottom
End Property

Public Property Let SelectionRectangleOffsetBottom(ByVal Value As Long)
    mp_lSelectionRectangleOffsetBottom = Value
End Property

Public Property Get SelectionRectangleOffsetLeft() As Long
    SelectionRectangleOffsetLeft = mp_lSelectionRectangleOffsetLeft
End Property

Public Property Let SelectionRectangleOffsetLeft(ByVal Value As Long)
    mp_lSelectionRectangleOffsetLeft = Value
End Property

Public Property Get SelectionRectangleOffsetRight() As Long
    SelectionRectangleOffsetRight = mp_lSelectionRectangleOffsetRight
End Property

Public Property Let SelectionRectangleOffsetRight(ByVal Value As Long)
    mp_lSelectionRectangleOffsetRight = Value
End Property

Public Property Get SelectionRectangleOffsetTop() As Long
    SelectionRectangleOffsetTop = mp_lSelectionRectangleOffsetTop
End Property

Public Property Let SelectionRectangleOffsetTop(ByVal Value As Long)
    mp_lSelectionRectangleOffsetTop = Value
End Property

Public Property Get Tag() As String
    Tag = mp_sTag
End Property

Public Property Let Tag(ByVal Value As String)
    mp_sTag = Value
End Property
    
Public Property Get Appearance() As E_STYLEAPPEARANCE
    Appearance = mp_yAppearance
End Property

Public Property Let Appearance(ByVal Value As E_STYLEAPPEARANCE)
    mp_yAppearance = Value
End Property

    
Public Property Get BackgroundPattern() As GRE_BACKGROUNDPATTERN
    BackgroundPattern = mp_yBackgroundPattern
End Property

Public Property Let BackgroundPattern(ByVal Value As GRE_BACKGROUNDPATTERN)
    mp_yBackgroundPattern = Value
End Property

Public Property Get BorderStyle() As E_STYLEBORDER
    BorderStyle = mp_yBorderStyle
End Property

Public Property Let BorderStyle(ByVal Value As E_STYLEBORDER)
    mp_yBorderStyle = Value
End Property

Public Property Get CaptionAlignmentHorizontal() As GRE_HORIZONTALALIGNMENT
    CaptionAlignmentHorizontal = mp_yCaptionAlignmentHorizontal
End Property

Public Property Let CaptionAlignmentHorizontal(ByVal Value As GRE_HORIZONTALALIGNMENT)
    mp_yCaptionAlignmentHorizontal = Value
End Property

Public Property Get CaptionAlignmentVertical() As GRE_VERTICALALIGNMENT
    CaptionAlignmentVertical = mp_yCaptionAlignmentVertical
End Property

Public Property Let CaptionAlignmentVertical(ByVal Value As GRE_VERTICALALIGNMENT)
    mp_yCaptionAlignmentVertical = Value
End Property

Public Property Get CaptionPlacement() As E_CAPTIONPLACEMENT
    CaptionPlacement = mp_yCaptionPlacement
End Property

Public Property Let CaptionPlacement(ByVal Value As E_CAPTIONPLACEMENT)
    mp_yCaptionPlacement = Value
End Property
    
Public Property Get GradientFillMode() As GRE_GRADIENTFILLMODE
    GradientFillMode = mp_yGradientFillMode
End Property

Public Property Let GradientFillMode(ByVal Value As GRE_GRADIENTFILLMODE)
    mp_yGradientFillMode = Value
End Property
    

    
Public Property Get PictureAlignmentHorizontal() As GRE_HORIZONTALALIGNMENT
    PictureAlignmentHorizontal = mp_yPictureAlignmentHorizontal
End Property

Public Property Let PictureAlignmentHorizontal(ByVal Value As GRE_HORIZONTALALIGNMENT)
    mp_yPictureAlignmentHorizontal = Value
End Property

Public Property Get PictureAlignmentVertical() As GRE_VERTICALALIGNMENT
    PictureAlignmentVertical = mp_yPictureAlignmentVertical
End Property

Public Property Let PictureAlignmentVertical(ByVal Value As GRE_VERTICALALIGNMENT)
    mp_yPictureAlignmentVertical = Value
End Property

Public Property Get Placement() As E_PLACEMENT
    Placement = mp_yPlacement
End Property

Public Property Let Placement(ByVal Value As E_PLACEMENT)
    mp_yPlacement = Value
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Style"
    oXML.InitializeWriter
    oXML.WriteProperty "Key", mp_sKey
    oXML.WriteProperty "CaptionVisible", mp_bCaptionVisible
    oXML.WriteProperty "ClipCaption", mp_bClipCaption
    oXML.WriteProperty "DrawCaptionInVisibleArea", mp_bDrawCaptionInVisibleArea
    oXML.WriteProperty "SelectionRectangleVisible", mp_bSelectionRectangleVisible
    oXML.WriteProperty "UseMask", mp_bUseMask
    oXML.WriteProperty "BackColor", mp_clrBackColor
    oXML.WriteProperty "BorderColor", mp_clrBorderColor

    oXML.WriteProperty "EndGradientColor", mp_clrEndGradientColor
    oXML.WriteProperty "ForeColor", mp_clrForeColor

    oXML.WriteProperty "StartGradientColor", mp_clrStartGradientColor
    oXML.WritePropertyFont "Font", mp_Font
    oXML.WriteProperty "HatchFactor", mp_iHatchFactor
    
    oXML.WriteProperty "CaptionFlags", mp_lCaptionFlags
    oXML.WriteProperty "CaptionOffsetBottom", mp_lCaptionOffsetBottom
    oXML.WriteProperty "CaptionOffsetLeft", mp_lCaptionOffsetLeft
    oXML.WriteProperty "CaptionOffsetRight", mp_lCaptionOffsetRight
    oXML.WriteProperty "CaptionOffsetTop", mp_lCaptionOffsetTop
    oXML.WriteProperty "CaptionXMargin", mp_lCaptionXMargin
    oXML.WriteProperty "CaptionYMargin", mp_lCaptionYMargin

    oXML.WriteProperty "OffsetBottom", mp_lOffsetBottom
    oXML.WriteProperty "OffsetTop", mp_lOffsetTop
    oXML.WriteProperty "PictureXMargin", mp_lPictureXMargin
    oXML.WriteProperty "PictureYMargin", mp_lPictureYMargin
    oXML.WriteProperty "SelectionRectangleOffsetBottom", mp_lSelectionRectangleOffsetBottom
    oXML.WriteProperty "SelectionRectangleOffsetLeft", mp_lSelectionRectangleOffsetLeft
    oXML.WriteProperty "SelectionRectangleOffsetRight", mp_lSelectionRectangleOffsetRight
    oXML.WriteProperty "SelectionRectangleOffsetTop", mp_lSelectionRectangleOffsetTop

    oXML.WriteProperty "Tag", mp_sTag
    oXML.WriteProperty "Appearance", mp_yAppearance

    oXML.WriteProperty "BackgroundPattern", mp_yBackgroundPattern
    oXML.WriteProperty "BorderStyle", mp_yBorderStyle
    oXML.WriteProperty "ButtonStyle", mp_yButtonStyle
    oXML.WriteProperty "CaptionAlignmentHorizontal", mp_yCaptionAlignmentHorizontal
    oXML.WriteProperty "CaptionAlignmentVertical", mp_yCaptionAlignmentVertical
    oXML.WriteProperty "CaptionPlacement", mp_yCaptionPlacement

    oXML.WriteProperty "GradientFillMode", mp_yGradientFillMode

    oXML.WriteProperty "PictureAlignmentHorizontal", mp_yPictureAlignmentHorizontal
    oXML.WriteProperty "PictureAlignmentVertical", mp_yPictureAlignmentVertical
    oXML.WriteProperty "Placement", mp_yPlacement
    
    oXML.WriteObject mp_oTaskStyle.GetXML
    oXML.WriteObject mp_oMilestoneStyle.GetXML
    oXML.WriteObject mp_oPredecessorStyle.GetXML

    
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Style"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadProperty "Key", mp_sKey
    oXML.ReadProperty "CaptionVisible", mp_bCaptionVisible
    oXML.ReadProperty "ClipCaption", mp_bClipCaption
    oXML.ReadProperty "DrawCaptionInVisibleArea", mp_bDrawCaptionInVisibleArea
    oXML.ReadProperty "SelectionRectangleVisible", mp_bSelectionRectangleVisible
    oXML.ReadProperty "UseMask", mp_bUseMask
    oXML.ReadPropertyColor "BackColor", mp_clrBackColor
    oXML.ReadPropertyColor "BorderColor", mp_clrBorderColor

    oXML.ReadPropertyColor "EndGradientColor", mp_clrEndGradientColor
    oXML.ReadPropertyColor "ForeColor", mp_clrForeColor

    oXML.ReadPropertyColor "StartGradientColor", mp_clrStartGradientColor
    oXML.ReadPropertyFont "Font", mp_Font
    oXML.ReadProperty "HatchFactor", mp_iHatchFactor
    oXML.ReadProperty "CaptionFlags", mp_lCaptionFlags
    oXML.ReadProperty "CaptionOffsetBottom", mp_lCaptionOffsetBottom
    oXML.ReadProperty "CaptionOffsetLeft", mp_lCaptionOffsetLeft
    oXML.ReadProperty "CaptionOffsetRight", mp_lCaptionOffsetRight
    oXML.ReadProperty "CaptionOffsetTop", mp_lCaptionOffsetTop
    oXML.ReadProperty "CaptionXMargin", mp_lCaptionXMargin
    oXML.ReadProperty "CaptionYMargin", mp_lCaptionYMargin

    oXML.ReadProperty "OffsetBottom", mp_lOffsetBottom
    oXML.ReadProperty "OffsetTop", mp_lOffsetTop
    oXML.ReadProperty "PictureXMargin", mp_lPictureXMargin
    oXML.ReadProperty "PictureYMargin", mp_lPictureYMargin
    oXML.ReadProperty "SelectionRectangleOffsetBottom", mp_lSelectionRectangleOffsetBottom
    oXML.ReadProperty "SelectionRectangleOffsetLeft", mp_lSelectionRectangleOffsetLeft
    oXML.ReadProperty "SelectionRectangleOffsetRight", mp_lSelectionRectangleOffsetRight
    oXML.ReadProperty "SelectionRectangleOffsetTop", mp_lSelectionRectangleOffsetTop

    oXML.ReadProperty "Tag", mp_sTag
    oXML.ReadProperty "Appearance", mp_yAppearance

    oXML.ReadProperty "BackgroundPattern", mp_yBackgroundPattern
    oXML.ReadProperty "BorderStyle", mp_yBorderStyle
    oXML.ReadProperty "ButtonStyle", mp_yButtonStyle
    oXML.ReadProperty "CaptionAlignmentHorizontal", mp_yCaptionAlignmentHorizontal
    oXML.ReadProperty "CaptionAlignmentVertical", mp_yCaptionAlignmentVertical
    oXML.ReadProperty "CaptionPlacement", mp_yCaptionPlacement

    oXML.ReadProperty "GradientFillMode", mp_yGradientFillMode

    oXML.ReadProperty "PictureAlignmentHorizontal", mp_yPictureAlignmentHorizontal
    oXML.ReadProperty "PictureAlignmentVertical", mp_yPictureAlignmentVertical
    oXML.ReadProperty "Placement", mp_yPlacement
    
    mp_oTaskStyle.SetXML oXML.ReadObject("TaskStyle")
    mp_oMilestoneStyle.SetXML oXML.ReadObject("MilestoneStyle")
    mp_oPredecessorStyle.SetXML oXML.ReadObject("PredecessorStyle")
    


End Sub

