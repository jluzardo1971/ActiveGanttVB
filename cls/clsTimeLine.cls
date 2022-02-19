VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimeLine"
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
'// Object Member Variables
Private mp_oView As clsView
Private mp_oScrollBar As clsTimeLineScrollBar
Private mp_oTierArea As clsTierArea
Private mp_oTickMarkArea As clsTickMarkArea
Private mp_sStyleIndex As String
Private mp_yAppearance As E_BORDERSTYLE
Private mp_clrBackColor As OLE_COLOR
Private mp_clrForeColor As OLE_COLOR
Private mp_dtEndDate As Date
Private mp_oProgressLine As clsProgressLine

Private mp_dtStartDate As Date

Private mp_lEnd As Long
Private mp_lStart As Long

'// ---------------------------------------------------------------------------------------------------------------------
'// Initialization & Destruction
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl, ByRef oView As clsView)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oView = oView
    Set mp_oScrollBar = New clsTimeLineScrollBar
    mp_oScrollBar.Initialize mp_oControl
    Set mp_oTierArea = New clsTierArea
    mp_oTierArea.Initialize mp_oControl, Me
    Set mp_oTickMarkArea = New clsTickMarkArea
    mp_oTickMarkArea.Initialize mp_oControl, Me, True
    
    Set mp_oProgressLine = New clsProgressLine
    mp_oProgressLine.Initialize mp_oControl, Me
    mp_sStyleIndex = "0"
    mp_yAppearance = E_BORDERSTYLE.TLB_3D
    mp_clrBackColor = GRE_COLORS.CLR_BUTTONFACE
    mp_clrForeColor = GRE_COLORS.CLR_BLACK
    mp_dtStartDate = mp_oControl.MathLib.DateTimeAdd("h", -3, mp_oControl.MathLib.CurrentDateTime)
End Sub

Private Sub Class_Terminate()
    Set mp_oView = Nothing
    Set mp_oTierArea = Nothing
    Set mp_oTickMarkArea = Nothing
    Set mp_oScrollBar = Nothing
    Set mp_oProgressLine = Nothing
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Object Member Variables
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get ScrollBar() As clsTimeLineScrollBar
    Set ScrollBar = mp_oScrollBar
End Property

Public Property Get ProgressLine() As clsProgressLine
    Set ProgressLine = mp_oProgressLine
End Property

Public Property Get TierArea() As clsTierArea
    Set TierArea = mp_oTierArea
End Property

Public Property Get TickMarkArea() As clsTickMarkArea
    Set TickMarkArea = mp_oTickMarkArea
End Property

Public Property Get StyleIndex() As String
    StyleIndex = mp_sStyleIndex
End Property

Public Property Let StyleIndex(ByVal Value As String)
    If Value = "" Then
        Value = "0"
    End If
    mp_sStyleIndex = Value
End Property

Public Property Get Appearance() As E_BORDERSTYLE
    Appearance = mp_yAppearance
End Property

Public Property Let Appearance(ByVal Value As E_BORDERSTYLE)
    mp_yAppearance = Value
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mp_clrBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
    mp_clrBackColor = mp_oControl.GrphLib.ConvertColor(Value)
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mp_clrForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
    mp_clrForeColor = mp_oControl.GrphLib.ConvertColor(Value)
End Property

Public Property Get EndDate() As Date
    EndDate = mp_dtEndDate
End Property

Friend Property Let f_EndDate(ByVal Value As Date)
    mp_dtEndDate = Value
End Property

Public Property Get StartDate() As Date
    StartDate = mp_dtStartDate
End Property

Friend Property Let f_StartDate(ByVal Value As Date)
    mp_dtStartDate = Value
End Property

Friend Property Get f_lStart() As Long
    f_lStart = mp_lStart
End Property

Friend Property Let f_lStart(ByVal Value As Long)
    mp_lStart = Value
End Property

Friend Property Get f_lEnd() As Long
    f_lEnd = mp_lEnd
End Property

Friend Property Let f_lEnd(ByVal Value As Long)
    mp_lEnd = Value
End Property

Public Sub Move(ByVal Interval As String)
    Dim sInterval As String
    Dim lFactor As Long
    Dim lSign As Long
    If mp_oControl.StrLib.StrLeft(Interval, 1) = "-" Then
        lSign = -1
        Interval = mp_oControl.StrLib.StrRight(Interval, mp_oControl.StrLib.StrLen(Interval) - 1)
    Else
        lSign = 1
    End If
    mp_oControl.mp_ProcessInterval Interval, sInterval, lFactor
    mp_dtStartDate = mp_oControl.MathLib.DateTimeAdd(sInterval, lFactor * lSign, mp_dtStartDate)
End Sub

Public Sub Position(ByVal TimeLineStartDate As Date)
    mp_dtStartDate = TimeLineStartDate
End Sub

Public Property Get Height() As Long
    Height = Bottom - Top
End Property

Public Property Get Top() As Long
    Top = mp_oControl.mt_BorderThickness
End Property

Public Property Get Bottom() As Long
    Dim lReturn As Long
    Dim lUpperTierHeight As Long
    Dim lMiddleTierHeight As Long
    Dim lLowerTierHeight As Long
    Dim lTickMarkAreaHeight As Long
    lReturn = 0
    lUpperTierHeight = 0
    lLowerTierHeight = 0
    lTickMarkAreaHeight = 0
    If (mp_oTierArea.UpperTier.Visible = True) Then
        lUpperTierHeight = mp_oTierArea.UpperTier.Height
    End If
    If (mp_oTierArea.MiddleTier.Visible = True) Then
        lMiddleTierHeight = mp_oTierArea.MiddleTier.Height
    End If
    If (mp_oTierArea.LowerTier.Visible = True) Then
        lLowerTierHeight = mp_oTierArea.LowerTier.Height
    End If
    If (mp_oTickMarkArea.Visible = True) Then
        lTickMarkAreaHeight = mp_oTickMarkArea.Height
    End If
    lReturn = lUpperTierHeight + lMiddleTierHeight + lLowerTierHeight + lTickMarkAreaHeight
    If (lReturn > 0) Then
        lReturn = lReturn + 3
    End If
    lReturn = lReturn + mp_oControl.mt_BorderThickness
    Bottom = lReturn
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Position
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Calculate()
    If mp_oScrollBar.Enabled = True Then
        mp_dtStartDate = mp_oControl.MathLib.DateTimeAdd(mp_oScrollBar.f_Interval, mp_oScrollBar.Value * mp_oScrollBar.f_Factor, mp_oScrollBar.StartDate)
    End If
    Select Case mp_yAppearance
        Case E_BORDERSTYLE.TLB_NONE
            mp_lStart = mp_oControl.Splitter.Right()
            mp_lEnd = mp_oControl.mt_RightMargin()
        Case E_BORDERSTYLE.TLB_SINGLE
            mp_lStart = mp_oControl.Splitter.Right() + 1
            mp_lEnd = mp_oControl.mt_RightMargin() - 1
        Case E_BORDERSTYLE.TLB_3D
            mp_lStart = mp_oControl.Splitter.Right() + 2
            mp_lEnd = mp_oControl.mt_RightMargin() - 2
    End Select
    mp_dtEndDate = mp_oControl.MathLib.DateTimeAdd(mp_oView.f_Interval, (mp_lEnd - mp_lStart) * mp_oView.f_Factor, mp_dtStartDate)
End Sub

Friend Function TiersTickMarksPosition(ByVal v_yTier As String) As Long
    Dim lReturn As Long
    Dim lUpperTierHeight As Long
    Dim lMiddleTierHeight As Long
    Dim lLowerTierHeight As Long
    Dim lTickMarkAreaHeight As Long
    lReturn = 0
    lUpperTierHeight = 0
    lLowerTierHeight = 0
    lTickMarkAreaHeight = 0
    If (mp_oTierArea.UpperTier.Visible = True) Then
        lUpperTierHeight = mp_oTierArea.UpperTier.Height
    End If
    If (mp_oTierArea.MiddleTier.Visible = True) Then
        lMiddleTierHeight = mp_oTierArea.MiddleTier.Height
    End If
    If (mp_oTierArea.LowerTier.Visible = True) Then
        lLowerTierHeight = mp_oTierArea.LowerTier.Height
    End If
    If (mp_oTickMarkArea.Visible = True) Then
        lTickMarkAreaHeight = mp_oTickMarkArea.Height
    End If
    lReturn = lUpperTierHeight + lMiddleTierHeight + lLowerTierHeight + lTickMarkAreaHeight
    If (lReturn > 0) Then
        lReturn = lReturn + 3
    End If
    lReturn = lReturn + mp_oControl.mt_BorderThickness
    Select Case (v_yTier)
        Case "UpperTier"
            lReturn = lReturn - lUpperTierHeight - lMiddleTierHeight - lLowerTierHeight - lTickMarkAreaHeight
        Case "MiddleTier"
            lReturn = lReturn - lMiddleTierHeight - lLowerTierHeight - lTickMarkAreaHeight
        Case "LowerTier"
            lReturn = lReturn - lLowerTierHeight - lTickMarkAreaHeight
        Case "TickMarkArea"
            lReturn = lReturn - lTickMarkAreaHeight
        Case Else
            MsgBox "TiersTickMarksPosition Error"
    End Select
    TiersTickMarksPosition = lReturn
End Function

'// ---------------------------------------------------------------------------------------------------------------------
'// Draw
'// ---------------------------------------------------------------------------------------------------------------------
Friend Sub Draw()
    Dim lBottom As Long
    Dim lTop As Long
    Dim lLeft As Long
    Dim lRight As Long
    If (Height = 0) Then
        Exit Sub
    End If
    lBottom = Bottom
    lTop = Top
    lLeft = mp_oControl.Splitter.Right()
    lRight = mp_oControl.mt_RightMargin()
    mp_oControl.GrphLib.ClipRegion lLeft, lTop, lRight, lBottom, True
    If (mp_sStyleIndex = "0") Then
        Select Case mp_yAppearance
            Case E_BORDERSTYLE.TLB_NONE
                mp_oControl.GrphLib.DrawLine lLeft, lTop, lRight, lBottom, GRE_LINETYPE.LT_FILLED, mp_clrBackColor, GRE_LINEDRAWSTYLE.LDS_SOLID
            Case E_BORDERSTYLE.TLB_SINGLE
                mp_oControl.GrphLib.DrawLine lLeft, lTop, lRight, lBottom, GRE_LINETYPE.LT_FILLED, mp_clrBackColor, GRE_LINEDRAWSTYLE.LDS_SOLID
                mp_oControl.GrphLib.DrawLine lLeft, lTop, lRight, lBottom, GRE_LINETYPE.LT_BORDER, GRE_COLORS.CLR_BLACK, GRE_LINEDRAWSTYLE.LDS_SOLID
                mp_oControl.GrphLib.ClipRegion lLeft + 1, lTop + 1, lRight - 1, lBottom - 1, True
            Case E_BORDERSTYLE.TLB_3D
                mp_oControl.GrphLib.DrawEdge lLeft, lTop, lRight, lBottom, mp_clrBackColor, GRE_BUTTONSTYLE.BT_NORMALWINDOWS, GRE_EDGETYPE.ET_RAISED, True
                mp_oControl.GrphLib.ClipRegion lLeft + 2, lTop + 2, lRight - 2, lBottom - 2, True
        End Select
    Else
        mp_oControl.mp_DrawItem lLeft, lRight, lTop, lBottom, mp_sStyleIndex, "", False, Nothing, 0, 0, Nothing
        Select Case mp_yAppearance
            Case E_BORDERSTYLE.TLB_NONE
            Case E_BORDERSTYLE.TLB_SINGLE
                mp_oControl.GrphLib.ClipRegion lLeft + 1, lTop + 1, lRight - 1, lBottom - 1, True
            Case E_BORDERSTYLE.TLB_3D
                mp_oControl.GrphLib.ClipRegion lLeft + 2, lTop + 2, lRight - 2, lBottom - 2, True
        End Select
    End If
    mp_oTierArea.UpperTier.Position
    mp_oTierArea.MiddleTier.Position
    mp_oTierArea.LowerTier.Position
    mp_oTickMarkArea.Draw
End Sub



'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TimeLine"
    oXML.InitializeWriter
    oXML.WriteProperty "StyleIndex", mp_sStyleIndex
    oXML.WriteProperty "Appearance", mp_yAppearance
    oXML.WriteProperty "BackColor", mp_clrBackColor
    oXML.WriteProperty "ForeColor", mp_clrForeColor
    oXML.WriteProperty "EndDate", mp_dtEndDate
    oXML.WriteProperty "StartDate", mp_dtStartDate
    oXML.WriteObject mp_oProgressLine.GetXML
    oXML.WriteObject mp_oScrollBar.GetXML
    oXML.WriteObject mp_oTierArea.GetXML
    oXML.WriteObject mp_oTickMarkArea.GetXML
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TimeLine"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadProperty "StyleIndex", mp_sStyleIndex
    oXML.ReadProperty "Appearance", mp_yAppearance
    oXML.ReadPropertyColor "BackColor", mp_clrBackColor
    oXML.ReadPropertyColor "ForeColor", mp_clrForeColor
    oXML.ReadProperty "EndDate", mp_dtEndDate
    oXML.ReadProperty "StartDate", mp_dtStartDate
    mp_oProgressLine.SetXML oXML.ReadObject("ProgressLine")
    mp_oScrollBar.SetXML oXML.ReadObject("ScrollBar")
    mp_oTierArea.SetXML oXML.ReadObject("TierArea")
    mp_oTickMarkArea.SetXML oXML.ReadObject("TickMarkArea")
End Sub

