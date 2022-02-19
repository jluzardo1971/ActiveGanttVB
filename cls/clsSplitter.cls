VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSplitter"
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
Private mp_lPosition As Long
Private mp_yAppearance As E_BORDERSTYLE

'// ---------------------------------------------------------------------------------------------------------------------
'// Initialization & Destruction
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    mp_lPosition = 125
    mp_yAppearance = E_BORDERSTYLE.TLB_3D
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Object Member Variables
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Appearance() As E_BORDERSTYLE
    Appearance = mp_yAppearance
End Property

Public Property Let Appearance(ByVal Value As E_BORDERSTYLE)
    mp_yAppearance = Value
End Property

Public Property Get Position() As Long
    Position = mp_lPosition
End Property

Public Property Let Position(ByVal Value As Long)
    If (Value <= 0) Then
        Exit Property
    End If
    mp_lPosition = Value
    If (mp_lPosition > mp_oControl.Columns.Width) Then
        mp_lPosition = mp_oControl.Columns.Width
        mp_oControl.f_oHScrollBar1.Value = 0
    End If
End Property

Friend Property Get Left() As Long
    If (mp_oControl.Columns.Count <> 0) Then
        Left = mp_lPosition + mp_oControl.mt_BorderThickness - 1
    Else
        If (mp_oControl.f_UserMode = True) Then
            Left = 0
        Else
            Left = 125 + mp_oControl.mt_BorderThickness - 1
        End If
    End If
End Property

Friend Property Get Right() As Long
    If (mp_oControl.Columns.Count <> 0) Then
        Right = mp_lPosition + mp_oControl.mt_BorderThickness + 6
    Else
        If (mp_oControl.f_UserMode = True) Then
            Right = mp_oControl.mt_BorderThickness
        Else
            Right = 125 + mp_oControl.mt_BorderThickness + 6
        End If
    End If
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Draw
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Draw()
    If mp_oControl.Columns.Count = 0 And mp_oControl.f_UserMode = True Then
        Exit Sub
    End If
    mp_oControl.GrphLib.ClipRegion Left + 1, 0, Left + 6, mp_oControl.mt_BottomMargin, True
    Select Case mp_yAppearance
        Case E_BORDERSTYLE.TLB_3D
            mp_oControl.GrphLib.DrawLine Left + 1, 0, Left + 1, mp_oControl.GrphLib.Height, GRE_LINETYPE.LT_NORMAL, GRE_COLORS.CLR_BUTTONFACE, GRE_LINEDRAWSTYLE.LDS_SOLID
            mp_oControl.GrphLib.DrawLine Left + 2, 0, Left + 2, mp_oControl.GrphLib.Height, GRE_LINETYPE.LT_NORMAL, GRE_COLORS.CLR_WHITE, GRE_LINEDRAWSTYLE.LDS_SOLID
            mp_oControl.GrphLib.DrawLine Left + 3, 0, Left + 3, mp_oControl.GrphLib.Height, GRE_LINETYPE.LT_NORMAL, GRE_COLORS.CLR_BUTTONFACE, GRE_LINEDRAWSTYLE.LDS_SOLID
            mp_oControl.GrphLib.DrawLine Left + 4, 0, Left + 4, mp_oControl.GrphLib.Height, GRE_LINETYPE.LT_NORMAL, GRE_COLORS.CLR_BUTTONFACE, GRE_LINEDRAWSTYLE.LDS_SOLID
            mp_oControl.GrphLib.DrawLine Left + 5, 0, Left + 5, mp_oControl.GrphLib.Height, GRE_LINETYPE.LT_NORMAL, GRE_COLORS.CLR_VERYDARKGREY, GRE_LINEDRAWSTYLE.LDS_SOLID
            mp_oControl.GrphLib.DrawLine Left + 6, 0, Left + 6, mp_oControl.GrphLib.Height, GRE_LINETYPE.LT_NORMAL, GRE_COLORS.CLR_ALMOSTBLACK, GRE_LINEDRAWSTYLE.LDS_SOLID
        Case E_BORDERSTYLE.TLB_NONE
            mp_oControl.GrphLib.DrawLine Left + 1, 0, Left + 6, mp_oControl.GrphLib.Height, GRE_LINETYPE.LT_FILLED, GRE_COLORS.CLR_BUTTONFACE, GRE_LINEDRAWSTYLE.LDS_SOLID
        Case E_BORDERSTYLE.TLB_SINGLE
            mp_oControl.GrphLib.DrawLine Left + 1, 0, Left + 1, mp_oControl.GrphLib.Height, GRE_LINETYPE.LT_NORMAL, GRE_COLORS.CLR_BLACK, GRE_LINEDRAWSTYLE.LDS_SOLID
            mp_oControl.GrphLib.DrawLine Left + 2, 0, Left + 5, mp_oControl.GrphLib.Height, GRE_LINETYPE.LT_FILLED, GRE_COLORS.CLR_BUTTONFACE, GRE_LINEDRAWSTYLE.LDS_SOLID
            mp_oControl.GrphLib.DrawLine Left + 6, 0, Left + 6, mp_oControl.GrphLib.Height, GRE_LINETYPE.LT_NORMAL, GRE_COLORS.CLR_BLACK, GRE_LINEDRAWSTYLE.LDS_SOLID
    End Select
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Custom Methods
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub f_AdjustPosition()
    If mp_oControl.Columns.Width < (mp_oControl.GrphLib.Width - 25) Then
        If mp_lPosition < mp_oControl.Columns.Width Then
            mp_lPosition = mp_oControl.Columns.Width
        End If
    End If
    If mp_lPosition > mp_oControl.Columns.Width Then
        mp_lPosition = mp_oControl.Columns.Width
        mp_oControl.f_oHScrollBar1.Value = 0
    End If
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Splitter"
    oXML.InitializeWriter
    oXML.WriteProperty "Appearance", mp_yAppearance
    oXML.WriteProperty "Position", mp_lPosition
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Splitter"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadProperty "Appearance", mp_yAppearance
    oXML.ReadProperty "Position", mp_lPosition
End Sub

