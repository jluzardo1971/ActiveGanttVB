VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHScrollBarEx"
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
Private WithEvents mp_oScrollBar As HScrollBar
Attribute mp_oScrollBar.VB_VarHelpID = -1
Private mp_bValueSet As Boolean
Private mp_yState As Integer
Private mp_lValueBuff As Integer

Private Enum E_SCROLLSTATE
    SS_CANTDISPLAY = 0
    SS_NOTNEEDED = 1
    SS_NEEDED = 2
    SS_SHOWN = 3
    SS_HIDDEN = 4
End Enum

Public Event ValueChanged(ByVal Offset As Integer)

Public Sub Initialize(ByRef Value As ActiveGanttVBCtl, ByRef Value2 As Object)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oScrollBar = Value2
    mp_lValueBuff = mp_oScrollBar.Value
End Sub

Private Sub mp_oScrollBar_Change()
    If mp_bValueSet = True Then
        Exit Sub
    End If
    RaiseEvent ValueChanged(mp_oScrollBar.Value - mp_lValueBuff)
End Sub

Public Property Get Value() As Integer
    Value = mp_oScrollBar.Value
End Property

Public Property Let Value(ByVal Value As Integer)
    mp_bValueSet = True
    mp_oScrollBar.Value = Value
    mp_lValueBuff = Value
    mp_bValueSet = False
End Property

Public Property Get Min() As Integer
    Min = mp_oScrollBar.Min
End Property

Public Property Let Min(ByVal Value As Integer)
    mp_oScrollBar.Min = Value
End Property

Public Property Get Max() As Integer
    Max = mp_oScrollBar.Max
End Property

Public Property Let Max(ByVal Value As Integer)
    If Value < mp_oScrollBar.Min Then
        Value = mp_oScrollBar.Min
    End If
    mp_oScrollBar.Max = Value
End Property

Public Property Get SmallChange() As Integer
    SmallChange = mp_oScrollBar.SmallChange
End Property

Public Property Let SmallChange(ByVal Value As Integer)
    mp_oScrollBar.SmallChange = Value
End Property

Public Property Get LargeChange() As Integer
    LargeChange = mp_oScrollBar.LargeChange
End Property

Public Property Let LargeChange(ByVal Value As Integer)
    mp_oScrollBar.LargeChange = Value
End Property

Public Property Get State() As Integer
    State = mp_yState
End Property

Public Property Let State(ByVal Value As Integer)
    mp_yState = Value
    Select Case mp_yState
        Case E_SCROLLSTATE.SS_CANTDISPLAY
            mp_yState = E_SCROLLSTATE.SS_HIDDEN
            mp_oScrollBar.Visible = False
        Case E_SCROLLSTATE.SS_NOTNEEDED
            If mp_oControl.ScrollBarBehaviour = E_SCROLLBEHAVIOUR.SB_DISABLE Then
                mp_yState = E_SCROLLSTATE.SS_SHOWN
                mp_oScrollBar.Enabled = False
                mp_oScrollBar.Visible = True
            Else
                mp_yState = E_SCROLLSTATE.SS_HIDDEN
                mp_oScrollBar.Visible = False
            End If
        Case E_SCROLLSTATE.SS_NEEDED
            mp_yState = E_SCROLLSTATE.SS_SHOWN
            mp_oScrollBar.Enabled = True
            mp_oScrollBar.Visible = True
    End Select
End Property

Public Property Get Left() As Long
    Left = mp_oScrollBar.Left
End Property

Public Property Let Left(ByVal Value As Long)
    mp_oScrollBar.Left = Value
End Property

Public Property Get Top() As Long
    Top = mp_oScrollBar.Top
End Property

Public Property Let Top(ByVal Value As Long)
    mp_oScrollBar.Top = Value
End Property

Public Property Get Width() As Long
    Width = mp_oScrollBar.Width
End Property

Public Property Let Width(ByVal Value As Long)
    mp_oScrollBar.Width = Value
End Property

Public Property Get Height() As Long
    Height = mp_oScrollBar.Height
End Property

Public Property Let Height(ByVal Value As Long)
    mp_oScrollBar.Height = Value
End Property

Public Property Get Enabled() As Boolean
    Enabled = mp_oScrollBar.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
    mp_oScrollBar.Enabled = Value
End Property


