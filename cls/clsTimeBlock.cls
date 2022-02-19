VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimeBlock"
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
Private mp_dtEndDate As Date
Private mp_dtStartDate As Date
Private mp_sStyleIndex As String
Private mp_sTag As String
'// Metrics
Private mp_lLeft As Long
Private mp_lTop As Long
Private mp_lRight As Long
Private mp_lBottom As Long
Private mp_bVisible As Boolean

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
    mp_dtEndDate = mp_oControl.MathLib.CurrentDateTime
    mp_dtStartDate = mp_oControl.MathLib.CurrentDateTime
    mp_sStyleIndex = "0"
    mp_sTag = ""
    '// Metrics
    mp_lLeft = 0
    mp_lTop = 0
    mp_lRight = 0
    mp_lBottom = 0
    mp_bVisible = False
End Sub

Private Sub Class_Terminate()
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// clsItemBase
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Key() As String
    Key = mp_sKey
End Property

Public Property Let Key(ByVal Value As String)
    mp_oControl.TimeBlocks.oCollection.mp_SetKey mp_sKey, Value, 50211
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

Public Property Get EndDate() As Date
    EndDate = mp_dtEndDate
End Property

Public Property Let EndDate(ByVal Value As Date)
    mp_dtEndDate = Value
End Property

Public Property Get StartDate() As Date
    StartDate = mp_dtStartDate
End Property

Public Property Let StartDate(ByVal Value As Date)
    mp_dtStartDate = Value
End Property

Public Property Get StyleIndex() As String
    StyleIndex = mp_sStyleIndex
End Property

Public Property Let StyleIndex(ByVal Value As String)
    If mp_oControl.StrLib.StrTrim(Value) = "" Then
        Value = "0"
    End If
    mp_sStyleIndex = Value
End Property

Public Property Get Tag() As String
    Tag = mp_sTag
End Property

Public Property Let Tag(ByVal Value As String)
    mp_sTag = Value
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Metrics
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get LeftTrim() As Long
    If mp_lLeft < mp_oControl.Splitter.Right Then
        LeftTrim = mp_oControl.Splitter.Right
    Else
        LeftTrim = mp_lLeft
    End If
End Property

Public Property Get RightTrim() As Long
    If mp_lRight > mp_oControl.mt_RightMargin Then
        RightTrim = mp_oControl.mt_RightMargin
    Else
        RightTrim = mp_lRight
    End If
End Property

Public Property Get Left() As Long
    Left = mp_lLeft
End Property

Friend Property Let f_lLeft(ByVal Value As Long)
    mp_lLeft = Value
End Property

Public Property Get Top() As Long
    Top = mp_lTop
End Property

Friend Property Let f_lTop(ByVal Value As Long)
    mp_lTop = Value
End Property

Public Property Get Right() As Long
    Right = mp_lRight
End Property

Friend Property Let f_lRight(ByVal Value As Long)
    mp_lRight = Value
End Property

Public Property Get Bottom() As Long
    Bottom = mp_lBottom
End Property

Friend Property Let f_lBottom(ByVal Value As Long)
    mp_lBottom = Value
End Property

Public Property Get Visible() As Boolean
    Visible = mp_bVisible
End Property

Friend Property Let f_bVisible(ByVal Value As Boolean)
    mp_bVisible = Value
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TimeBlock"
    oXML.InitializeWriter
    oXML.WriteProperty "Key", mp_sKey
    oXML.WriteProperty "EndDate", mp_dtEndDate
    oXML.WriteProperty "StartDate", mp_dtStartDate
    oXML.WriteProperty "StyleIndex", mp_sStyleIndex
    oXML.WriteProperty "Tag", mp_sTag
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TimeBlock"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadProperty "Key", mp_sKey
    oXML.ReadProperty "EndDate", mp_dtEndDate
    oXML.ReadProperty "StartDate", mp_dtStartDate
    oXML.ReadProperty "StyleIndex", mp_sStyleIndex
    oXML.ReadProperty "Tag", mp_sTag
End Sub

