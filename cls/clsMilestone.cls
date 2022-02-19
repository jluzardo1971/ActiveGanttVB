VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMilestone"
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
Private mp_clsPredecessors As clsPredecessors
Private mp_dtMilestoneDate As Date
Private mp_sCaption As String
Private mp_sLayerIndex As String
Private mp_oPicture As StdPicture
Private mp_sRowKey As String
Private mp_sStyleIndex As String
Private mp_sTag As String
Private mp_yAllowedMovement As E_MOVEMENTTYPE
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
    Set mp_clsPredecessors = New clsPredecessors
    mp_clsPredecessors.Initialize mp_oControl, Nothing, Me
    mp_dtMilestoneDate = mp_oControl.MathLib.CurrentDateTime
    mp_sCaption = ""
    mp_sLayerIndex = "0"
    Set mp_oPicture = Nothing
    mp_sRowKey = ""
    mp_sStyleIndex = "0"
    mp_sTag = ""
    mp_yAllowedMovement = E_MOVEMENTTYPE.MT_UNRESTRICTED
    '// Metrics
    mp_lLeft = 0
    mp_lTop = 0
    mp_lRight = 0
    mp_lBottom = 0
    mp_bVisible = False
End Sub

Private Sub Class_Terminate()
    Set mp_clsPredecessors = Nothing
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// clsItemBase
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Key() As String
    Key = mp_sKey
End Property

Public Property Let Key(ByVal Value As String)
    mp_oControl.Milestones.oCollection.mp_SetKey mp_sKey, Value, 50330
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

Public Property Get Predecessors() As clsPredecessors
    Set Predecessors = mp_clsPredecessors
End Property

Public Property Get MilestoneDate() As Date
    MilestoneDate = mp_dtMilestoneDate
End Property

Public Property Let MilestoneDate(ByVal Value As Date)
    mp_dtMilestoneDate = Value
End Property

Public Property Get Picture() As StdPicture
    Set Picture = mp_oPicture
End Property

Public Property Set Picture(ByVal Value As StdPicture)
    Set mp_oPicture = Value
End Property
Public Property Get Caption() As String
    Caption = mp_sCaption
End Property

Public Property Let Caption(ByVal Value As String)
    mp_sCaption = Value
End Property

Public Property Get LayerIndex() As String
    LayerIndex = mp_sLayerIndex
End Property

Public Property Let LayerIndex(ByVal Value As String)
    If mp_oControl.StrLib.StrTrim(Value) = "" Then
        Value = "0"
    End If
    mp_sLayerIndex = Value
End Property

Public Property Get RowKey() As String
    RowKey = mp_sRowKey
End Property

Public Property Let RowKey(ByVal Value As String)
    If mp_oControl.Milestones.oCollection.m_bIgnoreKeyChecks = False Then
        If mp_oControl.Rows.oCollection.m_bDoesKeyExist(Value) = False Then
            mp_oControl.mp_ErrorReport 50381, "Invalid Row Key", "ActiveGanttVBCtl.clsMilestone.Let RowKey"
            Exit Property
        End If
    End If
    mp_sRowKey = Value
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

Public Property Get AllowedMovement() As E_MOVEMENTTYPE
    AllowedMovement = mp_yAllowedMovement
End Property

Public Property Let AllowedMovement(ByVal Value As E_MOVEMENTTYPE)
    mp_yAllowedMovement = Value
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

Friend Property Get f_bLeftVisible() As Boolean
    If LeftTrim = mp_lLeft Then
        f_bLeftVisible = True
    Else
        f_bLeftVisible = False
    End If
End Property

Friend Property Get f_bRightVisible() As Boolean
    If RightTrim = mp_lRight Then
        f_bRightVisible = True
    Else
        f_bRightVisible = False
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
'// Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Function InConflict() As Boolean
    InConflict = mp_oControl.mp_bDetectConflict(MilestoneDate, MilestoneDate, mp_sRowKey, Index, mp_sLayerIndex, E_TYPE.TP_MILESTONE)
End Function

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Milestone"
    oXML.InitializeWriter
    oXML.WriteProperty "Key", mp_sKey
    oXML.WriteProperty "MilestoneDate", mp_dtMilestoneDate
    oXML.WriteProperty "Caption", mp_sCaption
    oXML.WriteProperty "LayerIndex", mp_sLayerIndex
    oXML.WritePropertyPicture "Picture", mp_oPicture
    oXML.WriteProperty "RowKey", mp_sRowKey
    oXML.WriteProperty "StyleIndex", mp_sStyleIndex
    oXML.WriteProperty "Tag", mp_sTag
    oXML.WriteProperty "AllowedMovement", mp_yAllowedMovement
    oXML.WriteObject mp_clsPredecessors.GetXML
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Milestone"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadProperty "Key", mp_sKey
    oXML.ReadProperty "MilestoneDate", mp_dtMilestoneDate
    oXML.ReadProperty "Caption", mp_sCaption
    oXML.ReadProperty "LayerIndex", mp_sLayerIndex
    oXML.ReadPropertyPicture "Picture", mp_oPicture
    oXML.ReadProperty "RowKey", mp_sRowKey
    oXML.ReadProperty "StyleIndex", mp_sStyleIndex
    oXML.ReadProperty "Tag", mp_sTag
    oXML.ReadProperty "AllowedMovement", mp_yAllowedMovement
    mp_clsPredecessors.SetXML oXML.ReadObject("Predecessors")
End Sub


