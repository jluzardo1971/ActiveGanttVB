VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPredecessor"
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
Private mp_bVisible As Boolean
Private mp_clsPredecessors As clsPredecessors
Private mp_sPredecessorKey As String
Private mp_sStyleIndex As String
Private mp_sTag As String
Private mp_yPredecessorType As E_OBJECTTYPE


'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl, ByRef oPredecessors As clsPredecessors)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// clsItemBase
    mp_sKey = ""
    mp_lIndex = 0
    '// Object Member Variables
    mp_bVisible = True
    Set mp_clsPredecessors = oPredecessors
    mp_sPredecessorKey = ""
    mp_sStyleIndex = "0"
    mp_sTag = ""
    mp_yPredecessorType = E_OBJECTTYPE.OT_TASK
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
    mp_clsPredecessors.oCollection.mp_SetKey mp_sKey, Value, 50211
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

Public Property Get Visible() As Boolean
    Visible = mp_bVisible
End Property

Public Property Let Visible(ByVal Value As Boolean)
    mp_bVisible = Value
End Property

Public Property Get PredecessorKey() As String
    PredecessorKey = mp_sPredecessorKey
End Property

Public Property Let PredecessorKey(ByVal Value As String)
    mp_sPredecessorKey = Value
End Property

Public Property Get MilestoneKey() As String
    If mp_yPredecessorType = E_OBJECTTYPE.OT_MILESTONE Then
        MilestoneKey = PredecessorKey
    Else
        MilestoneKey = ""
    End If
End Property

Public Property Get TaskKey() As String
    If mp_yPredecessorType = E_OBJECTTYPE.OT_TASK Then
        TaskKey = PredecessorKey
    Else
        TaskKey = ""
    End If
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

Public Property Get PredecessorType() As E_OBJECTTYPE
    PredecessorType = mp_yPredecessorType
End Property

Public Property Let PredecessorType(ByVal Value As E_OBJECTTYPE)
    mp_yPredecessorType = Value
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Predecessor"
    oXML.InitializeWriter
    oXML.WriteProperty "Key", mp_sKey
    oXML.WriteProperty "Visible", mp_bVisible
    oXML.WriteProperty "PredecessorKey", mp_sPredecessorKey
    oXML.WriteProperty "StyleIndex", mp_sStyleIndex
    oXML.WriteProperty "Tag", mp_sTag
    oXML.WriteProperty "PredecessorType", mp_yPredecessorType
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Predecessor"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadProperty "Key", mp_sKey
    oXML.ReadProperty "Visible", mp_bVisible
    oXML.ReadProperty "PredecessorKey", mp_sPredecessorKey
    oXML.ReadProperty "StyleIndex", mp_sStyleIndex
    oXML.ReadProperty "Tag", mp_sTag
    oXML.ReadProperty "PredecessorType", mp_yPredecessorType
End Sub
