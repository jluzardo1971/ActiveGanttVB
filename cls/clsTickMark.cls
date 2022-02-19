VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTickMark"
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
Private mp_bDisplayCaption As Boolean
Private mp_bModulus As Boolean
Private mp_clsTickMarks As clsTickMarks
Private mp_sCaptionFormat As String
Private mp_sTag As String
Private mp_yInterval As Integer
Private mp_yTickMarkType As E_TICKMARKTYPES

'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl, ByRef oTickMarks As clsTickMarks)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// clsItemBase
    mp_sKey = ""
    mp_lIndex = 0
    '// Object Member Variables
    mp_bDisplayCaption = False
    mp_bModulus = False
    Set mp_clsTickMarks = oTickMarks
    mp_sCaptionFormat = ""
    mp_sTag = ""
    mp_yInterval = 0
    mp_yTickMarkType = E_TICKMARKTYPES.TLT_SMALL
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
    mp_clsTickMarks.oCollection.mp_SetKey mp_sKey, Value, 50000
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

Public Property Get DisplayCaption() As Boolean
    DisplayCaption = mp_bDisplayCaption
End Property

Public Property Let DisplayCaption(ByVal Value As Boolean)
    mp_bDisplayCaption = Value
End Property

Public Property Get Modulus() As Boolean
    Modulus = mp_bModulus
End Property

Public Property Let Modulus(ByVal Value As Boolean)
    mp_bModulus = Value
End Property

Public Property Get CaptionFormat() As String
    CaptionFormat = mp_sCaptionFormat
End Property

Public Property Let CaptionFormat(ByVal Value As String)
    mp_sCaptionFormat = Value
End Property

Public Property Get Tag() As String
    Tag = mp_sTag
End Property

Public Property Let Tag(ByVal Value As String)
    mp_sTag = Value
End Property

Public Property Get Interval() As Integer
    Interval = mp_yInterval
End Property

Public Property Let Interval(ByVal Value As Integer)
    mp_yInterval = Value
End Property

Public Property Get TickMarkType() As E_TICKMARKTYPES
    TickMarkType = mp_yTickMarkType
End Property

Public Property Let TickMarkType(ByVal Value As E_TICKMARKTYPES)
    mp_yTickMarkType = Value
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TickMark"
    oXML.InitializeWriter
    oXML.WriteProperty "DisplayCaption", mp_bDisplayCaption
    oXML.WriteProperty "Modulus", mp_bModulus
    oXML.WriteProperty "CaptionFormat", mp_sCaptionFormat
    oXML.WriteProperty "Tag", mp_sTag
    oXML.WriteProperty "Interval", mp_yInterval
    oXML.WriteProperty "TickMarkType", mp_yTickMarkType
    oXML.WriteProperty "Key", mp_sKey
    oXML.WriteProperty "Index", mp_lIndex
    
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TickMark"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadProperty "DisplayCaption", mp_bDisplayCaption
    oXML.ReadProperty "Modulus", mp_bModulus
    oXML.ReadProperty "CaptionFormat", mp_sCaptionFormat
    oXML.ReadProperty "Tag", mp_sTag
    oXML.ReadProperty "Interval", mp_yInterval
    oXML.ReadProperty "TickMarkType", mp_yTickMarkType
    oXML.ReadProperty "Key", mp_sKey
    oXML.ReadProperty "Index", mp_lIndex
End Sub
