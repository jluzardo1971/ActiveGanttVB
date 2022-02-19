VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCell"
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
Private mp_lIndex As Long
Private mp_sKey As String
'// Object Member Variables
Private mp_sCaption As String
Private mp_oPicture As StdPicture
Private mp_sStyleIndex As String
Private mp_sTag As String
Private mp_oCells As clsCells

'// ---------------------------------------------------------------------------------------------------------------------
'// Initialization & Destruction
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl, ByRef oCells As clsCells)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// clsItemBase
    mp_sKey = ""
    mp_lIndex = 0
    '// Object Member Variables
    mp_sCaption = ""
    Set mp_oPicture = Nothing
    mp_sStyleIndex = "0"
    mp_sTag = ""
    Set mp_oCells = oCells
End Sub

Private Sub Class_Terminate()
    Set mp_oControl = Nothing
    Set mp_oCells = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// clsItemBase
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Key() As String
    Key = mp_sKey
End Property

Public Property Let Key(ByVal Value As String)
    mp_oCells.oCollection.mp_SetKey mp_sKey, Value, 50000
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

Public Property Get RowKey() As String
    RowKey = mp_oCells.Row.Key
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Metrics
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Left() As Long
    Left = mp_oControl.Columns.Item(mp_lIndex).Left
End Property

Public Property Get Top() As Long
    Top = mp_oCells.Row.Top
End Property

Public Property Get Right() As Long
    Right = mp_oControl.Columns.Item(mp_lIndex).Right
End Property

Public Property Get Bottom() As Long
    Bottom = mp_oCells.Row.Bottom
End Property

Public Property Get LeftTrim() As Long
    LeftTrim = mp_oControl.Columns.Item(mp_lIndex).LeftTrim
End Property

Public Property Get RightTrim() As Long
    RightTrim = mp_oControl.Columns.Item(mp_lIndex).RightTrim
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Cell"
    oXML.InitializeWriter
    oXML.WriteProperty "Key", mp_sKey
    oXML.WriteProperty "Caption", mp_sCaption
    oXML.WritePropertyPicture "Picture", mp_oPicture
    oXML.WriteProperty "Tag", mp_sTag
    oXML.WriteProperty "StyleIndex", mp_sStyleIndex
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Cell"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadProperty "Key", mp_sKey
    oXML.ReadProperty "Caption", mp_sCaption
    oXML.ReadPropertyPicture "Picture", mp_oPicture
    oXML.ReadProperty "Tag", mp_sTag
    oXML.ReadProperty "StyleIndex", mp_sStyleIndex
End Sub



