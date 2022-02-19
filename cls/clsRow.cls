VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRow"
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
Private mp_bContainer As Boolean
Private mp_bMergeCells As Boolean
Private mp_lHeight As Long
Private mp_sCaption As String
Private mp_oPicture As StdPicture
Private mp_sStyleIndex As String
Private mp_sTag As String
Private mp_oCells As clsCells
Private mp_sClientAreaStyleIndex As String
'// Metrics
Private mp_lTop As Long
Private mp_lBottom As Long
Private mp_bVisible As Boolean



'// ---------------------------------------------------------------------------------------------------------------------
'// Initialization & Destruction
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// clsItemBase
    mp_sKey = ""
    mp_lIndex = 0
    '// Object Member Variables
    mp_bContainer = True
    mp_bMergeCells = False
    mp_lHeight = mp_oControl.DefaultValues.RowHeight
    mp_sCaption = ""
    Set mp_oPicture = Nothing
    mp_sStyleIndex = "0"
    mp_sTag = ""
    Set mp_oCells = New clsCells
    mp_oCells.Initialize mp_oControl, Me
    mp_sClientAreaStyleIndex = ""
    '// Metrics
    mp_lTop = 0
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
    mp_oControl.Rows.oCollection.mp_SetKey mp_sKey, Value, 50223
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

Public Property Get Container() As Boolean
    Container = mp_bContainer
End Property

Public Property Let Container(ByVal Value As Boolean)
    mp_bContainer = Value
End Property

Public Property Get MergeCells() As Boolean
    MergeCells = mp_bMergeCells
End Property

Public Property Let MergeCells(ByVal Value As Boolean)
    mp_bMergeCells = Value
End Property

Public Property Get Height() As Long
    Height = mp_lHeight
End Property

Public Property Let Height(ByVal Value As Long)
    mp_lHeight = Value
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

Public Property Get Cells() As clsCells
    Set Cells = mp_oCells
End Property

Public Property Get ClientAreaStyleIndex() As String
    ClientAreaStyleIndex = mp_sClientAreaStyleIndex
End Property

Public Property Let ClientAreaStyleIndex(ByVal Value As String)
    mp_sClientAreaStyleIndex = Value
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Metrics
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Left() As Long
    Left = mp_oControl.mt_LeftMargin()
End Property

Public Property Get Top() As Long
    Top = mp_lTop
End Property

Friend Property Let f_lTop(ByVal Value As Long)
    mp_lTop = Value
End Property

Public Property Get Right() As Long
    Right = mp_oControl.Splitter.Left()
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
    oXML.Initialize mp_oControl, "Row"
    oXML.InitializeWriter
    oXML.WriteProperty "Key", mp_sKey
    oXML.WriteProperty "Container", mp_bContainer
    oXML.WriteProperty "MergeCells", mp_bMergeCells
    oXML.WriteProperty "Height", mp_lHeight
    oXML.WriteProperty "Caption", mp_sCaption
    oXML.WritePropertyPicture "Picture", mp_oPicture
    oXML.WriteProperty "StyleIndex", mp_sStyleIndex
    oXML.WriteProperty "Tag", mp_sTag
    oXML.WriteProperty "ClientAreaStyleIndex", mp_sClientAreaStyleIndex
    oXML.WriteObject mp_oCells.GetXML
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Row"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadProperty "Key", mp_sKey
    oXML.ReadProperty "Container", mp_bContainer
    oXML.ReadProperty "MergeCells", mp_bMergeCells
    oXML.ReadProperty "Height", mp_lHeight
    oXML.ReadProperty "Caption", mp_sCaption
    oXML.ReadPropertyPicture "Picture", mp_oPicture
    oXML.ReadProperty "StyleIndex", mp_sStyleIndex
    oXML.ReadProperty "Tag", mp_sTag
    oXML.ReadProperty "ClientAreaStyleIndex", mp_sClientAreaStyleIndex
    mp_oCells.SetXML oXML.ReadObject("Cells")
End Sub
