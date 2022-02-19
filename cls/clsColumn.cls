VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColumn"
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
'// Object member variables
Private mp_lWidth As Long
Private mp_sCaption As String
Private mp_oPicture As StdPicture
Private mp_sStyleIndex As String
Private mp_sTag As String
'// Metrics
Private mp_lLeft As Long
Private mp_lRight As Long
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
    '// Object member variables
    mp_lWidth = mp_oControl.DefaultValues.ColumnWidth
    mp_sCaption = ""
    Set mp_oPicture = Nothing
    mp_sStyleIndex = "0"
    mp_sTag = ""
    '// Metrics
    mp_lLeft = 0
    mp_lRight = 0
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
    mp_oControl.Columns.oCollection.mp_SetKey mp_sKey, Value, 50000
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

Public Property Get Width() As Long
    Width = mp_lWidth
End Property

Public Property Let Width(ByVal Value As Long)
    mp_lWidth = Value
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

'// ---------------------------------------------------------------------------------------------------------------------
'// Metrics
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get LeftTrim() As Long
    If mp_lLeft < mp_oControl.mt_LeftMargin Then
        LeftTrim = mp_oControl.mt_LeftMargin
    Else
        LeftTrim = mp_lLeft
    End If
End Property

Public Property Get RightTrim() As Long
    If mp_lRight > mp_oControl.Splitter.Left Then
        RightTrim = mp_oControl.Splitter.Left
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
    Top = mp_oControl.CurrentViewObject.TimeLine.Top
End Property

Public Property Get Right() As Long
    Right = mp_lRight
End Property

Friend Property Let f_lRight(ByVal Value As Long)
    mp_lRight = Value
End Property

Public Property Get Bottom() As Long
    Bottom = mp_oControl.CurrentViewObject.TimeLine.Bottom
End Property

Public Property Get Visible() As Boolean
    Visible = mp_bVisible
End Property

Friend Property Let f_bVisible(ByVal Value As Boolean)
    If Value = False Then
        mp_lLeft = 0
        mp_lRight = 0
    End If
    mp_bVisible = Value
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Column"
    oXML.InitializeWriter
    oXML.WriteProperty "Key", mp_sKey
    oXML.WriteProperty "Width", mp_lWidth
    oXML.WriteProperty "Caption", mp_sCaption
    oXML.WritePropertyPicture "Picture", mp_oPicture
    oXML.WriteProperty "StyleIndex", mp_sStyleIndex
    oXML.WriteProperty "Tag", mp_sTag

    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Column"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadProperty "Key", mp_sKey
    oXML.ReadProperty "Width", mp_lWidth
    oXML.ReadProperty "Caption", mp_sCaption
    oXML.ReadPropertyPicture "Picture", mp_oPicture
    oXML.ReadProperty "StyleIndex", mp_sStyleIndex
    oXML.ReadProperty "Tag", mp_sTag

End Sub
