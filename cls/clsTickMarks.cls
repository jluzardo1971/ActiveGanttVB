VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTickMarks"
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
Private mp_oCollection As clsCollectionBase


'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oCollection = New clsCollectionBase
    mp_oCollection.Initialize Value, "TickMark"
End Sub

Private Sub Class_Terminate()
    Set mp_oCollection = Nothing
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Properties
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Count() As Long
    Count = mp_oCollection.m_lCount
End Property

Public Property Get Item(ByVal Index As String) As clsTickMark
    Set Item = mp_oCollection.m_oItem(Index, 50000, 50000, 50000, 50000)
End Property

Friend Property Get oCollection() As clsCollectionBase
    Set oCollection = mp_oCollection
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub Add(ByVal Interval As Integer, ByVal TickMarkType As E_TICKMARKTYPES, Optional ByVal DisplayCaption As Boolean = False, Optional ByVal CaptionFormat As String = "", Optional ByVal Modulus As Boolean = False, Optional ByVal Key As String = "")
    mp_oCollection.AddMode = True
    Dim oTickMark As New clsTickMark
    oTickMark.Initialize mp_oControl, Me
    oTickMark.Interval = Interval
    oTickMark.TickMarkType = TickMarkType
    oTickMark.DisplayCaption = DisplayCaption
    oTickMark.CaptionFormat = CaptionFormat
    oTickMark.Modulus = Modulus
    oTickMark.Key = Key
    mp_oCollection.m_Add oTickMark, Key, 50000, 50000
    Set oTickMark = Nothing
End Sub

Public Sub Clear()
    mp_oCollection.m_Clear
End Sub

Public Sub Remove(ByVal Index As String)
    mp_oCollection.m_Remove Index, 50000, 50000, 50000, 50000
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim lIndex As Long
    Dim oTickMark As clsTickMark
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TickMarks"
    oXML.InitializeWriter
    For lIndex = 1 To Count
        Set oTickMark = mp_oCollection.m_oReturnArrayElement(lIndex)
        oXML.WriteObject oTickMark.GetXML
    Next lIndex
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim lIndex As Long
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TickMarks"
    oXML.SetXML sXML
    oXML.InitializeReader
    mp_oCollection.m_Clear
    For lIndex = 1 To oXML.ReadCollectionCount
        Dim oTickMark As New clsTickMark
        oTickMark.Initialize mp_oControl, Me
        oTickMark.SetXML oXML.ReadCollectionObject(lIndex)
        mp_oCollection.AddMode = True
        mp_oCollection.m_Add oTickMark, oTickMark.Key, 50000, 50000
        Set oTickMark = Nothing
    Next lIndex
End Sub
