VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPercentageGroups"
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
    mp_oCollection.Initialize Value, "PercentageGroup"
    
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

Public Property Get Item(ByVal Index As String) As clsPercentageGroup
    Set Item = mp_oCollection.m_oItem(Index, 50349, 50350, 50351, 50352)
End Property

Friend Property Get oCollection() As clsCollectionBase
    Set oCollection = mp_oCollection
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub Add(ByVal Key As String, Optional ByVal Visible As Boolean = True, Optional ByVal StyleIndex As String = "0")
    mp_oCollection.AddMode = True
    Dim oPercentageGroup As New clsPercentageGroup
    oPercentageGroup.Initialize mp_oControl
    oPercentageGroup.StyleIndex = StyleIndex
    oPercentageGroup.Visible = Visible
    oPercentageGroup.Key = Key
    mp_oCollection.m_Add oPercentageGroup, Key, 50337, 50338
    Set oPercentageGroup = Nothing
End Sub

Public Sub Clear()
    mp_oControl.Percentages.Clear
    mp_oCollection.m_Clear
End Sub

Public Sub Remove(ByVal Index As String)
    Dim sRIndex As String
    Dim sRKey As String
    mp_oCollection.m_GetKeyAndIndex Index, sRKey, sRIndex
    mp_oControl.Percentages.oCollection.m_CollectionRemoveWhere "PercentageGroupIndex", sRKey, sRIndex
    mp_oCollection.m_Remove Index, 50369, 50370, 50371, 50372
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim lIndex As Long
    Dim oPercentageGroup As clsPercentageGroup
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "PercentageGroups"
    oXML.InitializeWriter
    For lIndex = 1 To Count
        Set oPercentageGroup = mp_oCollection.m_oReturnArrayElement(lIndex)
        oXML.WriteObject oPercentageGroup.GetXML
    Next lIndex
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim lIndex As Long
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "PercentageGroups"
    oXML.SetXML sXML
    oXML.InitializeReader
    mp_oCollection.m_Clear
    For lIndex = 1 To oXML.ReadCollectionCount
        Dim oPercentageGroup As New clsPercentageGroup
        oPercentageGroup.Initialize mp_oControl
        oPercentageGroup.SetXML oXML.ReadCollectionObject(lIndex)
        mp_oCollection.AddMode = True
        mp_oCollection.m_Add oPercentageGroup, oPercentageGroup.Key, 50000, 50000
        Set oPercentageGroup = Nothing
    Next lIndex
End Sub





