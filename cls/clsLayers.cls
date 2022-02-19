VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLayers"
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
Private mp_oDefaultLayer As clsLayer


'// ---------------------------------------------------------------------------------------------------------------------
'// Initialization & Destruction
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oCollection = New clsCollectionBase
    mp_oCollection.Initialize Value, "Layer"
    Set mp_oDefaultLayer = New clsLayer
    mp_oDefaultLayer.Initialize Value
    
End Sub

Private Sub Class_Terminate()
    Set mp_oDefaultLayer = Nothing
    Set mp_oCollection = Nothing
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Properties
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Count() As Long
    Count = mp_oCollection.m_lCount
End Property

Public Property Get Item(ByVal Index As String) As clsLayer
    Set Item = mp_oCollection.m_oItem(Index, 50341, 50342, 50343, 50344)
End Property

Friend Property Get FItem(ByVal Index As String) As clsLayer
    If Index = "0" Then
        Set FItem = mp_oDefaultLayer
    Else
        Set FItem = mp_oCollection.m_oItem(Index, 50357, 50358, 50359, 50360)
    End If
End Property

Friend Property Get oCollection() As clsCollectionBase
    Set oCollection = mp_oCollection
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub Add(ByVal Key As String, Optional ByVal Visible As Boolean = True)
    mp_oCollection.AddMode = True
    Dim oLayer As New clsLayer
    oLayer.Initialize mp_oControl
    oLayer.Key = Key
    oLayer.Visible = Visible
    mp_oCollection.m_Add oLayer, Key, 50333, 50334
    Set oLayer = Nothing
End Sub

Public Sub Clear()
    mp_oControl.Tasks.oCollection.m_CollectionRemoveWhereNot "LayerIndex", "0"
    mp_oControl.Milestones.oCollection.m_CollectionRemoveWhereNot "LayerIndex", "0"
    mp_oControl.CurrentLayer = "0"
    mp_oCollection.m_Clear
End Sub

Public Sub Remove(ByVal Index As String)
    Dim sRIndex As String
    Dim sRKey As String
    mp_oCollection.m_GetKeyAndIndex Index, sRKey, sRIndex
    mp_oControl.Tasks.oCollection.m_CollectionRemoveWhere "LayerIndex", sRKey, sRIndex
    mp_oControl.Milestones.oCollection.m_CollectionRemoveWhere "LayerIndex", sRKey, sRIndex
    mp_oControl.CurrentLayer = "0"
    mp_oCollection.m_Remove Index, 50361, 50362, 50363, 50364
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim lIndex As Long
    Dim oLayer As clsLayer
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Layers"
    oXML.InitializeWriter
    For lIndex = 1 To Count
        Set oLayer = mp_oCollection.m_oReturnArrayElement(lIndex)
        oXML.WriteObject oLayer.GetXML
    Next lIndex
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim lIndex As Long
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Layers"
    oXML.SetXML sXML
    oXML.InitializeReader
    mp_oCollection.m_Clear
    For lIndex = 1 To oXML.ReadCollectionCount
        Dim oLayer As New clsLayer
        oLayer.Initialize mp_oControl
        oLayer.SetXML oXML.ReadCollectionObject(lIndex)
        mp_oCollection.AddMode = True
        mp_oCollection.m_Add oLayer, oLayer.Key, 50000, 50000
        Set oLayer = Nothing
    Next lIndex
End Sub
