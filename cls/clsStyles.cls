VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStyles"
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
Private mp_oDefaultStyle As clsStyle

'// ---------------------------------------------------------------------------------------------------------------------
'// Initialization & Destruction
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oCollection = New clsCollectionBase
    mp_oCollection.Initialize Value, "Style"
    Set mp_oDefaultStyle = New clsStyle
    mp_oDefaultStyle.Initialize Value
End Sub

Private Sub Class_Terminate()
    Set mp_oDefaultStyle = Nothing
    Set mp_oCollection = Nothing
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Properties
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Count() As Long
    Count = mp_oCollection.m_lCount
End Property

Public Property Get Item(ByVal Index As String) As clsStyle
    Set Item = mp_oCollection.m_oItem(Index, 50185, 50186, 50187, 50188)
End Property

Friend Property Get FItem(ByVal Index As String) As clsStyle
    If Index = "0" Then
        Set FItem = mp_oDefaultStyle
    Else
        Set FItem = mp_oCollection.m_oItem(Index, 50193, 50194, 50195, 50188)
    End If
End Property

Friend Property Get oCollection() As clsCollectionBase
    Set oCollection = mp_oCollection
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub Add(Optional ByVal Key As String = "")
    mp_oCollection.AddMode = True
    Dim oStyle As New clsStyle
    Key = mp_oControl.StrLib.StrTrim(Key)
    oStyle.Initialize mp_oControl
    oStyle.Key = Key
    mp_oCollection.m_Add oStyle, Key, 50183, 50184
    Set oStyle = Nothing
End Sub

Public Sub Clear()
    Dim lIndex As Long
    Dim oRow As clsRow
    Dim oTask As clsTask
    Dim oMilestone As clsMilestone

    mp_oControl.Columns.oCollection.m_CollectionChangeAll "StyleIndex", "0"
    
    mp_oControl.Rows.oCollection.m_CollectionChangeAll "StyleIndex", "0"
    For lIndex = 1 To mp_oControl.Rows.Count
        Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex)
        oRow.Cells.oCollection.m_CollectionChangeAll "StyleIndex", "0"
    Next lIndex
    
    mp_oControl.Tasks.oCollection.m_CollectionChangeAll "StyleIndex", "0"
    For lIndex = 1 To mp_oControl.Tasks.Count
        Set oTask = mp_oControl.Tasks.oCollection.m_oReturnArrayElement(lIndex)
        oTask.Predecessors.oCollection.m_CollectionChangeAll "StyleIndex", "0"
    Next lIndex
    
    mp_oControl.Milestones.oCollection.m_CollectionChangeAll "StyleIndex", "0"
    For lIndex = 1 To mp_oControl.Milestones.Count
        Set oMilestone = mp_oControl.Milestones.oCollection.m_oReturnArrayElement(lIndex)
        oMilestone.Predecessors.oCollection.m_CollectionChangeAll "StyleIndex", "0"
    Next lIndex

    mp_oControl.TimeBlocks.oCollection.m_CollectionChangeAll "StyleIndex", "0"
    mp_oControl.PercentageGroups.oCollection.m_CollectionChangeAll "StyleIndex", "0"

    mp_oCollection.m_Clear
End Sub

Public Sub Remove(ByVal Index As String)
    Dim lIndex As Long
    Dim oRow As clsRow
    Dim oTask As clsTask
    Dim oMilestone As clsMilestone
    Dim sRIndex As String
    Dim sRKey As String
    
    mp_oCollection.m_GetKeyAndIndex Index, sRKey, sRIndex
    
    mp_oControl.Columns.oCollection.m_CollectionChange "StyleIndex", sRKey, sRIndex, "0"
    
    mp_oControl.Rows.oCollection.m_CollectionChange "StyleIndex", sRKey, sRIndex, "0"
    For lIndex = 1 To mp_oControl.Rows.Count
        Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex)
        oRow.Cells.oCollection.m_CollectionChange "StyleIndex", sRKey, sRIndex, "0"
    Next lIndex
    
    mp_oControl.Tasks.oCollection.m_CollectionChange "StyleIndex", sRKey, sRIndex, "0"
    For lIndex = 1 To mp_oControl.Tasks.Count
        Set oTask = mp_oControl.Tasks.oCollection.m_oReturnArrayElement(lIndex)
        oTask.Predecessors.oCollection.m_CollectionChange "StyleIndex", sRKey, sRIndex, "0"
    Next lIndex
    
    mp_oControl.Milestones.oCollection.m_CollectionChange "StyleIndex", sRKey, sRIndex, "0"
    For lIndex = 1 To mp_oControl.Milestones.Count
        Set oMilestone = mp_oControl.Milestones.oCollection.m_oReturnArrayElement(lIndex)
        oMilestone.Predecessors.oCollection.m_CollectionChange "StyleIndex", sRKey, sRIndex, "0"
    Next lIndex

    mp_oControl.TimeBlocks.oCollection.m_CollectionChange "StyleIndex", sRKey, sRIndex, "0"
    mp_oControl.PercentageGroups.oCollection.m_CollectionChange "StyleIndex", sRKey, sRIndex, "0"
    mp_oCollection.m_Remove Index, 50189, 50190, 50191, 50192
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim lIndex As Long
    Dim oStyle As clsStyle
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Styles"
    oXML.InitializeWriter
    For lIndex = 1 To Count
        Set oStyle = mp_oCollection.m_oReturnArrayElement(lIndex)
        oXML.WriteObject oStyle.GetXML
    Next lIndex
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim lIndex As Long
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Styles"
    oXML.SetXML sXML
    oXML.InitializeReader
    mp_oCollection.m_Clear
    For lIndex = 1 To oXML.ReadCollectionCount
        Dim oStyle As New clsStyle
        oStyle.Initialize mp_oControl
        oStyle.SetXML oXML.ReadCollectionObject(lIndex)
        mp_oCollection.AddMode = True
        mp_oCollection.m_Add oStyle, oStyle.Key, 50000, 50000
        Set oStyle = Nothing
    Next lIndex
End Sub
