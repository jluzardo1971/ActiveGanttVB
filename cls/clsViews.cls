VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsViews"
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
Private mp_oDefaultView As clsView


'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oCollection = New clsCollectionBase
    mp_oCollection.Initialize Value, "View"
    
    Set mp_oDefaultView = New clsView
    mp_oDefaultView.Initialize mp_oControl
    
    mp_oDefaultView.Interval = "1n"
    mp_oDefaultView.TimeLine.TierArea.UpperTier.TierType = E_TIERTYPE.ST_MONTH
    mp_oDefaultView.TimeLine.TierArea.LowerTier.TierType = E_TIERTYPE.ST_DAYOFWEEK
    mp_oDefaultView.TimeLine.TickMarkArea.Interval = "5n"
    mp_oDefaultView.ClientArea.ToolTipFormat = "Long Time"
    mp_oDefaultView.TimeLine.TickMarkArea.TickMarks.Add 0, E_TICKMARKTYPES.TLT_BIG, True, "Hh:Nnam/pm", False
    mp_oDefaultView.TimeLine.TickMarkArea.TickMarks.Add 15, E_TICKMARKTYPES.TLT_MEDIUM, False, "", False
    mp_oDefaultView.TimeLine.TickMarkArea.TickMarks.Add 45, E_TICKMARKTYPES.TLT_MEDIUM, False, "", False
    mp_oDefaultView.TimeLine.TickMarkArea.TickMarks.Add 1, E_TICKMARKTYPES.TLT_SMALL, False, "", True
End Sub

Private Sub Class_Terminate()
    Set mp_oCollection = Nothing
    Set mp_oDefaultView = Nothing
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Properties
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Count() As Long
    Count = mp_oCollection.m_lCount
End Property

Public Property Get Item(ByVal Index As String) As clsView
    Set Item = mp_oCollection.m_oItem(Index, 50000, 50000, 50000, 50000)
End Property

Friend Property Get FItem(ByVal Index As String) As clsView
    If Index = "0" Then
        Set FItem = mp_oDefaultView
    Else
        Set FItem = mp_oCollection.m_oItem(Index, 50000, 50000, 50000, 50000)
    End If
End Property

Friend Property Get oCollection() As clsCollectionBase
    Set oCollection = mp_oCollection
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub Add(ByVal Interval As String, ByVal TickMarkAreaInterval As String, ByVal UpperTierType As E_TIERTYPE, ByVal MiddleTierType As E_TIERTYPE, ByVal LowerTierType As E_TIERTYPE, Optional ByVal Key As String = "")
    mp_oCollection.AddMode = True
    Dim oView As New clsView
    oView.Initialize mp_oControl
    oView.Interval = Interval
    oView.TimeLine.TickMarkArea.Interval = TickMarkAreaInterval
    oView.TimeLine.TierArea.UpperTier.TierType = UpperTierType
    oView.TimeLine.TierArea.MiddleTier.TierType = MiddleTierType
    oView.TimeLine.TierArea.LowerTier.TierType = LowerTierType
    oView.Key = Key
    mp_oCollection.m_Add oView, Key, 50000, 50000
    Set oView = Nothing
End Sub

Public Sub Clear()
    mp_oControl.CurrentView = "0"
    mp_oCollection.m_Clear
End Sub

Public Sub Remove(ByVal Index As String)
    mp_oControl.CurrentView = "0"
    mp_oCollection.m_Remove Index, 50000, 50000, 50000, 50000
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim lIndex As Long
    Dim oView As clsView
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Views"
    oXML.InitializeWriter
    For lIndex = 1 To Count
        Set oView = mp_oCollection.m_oReturnArrayElement(lIndex)
        oXML.WriteObject oView.GetXML
    Next lIndex
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim lIndex As Long
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Views"
    oXML.SetXML sXML
    oXML.InitializeReader
    mp_oCollection.m_Clear
    For lIndex = 1 To oXML.ReadCollectionCount
        Dim oView As New clsView
        oView.Initialize mp_oControl
        oView.SetXML oXML.ReadCollectionObject(lIndex)
        mp_oCollection.AddMode = True
        mp_oCollection.m_Add oView, oView.Key, 50000, 50000
        Set oView = Nothing
    Next lIndex
End Sub
