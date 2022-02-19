VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTierColors"
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

Private mp_yTierType As E_TIERTYPE

'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl, ByVal yTierType As E_TIERTYPE)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oCollection = New clsCollectionBase
    mp_oCollection.Initialize Value, "TierColor"
    
    mp_yTierType = yTierType
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

Public Property Get Item(ByVal Index As String) As clsTierColor
    Set Item = mp_oCollection.m_oItem(Index, 50000, 50000, 50000, 50000)
End Property

Friend Property Get oCollection() As clsCollectionBase
    Set oCollection = mp_oCollection
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Methods
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Add(ByVal BackColor As OLE_COLOR, ByVal ForeColor As OLE_COLOR, Optional ByVal Key As String = "")
    mp_oCollection.AddMode = True
    Dim oTierColor As New clsTierColor
    oTierColor.Initialize mp_oControl, Me
    oTierColor.BackColor = BackColor
    oTierColor.ForeColor = ForeColor
    oTierColor.Key = Key
    mp_oCollection.m_Add oTierColor, Key, 50000, 500
    Set oTierColor = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Custom Methods
'// ---------------------------------------------------------------------------------------------------------------------

Private Function mp_CollectionName() As String
    Select Case mp_yTierType
        Case E_TIERTYPE.ST_MINUTE
            mp_CollectionName = "TierColorsMinuteInterval"
        Case E_TIERTYPE.ST_HOUR
            mp_CollectionName = "TierColorsHourInterval"
        Case E_TIERTYPE.ST_DAY
            mp_CollectionName = "TierColorsDayInterval"
        Case E_TIERTYPE.ST_DAYOFWEEK
            mp_CollectionName = "TierColorsDayOfWeekInterval"
        Case E_TIERTYPE.ST_DAYOFYEAR
            mp_CollectionName = "TierColorsDayOfYearInterval"
        Case E_TIERTYPE.ST_WEEK
            mp_CollectionName = "TierColorsWeekInterval"
        Case E_TIERTYPE.ST_MONTH
            mp_CollectionName = "TierColorsMonthInterval"
        Case E_TIERTYPE.ST_QUARTER
            mp_CollectionName = "TierColorsQuarterInterval"
        Case E_TIERTYPE.ST_YEAR
            mp_CollectionName = "TierColorsYearInterval"
    End Select
End Function

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim lIndex As Long
    Dim oTierColor As clsTierColor
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, mp_CollectionName
    oXML.InitializeWriter
    For lIndex = 1 To Count
        Set oTierColor = mp_oCollection.m_oReturnArrayElement(lIndex)
        oXML.WriteObject oTierColor.GetXML
    Next lIndex
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim lIndex As Long
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, mp_CollectionName
    oXML.SetXML sXML
    oXML.InitializeReader
    mp_oCollection.m_Clear
    For lIndex = 1 To oXML.ReadCollectionCount
        Dim oTierColor As New clsTierColor
        oTierColor.Initialize mp_oControl, Me
        oTierColor.SetXML oXML.ReadCollectionObject(lIndex)
        mp_oCollection.AddMode = True
        mp_oCollection.m_Add oTierColor, oTierColor.Key, 50000, 50000
        Set oTierColor = Nothing
    Next lIndex
End Sub
