VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsView"
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
Private mp_oClientArea As clsClientArea
Private mp_oTimeLine As clsTimeLine
Private mp_sTag As String
Private mp_sScrollInterval As String
Private mp_sInterval As String
Private mp_lFactor As Long


'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// clsItemBase
    mp_sKey = ""
    mp_lIndex = 0
    '// Object Member Variables
    Set mp_oTimeLine = New clsTimeLine
    mp_oTimeLine.Initialize mp_oControl, Me
    Set mp_oClientArea = New clsClientArea
    mp_oClientArea.Initialize mp_oControl, mp_oTimeLine
    mp_sTag = ""
    mp_sScrollInterval = ""
    mp_sInterval = ""
    mp_lFactor = 0
End Sub

Private Sub Class_Terminate()
    Set mp_oTimeLine = Nothing
    Set mp_oClientArea = Nothing
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// clsItemBase
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Key() As String
    Key = mp_sKey
End Property

Public Property Let Key(ByVal Value As String)
    mp_oControl.Views.oCollection.mp_SetKey mp_sKey, Value, 50211
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

Public Property Get TimeLine() As clsTimeLine
    Set TimeLine = mp_oTimeLine
End Property

Public Property Get ClientArea() As clsClientArea
    Set ClientArea = mp_oClientArea
End Property

Public Property Get Tag() As String
    Tag = mp_sTag
End Property

Public Property Let Tag(ByVal Value As String)
    mp_sTag = Value
End Property

Friend Property Get f_ScrollInterval() As String
    f_ScrollInterval = mp_sScrollInterval
End Property

Public Property Get Interval() As String
    Interval = mp_lFactor & mp_sInterval
End Property

Public Property Let Interval(ByVal Value As String)
    mp_oControl.mp_ProcessInterval Value, mp_sInterval, mp_lFactor
    Select Case mp_sInterval
        Case "s"
            mp_sScrollInterval = "n"
        Case "n"
            mp_sScrollInterval = "h"
        Case "h"
            mp_sScrollInterval = "d"
        Case "d"
            mp_sScrollInterval = "ww"
        Case "w"
            mp_sScrollInterval = "ww"
        Case "y"
            mp_sScrollInterval = "ww"
        Case "ww"
            mp_sScrollInterval = "m"
        Case "m"
            mp_sScrollInterval = "yyyy"
        Case "yyyy"
            mp_sScrollInterval = "yyyy"
    End Select
End Property

Friend Property Get f_Factor() As Long
    f_Factor = mp_lFactor
End Property

Friend Property Get f_Interval() As String
    f_Interval = mp_sInterval
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "View"
    oXML.InitializeWriter
    oXML.WriteProperty "Tag", mp_sTag
    oXML.WriteProperty "ScrollInterval", mp_sScrollInterval
    oXML.WriteProperty "Interval", mp_sInterval
    oXML.WriteProperty "Factor", mp_lFactor
    oXML.WriteObject mp_oClientArea.GetXML
    oXML.WriteObject mp_oTimeLine.GetXML
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "View"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadProperty "Tag", mp_sTag
    oXML.ReadProperty "ScrollInterval", mp_sScrollInterval
    oXML.ReadProperty "Interval", mp_sInterval
    oXML.ReadProperty "Factor", mp_lFactor
    mp_oClientArea.SetXML oXML.ReadObject("ClientArea")
    mp_oTimeLine.SetXML oXML.ReadObject("TimeLine")
End Sub
