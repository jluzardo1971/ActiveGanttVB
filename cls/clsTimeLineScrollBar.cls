VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimeLineScrollBar"
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
Private mp_dtStartDate As Date
Private mp_sInterval As String
Private mp_lFactor As Long



'// ---------------------------------------------------------------------------------------------------------------------
'// Initialization & Destruction
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    mp_dtStartDate = mp_oControl.MathLib.CurrentDateTime
    mp_sInterval = "n"
    mp_lFactor = 1
    mp_oControl.f_oHScrollBar2.Enabled = False
End Sub

Private Sub Class_Terminate()
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Object Member Variables
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Interval() As String
    Interval = mp_lFactor & mp_sInterval
End Property

Public Property Let Interval(ByVal Value As String)
    mp_oControl.mp_ProcessInterval Value, mp_sInterval, mp_lFactor
End Property

Friend Property Get f_Interval() As String
    f_Interval = mp_sInterval
End Property

Friend Property Get f_Factor() As Long
    f_Factor = mp_lFactor
End Property

Public Property Get Value() As Integer
    Value = mp_oControl.f_oHScrollBar2.Value
End Property

Public Property Let Value(ByVal lValue As Integer)
    If lValue < 0 Then
        lValue = 0
    End If
    If lValue > mp_oControl.f_oHScrollBar2.Max Then
        lValue = mp_oControl.f_oHScrollBar2.Max
    End If
    mp_oControl.f_oHScrollBar2.Value = lValue
End Property

Public Property Get Enabled() As Boolean
    Enabled = mp_oControl.f_oHScrollBar2.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
    mp_oControl.f_oHScrollBar2.Enabled = Value
End Property

Public Property Get LargeChange() As Integer
    LargeChange = mp_oControl.f_oHScrollBar2.LargeChange
End Property

Public Property Let LargeChange(ByVal Value As Integer)
    mp_oControl.f_oHScrollBar2.LargeChange = Value
End Property

Public Property Get Max() As Integer
    Max = mp_oControl.f_oHScrollBar2.Max
End Property

Public Property Let Max(ByVal Value As Integer)
    mp_oControl.f_oHScrollBar2.Max = Value
End Property

Public Property Get SmallChange() As Integer
    SmallChange = mp_oControl.f_oHScrollBar2.SmallChange
End Property

Public Property Let SmallChange(ByVal Value As Integer)
    mp_oControl.f_oHScrollBar2.SmallChange = Value
End Property

Public Property Get StartDate() As Date
    StartDate = mp_dtStartDate
End Property

Public Property Let StartDate(ByVal Value As Date)
    mp_dtStartDate = Value
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "ScrollBar"
    oXML.InitializeWriter

    oXML.WriteProperty "StartDate", mp_dtStartDate
    oXML.WriteProperty "Interval", mp_sInterval
    oXML.WriteProperty "Factor", mp_lFactor
    
    Dim iValue As Integer
    Dim bEnabled As Boolean
    Dim iLargeChange As Integer
    Dim iMax As Integer
    Dim iSmallChange As Integer
    
    iValue = mp_oControl.f_oHScrollBar2.Value
    bEnabled = mp_oControl.f_oHScrollBar2.Enabled
    iLargeChange = mp_oControl.f_oHScrollBar2.LargeChange
    iMax = mp_oControl.f_oHScrollBar2.Max
    iSmallChange = mp_oControl.f_oHScrollBar2.SmallChange
    
    oXML.WriteProperty "Value", iValue
    oXML.WriteProperty "Enabled", bEnabled
    oXML.WriteProperty "LargeChange", iLargeChange
    oXML.WriteProperty "Max", iMax
    oXML.WriteProperty "SmallChange", iSmallChange
    
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "ScrollBar"
    oXML.SetXML sXML
    oXML.InitializeReader
    
    oXML.ReadProperty "StartDate", mp_dtStartDate
    oXML.ReadProperty "Interval", mp_sInterval
    oXML.ReadProperty "Factor", mp_lFactor
    
    Dim iValue As Integer
    Dim bEnabled As Boolean
    Dim iLargeChange As Integer
    Dim iMax As Integer
    Dim iSmallChange As Integer

    oXML.ReadProperty "Value", iValue
    oXML.ReadProperty "Enabled", bEnabled
    oXML.ReadProperty "LargeChange", iLargeChange
    oXML.ReadProperty "Max", iMax
    oXML.ReadProperty "SmallChange", iSmallChange
    
    mp_oControl.f_oHScrollBar2.Value = iValue
    mp_oControl.f_oHScrollBar2.Enabled = bEnabled
    mp_oControl.f_oHScrollBar2.LargeChange = iLargeChange
    mp_oControl.f_oHScrollBar2.Max = iMax
    mp_oControl.f_oHScrollBar2.SmallChange = iSmallChange
End Sub
