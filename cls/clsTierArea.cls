VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTierArea"
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
Private mp_oUpperTier As clsTier
Private mp_oMiddleTier As clsTier
Private mp_oLowerTier As clsTier
Private mp_oTierFormat As clsTierFormat
Private mp_oTierAppearance As clsTierAppearance
Private mp_oTimeLine As clsTimeLine

'// -----------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// -----------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl, ByRef oTimeLine As clsTimeLine)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oTimeLine = oTimeLine
    Set mp_oUpperTier = New clsTier
    mp_oUpperTier.Initialize mp_oControl, Me, E_TIERPOSITION.SP_UPPER, True
    Set mp_oMiddleTier = New clsTier
    mp_oMiddleTier.Initialize mp_oControl, Me, E_TIERPOSITION.SP_MIDDLE, True
    Set mp_oLowerTier = New clsTier
    mp_oLowerTier.Initialize mp_oControl, Me, E_TIERPOSITION.SP_LOWER, True
    Set mp_oTierFormat = New clsTierFormat
    mp_oTierFormat.Initialize mp_oControl, True
    Set mp_oTierAppearance = New clsTierAppearance
    mp_oTierAppearance.Initialize mp_oControl, True
End Sub

Private Sub Class_Terminate()
    Set mp_oUpperTier = Nothing
    Set mp_oMiddleTier = Nothing
    Set mp_oLowerTier = Nothing
    Set mp_oTierFormat = Nothing
    Set mp_oTierAppearance = Nothing
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Object Member Variables
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get UpperTier() As clsTier
    Set UpperTier = mp_oUpperTier
End Property

Public Property Get MiddleTier() As clsTier
    Set MiddleTier = mp_oMiddleTier
End Property

Public Property Get LowerTier() As clsTier
    Set LowerTier = mp_oLowerTier
End Property

Public Property Get TierFormat() As clsTierFormat
    Set TierFormat = mp_oTierFormat
End Property

Public Property Get TierAppearance() As clsTierAppearance
    Set TierAppearance = mp_oTierAppearance
End Property

Friend Property Get TimeLine() As clsTimeLine
    Set TimeLine = mp_oTimeLine
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TierArea"
    oXML.InitializeWriter
    oXML.WriteObject mp_oUpperTier.GetXML
    oXML.WriteObject mp_oMiddleTier.GetXML
    oXML.WriteObject mp_oLowerTier.GetXML
    oXML.WriteObject mp_oTierFormat.GetXML
    oXML.WriteObject mp_oTierAppearance.GetXML
    
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TierArea"
    oXML.SetXML sXML
    oXML.InitializeReader
    mp_oUpperTier.SetXML oXML.ReadObject("UpperTier")
    mp_oMiddleTier.SetXML oXML.ReadObject("MiddleTier")
    mp_oLowerTier.SetXML oXML.ReadObject("LowerTier")
    mp_oTierFormat.SetXML oXML.ReadObject("TierFormat")
    mp_oTierAppearance.SetXML oXML.ReadObject("TierAppearance")
End Sub
