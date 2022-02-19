VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDefaultValues"
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

'// Parent Control Pointer
Private mp_oControl As ActiveGanttVBCtl
'// Object member variables
Private mp_oUpperTier As clsTier
Private mp_oMiddleTier As clsTier
Private mp_oLowerTier As clsTier
Private mp_oTickMarkArea As clsTickMarkArea
Private mp_oTierFormat As clsTierFormat
Private mp_oTierAppearance As clsTierAppearance
Private mp_lRowHeight As Long
Private mp_lColumnWidth As Long


'// ---------------------------------------------------------------------------------------------------------------------
'// Initialization & Destruction
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oUpperTier = New clsTier
    mp_oUpperTier.Initialize mp_oControl, Nothing, E_TIERPOSITION.SP_UPPER, False
    Set mp_oLowerTier = New clsTier
    mp_oLowerTier.Initialize mp_oControl, Nothing, E_TIERPOSITION.SP_LOWER, False
    Set mp_oMiddleTier = New clsTier
    mp_oMiddleTier.Initialize mp_oControl, Nothing, E_TIERPOSITION.SP_MIDDLE, False
    Set mp_oTickMarkArea = New clsTickMarkArea
    mp_oTickMarkArea.Initialize mp_oControl, Nothing, False
    Set mp_oTierFormat = New clsTierFormat
    mp_oTierFormat.Initialize mp_oControl, False
    Set mp_oTierAppearance = New clsTierAppearance
    mp_oTierAppearance.Initialize mp_oControl, False
    mp_oUpperTier.Visible = True
    mp_oUpperTier.TierType = E_TIERTYPE.ST_MONTH
    mp_oMiddleTier.Visible = False
    mp_oMiddleTier.TierType = E_TIERTYPE.ST_WEEK
    mp_oLowerTier.Visible = True
    mp_oUpperTier.TierType = E_TIERTYPE.ST_DAYOFWEEK
    mp_lRowHeight = 40
    mp_lColumnWidth = 125
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

Public Property Get TickMarkArea() As clsTickMarkArea
    Set TickMarkArea = mp_oTickMarkArea
End Property

Public Property Get TierFormat() As clsTierFormat
    Set TierFormat = mp_oTierFormat
End Property

Public Property Get TierAppearance() As clsTierAppearance
    Set TierAppearance = mp_oTierAppearance
End Property

Public Property Get RowHeight() As Long
    RowHeight = mp_lRowHeight
End Property

Public Property Let RowHeight(ByVal Value As Long)
    mp_lRowHeight = Value
End Property

Public Property Let ColumnWidth(ByVal Value As Long)
    mp_lColumnWidth = Value
End Property

Public Property Get ColumnWidth() As Long
    ColumnWidth = mp_lColumnWidth
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "DefaultValues"
    oXML.InitializeWriter
    oXML.WriteProperty "RowHeight", mp_lRowHeight
    oXML.WriteProperty "ColumnWidth", mp_lColumnWidth
    oXML.WriteObject mp_oUpperTier.GetXML
    oXML.WriteObject mp_oMiddleTier.GetXML
    oXML.WriteObject mp_oLowerTier.GetXML
    oXML.WriteObject mp_oTickMarkArea.GetXML
    oXML.WriteObject mp_oTierFormat.GetXML
    oXML.WriteObject mp_oTierAppearance.GetXML
    
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "DefaultValues"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadProperty "RowHeight", mp_lRowHeight
    oXML.ReadProperty "ColumnWidth", mp_lColumnWidth
    mp_oUpperTier.SetXML oXML.ReadObject("UpperTier")
    mp_oMiddleTier.SetXML oXML.ReadObject("MiddleTier")
    mp_oLowerTier.SetXML oXML.ReadObject("LowerTier")
    mp_oTickMarkArea.SetXML oXML.ReadObject("TickMarkArea")
    mp_oTierFormat.SetXML oXML.ReadObject("TierFormat")
    mp_oTierAppearance.SetXML oXML.ReadObject("TierAppearance")
End Sub



