VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTierColor"
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
Private mp_clsTierColors As clsTierColors
Private mp_clrBackColor As OLE_COLOR
Private mp_clrForeColor As OLE_COLOR




'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl, ByRef oTierColors As clsTierColors)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// clsItemBase
    mp_sKey = ""
    mp_lIndex = 0
    '// Object Member Variables
    Set mp_clsTierColors = oTierColors
End Sub

Private Sub Class_Terminate()
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// clsItemBase
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Key() As String
    Key = mp_sKey
End Property

Public Property Let Key(ByVal Value As String)
    mp_clsTierColors.oCollection.mp_SetKey mp_sKey, Value, 50000
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

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mp_clrForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
    mp_clrForeColor = mp_oControl.GrphLib.ConvertColor(Value)
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mp_clrBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
    mp_clrBackColor = mp_oControl.GrphLib.ConvertColor(Value)
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TierColor"
    oXML.InitializeWriter
    oXML.WriteProperty "ForeColor", mp_clrForeColor
    oXML.WriteProperty "BackColor", mp_clrBackColor
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TierColor"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadPropertyColor "ForeColor", mp_clrForeColor
    oXML.ReadPropertyColor "BackColor", mp_clrBackColor
End Sub
