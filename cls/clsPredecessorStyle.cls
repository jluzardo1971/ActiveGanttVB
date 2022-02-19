VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPredecessorStyle"
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
Private mp_yStyle As GRE_CONNLINESTYLE
Private mp_yArrowHeads As GRE_ARROWHEADS
Private mp_yArrowSize As Integer
Private mp_yLineStyle As GRE_LINEDRAWSTYLE
Private mp_iLineWidth As Integer

'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    mp_yStyle = GRE_CONNLINESTYLE.PDS_NORMAL
    mp_yArrowHeads = GRE_ARROWHEADS.AH_END
    mp_yArrowSize = 2
    mp_yLineStyle = GRE_LINEDRAWSTYLE.LDS_SOLID
    mp_iLineWidth = 1
    
End Sub

Private Sub Class_Terminate()
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Object Member Variables
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Style() As GRE_CONNLINESTYLE
    Style = mp_yStyle
End Property

Public Property Let Style(ByVal Value As GRE_CONNLINESTYLE)
    mp_yStyle = Value
End Property

Public Property Get LineWidth() As Integer
    LineWidth = mp_iLineWidth
End Property

Public Property Let LineWidth(ByVal Value As Integer)
    mp_iLineWidth = Value
End Property

Public Property Get LineStyle() As GRE_LINEDRAWSTYLE
    LineStyle = mp_yLineStyle
End Property

Public Property Let LineStyle(ByVal Value As GRE_LINEDRAWSTYLE)
    mp_yLineStyle = Value
End Property

Public Property Get ArrowHeads() As GRE_ARROWHEADS
    ArrowHeads = mp_yArrowHeads
End Property

Public Property Let ArrowHeads(ByVal Value As GRE_ARROWHEADS)
    mp_yArrowHeads = Value
End Property

Public Property Get ArrowSize() As Integer
    ArrowSize = mp_yArrowSize
End Property

Public Property Let ArrowSize(ByVal Value As Integer)
    If Value < 1 Then Value = 1
    mp_yArrowSize = Value
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "PredecessorStyle"
    oXML.InitializeWriter
    oXML.WriteProperty "ArrowHeads", mp_yArrowHeads
    oXML.WriteProperty "ArrowSize", mp_yArrowSize
    oXML.WriteProperty "LineStyle", mp_yLineStyle
    oXML.WriteProperty "Style", mp_yStyle
    oXML.WriteProperty "LineWidth", mp_iLineWidth
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "PredecessorStyle"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadProperty "ArrowHeads", mp_yArrowHeads
    oXML.ReadProperty "ArrowSize", mp_yArrowSize
    oXML.ReadProperty "LineStyle", mp_yLineStyle
    oXML.ReadProperty "Style", mp_yStyle
    oXML.ReadProperty "LineWidth", mp_iLineWidth
End Sub
