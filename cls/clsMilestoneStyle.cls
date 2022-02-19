VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMilestoneStyle"
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
Private mp_clrBorderColor As OLE_COLOR
Private mp_clrFillColor As OLE_COLOR
Private mp_yShapeIndex As GRE_FIGURETYPE

'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    mp_clrBorderColor = GRE_COLORS.CLR_BLACK
    mp_clrFillColor = GRE_COLORS.CLR_BLACK
    mp_yShapeIndex = GRE_FIGURETYPE.FT_NONE
End Sub

Private Sub Class_Terminate()
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Object Member Variables
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = mp_clrBorderColor
End Property

Public Property Let BorderColor(ByVal Value As OLE_COLOR)
    mp_clrBorderColor = mp_oControl.GrphLib.ConvertColor(Value)
End Property

Public Property Get FillColor() As OLE_COLOR
    FillColor = mp_clrFillColor
End Property

Public Property Let FillColor(ByVal Value As OLE_COLOR)
    mp_clrFillColor = mp_oControl.GrphLib.ConvertColor(Value)
End Property

Public Property Get ShapeIndex() As GRE_FIGURETYPE
    ShapeIndex = mp_yShapeIndex
End Property

Public Property Let ShapeIndex(ByVal Value As GRE_FIGURETYPE)
    mp_yShapeIndex = Value
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "MilestoneStyle"
    oXML.InitializeWriter
    oXML.WriteProperty "BorderColor", mp_clrBorderColor
    oXML.WriteProperty "FillColor", mp_clrFillColor
    oXML.WriteProperty "ShapeIndex", mp_yShapeIndex
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "MilestoneStyle"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadPropertyColor "BorderColor", mp_clrBorderColor
    oXML.ReadPropertyColor "FillColor", mp_clrFillColor
    oXML.ReadProperty "ShapeIndex", mp_yShapeIndex
End Sub

