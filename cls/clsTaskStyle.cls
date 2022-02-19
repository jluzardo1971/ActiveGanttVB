VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTaskStyle"
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
Private mp_clrEndBorderColor As OLE_COLOR
Private mp_clrEndFillColor As OLE_COLOR
Private mp_clrStartBorderColor As OLE_COLOR
Private mp_clrStartFillColor As OLE_COLOR


Private mp_yEndShapeIndex As GRE_FIGURETYPE
Private mp_yStartShapeIndex As GRE_FIGURETYPE

Private mp_oEndPicture As StdPicture
Private mp_oMiddlePicture As StdPicture
Private mp_oStartPicture As StdPicture

'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    mp_clrEndBorderColor = GRE_COLORS.CLR_BLACK
    mp_clrEndFillColor = GRE_COLORS.CLR_BLACK
    mp_clrStartBorderColor = GRE_COLORS.CLR_BLACK
    mp_clrStartFillColor = GRE_COLORS.CLR_BLACK
    Set mp_oEndPicture = Nothing
    Set mp_oMiddlePicture = Nothing
    Set mp_oStartPicture = Nothing
    mp_yEndShapeIndex = GRE_FIGURETYPE.FT_NONE
    mp_yStartShapeIndex = GRE_FIGURETYPE.FT_NONE
    
End Sub

Private Sub Class_Terminate()
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Object Member Variables
'// ---------------------------------------------------------------------------------------------------------------------


Public Property Get EndBorderColor() As OLE_COLOR
    EndBorderColor = mp_clrEndBorderColor
End Property

Public Property Let EndBorderColor(ByVal Value As OLE_COLOR)
    mp_clrEndBorderColor = mp_oControl.GrphLib.ConvertColor(Value)
End Property
    
Public Property Get EndFillColor() As OLE_COLOR
    EndFillColor = mp_clrEndFillColor
End Property

Public Property Let EndFillColor(ByVal Value As OLE_COLOR)
    mp_clrEndFillColor = mp_oControl.GrphLib.ConvertColor(Value)
End Property

Public Property Get StartBorderColor() As OLE_COLOR
    StartBorderColor = mp_clrStartBorderColor
End Property

Public Property Let StartBorderColor(ByVal Value As OLE_COLOR)
    mp_clrStartBorderColor = mp_oControl.GrphLib.ConvertColor(Value)
End Property

Public Property Get StartFillColor() As OLE_COLOR
    StartFillColor = mp_clrStartFillColor
End Property

Public Property Let StartFillColor(ByVal Value As OLE_COLOR)
    mp_clrStartFillColor = mp_oControl.GrphLib.ConvertColor(Value)
End Property

Public Property Get EndPicture() As StdPicture
    Set EndPicture = mp_oEndPicture
End Property

Public Property Set EndPicture(ByVal Value As StdPicture)
    Set mp_oEndPicture = Value
End Property

Public Property Get MiddlePicture() As StdPicture
    Set MiddlePicture = mp_oMiddlePicture
End Property

Public Property Set MiddlePicture(ByVal Value As StdPicture)
    Set mp_oMiddlePicture = Value
End Property

Public Property Get StartPicture() As StdPicture
    Set StartPicture = mp_oStartPicture
End Property

Public Property Set StartPicture(ByVal Value As StdPicture)
    Set mp_oStartPicture = Value
End Property

Public Property Get StartShapeIndex() As GRE_FIGURETYPE
    StartShapeIndex = mp_yStartShapeIndex
End Property

Public Property Let StartShapeIndex(ByVal Value As GRE_FIGURETYPE)
    mp_yStartShapeIndex = Value
End Property

Public Property Get EndShapeIndex() As GRE_FIGURETYPE
    EndShapeIndex = mp_yEndShapeIndex
End Property

Public Property Let EndShapeIndex(ByVal Value As GRE_FIGURETYPE)
    mp_yEndShapeIndex = Value
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TaskStyle"
    oXML.InitializeWriter
    oXML.WriteProperty "EndBorderColor", mp_clrEndBorderColor
    oXML.WriteProperty "EndFillColor", mp_clrEndFillColor
    oXML.WriteProperty "StartBorderColor", mp_clrStartBorderColor
    oXML.WriteProperty "StartFillColor", mp_clrStartFillColor
    oXML.WritePropertyPicture "EndPicture", mp_oEndPicture
    oXML.WritePropertyPicture "MiddlePicture", mp_oMiddlePicture
    oXML.WritePropertyPicture "StartPicture", mp_oStartPicture
    oXML.WriteProperty "EndShapeIndex", mp_yEndShapeIndex
    oXML.WriteProperty "StartShapeIndex", mp_yStartShapeIndex
    
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TaskStyle"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadPropertyColor "EndBorderColor", mp_clrEndBorderColor
    oXML.ReadPropertyColor "EndFillColor", mp_clrEndFillColor
    oXML.ReadPropertyColor "StartBorderColor", mp_clrStartBorderColor
    oXML.ReadPropertyColor "StartFillColor", mp_clrStartFillColor
    oXML.ReadPropertyPicture "EndPicture", mp_oEndPicture
    oXML.ReadPropertyPicture "MiddlePicture", mp_oMiddlePicture
    oXML.ReadPropertyPicture "StartPicture", mp_oStartPicture
    oXML.ReadProperty "EndShapeIndex", mp_yEndShapeIndex
    oXML.ReadProperty "StartShapeIndex", mp_yStartShapeIndex

End Sub

