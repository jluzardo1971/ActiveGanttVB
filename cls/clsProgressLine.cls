VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProgressLine"
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
Private mp_clrForeColor As OLE_COLOR
Private mp_dtPosition As Date
Private mp_yLength As E_PROGRESSLINELENGTH
Private mp_yLineType As E_PROGRESSLINETYPE
Private mp_oTimeLine As clsTimeLine


'// ---------------------------------------------------------------------------------------------------------------------
'// Initialization & Destruction
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl, ByRef oTimeLine As clsTimeLine)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    mp_clrForeColor = GRE_COLORS.CLR_RED
    mp_dtPosition = mp_oControl.MathLib.CurrentDateTime
    mp_yLength = E_PROGRESSLINELENGTH.TLMA_TICKMARKAREA
    mp_yLineType = E_PROGRESSLINETYPE.TLMT_SYSTEMTIME
    Set mp_oTimeLine = oTimeLine
End Sub

Private Sub Class_Terminate()
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Object Member Variables
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Position() As Date
    Position = mp_dtPosition
End Property

Public Property Let Position(ByVal Value As Date)
    mp_dtPosition = Value
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mp_clrForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
    mp_clrForeColor = mp_oControl.GrphLib.ConvertColor(Value)
End Property

Public Property Get Length() As E_PROGRESSLINELENGTH
    Length = mp_yLength
End Property

Public Property Let Length(ByVal Value As E_PROGRESSLINELENGTH)
    mp_yLength = Value
End Property

Public Property Get LineType() As E_PROGRESSLINETYPE
    LineType = mp_yLineType
End Property

Public Property Let LineType(ByVal Value As E_PROGRESSLINETYPE)
    mp_yLineType = Value
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Draw Method
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Draw()
    Dim lXCoordinate As Long
    Dim yTimeLineMarkerLength As E_PROGRESSLINELENGTH
    Dim dtDate As Date
    If mp_yLineType = E_PROGRESSLINETYPE.TLMT_SYSTEMTIME Then
        dtDate = mp_oControl.MathLib.CurrentDateTime
    Else
        dtDate = mp_dtPosition
    End If
    If dtDate >= mp_oTimeLine.StartDate And dtDate <= mp_oTimeLine.EndDate Then
        yTimeLineMarkerLength = mp_yLength
        lXCoordinate = mp_oControl.MathLib.GetXCoordinateFromDate(mp_dtPosition)
        If mp_oTimeLine.TickMarkArea.Visible = False And yTimeLineMarkerLength = E_PROGRESSLINELENGTH.TLMA_BOTH Then
            yTimeLineMarkerLength = E_PROGRESSLINELENGTH.TLMA_CLIENTAREA
        End If
        If mp_oTimeLine.TickMarkArea.Visible = False And yTimeLineMarkerLength = E_PROGRESSLINELENGTH.TLMA_TICKMARKAREA Then
            yTimeLineMarkerLength = E_PROGRESSLINELENGTH.TLMA_NONE
        End If
        Select Case yTimeLineMarkerLength
            Case E_PROGRESSLINELENGTH.TLMA_TICKMARKAREA
                mp_oControl.GrphLib.DrawLine lXCoordinate, mp_oTimeLine.TiersTickMarksPosition("TickMarkArea"), lXCoordinate, mp_oTimeLine.Bottom, GRE_LINETYPE.LT_NORMAL, mp_clrForeColor, GRE_LINEDRAWSTYLE.LDS_SOLID
            Case E_PROGRESSLINELENGTH.TLMA_CLIENTAREA
                mp_oControl.GrphLib.DrawLine lXCoordinate, mp_oControl.CurrentViewObject.ClientArea.Top, lXCoordinate, mp_oControl.CurrentViewObject.ClientArea.Bottom, GRE_LINETYPE.LT_NORMAL, mp_clrForeColor, GRE_LINEDRAWSTYLE.LDS_SOLID
            Case E_PROGRESSLINELENGTH.TLMA_BOTH
                mp_oControl.GrphLib.DrawLine lXCoordinate, mp_oTimeLine.TiersTickMarksPosition("TickMarkArea"), lXCoordinate, mp_oControl.CurrentViewObject.ClientArea.Bottom, GRE_LINETYPE.LT_NORMAL, mp_clrForeColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        End Select
    End If
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "ProgressLine"
    oXML.InitializeWriter
    oXML.WriteProperty "ForeColor", mp_clrForeColor
    oXML.WriteProperty "Position", mp_dtPosition
    oXML.WriteProperty "Length", mp_yLength
    oXML.WriteProperty "LineType", mp_yLineType
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "ProgressLine"
    oXML.SetXML sXML
    oXML.InitializeReader
    oXML.ReadPropertyColor "ForeColor", mp_clrForeColor
    oXML.ReadProperty "Position", mp_dtPosition
    oXML.ReadProperty "Length", mp_yLength
    oXML.ReadProperty "LineType", mp_yLineType
End Sub




