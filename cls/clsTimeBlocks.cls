VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimeBlocks"
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


'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oCollection = New clsCollectionBase
    mp_oCollection.Initialize Value, "TimeBlock"
    
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

Public Property Get Item(ByVal Index As String) As clsTimeBlock
    Set Item = mp_oCollection.m_oItem(Index, 50215, 50216, 50217, 50218)
End Property

Friend Property Get oCollection() As clsCollectionBase
    Set oCollection = mp_oCollection
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub Add(ByVal StartDate As Date, ByVal EndDate As Date, Optional ByVal Key As String = "", Optional ByVal StyleIndex As String = "0")
    mp_oCollection.AddMode = True
    Dim oTimeBlock As New clsTimeBlock
    oTimeBlock.Initialize mp_oControl
    Key = mp_oControl.StrLib.StrTrim(Key)
    oTimeBlock.StartDate = StartDate
    oTimeBlock.EndDate = EndDate
    oTimeBlock.Key = Key
    oTimeBlock.StyleIndex = StyleIndex
    mp_oCollection.m_Add oTimeBlock, Key, 50213, 50214
    Set oTimeBlock = Nothing
End Sub

Public Sub Clear()
    mp_oCollection.m_Clear
End Sub

Public Sub Remove(ByVal Index As String)
    mp_oCollection.m_Remove Index, 50219, 50220, 50221, 50222
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Position
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Position()
    Dim oTimeBlock As clsTimeBlock
    Dim lIndex As Long
    For lIndex = 1 To Count
        Set oTimeBlock = mp_oCollection.m_oReturnArrayElement(lIndex)
        If (((oTimeBlock.StartDate >= mp_oControl.CurrentViewObject.TimeLine.StartDate And oTimeBlock.StartDate <= mp_oControl.CurrentViewObject.TimeLine.EndDate) Or (oTimeBlock.EndDate >= mp_oControl.CurrentViewObject.TimeLine.StartDate And oTimeBlock.EndDate <= mp_oControl.CurrentViewObject.TimeLine.EndDate)) Or (oTimeBlock.StartDate < mp_oControl.CurrentViewObject.TimeLine.StartDate And oTimeBlock.EndDate > mp_oControl.CurrentViewObject.TimeLine.EndDate)) Then
            oTimeBlock.f_bVisible = True
            oTimeBlock.f_lTop = mp_oControl.CurrentViewObject.ClientArea.Top
            If mp_oControl.TimeBlockBehaviour = E_TIMEBLOCKBEHAVIOUR.TBB_CONTROLEXTENTS Then
                oTimeBlock.f_lBottom = mp_oControl.mt_BottomMargin
            Else
                oTimeBlock.f_lBottom = mp_oControl.Rows.TopOffset
            End If
            oTimeBlock.f_lLeft = mp_oControl.MathLib.GetXCoordinateFromDate(oTimeBlock.StartDate)
            If oTimeBlock.Left < -100 Then
                oTimeBlock.f_lLeft = -100
            End If
            oTimeBlock.f_lRight = mp_oControl.MathLib.GetXCoordinateFromDate(oTimeBlock.EndDate)
            If oTimeBlock.Right > (mp_oControl.mt_RightMargin + 100) Then
                oTimeBlock.f_lRight = (mp_oControl.mt_RightMargin + 100)
            End If
            If (oTimeBlock.Right - oTimeBlock.Left) < 0 Then
                oTimeBlock.f_bVisible = False
            End If
        Else
            oTimeBlock.f_bVisible = False
        End If
    Next lIndex
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Draw
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Draw()
    Dim lIndex As Long
    Dim oTimeBlock As clsTimeBlock
    Dim bDraw As Boolean
    If Count = 0 Then
        Exit Sub
    End If
    For lIndex = 1 To Count
        Set oTimeBlock = mp_oCollection.m_oReturnArrayElement(lIndex)
        If oTimeBlock.Visible = True Then
            If mp_oControl.GrphLib.RectangleIntersection(oTimeBlock.LeftTrim, oTimeBlock.Top, oTimeBlock.RightTrim, oTimeBlock.Bottom) Then
                mp_oControl.GrphLib.ClipRegion oTimeBlock.LeftTrim, oTimeBlock.Top, oTimeBlock.RightTrim, oTimeBlock.Bottom, True
                bDraw = False
                mp_oControl.FireDraw E_EVENTTARGET.EVT_TIMEBLOCK, bDraw, lIndex, 0, mp_oControl.GrphLib.mp_lHdc
                If bDraw = False Then
                    mp_oControl.mp_DrawItem oTimeBlock.Left, oTimeBlock.Right, oTimeBlock.Top, oTimeBlock.Bottom, oTimeBlock.StyleIndex, "", False, Nothing, oTimeBlock.LeftTrim, oTimeBlock.RightTrim, Nothing
                End If
            End If
        End If
    Next lIndex
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim lIndex As Long
    Dim oTimeBlock As clsTimeBlock
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TimeBlocks"
    oXML.InitializeWriter
    For lIndex = 1 To Count
        Set oTimeBlock = mp_oCollection.m_oReturnArrayElement(lIndex)
        oXML.WriteObject oTimeBlock.GetXML
    Next lIndex
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim lIndex As Long
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "TimeBlocks"
    oXML.SetXML sXML
    oXML.InitializeReader
    mp_oCollection.m_Clear
    For lIndex = 1 To oXML.ReadCollectionCount
        Dim oTimeBlock As New clsTimeBlock
        oTimeBlock.Initialize mp_oControl
        oTimeBlock.SetXML oXML.ReadCollectionObject(lIndex)
        mp_oCollection.AddMode = True
        mp_oCollection.m_Add oTimeBlock, oTimeBlock.Key, 50000, 50000
        Set oTimeBlock = Nothing
    Next lIndex
End Sub

