VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColumns"
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
'// Initialization & Destruction
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oCollection = New clsCollectionBase
    mp_oCollection.Initialize Value, "Column"
    
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

Public Property Get Item(ByVal Index As String) As clsColumn
    Set Item = mp_oCollection.m_oItem(Index, 50000, 50000, 50000, 50000)
End Property

Friend Property Get oCollection() As clsCollectionBase
    Set oCollection = mp_oCollection
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub Add(ByVal Caption As String, Optional ByVal Key As String = "", Optional ByVal Width As Long = 125, Optional ByVal StyleIndex As String = "0")
    mp_oCollection.AddMode = True
    Dim oColumn As New clsColumn
    Caption = mp_oControl.StrLib.StrTrim(Caption)
    oColumn.Initialize mp_oControl
    oColumn.Caption = Caption
    oColumn.Width = Width
    oColumn.StyleIndex = StyleIndex
    oColumn.Key = Key
    Dim lIndex As Long
    Dim oRow As clsRow
    For lIndex = 1 To mp_oControl.Rows.Count
        Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex)
        oRow.Cells.Add
    Next lIndex
    mp_oCollection.m_Add oColumn, oColumn.Key, 50384, 50385
    mp_oControl.Splitter.f_AdjustPosition
End Sub

Public Sub Clear()
    mp_oControl.Rows.ClearCells
    mp_oControl.SelectedColumnIndex = 0
    mp_oCollection.m_Clear
    mp_oControl.Splitter.f_AdjustPosition
End Sub

Public Sub Remove(ByVal Index As String)
    Dim lIndex As Long
    Dim oRow As clsRow
    For lIndex = 1 To mp_oControl.Rows.Count
        Set oRow = mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex)
        oRow.Cells.Remove Index
    Next lIndex
    mp_oControl.SelectedColumnIndex = 0
    mp_oCollection.m_Remove Index, 50377, 50378, 50379, 50380
    mp_oControl.Splitter.f_AdjustPosition
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Custom Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub MoveColumn(ByVal OriginColumnIndex As Long, ByVal DestColumnIndex As Long)
    If OriginColumnIndex < 1 Or OriginColumnIndex > Count Then
        Exit Sub
    End If
    If DestColumnIndex < 1 Or DestColumnIndex > Count Then
        Exit Sub
    End If
    If DestColumnIndex = OriginColumnIndex Then
        Exit Sub
    End If
    mp_oCollection.m_lCopyAndMoveItems OriginColumnIndex, DestColumnIndex
End Sub

Friend Property Get Width() As Long
    Dim lIndex As Long
    Dim lResult As Long
    For lIndex = 1 To Count
        lResult = lResult + mp_oCollection.m_oItem(lIndex).Width
    Next lIndex
    Width = lResult
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Position
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Position()
    Dim lIndex As Long
    Dim oColumn As clsColumn
    Dim lLeft As Long
    lLeft = -mp_oControl.CurrentViewObject.TimeLine.ScrollBar.Value + mp_oControl.mt_LeftMargin
    For lIndex = 1 To Count
        Set oColumn = mp_oCollection.m_oReturnArrayElement(lIndex)
        oColumn.f_lLeft = lLeft
        oColumn.f_lRight = lLeft + oColumn.Width
        If oColumn.Right < mp_oControl.mt_LeftMargin Then
            oColumn.f_bVisible = False
        ElseIf oColumn.Left > mp_oControl.Splitter.Left Then
            oColumn.f_bVisible = False
        Else
            oColumn.f_bVisible = True
        End If
        If oColumn.Right > oColumn.Left Then
            oColumn.f_bVisible = True
        Else
            oColumn.f_bVisible = False
        End If
        lLeft = lLeft + oColumn.Width
    Next lIndex
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Draw
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Draw()
    Dim oColumn As clsColumn
    Dim lIndex As Long
    Dim bDraw As Boolean
    If Count = 0 Then
        Exit Sub
    End If
    If (mp_oControl.CurrentViewObject.TimeLine.Height = 0) Then
        Exit Sub
    End If
    For lIndex = 1 To Count
        Set oColumn = mp_oCollection.m_oReturnArrayElement(lIndex)
        If oColumn.Visible = True Then
            If mp_oControl.GrphLib.RectangleIntersection(oColumn.Left, oColumn.Top, oColumn.Right, oColumn.Bottom) = True Then
                mp_oControl.GrphLib.ClipRegion oColumn.LeftTrim, oColumn.Top, oColumn.RightTrim, oColumn.Bottom, True
                bDraw = False
                mp_oControl.FireDraw E_EVENTTARGET.EVT_COLUMN, bDraw, lIndex, 0, mp_oControl.GrphLib.mp_lHdc
                If bDraw = False Then
                    mp_oControl.mp_DrawItem oColumn.Left, oColumn.Right - 1, oColumn.Top, oColumn.Bottom, oColumn.StyleIndex, oColumn.Caption, (lIndex = mp_oControl.SelectedColumnIndex), oColumn.Picture, oColumn.LeftTrim, oColumn.RightTrim, Nothing
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
    Dim oColumn As clsColumn
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Columns"
    oXML.InitializeWriter
    For lIndex = 1 To Count
        Set oColumn = mp_oCollection.m_oReturnArrayElement(lIndex)
        oXML.WriteObject oColumn.GetXML
    Next lIndex
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim lIndex As Long
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Columns"
    oXML.SetXML sXML
    oXML.InitializeReader
    mp_oCollection.m_Clear
    For lIndex = 1 To oXML.ReadCollectionCount
        Dim oColumn As New clsColumn
        oColumn.Initialize mp_oControl
        oColumn.SetXML oXML.ReadCollectionObject(lIndex)
        mp_oCollection.AddMode = True
        mp_oCollection.m_Add oColumn, oColumn.Key, 50000, 50000
        Set oColumn = Nothing
    Next lIndex
End Sub
