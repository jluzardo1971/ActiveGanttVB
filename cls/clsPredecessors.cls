VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPredecessors"
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
Private mp_oTask As clsTask
Private mp_oMilestone As clsMilestone

Private Type T_PREDECESSOR
    udtObjectType As E_OBJECTTYPE
    lIndex As Long
    lLeft As Long
    lTop As Long
    lRight As Long
    lBottom As Long
    bVisible As Boolean
End Type


'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl, ByRef oTask As clsTask, ByRef oMilestone As clsMilestone)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set mp_oCollection = New clsCollectionBase
    mp_oCollection.Initialize Value, "Predecessor"
    Set mp_oTask = oTask
    Set mp_oMilestone = oMilestone
    
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

Public Property Get Item(ByVal Index As String) As clsPredecessor
    Set Item = mp_oCollection.m_oItem(Index, 50349, 50350, 50351, 50352)
End Property

Friend Property Get oCollection() As clsCollectionBase
    Set oCollection = mp_oCollection
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Collection Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub Add(ByVal PredecessorKey As String, Optional ByVal PredecessorType As E_OBJECTTYPE = E_OBJECTTYPE.OT_TASK, Optional ByVal Key As String = "", Optional ByVal StyleIndex As String = "0")
    mp_oCollection.AddMode = True
    Dim oPredecessor As New clsPredecessor
    oPredecessor.Initialize mp_oControl, Me
    oPredecessor.PredecessorType = PredecessorType
    oPredecessor.PredecessorKey = PredecessorKey
    oPredecessor.StyleIndex = StyleIndex
    oPredecessor.Key = Key
    mp_oCollection.m_Add oPredecessor, Key, 50337, 50338
    Set oPredecessor = Nothing
End Sub

Public Sub Clear()
    mp_oCollection.m_Clear
End Sub

Public Sub Remove(ByVal Index As String)
    mp_oCollection.m_Remove Index, 50369, 50370, 50371, 50372
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Draw()
    Dim lIndex2 As Long
    Dim oMain As T_PREDECESSOR
    Dim oLink As T_PREDECESSOR
    Dim oTask As clsTask
    Dim oMilestone As clsMilestone
    Dim oPredecessor As clsPredecessor
    Dim oStyle As clsStyle
    Dim oPredecessorStyle As clsPredecessorStyle
    Dim bCustomDraw As Boolean
    mp_oControl.GrphLib.ClipRegion mp_oControl.Splitter.Right, mp_oControl.CurrentViewObject.ClientArea.Top, mp_oControl.mt_RightMargin, mp_oControl.CurrentViewObject.ClientArea.Bottom, True
    If Not mp_oTask Is Nothing Then
        oMain.udtObjectType = E_OBJECTTYPE.OT_TASK
        oMain.lIndex = mp_oTask.Index
        oMain.lLeft = mp_oTask.Left
        oMain.lTop = mp_oTask.Top
        oMain.lRight = mp_oTask.Right
        oMain.lBottom = mp_oTask.Bottom
        If mp_oTask.Top = mp_oTask.Bottom Then
            oMain.bVisible = False
        Else
            oMain.bVisible = True
        End If
    Else
        oMain.udtObjectType = E_OBJECTTYPE.OT_MILESTONE
        oMain.lIndex = mp_oMilestone.Index
        oMain.lLeft = mp_oMilestone.Left
        oMain.lTop = mp_oMilestone.Top
        oMain.lRight = mp_oMilestone.Right
        oMain.lBottom = mp_oMilestone.Bottom
        If mp_oMilestone.Top = mp_oMilestone.Bottom Then
            oMain.bVisible = False
        Else
            oMain.bVisible = True
        End If
    End If
    If Count > 0 Then
        For lIndex2 = 1 To Count
            Set oPredecessor = mp_oCollection.m_oReturnArrayElement(lIndex2)
            If oPredecessor.Visible = True And oMain.bVisible = True Then
                Set oStyle = mp_oControl.Styles.FItem(oPredecessor.StyleIndex)
                If oPredecessor.PredecessorType = E_OBJECTTYPE.OT_TASK Then
                    Set oTask = mp_oControl.Tasks.oCollection.m_oReturnArrayElementKey(oPredecessor.PredecessorKey)
                    oLink.udtObjectType = E_OBJECTTYPE.OT_TASK
                    oLink.lIndex = oTask.Index
                    oLink.lLeft = oTask.Left
                    oLink.lTop = oTask.Top
                    oLink.lRight = oTask.Right
                    oLink.lBottom = oTask.Bottom
                    If oTask.Top = oTask.Bottom Then
                        oLink.bVisible = False
                    Else
                        oLink.bVisible = True
                    End If
                Else
                    Set oMilestone = mp_oControl.Milestones.oCollection.m_oReturnArrayElementKey(oPredecessor.PredecessorKey)
                    oLink.udtObjectType = E_OBJECTTYPE.OT_MILESTONE
                    oLink.lIndex = oMilestone.Index
                    oLink.lLeft = oMilestone.Left
                    oLink.lTop = oMilestone.Top
                    oLink.lRight = oMilestone.Right
                    oLink.lBottom = oMilestone.Bottom
                    If oMilestone.Top = oMilestone.Bottom Then
                        oLink.bVisible = False
                    Else
                        oLink.bVisible = True
                    End If
                End If
                If oLink.bVisible = True Then
                    bCustomDraw = False
                    mp_oControl.FirePredecessorDraw bCustomDraw, oMain.lIndex, oMain.udtObjectType, oLink.lIndex, oLink.udtObjectType, mp_oControl.GrphLib.mp_lHdc
                    If bCustomDraw = False Then
                        Set oPredecessorStyle = oStyle.PredecessorStyle
                        mp_oControl.GrphLib.DrawConnectingLines oLink.lLeft, oLink.lTop, oLink.lRight, oLink.lBottom, oMain.lLeft, oMain.lTop, oMain.lRight, oMain.lBottom, oPredecessorStyle.Style, oPredecessorStyle.ArrowHeads, oPredecessorStyle.ArrowSize, oStyle.ForeColor, oPredecessorStyle.LineWidth, oPredecessorStyle.LineStyle
                    End If
                End If
            End If
        Next lIndex2
    End If

End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// XML
'// ---------------------------------------------------------------------------------------------------------------------

Public Function GetXML() As String
    Dim lIndex As Long
    Dim oPredecessor As clsPredecessor
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Predecessors"
    oXML.InitializeWriter
    For lIndex = 1 To Count
        Set oPredecessor = mp_oCollection.m_oReturnArrayElement(lIndex)
        oXML.WriteObject oPredecessor.GetXML
    Next lIndex
    GetXML = oXML.GetXML
End Function

Public Sub SetXML(ByVal sXML As String)
    Dim lIndex As Long
    Dim oXML As New clsXML
    oXML.Initialize mp_oControl, "Predecessors"
    oXML.SetXML sXML
    oXML.InitializeReader
    mp_oCollection.m_Clear
    For lIndex = 1 To oXML.ReadCollectionCount
        Dim oPredecessor As New clsPredecessor
        oPredecessor.Initialize mp_oControl, Me
        oPredecessor.SetXML oXML.ReadCollectionObject(lIndex)
        mp_oCollection.AddMode = True
        mp_oCollection.m_Add oPredecessor, oPredecessor.Key, 50000, 50000
        Set oPredecessor = Nothing
    Next lIndex
End Sub
