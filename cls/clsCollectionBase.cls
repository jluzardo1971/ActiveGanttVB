VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCollectionBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'// Object member variables
Private m_sObjectName As String
Private mp_aoCollection() As Object
Private mp_bAddMode As Boolean
Private mp_bDescending As Boolean
Private mp_bIgnoreKeyChecks As Boolean
Private mp_bSortCells As Boolean
Private mp_lCellIndex As Long
Private mp_oKeys As Collection
Private mp_sPropertyName As String
Private mp_ySortType As E_SORTTYPE


'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl, ByVal sObjectName As String)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object member variables
    m_sObjectName = sObjectName
    ReDim mp_aoCollection(0)
    mp_bAddMode = False
    mp_bDescending = False
    mp_bIgnoreKeyChecks = False
    mp_bSortCells = False
    mp_lCellIndex = 0
    Set mp_oKeys = New Collection
    mp_sPropertyName = ""
    mp_ySortType = E_SORTTYPE.ES_STRING
End Sub

Private Sub Class_Terminate()
    ReDim mp_aoCollection(0)
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub m_Add(ByRef r_oObject As Object, ByVal v_sKey As String, ByVal v_lErr1 As Long, ByVal v_lErr2 As Long, Optional ByVal v_bKeyRequired As Boolean = False, Optional ByVal v_lKeyError As Long)
    Dim lUpperBounds As Long
    If mp_bAddMode = False Then
        mp_oControl.mp_ErrorReport 50000, "AddMode must be set to true before executing oCollection.m_Add", "cls" & m_sObjectName & "s"
    End If
    If mp_oControl.StrLib.StrIsNumeric(v_sKey) Then
        mp_oControl.mp_ErrorReport v_lErr1, "Invalid " & m_sObjectName & " object key, key cannot be numeric", "ActiveGanttVBCtl.cls" & m_sObjectName & "s.Add"
        Exit Sub
    End If
    If v_sKey <> "" Then
        If m_bIsKeyUnique(v_sKey) = False Then
            mp_oControl.mp_ErrorReport v_lErr2, "Key is not unique in " & m_sObjectName & "s collection", "ActiveGanttVBCtl.cls" & m_sObjectName & "s.Add"
            Exit Sub
        End If
    End If
    lUpperBounds = UBound(mp_aoCollection, 1) + 1
    ReDim Preserve mp_aoCollection(lUpperBounds)
    r_oObject.Index = lUpperBounds
    Set mp_aoCollection(lUpperBounds) = r_oObject
    If v_sKey <> "" Then
        mp_oKeys.Add lUpperBounds, v_sKey
    End If
    mp_bAddMode = False
End Sub

Public Sub m_Clear()
    ReDim mp_aoCollection(0)
    m_ReconstKeys
End Sub

Public Function m_lCount() As Long
    m_lCount = UBound(mp_aoCollection, 1)
End Function

Public Property Get m_oItem(ByVal v_lIndex As String, Optional ByVal v_lErr1 As Long = 50000, Optional ByVal v_lErr2 As Long = 50000, Optional ByVal v_lErr3 As Long = 50000, Optional ByVal v_lErr4 As Long = 50000) As Object
Attribute m_oItem.VB_MemberFlags = "200"
    Dim lIndex As Long
    If Not mp_oControl.StrLib.StrIsNumeric(v_lIndex) Then
        If v_lIndex = "" Then
            mp_oControl.mp_ErrorReport v_lErr1, "Invalid " & m_sObjectName & " object key, key cannot be an empty string", "ActiveGanttVBCtl.cls" & m_sObjectName & "s.Get m_oItem"
            Exit Property
        End If
        lIndex = m_lFindIndexByKey(v_lIndex)
        If lIndex = -1 Then
            mp_oControl.mp_ErrorReport v_lErr2, m_sObjectName & " object not found, invalid key", "ActiveGanttVBCtl.cls" & m_sObjectName & "s.Get m_oItem"
            Exit Property
        End If
    Else
        lIndex = CLng(v_lIndex)
        If CStr(lIndex) <> v_lIndex Then
            mp_oControl.mp_ErrorReport v_lErr3, m_sObjectName & " object not found, invalid v_lIndex", "ActiveGanttVBCtl.cls" & m_sObjectName & "s.Get m_oItem"
            Exit Property
        End If
        If lIndex < 1 Or lIndex > UBound(mp_aoCollection, 1) Then
            mp_oControl.mp_ErrorReport v_lErr4, m_sObjectName & " object not found, v_lIndex out of bounds", "ActiveGanttVBCtl.cls" & m_sObjectName & "s.Get m_oItem"
            Exit Property
        End If
    End If
    Set m_oItem = mp_aoCollection(lIndex)
End Property

Public Sub m_Remove(ByVal Index As String, ByVal v_lErr1 As Long, ByVal v_lErr2 As Long, ByVal v_lErr3 As Long, ByVal v_lErr4 As Long)
    Dim lIndex As Long
    Dim lIteratorIndex As Long
    If Not mp_oControl.StrLib.StrIsNumeric(Index) Then
        If Index = "" Then
            mp_oControl.mp_ErrorReport v_lErr1, "Invalid " & m_sObjectName & " object key, key cannot be an empty string", "ActiveGanttVBCtl.cls" & m_sObjectName & "s.Remove"
            Exit Sub
        End If
        lIndex = m_lFindIndexByKey(Index)
        If lIndex = -1 Then
            mp_oControl.mp_ErrorReport v_lErr2, m_sObjectName & " object not found, invalid key", "ActiveGanttVBCtl.cls" & m_sObjectName & "s.Remove"
            Exit Sub
        End If
    Else
        lIndex = CLng(Index)
        If CStr(lIndex) <> Index Then
            mp_oControl.mp_ErrorReport v_lErr3, m_sObjectName & " object not found, invalid index", "ActiveGanttVBCtl.cls" & m_sObjectName & "s.Remove"
            Exit Sub
        End If
        If lIndex < 1 Or lIndex > UBound(mp_aoCollection, 1) Then
            mp_oControl.mp_ErrorReport v_lErr4, m_sObjectName & " object not found, index out of bounds", "ActiveGanttVBCtl.cls" & m_sObjectName & "s.Remove"
            Exit Sub
        End If
    End If
    mp_bIgnoreKeyChecks = True
    For lIteratorIndex = lIndex To (UBound(mp_aoCollection, 1) - 1)
        Set mp_aoCollection(lIteratorIndex) = mp_aoCollection(lIteratorIndex + 1)
    Next lIteratorIndex
    mp_bIgnoreKeyChecks = False
    ReDim Preserve mp_aoCollection(UBound(mp_aoCollection, 1) - 1)
    m_ReconstKeys
End Sub

Public Sub m_Sort(ByVal sPropertyName As String, ByVal bDescending As Boolean, ByVal SortType As E_SORTTYPE, ByVal StartIndex As Long, ByVal EndIndex As Long)
    Dim aTempArray As Variant
    mp_sPropertyName = sPropertyName
    mp_bDescending = bDescending
    mp_ySortType = SortType
    mp_bSortCells = False
    ReDim aTempArray(StartIndex To EndIndex)
    mp_Sort mp_aoCollection, aTempArray, StartIndex, EndIndex
    m_ReconstKeys
End Sub

Public Sub m_SortCells(ByVal CellIndex As Long, ByVal sPropertyName As String, ByVal bDescending As Boolean, ByVal SortType As E_SORTTYPE, ByVal StartIndex As Long, ByVal EndIndex As Long)
    Dim aTempArray As Variant
    mp_sPropertyName = sPropertyName
    mp_bDescending = bDescending
    mp_ySortType = SortType
    mp_bSortCells = True
    mp_lCellIndex = CellIndex
    ReDim aTempArray(StartIndex To EndIndex)
    mp_Sort mp_aoCollection, aTempArray, StartIndex, EndIndex
    m_ReconstKeys
End Sub

Public Function m_lCopyAndMoveItems(ByVal v_lOriginIndex As Long, ByVal v_lDestinationIndex As Long) As Long
    Dim Buffer As Object
    Dim lIndex As Long
    mp_bIgnoreKeyChecks = True
    Set Buffer = mp_aoCollection(v_lOriginIndex)
    If v_lOriginIndex > v_lDestinationIndex Then
        For lIndex = v_lOriginIndex To v_lDestinationIndex + 1 Step -1
            Set mp_aoCollection(lIndex) = mp_aoCollection(lIndex - 1)
        Next lIndex
    Else
        For lIndex = v_lOriginIndex To v_lDestinationIndex - 1
            Set mp_aoCollection(lIndex) = mp_aoCollection(lIndex + 1)
        Next lIndex
    End If
    Set mp_aoCollection(v_lDestinationIndex) = Buffer
    mp_bIgnoreKeyChecks = False
    m_lCopyAndMoveItems = lIndex
    m_ReconstKeys
End Function

'//******************************************************************************************
'// FRIENDS & AUXILIARY FUNCTIONS
'//******************************************************************************************

Public Function m_oReturnArrayElement(ByVal r_lIndex As Long) As Object
    Set m_oReturnArrayElement = mp_aoCollection(r_lIndex)
End Function

Public Function m_oReturnArrayElementKey(ByVal v_sKey As String) As Object
    Dim lIndex As Long
    If mp_oControl.StrLib.StrIsNumeric(v_sKey) Then
        Set m_oReturnArrayElementKey = mp_aoCollection(CLng(v_sKey))
    Else
        lIndex = m_lFindIndexByKey(v_sKey)
        If lIndex <> -1 Then
            Set m_oReturnArrayElementKey = mp_aoCollection(lIndex)
        Else
            mp_oControl.mp_ErrorReport 50000, "Key not found", "ActiveGanttVBCtl.cls" & m_sObjectName & "s.m_oReturnArrayElementKey"
        End If
    End If
End Function

Public Function m_lReturnIndex(ByVal v_sIndex As String, ByVal bIncludeDefault As Boolean) As Long
    Dim lIndex As Long
    Dim lReturn As Long
    If (mp_oControl.StrLib.StrIsNumeric(v_sIndex)) Then
        lIndex = CLng(v_sIndex)
        If (bIncludeDefault = True) Then
            If (lIndex >= 0 And lIndex <= m_lCount) Then
                lReturn = lIndex
            Else
                lReturn = -1
            End If
        Else
            If (lIndex >= 1 And lIndex <= m_lCount) Then
                lReturn = lIndex
            Else
                lReturn = -1
            End If
        End If
    Else
        lReturn = m_lFindIndexByKey(v_sIndex)
    End If
    m_lReturnIndex = lReturn
End Function

Public Function m_lFindIndexByKey(ByVal v_sKey As String) As Long
On Error GoTo ErrorHandler
    m_lFindIndexByKey = mp_oKeys.Item(v_sKey)
Exit Function

ErrorHandler:
    m_lFindIndexByKey = -1
End Function

Public Function m_bIsKeyUnique(ByVal v_sKey As String) As Boolean
On Error GoTo ErrorHandler
    Dim lDummy As Long
    lDummy = mp_oKeys.Item(v_sKey)
    m_bIsKeyUnique = False
Exit Function

ErrorHandler:
    m_bIsKeyUnique = True
End Function

Public Function m_bDoesKeyExist(ByVal v_sKey As String) As Boolean
On Error GoTo ErrorHandler
    Dim lDummy As Long
    lDummy = mp_oKeys.Item(v_sKey)
    m_bDoesKeyExist = True
Exit Function

ErrorHandler:
    m_bDoesKeyExist = False
End Function

Public Sub m_ReconstKeys()
    Dim lIndex As Long
    Dim lCount As Long
    Dim sKey As String
    Set mp_oKeys = Nothing
    Set mp_oKeys = New Collection
    lCount = UBound(mp_aoCollection)
    For lIndex = 1 To lCount
        sKey = mp_aoCollection(lIndex).Key
        mp_aoCollection(lIndex).Index = lIndex
        If sKey <> "" Then
            mp_oKeys.Add lIndex, sKey
        End If
    Next lIndex
End Sub

Public Property Get m_bIgnoreKeyChecks() As Boolean
    m_bIgnoreKeyChecks = mp_bIgnoreKeyChecks
End Property

Public Property Let m_bIgnoreKeyChecks(ByVal v_bNewValue As Boolean)
    mp_bIgnoreKeyChecks = v_bNewValue
End Property

Private Sub mp_Sort(r_aSortArray As Variant, r_aTempArray As Variant, Optional ByVal first As Long, Optional ByVal last As Long)
    Dim lArrayMBound As Long
    Dim lArrayLBound As Long
    Dim lArrayUBound As Long
    If IsMissing(first) Then
        lArrayLBound = LBound(r_aSortArray)
    Else
        lArrayLBound = CLng(first)
    End If
    If IsMissing(last) Then
        lArrayUBound = UBound(r_aSortArray)
    Else
        lArrayUBound = CLng(last)
    End If
    If lArrayUBound > lArrayLBound Then
        lArrayMBound = (lArrayUBound + lArrayLBound) \ 2
        mp_Sort r_aSortArray, r_aTempArray, lArrayLBound, lArrayMBound
        mp_Sort r_aSortArray, r_aTempArray, lArrayMBound + 1, lArrayUBound
        mp_Merge r_aSortArray, r_aTempArray, lArrayLBound, lArrayMBound + 1, lArrayUBound
    End If
End Sub

Private Sub mp_Merge(r_aSortArray As Variant, r_aTempArray As Variant, first As Long, mid As Long, last As Long)
    Dim i As Long
    Dim iLow As Long
    Dim nNumElements As Long
    Dim iTempPos As Long
    iLow = mid - 1
    iTempPos = first
    nNumElements = last - first + 1
    Do While first <= iLow And mid <= last
        If mp_bDescending = False Then
            If mp_oGetProperty(r_aSortArray(first)) <= mp_oGetProperty(r_aSortArray(mid)) Then
                Set r_aTempArray(iTempPos) = r_aSortArray(first)
                iTempPos = iTempPos + 1
                first = first + 1
            Else
                Set r_aTempArray(iTempPos) = r_aSortArray(mid)
                iTempPos = iTempPos + 1
                mid = mid + 1
            End If
        Else
            If mp_oGetProperty(r_aSortArray(first)) >= mp_oGetProperty(r_aSortArray(mid)) Then
                Set r_aTempArray(iTempPos) = r_aSortArray(first)
                iTempPos = iTempPos + 1
                first = first + 1
            Else
                Set r_aTempArray(iTempPos) = r_aSortArray(mid)
                iTempPos = iTempPos + 1
                mid = mid + 1
            End If
        End If
    Loop
    Do While first <= iLow
        Set r_aTempArray(iTempPos) = r_aSortArray(first)
        first = first + 1
        iTempPos = iTempPos + 1
    Loop
    Do While mid <= last
        Set r_aTempArray(iTempPos) = r_aSortArray(mid)
        mid = mid + 1
        iTempPos = iTempPos + 1
    Loop
    For i = 0 To nNumElements - 1
        Set r_aSortArray(last) = r_aTempArray(last)
        last = last - 1
    Next i
End Sub

Private Function mp_oGetProperty(obj As Variant) As Variant
    If mp_bSortCells = False Then
        Select Case mp_ySortType
            Case E_SORTTYPE.ES_NUMERIC
                mp_oGetProperty = CLng(CallByName(obj, mp_sPropertyName, VbGet))
            Case E_SORTTYPE.ES_STRING
                mp_oGetProperty = CStr(CallByName(obj, mp_sPropertyName, VbGet))
            Case E_SORTTYPE.ES_DATE
                mp_oGetProperty = CDate(CallByName(obj, mp_sPropertyName, VbGet))
        End Select
    Else
        Select Case mp_ySortType
            Case E_SORTTYPE.ES_NUMERIC
                mp_oGetProperty = CLng(CallByName(obj.Cell(mp_lCellIndex), mp_sPropertyName, VbGet))
            Case E_SORTTYPE.ES_STRING
                mp_oGetProperty = CStr(CallByName(obj.Cell(mp_lCellIndex), mp_sPropertyName, VbGet))
            Case E_SORTTYPE.ES_DATE
                mp_oGetProperty = CDate(CallByName(obj.Cell(mp_lCellIndex), mp_sPropertyName, VbGet))
        End Select
    End If
End Function

Public Sub mp_SetKey(ByRef sCurrentKey As String, ByVal sNewKey As String, ByVal ErrNumber As Long)
    If m_bIgnoreKeyChecks = False Then
        If mp_oControl.StrLib.StrIsNumeric(sNewKey) Or (sNewKey <> sCurrentKey And m_bIsKeyUnique(sNewKey) = False) Then
            mp_oControl.mp_ErrorReport ErrNumber, "Numeric or duplicate " & m_sObjectName & " object key", "ActiveGanttVBCtl.cls" & m_sObjectName & ".Let Key"
            Exit Sub
        End If
    End If
    sCurrentKey = sNewKey
    If mp_bAddMode = False Then
        m_ReconstKeys
    End If
End Sub

Public Property Get AddMode() As Boolean
    AddMode = mp_bAddMode
End Property

Public Property Let AddMode(ByVal Value As Boolean)
    mp_bAddMode = Value
End Property

Public Sub m_GetKeyAndIndex(ByVal sIndex As String, ByRef sKey As String, ByRef sReturnIndex As String)
    Dim oObject As Object
    Set oObject = m_oItem(sIndex)
    If oObject.Key <> "" Then
        sKey = oObject.Key
    Else
        sKey = mp_oControl.StrLib.StrCStr(oObject.lIndex)
    End If
    sReturnIndex = mp_oControl.StrLib.StrCStr(oObject.lIndex)
End Sub

Public Sub m_CollectionRemoveWhere(ByVal sPropertyName As String, ByVal sKey As String, ByVal sIndex As String)
    Dim lIndex As Long
    Dim oObject As Object
    Dim sPropertyValue As String
    For lIndex = m_lCount To 1 Step -1
        Set oObject = m_oReturnArrayElement(lIndex)
        sPropertyValue = CallByName(oObject, sPropertyName, VbGet)
        If sPropertyValue = sKey Or sPropertyValue = sIndex Then
            m_Remove sIndex, 50000, 50000, 50000, 50000
        End If
    Next lIndex
End Sub

Public Sub m_CollectionRemoveWhereNot(ByVal sPropertyName As String, ByVal sValue As String)
    Dim lIndex As Long
    Dim oObject As Object
    Dim sPropertyValue As String
    Dim sIndex As String
    For lIndex = m_lCount To 1 Step -1
        Set oObject = m_oReturnArrayElement(lIndex)
        sPropertyValue = CallByName(oObject, sPropertyName, VbGet)
        If sPropertyValue <> sValue Then
            sIndex = mp_oControl.StrLib.StrCStr(lIndex)
            m_Remove sIndex, 50000, 50000, 50000, 50000
        End If
    Next lIndex
End Sub

Public Sub m_CollectionChange(ByVal sPropertyName As String, ByVal sKey As String, ByVal sIndex As String, ByVal sNewValue As String)
    Dim lIndex As Long
    Dim oObject As Object
    Dim sPropertyValue As String
    For lIndex = 1 To m_lCount
        Set oObject = m_oReturnArrayElement(lIndex)
        sPropertyValue = CallByName(oObject, sPropertyName, VbGet)
        If sPropertyValue = sKey Or sPropertyValue = sIndex Then
            CallByName oObject, sPropertyName, VbLet, sNewValue
        End If
    Next lIndex
End Sub

Public Sub m_CollectionChangeAll(ByVal sPropertyName As String, ByVal sNewValue As String)
    Dim lIndex As Long
    Dim oObject As Object
    For lIndex = 1 To m_lCount
        Set oObject = m_oReturnArrayElement(lIndex)
        CallByName oObject, sPropertyName, VbLet, sNewValue
    Next lIndex
End Sub



