VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXML"
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
'// Object Member Variables
Private xDoc As DOMDocument40
Private oControlElement As IXMLDOMElement
Private oFontElement As IXMLDOMElement
Private mp_sObject As String
Private mp_yLevel As PE_LEVEL

Private Enum PE_LEVEL
    LVL_CONTROL = 0
    LVL_FONT = 5
End Enum

'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl, ByVal sObject As String)
    '// Parent Control Pointer
    Set mp_oControl = Value
    '// Object Member Variables
    Set xDoc = New DOMDocument40
    mp_sObject = sObject
End Sub

Private Sub Class_Initialize()
    '//
End Sub

Private Sub Class_Terminate()
    Set xDoc = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub InitializeWriter()
    xDoc.loadXML "<" & mp_sObject & "></" & mp_sObject & ">"
    Set oControlElement = GetDocumentElement(mp_sObject, 0)
    xDoc.preserveWhiteSpace = False
    mp_yLevel = PE_LEVEL.LVL_CONTROL
End Sub

Public Sub InitializeReader()
    Set oControlElement = GetDocumentElement(mp_sObject, 0)
    mp_yLevel = PE_LEVEL.LVL_CONTROL
End Sub

Public Sub WriteXML(ByVal sPath As String)
    xDoc.save sPath
End Sub

Public Sub ReadXML(ByVal sPath As String)
    xDoc.async = False
    xDoc.Load sPath
End Sub

Private Function ParentElement() As IXMLDOMElement
    Select Case mp_yLevel
        Case PE_LEVEL.LVL_CONTROL
            Set ParentElement = oControlElement
        Case PE_LEVEL.LVL_FONT
            Set ParentElement = oFontElement
    End Select
End Function

Private Function mp_oCreateEmptyDOMElement(ByVal sElementName As String) As IXMLDOMElement
    Dim oNodeBuff As IXMLDOMElement
    Set oNodeBuff = xDoc.createElement(sElementName)
    ParentElement.appendChild oNodeBuff
    Set mp_oCreateEmptyDOMElement = oNodeBuff
End Function

Private Function GetDocumentElement(ByVal TagName As String, ByVal lIndex As Long) As IXMLDOMElement
    Set GetDocumentElement = xDoc.getElementsByTagName(TagName).Item(lIndex)
End Function

'// GetElement

Public Sub SetXML(ByVal sXML As String)
    xDoc.async = False
    xDoc.loadXML sXML
End Sub

Public Function GetXML() As String
    GetXML = xDoc.xml
End Function

Public Function ReadCollectionCount() As Long
    ReadCollectionCount = ParentElement.childNodes.Length
End Function

Public Function ReadObject(ByVal sObjectName As String) As String
    Dim oNodeBuff As IXMLDOMElement
    ReadObject = ParentElement.getElementsByTagName(sObjectName).Item(0).xml
End Function

Public Function ReadCollectionObject(ByVal lIndex As Long) As String
    ReadCollectionObject = ParentElement.childNodes.Item(lIndex - 1).xml
End Function

Public Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As Variant)
    Select Case TypeName(sElementValue)
        Case "Long"
            sElementValue = lReadProperty(sElementName)
        Case "Integer"
            sElementValue = yReadProperty(sElementName)
        Case "String"
            sElementValue = sReadProperty(sElementName)
        Case "Boolean"
            sElementValue = bReadProperty(sElementName)
        Case "Date"
            sElementValue = dtReadProperty(sElementName)
        Case "Currency"
            sElementValue = cReadProperty(sElementName)
        Case "Single"
            sElementValue = fReadProperty(sElementName)
        Case Else
            MsgBox "XML Read Issue"
    End Select
End Sub

Private Function lReadProperty(ByVal v_sNodeName As String) As Long
    lReadProperty = ParentElement.getElementsByTagName(v_sNodeName).Item(0).nodeTypedValue
End Function

Private Function yReadProperty(ByVal v_sNodeName As String) As Integer
    yReadProperty = ParentElement.getElementsByTagName(v_sNodeName).Item(0).nodeTypedValue
End Function

Private Function sReadProperty(ByVal v_sNodeName As String) As String
    sReadProperty = ParentElement.getElementsByTagName(v_sNodeName).Item(0).nodeTypedValue
End Function

Private Function bReadProperty(ByVal v_sNodeName As String) As Boolean
    If ParentElement.getElementsByTagName(v_sNodeName).Item(0).nodeTypedValue = "true" Then
        bReadProperty = True
    Else
        bReadProperty = False
    End If
End Function

Private Function dtReadProperty(ByVal v_sNodeName As String) As Date
    dtReadProperty = mp_dtGetDateFromXML(ParentElement.getElementsByTagName(v_sNodeName).Item(0).nodeTypedValue)
End Function

Private Function mp_dtGetDateFromXML(ByVal sParam As String) As Date
    mp_dtGetDateFromXML = DateSerial(mid$(sParam, 1, 4), mid$(sParam, 6, 2), mid$(sParam, 9, 2)) + TimeSerial(mid$(sParam, 12, 2), mid$(sParam, 15, 2), mid$(sParam, 18, 2))
End Function

Private Function cReadProperty(ByVal v_sNodeName As String) As Currency
    Dim sBuff As String
    sBuff = ParentElement.getElementsByTagName(v_sNodeName).Item(0).nodeTypedValue
    sBuff = Replace(sBuff, ".", Format$(0, "."))
    cReadProperty = sBuff
End Function

Private Function fReadProperty(ByVal v_sNodeName As String) As Single
    fReadProperty = ParentElement.getElementsByTagName(v_sNodeName).Item(0).nodeTypedValue
End Function

Public Sub ReadPropertyColor(ByVal sElementName As String, ByRef clrColor As OLE_COLOR)
    clrColor = lReadProperty(sElementName)
    clrColor = mp_oControl.GrphLib.ConvertColor(clrColor)
End Sub


Public Sub WriteObject(ByVal sObjectText As String)
    Dim xDoc1 As DOMDocument40
    Dim oNodeBuff As IXMLDOMElement
    Set xDoc1 = New DOMDocument40
    xDoc1.loadXML sObjectText
    Set oNodeBuff = xDoc1.documentElement
    ParentElement.appendChild oNodeBuff
End Sub



Public Sub WritePropertyFont(ByVal sElementName As String, ByRef oFont As StdFont)
    Dim mp_yBackupLevel As PE_LEVEL
    Set oFontElement = mp_oCreateEmptyDOMElement(sElementName)
    mp_yBackupLevel = mp_yLevel
    mp_yLevel = PE_LEVEL.LVL_FONT
    WriteProperty "Name", oFont.Name
    WriteProperty "Size", Replace(CStr(oFont.Size), Format$(0, "."), ".")
    WriteProperty "Bold", oFont.Bold
    WriteProperty "Italic", oFont.Italic
    WriteProperty "Underline", oFont.Underline
    mp_yLevel = mp_yBackupLevel
End Sub

Public Sub ReadPropertyFont(ByVal sElementName, ByRef r_oFont As StdFont)
    Dim sName As String
    Dim cSize As Currency
    Dim bBold As Boolean
    Dim bItalic As Boolean
    Dim bUnderline As Boolean
    Dim mp_yBackupLevel As PE_LEVEL
    Set oFontElement = ParentElement.getElementsByTagName(sElementName).Item(0)
    mp_yBackupLevel = mp_yLevel
    mp_yLevel = PE_LEVEL.LVL_FONT
    ReadProperty "Name", sName
    ReadProperty "Size", cSize
    ReadProperty "Bold", bBold
    ReadProperty "Italic", bItalic
    ReadProperty "Underline", bUnderline
    mp_yLevel = mp_yBackupLevel
    r_oFont.Name = sName
    r_oFont.Size = cSize
    r_oFont.Bold = bBold
    r_oFont.Italic = bItalic
    r_oFont.Underline = bUnderline
End Sub

Public Sub WriteProperty(ByVal sElementName As String, ByVal sElementValue As Variant)
    Dim oNodeBuff As IXMLDOMElement
    Set oNodeBuff = xDoc.createElement(sElementName)
    If TypeName(sElementValue) = "Date" Then
        oNodeBuff.nodeTypedValue = mp_sGetXMLDateString(sElementValue)
    ElseIf TypeName(sElementValue) = "Boolean" Then
        If sElementValue = True Then
            oNodeBuff.nodeTypedValue = "true"
        Else
            oNodeBuff.nodeTypedValue = "false"
        End If
    Else
        oNodeBuff.nodeTypedValue = CStr(sElementValue)
    End If
    ParentElement.appendChild oNodeBuff
End Sub

Public Sub WritePropertyPicture(ByVal sElementName As String, ByRef oPicture As StdPicture)
    Dim oNodeBuff As IXMLDOMElement
    Dim lLen As Long
    Dim iFile As Integer
    Dim arrBytes() As Byte
    Dim lCount As Long
    Set oNodeBuff = xDoc.createElement(sElementName)
    If Not oPicture Is Nothing Then
        SavePicture oPicture, App.Path & "\AGTemp.bmp"
        iFile = FreeFile()
        Open App.Path & "\AGTemp.bmp" For Binary Access Read As iFile
        lLen = FileLen(App.Path & "\AGTemp.bmp")
        ReDim arrBytes(lLen - 1)
        Get iFile, , arrBytes
        Close iFile
        oNodeBuff.dataType = "bin.base64"
        oNodeBuff.nodeTypedValue = arrBytes()
    End If
    ParentElement.appendChild oNodeBuff
End Sub

Public Sub ReadPropertyPicture(ByVal sElementName As String, ByRef oPicture As StdPicture)
    Dim arrBuffer() As Byte
    Dim iFile As Integer
    If ParentElement.getElementsByTagName(sElementName).Item(0).Text <> "" Then
        arrBuffer = ParentElement.getElementsByTagName(sElementName).Item(0).nodeTypedValue
        iFile = FreeFile()
        Open App.Path & "\AGTemp.bmp" For Binary Access Write As iFile
        Put iFile, , arrBuffer
        Close iFile
        Set oPicture = LoadPicture(App.Path & "\AGTemp.bmp")
    Else
        Set oPicture = Nothing
    End If
End Sub

Private Function mp_sGetXMLDateString(ByVal dtParam As Date) As String
    mp_sGetXMLDateString = Year(dtParam) & "-" & Format$(Month(dtParam), "00") & "-" & Format$(Day(dtParam), "00") & "T" & Format$(Hour(dtParam), "00") & ":" & Format$(Minute(dtParam), "00") & ":" & Format$(Second(dtParam), "00")
End Function
