VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
Private mp_lImageDC As Long
Private mp_lInvertImageDC As Long
Private mp_lMaskDC As Long
Private mp_lImageOld As Long
Private mp_lInvertImageOld As Long
Private mp_lMaskOld As Long
Private mp_lWidth As Long
Private mp_lHeight As Long
'// Constants
Private Const mp_lSRCAND As Long = &H8800C6
Private Const mp_lSRCCOPY As Long = &HCC0020
Private Const mp_lSRCPAINT As Long = &HEE0086
Private Const mp_lCLRINVALID As Long = &HFFFF
'// Private Types
Private Type T_BITMAP
    lType As Long
    lWidth As Long
    lHeight As Long
    lWidthBytes As Long
    iPlanes As Integer
    iBitsPixel As Integer
    lBits As Long
End Type
'// Win32 API
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    '// Parent Control Pointer
    Set mp_oControl = Value
End Sub

Private Sub Class_Terminate()
    Dim lResult As Long
    If mp_lImageDC = 0 Then Exit Sub
    lResult = DeleteObject(SelectObject(mp_lInvertImageDC, mp_lInvertImageOld))
    lResult = DeleteObject(SelectObject(mp_lMaskDC, mp_lMaskOld))
    DeleteDC mp_lImageDC
    DeleteDC mp_lInvertImageDC
    DeleteDC mp_lMaskDC
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Properties
'// ---------------------------------------------------------------------------------------------------------------------

Public Property Get Width() As Long
    Width = mp_lWidth
End Property

Public Property Get Height() As Long
    Height = mp_lHeight
End Property

'// ---------------------------------------------------------------------------------------------------------------------
'// Methods
'// ---------------------------------------------------------------------------------------------------------------------

Public Sub SetBitmap(ByVal v_lNewBitmap As Long)
    Dim lResult As Long
    Dim lInvertImage As Long
    Dim lMask As Long
    Dim udtBitmapData As T_BITMAP
    lResult = GetObject(v_lNewBitmap, Len(udtBitmapData), udtBitmapData)
    If (lResult = 0) Then
        mp_oControl.mp_ErrorReport 50197, "Couldn't GetObject", "clsBitmap.SetBitmap"
    End If
    mp_lWidth = 0
    mp_lHeight = 0
    mp_lWidth = udtBitmapData.lWidth
    mp_lHeight = udtBitmapData.lHeight
    mp_lImageDC = CreateCompatibleDC(0)
    mp_lInvertImageDC = CreateCompatibleDC(0)
    mp_lMaskDC = CreateCompatibleDC(0)
    If (mp_lImageDC = 0) Or (mp_lInvertImageDC = 0) Or (mp_lMaskDC = 0) Then
        mp_oControl.mp_ErrorReport 50198, "Couldn't CreateCompatibleDC", "clsBitmap.SetBitmap"
    End If
    mp_lImageOld = SelectObject(mp_lImageDC, v_lNewBitmap)
    If (mp_lImageOld = 0) Then
        mp_oControl.mp_ErrorReport 50199, "Couldn't SelectObject(mp_lImageDC, v_lNewBitmap)", "clsBitmap.SetBitmap"
    End If
    lInvertImage = CreateCompatibleBitmap(mp_lImageDC, mp_lWidth, mp_lHeight)
    If (lInvertImage = 0) Then
        mp_oControl.mp_ErrorReport 50200, "Couldn't CreateCompatibleBitmap(mp_lImageDC, mp_lWidth, mp_lHeight)", "clsBitmap.SetBitmap"
    End If
    mp_lInvertImageOld = SelectObject(mp_lInvertImageDC, lInvertImage)
    If (mp_lInvertImageOld = 0) Then
        mp_oControl.mp_ErrorReport 50201, "Couldn't SelectObject(mp_lInvertImageDC, lInvertImage)", "clsBitmap.SetBitmap"
    End If
    lMask = CreateCompatibleBitmap(mp_lMaskDC, mp_lWidth, mp_lHeight)
    If (lMask = 0) Then
        mp_oControl.mp_ErrorReport 50202, "Couldn't CreateCompatibleBitmap(mp_lMaskDC, mp_lWidth, mp_lHeight)", "clsBitmap.SetBitmap"
    End If
    mp_lMaskOld = SelectObject(mp_lMaskDC, lMask)
    If (mp_lMaskOld = 0) Then
        mp_oControl.mp_ErrorReport 50203, "Couldn't SelectObject(mp_lMaskDC, lMask)", "clsBitmap.SetBitmap"
    End If
    lResult = BitBlt(mp_lInvertImageDC, 0, 0, mp_lWidth, mp_lHeight, mp_lImageDC, 0, 0, mp_lSRCCOPY)
    If (lResult = 0) Then
        mp_oControl.mp_ErrorReport 50204, "Couldn't BitBlt(mp_lInvertImageDC, 0, 0, mp_lWidth, mp_lHeight, mp_lImageDC, 0, 0, mp_lSRCCOPY)", "clsBitmap.SetBitmap"
    End If
    lResult = BitBlt(mp_lMaskDC, 0, 0, mp_lWidth, mp_lHeight, mp_lImageDC, 0, 0, mp_lSRCCOPY)
    If (lResult = 0) Then
        mp_oControl.mp_ErrorReport 50205, "Couldn't BitBlt(mp_lMaskDC, 0, 0, mp_lWidth, mp_lHeight, mp_lImageDC, 0, 0, mp_lSRCCOPY)", "clsBitmap.SetBitmap"
    End If
    lResult = SetBkColor(mp_lInvertImageDC, vbBlack)
    If (lResult = mp_lCLRINVALID) Then
        mp_oControl.mp_ErrorReport 50206, "Couldn't SetBkColor(mp_lInvertImageDC, vbBlack)", "clsBitmap.SetBitmap"
    End If
    lResult = SetTextColor(mp_lInvertImageDC, vbWhite)
    If (lResult = mp_lCLRINVALID) Then
        mp_oControl.mp_ErrorReport 50207, "Couldn't SetTextColor(mp_lInvertImageDC, vbWhite)", "clsBitmap.SetBitmap"
    End If
    lResult = BitBlt(mp_lInvertImageDC, 0, 0, mp_lWidth, mp_lHeight, mp_lMaskDC, 0, 0, mp_lSRCAND)
    If (lResult = 0) Then
        mp_oControl.mp_ErrorReport 50208, "Couldn't BitBlt(mp_lInvertImageDC, 0, 0, mp_lWidth, mp_lHeight, mp_lMaskDC, 0, 0, mp_lSRCAND)", "clsBitmap.SetBitmap"
    End If
    lResult = DeleteObject(lInvertImage)
    If (lResult = 0) Then
        mp_oControl.mp_ErrorReport 50209, "Couldn't DeleteObject(lInvertImage)", "clsBitmap.SetBitmap"
    End If
    lResult = DeleteObject(lMask)
    If (lResult = 0) Then
        mp_oControl.mp_ErrorReport 50210, "Couldn't DeleteObject(lMask)", "clsBitmap.SetBitmap"
    End If
End Sub

Public Sub PaintPicture(ByVal v_lhdc As Long, ByVal v_xDest As Long, ByVal v_yDest As Long, ByVal v_bTransparent As Boolean, ByVal v_mp_lWidth As Long, ByVal v_mp_lHeight As Long, ByVal v_lXOrigin As Long, ByVal v_lYOrigin As Long)
    If v_bTransparent = True Then
        BitBlt v_lhdc, v_xDest, v_yDest, v_mp_lWidth, v_mp_lHeight, mp_lMaskDC, v_lXOrigin, v_lYOrigin, mp_lSRCAND
        BitBlt v_lhdc, v_xDest, v_yDest, v_mp_lWidth, v_mp_lHeight, mp_lInvertImageDC, v_lXOrigin, v_lYOrigin, mp_lSRCPAINT
    Else
        BitBlt v_lhdc, v_xDest, v_yDest, v_mp_lWidth, v_mp_lHeight, mp_lImageDC, v_lXOrigin, v_lYOrigin, mp_lSRCCOPY
    End If
End Sub




