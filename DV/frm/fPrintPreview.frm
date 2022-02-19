VERSION 5.00
Begin VB.Form fPrintPreview 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Preview"
   ClientHeight    =   8115
   ClientLeft      =   1890
   ClientTop       =   1935
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Height          =   255
      Left            =   9960
      TabIndex        =   12
      Top             =   6960
      Width           =   255
   End
   Begin VB.VScrollBar mp_lVScroll 
      Height          =   6975
      Left            =   9960
      TabIndex        =   11
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar mp_lHScroll 
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   6960
      Width           =   9975
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   7200
      Width           =   10215
      Begin VB.CommandButton cmdYMarginMinus 
         Caption         =   "-"
         Height          =   255
         Left            =   6840
         TabIndex        =   9
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdYMarginPlus 
         Caption         =   "+"
         Height          =   255
         Left            =   7320
         TabIndex        =   8
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdXMarginMinus 
         Caption         =   "-"
         Height          =   255
         Left            =   6840
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdXMarginPlus 
         Caption         =   "+"
         Height          =   255
         Left            =   7320
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdNextPage 
         Caption         =   "+"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdPreviousPage 
         Caption         =   "-"
         Height          =   255
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblYMargin 
         Caption         =   "YMargin ="
         Height          =   255
         Left            =   5400
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblXMargin 
         Caption         =   "XMargin ="
         Height          =   255
         Left            =   5400
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblPageCaption 
         Caption         =   "Page:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "fPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Private mp_lPageNumber As Long
Public mp_Parent As fPrintDialog

Dim mp_lPhysicalOffsetX As Long
Dim mp_lPhysicalOffsetY As Long
Dim mp_lPhysicalWidth As Long
Dim mp_lPhysicalHeight As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Const PS_SOLID = 0



Private Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Private Const PHYSICALOFFSETX = 112 '  Physical Printable Area x margin
Private Const PHYSICALOFFSETY = 113 '  Physical Printable Area y margin
Private Const PHYSICALWIDTH = 110 '  Physical Width in device units
Private Const PHYSICALHEIGHT = 111 '  Physical Height in device units



Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long








Private Sub cmdNextPage_Click()
    If mp_lPageNumber < fPrintDialog.TotalPages Then
        mp_lPageNumber = mp_lPageNumber + 1
    End If
    WritePageNumber
    Me.Cls
    Form_Paint
End Sub

Private Sub cmdPreviousPage_Click()
    If mp_lPageNumber > 1 Then
        mp_lPageNumber = mp_lPageNumber - 1
    End If
    WritePageNumber
    Me.Cls
    Form_Paint
End Sub

Private Sub cmdXMarginMinus_Click()
    mp_Parent.XMargin = mp_Parent.XMargin - 10
    Me.Cls
    Form_Paint
    WriteXMargin
End Sub

Private Sub cmdXMarginPlus_Click()
    mp_Parent.XMargin = mp_Parent.XMargin + 10
    Me.Cls
    Form_Paint
    WriteXMargin
End Sub

Private Sub cmdYMarginMinus_Click()
    mp_Parent.YMargin = mp_Parent.YMargin - 10
    Me.Cls
    Form_Paint
    WriteYMargin
End Sub

Private Sub cmdYMarginPlus_Click()
    mp_Parent.YMargin = mp_Parent.YMargin + 10
    Me.Cls
    Form_Paint
    WriteYMargin
End Sub

Private Sub Form_Activate()
    Me.Cls
    Form_Paint
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    mp_lPageNumber = 1
    WritePageNumber
    WriteXMargin
    WriteYMargin

    Dim hdc As Long
    Dim sPrinterName As String
    sPrinterName = mp_Parent.cboPrinterName.Text
    hdc = CreateDC("WINSPOOL", sPrinterName, vbNullString, ByVal 0&)
    If (mp_Parent.cboOrientation.ListIndex = 0) Then
        mp_lPhysicalOffsetX = GetDeviceCaps(hdc, PHYSICALOFFSETX) / GetDeviceCaps(hdc, LOGPIXELSX) * GetDeviceCaps(Me.hdc, LOGPIXELSX)
        mp_lPhysicalOffsetY = GetDeviceCaps(hdc, PHYSICALOFFSETY) / GetDeviceCaps(hdc, LOGPIXELSY) * GetDeviceCaps(Me.hdc, LOGPIXELSY)
        mp_lPhysicalWidth = GetDeviceCaps(hdc, PHYSICALWIDTH) / GetDeviceCaps(hdc, LOGPIXELSX) * GetDeviceCaps(Me.hdc, LOGPIXELSX)
        mp_lPhysicalHeight = GetDeviceCaps(hdc, PHYSICALHEIGHT) / GetDeviceCaps(hdc, LOGPIXELSY) * GetDeviceCaps(Me.hdc, LOGPIXELSY)
    Else
        mp_lPhysicalOffsetY = GetDeviceCaps(hdc, PHYSICALOFFSETX) / GetDeviceCaps(hdc, LOGPIXELSX) * GetDeviceCaps(Me.hdc, LOGPIXELSX)
        mp_lPhysicalOffsetX = GetDeviceCaps(hdc, PHYSICALOFFSETY) / GetDeviceCaps(hdc, LOGPIXELSY) * GetDeviceCaps(Me.hdc, LOGPIXELSY)
        mp_lPhysicalHeight = GetDeviceCaps(hdc, PHYSICALWIDTH) / GetDeviceCaps(hdc, LOGPIXELSX) * GetDeviceCaps(Me.hdc, LOGPIXELSX)
        mp_lPhysicalWidth = GetDeviceCaps(hdc, PHYSICALHEIGHT) / GetDeviceCaps(hdc, LOGPIXELSY) * GetDeviceCaps(Me.hdc, LOGPIXELSY)
    End If
    DeleteDC hdc

    mp_lHScroll.Min = 0
    mp_lHScroll.Max = mp_lPhysicalWidth
    mp_lHScroll.LargeChange = 500
    mp_lHScroll.SmallChange = 50
    mp_lHScroll.Value = 0
    
    mp_lVScroll.Min = 0
    mp_lVScroll.Max = mp_lPhysicalHeight
    mp_lVScroll.LargeChange = 500
    mp_lVScroll.SmallChange = 50
    mp_lVScroll.Value = 0
    
End Sub

Private Sub WritePageNumber()
    lblPageCaption.Caption = "Page: " & mp_lPageNumber & " of " & mp_Parent.TotalPages
End Sub

Private Sub WriteXMargin()
    lblXMargin.Caption = "XMargin = " & mp_Parent.XMargin
End Sub

Private Sub WriteYMargin()
    lblYMargin.Caption = "YMargin = " & mp_Parent.YMargin
End Sub

Private Sub Form_Paint()
    Dim hPen As Long
    Dim hDefaultPen As Long
    Dim hBrush As Long
    Dim hDefaultBrush As Long
    Dim lIncrement As Long
    Dim lInch As Long
    Dim udtPageRect As RECT
    Dim udtRuler As RECT
    Dim X As Long
    Dim Y As Long
    Dim Points(1) As POINTAPI
    Dim sCaption As String
    lIncrement = GetDeviceCaps(Me.hdc, LOGPIXELSX) / 16
    lInch = GetDeviceCaps(Me.hdc, LOGPIXELSX)
    hBrush = CreateSolidBrush(&H808080)
    GetClientRect Me.hwnd, udtPageRect
    FillRect Me.hdc, udtPageRect, hBrush
    DeleteObject hBrush
    hBrush = CreateSolidBrush(RGB(255, 255, 255))
    udtPageRect.Left = mp_Parent.m_lXPreviewMargin - mp_lHScroll.Value
    udtPageRect.Top = mp_Parent.m_lYPreviewMargin - mp_lVScroll.Value
    udtPageRect.Right = udtPageRect.Left + mp_lPhysicalWidth
    udtPageRect.Bottom = udtPageRect.Top + mp_lPhysicalHeight
    FillRect Me.hdc, udtPageRect, hBrush
    

    mp_Parent.PrintControl Me.hdc, mp_lPageNumber, False, mp_lPhysicalOffsetX, mp_lPhysicalOffsetY, mp_lHScroll.Value, mp_lVScroll.Value
    
    hPen = CreatePen(PS_SOLID, 1, RGB(0, 0, 0))
    hDefaultPen = SelectObject(Me.hdc, hPen)
    
    
    hBrush = CreateSolidBrush(13160660)
    
    udtRuler.Left = 0
    udtRuler.Top = 0
    udtRuler.Right = 670
    udtRuler.Bottom = 25
    FillRect Me.hdc, udtRuler, hBrush
    For X = mp_Parent.m_lXPreviewMargin To mp_lPhysicalWidth + mp_Parent.m_lXPreviewMargin Step lIncrement
        Points(0).X = X - mp_lHScroll.Value
        If ((X - mp_Parent.m_lXPreviewMargin) Mod 96 = 0) Then
            sCaption = (X - mp_Parent.m_lXPreviewMargin) / lInch
            TextOut Me.hdc, X - mp_lHScroll.Value, 0, sCaption, Len(sCaption)
            Points(0).Y = 14
        ElseIf ((X - mp_Parent.m_lXPreviewMargin) Mod 48 = 0) Then
            Points(0).Y = 16
        ElseIf ((X - mp_Parent.m_lXPreviewMargin) Mod 24 = 0) Then
            Points(0).Y = 18
        Else
            Points(0).Y = 20
        End If
        Points(1).X = X - mp_lHScroll.Value
        Points(1).Y = 25
        Polyline Me.hdc, Points(0), 2
    Next X
    
    udtRuler.Left = 0
    udtRuler.Top = 0
    udtRuler.Right = 25
    udtRuler.Bottom = 500
    FillRect Me.hdc, udtRuler, hBrush
    For Y = mp_Parent.m_lYPreviewMargin To mp_lPhysicalHeight + mp_Parent.m_lYPreviewMargin Step lIncrement
        If ((Y - mp_Parent.m_lXPreviewMargin) Mod 96 = 0) Then
            sCaption = (Y - mp_Parent.m_lYPreviewMargin) / lInch
            TextOut Me.hdc, 0, Y - mp_lVScroll.Value, sCaption, Len(sCaption)
            Points(0).X = 14
        ElseIf ((Y - mp_Parent.m_lXPreviewMargin) Mod 48 = 0) Then
            Points(0).X = 16
        ElseIf ((Y - mp_Parent.m_lXPreviewMargin) Mod 24 = 0) Then
            Points(0).X = 18
        Else
            Points(0).X = 20
        End If
        Points(0).Y = Y - mp_lVScroll.Value
        Points(1).X = 25
        Points(1).Y = Y - mp_lVScroll.Value
        Polyline Me.hdc, Points(0), 2
    Next Y
    udtRuler.Left = 0
    udtRuler.Top = 0
    udtRuler.Right = 25
    udtRuler.Bottom = 25
    FillRect Me.hdc, udtRuler, hBrush
    SelectObject Me.hdc, hDefaultPen
    DeleteObject hBrush
End Sub

Private Sub mp_lHScroll_Change()
    Me.Cls
    Form_Paint
End Sub

Private Sub mp_lVScroll_Change()
    Me.Cls
    Form_Paint
End Sub
