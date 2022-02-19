VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphics"
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

Private Const MM_ISOTROPIC = 7
Private Const LF_FACESIZE = 32

Private Type T_PRECT
    lLeft As Long
    lTop As Long
    lRight As Long
    lBottom As Long
End Type

Private Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long

Private Declare Function LPtoDP Lib "gdi32" (ByVal hdc As Long, lpPoint As GRE_POINTTYPE, ByVal nCount As Long) As Long
Private Declare Function SetViewportExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As GRE_POINTTYPE) As Long
Private Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As GRE_POINTTYPE) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As GRE_POINTTYPE) As Long
Private Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As GRE_POINTTYPE) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As T_PRECT, ByVal wFormat As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As T_PRECT) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As GRE_POINTTYPE, ByVal nCount As Long) As Long
Private Declare Function SetPolyFillMode Lib "GDI32.DLL" (ByVal hdc As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function Polygon Lib "GDI32.DLL" (ByVal hdc As Long, lpPoint As GRE_POINTTYPE, ByVal nCount As Long) As Long
Private Declare Function CreateSolidBrush Lib "GDI32.DLL" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "GDI32.DLL" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As T_PRECT, lpSrc1Rect As T_PRECT, lpSrc2Rect As T_PRECT) As Long
Private Declare Function IntersectClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As T_PRECT, ByVal bErase As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As T_PRECT, ByVal hBrush As Long) As Long

'// Parent Control Pointer
Private mp_oControl As ActiveGanttVBCtl
'// Object Member Variables
Private mp_udtPreviousClipRegion As T_PRECT
Private mp_audtActiveReversibleFrames() As T_PRECT
Private mp_bCustomPrinting As Boolean
Private mp_lCustomDC As Long
Private mp_lBackupDC As Long
Private mp_lPWidth As Long
Private mp_lPHeight As Long
Private mp_udtPrintingClipRegion As T_PRECT
Private mp_lFocusLeft As Long
Private mp_lFocusTop As Long
Private mp_lFocusRight As Long
Private mp_lFocusBottom  As Long

'// ---------------------------------------------------------------------------------------------------------------------
'// Construction/Destruction & Initialization
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    Set mp_oControl = Value
    ReDim mp_audtActiveReversibleFrames(0)
    mp_bCustomPrinting = False
End Sub

Private Sub Class_Terminate()
    Set mp_oControl = Nothing
End Sub

'// ---------------------------------------------------------------------------------------------------------------------
'// Object Member Variables
'// ---------------------------------------------------------------------------------------------------------------------

Friend Property Get f_FocusLeft() As Long
    f_FocusLeft = mp_lFocusLeft
End Property

Friend Property Let f_FocusLeft(ByVal Value As Long)
    mp_lFocusLeft = Value
End Property

Friend Property Get f_FocusTop() As Long
    f_FocusTop = mp_lFocusTop
End Property

Friend Property Let f_FocusTop(ByVal Value As Long)
    mp_lFocusTop = Value
End Property

Friend Property Get f_FocusRight() As Long
    f_FocusRight = mp_lFocusRight
End Property

Friend Property Let f_FocusRight(ByVal Value As Long)
    mp_lFocusRight = Value
End Property

Friend Property Get f_FocusBottom() As Long
    f_FocusBottom = mp_lFocusBottom
End Property

Friend Property Let f_FocusBottom(ByVal Value As Long)
    mp_lFocusBottom = Value
End Property

Friend Property Get mp_lHdc() As Long
    If mp_bCustomPrinting = False Then
        mp_lHdc = mp_oControl.f_Hdc
    Else
        mp_lHdc = mp_lCustomDC
    End If
End Property

Public Property Get CustomPrinting() As Boolean
    CustomPrinting = mp_bCustomPrinting
End Property

Public Property Let CustomPrinting(ByVal Value As Boolean)
    mp_bCustomPrinting = Value
End Property

Public Property Let CustomDC(ByVal Value As Long)
    mp_lCustomDC = Value
End Property

Public Property Get CustomWidth() As Long
    CustomWidth = mp_lPWidth
End Property

Public Property Let CustomWidth(ByVal Value As Long)
    mp_lPWidth = Value
End Property

Public Property Get CustomHeight() As Long
    CustomHeight = mp_lPHeight
End Property

Public Property Let CustomHeight(ByVal Value As Long)
    mp_lPHeight = Value
End Property

Public Function Width() As Long
    If mp_bCustomPrinting = False Then
        Width = mp_oControl.f_Width
    Else
        Width = mp_lPWidth
    End If
End Function

Public Function Height() As Long
    If mp_bCustomPrinting = False Then
        Height = mp_oControl.f_Height
    Else
        Height = mp_lPHeight
    End If
End Function

Public Sub DrawArrow(ByVal v_X As Long, ByVal v_Y As Long, ByVal v_ArrowDirection As GRE_ARROWDIRECTION, ByVal v_ArrowSize As Integer, ByVal v_lColor As OLE_COLOR)
    Dim Points() As GRE_POINTTYPE
    Select Case v_ArrowDirection
        Case GRE_ARROWDIRECTION.AWD_UP
            ReDim Points(2)
            Points(0).X = v_X
            Points(0).Y = v_Y + 1
            Points(1).X = v_X - v_ArrowSize
            Points(1).Y = v_Y + v_ArrowSize + 1
            Points(2).X = v_X + v_ArrowSize
            Points(2).Y = v_Y + v_ArrowSize + 1
        Case GRE_ARROWDIRECTION.AWD_DOWN
            ReDim Points(2)
            Points(0).X = v_X
            Points(0).Y = v_Y - 1
            Points(1).X = v_X - v_ArrowSize
            Points(1).Y = v_Y - v_ArrowSize - 1
            Points(2).X = v_X + v_ArrowSize
            Points(2).Y = v_Y - v_ArrowSize - 1
        Case GRE_ARROWDIRECTION.AWD_LEFT
            ReDim Points(2)
            Points(0).X = v_X + 1
            Points(0).Y = v_Y
            Points(1).X = v_X + v_ArrowSize + 1
            Points(1).Y = v_Y + v_ArrowSize
            Points(2).X = v_X + v_ArrowSize + 1
            Points(2).Y = v_Y - v_ArrowSize
        Case GRE_ARROWDIRECTION.AWD_RIGHT
            ReDim Points(2)
            Points(0).X = v_X - 1
            Points(0).Y = v_Y
            Points(1).X = v_X - v_ArrowSize - 1
            Points(1).Y = v_Y + v_ArrowSize
            Points(2).X = v_X - v_ArrowSize - 1
            Points(2).Y = v_Y - v_ArrowSize
    End Select
    DrawPolygon v_lColor, Points(0), 3
End Sub

Private Sub DrawPolygon(ByVal v_lColor As OLE_COLOR, ByRef r_oPoints As GRE_POINTTYPE, ByVal v_Len As Long)
    Dim hBrush As Long
    Dim HoldBrush As Long
    Dim RetVal As Long
    Dim hPen As Long
    Dim HoldPen As Long
    hBrush = CreateSolidBrush(ConvertColor(v_lColor))
    HoldBrush = SelectObject(mp_lHdc, hBrush)
    hPen = CreatePen(GRE_LINEDRAWSTYLE.LDS_SOLID, 1, ConvertColor(v_lColor))
    HoldPen = SelectObject(mp_lHdc, hPen)
    RetVal = SetPolyFillMode(mp_lHdc, 2)
    RetVal = Polygon(mp_lHdc, r_oPoints, v_Len)
    RetVal = SelectObject(mp_lHdc, HoldBrush)
    RetVal = DeleteObject(hBrush)
    RetVal = SelectObject(mp_lHdc, HoldPen)
    RetVal = DeleteObject(hPen)
End Sub

Public Function ConvertColor(ByVal dwOleColor As OLE_COLOR) As Long
   Dim clrref As Long
   OleTranslateColor dwOleColor, 0, clrref
   ConvertColor = clrref
End Function

Public Sub DrawEdge(ByVal v_X1 As Single, ByVal v_Y1 As Single, ByVal v_X2 As Single, ByVal v_Y2 As Single, ByVal clrBackColor As OLE_COLOR, ByVal v_yButtonStyle As GRE_BUTTONSTYLE, Optional ByVal v_lEdgeType As GRE_EDGETYPE = GRE_EDGETYPE.ET_RAISED, Optional ByVal v_bFilled As Boolean = False)
    If RectangleIntersection(v_X1, v_Y1, v_X2, v_Y2) = False Then
        Exit Sub
    End If
    Dim lExteriorLeftTopColor As OLE_COLOR
    Dim lInteriorLeftTopColor As OLE_COLOR
    Dim lExteriorRightBottomColor As OLE_COLOR
    Dim lInteriorRightBottomColor As OLE_COLOR
    If v_yButtonStyle = GRE_BUTTONSTYLE.BT_NORMALWINDOWS Then
        Select Case v_lEdgeType
            Case GRE_EDGETYPE.ET_RAISED
                lExteriorLeftTopColor = GRE_COLORS.CLR_WHITE
                lInteriorLeftTopColor = GRE_COLORS.CLR_BUTTONFACE
                lInteriorRightBottomColor = GRE_COLORS.CLR_DARKGREY
                lExteriorRightBottomColor = GRE_COLORS.CLR_VERYDARKGREY
            Case GRE_EDGETYPE.ET_SUNKEN
                lExteriorLeftTopColor = GRE_COLORS.CLR_DARKGREY
                lInteriorLeftTopColor = GRE_COLORS.CLR_VERYDARKGREY
                lInteriorRightBottomColor = GRE_COLORS.CLR_BUTTONFACE
                lExteriorRightBottomColor = GRE_COLORS.CLR_VERYLIGHTGREY
        End Select
        DrawLine v_X1, v_Y1, v_X2, v_Y1, GRE_LINETYPE.LT_NORMAL, lExteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        DrawLine v_X1, v_Y1, v_X1, v_Y2, GRE_LINETYPE.LT_NORMAL, lExteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        DrawLine v_X1 + 1, v_Y1 + 1, v_X2 - 1, v_Y1 + 1, GRE_LINETYPE.LT_NORMAL, lInteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        DrawLine v_X1 + 1, v_Y1 + 1, v_X1 + 1, v_Y2 - 1, GRE_LINETYPE.LT_NORMAL, lInteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        DrawLine v_X1, v_Y2, v_X2, v_Y2, GRE_LINETYPE.LT_NORMAL, lExteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        DrawLine v_X2, v_Y2, v_X2, v_Y1 - 1, GRE_LINETYPE.LT_NORMAL, lExteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        DrawLine v_X1 + 1, v_Y2 - 1, v_X2 - 1, v_Y2 - 1, GRE_LINETYPE.LT_NORMAL, lInteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        DrawLine v_X2 - 1, v_Y2 - 1, v_X2 - 1, v_Y1, GRE_LINETYPE.LT_NORMAL, lInteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        If v_bFilled = True Then
            DrawLine v_X1 + 2, v_Y1 + 2, v_X2 - 2, v_Y2 - 2, GRE_LINETYPE.LT_FILLED, clrBackColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        End If
    Else
        Select Case v_lEdgeType
            Case GRE_EDGETYPE.ET_RAISED
                lExteriorLeftTopColor = GRE_COLORS.CLR_WHITE
                lExteriorRightBottomColor = GRE_COLORS.CLR_VERYDARKGREY
            Case GRE_EDGETYPE.ET_SUNKEN
                lExteriorLeftTopColor = GRE_COLORS.CLR_DARKGREY
                lExteriorRightBottomColor = GRE_COLORS.CLR_VERYLIGHTGREY
        End Select
        DrawLine v_X1, v_Y1, v_X2, v_Y1, GRE_LINETYPE.LT_NORMAL, lExteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        DrawLine v_X1, v_Y1, v_X1, v_Y2, GRE_LINETYPE.LT_NORMAL, lExteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        DrawLine v_X1, v_Y2, v_X2, v_Y2, GRE_LINETYPE.LT_NORMAL, lExteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        DrawLine v_X2, v_Y2, v_X2, v_Y1 - 1, GRE_LINETYPE.LT_NORMAL, lExteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        If v_bFilled = True Then
            DrawLine v_X1 + 1, v_Y1 + 1, v_X2 - 1, v_Y2 - 1, GRE_LINETYPE.LT_FILLED, clrBackColor, GRE_LINEDRAWSTYLE.LDS_SOLID
        End If
    End If
End Sub

Public Sub DrawLine(ByVal v_X1 As Long, ByVal v_Y1 As Long, ByVal v_X2 As Long, ByVal v_Y2 As Long, ByVal v_yStyle As GRE_LINETYPE, ByVal v_lColor As OLE_COLOR, ByVal v_lDrawStyle As GRE_LINEDRAWSTYLE, Optional ByVal v_lWidth As Integer = 1, Optional ByVal v_bCreatePens As Boolean = True)
    Dim Points() As GRE_POINTTYPE
    Dim hBrush As Long
    Dim HoldBrush As Long
    Dim hPen As Long
    Dim HoldPen As Long
    If v_bCreatePens = True Then
        hPen = CreatePen(v_lDrawStyle, v_lWidth, ConvertColor(v_lColor))
        HoldPen = SelectObject(mp_lHdc, hPen)
    End If
    Select Case v_yStyle
        Case GRE_LINETYPE.LT_NORMAL
            If LineIntersection(v_X1, v_Y1, v_X2, v_Y2) = True Then
                ReDim Points(1)
                Points(0).X = v_X1
                Points(0).Y = v_Y1
                Points(1).X = v_X2
                Points(1).Y = v_Y2
                Polyline mp_lHdc, Points(0), 2
            End If
        Case GRE_LINETYPE.LT_BORDER
            If RectangleIntersection(v_X1, v_Y1, v_X2, v_Y2) = True Then
                ReDim Points(4)
                Points(0).X = v_X1
                Points(0).Y = v_Y1
                Points(1).X = v_X2
                Points(1).Y = v_Y1
                Points(2).X = v_X2
                Points(2).Y = v_Y2
                Points(3).X = v_X1
                Points(3).Y = v_Y2
                Points(4).X = v_X1
                Points(4).Y = v_Y1
                Polyline mp_lHdc, Points(0), 5
            End If
        Case GRE_LINETYPE.LT_FILLED
            If RectangleIntersection(v_X1, v_Y1, v_X2, v_Y2) = True Then
                ReDim Points(4)
                Points(0).X = v_X1
                Points(0).Y = v_Y1
                Points(1).X = v_X2
                Points(1).Y = v_Y1
                Points(2).X = v_X2
                Points(2).Y = v_Y2
                Points(3).X = v_X1
                Points(3).Y = v_Y2
                hBrush = CreateSolidBrush(ConvertColor(v_lColor))
                HoldBrush = SelectObject(mp_lHdc, hBrush)
                SetPolyFillMode mp_lHdc, 2
                Polygon mp_lHdc, Points(0), 4
                SelectObject mp_lHdc, HoldBrush
                DeleteObject hBrush
            End If
    End Select
    If v_bCreatePens = True Then
        SelectObject mp_lHdc, HoldPen
        DeleteObject hPen
    End If
End Sub

Public Sub DrawFigure(ByVal v_X As Long, ByVal v_Y As Long, ByVal v_dx As Long, ByVal v_dy As Long, ByVal v_yFigureType As GRE_FIGURETYPE, ByVal v_lBorderColor As OLE_COLOR, ByVal v_lFillColor As OLE_COLOR, ByVal v_yBorderStyle As GRE_LINEDRAWSTYLE)
    Dim Points() As GRE_POINTTYPE
    Dim RetVal As Long
    Dim hBrush As Long
    Dim HoldBrush As Long
    Dim NumPoints As Long
    Dim bCircle As Boolean
    If v_dx Mod 2 <> 0 Then
        v_dx = v_dx + 1
        v_dy = v_dy + 1
    End If
    Select Case v_yFigureType
        Case GRE_FIGURETYPE.FT_PROJECTUP
            NumPoints = 5
            ReDim Points(NumPoints - 1)
            Points(0).X = v_X
            Points(0).Y = v_Y
            Points(1).X = v_X + v_dx / 2
            Points(1).Y = v_Y + v_dy / 2
            Points(2).X = v_X + v_dx / 2
            Points(2).Y = v_Y + v_dy
            Points(3).X = v_X - v_dx / 2
            Points(3).Y = v_Y + v_dy
            Points(4).X = v_X - v_dx / 2
            Points(4).Y = v_Y + v_dy / 2
        Case GRE_FIGURETYPE.FT_PROJECTDOWN
            NumPoints = 5
            ReDim Points(NumPoints - 1)
            Points(0).X = v_X + v_dx / 2
            Points(0).Y = v_Y
            Points(1).X = v_X + v_dx / 2
            Points(1).Y = v_Y + v_dy / 2
            Points(2).X = v_X
            Points(2).Y = v_Y + v_dy
            Points(3).X = v_X - v_dx / 2
            Points(3).Y = v_Y + v_dy / 2
            Points(4).X = v_X - v_dx / 2
            Points(4).Y = v_Y
        Case GRE_FIGURETYPE.FT_DIAMOND
            NumPoints = 4
            ReDim Points(NumPoints - 1)
            Points(0).X = v_X
            Points(0).Y = v_Y
            Points(1).X = v_X + v_dx / 2
            Points(1).Y = v_Y + v_dy / 2
            Points(2).X = v_X
            Points(2).Y = v_Y + v_dy
            Points(3).X = v_X - v_dx / 2
            Points(3).Y = v_Y + v_dy / 2
        Case GRE_FIGURETYPE.FT_CIRCLEDIAMOND
            NumPoints = 4
            ReDim Points(NumPoints - 1)
            Points(0).X = v_X
            Points(0).Y = v_Y + v_dy / 4
            Points(1).X = v_X + v_dx / 4
            Points(1).Y = v_Y + v_dy / 2
            Points(2).X = v_X
            Points(2).Y = v_Y + (3 * v_dy) / 4
            Points(3).X = v_X - v_dx / 4
            Points(3).Y = v_Y + v_dy / 2
            bCircle = True
        Case GRE_FIGURETYPE.FT_TRIANGLEUP
            NumPoints = 3
            ReDim Points(NumPoints - 1)
            Points(0).X = v_X
            Points(0).Y = v_Y
            Points(1).X = v_X + v_dx / 2
            Points(1).Y = v_Y + v_dy
            Points(2).X = v_X - v_dx / 2
            Points(2).Y = v_Y + v_dy
        Case GRE_FIGURETYPE.FT_TRIANGLEDOWN
            NumPoints = 3
            ReDim Points(NumPoints - 1)
            Points(0).X = v_X + v_dx / 2
            Points(0).Y = v_Y
            Points(1).X = v_X - v_dx / 2
            Points(1).Y = v_Y
            Points(2).X = v_X
            Points(2).Y = v_Y + v_dy
        Case GRE_FIGURETYPE.FT_TRIANGLERIGHT
            NumPoints = 3
            ReDim Points(NumPoints - 1)
            Points(0).X = v_X
            Points(0).Y = v_Y
            Points(1).X = v_X
            Points(1).Y = v_Y + v_dy
            Points(2).X = v_X + v_dx
            Points(2).Y = v_Y + v_dy / 2
        Case GRE_FIGURETYPE.FT_TRIANGLELEFT
            NumPoints = 3
            ReDim Points(NumPoints - 1)
            Points(0).X = v_X
            Points(0).Y = v_Y
            Points(1).X = v_X
            Points(1).Y = v_Y + v_dy
            Points(2).X = v_X - v_dx
            Points(2).Y = v_Y + v_dy / 2
        Case GRE_FIGURETYPE.FT_CIRCLETRIANGLEUP
            NumPoints = 3
            ReDim Points(NumPoints - 1)
            Points(0).X = v_X
            Points(0).Y = v_Y + v_dy / 4
            Points(1).X = v_X + v_dx / 4
            Points(1).Y = v_Y + (3 * v_dy) / 4
            Points(2).X = v_X - v_dx / 4
            Points(2).Y = v_Y + (3 * v_dy) / 4
            bCircle = True
        Case GRE_FIGURETYPE.FT_CIRCLETRIANGLEDOWN
            NumPoints = 3
            ReDim Points(NumPoints - 1)
            Points(0).X = v_X - v_dx / 4
            Points(0).Y = v_Y + v_dy / 4
            Points(1).X = v_X + v_dx / 4
            Points(1).Y = v_Y + v_dy / 4
            Points(2).X = v_X
            Points(2).Y = v_Y + (3 * v_dy) / 4
            bCircle = True
        Case GRE_FIGURETYPE.FT_ARROWUP
            NumPoints = 7
            ReDim Points(NumPoints - 1)
            Points(0).X = v_X
            Points(0).Y = v_Y
            Points(1).X = v_X + v_dx / 2
            Points(1).Y = v_Y + v_dy / 2
            Points(2).X = v_X + v_dx / 4
            Points(2).Y = v_Y + v_dy / 2
            Points(3).X = v_X + v_dx / 4
            Points(3).Y = v_Y + v_dy
            Points(4).X = v_X - v_dx / 4
            Points(4).Y = v_Y + v_dy
            Points(5).X = v_X - v_dx / 4
            Points(5).Y = v_Y + v_dy / 2
            Points(6).X = v_X - v_dx / 2
            Points(6).Y = v_Y + v_dy / 2
        Case GRE_FIGURETYPE.FT_ARROWDOWN
            NumPoints = 7
            ReDim Points(NumPoints - 1)
            Points(0).X = v_X - v_dx / 4
            Points(0).Y = v_Y
            Points(1).X = v_X + v_dx / 4
            Points(1).Y = v_Y
            Points(2).X = v_X + v_dx / 4
            Points(2).Y = v_Y + v_dy / 2
            Points(3).X = v_X + v_dx / 2
            Points(3).Y = v_Y + v_dy / 2
            Points(4).X = v_X
            Points(4).Y = v_Y + v_dy
            Points(5).X = v_X - v_dx / 2
            Points(5).Y = v_Y + v_dy / 2
            Points(6).X = v_X - v_dx / 4
            Points(6).Y = v_Y + v_dy / 2
        Case GRE_FIGURETYPE.FT_CIRCLEARROWUP
            NumPoints = 7
            ReDim Points(NumPoints - 1)
            Points(0).X = v_X
            Points(0).Y = v_Y + v_dy / 4
            Points(1).X = v_X + v_dx / 4
            Points(1).Y = v_Y + v_dy / 2
            Points(2).X = v_X + v_dx / 8
            Points(2).Y = v_Y + v_dy / 2
            Points(3).X = v_X + v_dx / 8
            Points(3).Y = v_Y + (3 * v_dy) / 4
            Points(4).X = v_X - v_dx / 8
            Points(4).Y = v_Y + (3 * v_dy) / 4
            Points(5).X = v_X - v_dx / 8
            Points(5).Y = v_Y + v_dy / 2
            Points(6).X = v_X - v_dx / 4
            Points(6).Y = v_Y + v_dy / 2
            bCircle = True
        Case GRE_FIGURETYPE.FT_CIRCLEARROWDOWN
            NumPoints = 7
            ReDim Points(NumPoints - 1)
            Points(0).X = v_X - v_dx / 8
            Points(0).Y = v_Y + v_dy / 4
            Points(1).X = v_X + v_dx / 8
            Points(1).Y = v_Y + v_dy / 4
            Points(2).X = v_X + v_dx / 8
            Points(2).Y = v_Y + v_dy / 2
            Points(3).X = v_X + v_dx / 4
            Points(3).Y = v_Y + v_dy / 2
            Points(4).X = v_X
            Points(4).Y = v_Y + (3 * v_dy) / 4
            Points(5).X = v_X - v_dx / 4
            Points(5).Y = v_Y + v_dy / 2
            Points(6).X = v_X - v_dx / 8
            Points(6).Y = v_Y + v_dy / 2
            bCircle = True
        Case GRE_FIGURETYPE.FT_SMALLPROJECTUP
            NumPoints = 5
            ReDim Points(NumPoints - 1)
            Points(0).X = v_X
            Points(0).Y = v_Y + v_dy / 2
            Points(1).X = v_X + v_dx / 4
            Points(1).Y = v_Y + (3 * v_dy) / 4
            Points(2).X = v_X + v_dx / 4
            Points(2).Y = v_Y + v_dy
            Points(3).X = v_X - v_dx / 4
            Points(3).Y = v_Y + v_dy
            Points(4).X = v_X - v_dx / 4
            Points(4).Y = v_Y + (3 * v_dy) / 4
        Case GRE_FIGURETYPE.FT_SMALLPROJECTDOWN
            NumPoints = 5
            ReDim Points(NumPoints - 1)
            Points(0).X = v_X + v_dx / 4
            Points(0).Y = v_Y
            Points(1).X = v_X + v_dx / 4
            Points(1).Y = v_Y + v_dy / 4
            Points(2).X = v_X
            Points(2).Y = v_Y + v_dy / 2
            Points(3).X = v_X - v_dx / 4
            Points(3).Y = v_Y + v_dy / 4
            Points(4).X = v_X - v_dx / 4
            Points(4).Y = v_Y
        Case GRE_FIGURETYPE.FT_RECTANGLE
            NumPoints = 4
            ReDim Points(NumPoints - 1)
            Points(0).X = v_X - v_dx / 8
            Points(0).Y = v_Y
            Points(1).X = v_X + v_dx / 8
            Points(1).Y = v_Y
            Points(2).X = v_X + v_dx / 8
            Points(2).Y = v_Y + v_dy
            Points(3).X = v_X - v_dx / 8
            Points(3).Y = v_Y + v_dy
        Case GRE_FIGURETYPE.FT_SQUARE
            NumPoints = 4
            ReDim Points(NumPoints - 1)
            Points(0).X = v_X - v_dx / 4
            Points(0).Y = v_Y + v_dx / 4
            Points(1).X = v_X + v_dx / 4
            Points(1).Y = v_Y + v_dx / 4
            Points(2).X = v_X + v_dx / 4
            Points(2).Y = v_Y + (3 * v_dy) / 4
            Points(3).X = v_X - v_dx / 4
            Points(3).Y = v_Y + (3 * v_dy) / 4
        Case GRE_FIGURETYPE.FT_CIRCLE
             hBrush = CreateSolidBrush(v_lFillColor)
             HoldBrush = SelectObject(mp_lHdc, hBrush)
             Ellipse mp_lHdc, v_X - v_dx / 2, v_Y, v_X + v_dx / 2, v_Y + v_dy
             RetVal = DeleteObject(hBrush)
             SelectObject mp_lHdc, hBrush
             Exit Sub
        Case Else
            Exit Sub
    End Select
    If bCircle = True Then
        hBrush = CreateSolidBrush(RGB(255, 255, 255))
        HoldBrush = SelectObject(mp_lHdc, hBrush)
        Ellipse mp_lHdc, v_X - v_dx / 2, v_Y, v_X + v_dx / 2, v_Y + v_dy
        RetVal = DeleteObject(hBrush)
        hBrush = CreateSolidBrush(v_lFillColor)
        Ellipse mp_lHdc, v_X - v_dx / 2, v_Y, v_X + v_dx / 2, v_Y + v_dy
        SelectObject mp_lHdc, hBrush
    Else
        hBrush = CreateSolidBrush(v_lFillColor)
        HoldBrush = SelectObject(mp_lHdc, hBrush)
    End If
    RetVal = SetPolyFillMode(mp_lHdc, 2)
    RetVal = Polygon(mp_lHdc, Points(0), NumPoints)
    RetVal = SelectObject(mp_lHdc, HoldBrush)
    RetVal = DeleteObject(hBrush)
End Sub

Public Sub DrawFocusRectangle(ByVal v_X1 As Long, ByVal v_Y1 As Long, ByVal v_X2 As Long, ByVal v_Y2 As Long)
    Dim HoldPen As Long
    Dim hPen As Long
    Dim RetVal As Long
    Dim udtRect As T_PRECT
    udtRect.lLeft = v_X1
    udtRect.lTop = v_Y1
    udtRect.lRight = v_X2 + 1
    udtRect.lBottom = v_Y2 + 1
    hPen = CreatePen(GRE_LINEDRAWSTYLE.LDS_SOLID, 1, ConvertColor(GRE_COLORS.CLR_BLACK))
    HoldPen = SelectObject(mp_lHdc, hPen)
    RetVal = DrawFocusRect(mp_lHdc, udtRect)
    RetVal = SelectObject(mp_lHdc, HoldPen)
    RetVal = DeleteObject(hPen)
End Sub

Public Sub DrawReversibleFrame(ByVal v_X1 As Long, ByVal v_Y1 As Long, ByVal v_X2 As Long, ByVal v_Y2 As Long)
    Dim HoldPen As Long
    Dim hPen As Long
    Dim RetVal As Long
    Dim udtRect As T_PRECT
    Dim lARFUBound As Long
    udtRect.lLeft = v_X1
    udtRect.lTop = v_Y1
    udtRect.lRight = v_X2 + 1
    udtRect.lBottom = v_Y2 + 1
    hPen = CreatePen(GRE_LINEDRAWSTYLE.LDS_SOLID, 1, ConvertColor(GRE_COLORS.CLR_BLACK))
    HoldPen = SelectObject(mp_lHdc, hPen)
    RetVal = DrawFocusRect(mp_lHdc, udtRect)
    mp_oControl.f_DrawLine 0, 0
    lARFUBound = UBound(mp_audtActiveReversibleFrames, 1) + 1
    ReDim Preserve mp_audtActiveReversibleFrames(lARFUBound)
    mp_audtActiveReversibleFrames(lARFUBound).lLeft = udtRect.lLeft
    mp_audtActiveReversibleFrames(lARFUBound).lTop = udtRect.lTop
    mp_audtActiveReversibleFrames(lARFUBound).lRight = udtRect.lRight
    mp_audtActiveReversibleFrames(lARFUBound).lBottom = udtRect.lBottom
    RetVal = SelectObject(mp_lHdc, HoldPen)
    RetVal = DeleteObject(hPen)
End Sub

Public Sub EraseReversibleFrames()
    Dim HoldPen As Long
    Dim hPen As Long
    Dim RetVal As Long
    Dim lIndex As Long
    For lIndex = 1 To UBound(mp_audtActiveReversibleFrames, 1)
        hPen = CreatePen(GRE_LINEDRAWSTYLE.LDS_SOLID, 1, ConvertColor(GRE_COLORS.CLR_BLACK))
        HoldPen = SelectObject(mp_lHdc, hPen)
        RetVal = DrawFocusRect(mp_lHdc, mp_audtActiveReversibleFrames(lIndex))
        mp_oControl.f_DrawLine 0, 0
        RetVal = SelectObject(mp_lHdc, HoldPen)
        RetVal = DeleteObject(hPen)
    Next lIndex
    ReDim mp_audtActiveReversibleFrames(0)
End Sub

Public Sub DrawHatch(ByVal v_X1 As Long, ByVal v_Y1 As Long, ByVal v_X2 As Long, ByVal v_Y2 As Long, ByVal v_lColor As OLE_COLOR, ByVal v_lDrawStyle As GRE_BACKGROUNDPATTERN, ByVal v_iHatchFactor As Integer)
    Dim tmp As Single
    Dim c As Integer
    Dim c1 As Integer
    Dim c2 As Integer
    Dim i1 As Integer
    Dim j1 As Integer
    Dim i2 As Integer
    Dim j2 As Integer
    Dim hPen As Long
    Dim HoldPen
    If v_X1 > v_X2 Then
        tmp = v_X1
        v_X1 = v_X2
        v_X2 = tmp
    End If
    If v_Y1 > v_Y2 Then
        tmp = v_Y1
        v_Y1 = v_Y2
        v_Y2 = tmp
    End If
    If v_lDrawStyle = GRE_BACKGROUNDPATTERN.FP_TRANSPARENT Then
        Exit Sub
    End If
    hPen = CreatePen(LDS_SOLID, 1, v_lColor)
    HoldPen = SelectObject(mp_lHdc, hPen)
    If v_lDrawStyle = GRE_BACKGROUNDPATTERN.FP_HORIZONTALLINE Or v_lDrawStyle = GRE_BACKGROUNDPATTERN.FP_CROSS Then
        For j1 = (v_Y1 + v_iHatchFactor) To v_Y2 Step v_iHatchFactor
            DrawLine v_X1, j1, v_X2, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False
        Next j1
    End If
    If v_lDrawStyle = GRE_BACKGROUNDPATTERN.FP_VERTICALLINE Or v_lDrawStyle = GRE_BACKGROUNDPATTERN.FP_CROSS Then
        For j1 = (v_X1 + v_iHatchFactor) To v_X2 Step v_iHatchFactor
            DrawLine j1, v_Y1, j1, v_Y2, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False
        Next j1
    End If
    If v_lDrawStyle = GRE_BACKGROUNDPATTERN.FP_UPWARDDIAGONAL Or v_lDrawStyle = GRE_BACKGROUNDPATTERN.FP_DIAGONALCROSS Then
        c1 = Int((v_Y1 + v_X1) / v_iHatchFactor + 1)
        c2 = Int((v_Y2 + v_X2) / v_iHatchFactor)
        For c = c1 To c2
            i1 = v_X1
            i2 = v_X2
            j1 = c * v_iHatchFactor - i1
            j2 = c * v_iHatchFactor - i2
            If j2 < v_Y1 Then
                i2 = c * v_iHatchFactor - v_Y1
                j2 = c * v_iHatchFactor - i2
            End If
            If j1 > v_Y2 Then
                i1 = c * v_iHatchFactor - v_Y2
                j1 = c * v_iHatchFactor - i1
            End If
            DrawLine i1, j1, i2, j2, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False
        Next c
    End If
    If v_lDrawStyle = GRE_BACKGROUNDPATTERN.FP_DOWNWARDDIAGONAL Or v_lDrawStyle = GRE_BACKGROUNDPATTERN.FP_DIAGONALCROSS Then
        c1 = Int((v_Y1 - v_X2) / v_iHatchFactor + 1)
        c2 = Int((v_Y2 - v_X1) / v_iHatchFactor)
        For c = c1 To c2
            i1 = v_X1
            i2 = v_X2
            j1 = i1 + c * v_iHatchFactor
            j2 = i2 + c * v_iHatchFactor
            If j1 < v_Y1 Then
                i1 = v_Y1 - c * v_iHatchFactor
                j1 = i1 + c * v_iHatchFactor
            End If
            If j2 > v_Y2 Then
                i2 = v_Y2 - c * v_iHatchFactor
                j2 = i2 + c * v_iHatchFactor
            End If
            DrawLine i1, j1, i2, j2, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False
        Next c
    End If
    If v_lDrawStyle = GRE_BACKGROUNDPATTERN.FP_LIGHT Then
        For j1 = (v_Y1 + 1) To (v_Y2 - 1)
            If j1 Mod 2 = 0 Then
                For j2 = (v_X1 + 1) To (v_X2 - 1) Step 4
                    DrawLine j2, j1, j2 + 1, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False
                Next j2
            Else
                For j2 = (v_X1 + 3) To (v_X2 - 1) Step 4
                    DrawLine j2, j1, j2 + 1, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False
                Next j2
            End If
        Next j1
    End If
    If v_lDrawStyle = GRE_BACKGROUNDPATTERN.FP_MEDIUM Then
        For j1 = (v_Y1 + 1) To (v_Y2 - 1)
            If j1 Mod 2 = 0 Then
                For j2 = (v_X1 + 1) To (v_X2 - 1) Step 2
                    DrawLine j2, j1, j2 + 1, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False
                Next j2
            Else
                For j2 = (v_X1 + 2) To (v_X2 - 1) Step 2
                    DrawLine j2, j1, j2 + 1, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False
                Next j2
            End If
        Next j1
    End If
    If v_lDrawStyle = GRE_BACKGROUNDPATTERN.FP_DARK Then
        For j1 = (v_Y1 + 1) To (v_Y2 - 1)
            If j1 Mod 2 = 0 Then
                For j2 = (v_X1 + 1) To (v_X2 - 1) Step 4
                    If j2 + 3 < v_X2 Then
                        DrawLine j2, j1, j2 + 3, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False
                    Else
                        DrawLine j2, j1, v_X2, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False
                    End If
                Next j2
            Else
                DrawLine v_X1, j1, v_X1 + 2, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False
                For j2 = (v_X1 + 3) To (v_X2 - 1) Step 4
                    If j2 + 3 < v_X2 Then
                        DrawLine j2, j1, j2 + 3, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False
                    Else
                        DrawLine j2, j1, v_X2, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False
                    End If
                Next j2
            End If
        Next j1
    End If
    SelectObject mp_lHdc, HoldPen
    DeleteObject hPen

End Sub

Private Sub mp_DrawPolyLine(ByVal v_lColor As OLE_COLOR, ByVal v_lWidth As Long, ByVal v_lDrawStyle As GRE_LINEDRAWSTYLE, ByRef r_oPoints As GRE_POINTTYPE, ByVal v_Len As Long)
    Dim RetVal As Long
    Dim hPen As Long
    Dim HoldPen
    hPen = CreatePen(v_lDrawStyle, v_lWidth, ConvertColor(v_lColor))
    HoldPen = SelectObject(mp_lHdc, hPen)
    RetVal = Polyline(mp_lHdc, r_oPoints, v_Len)
    RetVal = SelectObject(mp_lHdc, HoldPen)
    RetVal = DeleteObject(hPen)
End Sub

Public Sub DrawTextEx(ByVal v_X1 As Long, ByVal v_Y1 As Long, ByVal v_X2 As Long, ByVal v_Y2 As Long, ByVal v_sParam As String, ByVal v_lFlags As GRE_DRAWTEXTCONSTANTS, ByVal v_lColor As OLE_COLOR, ByVal v_oFont As Font, Optional ByVal v_bClip As Boolean = True)
    Dim udtRect As T_PRECT
    Dim hBrush As Long
    Dim HoldBrush As Long
    Dim holdFont As Long
    Dim oFont As New StdFont
    If RectangleIntersection(v_X1, v_Y1, v_X2, v_Y2) = False Then Exit Sub
    Set oFont = CloneFont(v_oFont)
    udtRect.lLeft = v_X1
    udtRect.lTop = v_Y1
    udtRect.lRight = v_X2
    udtRect.lBottom = v_Y2
    Dim FontI As IFont
    Set FontI = oFont
    holdFont = SelectObject(mp_lHdc, FontI.hFont)
    SetTextColor mp_lHdc, ConvertColor(v_lColor)
    DrawText mp_lHdc, v_sParam, mp_oControl.StrLib.StrLen(v_sParam), udtRect, v_lFlags
    SelectObject mp_lHdc, holdFont
End Sub

Public Sub TextOutEx(ByVal v_lLeft As Long, ByVal v_lTop As Long, ByVal v_lRight As Long, ByVal v_lBottom As Long, ByVal v_sParam As String, ByVal v_yHPos As GRE_HORIZONTALALIGNMENT, ByVal v_yVPos As GRE_VERTICALALIGNMENT, ByVal v_lColor As OLE_COLOR, ByVal v_oFont As Font, Optional ByVal v_bClip As Boolean = True)
    Dim udtRect As T_PRECT
    Dim lFlags As Long
    Dim holdFont As Long
    Dim oFont As New StdFont
    
    If RectangleIntersection(v_lLeft, v_lTop, v_lRight, v_lBottom) = False Then Exit Sub
    If v_bClip = False And mp_oControl.mp_lStrWidth(v_sParam, v_oFont) > (v_lRight - v_lLeft) Then Exit Sub
    If v_bClip = False And mp_oControl.mp_lStrHeight(v_sParam, v_oFont) > (v_lBottom - v_lTop) Then Exit Sub
    Set oFont = CloneFont(v_oFont)
    udtRect.lLeft = v_lLeft
    udtRect.lTop = v_lTop
    udtRect.lRight = v_lRight
    udtRect.lBottom = v_lBottom
    Dim FontI As IFont
    Set FontI = oFont
    holdFont = SelectObject(mp_lHdc, FontI.hFont)
    SetTextColor mp_lHdc, v_lColor
    Select Case v_yHPos
        Case GRE_HORIZONTALALIGNMENT.HAL_LEFT
            lFlags = GRE_DRAWTEXTCONSTANTS.DT_LEFT
        Case GRE_HORIZONTALALIGNMENT.HAL_CENTER
            lFlags = GRE_DRAWTEXTCONSTANTS.DT_CENTER
        Case GRE_HORIZONTALALIGNMENT.HAL_RIGHT
            lFlags = GRE_DRAWTEXTCONSTANTS.DT_RIGHT
    End Select
    Select Case v_yVPos
        Case GRE_VERTICALALIGNMENT.VAL_TOP
            lFlags = lFlags Or GRE_DRAWTEXTCONSTANTS.DT_TOP
        Case GRE_VERTICALALIGNMENT.VAL_CENTER
            lFlags = lFlags Or GRE_DRAWTEXTCONSTANTS.DT_VCENTER
        Case GRE_VERTICALALIGNMENT.VAL_BOTTOM
            lFlags = lFlags Or GRE_DRAWTEXTCONSTANTS.DT_BOTTOM
    End Select
    lFlags = lFlags Or GRE_DRAWTEXTCONSTANTS.DT_SINGLELINE
    DrawText mp_lHdc, v_sParam, mp_oControl.StrLib.StrLen(v_sParam), udtRect, lFlags
    SelectObject mp_lHdc, holdFont
End Sub

Public Sub ClipRegion(ByVal v_X1 As Long, ByVal v_Y1 As Long, ByVal v_X2 As Long, ByVal v_Y2 As Long, ByVal v_bStore As Boolean)
    Dim lRegionHandle As Long
    Dim udtPoints(1) As GRE_POINTTYPE
    v_X2 = v_X2 + 1
    v_Y2 = v_Y2 + 1
    If v_bStore = True Then
        mp_udtPreviousClipRegion.lLeft = v_X1
        mp_udtPreviousClipRegion.lRight = v_X2 - 1
        mp_udtPreviousClipRegion.lTop = v_Y1
        mp_udtPreviousClipRegion.lBottom = v_Y2 - 1
    End If
    If mp_bCustomPrinting = False Then
        udtPoints(0).X = v_X1
        udtPoints(0).Y = v_Y1
        udtPoints(1).X = v_X2
        udtPoints(1).Y = v_Y2
        LPtoDP mp_lHdc, udtPoints(0), 2
        lRegionHandle = CreateRectRgn(udtPoints(0).X, udtPoints(0).Y, udtPoints(1).X, udtPoints(1).Y)
        SelectClipRgn mp_lHdc, lRegionHandle
    Else
        udtPoints(0).X = mp_udtPrintingClipRegion.lLeft
        udtPoints(0).Y = mp_udtPrintingClipRegion.lTop
        udtPoints(1).X = mp_udtPrintingClipRegion.lRight
        udtPoints(1).Y = mp_udtPrintingClipRegion.lBottom
        LPtoDP mp_lHdc, udtPoints(0), 2
        lRegionHandle = CreateRectRgn(udtPoints(0).X, udtPoints(0).Y, udtPoints(1).X, udtPoints(1).Y)
        SelectClipRgn mp_lHdc, lRegionHandle
        IntersectClipRect mp_lHdc, v_X1, v_Y1, v_X2, v_Y2
    End If
    DeleteObject lRegionHandle
End Sub

Public Sub RestorePreviousClipRegion()
    ClipRegion mp_udtPreviousClipRegion.lLeft, mp_udtPreviousClipRegion.lTop, mp_udtPreviousClipRegion.lRight, mp_udtPreviousClipRegion.lBottom, False
End Sub

Public Sub ClearClipRegion()
    Dim lRegionHandle As Long
    lRegionHandle = CreateRectRgn(0, 0, Width, Height)
    SelectClipRgn mp_lHdc, lRegionHandle
    DeleteObject lRegionHandle
End Sub

Public Sub DrawConnectingLines(ByVal v_AX1 As Long, ByVal v_AY1 As Long, ByVal v_AX2 As Long, ByVal v_AY2 As Long, ByVal v_BX1 As Long, ByVal v_BY1 As Long, ByVal v_BX2 As Long, ByVal v_BY2 As Long, ByVal v_yLineStyle As GRE_CONNLINESTYLE, ByVal v_yArrowHeads As GRE_ARROWHEADS, ByVal v_yArrowSize As Integer, ByVal v_lColor As OLE_COLOR, ByVal v_lLineWidth As Long, ByVal v_lLineStyle As GRE_LINEDRAWSTYLE)
    Dim X1 As Long
    Dim X2 As Long
    Dim Y1 As Long
    Dim Y2 As Long
    Dim dy1 As Long
    Dim dy2 As Long
    Dim Points() As GRE_POINTTYPE
    Select Case v_yLineStyle
        Case GRE_CONNLINESTYLE.PDS_NORMAL
            X1 = v_AX2
            Y1 = ((v_AY2 - v_AY1) / 2) + v_AY1
            X2 = v_BX1 + ((v_BY2 - v_BY1) / 2)
            Y2 = ((v_BY2 - v_BY1) / 2) + v_BY1
            If Y1 > Y2 Then
                Y2 = v_BY2
                If v_AX2 <= v_BX1 Then
                    If v_yArrowHeads = GRE_ARROWHEADS.AH_END Or v_yArrowHeads = GRE_ARROWHEADS.AH_BOTH Then
                        DrawArrow X2, Y2, GRE_ARROWDIRECTION.AWD_UP, v_yArrowSize, v_lColor
                    End If
                    ReDim Points(2)
                    Points(0).X = X1
                    Points(0).Y = Y1
                    Points(1).X = X2
                    Points(1).Y = Y1
                    Points(2).X = X2
                    Points(2).Y = Y2
                    mp_DrawPolyLine v_lColor, v_lLineWidth, v_lLineStyle, Points(0), 3
                End If
            ElseIf Y1 < Y2 Then
                Y2 = v_BY1
                If v_AX2 <= v_BX1 Then
                    If v_yArrowHeads = GRE_ARROWHEADS.AH_END Or v_yArrowHeads = GRE_ARROWHEADS.AH_BOTH Then
                        DrawArrow X2, Y2, GRE_ARROWDIRECTION.AWD_DOWN, v_yArrowSize, v_lColor
                    End If
                    ReDim Points(2)
                    Points(0).X = X1
                    Points(0).Y = Y1
                    Points(1).X = X2
                    Points(1).Y = Y1
                    Points(2).X = X2
                    Points(2).Y = Y2
                    mp_DrawPolyLine v_lColor, v_lLineWidth, v_lLineStyle, Points(0), 3
                End If
            ElseIf Y1 = Y2 Then
                X2 = v_BX1
                If v_AX2 <= v_BX1 Then
                    If v_yArrowHeads = AH_END Or v_yArrowHeads = AH_BOTH Then
                        DrawArrow X2, Y2, GRE_ARROWDIRECTION.AWD_RIGHT, v_yArrowSize, v_lColor
                    End If
                    DrawLine X1, Y1, X2, Y2, LT_NORMAL, v_lColor, v_lLineStyle, v_lLineWidth
                End If
            End If
            If v_AX2 <= v_BX1 Then
                If v_yArrowHeads = AH_START Or v_yArrowHeads = AH_BOTH Then
                    DrawArrow X1, Y1, GRE_ARROWDIRECTION.AWD_LEFT, v_yArrowSize, v_lColor
                End If
            End If
        Case GRE_CONNLINESTYLE.PDS_ENHANCED
            X1 = v_AX2
            Y1 = ((v_AY2 - v_AY1) / 2) + v_AY1
            X2 = v_BX1
            Y2 = ((v_BY2 - v_BY1) / 2) + v_BY1
            dy1 = (v_AY2 - v_AY1)
            dy2 = (v_BY2 - v_BY1)
            If (X2 > X1) And (Y1 <> Y2) Then
                ReDim Points(3)
                Points(0).X = X1
                Points(0).Y = Y1
                Points(1).X = X1 + (X2 - X1) / 2
                Points(1).Y = Y1
                Points(2).X = X1 + (X2 - X1) / 2
                Points(2).Y = Y2
                Points(3).X = X2
                Points(3).Y = Y2
                mp_DrawPolyLine v_lColor, v_lLineWidth, v_lLineStyle, Points(0), 4
            ElseIf (X2 <= X1) And (Y1 <> Y2) Then
                ReDim Points(5)
                Points(0).X = X1
                Points(0).Y = Y1
                Points(1).X = X1 + dy1
                Points(1).Y = Y1
                Points(2).X = X1 + dy1
                Points(2).Y = Y1 + Abs((Y2 - Y1) / 2)
                Points(3).X = X2 - dy2
                Points(3).Y = Y1 + Abs((Y2 - Y1) / 2)
                Points(4).X = X2 - dy2
                Points(4).Y = Y2
                Points(5).X = X2
                Points(5).Y = Y2
                mp_DrawPolyLine v_lColor, v_lLineWidth, v_lLineStyle, Points(0), 6
            ElseIf (Y1 = Y2) And (X2 > X1) Then
                DrawLine X1, Y1, X2, Y2, LT_NORMAL, v_lColor, v_lLineStyle, v_lLineWidth
            End If
            If (v_yArrowHeads = GRE_ARROWHEADS.AH_END) Or (v_yArrowHeads = GRE_ARROWHEADS.AH_BOTH) Then
                DrawArrow X2, Y2, GRE_ARROWDIRECTION.AWD_RIGHT, v_yArrowSize, v_lColor
            End If
            If (v_yArrowHeads = GRE_ARROWHEADS.AH_START) Or (v_yArrowHeads = GRE_ARROWHEADS.AH_BOTH) Then
                DrawArrow X1, Y1, GRE_ARROWDIRECTION.AWD_LEFT, v_yArrowSize, v_lColor
            End If
        Case GRE_CONNLINESTYLE.PDS_STRAIGHTLINES
            X1 = v_AX2
            Y1 = ((v_AY2 - v_AY1) / 2) + v_AY1
            X2 = v_BX1
            Y2 = ((v_BY2 - v_BY1) / 2) + v_BY1
            dy1 = (v_AY2 - v_AY1)
            dy2 = (v_BY2 - v_BY1)
            If (X2 > X1) And (Y1 <> Y2) Then
                DrawLine X1, Y1, X2, Y2, LT_NORMAL, v_lColor, v_lLineStyle, v_lLineWidth
            ElseIf (X2 <= X1) And (Y1 <> Y2) Then
                ReDim Points(5)
                Points(0).X = X1
                Points(0).Y = Y1
                Points(1).X = X1 + dy1
                Points(1).Y = Y1
                Points(2).X = X2 - dy2
                Points(2).Y = Y2
                Points(3).X = X2
                Points(3).Y = Y2
                mp_DrawPolyLine v_lColor, v_lLineWidth, v_lLineStyle, Points(0), 4
            ElseIf (Y1 = Y2) And (X2 > X1) Then
                DrawLine X1, Y1, X2, Y2, LT_NORMAL, v_lColor, v_lLineStyle, v_lLineWidth
            End If
            If (v_yArrowHeads = GRE_ARROWHEADS.AH_END) Or (v_yArrowHeads = GRE_ARROWHEADS.AH_BOTH) Then
                DrawArrow X2, Y2, GRE_ARROWDIRECTION.AWD_RIGHT, v_yArrowSize, v_lColor
            End If
            If (v_yArrowHeads = GRE_ARROWHEADS.AH_START) Or (v_yArrowHeads = GRE_ARROWHEADS.AH_BOTH) Then
                DrawArrow X1, Y1, GRE_ARROWDIRECTION.AWD_LEFT, v_yArrowSize, v_lColor
            End If
    End Select
End Sub

Public Function RectangleIntersection(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
    Dim udtRect As T_PRECT
    Dim udtDestRect As T_PRECT
    If mp_bCustomPrinting = False Then
        RectangleIntersection = True
        Exit Function
    End If
    udtRect.lLeft = X1
    udtRect.lTop = Y1
    udtRect.lRight = X2
    udtRect.lBottom = Y2
    If IntersectRect(udtDestRect, udtRect, mp_udtPrintingClipRegion) <> 0 Then
        RectangleIntersection = True
    Else
        RectangleIntersection = False
    End If
End Function

Public Function LineIntersection(ByRef X1 As Long, ByRef Y1 As Long, ByRef X2 As Long, ByRef Y2 As Long) As Boolean
    Dim xIntersect As Boolean
    Dim yIntersect As Boolean
    If mp_bCustomPrinting = False Then
        LineIntersection = True
        Exit Function
    End If
    If (X2 < X1) Then
        Dim X3 As Long
        X3 = X2
        X2 = X1
        X1 = X3
    End If
    If (Y2 < Y1) Then
        Dim Y3 As Long
        Y3 = Y2
        Y2 = Y1
        Y1 = Y3
    End If
    If (X1 >= mp_udtPrintingClipRegion.lLeft And X1 <= mp_udtPrintingClipRegion.lRight) Then
        xIntersect = True
    ElseIf (X2 >= mp_udtPrintingClipRegion.lLeft And X2 <= mp_udtPrintingClipRegion.lRight) Then
        xIntersect = True
    ElseIf (X1 < mp_udtPrintingClipRegion.lLeft And X2 > mp_udtPrintingClipRegion.lRight) Then
        xIntersect = True
    End If
    If (Y1 >= mp_udtPrintingClipRegion.lTop And Y1 <= mp_udtPrintingClipRegion.lBottom) Then
        yIntersect = True
    ElseIf (Y2 >= mp_udtPrintingClipRegion.lTop And Y2 <= mp_udtPrintingClipRegion.lBottom) Then
        yIntersect = True
    ElseIf (Y1 < mp_udtPrintingClipRegion.lTop And Y2 > mp_udtPrintingClipRegion.lBottom) Then
        yIntersect = True
    End If
    If xIntersect = True And yIntersect = True Then
        If X1 < mp_udtPrintingClipRegion.lLeft Then X1 = mp_udtPrintingClipRegion.lLeft
        If X2 > mp_udtPrintingClipRegion.lRight Then X2 = mp_udtPrintingClipRegion.lRight
        If Y1 < mp_udtPrintingClipRegion.lTop Then Y1 = mp_udtPrintingClipRegion.lTop
        If Y2 > mp_udtPrintingClipRegion.lBottom Then Y2 = mp_udtPrintingClipRegion.lBottom
        LineIntersection = True
    Else
        LineIntersection = False
    End If
End Function

Public Sub StartPrintControl(ByVal DestHdc As Long, ByVal XOrigin As Long, ByVal YOrigin As Long, ByVal XOriginExtents As Long, ByVal YOriginExtents As Long, ByVal MarginX As Long, ByVal MarginY As Long, ByVal DestScale As Long)
    Dim oWindowOrgEx As GRE_POINTTYPE
    Dim oWindowExtEx As GRE_POINTTYPE
    Dim oViewPortOrgEx As GRE_POINTTYPE
    Dim oViewPortExtEx As GRE_POINTTYPE
    Dim lMapMode As Long
    mp_lBackupDC = SaveDC(DestHdc)
    lMapMode = SetMapMode(DestHdc, MM_ISOTROPIC)
    SetWindowOrgEx DestHdc, -MarginX + XOrigin, -MarginY + YOrigin, oWindowOrgEx
    SetWindowExtEx DestHdc, 100, 100, oWindowExtEx
    SetViewportOrgEx DestHdc, 0, 0, oViewPortOrgEx
    SetViewportExtEx DestHdc, DestScale, DestScale, oViewPortExtEx
    mp_udtPrintingClipRegion.lLeft = XOrigin
    mp_udtPrintingClipRegion.lTop = YOrigin
    mp_udtPrintingClipRegion.lRight = XOrigin + XOriginExtents
    mp_udtPrintingClipRegion.lBottom = YOrigin + YOriginExtents
    mp_lCustomDC = DestHdc
    ClipRegion mp_udtPrintingClipRegion.lLeft, mp_udtPrintingClipRegion.lTop, mp_udtPrintingClipRegion.lRight, mp_udtPrintingClipRegion.lBottom, False
End Sub

Public Sub EndPrintControl()
    RestoreDC mp_lCustomDC, mp_lBackupDC
End Sub

Public Sub TilePictureHorizontal(ByVal PictureHandle As Long, ByVal v_X1 As Long, ByVal v_Y1 As Long, ByVal v_X2 As Long, ByVal v_Y2 As Long, ByVal v_bTransparent As Boolean)
    Dim oPic As New clsBitmap
    Dim X As Long
    Dim lPictureWidth As Long
    Dim lPictureHeight As Long
    oPic.Initialize mp_oControl
    SetTextColor mp_lHdc, RGB(0, 0, 0)
    oPic.SetBitmap PictureHandle
    lPictureWidth = oPic.Width
    lPictureHeight = oPic.Height
    Do While X < (v_X2 - v_X1)
        If (X + lPictureWidth) > (v_X2 - v_X1) Then
            oPic.PaintPicture mp_lHdc, v_X1 + X, v_Y1, v_bTransparent, v_X2 - (X + v_X1), lPictureHeight, 0, 0
        Else
            oPic.PaintPicture mp_lHdc, v_X1 + X, v_Y1, v_bTransparent, lPictureWidth, lPictureHeight, 0, 0
        End If
        X = X + lPictureWidth
    Loop
    Set oPic = Nothing
End Sub

Public Sub PaintPicture(ByVal PictureHandle As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal v_lXOrigin As Long, ByVal v_lYOrigin As Long, ByVal v_bTransparent As Boolean)
    Dim oPic As New clsBitmap
    oPic.Initialize mp_oControl
    SetTextColor mp_lHdc, RGB(0, 0, 0)
    oPic.SetBitmap PictureHandle
    oPic.PaintPicture mp_lHdc, X1, Y1, v_bTransparent, X2 - X1, Y2 - Y1, v_lXOrigin, v_lYOrigin
    Set oPic = Nothing
End Sub

Public Sub DrawPicture(ByRef v_oPicture As StdPicture, ByRef v_yHorizontalAlignment As GRE_HORIZONTALALIGNMENT, ByRef v_yVerticalAlignment As GRE_VERTICALALIGNMENT, ByVal v_lPictureXMargin As Long, ByVal v_lPictureYMargin As Long, ByRef v_lLeft As Long, ByRef v_lRight As Long, ByRef v_lTop As Long, ByRef v_lBottom As Long, ByVal v_bTransparent As Boolean)
    Dim oPicture As New clsBitmap
    oPicture.Initialize mp_oControl
    Dim bDrawPicture As Boolean
    Dim bHorizontalSmall As Boolean
    Dim bVerticalSmall As Boolean
    Dim XOrigin As Long
    Dim YOrigin As Long
    Dim xDest As Long
    Dim yDest As Long
    Dim lxWidth As Long
    Dim lyHeight As Long
    Dim lPictureHeight As Long
    Dim lPictureWidth As Long
    If v_oPicture Is Nothing Then
        Exit Sub
    End If
    SetTextColor mp_lHdc, RGB(0, 0, 0)
    oPicture.SetBitmap v_oPicture.Handle
    lPictureHeight = oPicture.Height
    lPictureWidth = oPicture.Width
    If v_yHorizontalAlignment = GRE_HORIZONTALALIGNMENT.HAL_CENTER Then
        v_lPictureXMargin = 0
    End If
    If v_yVerticalAlignment = GRE_VERTICALALIGNMENT.VAL_CENTER Then
        v_lPictureYMargin = 0
    End If
    bDrawPicture = True
    If (v_lRight - v_lLeft) < (lPictureWidth + v_lPictureXMargin) Then
        lxWidth = v_lRight - v_lLeft - v_lPictureXMargin
        If lxWidth <= 0 Then bDrawPicture = False
        bHorizontalSmall = True
    Else
        lxWidth = lPictureWidth
        bHorizontalSmall = False
    End If
    If (v_lBottom - v_lTop) < (lPictureHeight + v_lPictureYMargin) Then
        lyHeight = v_lBottom - v_lTop - v_lPictureYMargin
        If lyHeight <= 0 Then bDrawPicture = False
        bVerticalSmall = True
    Else
        lyHeight = lPictureHeight
        bVerticalSmall = False
    End If
    If bHorizontalSmall = False Then
        Select Case v_yHorizontalAlignment
            Case GRE_HORIZONTALALIGNMENT.HAL_LEFT
                xDest = v_lLeft + v_lPictureXMargin
            Case GRE_HORIZONTALALIGNMENT.HAL_CENTER
                xDest = ((v_lRight - v_lLeft) - lPictureWidth) / 2 + v_lLeft
            Case GRE_HORIZONTALALIGNMENT.HAL_RIGHT
                xDest = v_lRight - lPictureWidth - v_lPictureXMargin
        End Select
        XOrigin = 0
    Else
        Select Case v_yHorizontalAlignment
            Case GRE_HORIZONTALALIGNMENT.HAL_LEFT
                XOrigin = 0
                xDest = v_lLeft + v_lPictureXMargin
            Case GRE_HORIZONTALALIGNMENT.HAL_CENTER
                XOrigin = (lPictureWidth - lxWidth) / 2
                xDest = v_lLeft
            Case GRE_HORIZONTALALIGNMENT.HAL_RIGHT
                XOrigin = lPictureWidth - lxWidth
                xDest = v_lRight - lxWidth - v_lPictureXMargin
        End Select
    End If
    If bVerticalSmall = False Then
        Select Case v_yVerticalAlignment
            Case GRE_VERTICALALIGNMENT.VAL_TOP
                yDest = v_lTop + v_lPictureYMargin
            Case GRE_VERTICALALIGNMENT.VAL_CENTER
                yDest = ((v_lBottom - v_lTop) - lPictureHeight) / 2 + v_lTop
            Case GRE_VERTICALALIGNMENT.VAL_BOTTOM
                yDest = v_lBottom - lPictureHeight - v_lPictureYMargin
        End Select
        YOrigin = 0
    Else
        Select Case v_yVerticalAlignment
            Case GRE_VERTICALALIGNMENT.VAL_TOP
                YOrigin = 0
                yDest = v_lTop + v_lPictureYMargin
            Case GRE_VERTICALALIGNMENT.VAL_CENTER
                YOrigin = (lPictureHeight - lyHeight) / 2
                yDest = v_lTop
            Case GRE_VERTICALALIGNMENT.VAL_BOTTOM
                YOrigin = lPictureHeight - lyHeight
                yDest = v_lBottom - lyHeight - v_lPictureYMargin
        End Select
    End If
    If bDrawPicture = True Then
        oPicture.PaintPicture mp_lHdc, xDest, yDest, v_bTransparent, lxWidth, lyHeight, XOrigin, YOrigin
    End If
    Set oPicture = Nothing
End Sub

Public Sub DrawReversibleFrameEx()
    DrawReversibleFrame mp_lFocusLeft, mp_lFocusTop, mp_lFocusRight, mp_lFocusBottom
End Sub

Public Sub ResetFocusRectangle()
    EraseReversibleFrames
    mp_lFocusLeft = 0
    mp_lFocusTop = 0
    mp_lFocusRight = 0
    mp_lFocusBottom = 0
End Sub

Public Sub InvalidateRectangle(ByVal hWnd As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal Width As Long, ByVal Height As Long)
    Dim oRect As T_PRECT
    oRect.lLeft = X1
    oRect.lTop = Y1
    oRect.lRight = X1 + Width
    oRect.lBottom = Y1 + Height
    InvalidateRect hWnd, oRect, 1
End Sub

Private Function CloneFont(Font As IFont) As StdFont
    Font.Clone CloneFont
End Function

Public Sub GradientFill(ByVal v_X1 As Long, ByVal v_Y1 As Long, ByVal v_X2 As Long, ByVal v_Y2 As Long, ByVal clrStartColor As OLE_COLOR, ByVal clrEndColor As OLE_COLOR, ByVal iGradientType As GRE_GRADIENTFILLMODE)
    Dim nWidth As Long: nWidth = v_X2 - v_X1
    Dim nHeight As Long: nHeight = v_Y2 - v_Y1
    Dim lStartR As Long: lStartR = GetRValue(clrStartColor)
    Dim lStartG As Long: lStartG = GetGValue(clrStartColor)
    Dim lStartB As Long: lStartB = GetBValue(clrStartColor)
    Dim lEndR As Long: lEndR = GetRValue(clrEndColor)
    Dim lEndG As Long: lEndG = GetGValue(clrEndColor)
    Dim lEndB As Long: lEndB = GetBValue(clrEndColor)
    Dim lRed As Long
    Dim lGreen As Long
    Dim lBlue As Long
    Dim temp As T_PRECT
    Dim hBrush As Long
    Dim i As Long
    If (iGradientType = GRE_GRADIENTFILLMODE.GDT_VERTICAL) Then
        For i = 0 To nHeight
            lRed = lStartR + (i * (lEndR - lStartR) / nHeight)
            lGreen = lStartG + (i * (lEndG - lStartG) / nHeight)
            lBlue = lStartB + (i * (lEndB - lStartB) / nHeight)
            temp.lLeft = v_X1
            temp.lTop = v_Y1 + i
            temp.lRight = v_X2 + 1
            temp.lBottom = v_Y1 + i + 1
            hBrush = CreateSolidBrush(RGB(lRed, lGreen, lBlue))
            FillRect mp_lHdc, temp, hBrush
            DeleteObject hBrush
        Next i
    ElseIf (iGradientType = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL) Then
        For i = 0 To nWidth
            lRed = lStartR + (i * (lEndR - lStartR) / nWidth)
            lGreen = lStartG + (i * (lEndG - lStartG) / nWidth)
            lBlue = lStartB + (i * (lEndB - lStartB) / nWidth)
            temp.lLeft = v_X1 + i
            temp.lTop = v_Y1
            temp.lRight = v_X1 + i + 1
            temp.lBottom = v_Y2 + 1
            hBrush = CreateSolidBrush(RGB(lRed, lGreen, lBlue))
            FillRect mp_lHdc, temp, hBrush
            DeleteObject hBrush
        Next i
    End If
End Sub

Public Function GetRValue(ByVal Color As OLE_COLOR) As Byte
    GetRValue = Color And &HFF&
End Function

Public Function GetGValue(ByVal Color As OLE_COLOR) As Byte
    GetGValue = (Color And &HFF00&) \ 256
End Function

Public Function GetBValue(ByVal Color As OLE_COLOR) As Byte
    GetBValue = (Color And &HFF0000) \ 65536
End Function






