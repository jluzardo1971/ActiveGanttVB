VERSION 5.00
Begin VB.Form fAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About ActiveGantt Scheduler Component"
   ClientHeight    =   3210
   ClientLeft      =   1830
   ClientTop       =   1935
   ClientWidth     =   5625
   Icon            =   "fAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "fAbout.frx":0442
   ScaleHeight     =   3210
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Frame fraForm 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         Picture         =   "fAbout.frx":074C
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblRegister 
         Caption         =   "secure online order form"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label lblRegisterCaption 
         Caption         =   "Buy Now:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblURL 
         Alignment       =   2  'Center
         Caption         =   "http://www.sourcecodestore.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   5055
      End
      Begin VB.Label lblSalesCaption 
         Caption         =   "Sales:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblTechnicalSupportCaption 
         Caption         =   "Technical Support:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblSales 
         Caption         =   "sales@sourcecodestore.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label lblTechnicalSupport 
         Caption         =   "support@sourcecodestore.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   2040
         Width           =   3255
      End
      Begin VB.Label lblTitle2 
         Caption         =   "Visual Basic 6.0 ActiveX Control"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblTitle3 
         Caption         =   "Copyright ©2002-2004 The Source Code Store LLC"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label lblTitle1 
         Caption         =   "ActiveGanttVB Scheduler Component, Version 1.1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   4575
      End
   End
End
Attribute VB_Name = "fAbout"
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

'// Parent Control Pointer
Private mp_oControl As ActiveGanttVBCtl
'// Declares
Private Const SW_SHOW = 5
Private Const SW_SHOWNORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

'// ---------------------------------------------------------------------------------------------------------------------
'// Initialization & Destruction
'// ---------------------------------------------------------------------------------------------------------------------

Friend Sub Initialize(ByRef Value As ActiveGanttVBCtl)
    Set mp_oControl = Value
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    lblTitle1.Caption = "ActiveGanttVB Scheduler Component, Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblURL.Caption = "http://www.sourcecodestore.com"
    lblTechnicalSupport.Caption = "Technical Support Page"
    lblSales.Caption = "sales@sourcecodestore.com"
    lblRegister.Tag = "http://www.sourcecodestore.com/onlinestore.htm"
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = 0
End Sub

Private Sub fraForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = 0
End Sub

Private Sub lblRegister_Click()
    mp_ShowInBrowser lblRegister.Tag
End Sub

Private Sub lblRegister_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = 99
End Sub

Private Sub lblSales_Click()
    ShellExecute Me.hWnd, "Open", "mailto:" & lblSales.Caption, vbNullString, App.Path, vbNormalFocus
End Sub

Private Sub lblSales_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = 99
End Sub

Private Sub lblTechnicalSupport_Click()
    mp_ShowInBrowser "http://www.sourcecodestore.com/support.htm"
End Sub

Private Sub lblTechnicalSupport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = 99
End Sub

Private Sub lblURL_Click()
    mp_ShowInBrowser lblURL.Caption
End Sub

Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = 99
End Sub

Private Sub mp_ShowInBrowser(ByVal sFileName As String)
    Dim FileName As String
    Dim Dummy As String
    Dim BrowserExec As String * 255
    Dim RetVal As Long
    Dim FileNumber As Integer
    BrowserExec = Space(255)
    FileName = "C:\temphtm.HTM"
    FileNumber = FreeFile
    Open FileName For Output As #FileNumber
        Write #FileNumber, "<HTML> <\HTML>"
    Close #FileNumber
    RetVal = FindExecutable(FileName, Dummy, BrowserExec)
    BrowserExec = mp_oControl.StrLib.StrTrim(BrowserExec)
    If RetVal <= 32 Or IsEmpty(BrowserExec) Then
        MsgBox "Could not find associated Browser", vbExclamation, "Browser Not Found"
    Else
        RetVal = ShellExecute(Me.hWnd, "open", BrowserExec, _
          sFileName, Dummy, SW_SHOWNORMAL)
        If RetVal <= 32 Then
            MsgBox "Web Page not Opened", vbExclamation, "URL Failed"
        End If
    End If
    Kill FileName
End Sub
