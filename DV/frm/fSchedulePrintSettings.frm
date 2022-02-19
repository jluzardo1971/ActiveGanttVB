VERSION 5.00
Begin VB.Form fSchedulePrintSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Schedule Print Settings"
   ClientHeight    =   2025
   ClientLeft      =   1890
   ClientTop       =   1935
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame fraPrintSettings 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtStartDate 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtEndDate 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblHeight 
         Caption         =   "Height:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblStartDate 
         Caption         =   "StartDate:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblEndDate 
         Caption         =   "End Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
   End
End
Attribute VB_Name = "fSchedulePrintSettings"
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

Public mp_Parent As fPrintDialog

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mp_Parent.mp_oControl.Printer.Initialize txtStartDate.Text, txtEndDate.Text, txtHeight.Text
    fPrintDialog.txtEndPage = fPrintDialog.TotalPages
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    txtStartDate.Text = mp_Parent.mp_oControl.Printer.PrintAreaStartDate
    txtEndDate.Text = mp_Parent.mp_oControl.Printer.PrintAreaEndDate
    txtHeight.Text = -1
End Sub
