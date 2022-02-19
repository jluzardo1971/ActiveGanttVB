VERSION 5.00
Begin VB.Form fProjectManagement08B 
   Caption         =   "Add Row"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox chkContainer 
         Caption         =   "Container"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CheckBox chkTitle 
         Caption         =   "Title"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Caption:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "fProjectManagement08B"
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


Private Sub cmdCancel_Click()
    fProjectManagement08A.bRowOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    fProjectManagement08A.bRowOK = True
    fProjectManagement08A.sRowCaption = txtCaption.Text
    fProjectManagement08A.bRowContainer = chkContainer.Value
    fProjectManagement08A.bRowTitle = chkTitle.Value
    Unload Me
End Sub

Private Sub Form_Load()
    fProjectManagement08A.bRowOK = False
    txtCaption.Text = fProjectManagement08A.sRowCaption
    chkContainer.Value = Abs(CInt(fProjectManagement08A.bRowContainer))
    chkTitle.Value = Abs(CInt(fProjectManagement08A.bRowTitle))
End Sub
