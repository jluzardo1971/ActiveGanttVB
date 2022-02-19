VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fProjectManagement08 
   Caption         =   "Form1"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraForm 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   6840
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   6840
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "AddNew"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   6840
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   11245
         _Version        =   393216
         Cols            =   4
      End
   End
End
Attribute VB_Name = "fProjectManagement08"
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


Private Sub mp_RequeryTable()
    Dim tb_Projects As New ADODB.Recordset
    Dim ProjectRepositoryMDB As New ADODB.Connection
    Dim lIndex As Long
    
    If MSFlexGrid1.Rows > 2 Then
        lIndex = MSFlexGrid1.Rows - 2
        Do While lIndex >= 1
            MSFlexGrid1.RemoveItem lIndex
            lIndex = lIndex - 1
        Loop
    End If

    mp_SetTitle "ID", 0, 0, 10
    mp_SetTitle "Project Description", 0, 1, 50
    mp_SetTitle "Start Date", 0, 2, 20
    mp_SetTitle "End Date", 0, 3, 20
    
    MSFlexGrid1.Row = 1
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Text = "*"
    

    ProjectRepositoryMDB.Open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "\frm\ProjectRepository.mdb;UID=;PWD="
    tb_Projects.CursorType = adOpenStatic
    tb_Projects.LockType = adLockReadOnly
    tb_Projects.Open "SELECT * FROM tb_Projects ORDER BY lProjectID", ProjectRepositoryMDB
    
    Do While tb_Projects.EOF = False
        MSFlexGrid1.AddItem tb_Projects!lProjectID & vbTab & tb_Projects!sProjectDescription & vbTab & tb_Projects!dtStartDate & vbTab & tb_Projects!dtEndDate, MSFlexGrid1.Rows - 1
        tb_Projects.MoveNext
    Loop
    
    tb_Projects.Close
    ProjectRepositoryMDB.Close
End Sub

Private Sub mp_SetTitle(ByVal sCaption As String, ByVal Row As Long, ByVal Col As Long, ByVal lWidth As Long)
    MSFlexGrid1.Row = Row
    MSFlexGrid1.Col = Col
    MSFlexGrid1.Text = sCaption
    MSFlexGrid1.ColWidth(Col) = (MSFlexGrid1.Width - (Screen.TwipsPerPixelX * 6)) * (lWidth / 100)
End Sub

Private Sub cmdAddNew_Click()
    fProjectManagement08A.bAddNew = True
    fProjectManagement08A.Show 1, Me
    mp_RequeryTable
End Sub

Private Sub Form_Load()
    Me.Caption = "The Source Code Store - ActiveGantt Scheduler Control - Project Management Example"
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    mp_RequeryTable
End Sub

Private Sub MSFlexGrid1_DblClick()
    If MSFlexGrid1.Row = MSFlexGrid1.Rows - 1 Then
        cmdAddNew_Click
    Else
        If MSFlexGrid1.Row = 0 Then
            Exit Sub
        End If
        MSFlexGrid1.Col = 0
        fProjectManagement08A.bAddNew = False
        fProjectManagement08A.lProjectID = MSFlexGrid1.Text
        fProjectManagement08A.Show 1, Me
    End If
End Sub
