VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Source Code Store - ActiveGantt Schedule Control - Main Screen"
   ClientHeight    =   7665
   ClientLeft      =   1650
   ClientTop       =   1935
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "fMain.frx":0000
   ScaleHeight     =   7665
   ScaleWidth      =   12000
   Begin MSComctlLib.ImageList imgTreeView 
      Left            =   1080
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":065E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":09B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgBanner 
      Left            =   120
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   784
      ImageHeight     =   61
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":0D06
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":23DCA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBanner 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      Picture         =   "fMain.frx":46E8E
      ScaleHeight     =   975
      ScaleWidth      =   11775
      TabIndex        =   3
      Top             =   0
      Width           =   11775
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   10680
      TabIndex        =   1
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Frame fraForm 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   11775
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   5535
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   9763
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgTreeView"
         Appearance      =   1
      End
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      BackColor       =   &H00A37B7C&
      Caption         =   "Copyright ©2002-2004 The Source Code Store. All Rights Reserved. All trademarks are property of their legal owner."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   7320
      Width           =   10455
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ----------------------------------------------------------------------------------------
'//                              COPYRIGHT NOTICE
'// ----------------------------------------------------------------------------------------
'//
'// The Source Code Store ACTIVEGANTT COMPONENT FOR VISUAL BASIC 6.0. VERSION 2
'// Copyright (c) 2002-2004, Julio Luzardo.
'//
'// All Rights Reserved. No parts of this file may be reproduced or transmitted in any
'// form or by any means without the written permission of the author.
'// ----------------------------------------------------------------------------------------

Option Explicit

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    'App.HelpFile = "..\agvb20.chm"
    lblCopyright.BackColor = &HA37B7C
    Set picBanner.Picture = imgBanner.ListImages.Item(1).Picture
    
    TreeView1.Nodes.Add , , "Project", "Project Management Examples:", 1, 2
    TreeView1.Nodes.Add "Project", tvwChild, "Project01", "Graphical Project Management Example (Uses Styles with Imagelists)", 3
    TreeView1.Nodes.Add "Project", tvwChild, "Project02", "Drawn Project Management Example (Uses Styles with Hatch Patterns and Predecessor Objects)", 3
    TreeView1.Nodes.Add "Project", tvwChild, "Project03", "Treeview Example (Uses a user-drawn treeview on the left hand side to show or hide rows)", 3
    TreeView1.Nodes.Add "Project", tvwChild, "Project04", "Percentage Complete Example (Uses Percentages and Percentage Groups to track project progress)", 3
    TreeView1.Nodes.Add "Project", tvwChild, "Project05", "Layers Examples (Uses Layers to show or hide project detail)", 3
    TreeView1.Nodes.Add "Project", tvwChild, "Project06", "ScrollBar Example (Uses the ScrollBar to limit time shown)", 3
    TreeView1.Nodes.Add "Project", tvwChild, "Project07", "Views Example (Uses views to limit the amount of time shown in the TimeLine)", 3
    TreeView1.Nodes.Add "Project", tvwChild, "Project08", "Database Example (Uses an Access 2000 database to populate the schedule)", 3
    TreeView1.Nodes.Add "Project", tvwChild, "Project09", "Tooltips Example (Demonstrates the use of custom tooltips)", 3
    TreeView1.Nodes.Add , , "TV", "TV Scheduling Examples:", 1, 2
    TreeView1.Nodes.Add "TV", tvwChild, "TV01", "TV Scheduling Example (Demonstrates futher use of ImageLists within styles and custom drawing of Tiers)", 3
    TreeView1.Nodes.Add "TV", tvwChild, "TV02", "Database TV Scheduling Example (Uses an Access 2000 database to populate the schedule)", 3
    TreeView1.Nodes.Add , , "Functional", "Functional Examples:", 1, 2
    TreeView1.Nodes.Add "Functional", tvwChild, "Functional01", "Zooming with Views (Uses views to demonstrate zooming in and out)", 3
    TreeView1.Nodes.Add "Functional", tvwChild, "Functional02", "Adding Items to a Collection (Demonstrates the fastest way to add items to a collection)", 3
    TreeView1.Nodes.Add "Functional", tvwChild, "Functional03", "Popup Menus (Demonstrates a way to display popup menus using events and the CancelUIOperations method)", 3
    TreeView1.Nodes.Item("Project").Expanded = True
    TreeView1.Nodes.Item("Project").Bold = True
    TreeView1.Nodes.Item("TV").Expanded = True
    TreeView1.Nodes.Item("TV").Bold = True
    TreeView1.Nodes.Item("Functional").Expanded = True
    TreeView1.Nodes.Item("Functional").Bold = True
End Sub

Private Sub TreeView1_DblClick()
    Select Case TreeView1.SelectedItem.Key
        Case "Project01"
            fProjectManagement01.Show 1, Me
        Case "Project02"
            fProjectManagement02.Show 1, Me
        Case "Project03"
            fProjectManagement03.Show 1, Me
        Case "Project09"
            fProjectManagement09.Show 1, Me
        Case "TV01"
            fTVScheduling01.Show 1, Me
    End Select
End Sub

Private Sub cmdExit_Click()
    End
End Sub
