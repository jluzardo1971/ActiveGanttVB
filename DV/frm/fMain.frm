VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A4F6894D-7F88-4359-BACE-61F7DE949168}#1.0#0"; "XTTreeviewVB.ocx"
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
      Begin XTTreeviewVB.XTTreeviewVBCtl XTTreeviewVBCtl1 
         Height          =   5775
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   10186
      End
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      BackColor       =   &H00A37B7C&
      Caption         =   "Copyright ©2002-2004 The Source Code Store LLC. All Rights Reserved. All trademarks are property of their legal owner."
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
'// The Source Code Store LLC
'// ACTIVEGANTT SCHEDULER COMPONENT FOR VISUAL BASIC 6
'// ACTIVEX COMPONENT
'// Copyright (c) 2002-2004 The Source Code Store LLC
'//
'// All Rights Reserved. No parts of this file may be reproduced or transmitted in any
'// form or by any means without the written permission of the author.
'// ----------------------------------------------------------------------------------------
Option Explicit

Private mp_sParent As String

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    lblCopyright.BackColor = &HA37B7C
    Set picBanner.Picture = imgBanner.ListImages.Item(1).Picture
    XTTreeviewVBCtl1.DefaultValues.NodeHeight = 15
    XTTreeviewVBCtl1.FullRowSelect = True
    
    AddHeading "Project", "Project Management Examples:"
    AddItem "Project01", "Graphical Project Management Example (Demonstrates the use of the TaskStyle class)"
    AddItem "Project02", "Drawn Project Management Example (Uses Styles with Hatch Patterns and Predecessor Objects)"
    AddItem "Project03", "Treeview Example (Uses a XTTreeview control on the left hand side to show or hide rows)"
    AddItem "Project04", "Percentage Complete Example (Uses Percentages and Percentage Groups to track project progress)"
    AddItem "Project05", "Layers Examples (Uses Layers to show or hide project detail)"
    AddItem "Project06", "Setting a Time Limit using the TimeLine ScrollBar (Uses the TimeLine ScrollBar to limit time shown)"
    AddItem "Project07", "Setting a Time Limit using Views (Like having read/write TimeLine.StartDate and TimeLine.EndDate properties)"
    AddItem "Project08", "Database Example (Uses an Access 2000 database and XML to maintain a Repository of schedules)"
    AddItem "Project09", "Tooltips Example (Demonstrates the use of custom tooltips)"
    
    AddHeading "TV", "TV Scheduling Examples:"
    AddItem "TV01", "TV Scheduling Example (Demonstrates futher use of ImageLists within styles and custom drawing of Tiers)"
    AddItem "TV02", "Database TV Scheduling Example (Uses an Access 2000 database to populate the schedule)"
    
    AddHeading "Functional", "Functional Examples:"
    AddItem "Functional01", "Zooming with Views (Uses views to demonstrate zooming in and out)"
    AddItem "Functional02", "Adding Items to a Collection (Demonstrates the fastest way to add items to a collection)"
    AddItem "Functional03", "Popup Menus (Demonstrates a way to display popup menus using events and the CancelUIOperations method)"
    
    XTTreeviewVBCtl1.Redraw

End Sub

Private Sub AddHeading(ByVal Key As String, ByVal Caption As String)
    XTTreeviewVBCtl1.Nodes.Add "", E_RELATIONSHIP.RS_CHILD, Key, Caption
    Set XTTreeviewVBCtl1.Nodes.Item(mp_sNodeIndex).Image = imgTreeView.ListImages.Item(1).Picture
    Set XTTreeviewVBCtl1.Nodes.Item(mp_sNodeIndex).ExpandedImage = imgTreeView.ListImages.Item(2).Picture
    XTTreeviewVBCtl1.Nodes.Item(mp_sNodeIndex).Font.Bold = True
    mp_sParent = Key
End Sub

Private Sub AddItem(ByVal Key As String, ByVal Caption As String)
    XTTreeviewVBCtl1.Nodes.Add mp_sParent, E_RELATIONSHIP.RS_CHILD, Key, Caption
    Set XTTreeviewVBCtl1.Nodes.Item(mp_sNodeIndex).Image = imgTreeView.ListImages.Item(3).Picture
End Sub

Private Function mp_sNodeIndex() As String
    mp_sNodeIndex = XTTreeviewVBCtl1.Nodes.Count
End Function

Private Sub XTTreeviewVBCtl1_DblClick()
    Select Case XTTreeviewVBCtl1.SelectedItem.Key
        Case "Project01"
            fProjectManagement01.Show 1, Me
        Case "Project02"
            fProjectManagement02.Show 1, Me
        Case "Project03"
            fProjectManagement03.Show 1, Me
        Case "Project04"
            fProjectManagement04.Show 1, Me
        Case "Project05"
            fProjectManagement05.Show 1, Me
        Case "Project06"
            fProjectManagement06.Show 1, Me
        Case "Project07"
            fProjectManagement07.Show 1, Me
        Case "Project08"
            fProjectManagement08.Show 1, Me
        Case "Project09"
            fProjectManagement09.Show 1, Me
        Case "TV01"
            fTVScheduling01.Show 1, Me
        Case "TV02"
            fTVScheduling02.Show 1, Me
        Case "Functional01"
            fFunctional01.Show 1, Me
        Case "Functional02"
            fFunctional02.Show 1, Me
        Case "Functional03"
            fFunctional03.Show 1, Me
    End Select
End Sub

Private Sub cmdExit_Click()
    End
End Sub

