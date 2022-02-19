VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CCC1C7D1-F592-4261-9A57-CA48F869B175}#1.0#0"; "ActiveGanttVB2.ocx"
Begin VB.Form fTVScheduling01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Source Code Store - ActiveGantt Scheduler Control - TV Schedule Example"
   ClientHeight    =   8490
   ClientLeft      =   3240
   ClientTop       =   1995
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraForm 
      Height          =   8175
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   11775
      Begin ActiveGanttVB.ActiveGanttVBCtl ActiveGanttVBCtl1 
         Height          =   7815
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   13785
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList imgESPN 
         Left            =   120
         Top             =   2400
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   35
         ImageHeight     =   27
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":0BB8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgDW 
         Left            =   120
         Top             =   1680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   28
         ImageHeight     =   27
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":1770
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":20A0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgMSNBC 
         Left            =   120
         Top             =   960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   51
         ImageHeight     =   27
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":29D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":3A98
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":4B60
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgChannels 
         Left            =   120
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   150
         ImageHeight     =   27
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   15
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":5C28
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":8C28
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":BC28
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":EC28
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":11C28
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":14C28
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":17C28
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":1AC28
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":1DC28
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":20C28
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":23C28
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":26C28
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":29C28
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":2CC28
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTVScheduling01.frx":2FC28
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Copyright ©2002 The Source Code Store. All trademarks are property of their legal owner."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8160
      Visible         =   0   'False
      Width           =   12015
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuPrint 
         Caption         =   "Print..."
      End
      Begin VB.Menu mnuLine010 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "fTVScheduling01"
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

Private Enum E_FONTTYPES
    FNT_NORMAL = 0
    FNT_BOLD = 1
    FNT_BOLD9 = 2
    FNT_BOLD11 = 3
End Enum

Private Function NewDate(ByVal Month As Integer, ByVal Day As Integer, ByVal Year As Integer, ByVal Hour As Integer, ByVal Minute As Integer)
    NewDate = DateSerial(Year, Month, Day) + TimeSerial(Hour, Minute, 0)
End Function

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    With ActiveGanttVBCtl1
        Me.Caption = "The Source Code Store - ActiveGantt Scheduler Control Version " & .Version & " - TV Schedule Example"
        .DefaultValues.RowHeight = 30
        With .Styles
            .Add "TVStations"
            With .Item("TVStations")
                .Appearance = E_STYLEAPPEARANCE.SA_CELL
                .BackColor = mp_ConvertColor(&H80000005)
            End With
            .Add "ABC"
            With .Item("ABC")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .BorderStyle = E_STYLEBORDER.SBR_SINGLE
                .BackColor = mp_ConvertColor(RGB(255, 204, 0))
                mp_ConvertFont .Font, E_FONTTYPES.FNT_BOLD
            End With
            .Add "NBC"
            With .Item("NBC")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .BackgroundPattern = GRE_BACKGROUNDPATTERN.FP_GRADIENT
                .GradientFillMode = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL
                .StartGradientColor = mp_ConvertColor(RGB(255, 142, 0))
                .EndGradientColor = mp_ConvertColor(RGB(206, 8, 8))
                .BorderStyle = E_STYLEBORDER.SBR_SINGLE
                .ForeColor = mp_ConvertColor(&H80000005)
                mp_ConvertFont .Font, E_FONTTYPES.FNT_BOLD
            End With
            .Add "MSNBC"
            With .Item("MSNBC")
                .UseMask = False
                .Appearance = E_STYLEAPPEARANCE.SA_GRAPHICAL
                Set .TaskStyle.StartPicture = imgMSNBC.ListImages.Item(1).Picture
                Set .TaskStyle.MiddlePicture = imgMSNBC.ListImages.Item(2).Picture
                Set .TaskStyle.EndPicture = imgMSNBC.ListImages.Item(3).Picture
                .ForeColor = mp_ConvertColor(&H80000005)
                mp_ConvertFont .Font, E_FONTTYPES.FNT_BOLD
            End With
            .Add "AE"
            With .Item("AE")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .BackgroundPattern = GRE_BACKGROUNDPATTERN.FP_GRADIENT
                .GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
                .BorderStyle = E_STYLEBORDER.SBR_SINGLE
                .StartGradientColor = mp_ConvertColor(RGB(13, 7, 51))
                .EndGradientColor = mp_ConvertColor(RGB(66, 44, 41))
                .ForeColor = mp_ConvertColor(RGB(255, 204, 0))
                .BorderColor = mp_ConvertColor(RGB(255, 204, 0))
                mp_ConvertFont .Font, E_FONTTYPES.FNT_BOLD
            End With
            .Add "HBO"
            With .Item("HBO")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .BackgroundPattern = GRE_BACKGROUNDPATTERN.FP_GRADIENT
                .GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
                .BorderStyle = E_STYLEBORDER.SBR_SINGLE
                .StartGradientColor = mp_ConvertColor(RGB(156, 174, 214))
                .EndGradientColor = mp_ConvertColor(RGB(198, 207, 231))
                .BackColor = mp_ConvertColor(RGB(153, 172, 215))
                .ForeColor = mp_ConvertColor(&H80000005)
                mp_ConvertFont .Font, E_FONTTYPES.FNT_BOLD
            End With
            .Add "CNN"
            With .Item("CNN")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .BorderStyle = E_STYLEBORDER.SBR_SINGLE
                .BackColor = mp_ConvertColor(RGB(206, 5, 2))
                .ForeColor = mp_ConvertColor(&H80000005)
                mp_ConvertFont .Font, E_FONTTYPES.FNT_BOLD
            End With
            .Add "CNMXRAI"
            With .Item("CNMXRAI")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .BorderStyle = E_STYLEBORDER.SBR_SINGLE
                .BackColor = mp_ConvertColor(&H80000005)
                mp_ConvertFont .Font, E_FONTTYPES.FNT_BOLD
            End With
            .Add "DW"
            With .Item("DW")
                .UseMask = False
                .Appearance = E_STYLEAPPEARANCE.SA_GRAPHICAL
                Set .TaskStyle.StartPicture = imgDW.ListImages.Item(1).Picture
                Set .TaskStyle.MiddlePicture = imgDW.ListImages.Item(1).Picture
                Set .TaskStyle.EndPicture = imgDW.ListImages.Item(2).Picture
                .ForeColor = mp_ConvertColor(&H80000005)
                mp_ConvertFont .Font, E_FONTTYPES.FNT_BOLD9
            End With
            .Add "ESPN"
            With .Item("ESPN")
                .UseMask = False
                .Appearance = E_STYLEAPPEARANCE.SA_GRAPHICAL
                Set .TaskStyle.StartPicture = imgESPN.ListImages.Item(1).Picture
                Set .TaskStyle.MiddlePicture = imgESPN.ListImages.Item(1).Picture
                Set .TaskStyle.EndPicture = imgESPN.ListImages.Item(2).Picture
                mp_ConvertFont .Font, E_FONTTYPES.FNT_BOLD9
            End With
            .Add "DISC"
            With .Item("DISC")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .BorderStyle = E_STYLEBORDER.SBR_SINGLE
                .BackColor = mp_ConvertColor(RGB(49, 68, 139))
                .ForeColor = mp_ConvertColor(&H80000005)
                mp_ConvertFont .Font, E_FONTTYPES.FNT_BOLD
            End With
            .Add "CART"
            With .Item("CART")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .BackgroundPattern = GRE_BACKGROUNDPATTERN.FP_GRADIENT
                .GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
                .BorderStyle = E_STYLEBORDER.SBR_SINGLE
                .StartGradientColor = mp_ConvertColor(RGB(0, 0, 0))
                .EndGradientColor = mp_ConvertColor(RGB(60, 60, 60))
                .BorderColor = mp_ConvertColor(&H80000005)
                .ForeColor = mp_ConvertColor(&H80000005)
                mp_ConvertFont .Font, E_FONTTYPES.FNT_BOLD
            End With
            .Add "CBS"
            With .Item("CBS")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .BorderStyle = E_STYLEBORDER.SBR_SINGLE
                .BackColor = mp_ConvertColor(RGB(4, 54, 155))
                .ForeColor = mp_ConvertColor(RGB(255, 204, 0))
                mp_ConvertFont .Font, E_FONTTYPES.FNT_BOLD
            End With
            .Add "BBC"
            With .Item("BBC")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .BorderStyle = E_STYLEBORDER.SBR_SINGLE
                .BackColor = mp_ConvertColor(RGB(153, 153, 153))
                .ForeColor = mp_ConvertColor(&H80000005)
                mp_ConvertFont .Font, E_FONTTYPES.FNT_BOLD
            End With
            .Add "WEATH"
            With .Item("WEATH")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .BorderStyle = E_STYLEBORDER.SBR_SINGLE
                .BackColor = mp_ConvertColor(RGB(116, 140, 188))
                .ForeColor = mp_ConvertColor(&H80000005)
                mp_ConvertFont .Font, E_FONTTYPES.FNT_BOLD
            End With
            .Add "TimeLineStyle"
            With .Item("TimeLineStyle")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .BorderStyle = E_STYLEBORDER.SBR_SINGLE
                .BorderColor = mp_ConvertColor(RGB(0, 0, 0))
                .BackgroundPattern = GRE_BACKGROUNDPATTERN.FP_GRADIENT
                .GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
                .StartGradientColor = mp_ConvertColor(RGB(214, 16, 8))
                .EndGradientColor = mp_ConvertColor(RGB(255, 142, 0))
                mp_ConvertFont .Font, E_FONTTYPES.FNT_BOLD11
                .CaptionAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM
                .ForeColor = mp_ConvertColor(RGB(255, 255, 255))
            End With
            .Add "LowerTierStyle"
            With .Item("LowerTierStyle")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .BorderStyle = E_STYLEBORDER.SBR_SINGLE
                .BorderColor = mp_ConvertColor(RGB(255, 255, 255))
                .BackgroundPattern = GRE_BACKGROUNDPATTERN.FP_GRADIENT
                .ForeColor = mp_ConvertColor(RGB(255, 255, 255))
                .GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
                .StartGradientColor = mp_ConvertColor(RGB(152, 116, 253))
                .EndGradientColor = mp_ConvertColor(RGB(35, 51, 251))
                .CaptionVisible = True
                .CaptionAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_CENTER
                .ClipCaption = True
                mp_ConvertFont .Font, E_FONTTYPES.FNT_NORMAL
                .DrawCaptionInVisibleArea = True
            End With
            .Add "UpperTierStyle"
            With .Item("UpperTierStyle")
                .Appearance = E_STYLEAPPEARANCE.SA_FLAT
                .BorderStyle = E_STYLEBORDER.SBR_SINGLE
                .BorderColor = mp_ConvertColor(RGB(255, 255, 255))
                .BackgroundPattern = GRE_BACKGROUNDPATTERN.FP_GRADIENT
                .ForeColor = mp_ConvertColor(RGB(0, 0, 0))
                .GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
                .StartGradientColor = mp_ConvertColor(RGB(156, 174, 214))
                .EndGradientColor = mp_ConvertColor(RGB(198, 207, 231))
                .CaptionVisible = True
                .CaptionAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_CENTER
                .ClipCaption = True
                mp_ConvertFont .Font, E_FONTTYPES.FNT_BOLD
                .DrawCaptionInVisibleArea = True
            End With
        End With
        With .Rows
            .Add "ABC", , True, , 1
            Set .Item("ABC").Picture = imgChannels.ListImages.Item(1).Picture
            .Add "NBC", , True, , 1
            Set .Item("NBC").Picture = imgChannels.ListImages.Item(2).Picture
            .Add "MSNBC", , True, , 1
            Set .Item("MSNBC").Picture = imgChannels.ListImages.Item(3).Picture
            .Add "DW", , True, , 1
            Set .Item("DW").Picture = imgChannels.ListImages.Item(4).Picture
            .Add "AE", , True, , 1
            Set .Item("AE").Picture = imgChannels.ListImages.Item(5).Picture
            .Add "HBO", , True, , 1
            Set .Item("HBO").Picture = imgChannels.ListImages.Item(6).Picture
            .Add "ESPN", , True, , 1
            Set .Item("ESPN").Picture = imgChannels.ListImages.Item(7).Picture
            .Add "CNN", , True, , 1
            Set .Item("CNN").Picture = imgChannels.ListImages.Item(8).Picture
            .Add "CNMX", , True, , 1
            Set .Item("CNMX").Picture = imgChannels.ListImages.Item(9).Picture
            .Add "CBS", , True, , 1
            Set .Item("CBS").Picture = imgChannels.ListImages.Item(10).Picture
            .Add "CART", , True, , 1
            Set .Item("CART").Picture = imgChannels.ListImages.Item(11).Picture
            .Add "DISC", , True, , 1
            Set .Item("DISC").Picture = imgChannels.ListImages.Item(12).Picture
            .Add "RAI", , True, , 1
            Set .Item("RAI").Picture = imgChannels.ListImages.Item(13).Picture
            .Add "BBC", , True, , 1
            Set .Item("BBC").Picture = imgChannels.ListImages.Item(14).Picture
            .Add "WEATH", , True, , 1
            Set .Item("WEATH").Picture = imgChannels.ListImages.Item(15).Picture
        End With
        With .Tasks
            .Add "Spin City", "ABC", NewDate(3, 11, 2003, 17, 0), NewDate(3, 11, 2003, 17, 30), , "ABC"
            .Add "Spin City", "ABC", NewDate(3, 11, 2003, 17, 30), NewDate(3, 11, 2003, 18, 0), , "ABC"
            .Add "Boston 24/7", "ABC", NewDate(3, 11, 2003, 18, 0), NewDate(3, 11, 2003, 19, 0), , "ABC"
            .Add "SPY TV", "NBC", NewDate(3, 11, 2003, 17, 0), NewDate(3, 11, 2003, 18, 0), , "NBC"
            .Add "Frasier", "NBC", NewDate(3, 11, 2003, 18, 0), NewDate(3, 11, 2003, 19, 0), , "NBC"
            .Add "Hardball with Chris Matthews", "MSNBC", NewDate(3, 11, 2003, 17, 0), NewDate(3, 11, 2003, 18, 0), , "MSNBC"
            .Add "MSNBC Live", "MSNBC", NewDate(3, 11, 2003, 18, 0), NewDate(3, 11, 2003, 19, 0), , "MSNBC"
            .Add "Deutschland Heute", "DW", NewDate(3, 11, 2003, 17, 0), NewDate(3, 11, 2003, 17, 30), , "DW"
            .Add "Journal", "DW", NewDate(3, 11, 2003, 17, 30), NewDate(3, 11, 2003, 18, 0), , "DW"
            .Add "Close Up", "DW", NewDate(3, 11, 2003, 18, 0), NewDate(3, 11, 2003, 18, 30), , "DW"
            .Add "Reiseland Deutschland", "DW", NewDate(3, 11, 2003, 18, 30), NewDate(3, 11, 2003, 19, 0), , "DW"
            .Add "Any Given Sunday", "HBO", NewDate(3, 11, 2003, 15, 0), NewDate(3, 11, 2003, 18, 15), , "HBO"
            .Add "Silk Hope", "HBO", NewDate(3, 11, 2003, 18, 15), NewDate(3, 11, 2003, 20, 30), , "HBO"
            .Add "FIM Road Racing World", "ESPN", NewDate(3, 11, 2003, 17, 0), NewDate(3, 11, 2003, 18, 0), , "ESPN"
            .Add "Sports Center", "ESPN", NewDate(3, 11, 2003, 18, 0), NewDate(3, 11, 2003, 20, 30), , "ESPN"
            .Add "Newsbiz Today", "CNN", NewDate(3, 11, 2003, 17, 0), NewDate(3, 11, 2003, 17, 30), , "CNN"
            .Add "Larry King Live", "CNN", NewDate(3, 11, 2003, 17, 30), NewDate(3, 11, 2003, 18, 15), , "CNN"
            .Add "World News", "CNN", NewDate(3, 11, 2003, 18, 15), NewDate(3, 11, 2003, 19, 50), , "CNN"
            .Add "Dragon Balls Z", "CART", NewDate(3, 11, 2003, 17, 0), NewDate(3, 11, 2003, 17, 30), , "CART"
            .Add "PowerPuff Girls", "CART", NewDate(3, 11, 2003, 17, 30), NewDate(3, 11, 2003, 18, 0), , "CART"
            .Add "Time Squad", "CART", NewDate(3, 11, 2003, 18, 0), NewDate(3, 11, 2003, 19, 30), , "CART"
            .Add "FBI Files", "DISC", NewDate(3, 11, 2003, 17, 0), NewDate(3, 11, 2003, 17, 45), , "DISC"
            .Add "The New Detectives", "DISC", NewDate(3, 11, 2003, 17, 45), NewDate(3, 11, 2003, 18, 15), , "DISC"
            .Add "Shark Files", "DISC", NewDate(3, 11, 2003, 18, 15), NewDate(3, 11, 2003, 19, 30), , "DISC"
            .Add "Biography", "AE", NewDate(3, 11, 2003, 17, 0), NewDate(3, 11, 2003, 18, 5), , "AE"
            .Add "Law and Order", "AE", NewDate(3, 11, 2003, 18, 5), NewDate(3, 11, 2003, 18, 45), , "AE"
            .Add "The view", "AE", NewDate(3, 11, 2003, 18, 45), NewDate(3, 11, 2003, 19, 30), , "AE"
            .Add "The Early Show", "CBS", NewDate(3, 11, 2003, 17, 0), NewDate(3, 11, 2003, 18, 5), , "CBS"
            .Add "Evening News", "CBS", NewDate(3, 11, 2003, 18, 5), NewDate(3, 11, 2003, 18, 45), , "CBS"
            .Add "60 Minutes", "CBS", NewDate(3, 11, 2003, 18, 45), NewDate(3, 11, 2003, 19, 30), , "CBS"
            .Add "Qui Roma", "RAI", NewDate(3, 11, 2003, 17, 0), NewDate(3, 11, 2003, 17, 40), , "CNMXRAI"
            .Add "C'era una Volta", "RAI", NewDate(3, 11, 2003, 17, 40), NewDate(3, 11, 2003, 20, 30), , "CNMXRAI"
            .Add "World Service", "BBC", NewDate(3, 11, 2003, 16, 0), NewDate(3, 11, 2003, 18, 0), , "BBC"
            .Add "World Service", "BBC", NewDate(3, 11, 2003, 18, 0), NewDate(3, 11, 2003, 20, 0), , "BBC"
            .Add "World Weather", "WEATH", NewDate(3, 11, 2003, 15, 45), NewDate(3, 11, 2003, 17, 45), , "WEATH"
            .Add "World Weather", "WEATH", NewDate(3, 11, 2003, 17, 45), NewDate(3, 11, 2003, 19, 45), , "WEATH"
            .Add "Space Cowboys", "CNMX", NewDate(3, 11, 2003, 15, 0), NewDate(3, 11, 2003, 17, 50), , "CNMXRAI"
            .Add "Love And Basketball", "CNMX", NewDate(3, 11, 2003, 17, 50), NewDate(3, 11, 2003, 20, 30), , "CNMXRAI"
        End With
        .Columns.Add "Channels:", "", 155, "TimeLineStyle"
        .Splitter.Position = 155
        
        .Views.Add "10s", "30n", E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM
        .Views.Item(1).TimeLine.TierArea.UpperTier.Interval = "1m"
        .Views.Item(1).TimeLine.TierArea.UpperTier.Height = 17
        .Views.Item(1).TimeLine.TierArea.LowerTier.Interval = "1d"
        .Views.Item(1).TimeLine.TierArea.LowerTier.Height = 17
        .Views.Item(1).ClientArea.ToolTipsVisible = False
        .Views.Item(1).TimeLine.StyleIndex = "TimeLineStyle"
        .Views.Item(1).TimeLine.Appearance = E_BORDERSTYLE.TLB_SINGLE
        .Views.Item(1).TimeLine.ForeColor = mp_ConvertColor(RGB(255, 255, 255))
        With .Views.Item(1).TimeLine.TickMarkArea
            .TickMarks.Add 60, E_TICKMARKTYPES.TLT_BIG, False, "", True
            .TickMarks.Add 30, E_TICKMARKTYPES.TLT_MEDIUM, True, "Hh:Nnam/pm", True
        End With
        mp_ConvertFont .Views.Item("1").TimeLine.TickMarkArea.Font, E_FONTTYPES.FNT_NORMAL
        .CurrentView = "1"
        .CurrentViewObject.TimeLine.Position NewDate(3, 11, 2003, 17, 0)
        .Redraw
    End With
End Sub

Private Sub ActiveGanttVBCtl1_CustomTierDraw(ByVal Position As ActiveGanttVB.E_TIERPOSITION, ByVal StartDate As Date, ByVal EndDate As Date, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal LeftTrim As Long, ByVal RightTrim As Long, ByVal lHdc As Long, Caption As String, StyleIndex As String)
    If Position = E_TIERPOSITION.SP_LOWER Then
        StyleIndex = "LowerTierStyle"
        Caption = Format(StartDate, "dddd d")
    ElseIf Position = E_TIERPOSITION.SP_UPPER Then
        StyleIndex = "UpperTierStyle"
        Caption = Format(StartDate, "mmmm yyyy")
    End If
End Sub

Private Function mp_ConvertColor(ByVal Color As OLE_COLOR) As OLE_COLOR
'// Function for Simple VB.Net Migrations
    mp_ConvertColor = Color
End Function

Private Sub mp_ConvertFont(ByRef oFont As Font, ByVal oFontType As E_FONTTYPES)
'// Function for Simple VB.Net Migrations
    Select Case oFontType
        Case FNT_NORMAL
            oFont.Name = "Verdana"
        Case FNT_BOLD
            oFont.Name = "Verdana"
            oFont.Bold = True
        Case FNT_BOLD9
            oFont.Name = "Verdana"
            oFont.Bold = True
            oFont.Size = 9
        Case FNT_BOLD11
            oFont.Name = "Verdana"
            oFont.Bold = True
            oFont.Size = 11
    End Select
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuPrint_Click()
    Set fPrintDialog.mp_oControl = ActiveGanttVBCtl1
    fPrintDialog.Show 1, Me
    ActiveGanttVBCtl1.Printer.Terminate
    ActiveGanttVBCtl1.Redraw
End Sub

































































