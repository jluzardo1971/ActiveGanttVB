VERSION 5.00
Begin VB.UserControl ColorPickerEx 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   ScaleHeight     =   3600
   ScaleWidth      =   5040
   Begin VB.CommandButton Command2 
      Caption         =   "ForeColor..."
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdBackColor 
      Caption         =   "BackColor..."
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "ColorPickerEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
