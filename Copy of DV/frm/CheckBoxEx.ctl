VERSION 5.00
Begin VB.UserControl CheckBoxEx 
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   720
   ScaleWidth      =   4800
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "CheckBoxEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mp_bChecked As Boolean
Private mp_sCaption As String

Private Sub Check1_Click()
    If Check1.Value = 0 Then
        mp_bChecked = False
    Else
        mp_bChecked = True
    End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "DisplayName" Then
        Check1.Caption = Right$(UserControl.Ambient.DisplayName, Len(UserControl.Ambient.DisplayName) - 3)
    End If
End Sub

Private Sub UserControl_InitProperties()
    mp_bChecked = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mp_bChecked = PropBag.ReadProperty("Checked", False)
    Check1.Caption = PropBag.ReadProperty("Caption", "")
End Sub

Private Sub UserControl_Resize()
    Check1.Left = 0
    Check1.Top = 0
    Check1.Width = UserControl.Width
    Check1.Height = UserControl.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Checked", mp_bChecked, False
    PropBag.WriteProperty "Caption", Check1.Caption, ""
End Sub

Public Property Get Checked() As Boolean
    Checked = mp_bChecked
End Property

Public Property Let Checked(ByVal Value As Boolean)
    mp_bChecked = Value
    If mp_bChecked = False Then
        Check1.Value = 0
    Else
        Check1.Value = 1
    End If
    PropertyChanged "Checked"
End Property

