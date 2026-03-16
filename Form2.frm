VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Password"
   ClientHeight    =   615
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   615
   ScaleWidth      =   4020
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox matkhau 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Maät Khaåu"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
        Case 27:   Unload Me
 End Select
End Sub

Private Sub matkhau_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        If matkhau.Text = "123@saoviet123" Then
                Form1.Show
                Unload Me
                Else
                matkhau.Text = ""
                MsgBox "Mat Khau Khong Dung", vbOK, "Thong bao"
        End If
  End If
End Sub
