VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "KEY VIETSTAR"
   ClientHeight    =   1815
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Daùn key"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy maõ"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6120
      TabIndex        =   14
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   9
      Left            =   4680
      TabIndex        =   10
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   8
      Left            =   3480
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   7
      Left            =   2280
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   5
      Left            =   1080
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   6
      Left            =   6960
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   4
      Left            =   4680
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   3
      Left            =   3480
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Laáy key"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6120
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Key PM"
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
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Maõ soá"
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
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Maõ soá thueá"
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
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function Int_StrToCodes(Str As String) As Long
    Dim i As Long, Length As Integer, k As Long, kq As Long
    
    Length = Len(Str)
    If Length > 0 Then
        If Length > 12 Then
            For i = Length To 1 Step -1
                k = Asc(Right(Str, i))
                If (k Mod 2 = 0) Then
                    kq = kq + 4 * i * (k ^ 2)
                Else
                    kq = kq + 5 * i * (k ^ 2)
                End If
            Next
        Else
            For i = Length To 1 Step -1
                k = Asc(Right(Str, i))
                If (k Mod 2 = 0) Then
                    kq = kq + 8 * i * (k ^ 3)
                Else
                    kq = kq + 9 * i * (k ^ 3)
                End If
            Next
        End If
    End If
    Int_StrToCodes = kq
End Function
Public Function Int_StrToCode(Str As String) As Long
    Dim i As Long, Length As Integer, k As Long, kq As Long
    
    Length = Len(Str)
    If Length > 0 Then
        If Length > 12 Then
            For i = Length To 1 Step -1
                k = Asc(Right(Str, i))
                kq = kq + 2 * i * (k ^ 2)
            Next
        Else
            For i = Length To 1 Step -1
                k = Asc(Right(Str, i))
                kq = kq + 8 * i * (k ^ 3)
            Next
        End If
    End If
    Int_StrToCode = kq
End Function
Private Sub Command1_Click()
 Dim a(5) As String
 On Error GoTo Error_Handler
 a(0) = Mid(Text1(1).Text, 2, 4)
 a(1) = Mid(Text1(2).Text, 2, 4)
 a(2) = Mid(Text1(3).Text, 2, 4)
 a(3) = Mid(Text1(1).Text, 1, 1) + Mid(Text1(2).Text, 1, 1) + Mid(Text1(3).Text, 1, 1) + Text1(4).Text
 
 Dim st As String
 st = Chr(8) + Chr(13) + Chr(27) + Chr(27) + Chr(8) + Chr(13) + Chr(27) + Chr(27) + Chr(8) + Chr(13) + Chr(27) + Chr(27) + Chr(127)
 Text1(6).Text = CStr(Int_StrToCodes(Text1(0).Text + st) + Int_StrToCodes(a(0) + a(1) + a(2) + st) + Int(a(3)) + 2006) + "-" + CStr(Int_StrToCodes(CStr(Int_StrToCodes(Text1(0).Text + st) + Int_StrToCodes(a(0) + a(1) + a(2) + st) + Int(a(3)) + 2006)))
 Text1(5).Text = Mid(Text1(6).Text, 1, 4)
 Text1(7).Text = Mid(Text1(6).Text, 5, 4)
 Text1(8).Text = Mid(Text1(6).Text, 10, 4)
 Text1(9).Text = Mid(Text1(6).Text, 14, 4)
Error_Handler:
End Sub

Private Sub Command2_Click()
On Error GoTo Error_Handler
 Clipboard.Clear
  Clipboard.SetText Text1(5).Text + "*" + Text1(7).Text + "*" + Text1(8).Text + "*" + Text1(9).Text
Error_Handler:
End Sub

Private Sub Command3_Click()
On Error GoTo Error_Handler
Text1(0).Text = Clipboard.GetText()
Command1_Click
Error_Handler:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
        Case 27:   Unload Me
 End Select
End Sub

Private Sub Text1_Change(Index As Integer)
Dim a() As String
 On Error GoTo Error_Handler
If (Len(Text1(0).Text) > 13) Then
    a = Split(Text1(0).Text, "*")
    
    Text1(0).Text = Trim(a(0))
    Text1(1).Text = Trim(a(1))
    Text1(2).Text = Trim(a(2))
    Text1(3).Text = Trim(a(3))
    Text1(4).Text = Trim(a(4))
End If
Error_Handler:
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))


End Sub
