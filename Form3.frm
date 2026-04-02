VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10665
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMST 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Index           =   1
      Left            =   5640
      TabIndex        =   27
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   3840
      Visible         =   0   'False
      Width           =   9495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "Thoâng tin ñaêng kyù license "
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   19
      Top             =   0
      Width           =   10095
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Soá chöùng töø"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000016&
         Caption         =   "..."
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   24
         Top             =   1080
         Width           =   6495
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000016&
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   23
         Top             =   720
         Width           =   5415
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   " Hình thöùc ñaêng kyù"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000016&
         Caption         =   "..."
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   21
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   " Teân Cty"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
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
      Index           =   12
      Left            =   7920
      TabIndex        =   18
      Top             =   3120
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
      Index           =   11
      Left            =   6360
      TabIndex        =   17
      Top             =   3120
      Width           =   1575
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
      Index           =   10
      Left            =   1320
      TabIndex        =   15
      Top             =   3120
      Width           =   615
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
      Left            =   1920
      TabIndex        =   14
      Top             =   3120
      Width           =   1815
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
      Left            =   3720
      TabIndex        =   13
      Top             =   3120
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
      Index           =   9
      Left            =   5040
      TabIndex        =   12
      Top             =   3120
      Width           =   1335
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
      Left            =   9360
      TabIndex        =   11
      Top             =   3120
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
      Index           =   6
      Left            =   7920
      TabIndex        =   10
      Top             =   2400
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
      Index           =   5
      Left            =   6360
      TabIndex        =   9
      Top             =   2400
      Width           =   1575
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
      Left            =   9360
      TabIndex        =   7
      Top             =   2400
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
      Height          =   405
      Index           =   1
      Left            =   1320
      TabIndex        =   6
      Top             =   2400
      Width           =   615
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
      Left            =   1920
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
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
      Left            =   3720
      TabIndex        =   4
      Top             =   2400
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
      Index           =   4
      Left            =   5040
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtMST 
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
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Width           =   3495
   End
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
      Left            =   9360
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      Caption         =   "Mac"
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
      Index           =   6
      Left            =   5040
      TabIndex        =   28
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
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
      Index           =   5
      Left            =   360
      TabIndex        =   16
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
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
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   8
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
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
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SECRET_KEY_MST As Long = &H7B4D8E2F

Private Const RANDOM_MIN As Long = 10
Private Const RANDOM_MAX As Long = 99

Private Const SECRET_KEYMac As Long = &H5A3C9F2E

Private Const SECRET_KEY2 As Long = 13579
Private Const BASE36 As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const SECRET_KEY As Byte = &H5A    ' Khóa bí m?t
Private Const CHARSET As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

'Mã hóa cho so chung tu
Public Function EncodeLicense6(ByVal num As Long, ByRef randomNum As Long) As String
'Dim randomNum As Long
    Dim combined As Variant   ' <--- Dùng Variant
    Dim encrypted As Variant  ' <--- Dùng Variant
    Dim result As String

    On Error GoTo ErrorHandler

    ' Gi?i h?n s? d?u vào du?i 2 tri?u
    If num > 1999999 Then
        EncodeLicense6 = "ERROR"
        Exit Function
    End If

    ' T?o s? random 3 ch? s?
    'Randomize
    'randomNum = Int((RANDOM_MAX - RANDOM_MIN + 1) * Rnd + RANDOM_MIN)

    ' K?t h?p: random (3 ch? s?) + num (t?i da 1,999,999)
    combined = randomNum * 2000000
    combined = combined + num

    ' Mã hóa
    encrypted = combined Xor SECRET_KEY
    encrypted = encrypted + 12345

    ' Gi?i h?n trong 36^6 = 2,176,782,336
    encrypted = encrypted - Int(encrypted / 2176782336#) * 2176782336#

    ' Chuy?n sang base36 (6 ký t?)
    result = ""
    Do While encrypted > 0
        result = Mid(CHARSET, (encrypted Mod 36) + 1, 1) & result
        encrypted = Int(encrypted / 36)
    Loop

    EncodeLicense6 = Right(String(6, "0") & result, 6)
    Exit Function

ErrorHandler:
    EncodeLicense6 = "ERROR"
End Function

' ==============================================
' DECODE: Mã 6 ký t? -> s? g?c và random
' ==============================================
Public Function DecodeLicense6(ByVal txt As String, ByRef randomNum As Long) As Long
    Dim i As Integer
    Dim result As Variant   ' <--- Dùng Variant
    Dim combined As Variant
    
    On Error GoTo ErrorHandler
    
    ' Gi?i mã base36
    result = 0
    For i = 1 To Len(txt)
        result = result * 36 + (InStr(CHARSET, Mid(txt, i, 1)) - 1)
    Next
    
    ' Gi?i mã
    result = result - 12345
    combined = result Xor SECRET_KEY
    
    ' Tách random và s? g?c
    randomNum = Int(combined / 2000000)
    DecodeLicense6 = combined - (randomNum * 2000000)
    Exit Function
    
ErrorHandler:
    DecodeLicense6 = -1
End Function

' ==============================================
' ENCODE: Có th? truy?n random vào (dùng cho test)
' ==============================================

' ==============================================
' ENCODE: Có th? truy?n random vào (dùng cho test)
' ==============================================
Public Function EncodeLicense6Ex(ByVal num As Long, ByVal randomNum As Long) As String
    Dim combined As Currency   ' <--- Ð?i thành Currency
    Dim encrypted As Long
    Dim result As String
    
    On Error GoTo ErrorHandler
    
    If num > 1999999 Or randomNum < 100 Or randomNum > 999 Then
        EncodeLicense6Ex = "ERROR"
        Exit Function
    End If
    
    ' Dùng Currency d? tránh tràn
    combined = CDec(randomNum) * 2000000 + num
    
    encrypted = (CLng(combined) Xor SECRET_KEY) + 12345
    encrypted = encrypted Mod 2176782336#
    
    result = ""
    Do While encrypted > 0
        result = Mid(CHARSET, (encrypted Mod 36) + 1, 1) & result
        encrypted = encrypted \ 36
    Loop
    
    EncodeLicense6Ex = Right(String(6, "0") & result, 6)
    Exit Function
    
ErrorHandler:
    EncodeLicense6Ex = "ERROR"
End Function

' ==============================================
' GI? L?I HÀM CU (4 ký t?) d? tuong thích
' ==============================================
Public Function EncodeLicense4(ByVal num As Long) As String
    Dim encrypted As Long
    Dim result As String
    
    If num > 999999 Then
        EncodeLicense4 = "ERROR"
        Exit Function
    End If
    
    encrypted = ((num Xor SECRET_KEY) + 12345) Mod 1679616
    
    result = ""
    Do While encrypted > 0
        result = Mid(CHARSET, (encrypted Mod 36) + 1, 1) & result
        encrypted = encrypted \ 36
    Loop
    
    EncodeLicense4 = Right("0000" & result, 4)
End Function

Public Function DecodeLicense4(ByVal txt As String) As Long
    Dim i As Integer
    Dim result As Long
    
    result = 0
    For i = 1 To Len(txt)
        result = result * 36 + (InStr(CHARSET, Mid(txt, i, 1)) - 1)
    Next
    
    result = result - 12345
    result = result Xor SECRET_KEY
    
    DecodeLicense4 = result
End Function

Public Function EncodeLicense2(ByVal num As Long) As String
    Dim encrypted As Long
    Dim result As String
    
    encrypted = (num Xor SECRET_KEY) + 12345
    
    result = ""
    Do While encrypted > 0
        result = Mid(CHARSET, (encrypted Mod 36) + 1, 1) & result
        encrypted = encrypted \ 36
    Loop
    
    EncodeLicense2 = result
End Function

Public Function DecodeLicense2(ByVal txt As String) As Long
    Dim i As Integer
    Dim result As Long
    
    result = 0
    For i = 1 To Len(txt)
        result = result * 36 + (InStr(CHARSET, Mid(txt, i, 1)) - 1)
    Next
    
    result = result - 12345
    result = result Xor SECRET_KEY
    
    DecodeLicense2 = result
End Function

'Cho mac
 Public Function EncodeMAC12(ByVal macAddress As String) As String
    Dim cleanMAC As String
    Dim p1 As Long, p2 As Long
    Dim code1 As String, code2 As String
    
    ' Xóa d?u :
    cleanMAC = UCase(Replace(macAddress, ":", ""))
    
    If Len(cleanMAC) <> 12 Then
        EncodeMAC12 = "ERROR"
        Exit Function
    End If
    
    ' Tách 12 hex thành 2 ph?n, m?i ph?n 6 hex (3 byte)
    p1 = CLng("&H" & Mid(cleanMAC, 1, 6))   ' 6 hex d?u
    p2 = CLng("&H" & Mid(cleanMAC, 7, 6))   ' 6 hex cu?i
    
    ' Encode m?i ph?n thành 6 ký t?
    code1 = Encode6(p1)
    code2 = Encode6(p2)
    
    ' N?i l?i v?i d?u -
    EncodeMAC12 = code1 & "-" & code2
End Function

' ==============================================
' DECODE MAC t? 12 ký t? (2 block x 6 ký t?)
' ==============================================
Public Function DecodeMAC12(ByVal code As String) As String
    Dim parts() As String
    Dim p1 As Long, p2 As Long
    Dim hexMAC As String
    
    ' Tách 2 ph?n
    parts = Split(code, "-")
    
    If UBound(parts) <> 1 Then
        DecodeMAC12 = "ERROR"
        Exit Function
    End If
    
    ' Gi?i mã t?ng ph?n
    p1 = Decode6(parts(0))
    p2 = Decode6(parts(1))
    
    ' Ghép l?i thành MAC 12 hex
    hexMAC = Right("000000" & Hex(p1), 6) & _
             Right("000000" & Hex(p2), 6)
    
    DecodeMAC12 = UCase(hexMAC)
End Function

' ==============================================
' ENCODE: 6 hex (3 byte) -> 6 ký t? base36 (có random)
' ==============================================
Public Function Encode6(ByVal num As Long) As String
    Dim randomNum As Long
    Dim combined As Variant
    Dim encrypted As Variant
    Dim result As String
    
    On Error GoTo ErrorHandler
    
    ' Gi?i h?n num trong 0-16,777,215 (6 hex = 24 bit)
    num = num And &HFFFFFF
    
    ' T?o random 3 ch? s?
    Randomize
    randomNum = Int((RANDOM_MAX - RANDOM_MIN + 1) * Rnd + RANDOM_MIN)
    
    ' K?t h?p: random (3 ch? s?) * 16,777,216 + num
    ' 16,777,216 = 2^24
    combined = CDec(randomNum) * 16777216
    combined = combined + num
    
    ' Mã hóa
    encrypted = combined Xor SECRET_KEY2
    encrypted = encrypted + 12345
    
    ' Gi?i h?n trong 36^6 = 2,176,782,336
    encrypted = encrypted - Int(encrypted / 2176782336#) * 2176782336#
    
    ' Chuy?n sang base36 (6 ký t?)
    result = ""
    Do While encrypted > 0
        result = Mid(CHARSET, (encrypted Mod 36) + 1, 1) & result
        encrypted = Int(encrypted / 36)
    Loop
    
    Encode6 = Right(String(6, "0") & result, 6)
    Exit Function
    
ErrorHandler:
    Encode6 = "ERROR"
End Function

' ==============================================
' DECODE: 6 ký t? base36 -> 6 hex (3 byte)
' ==============================================
Public Function Decode6(ByVal txt As String) As Long
    Dim i As Integer
    Dim result As Variant
    Dim combined As Variant
    
    On Error GoTo ErrorHandler
    
    ' Gi?i mã base36
    result = 0
    For i = 1 To Len(txt)
        result = result * 36 + (InStr(CHARSET, Mid(txt, i, 1)) - 1)
    Next
    
    ' Gi?i mã
    result = result - 12345
    combined = result Xor SECRET_KEY2
    
    ' L?y num (b? random)
    Decode6 = combined Mod 16777216
    Exit Function
    
ErrorHandler:
    Decode6 = -1
End Function

' ==============================================
' Ð?NH D?NG MAC CÓ D?U ":"
' ==============================================
Public Function FormatMAC(ByVal mac As String) As String
    Dim result As String
    Dim i As Integer
    
    If Len(mac) <> 12 Then
        FormatMAC = mac
        Exit Function
    End If
    
    result = ""
    For i = 1 To Len(mac) Step 2
        If result <> "" Then result = result & ":"
        result = result & Mid(mac, i, 2)
    Next i
    
    FormatMAC = result
End Function
 

'random 6 cho 1_2
Public Function EncodeWithRandom6(ByVal strValue As String) As String
    Dim parts() As String
    Dim num1 As Long   ' <--- S?A: num1, không ph?i numl
    Dim num2 As Long
    Dim randomNum As Long
    Dim combined As Long
    
    On Error GoTo ErrorHandler
    
    ' Tách chu?i d?ng "1_2"
    parts = Split(strValue, "_")
    
    If UBound(parts) <> 1 Then
        EncodeWithRandom6 = "ERROR"
        Exit Function
    End If
    
    num1 = CLng(parts(0))
    num2 = CLng(parts(1))
    
    ' Gi?i h?n m?i s? trong 0-999
    If num1 > 999 Or num2 > 999 Then   ' <--- S?A: num1
        EncodeWithRandom6 = "ERROR"
        Exit Function
    End If
    
    ' T?o s? random 4 ch? s?
    Randomize
    randomNum = Int((RANDOM_MAX - RANDOM_MIN + 1) * Rnd + RANDOM_MIN)
    
    ' K?t h?p: randomNum (4 ch? s?) + num1 (3 ch? s?) + num2 (3 ch? s?)
    combined = randomNum * 1000000 + num1 * 1000 + num2   ' <--- S?A: num1
    
    ' Mã hóa thành 6 ký t?
    EncodeWithRandom6 = ToBase36_6(combined)
    Exit Function
    
ErrorHandler:
    EncodeWithRandom6 = "ERROR"
End Function

' ==============================================
' DECODE: 6 ký t? -> chu?i "x_y" và random
' ==============================================
Public Function DecodeWithRandom6(ByVal code As String, ByRef randomNum As Long) As String
    Dim combined As Long
    Dim num1 As Long
    Dim num2 As Long
    
    On Error GoTo ErrorHandler
    
    ' Gi?i mã base36
    combined = FromBase36_6(code)
    
    ' Tách thành random, num1, num2
    randomNum = combined \ 1000000
    num1 = (combined Mod 1000000) \ 1000
    num2 = combined Mod 1000
    
    DecodeWithRandom6 = num1 & "_" & num2
    Exit Function
    
ErrorHandler:
    DecodeWithRandom6 = "ERROR"
End Function

' ==============================================
' ENCODE: Có th? truy?n random vào (dùng cho test)
' ==============================================
Public Function EncodeWithRandom6Ex(ByVal strValue As String, ByVal randomNum As Long) As String
    Dim parts() As String
    Dim num1 As Long
    Dim num2 As Long
    Dim combined As Long
    
    On Error GoTo ErrorHandler
    
    ' Tách chu?i d?ng "1_2"
    parts = Split(strValue, "_")
    
    If UBound(parts) <> 1 Then
        EncodeWithRandom6Ex = "ERROR"
        Exit Function
    End If
    
    num1 = CLng(parts(0))
    num2 = CLng(parts(1))
    
    ' Gi?i h?n
    If num1 > 999 Or num2 > 999 Or randomNum < 1000 Or randomNum > 9999 Then
        EncodeWithRandom6Ex = "ERROR"
        Exit Function
    End If
    
    ' K?t h?p: randomNum (4 ch? s?) + num1 (3 ch? s?) + num2 (3 ch? s?)
    combined = randomNum * 1000000 + num1 * 1000 + num2
    
    ' Mã hóa thành 6 ký t?
    EncodeWithRandom6Ex = ToBase36_6(combined)
    Exit Function
    
ErrorHandler:
    EncodeWithRandom6Ex = "ERROR"
End Function

' ==============================================
' HÀM H? TR?: Chuy?n s? sang base36 (6 ký t?)
' ==============================================
Private Function ToBase36_6(ByVal num As Long) As String
    Dim result As String
    
    If num = 0 Then
        ToBase36_6 = "000000"
        Exit Function
    End If
    
    result = ""
    Do While num > 0
        result = Mid(CHARSET, (num Mod 36) + 1, 1) & result
        num = num \ 36
    Loop
    
    ToBase36_6 = Right(String(6, "0") & result, 6)
End Function

' ==============================================
' HÀM H? TR?: Gi?i mã base36
' ==============================================
Private Function FromBase36_6(ByVal txt As String) As Long
    Dim i As Integer

    FromBase36_6 = 0
    For i = 1 To Len(txt)
        FromBase36_6 = FromBase36_6 * 36 + (InStr(CHARSET, Mid(txt, i, 1)) - 1)
    Next i
End Function
'cho mst
 Public Function EncodeMST8(ByVal mst As String, ByRef randomNum As Long) As String
    Dim cleanMST As String
    Dim totalLen As Integer
    Dim p1 As String, p2 As String
    Dim len1 As Integer, len2 As Integer
    Dim mixed1 As String, mixed2 As String
    
    On Error GoTo ErrorHandler
    
    ' Xóa d?u -
    cleanMST = Replace(mst, "-", "")
    
    ' Ki?m tra ch? có s?
    If Not IsNumeric(cleanMST) Then
        EncodeMST8 = "ERROR"
        Exit Function
    End If
    
    totalLen = Len(cleanMST)
    
    ' Phân chia theo d? dài
    Select Case totalLen
        Case 10
            len1 = 5: len2 = 5
            p1 = Left(cleanMST, len1)
            p2 = Right(cleanMST, len2)
        Case 12
            len1 = 6: len2 = 6
            p1 = Left(cleanMST, len1)
            p2 = Right(cleanMST, len2)
        Case 13
            len1 = 6: len2 = 7
            p1 = Left(cleanMST, len1)
            p2 = Right(cleanMST, len2)
        Case Else
            EncodeMST8 = "ERROR"
            Exit Function
    End Select
    
    ' T?o s? random 2 ch? s? (10-99)
    Randomize
    randomNum = Int((RANDOM_MAX - RANDOM_MIN + 1) * Rnd + RANDOM_MIN)
    
    ' Tr?n random vào d? li?u
    mixed1 = MixString(p1, randomNum)
    mixed2 = MixString(p2, randomNum)
    
    ' Mã: 2 ký t? random + 2 ký t? d? dài + 8 ký t? d? li?u dã tr?n
    EncodeMST8 = ToBase36_2_Fixed(randomNum) & _
                 ToBase36_1(len1) & ToBase36_1(len2) & _
                 EncodeString(mixed1) & EncodeString(mixed2)
    Exit Function
    
ErrorHandler:
    EncodeMST8 = "ERROR"
End Function

' ==============================================
' DECODE: 12 ký t? -> MST
' ==============================================
Public Function DecodeMST8(ByVal code As String, ByRef randomNum As Long) As String
    Dim len1 As Integer, len2 As Integer
    Dim mixed1 As String, mixed2 As String
    Dim p1 As String, p2 As String
    Dim lenCode1 As String, lenCode2 As String
    Dim dataCode1 As String, dataCode2 As String
    Dim randomCode As String
    Dim rdn As Long
    On Error GoTo ErrorHandler

    ' Ki?m tra d? dài
    If Len(code) <> 12 Then
        DecodeMST8 = "ERROR"
        Exit Function
    End If

    ' L?y 2 ký t? random
    randomCode = Mid(code, 1, 2)
    rdn = FromBase36_2_Fixed(randomCode)

    ' L?y 2 ký t? d? dài
    lenCode1 = Mid(code, 3, 1)
    lenCode2 = Mid(code, 4, 1)
    len1 = FromBase36_1(lenCode1)
    len2 = FromBase36_1(lenCode2)

    ' L?y 8 ký t? d? li?u dã encode
    dataCode1 = Mid(code, 5, 4)
    dataCode2 = Mid(code, 9, 4)

    ' Gi?i mã d? li?u
    mixed1 = DecodeString(dataCode1)
    mixed2 = DecodeString(dataCode2)

    ' Tách random ra kh?i d? li?u
    p1 = UnmixString(mixed1, rdn)
    p2 = UnmixString(mixed2, rdn)

    ' Ghép l?i v?i dúng d? dài (gi? s? 0 d?u)
    DecodeMST8 = Right(String(len1, "0") & p1, len1) & Right(String(len2, "0") & p2, len2)
    Exit Function

ErrorHandler:
    DecodeMST8 = "ERROR"
End Function

' ==============================================
' ENCODE: Có th? truy?n random vào (dùng cho test)
' ==============================================
Public Function EncodeMST8Ex(ByVal mst As String, ByVal randomNum As Integer) As String
    Dim cleanMST As String
    Dim totalLen As Integer
    Dim p1 As String, p2 As String
    Dim len1 As Integer, len2 As Integer
    Dim mixed1 As String, mixed2 As String
    
    On Error GoTo ErrorHandler
    
    cleanMST = Replace(mst, "-", "")
    
    If Not IsNumeric(cleanMST) Then
        EncodeMST8Ex = "ERROR"
        Exit Function
    End If
    
    totalLen = Len(cleanMST)
    
    Select Case totalLen
        Case 10
            len1 = 5: len2 = 5
            p1 = Left(cleanMST, len1)
            p2 = Right(cleanMST, len2)
        Case 12
            len1 = 6: len2 = 6
            p1 = Left(cleanMST, len1)
            p2 = Right(cleanMST, len2)
        Case 13
            len1 = 6: len2 = 7
            p1 = Left(cleanMST, len1)
            p2 = Right(cleanMST, len2)
        Case Else
            EncodeMST8Ex = "ERROR"
            Exit Function
    End Select
    
    ' Tr?n random vào d? li?u
    mixed1 = MixString(p1, randomNum)
    mixed2 = MixString(p2, randomNum)
    
    EncodeMST8Ex = ToBase36_2_Fixed(randomNum) & _
                   ToBase36_1(len1) & ToBase36_1(len2) & _
                   EncodeString(mixed1) & EncodeString(mixed2)
    Exit Function
    
ErrorHandler:
    EncodeMST8Ex = "ERROR"
End Function

' ==============================================
' HÀM TR?N: Tr?n random vào chu?i s?
' ==============================================
Private Function MixString(ByVal s As String, ByVal key As Integer) As String
    Dim result As String
    Dim i As Integer
    Dim digit As Integer
    Dim keyDigit As Integer
    
    result = ""
    For i = 1 To Len(s)
        digit = CLng(Mid(s, i, 1))
        ' L?y ch? s? th? i c?a key (l?p l?i n?u key ng?n)
        keyDigit = (key \ (10 ^ ((i - 1) Mod 2))) Mod 10
        digit = (digit + keyDigit) Mod 10
        result = result & digit
    Next i
    
    MixString = result
End Function

' ==============================================
' HÀM TÁCH: Tách random kh?i chu?i s?
' ==============================================
Private Function UnmixString(ByVal s As String, ByVal key As Integer) As String
    Dim result As String
    Dim i As Integer
    Dim digit As Integer
    Dim keyDigit As Integer
    
    result = ""
    For i = 1 To Len(s)
        digit = CLng(Mid(s, i, 1))
        ' L?y ch? s? th? i c?a key (l?p l?i n?u key ng?n)
        keyDigit = (key \ (10 ^ ((i - 1) Mod 2))) Mod 10
        digit = (digit - keyDigit) Mod 10
        If digit < 0 Then digit = digit + 10
        result = result & digit
    Next i
    
    UnmixString = result
End Function

' ==============================================
' Encode 1 ký t? (0-35)
' ==============================================
Private Function ToBase36_1(ByVal num As Integer) As String
    If num < 0 Or num > 35 Then
        ToBase36_1 = "0"
    Else
        ToBase36_1 = Mid(CHARSET, num + 1, 1)
    End If
End Function

' ==============================================
' Decode 1 ký t? (0-35)
' ==============================================
Private Function FromBase36_1(ByVal c As String) As Integer
    Dim pos As Integer
    pos = InStr(CHARSET, c)
    If pos > 0 Then
        FromBase36_1 = pos - 1
    Else
        FromBase36_1 = 0
    End If
End Function

' ==============================================
' Encode 2 ký t? (LUÔN RA 2 KÝ T?) cho random 10-99
' ==============================================
Private Function ToBase36_2_Fixed(ByVal num As Integer) As String
    Dim high As Integer, low As Integer
    
    high = num \ 36
    low = num Mod 36
    
    ToBase36_2_Fixed = Mid(CHARSET, high + 1, 1) & Mid(CHARSET, low + 1, 1)
End Function

' ==============================================
' Decode 2 ký t? (LUÔN T? 2 KÝ T?) -> s?
' ==============================================
Private Function FromBase36_2_Fixed(ByVal code As String) As Integer
    Dim high As Integer, low As Integer
    
    If Len(code) <> 2 Then
        FromBase36_2_Fixed = 0
        Exit Function
    End If
    
    high = InStr(CHARSET, Mid(code, 1, 1)) - 1
    low = InStr(CHARSET, Mid(code, 2, 1)) - 1
    
    If high < 0 Then high = 0
    If low < 0 Then low = 0
    
    FromBase36_2_Fixed = high * 36 + low
End Function

' ==============================================
' Encode chu?i s? thành 4 ký t? base36
' ==============================================
Private Function EncodeString(ByVal s As String) As String
    Dim num As Variant
    Dim result As String
    Dim i As Integer
    
    num = CDec(0)
    For i = 1 To Len(s)
        num = num * 10 + CLng(Mid(s, i, 1))
    Next i
    
    If num = 0 Then
        EncodeString = "0000"
        Exit Function
    End If
    
    result = ""
    Do While num > 0
        result = Mid(CHARSET, (num Mod 36) + 1, 1) & result
        num = Int(num / 36)
    Loop
    
    EncodeString = Right(String(4, "0") & result, 4)
End Function

' ==============================================
' Decode 4 ký t? base36 thành chu?i s?
' ==============================================
Private Function DecodeString(ByVal code As String) As String
    Dim i As Integer
    Dim num As Variant

    num = CDec(0)
    For i = 1 To Len(code)
        num = num * 36 + (InStr(CHARSET, Mid(code, i, 1)) - 1)
    Next i

    DecodeString = CStr(num)
End Function

Private Sub Command3_Click()
    Text2.Text = Clipboard.GetText()
    Dim s As String
    Dim arr() As String
    Dim arr2() As String

    Dim i As Integer
    Dim encoded10 As String
    s = Text2.Text
    arr = Split(s, "*")

    For i = 0 To UBound(arr)
        Debug.Print arr(i)
        arr2 = Split(arr(4), "-")

        If i = 0 Then
            Text1(1).Text = arr(i)
        End If
        If i = 1 Then
            Text1(2).Text = arr(i)
            txtMST(0).Text = DecodeMST8(arr(i), Text1(1).Text)
        End If
        If i = 2 Then
            Text1(3).Text = arr(i)
        End If
        If i = 3 Then
            Text1(4).Text = arr(i)
        End If
        If i = 4 Then
            Text1(5).Text = arr2(0)
            Text1(6).Text = arr2(1)
            txtMST(1).Text = FormatMAC(DecodeMAC12(arr(4)))
        End If

        'Thuc thi viec ma hoa key de gui nguoc lai cho user
        'Thuc hien random ma truoc
        Randomize
        Dim randomNum As Long
        Do
            randomNum = Int((RANDOM_MAX - RANDOM_MIN + 1) * Rnd + RANDOM_MIN)
        Loop While randomNum = CLng(Text1(1).Text)
        Text1(10).Text = randomNum
        'encode mst
        Text1(7).Text = EncodeMST8(txtMST(0).Text, randomNum)
    Next
End Sub

