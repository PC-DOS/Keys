VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keyboard Test - PC-DOS Workshop"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   15270
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "F12"
      Height          =   375
      Index           =   123
      Left            =   11880
      TabIndex        =   13
      Top             =   240
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "标准按键区键位测试(按下一个键,对应的键会消失)"
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12735
      Begin VB.CommandButton Command1 
         Caption         =   "Ctrl"
         Height          =   495
         Index           =   17
         Left            =   120
         TabIndex        =   73
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Win"
         Height          =   495
         Index           =   91
         Left            =   1440
         MaskColor       =   &H00000000&
         Picture         =   "Form1.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   3120
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Alt"
         Height          =   495
         Index           =   18
         Left            =   2400
         TabIndex        =   71
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Space"
         Height          =   495
         Index           =   32
         Left            =   3360
         TabIndex        =   70
         Top             =   3120
         Width           =   6255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Alt"
         Height          =   495
         Index           =   998
         Left            =   9720
         TabIndex        =   69
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Menu"
         Height          =   495
         Index           =   93
         Left            =   10680
         MaskColor       =   &H000000FF&
         Picture         =   "Form1.frx":06F4
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   3120
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ctrl"
         Height          =   495
         Index           =   999
         Left            =   11400
         TabIndex        =   67
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Shift"
         Height          =   495
         Index           =   997
         Left            =   9960
         TabIndex        =   66
         Top             =   2520
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "/"
         Height          =   495
         Index           =   191
         Left            =   9240
         TabIndex        =   65
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         Height          =   495
         Index           =   190
         Left            =   8520
         TabIndex        =   64
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   ","
         Height          =   495
         Index           =   188
         Left            =   7800
         TabIndex        =   63
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "M"
         Height          =   495
         Index           =   77
         Left            =   7080
         TabIndex        =   62
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "N"
         Height          =   495
         Index           =   78
         Left            =   6360
         TabIndex        =   61
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "B"
         Height          =   495
         Index           =   66
         Left            =   5640
         TabIndex        =   60
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "V"
         Height          =   495
         Index           =   86
         Left            =   4920
         TabIndex        =   59
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "C"
         Height          =   495
         Index           =   67
         Left            =   4200
         TabIndex        =   58
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         Height          =   495
         Index           =   88
         Left            =   3480
         TabIndex        =   57
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Z"
         Height          =   495
         Index           =   90
         Left            =   2760
         TabIndex        =   56
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Shift"
         Height          =   495
         Index           =   16
         Left            =   120
         TabIndex        =   55
         Top             =   2520
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "'"
         Height          =   495
         Index           =   222
         Left            =   9000
         TabIndex        =   54
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   ":"
         Height          =   495
         Index           =   186
         Left            =   8160
         TabIndex        =   53
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "L"
         Height          =   495
         Index           =   76
         Left            =   7440
         TabIndex        =   52
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "K"
         Height          =   495
         Index           =   75
         Left            =   6720
         TabIndex        =   51
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "J"
         Height          =   495
         Index           =   74
         Left            =   6000
         TabIndex        =   50
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "H"
         Height          =   495
         Index           =   72
         Left            =   5280
         TabIndex        =   49
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "G"
         Height          =   495
         Index           =   71
         Left            =   4560
         TabIndex        =   48
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F"
         Height          =   495
         Index           =   70
         Left            =   3840
         TabIndex        =   47
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "D"
         Height          =   495
         Index           =   68
         Left            =   3120
         TabIndex        =   46
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "S"
         Height          =   495
         Index           =   83
         Left            =   2400
         TabIndex        =   45
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "A"
         Height          =   495
         Index           =   65
         Left            =   1680
         TabIndex        =   44
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Caps Lock"
         Height          =   495
         Index           =   20
         Left            =   120
         TabIndex        =   43
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Enter"
         Height          =   1095
         Index           =   13
         Left            =   9840
         TabIndex        =   42
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "}"
         Height          =   495
         Index           =   221
         Left            =   9000
         TabIndex        =   41
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "{"
         Height          =   495
         Index           =   219
         Left            =   8160
         TabIndex        =   40
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "P"
         Height          =   495
         Index           =   80
         Left            =   7440
         TabIndex        =   39
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "O"
         Height          =   495
         Index           =   79
         Left            =   6720
         TabIndex        =   38
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "I"
         Height          =   495
         Index           =   73
         Left            =   6000
         TabIndex        =   37
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "U"
         Height          =   495
         Index           =   85
         Left            =   5280
         TabIndex        =   36
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Y"
         Height          =   495
         Index           =   89
         Left            =   4560
         TabIndex        =   35
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "T"
         Height          =   495
         Index           =   84
         Left            =   3840
         TabIndex        =   34
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "R"
         Height          =   495
         Index           =   82
         Left            =   3120
         TabIndex        =   33
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "E"
         Height          =   495
         Index           =   69
         Left            =   2400
         TabIndex        =   32
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "W"
         Height          =   495
         Index           =   87
         Left            =   1680
         TabIndex        =   31
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Q"
         Height          =   495
         Index           =   81
         Left            =   960
         TabIndex        =   30
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Tab"
         Height          =   495
         Index           =   9
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Back Space"
         Height          =   495
         Index           =   8
         Left            =   10560
         TabIndex        =   28
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "|"
         Height          =   495
         Index           =   220
         Left            =   9840
         TabIndex        =   27
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+"
         Height          =   495
         Index           =   187
         Left            =   9000
         TabIndex        =   26
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "_"
         Height          =   495
         Index           =   189
         Left            =   8160
         TabIndex        =   25
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "0"
         Height          =   495
         Index           =   48
         Left            =   7440
         TabIndex        =   24
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "9"
         Height          =   495
         Index           =   57
         Left            =   6720
         TabIndex        =   23
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "8"
         Height          =   495
         Index           =   56
         Left            =   6000
         TabIndex        =   22
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "7"
         Height          =   495
         Index           =   55
         Left            =   5280
         TabIndex        =   21
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "6"
         Height          =   495
         Index           =   54
         Left            =   4560
         TabIndex        =   20
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "5"
         Height          =   495
         Index           =   53
         Left            =   3840
         TabIndex        =   19
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "4"
         Height          =   495
         Index           =   52
         Left            =   3120
         TabIndex        =   18
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "3"
         Height          =   495
         Index           =   51
         Left            =   2400
         TabIndex        =   17
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "2"
         Height          =   495
         Index           =   50
         Left            =   1680
         TabIndex        =   16
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   495
         Index           =   49
         Left            =   960
         TabIndex        =   15
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "~"
         Height          =   495
         Index           =   192
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F11"
         Height          =   375
         Index           =   122
         Left            =   10800
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F10"
         Height          =   375
         Index           =   121
         Left            =   9840
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F9"
         Height          =   375
         Index           =   120
         Left            =   8880
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F8"
         Height          =   375
         Index           =   119
         Left            =   7920
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F7"
         Height          =   375
         Index           =   118
         Left            =   6960
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F6"
         Height          =   375
         Index           =   117
         Left            =   6000
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F5"
         Height          =   375
         Index           =   116
         Left            =   5040
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F4"
         Height          =   375
         Index           =   115
         Left            =   4080
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F3"
         Height          =   375
         Index           =   114
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F2"
         Height          =   375
         Index           =   113
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F1"
         Height          =   375
         Index           =   112
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Esc"
         Height          =   375
         Index           =   27
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "高级测试"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12960
      TabIndex        =   81
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "小键盘键位测试"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12960
      TabIndex        =   80
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "功能区键位测试"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12960
      TabIndex        =   79
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "友好的按键名称"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12960
      TabIndex        =   78
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12960
      TabIndex        =   77
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "键位代码"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12960
      TabIndex        =   76
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12960
      TabIndex        =   75
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "复位虚拟按键"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12960
      TabIndex        =   74
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Activate()
Me.SetFocus
End Sub
Private Sub Form_Initialize()
On Error Resume Next
If App.PrevInstance = False Then
Dim ans As Integer
ans = MsgBox("此应用程序仅适用与对Windows95/Windows98标准QWERTY键盘(美式101键盘)进行测试,如果您使用的是其它类型的键盘(如:麦金托什(苹果)键盘或日文键盘),不推荐使用本程序" & vbCrLf & "为了获得准确的测试结果,建议使用Microsoft(R) Windows XP操作系统作为测试环境并且关闭输入法" & vbCrLf & "继续运行?" & vbCrLf & vbCrLf & "点击[是]继续" & vbCrLf & "点击[否]退出", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Load Form1
Else
End
End If
Else
MsgBox "对不起,本程序不允许同时执行2个及以上的实例,程序将退出", vbCritical, "Error"
End
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyF4 And Shift = vbAltMask Then
End
End If
Command1(KeyCode).Visible = False
If KeyCode = 16 Then
Command1(997).Visible = False
End If
If KeyCode = 17 Then
Command1(999).Visible = False
End If
If KeyCode = 18 Then
Command1(998).Visible = False
End If
Label2.Caption = ""
Label4.Caption = ""
Label2.Caption = KeyCode
Label4.FontName = Command1(KeyCode).FontName
Label4.Caption = Command1(KeyCode).Caption
If Label4.Caption = "" Then
Label4.FontName = Form2.Command1(KeyCode).FontName
Label4.Caption = Form2.Command1(KeyCode).Caption
If Label4.Caption = "" Then
Label4.FontName = Form3.Command1(KeyCode).FontName
Label4.Caption = Form3.Command1(KeyCode).Caption
If Label4.Caption = "" Then
Label4.FontName = "Tahoma"
Label4.Caption = "Undefined"
End If
End If
End If
Me.SetFocus
KeyCode = 0
Dim a As Object
Dim cunt As Integer
cunt = 0
For Each a In Command1
If a.Visible = False Then
cunt = cunt + 1
If cunt = Command1.Count Then
MsgBox "测试完毕,所有标准按键都可以正常使用!" & vbCrLf & "请单击'确定',复位虚拟按键", vbExclamation, "Congratulations!"
Dim b As Object
For Each b In Command1
b.FontName = "Tahoma"
b.Enabled = False
b.Visible = True
b.TabStop = False
Next
Label4.Caption = ""
Label2.Caption = ""
End If
End If
Next
End Sub
Private Sub Form_Load()
On Error Resume Next
With Form2.Command1(37)
.FontName = "Marlett"
.FontSize = 18
.FontBold = True
End With
With Form2.Command1(38)
.FontName = "Marlett"
.FontSize = 18
.FontBold = True
End With
With Form2.Command1(39)
.FontName = "Marlett"
.FontSize = 18
.FontBold = True
End With
With Form2.Command1(40)
.FontName = "Marlett"
.FontSize = 18
.FontBold = True
End With
With Me
.Left = Screen.Width / 2 - Me.Width / 2
.Top = Screen.Height / 2 - Me.Height / 2
.KeyPreview = True
.BackColor = vbBlack
Dim a As Object
For Each a In Command1
a.FontName = "Tahoma"
a.Enabled = False
a.Visible = True
a.TabStop = False
Next
.Label1.Enabled = True
End With
With Frame1
.BackColor = vbBlack
.ForeColor = vbWhite
End With
Label4.Caption = ""
Label2.Caption = ""
Command1(192).Caption = "~" & vbCrLf & "`"
Command1(189).Caption = "_" & vbCrLf & "-"
Command1(220).Caption = "|" & vbCrLf & "\"
Command1(219).Caption = "{" & vbCrLf & "["
Command1(221).Caption = "}" & vbCrLf & "]"
Command1(186).Caption = ":" & vbCrLf & ";"
Command1(222).Caption = Chr(34) & vbCrLf & " '"
Command1(188).Caption = "<" & vbCrLf & ","
Command1(190).Caption = ">" & vbCrLf & "."
Command1(191).Caption = "?" & vbCrLf & "/"
Command1(187).Caption = "+" & vbCrLf & "="
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Unload Me
Unload Form1
Unload Form2
Unload Form3
Unload Form4
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload Me
Unload Form1
Unload Form2
Unload Form3
Unload Form4
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End
End Sub
Private Sub Label1_Click()
On Error Resume Next
Dim a As Object
For Each a In Command1
a.FontName = "Tahoma"
a.Enabled = False
a.Visible = True
a.TabStop = False
Next
Label4.Caption = ""
Label2.Caption = ""
End Sub
Private Sub Label6_Click()
On Error Resume Next
Form2.Show
'Unload Me
Me.Hide
End Sub
Private Sub Label7_Click()
On Error Resume Next
Form3.Show
'Unload Me
Me.Hide
End Sub
Private Sub Label8_Click()
On Error Resume Next
Form4.Show
'Unload Me
Me.Hide
End Sub
