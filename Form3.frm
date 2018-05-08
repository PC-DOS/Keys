VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Number And Compute Area Test"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6630
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "虚拟小键盘"
      ForeColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton Command1 
         Caption         =   "*"
         Height          =   615
         Index           =   106
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "-"
         Height          =   615
         Index           =   109
         Left            =   2640
         TabIndex        =   16
         Top             =   240
         Width           =   705
      End
      Begin VB.CommandButton Command1 
         Caption         =   "="
         Height          =   1335
         Index           =   13
         Left            =   2640
         TabIndex        =   15
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+"
         Height          =   1335
         Index           =   107
         Left            =   2640
         TabIndex        =   14
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "/"
         Height          =   615
         Index           =   111
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         Height          =   615
         Index           =   110
         Left            =   1800
         TabIndex        =   12
         Top             =   3120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "0"
         Height          =   615
         Index           =   96
         Left            =   120
         TabIndex        =   11
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "3"
         Height          =   615
         Index           =   99
         Left            =   1800
         TabIndex        =   10
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "2"
         Height          =   615
         Index           =   98
         Left            =   960
         TabIndex        =   9
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   615
         Index           =   97
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "6"
         Height          =   615
         Index           =   102
         Left            =   1800
         TabIndex        =   7
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "5"
         Height          =   615
         Index           =   101
         Left            =   960
         TabIndex        =   6
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "4"
         Height          =   615
         Index           =   100
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "9"
         Height          =   615
         Index           =   105
         Left            =   1800
         TabIndex        =   4
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "8"
         Height          =   615
         Index           =   104
         Left            =   960
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "7"
         Height          =   615
         Index           =   103
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Num Lock"
         Height          =   615
         Index           =   144
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "友好的按键名称"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   3240
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
      Height          =   495
      Left            =   3720
      TabIndex        =   22
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "键位代码"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   21
      Top             =   2280
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
      Left            =   3720
      TabIndex        =   20
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "复位所有虚拟按键"
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
      Left            =   3720
      TabIndex        =   19
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "退出"
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
      Left            =   3720
      TabIndex        =   18
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 27 Then
Unload Me
Exit Sub
End If
If KeyCode = vbKeyF4 And Shift = vbAltMask Then
Unload Me
Exit Sub
End If
Command1(KeyCode).Visible = False
Label2.Caption = ""
Label4.Caption = ""
Label2.Caption = KeyCode
Label4.FontName = Command1(KeyCode).FontName
Label4.Caption = Command1(KeyCode).Caption
If Label4.Caption = "" Then
Label4.FontName = Form1.Command1(KeyCode).FontName
Label4.Caption = Form1.Command1(KeyCode).Caption
If Label4.Caption = "" Then
Label4.FontName = Form2.Command1(KeyCode).FontName
Label4.Caption = Form2.Command1(KeyCode).Caption
If Label4.Caption = "" Then
Label4.FontName = "Tahoma"
Label4.Caption = "Undefined"
End If
End If
End If
Form3.SetFocus
KeyCode = 0
Dim a As Object
Dim cunt As Integer
cunt = 0
For Each a In Command1
If a.Visible = False Then
cunt = cunt + 1
If cunt = Command1.Count Then
MsgBox "测试完毕,所有小键盘按键都可以正常使用!" & vbCrLf & "请单击'确定',复位虚拟按键", vbExclamation, "Congratulations!"
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
With Me
.Icon = Form1.Icon
.Left = Screen.Width / 2 - Me.Width / 2
.Top = Screen.Height / 2 - Me.Height / 2
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
End With
Dim a As Object
For Each a In Me.Command1
With a
a.FontName = "Tahoma"
.Enabled = False
.TabStop = False
.Visible = True
End With
Next
Label4.Caption = ""
Label2.Caption = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Form1.Show
End Sub
Private Sub Label1_Click()
On Error Resume Next
Dim a As Object
For Each a In Me.Command1
With a
a.FontName = "Tahoma"
.Enabled = False
.TabStop = False
.Visible = True
End With
Next
Label4.Caption = ""
Label2.Caption = ""
End Sub
Private Sub Label6_Click()
On Error Resume Next
Unload Me
End Sub
