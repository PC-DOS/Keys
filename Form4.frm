VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advanced Test"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Timer ReplaceCunter 
      Interval        =   100
      Left            =   1440
      Top             =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "按键重复速度"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Text1 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1080
         Width           =   4215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "字符每秒"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3480
         TabIndex        =   7
         Top             =   2160
         Width           =   840
      End
      Begin VB.Label Sec 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "字符每100毫秒"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   840
         TabIndex        =   5
         Top             =   2160
         Width           =   1365
      End
      Begin VB.Label MSec 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "测试区域"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "    在下方的文本框中按下一个字母/数字键并至少保持1秒钟,即可测出按键字符重复速度"
         Height          =   495
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "Form4.frx":0000
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type KeyTestResult
TxtLength As Long
Times100Ms As Long
TimesPerMSecond As Long
TimesPerSecond As Long
End Type
Dim ktst As KeyTestResult
Private Sub Command1_Click()
On Error Resume Next
Form1.Show
Unload Me
End Sub
Private Sub Form_Activate()
Command1.SetFocus
End Sub
Private Sub Form_Load()
On Error Resume Next
Command1.Cancel = True
With Me
.Left = Screen.Width / 2 - Me.Width / 2
.Top = Screen.Height / 2 - Me.Height / 2
.Icon = Form1.Icon
End With
With Me.ReplaceCunter
.Interval = 1000
.Enabled = True
End With
With ktst
.TxtLength = 0
.Times100Ms = 0
.TimesPerMSecond = 0
.TimesPerSecond = 0
End With
With Me.MSec
.Caption = ktst.TimesPerMSecond
End With
With Me.Sec
.Caption = ktst.TimesPerSecond
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Form1.Show
End Sub
Private Sub ReplaceCunter_Timer()
On Error Resume Next
ktst.TxtLength = Len(Text1.Text)
ktst.Times100Ms = ktst.Times100Ms + 1
ktst.TimesPerSecond = ktst.TxtLength / ktst.Times100Ms
ktst.TimesPerMSecond = ktst.TimesPerSecond / 10
With Me.MSec
.Caption = ktst.TimesPerMSecond
End With
With Me.Sec
.Caption = ktst.TimesPerSecond
End With
End Sub
Private Sub Text1_GotFocus()
On Error Resume Next
With Me.ReplaceCunter
.Interval = 1000
.Enabled = True
End With
With ktst
.TxtLength = 0
.Times100Ms = 0
.TimesPerMSecond = 0
.TimesPerSecond = 0
End With
With Me.MSec
.Caption = ktst.TimesPerMSecond
End With
With Me.Sec
.Caption = ktst.TimesPerSecond
End With
End Sub
Private Sub Text1_LostFocus()
On Error Resume Next
Text1.Text = ""
With Me.ReplaceCunter
.Interval = 1000
.Enabled = False
End With
With ktst
.TxtLength = 0
.Times100Ms = 0
.TimesPerMSecond = 0
.TimesPerSecond = 0
End With
With Me.MSec
.Caption = ktst.TimesPerMSecond
End With
With Me.Sec
.Caption = ktst.TimesPerSecond
End With
End Sub
