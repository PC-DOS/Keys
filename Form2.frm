VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Function Area Test"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   4800
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "虚拟功能区"
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   15.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   40
         Left            =   1800
         TabIndex        =   12
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   15.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   37
         Left            =   720
         TabIndex        =   11
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   18
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   39
         Left            =   2880
         TabIndex        =   10
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Scoll Lock"
         Height          =   375
         Index           =   145
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Home"
         Height          =   495
         Index           =   36
         Left            =   1800
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "End"
         Height          =   495
         Index           =   35
         Left            =   1800
         TabIndex        =   7
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Page Up"
         Height          =   495
         Index           =   33
         Left            =   2880
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Page Down"
         Height          =   495
         Index           =   34
         Left            =   2880
         TabIndex        =   5
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Insert"
         Height          =   495
         Index           =   45
         Left            =   720
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Delete"
         Height          =   495
         Index           =   46
         Left            =   720
         TabIndex        =   3
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   15.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   38
         Left            =   1800
         TabIndex        =   2
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Pause/Break"
         Height          =   375
         Index           =   19
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
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
      Left            =   3120
      TabIndex        =   18
      Top             =   3840
      Width           =   1575
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
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   2895
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
      Left            =   120
      TabIndex        =   16
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "键位代码"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4560
      Width           =   1095
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
      Left            =   2520
      TabIndex        =   14
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "友好的按键名称"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   4560
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
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
Label4.FontName = Form3.Command1(KeyCode).FontName
Label4.Caption = Form3.Command1(KeyCode).Caption
If Label4.Caption = "" Then
Label4.FontName = "Tahoma"
Label4.Caption = "Undefined"
End If
End If
End If
Form2.SetFocus
KeyCode = 0
Dim a As Object
Dim cunt As Integer
cunt = 0
For Each a In Command1
If a.Visible = False Then
cunt = cunt + 1
If cunt = Command1.Count Then
MsgBox "测试完毕,所有功能区按键都可以正常使用!" & vbCrLf & "请单击'确定',复位虚拟按键", vbExclamation, "Congratulations!"
Dim b As Object
For Each b In Command1
b.FontName = "Tahoma"
b.Enabled = False
b.Visible = True
b.TabStop = False
Next
With Me.Command1(37)
.FontName = "Marlett"
.FontSize = 18
.FontBold = True
End With
With Me.Command1(38)
.FontName = "Marlett"
.FontSize = 18
.FontBold = True
End With
With Me.Command1(39)
.FontName = "Marlett"
.FontSize = 18
.FontBold = True
End With
With Me.Command1(40)
.FontName = "Marlett"
.FontSize = 18
.FontBold = True
End With
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
With Me.Command1(37)
.FontName = "Marlett"
.FontSize = 18
.FontBold = True
End With
With Me.Command1(38)
.FontName = "Marlett"
.FontSize = 18
.FontBold = True
End With
With Me.Command1(39)
.FontName = "Marlett"
.FontSize = 18
.FontBold = True
End With
With Me.Command1(40)
.FontName = "Marlett"
.FontSize = 18
.FontBold = True
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
With Me.Command1(37)
.FontName = "Marlett"
.FontSize = 18
.FontBold = True
End With
With Me.Command1(38)
.FontName = "Marlett"
.FontSize = 18
.FontBold = True
End With
With Me.Command1(39)
.FontName = "Marlett"
.FontSize = 18
.FontBold = True
End With
With Me.Command1(40)
.FontName = "Marlett"
.FontSize = 18
.FontBold = True
End With
Label4.Caption = ""
Label2.Caption = ""
End Sub
Private Sub Label6_Click()
On Error Resume Next
Unload Me
End Sub
