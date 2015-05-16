VERSION 5.00
Begin VB.Form Form1主界面 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "主界面"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label Lab装饰 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "东方幻想冲锋3D"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Index           =   0
      Left            =   11400
      TabIndex        =   7
      Top             =   10680
      Width           =   3480
   End
   Begin VB.Label Lab选项 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "English"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   435
      Index           =   3
      Left            =   1200
      TabIndex        =   6
      Top             =   8400
      Width           =   1680
   End
   Begin VB.Label Lab选项 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "剧情模式"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   435
      Index           =   6
      Left            =   1200
      TabIndex        =   5
      Top             =   6240
      Width           =   1860
   End
   Begin VB.Label Lab选项 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "生存模式"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   435
      Index           =   5
      Left            =   1200
      TabIndex        =   4
      Top             =   6960
      Width           =   1860
   End
   Begin VB.Label Lab选项 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "制作名单"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   435
      Index           =   4
      Left            =   1200
      TabIndex        =   3
      Top             =   7680
      Width           =   1860
   End
   Begin VB.Label Lab选项 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "字体显示有误请点我"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   11040
      Width           =   2295
   End
   Begin VB.Label Lab选项 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "选项设置"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   435
      Index           =   2
      Left            =   1200
      TabIndex        =   1
      Top             =   9120
      Width           =   1860
   End
   Begin VB.Label Lab选项 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "退出游戏"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   435
      Index           =   1
      Left            =   1200
      TabIndex        =   0
      Top             =   9840
      Width           =   1860
   End
End
Attribute VB_Name = "Form1主界面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function 切换界面语言(语言 As Long)
设置(4) = 语言
Select Case 语言
Case 0 '简体中文
  Lab选项(0).Caption = "字体显示有误请点我"
  Lab选项(1).Caption = "退出游戏"
  Lab选项(2).Caption = "选项设置"
  Lab选项(3).Caption = "English"
  Lab选项(4).Caption = "制作名单"
  Lab选项(5).Caption = "生存模式"
  Lab选项(6).Caption = "剧情模式"
  Lab装饰(0).Caption = "东方现代战争"
Case 1 'English
  Lab选项(0).Caption = "Click here if the words indicate wrong"
  Lab选项(1).Caption = "Exit"
  Lab选项(2).Caption = "Options"
  Lab选项(3).Caption = "简体中文"
  Lab选项(4).Caption = "Credits"
  Lab选项(5).Caption = "Survival Mode"
  Lab选项(6).Caption = "Story Mode"
  Lab装饰(0).Caption = "Touhou Modern War"
End Select
End Function
Private Sub Form_Activate()
On Error Resume Next
已阅读说明 = True
Dir1.Path = App.Path
Dir1.Path = App.Path
读取设置
切换界面语言 设置(4)
'——公共变量——
End Sub

Private Sub Lab选项_Click(index As Integer)
BG.BGM(9).Controls.Play
Select Case index
Case 0 '系统未安装楷体字体
  For i = 0 To Lab选项.UBound: Lab选项(i).Font = "宋体": Next
  Lab选项(0).Caption = "已切换宋体字，建议安装完整操作系统"
Case 1: Unload BG: Unload Me: End
Case 2: Form2设置界面.Show , vbNormalFocus
Case 3 '多语言切换
If Lab选项(3).Caption = "English" Then
  切换界面语言 1
Else
  切换界面语言 0
End If
Case 4: Form3制作名单豪华版.Show , vbNormalFocus
Case 5 '生存模式
  BG.BGM(0).url = ""
  Stage生存.Show , vbNormalFocus
Case 6 '剧情模式
  BG.BGM(0).url = ""
  Stage01.Show , vbNormalFocus
End Select
End Sub

Private Sub Lab选项_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Lab选项(index).ForeColor = vbWhite
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To Lab选项.UBound
Lab选项(i).ForeColor = &HC0C0C0
Next
End Sub
