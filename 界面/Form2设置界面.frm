VERSION 5.00
Begin VB.Form Form2设置界面 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "选项设置"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "Form2设置界面.frx":0000
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox Com 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   6
      ItemData        =   "Form2设置界面.frx":6248F
      Left            =   4800
      List            =   "Form2设置界面.frx":62499
      TabIndex        =   28
      Tag             =   "开关"
      Text            =   "WMP(慢差,兼容)"
      ToolTipText     =   "效率很高，尽量保留"
      Top             =   7440
      Width           =   1935
   End
   Begin VB.ComboBox Com 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   4
      ItemData        =   "Form2设置界面.frx":624C1
      Left            =   5400
      List            =   "Form2设置界面.frx":624CB
      TabIndex        =   25
      Tag             =   "开关"
      Text            =   "简体中文"
      ToolTipText     =   "效率很高，尽量保留"
      Top             =   5280
      Width           =   1335
   End
   Begin VB.ComboBox Com 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   5
      ItemData        =   "Form2设置界面.frx":624E2
      Left            =   13320
      List            =   "Form2设置界面.frx":624EC
      TabIndex        =   23
      Tag             =   "开关"
      Text            =   "   F键"
      ToolTipText     =   "照顾老FPS和COD玩家"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Tex无用 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   2
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "左Alt\Ctrl"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Tex无用 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   1
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "左Shift"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox Tex无用 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "鼠标右键"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Tex焦点 
      Height          =   270
      Left            =   7200
      TabIndex        =   15
      Text            =   "九州烽火工作室"
      Top             =   120
      Width           =   1335
   End
   Begin VB.HScrollBar HSc 
      Height          =   225
      Index           =   3
      Left            =   3720
      Max             =   100
      TabIndex        =   14
      Top             =   4680
      Value           =   1
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   360
      Top             =   240
   End
   Begin VB.ComboBox Com 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   2
      ItemData        =   "Form2设置界面.frx":62500
      Left            =   5880
      List            =   "Form2设置界面.frx":6250A
      TabIndex        =   9
      Text            =   "低Low"
      ToolTipText     =   "三线型或双线性过滤远处贴图"
      Top             =   4080
      Width           =   855
   End
   Begin VB.ComboBox Com 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      ItemData        =   "Form2设置界面.frx":6251D
      Left            =   5880
      List            =   "Form2设置界面.frx":6252A
      TabIndex        =   7
      Tag             =   "开关"
      Text            =   "关Off"
      ToolTipText     =   "带投影的动态阴影，较慢"
      Top             =   3480
      Width           =   855
   End
   Begin VB.ComboBox Com 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      ItemData        =   "Form2设置界面.frx":62544
      Left            =   5880
      List            =   "Form2设置界面.frx":62551
      TabIndex        =   6
      Tag             =   "开关"
      Text            =   "关Off"
      ToolTipText     =   "效率很高，尽量保留"
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Lab说明 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "播放核心"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   6
      Left            =   960
      TabIndex        =   27
      Tag             =   "开关"
      ToolTipText     =   "如您已安装DX8建议使用BASS高音质高速引擎"
      Top             =   7440
      Width           =   900
   End
   Begin VB.Label Lab分类 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "声音(Sound)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   2
      Left            =   720
      TabIndex        =   26
      Top             =   6720
      Width           =   2115
   End
   Begin VB.Label Lab说明 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Language"
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   4
      Left            =   960
      TabIndex        =   24
      Tag             =   "开关"
      ToolTipText     =   "上车、治疗、维修等"
      Top             =   5280
      Width           =   960
   End
   Begin VB.Label Lab按钮 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "保存(Save)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   12810
      TabIndex        =   22
      Top             =   10440
      Width           =   1605
   End
   Begin VB.Image Cmd 
      Height          =   705
      Index           =   1
      Left            =   12600
      Picture         =   "Form2设置界面.frx":6256B
      Stretch         =   -1  'True
      Top             =   10200
      Width           =   1995
   End
   Begin VB.Label Lab没用说明 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "保持下蹲"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Index           =   2
      Left            =   8640
      TabIndex        =   19
      ToolTipText     =   "此项仅供查看，不可修改"
      Top             =   4680
      Width           =   900
   End
   Begin VB.Label Lab没用说明 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "小跑冲刺"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Index           =   1
      Left            =   8640
      TabIndex        =   18
      ToolTipText     =   "此项仅供查看，不可修改"
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label Lab没用说明 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "精确瞄准"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Index           =   0
      Left            =   8640
      TabIndex        =   16
      ToolTipText     =   "此项仅供查看，不可修改"
      Top             =   3480
      Width           =   900
   End
   Begin VB.Image 装饰框 
      Height          =   495
      Index           =   11
      Left            =   8400
      Picture         =   "Form2设置界面.frx":6272B
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   6315
   End
   Begin VB.Image 装饰框 
      Height          =   495
      Index           =   10
      Left            =   8400
      Picture         =   "Form2设置界面.frx":62771
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   6315
   End
   Begin VB.Image 装饰框 
      Height          =   495
      Index           =   9
      Left            =   8400
      Picture         =   "Form2设置界面.frx":627B7
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   6315
   End
   Begin VB.Label Lab按钮 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "返回(Back)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   13
      Top             =   10440
      Width           =   1605
   End
   Begin VB.Image Cmd 
      Height          =   705
      Index           =   0
      Left            =   840
      Picture         =   "Form2设置界面.frx":627FD
      Stretch         =   -1  'True
      Top             =   10200
      Width           =   1995
   End
   Begin VB.Label Lab说明 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "动作热键"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   5
      Left            =   8640
      TabIndex        =   12
      Tag             =   "开关"
      ToolTipText     =   "上车、治疗、维修等"
      Top             =   2880
      Width           =   900
   End
   Begin VB.Label Lab数值 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   3
      Left            =   3480
      TabIndex        =   11
      Top             =   4680
      Width           =   120
   End
   Begin VB.Label Lab说明 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "特效细节"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   3
      Left            =   960
      TabIndex        =   10
      Tag             =   "滚动条"
      ToolTipText     =   "特效质量与细节"
      Top             =   4680
      Width           =   900
   End
   Begin VB.Label Lab说明 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "贴图过滤"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   2
      Left            =   960
      TabIndex        =   8
      Tag             =   "开关"
      ToolTipText     =   "三线型(低)远处会模糊，但速度快"
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label Lab说明 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "动态阴影"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   960
      TabIndex        =   5
      Tag             =   "开关"
      ToolTipText     =   "某些场景或物体的动态阴影"
      Top             =   3480
      Width           =   900
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "鼠标停留在选项和调节控件上有提示说明"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   5760
      TabIndex        =   4
      Top             =   1320
      Width           =   3780
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "选项设置(Options)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   0
      Left            =   5760
      TabIndex        =   3
      Top             =   600
      Width           =   3960
   End
   Begin VB.Label Lab说明 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "模型光影"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Tag             =   "开关"
      ToolTipText     =   "人物、武器、场景等光影"
      Top             =   2880
      Width           =   900
   End
   Begin VB.Label Lab分类 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "操作(Control)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   1
      Left            =   8400
      TabIndex        =   1
      Top             =   2040
      Width           =   2505
   End
   Begin VB.Label Lab分类 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "画质(Video)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   2040
      Width           =   2115
   End
   Begin VB.Image 装饰框 
      Height          =   495
      Index           =   0
      Left            =   720
      Picture         =   "Form2设置界面.frx":629BD
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   6315
   End
   Begin VB.Image 装饰框 
      Height          =   495
      Index           =   1
      Left            =   720
      Picture         =   "Form2设置界面.frx":62A03
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   6315
   End
   Begin VB.Image 装饰框 
      Height          =   495
      Index           =   5
      Left            =   720
      Picture         =   "Form2设置界面.frx":62A49
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   6315
   End
   Begin VB.Image 装饰框 
      Height          =   495
      Index           =   4
      Left            =   720
      Picture         =   "Form2设置界面.frx":62A8F
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   6315
   End
   Begin VB.Image 装饰框 
      Height          =   495
      Index           =   8
      Left            =   8400
      Picture         =   "Form2设置界面.frx":62AD5
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   6315
   End
   Begin VB.Image 装饰框 
      Height          =   495
      Index           =   2
      Left            =   720
      Picture         =   "Form2设置界面.frx":62B1B
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   6315
   End
   Begin VB.Image 装饰框 
      Height          =   495
      Index           =   3
      Left            =   720
      Picture         =   "Form2设置界面.frx":62B61
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   6315
   End
End
Attribute VB_Name = "Form2设置界面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tv As New TVEngine '调用tv3d所必需的
Dim Inp As New TVInputEngine
Public Mx As Long, My As Long, B1 As Boolean, B2 As Boolean, Roll As Long   '接收鼠标信息
Dim 临时输入值 As Long
Public Function 刷新设置()
On Error Resume Next
Dim NR As String
Open App.Path & "\Save\Options.ini" For Input As #2
For i = 0 To UBound(设置)
  Line Input #2, NR
  设置(i) = Val(NR)
  Select Case Lab说明(i).Tag
  Case "开关": Com(i).ListIndex = 设置(i)
  Case "滚动条": HSc(i).value = 设置(i)
  'Case "按键": Tex(i).Text = Asc_to_Key(Val(设置(i)))
  'Case "文本框": Tex(i).Text = 设置文本(i)
  End Select
Next
Close #2
End Function
Private Sub Cmd_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
BG.BGM(9).Controls.Play
Cmd(index).Picture = LoadPicture(App.Path & "\Pic\System\UI\cmd1.gif")
Cmd(index).WhatsThisHelpID = 1: Timer1.Enabled = True
Select Case index
Case 0 '返回
  'BASSset 0, 0
  Unload Me
Case 1 '保存
  Open App.Path & "\Save\Options.ini" For Output As #1
  For i = 0 To UBound(设置)
    Select Case Lab说明(i).Tag
    Case "开关": Print #1, Com(i).ListIndex
    Case "滚动条": Print #1, HSc(i).value
    'Case "按键": Print #1, Key_to_Asc(Tex(i).Text)
    'Case "文本框": Print #1, Tex(i).Text
    End Select
  Next
  Close #1
  刷新设置
End Select
End Sub
Private Sub Cmd_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Lab按钮(index).ForeColor = vbWhite
End Sub
Private Sub Form_Load()
On Error Resume Next
Inp.Initialize '初始化按键检测
刷新设置
Tex焦点.Left = Me.Width
For i = 装饰框.lbound To 装饰框.UBound
  装饰框(i).Picture = LoadPicture(App.Path & "\Pic\System\UI\cmdBG.gif")
Next
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
临时输入值 = KeyAscii
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For woc = Cmd.lbound To Cmd.UBound: Lab按钮(woc).ForeColor = &HC0C0C0: Next
End Sub
Private Sub HSc_Scroll(index As Integer)
Lab数值(index) = HSc(index).value
End Sub
Private Sub HSc_Change(index As Integer)
Lab数值(index) = HSc(index).value \ 10
End Sub
Private Sub Lab按钮_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Cmd_MouseDown(index, Button, Shift, X, Y)
End Sub
Private Sub Lab按钮_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Lab按钮(index).ForeColor = vbWhite
End Sub
Private Sub Tex_Click(index As Integer)
On Error Resume Next
If Lab说明(index).Tag <> "按键" Then Exit Sub
MsgBox "点击“确定”后按下您想应用的新键位" & vbCrLf & "Press your new key after clicking ""确定"""
临时输入值 = 0
Do Until 临时输入值 > 0
  DoEvents
Loop
'Tex(Index).Text = Asc_to_Key(临时输入值)
Tex焦点.SetFocus
End Sub
Private Sub Tex无用_Click(index As Integer)
MsgBox "此项内容仅供查看，不可修改" & vbCrLf & "（其实是为了美观凑数qwq）"
End Sub
Private Sub Timer1_Timer()
For woc = Cmd.lbound To Cmd.UBound
If Cmd(woc).WhatsThisHelpID = 1 Then
  Cmd(woc).Picture = LoadPicture(App.Path & "\Pic\System\UI\cmd0.gif")
  Timer1.Enabled = False
End If
Next
End Sub

