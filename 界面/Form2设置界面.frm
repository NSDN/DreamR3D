VERSION 5.00
Begin VB.Form Form2���ý��� 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "ѡ������"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "Form2���ý���.frx":0000
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox Com 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   6
      ItemData        =   "Form2���ý���.frx":6248F
      Left            =   4800
      List            =   "Form2���ý���.frx":62499
      TabIndex        =   28
      Tag             =   "����"
      Text            =   "WMP(����,����)"
      ToolTipText     =   "Ч�ʺܸߣ���������"
      Top             =   7440
      Width           =   1935
   End
   Begin VB.ComboBox Com 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   4
      ItemData        =   "Form2���ý���.frx":624C1
      Left            =   5400
      List            =   "Form2���ý���.frx":624CB
      TabIndex        =   25
      Tag             =   "����"
      Text            =   "��������"
      ToolTipText     =   "Ч�ʺܸߣ���������"
      Top             =   5280
      Width           =   1335
   End
   Begin VB.ComboBox Com 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   5
      ItemData        =   "Form2���ý���.frx":624E2
      Left            =   13320
      List            =   "Form2���ý���.frx":624EC
      TabIndex        =   23
      Tag             =   "����"
      Text            =   "   F��"
      ToolTipText     =   "�չ���FPS��COD���"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Tex���� 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   2
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "��Alt\Ctrl"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Tex���� 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   1
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "��Shift"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox Tex���� 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "����Ҽ�"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Tex���� 
      Height          =   270
      Left            =   7200
      TabIndex        =   15
      Text            =   "���ݷ������"
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
      ItemData        =   "Form2���ý���.frx":62500
      Left            =   5880
      List            =   "Form2���ý���.frx":6250A
      TabIndex        =   9
      Text            =   "��Low"
      ToolTipText     =   "�����ͻ�˫���Թ���Զ����ͼ"
      Top             =   4080
      Width           =   855
   End
   Begin VB.ComboBox Com 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      ItemData        =   "Form2���ý���.frx":6251D
      Left            =   5880
      List            =   "Form2���ý���.frx":6252A
      TabIndex        =   7
      Tag             =   "����"
      Text            =   "��Off"
      ToolTipText     =   "��ͶӰ�Ķ�̬��Ӱ������"
      Top             =   3480
      Width           =   855
   End
   Begin VB.ComboBox Com 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      ItemData        =   "Form2���ý���.frx":62544
      Left            =   5880
      List            =   "Form2���ý���.frx":62551
      TabIndex        =   6
      Tag             =   "����"
      Text            =   "��Off"
      ToolTipText     =   "Ч�ʺܸߣ���������"
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Lab˵�� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "���ź���"
      BeginProperty Font 
         Name            =   "����"
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
      Tag             =   "����"
      ToolTipText     =   "�����Ѱ�װDX8����ʹ��BASS�����ʸ�������"
      Top             =   7440
      Width           =   900
   End
   Begin VB.Label Lab���� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "����(Sound)"
      BeginProperty Font 
         Name            =   "����"
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
   Begin VB.Label Lab˵�� 
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
      Tag             =   "����"
      ToolTipText     =   "�ϳ������ơ�ά�޵�"
      Top             =   5280
      Width           =   960
   End
   Begin VB.Label Lab��ť 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "����(Save)"
      BeginProperty Font 
         Name            =   "����"
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
      Picture         =   "Form2���ý���.frx":6256B
      Stretch         =   -1  'True
      Top             =   10200
      Width           =   1995
   End
   Begin VB.Label Labû��˵�� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "�����¶�"
      BeginProperty Font 
         Name            =   "����"
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
      ToolTipText     =   "��������鿴�������޸�"
      Top             =   4680
      Width           =   900
   End
   Begin VB.Label Labû��˵�� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "С�ܳ��"
      BeginProperty Font 
         Name            =   "����"
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
      ToolTipText     =   "��������鿴�������޸�"
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label Labû��˵�� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "��ȷ��׼"
      BeginProperty Font 
         Name            =   "����"
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
      ToolTipText     =   "��������鿴�������޸�"
      Top             =   3480
      Width           =   900
   End
   Begin VB.Image װ�ο� 
      Height          =   495
      Index           =   11
      Left            =   8400
      Picture         =   "Form2���ý���.frx":6272B
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   6315
   End
   Begin VB.Image װ�ο� 
      Height          =   495
      Index           =   10
      Left            =   8400
      Picture         =   "Form2���ý���.frx":62771
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   6315
   End
   Begin VB.Image װ�ο� 
      Height          =   495
      Index           =   9
      Left            =   8400
      Picture         =   "Form2���ý���.frx":627B7
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   6315
   End
   Begin VB.Label Lab��ť 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "����(Back)"
      BeginProperty Font 
         Name            =   "����"
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
      Picture         =   "Form2���ý���.frx":627FD
      Stretch         =   -1  'True
      Top             =   10200
      Width           =   1995
   End
   Begin VB.Label Lab˵�� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "�����ȼ�"
      BeginProperty Font 
         Name            =   "����"
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
      Tag             =   "����"
      ToolTipText     =   "�ϳ������ơ�ά�޵�"
      Top             =   2880
      Width           =   900
   End
   Begin VB.Label Lab��ֵ 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
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
   Begin VB.Label Lab˵�� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "��Чϸ��"
      BeginProperty Font 
         Name            =   "����"
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
      Tag             =   "������"
      ToolTipText     =   "��Ч������ϸ��"
      Top             =   4680
      Width           =   900
   End
   Begin VB.Label Lab˵�� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "��ͼ����"
      BeginProperty Font 
         Name            =   "����"
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
      Tag             =   "����"
      ToolTipText     =   "������(��)Զ����ģ�������ٶȿ�"
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label Lab˵�� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "��̬��Ӱ"
      BeginProperty Font 
         Name            =   "����"
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
      Tag             =   "����"
      ToolTipText     =   "ĳЩ����������Ķ�̬��Ӱ"
      Top             =   3480
      Width           =   900
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "���ͣ����ѡ��͵��ڿؼ�������ʾ˵��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ѡ������(Options)"
      BeginProperty Font 
         Name            =   "����"
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
   Begin VB.Label Lab˵�� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "ģ�͹�Ӱ"
      BeginProperty Font 
         Name            =   "����"
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
      Tag             =   "����"
      ToolTipText     =   "��������������ȹ�Ӱ"
      Top             =   2880
      Width           =   900
   End
   Begin VB.Label Lab���� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "����(Control)"
      BeginProperty Font 
         Name            =   "����"
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
   Begin VB.Label Lab���� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "����(Video)"
      BeginProperty Font 
         Name            =   "����"
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
   Begin VB.Image װ�ο� 
      Height          =   495
      Index           =   0
      Left            =   720
      Picture         =   "Form2���ý���.frx":629BD
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   6315
   End
   Begin VB.Image װ�ο� 
      Height          =   495
      Index           =   1
      Left            =   720
      Picture         =   "Form2���ý���.frx":62A03
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   6315
   End
   Begin VB.Image װ�ο� 
      Height          =   495
      Index           =   5
      Left            =   720
      Picture         =   "Form2���ý���.frx":62A49
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   6315
   End
   Begin VB.Image װ�ο� 
      Height          =   495
      Index           =   4
      Left            =   720
      Picture         =   "Form2���ý���.frx":62A8F
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   6315
   End
   Begin VB.Image װ�ο� 
      Height          =   495
      Index           =   8
      Left            =   8400
      Picture         =   "Form2���ý���.frx":62AD5
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   6315
   End
   Begin VB.Image װ�ο� 
      Height          =   495
      Index           =   2
      Left            =   720
      Picture         =   "Form2���ý���.frx":62B1B
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   6315
   End
   Begin VB.Image װ�ο� 
      Height          =   495
      Index           =   3
      Left            =   720
      Picture         =   "Form2���ý���.frx":62B61
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   6315
   End
End
Attribute VB_Name = "Form2���ý���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tv As New TVEngine '����tv3d�������
Dim Inp As New TVInputEngine
Public Mx As Long, My As Long, B1 As Boolean, B2 As Boolean, Roll As Long   '���������Ϣ
Dim ��ʱ����ֵ As Long
Public Function ˢ������()
On Error Resume Next
Dim NR As String
Open App.Path & "\Save\Options.ini" For Input As #2
For i = 0 To UBound(����)
  Line Input #2, NR
  ����(i) = Val(NR)
  Select Case Lab˵��(i).Tag
  Case "����": Com(i).ListIndex = ����(i)
  Case "������": HSc(i).value = ����(i)
  'Case "����": Tex(i).Text = Asc_to_Key(Val(����(i)))
  'Case "�ı���": Tex(i).Text = �����ı�(i)
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
Case 0 '����
  'BASSset 0, 0
  Unload Me
Case 1 '����
  Open App.Path & "\Save\Options.ini" For Output As #1
  For i = 0 To UBound(����)
    Select Case Lab˵��(i).Tag
    Case "����": Print #1, Com(i).ListIndex
    Case "������": Print #1, HSc(i).value
    'Case "����": Print #1, Key_to_Asc(Tex(i).Text)
    'Case "�ı���": Print #1, Tex(i).Text
    End Select
  Next
  Close #1
  ˢ������
End Select
End Sub
Private Sub Cmd_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Lab��ť(index).ForeColor = vbWhite
End Sub
Private Sub Form_Load()
On Error Resume Next
Inp.Initialize '��ʼ���������
ˢ������
Tex����.Left = Me.Width
For i = װ�ο�.lbound To װ�ο�.UBound
  װ�ο�(i).Picture = LoadPicture(App.Path & "\Pic\System\UI\cmdBG.gif")
Next
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
��ʱ����ֵ = KeyAscii
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For woc = Cmd.lbound To Cmd.UBound: Lab��ť(woc).ForeColor = &HC0C0C0: Next
End Sub
Private Sub HSc_Scroll(index As Integer)
Lab��ֵ(index) = HSc(index).value
End Sub
Private Sub HSc_Change(index As Integer)
Lab��ֵ(index) = HSc(index).value \ 10
End Sub
Private Sub Lab��ť_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Cmd_MouseDown(index, Button, Shift, X, Y)
End Sub
Private Sub Lab��ť_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Lab��ť(index).ForeColor = vbWhite
End Sub
Private Sub Tex_Click(index As Integer)
On Error Resume Next
If Lab˵��(index).Tag <> "����" Then Exit Sub
MsgBox "�����ȷ������������Ӧ�õ��¼�λ" & vbCrLf & "Press your new key after clicking ""ȷ��"""
��ʱ����ֵ = 0
Do Until ��ʱ����ֵ > 0
  DoEvents
Loop
'Tex(Index).Text = Asc_to_Key(��ʱ����ֵ)
Tex����.SetFocus
End Sub
Private Sub Tex����_Click(index As Integer)
MsgBox "�������ݽ����鿴�������޸�" & vbCrLf & "����ʵ��Ϊ�����۴���qwq��"
End Sub
Private Sub Timer1_Timer()
For woc = Cmd.lbound To Cmd.UBound
If Cmd(woc).WhatsThisHelpID = 1 Then
  Cmd(woc).Picture = LoadPicture(App.Path & "\Pic\System\UI\cmd0.gif")
  Timer1.Enabled = False
End If
Next
End Sub

