VERSION 5.00
Begin VB.Form Form1������ 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "������"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.Label Labװ�� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "����������3D"
      BeginProperty Font 
         Name            =   "����_GB2312"
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
   Begin VB.Label Labѡ�� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "English"
      BeginProperty Font 
         Name            =   "����"
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
   Begin VB.Label Labѡ�� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "����ģʽ"
      BeginProperty Font 
         Name            =   "����_GB2312"
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
   Begin VB.Label Labѡ�� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "����ģʽ"
      BeginProperty Font 
         Name            =   "����_GB2312"
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
   Begin VB.Label Labѡ�� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����_GB2312"
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
   Begin VB.Label Labѡ�� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "������ʾ���������"
      BeginProperty Font 
         Name            =   "����"
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
   Begin VB.Label Labѡ�� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "ѡ������"
      BeginProperty Font 
         Name            =   "����_GB2312"
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
   Begin VB.Label Labѡ�� 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "�˳���Ϸ"
      BeginProperty Font 
         Name            =   "����_GB2312"
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
Attribute VB_Name = "Form1������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function �л���������(���� As Long)
����(4) = ����
Select Case ����
Case 0 '��������
  Labѡ��(0).Caption = "������ʾ���������"
  Labѡ��(1).Caption = "�˳���Ϸ"
  Labѡ��(2).Caption = "ѡ������"
  Labѡ��(3).Caption = "English"
  Labѡ��(4).Caption = "��������"
  Labѡ��(5).Caption = "����ģʽ"
  Labѡ��(6).Caption = "����ģʽ"
  Labװ��(0).Caption = "�����ִ�ս��"
Case 1 'English
  Labѡ��(0).Caption = "Click here if the words indicate wrong"
  Labѡ��(1).Caption = "Exit"
  Labѡ��(2).Caption = "Options"
  Labѡ��(3).Caption = "��������"
  Labѡ��(4).Caption = "Credits"
  Labѡ��(5).Caption = "Survival Mode"
  Labѡ��(6).Caption = "Story Mode"
  Labװ��(0).Caption = "Touhou Modern War"
End Select
End Function
Private Sub Form_Activate()
On Error Resume Next
���Ķ�˵�� = True
Dir1.Path = App.Path
Dir1.Path = App.Path
��ȡ����
�л��������� ����(4)
'����������������
End Sub

Private Sub Labѡ��_Click(index As Integer)
BG.BGM(9).Controls.Play
Select Case index
Case 0 'ϵͳδ��װ��������
  For i = 0 To Labѡ��.UBound: Labѡ��(i).Font = "����": Next
  Labѡ��(0).Caption = "���л������֣����鰲װ��������ϵͳ"
Case 1: Unload BG: Unload Me: End
Case 2: Form2���ý���.Show , vbNormalFocus
Case 3 '�������л�
If Labѡ��(3).Caption = "English" Then
  �л��������� 1
Else
  �л��������� 0
End If
Case 4: Form3��������������.Show , vbNormalFocus
Case 5 '����ģʽ
  BG.BGM(0).url = ""
  Stage����.Show , vbNormalFocus
Case 6 '����ģʽ
  BG.BGM(0).url = ""
  Stage01.Show , vbNormalFocus
End Select
End Sub

Private Sub Labѡ��_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Labѡ��(index).ForeColor = vbWhite
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To Labѡ��.UBound
Labѡ��(i).ForeColor = &HC0C0C0
Next
End Sub
