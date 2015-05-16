VERSION 5.00
Begin VB.Form Form3制作名单 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "制作名单"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3240
      Top             =   0
   End
   Begin VB.ListBox List1 
      Height          =   1320
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.ListBox List 
      Height          =   1320
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label CmdSkip 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   ">>跳过Skip"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   13200
      TabIndex        =   1
      Top             =   10920
      Width           =   1800
   End
   Begin VB.Label Lrc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "感谢您的试玩！Thanks for playing!"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
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
      Left            =   3480
      TabIndex        =   0
      Top             =   11040
      Width           =   7845
   End
End
Attribute VB_Name = "Form3制作名单"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ending As Long
Private Sub CmdSkip_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Dim NR As String
Lrc(0) = ""
Open App.Path & "\Pic\System\Credits.ini" For Input As #1
Do Until EOF(1)
  Line Input #1, NR
  List1.AddItem NR
Loop
Close #1
BG.BGM(0).url = App.Path & "\Audio\BGM\Credits.mp3"
For i = 1 To 10
  Load Lrc(i)
  With Lrc(i)
  .Left = Lrc(0).Left
  .top = Lrc(i - 1).top + 80
  .Caption = List1.List(0)
  .Visible = True
  End With
  List1.RemoveItem (0)
Next
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
For i = 1 To Lrc.UBound
  Lrc(i).top = Lrc(i).top - 2.2
  If Lrc(i).top < -Lrc(i).Height Then
    If List1.ListCount > 0 Then
      Lrc(i).Caption = List1.List(0)
      List1.RemoveItem (0)
    Else
      Lrc(i).Caption = ""
      Ending = 1 + Ending
      If Ending >= Lrc.UBound Then Unload Me
    End If
    Lrc(i).top = Me.ScaleHeight
  End If
Next
End Sub
