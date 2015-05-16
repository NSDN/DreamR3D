VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form BG 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.FileListBox FileList 
      Height          =   270
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin WMPLibCtl.WindowsMediaPlayer BGM 
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   255
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   -1  'True
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   450
   End
   Begin WMPLibCtl.WindowsMediaPlayer BGM 
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   450
   End
   Begin WMPLibCtl.WindowsMediaPlayer BGM 
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   450
   End
   Begin WMPLibCtl.WindowsMediaPlayer BGM 
      Height          =   255
      Index           =   7
      Left            =   1200
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   450
   End
   Begin WMPLibCtl.WindowsMediaPlayer BGM 
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   450
   End
   Begin WMPLibCtl.WindowsMediaPlayer BGM 
      Height          =   255
      Index           =   5
      Left            =   3960
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   255
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   450
   End
   Begin WMPLibCtl.WindowsMediaPlayer BGM 
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   255
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   450
   End
   Begin WMPLibCtl.WindowsMediaPlayer BGM 
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   255
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   450
   End
   Begin WMPLibCtl.WindowsMediaPlayer BGM 
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   255
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   450
   End
   Begin WMPLibCtl.WindowsMediaPlayer BGM 
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   255
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   450
   End
End
Attribute VB_Name = "BG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
On Error Resume Next
BG.BGM(0).url = App.Path & "\Audio\BGM\Title.mp3"
BG.BGM(9).url = App.Path & "\Pic\System\UI\Click0.wav"
'loadSkin Get_OldSkin
For i = 0 To 5
  BGM(i).settings.playCount = 999999999
Next
For i = 0 To BGM.UBound
  BGM(i).settings.volume = 100
Next
For i = 0 To FileList(0).ListCount - 1
 If Len(FileList(0).List(i)) > 3 Then
  If Right(FileList(0).List(i), 3) = "dmp" Then Kill App.Path & "\" & FileList(0).List(i)
 End If
Next
Form1主界面.Show
End Sub

Private Sub Form_Load()
On Error Resume Next
FileList(0).Path = App.Path
With Me
.Width = Screen.Width
.Height = Screen.Height
.Left = 0
.top = 0
End With
End Sub
