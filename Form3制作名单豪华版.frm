VERSION 5.00
Begin VB.Form Form3制作名单豪华版 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "制作名单豪华版"
   ClientHeight    =   7620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   508
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form3制作名单豪华版"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mesh(0 To 10) As TVMesh, MeshVisible(0 To 99) As Boolean
Dim Enemy(0 To 9) As TVActor, EnmVisible(0 To 99) As Boolean
'——普通变量声明——
Dim T总 As Long
Dim LRCname As String, LRCtext As String, LRCcolor As Single, LRC消失时间 As Long
Dim 剧情进度 As Long: Dim 难度 As Long
Dim 调试模式 As Boolean
Dim CLRC完整(1 To 5) As String, CLRC实际(1 To 5) As String
Dim CLRCx As Long, CLRCy As Long, CLRCcolor As Single
Private Function CameraMoveTo(X As Long, Y As Long, z As Long, AngX As Long, AngY As Long, Speed As Single)
CameraPozX = CameraPozX + Speed * (X - CameraPozX)
CameraPozY = CameraPozY + Speed * (Y - CameraPozY)
CameraPozZ = CameraPozZ + Speed * (z - CameraPozZ)
CameraAngX = CameraAngX Mod 360: CameraAngY = CameraAngY Mod 360
CameraAngX = CameraAngX + Speed * (AngX - CameraAngX)
CameraAngY = CameraAngY + Speed * (AngY - CameraAngY)
End Function
Private Function CLRC(X As Long, Y As Long, 首行 As String, 颔行 As String, 颈行 As String, 尾行 As String, 末行 As String)
For j = 1 To 5: CLRC实际(j) = "": Next
CLRCx = X: CLRCy = Y
CLRC完整(1) = 首行: CLRC完整(2) = 颔行: CLRC完整(3) = 颈行: CLRC完整(4) = 尾行: CLRC完整(5) = 末行
End Function
Private Function DrawCLRC(颜色 As Single)
Dim HavePlayed As Boolean
For j = 1 To 5
  If Len(CLRC实际(j)) < Len(CLRC完整(j)) Then
    CLRC实际(j) = Left(CLRC完整(j), 1 + Len(CLRC实际(j)))
    If HavePlayed = False Then
    HavePlayed = True
    End If
  End If
Next
Lrc.Action_BeginText
Lrc.NormalFont_DrawText CLRC实际(1), CLRCx - Len(CLRC实际(1)) * 14, CLRCy, 颜色, 1
Lrc.NormalFont_DrawText CLRC实际(2), CLRCx - Len(CLRC实际(2)) * 14, CLRCy + 50, 颜色, 1
Lrc.NormalFont_DrawText CLRC实际(3), CLRCx - Len(CLRC实际(3)) * 14, CLRCy + 100, 颜色, 1
Lrc.NormalFont_DrawText CLRC实际(4), CLRCx - Len(CLRC实际(4)) * 14, CLRCy + 150, 颜色, 1
Lrc.NormalFont_DrawText CLRC实际(5), CLRCx - Len(CLRC实际(5)) * 14, CLRCy + 200, 颜色, 1
Lrc.Action_EndText
End Function
Private Function CreatLRC(名字 As String, 内容 As String, 名字颜色 As Single)
LRCname = 名字: LRCtext = 内容: LRCcolor = 名字颜色: LRC消失时间 = 3000
End Function
Private Function DrawLRC(名字 As String, 内容 As String, 名字颜色 As Single)
Dim LSX As Long
LSX = 准星X - Len(名字 & 内容) * 10 - 10
Lrc.NormalFont_DrawText 名字, LSX, Me.Height \ 15 - 100, 名字颜色, 1
Lrc.NormalFont_DrawText 内容, LSX + Len(名字) * 20 + 10, Me.Height \ 15 - 100, RGBA(1, 1, 1, 1), 1
End Function
Private Function 执行动作(对象类型 As Long, 对象编号 As Long, 动作名 As String, 播放速度 As Single, 是否循环 As Boolean)
Select Case 对象类型
Case 0 '玩家和队友
  Player(对象编号).SetAnimationByName 动作名
  Player(对象编号).SetAnimationLoop 是否循环
  Player(对象编号).PlayAnimation 播放速度
Case 1 '敌人
  Enemy(对象编号).SetAnimationByName 动作名
  Enemy(对象编号).SetAnimationLoop 是否循环
  Enemy(对象编号).PlayAnimation 播放速度
End Select
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 27 'ESC退出
  Unload Me
Case Asc("f") Or Asc("F")
  Open App.Path & "\时间轴.ini" For Append As #1
    Print #1, T总
  Close #1
End Select
End Sub

'============================声明结束===============================
Private Sub Form_Load()
On Error Resume Next
Randomize
'=====特殊设置=====
BASSready Me.hWnd
SE.Init Me.hWnd
'=====变量设置=====
调试模式 = True
If 调试模式 = True Then
  Tv.SetDebugFile "TV3D运行日志.txt"
  Tv.SetDebugMode True, True
End If
初始化视角参数 0, 0, 0, 0, 100
With Me
  .Width = 15360
  .Height = 11520
  .Left = 0
  .top = 0
  .Show '显示当前窗口，每次都加上错不了
End With
准星X = Me.Width \ 30: 准星Y = Me.Height \ 30
适配X = Me.Width / 15360: 适配Y = Me.Height / 11520
Tv.SetSearchDirectory App.Path & "\" '设定贴图读取目录为当前目录
Tv.SetVSync True '垂直同步开关
Tv.Init3DWindowed Me.hWnd '用窗口模式启动tv3d
Tv.ShowWinCursor False '隐藏鼠标
Inp.Initialize '初始化按键检测
Tv.SetAngleSystem TV_ANGLE_DEGREE
Scene.SetViewFrustum 45, 0    '可视范围，可视角度45
'=====路点=====
'=====贴图=====
TF.LoadTexture "Pic\Flash\white.jpg", "w" '天空盒
TF.LoadTexture "Pic\Flash\black.jpg", "b"
TF.LoadTexture "Pic\Flash\flash00.png", "flash00", , , TV_COLORKEY_USE_ALPHA_CHANNEL '枪火
TF.LoadTexture "Pic\Flash\flash01.png", "flash01", , , TV_COLORKEY_USE_ALPHA_CHANNEL
TF.LoadTexture "Pic\Flash\black.jpg", "height" '高度图

'=====天空特效=====
Atmos.SkyBox_Enable True '开启天空盒
Atmos.SkyBox_SetTexture GetTex("w"), GetTex("w"), GetTex("w"), GetTex("w"), GetTex("w"), GetTex("w") '设定贴图
Atmos.Fog_Enable False
'=====材质=====
MF.CreateMaterial "solid" '建立名为solid的材质
MF.SetAmbient GetMat("solid"), 0.8, 0.8, 0.8, 1    '环境光
MF.SetDiffuse GetMat("solid"), 1, 1, 1, 1 '扩散光，即物体的固有颜色
MF.SetEmissive GetMat("solid"), 0, 0, 0, 1  '自发光
MF.SetOpacity GetMat("solid"), 1  '不透明度
MF.SetSpecular GetMat("solid"), 0, 0, 0, 0 '高光色
MF.SetPower GetMat("solid"), 60 '散射强度

MF.CreateMaterial "map" '建立地图高光材质
MF.SetAmbient GetMat("map"), 0.8, 0.8, 0.8, 1   '环境光
MF.SetDiffuse GetMat("map"), 1, 1, 1, 1 '扩散光，即物体的固有颜色
MF.SetEmissive GetMat("map"), 1, 1, 1, 1 '自发光
MF.SetOpacity GetMat("map"), 1  '不透明度
MF.SetSpecular GetMat("map"), 1, 1, 1, 1 '高光色
MF.SetPower GetMat("map"), 15 '散射强度
'=====光影=====
'光影
LE.CreateDirectionalLight Vector(1, -1, 1), 1, 1, 1, "sun", 1 '添加一个平行光
If 设置(0) > 0 Then
  LE.SetSpecularLighting True '高光开关
  LE.SetLightProperties 0, True, False, False '灯光开启影子
End If
'=====角色=====
For i = 0 To UBound(Mesh): Set Mesh(i) = Scene.CreateMeshBuilder: Next
With Mesh(0)
.LoadTVM "Model\ZBD05\ZBD05.tvm", True, True
.SetScale 1, 1, 1
.SetPosition 10, 2, 15
.SetRotation 0, 180, 0
.SetLightingMode TV_LIGHTING_NORMAL
.SetTexture GetTex("b")
End With

For i = 0 To UBound(Enemy): Set Enemy(i) = Scene.CreateActor: Next '角色初始化
With Enemy(0)
.LoadTVA "Model\M1坦克\M1坦克.tva", True, True
.SetScale 0.03, 0.03, 0.03
.SetPosition 1, -5, 5
.SetRotation 0, -55, 0
End With
EnmVisible(0) = True

For i = 2 To 3
With Enemy(i)
.LoadTVA "Model\黑鹰\黑鹰.tva", True, True
.SetScale 0.02, 0.02, 0.02
.SetRotation 0, 90, 5
End With
执行动作 1, i, "idle", 1, True
Next
Enemy(2).SetPosition 16, 6, 5
Enemy(3).SetPosition 12, 2, 9

For i = 0 To UBound(Enemy)
With Enemy(i)
.SetLightingMode TV_LIGHTING_NORMAL
.SetTexture GetTex("b")
End With
Next
'=====特效====
'=====参数=====
Lrc.NormalFont_Create "", "隶书", 35, False, False, False
初始化视角参数 0, 0, 0, 0, 9999
PlayerHeight = 2: 游戏速度 = 1
BG.BGM(0).Controls.stop
BG.BGM(8).url = App.Path & "\Audio\BGM\Senya-Bad Apple.mp3"
BG.BGM(8).Controls.currentPosition = 0
CLRCcolor = RGBA(0, 0, 0, 1)
'================================主循环=======================================
Do
 Inp.GetMouseState Mx, My, B1, B2, , , Roll           '接收鼠标信息
 If 调试模式 = True Then
   CameraAngX = CameraAngX + 0.1 * Mx * 游戏速度
   CameraAngY = CameraAngY + 0.1 * My * 游戏速度
 End If
视角坐标更新 True
'设定摄像机
Camera.SetPosition CameraPozX, CameraPozY + PlayerHeight + CameraPoz偏移(2), CameraPozZ
Camera.SetRotation CameraAngY - 后坐力(1), CameraAngX + 后坐力(0), 0
'===============清屏与渲染===============
Tv.Clear '清屏
Atmos.Fog_Enable False
Atmos.Atmosphere_Render '渲染大气
Atmos.Fog_Enable True
For i = 0 To UBound(Enemy)
  If EnmVisible(i) = True Then Enemy(i).Render
Next
For i = 0 To UBound(Mesh)
  If MeshVisible(i) = True Then Mesh(i).Render
Next
Scene.FinalizeShadows '渲染影子
'===============角色事件===============

跳过此对象:
Select Case 剧情进度
End Select
'===============文字渲染===============
Lrc.NormalFont_DrawText T总, 10, 10, RGBA(1, 0, 0, 1), 1
If LRC消失时间 > 0 Then
  DrawLRC LRCname, LRCtext, LRCcolor
  LRC消失时间 = LRC消失时间 - Tv.TimeElapsed
End If
DrawCLRC CLRCcolor
Tv.RenderToScreen
DoEvents
Loop
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Shell App.Path & "\" & App.EXEName & ".exe 0", vbNormalFocus
End
End Sub
Private Sub Timer1_Timer()
T总 = 10 * BG.BGM(8).Controls.currentPosition
Select Case T总
Case 22: CLRC Me.Width \ 30, Me.Height \ 30 - 150, "-PV版制作名单-", "佐亚_洛克k上尉", "", "", ""
Case 90
  CLRC Me.Width \ 30, Me.Height \ 30 - 150, "-PV版制作名单-", "佐亚_洛克k上尉", "    （向Bad Apple致敬）", "-BGM-", "幽闭星光（Senya）"
  For j = 1 To 2
    CLRC实际(j) = CLRC完整(j)
  Next
Case 165:: GF.Flash 0, 0, 0, 500
Case 173:: GF.Flash 0, 0, 0, 500
Case 178:: GF.Flash 0, 0, 0, 500
Case 165: GF.Flash 0, 0, 0, 500
Case 200: GF.Flash 0, 0, 0, 500
Case 235: GF.Flash 0, 0, 0, 500
Case 270: GF.Flash 0, 0, 0, 500
Case 282: GF.Flash 0, 0, 0, 500
Case 298: CLRC -9999, 0, "", "", "", "", "": 剧情进度 = 1
Case 368: 剧情进度 = 3
Case 430: CLRC -9999, 0, "", "", "", "", "": 剧情进度 = 4
Case 440: CLRC Me.Width \ 30 + 300, 80, "-主程序-", "佐亚_洛克k上尉", "木龙华易", "        Drzzm32", "       Reity": EnmVisible(2) = True: EnmVisible(3) = True
Case 510: 剧情进度 = 5: EnmVisible(0) = False: CLRC -9999, 0, "", "", "", "", ""
Case 580: 剧情进度 = 6: CLRC Me.Width \ 30, Me.Height \ 30 - 150, "  -3D美术-", "木龙华易", "  -2D美术-", "佐亚_洛克k上尉", "木龙华易"
Case 718: GF.Flash 0, 0, 0, 500
Case 788: GF.Flash 0, 0, 0, 500
Case 848: GF.Flash 0, 0, 0, 500
Case 855: 剧情进度 = 7: GF.Flash 0, 0, 0, 500: CLRC -9999, 0, "", "", "", "", "": MeshVisible(0) = True: EnmVisible(2) = False: EnmVisible(3) = False
Case 870: 剧情进度 = 8
Case 3780: Unload Me
End Select
'————————
Select Case 剧情进度
Case 1 '坦克浮起
  If Enemy(0).GetPosition.Y < 2 Then
    Enemy(0).SetPosition Enemy(0).GetPosition.X, Enemy(0).GetPosition.Y + 0.3, Enemy(0).GetPosition.z
  Else
    CLRC Me.Width \ 30 - 300, Me.Width \ 30 - 200, "-策划-", "佐亚_洛克k上尉", "木龙华易", "吃桃的叫天子", ""
    剧情进度 = 2
  End If
Case 2: Enemy(0).SetPosition 0.01 + Enemy(0).GetPosition.X, Enemy(0).GetPosition.Y, Enemy(0).GetPosition.z
Case 3: CameraMoveTo 0, 3, 10, 180, 40, 0.019
Case 4: Enemy(0).SetPosition Enemy(0).GetPosition.X - 0.05, Enemy(0).GetPosition.Y, Enemy(0).GetPosition.z - 0.048
Case 5: CameraMoveTo 0, 3, 10, 85, 0, 0.08
Case 6 '直升机飞离
  Enemy(2).SetPosition Enemy(2).GetPosition.X, Enemy(2).GetPosition.Y, Enemy(2).GetPosition.z + 0.2
  Enemy(3).SetPosition Enemy(3).GetPosition.X, Enemy(3).GetPosition.Y, Enemy(3).GetPosition.z + 0.2
Case 7 '装甲车驶入
  Mesh(0).SetPosition Mesh(0).GetPosition.X, Mesh(0).GetPosition.Y, Mesh(0).GetPosition.z - 0.02
Case 8: Mesh(0).RotateY 0.01
End Select
End Sub

