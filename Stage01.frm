VERSION 5.00
Begin VB.Form Stage01 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "东方现代战争：绯红破晓5  ——九州烽火工作室"
   ClientHeight    =   5460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   364
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   467
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Tim敌人重生 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Stage01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mesh(0) As TVLandscape: Dim MeshTVA(0 To 2) As TVActor: Dim MeshSin(0 To 0) As TVMesh
Dim Enemy(1 To 35) As TVActor, EnemyGun(1 To 25) As TVActor, EnmType(1 To 25) As Long, EnmGunFire(1 To 25) As Long
Dim EnmLastView(1 To 25) As TV_3DVECTOR
Dim 角色位置(0 To 3) As TV_3DVECTOR '0玩家12角色3中转
Dim 角色方向(0 To 3) As TV_3DVECTOR
Dim VoiceT As Long, T总 As Long
'——普通变量声明——
Dim LRCname As String, LRCtext As String, LRCcolor As Single, LRC消失时间 As Long
Dim 剧情进度 As Long: Dim 难度 As Long
Dim 调试模式 As Boolean
Dim CLRC完整(1 To 4) As String, CLRC实际(1 To 4) As String, CLRCt(0 To 1) As Long
Private Function CLRC(首行 As String, 颔行 As String, 颈行 As String, 尾行 As String)
For j = 1 To 4: CLRC实际(j) = "": Next
CLRC完整(1) = 首行: CLRC完整(2) = 颔行: CLRC完整(3) = 颈行: CLRC完整(4) = 尾行
CLRCt(0) = 5000
End Function
Private Function DrawCLRC(颜色 As Single)
Dim HavePlayed As Boolean
If CLRCt(0) <= 0 Then Exit Function
CLRCt(1) = CLRCt(1) + Tv.TimeElapsed
CLRCt(0) = CLRCt(0) - Tv.TimeElapsed
If CLRCt(1) > 100 Then
For j = 1 To 4
  If Len(CLRC实际(j)) < Len(CLRC完整(j)) Then
    CLRC实际(j) = Left(CLRC完整(j), 1 + Len(CLRC实际(j)))
    If HavePlayed = False Then
    SEplay "Type0.wav", False
    HavePlayed = True
    End If
  End If
Next
CLRCt(1) = 0
End If
Lrc.Action_BeginText
Lrc.NormalFont_DrawText CLRC实际(1), 准星X - Len(CLRC实际(1)) * 10, 准星Y - 200, 颜色, 1
Lrc.NormalFont_DrawText CLRC实际(2), 准星X - Len(CLRC实际(2)) * 10, 准星Y - 150, 颜色, 1
Lrc.NormalFont_DrawText CLRC实际(3), 准星X - Len(CLRC实际(3)) * 10, 准星Y - 100, 颜色, 1
Lrc.NormalFont_DrawText CLRC实际(4), 准星X - Len(CLRC实际(4)) * 10, 准星Y - 50, 颜色, 1
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
Public Function 物理计算高度(位置 As TV_3DVECTOR, 越野高度 As Single, 下落速度 As Single) As Single
Dim Result As TVCollisionResult
物理计算高度 = Mesh(0).GetHeight(位置.X, 位置.z)
If 物理计算高度 > 位置.Y + 越野高度 Then 物理计算高度 = 位置.Y
If 物理计算高度 < 位置.Y - 下落速度 Then 物理计算高度 = 位置.Y - 下落速度
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
Private Function AImoveTo(对象编号 As Long, 移动类型 As Long, 目标位置 As TV_3DVECTOR, 移动速度 As Single)
Dim 临时坐标(0 To 6) As TV_3DVECTOR: Dim 物理(0 To 6) As Single
Select Case 移动类型
Case 0 '===（敌）面向对象===
    临时坐标(0) = 目标位置
    临时坐标(3) = Enemy(对象编号).GetRotation
    Enemy(对象编号).LookAtPoint 临时坐标(0)
    临时坐标(1) = Enemy(对象编号).GetRotation
    Enemy(对象编号).SetRotation 临时坐标(3).X, 临时坐标(1).Y + 90, 角色方向(3).z
    临时坐标(3) = Enemy(对象编号).GetRotation
   If 移动速度 > 360 Or Abs(角色方向(1).Y - 角色方向(3).Y) <= 移动速度 Then Exit Function
Case 1 '===（敌）直线移动===
    临时坐标(0) = Enemy(对象编号).GetPosition: 临时坐标(0).Y = 临时坐标(0).Y - 12
    物理(1) = 移动速度 * (目标位置.X - 临时坐标(0).X)
    物理(2) = 移动速度 * (目标位置.z - 临时坐标(0).z)
    Enemy(对象编号).SetPosition 临时坐标(0).X + 物理(1), 物理计算高度(临时坐标(0), 1, 3) + 12, 临时坐标(0).z + 物理(2)
    临时坐标(1) = Enemy(对象编号).GetPosition
    If Mesh(0).AdvancedCollide(Vector(临时坐标(1).X + 4, 临时坐标(1).Y, 临时坐标(1).z), Vector(临时坐标(1).X - 4, 临时坐标(1).Y, 临时坐标(1).z)).IsCollision Then Enemy(对象编号).SetPosition 临时坐标(1).X - 物理(1), 临时坐标(1).Y, 临时坐标(1).z
    If Mesh(0).AdvancedCollide(Vector(临时坐标(1).X, 临时坐标(1).Y, 临时坐标(1).z + 4), Vector(临时坐标(1).X, 临时坐标(1).Y, 临时坐标(1).z - 4)).IsCollision Then Enemy(对象编号).SetPosition 临时坐标(1).X, 临时坐标(1).Y, 临时坐标(1).z - 物理(2)
End Select
End Function
Private Function AIadv(对象类型 As Long, 对象编号 As Long, AI类型 As Long, 移动速度 As Single)
Dim 碰撞(0 To 2) As Boolean
Select Case 对象类型
Case 0 '队友
Case 1 '敌人
  角色位置(0) = Player(0).GetPosition
  角色位置(1) = Enemy(对象编号).GetPosition
  Select Case AI类型
  Case 1 '==僵尸==
      Select Case EnmState(对象编号)
      Case 0 '——靠近——
        AImoveTo 对象编号, 0, Player(0).GetPosition, 999
        AImoveTo 对象编号, 1, Player(0).GetPosition, 移动速度
        If 距离平方(Player(0).GetPosition, Enemy(对象编号).GetPosition) < 400 Then
          执行动作 1, 对象编号, "ref_shoot_wrench", 游戏速度, False
          If 血污残留时间 <= 20 Then 玩家受伤 0, 难度
          EnmState(对象编号) = 1
        End If
      Case 1 '——攻击——
        If Enemy(对象编号).IsAnimationFinished Then
          执行动作 1, 对象编号, "run2", 游戏速度, False
          EnmState(对象编号) = 0
        End If
      End Select
  Case 2 '==步兵==
      AImoveTo 对象编号, 1, Player(0).GetPosition, 0
      碰撞(0) = Mesh(0).AdvancedCollide(Enemy(对象编号).GetPosition, Player(0).GetPosition).IsCollision
      If 碰撞(0) = False Then EnmLastView(对象编号) = Vector(CameraPozX + 10 * 物理引擎(3), CameraPozY, CameraPozZ + 10 * 物理引擎(4))
      Select Case EnmState(对象编号)
      Case 0 '——接近玩家——
          If 碰撞(0) = False Then '玩家可视
            AImoveTo 对象编号, 0, Player(0).GetPosition, 999
            If 距离平方(Enemy(对象编号).GetPosition, Player(0).GetPosition) > 6000 Then
              AImoveTo 对象编号, 1, Player(0).GetPosition, 移动速度
            Else
              If VoiceT > 4000 Then SEplay "AI-fire" & CInt(2 * Rnd) & ".wav", True: VoiceT = 0 'voice
              Select Case EnmType(对象编号)
              Case 0: 执行动作 1, 对象编号, "ref_shoot_onehanded", 游戏速度, True
              Case 1: 执行动作 1, 对象编号, "lmg_fire", 游戏速度, True
              End Select
              EnmState(对象编号) = 1
            End If
          Else '玩家不可视
            If EnmLastView(i).X <> 0 Then AImoveTo i, 1, EnmLastView(对象编号), 2 * 移动速度
          End If
      Case 1 '——立正射击——
          If 碰撞(0) = False Then
              AImoveTo 对象编号, 0, Player(0).GetPosition, 999
              If 距离平方(Enemy(对象编号).GetPosition, Player(0).GetPosition) < 9000 Then
                If EnmT(对象编号) Mod 300 < 20 Then '攻击
                    SEplay "掩体着弹0.wav", False
                    SEplay "擦弹" & CInt(Rnd * 6) & ".wav", False
                    If 血污残留时间 <= 60 Then 玩家受伤 0, 难度
                End If '攻击结束
                If EnmT(对象编号) Mod 300 < 90 Then
                  Scr图形.Draw_Line3D Enemy(对象编号).GetPosition.X, Enemy(对象编号).GetPosition.Y + 4, Enemy(对象编号).GetPosition.z, CameraPozX + Rnd * 6 - 3, CameraPozY + Rnd * 6 - 3 + PlayerHeight, CameraPozZ + Rnd * 6 - 3, RGBA(1, 1, 0, 0.3)
                  If 设置(0) > 1 Then LE.CreatePointLight Enemy(对象编号).GetPosition, 1, 1, 0, 50, "EnmGunFire" & 对象编号
                End If
                EnmT(对象编号) = EnmT(对象编号) + Tv.TimeElapsed
                If EnmT(对象编号) > 4000 Then '换弹
                    If VoiceT > 3000 Then SEplay "AI-reload.wav", True: VoiceT = 0 'voice
                    SEplay "AI_reload.wav", False
                    Select Case EnmType(对象编号)
                    Case 0: 执行动作 1, 对象编号, "ref_reload_onehanded", 游戏速度, False
                    Case 1: 执行动作 1, 对象编号, "lmg_reload", 游戏速度, False
                    End Select
                    EnmState(i) = 2
                End If
              Else '距离远，切换追击
                  Select Case EnmType(对象编号)
                  Case 0: 执行动作 1, 对象编号, "run", 游戏速度, True
                  Case 1: 执行动作 1, 对象编号, "walk", 游戏速度, True
                  End Select
                  EnmState(对象编号) = 0
              End If
          Else '玩家不可视
              If VoiceT > 4000 Then SEplay "AI-go" & CInt(2 * Rnd) & ".wav", True: VoiceT = 0 'voice
              Select Case EnmType(对象编号)
              Case 0: 执行动作 1, 对象编号, "run", 游戏速度, True
              Case 1: 执行动作 1, 对象编号, "walk", 游戏速度, True
              End Select
              EnmLastView(对象编号) = Player(0).GetPosition
              EnmState(对象编号) = 0
          End If
      Case 2 '——换弹匣——
          If Enemy(对象编号).IsAnimationFinished Then
            Select Case EnmType(对象编号)
            Case 0: 执行动作 1, 对象编号, "run", 游戏速度, True
            Case 1: 执行动作 1, 对象编号, "walk", 游戏速度, True
            End Select
            EnmT(对象编号) = 0
            EnmState(对象编号) = 1
          End If
      End Select
  End Select
End Select
End Function
Public Function 玩家受伤(编号 As Long, 伤害 As Long)
If 编号 = 0 Then
  PlayerHP(0) = PlayerHP(0) - 伤害
  If PlayerHeight > 10 Then PlayerHeight = PlayerHeight - 5
  If 后坐力(1) < 2 Then 后坐力(1) = 后坐力(1) - 6
  血污残留时间 = 80
Else
End If
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 27 'ESC退出
  Tv.ShowWinCursor True: If MsgBox("暂停中，要现在放弃战斗吗？", vbYesNo) = vbYes Then Unload Me Else Tv.ShowWinCursor False
Case Asc("h") Or Asc("H") '防卡死
  MsgBox "当前位置：" & Player(0).GetPosition.X & "//" & Player(0).GetPosition.Y & "//" & Player(0).GetPosition.z & vbCrLf
  CameraPozY = 物理计算高度(Vector(CameraPozX, CameraPozY, CameraPozZ), 2000, 3) + PlayerHeight
  视角坐标更新 True '防卡死急救
Case Asc("f") Or Asc("F")

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
调试模式 = False
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
If 设置(2) > 0 Then Scene.SetTextureFilter TV_FILTER_BILINEAR
'=====路点=====
'=====贴图=====
Dim BGname As String: BGname = "sunny"
TF.LoadTexture "Pic\System\血污0.png", "血污0", , , TV_COLORKEY_USE_ALPHA_CHANNEL  '读取名为XX的贴图命名为XX，建立透明Alpha通道
TF.LoadTexture "Pic\BG\" & BGname & "_Back.jpg", "SKYBOX_Back" '天空盒
TF.LoadTexture "Pic\BG\" & BGname & "_Front.jpg", "SKYBOX_Front"
TF.LoadTexture "Pic\BG\" & BGname & "_Left.jpg", "SKYBOX_Left"
TF.LoadTexture "Pic\BG\" & BGname & "_Right.jpg", "SKYBOX_Right"
TF.LoadTexture "Pic\BG\" & BGname & "_Top.jpg", "SKYBOX_Up"
TF.LoadTexture "Pic\BG\" & BGname & "_Down.jpg", "SKYBOX_DOWN"
TF.LoadTexture "Pic\Flash\flash00.png", "flash00", , , TV_COLORKEY_USE_ALPHA_CHANNEL '枪火
TF.LoadTexture "Pic\Flash\flash01.png", "flash01", , , TV_COLORKEY_USE_ALPHA_CHANNEL
TF.LoadTexture "Pic\System\UI武器属性.png", "UI武器属性", , , TV_COLORKEY_USE_ALPHA_CHANNEL 'UI武器属性
TF.LoadTexture "Pic\Flash\sun.jpg", "sun", , , TV_COLORKEY_USE_ALPHA_CHANNEL '太阳
TF.LoadTexture "Pic\Flash\flare1.jpg", "flare1" '太阳光晕
TF.LoadTexture "Pic\Flash\flare2.jpg", "flare2"
TF.LoadTexture "Pic\Flash\flare3.jpg", "flare3"
TF.LoadTexture "Pic\Flash\flare4.jpg", "flare4"
TF.LoadTexture "Pic\Flash\flare4.jpg", "flare4"
TF.LoadTexture "Map\海滩\height.jpg", "height" '高度图
TF.LoadTexture "Map\海滩\mud.jpg", "main"
TF.LoadAlphaTexture "Map\海滩\mask.jpg", "mask" '遮罩图
TF.LoadTexture "Map\海滩\ggrass.dds", "main2"
TF.LoadTexture "Pic\Stage\Fire_smoke.jpg", "fire", , , TV_COLORKEY_MAGENTA '火焰和烟雾

'=====天空特效=====
Atmos.SkyBox_Enable True '开启天空盒
Atmos.SkyBox_SetTexture GetTex("SKYBOX_Front"), GetTex("SKYBOX_Back"), GetTex("SKYBOX_Left"), GetTex("SKYBOX_Right"), GetTex("SKYBOX_Up"), GetTex("SKYBOX_DOWN") '设定贴图
Atmos.Fog_Enable True                              '开启雾
Atmos.Fog_SetColor 0.1, 0.1, 0.1                         '颜色RGBA，例如红
Atmos.Fog_SetParameters 500, 2000, 0              '最近距离，最远距离，浓度
Atmos.Fog_SetType TV_FOG_LINEAR, TV_FOGTYPE_PIXEL  '雾的类型
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
'光晕
Atmos.Sun_SetBillboardSize 0.7 '设置太阳贴图大小
Atmos.Sun_SetPosition -400, 400, -400  '设置太阳位置
Atmos.Sun_SetTexture GetTex("sun") '赋予太阳贴图
If 设置(0) > 1 And 设置(4) >= 30 Then Atmos.Sun_Enable True '使太阳贴图生效
Atmos.LensFlare_SetLensNumber 4 '光晕层数
Atmos.LensFlare_SetLensParams 1, GetTex("flare1"), 7.5, 40, RGBA(1, 1, 1, 0.5), RGBA(1, 1, 1, 0.5)
Atmos.LensFlare_SetLensParams 2, GetTex("flare2"), 3, 18, RGBA(1, 1, 1, 0.5), RGBA(1, 1, 1, 0.5)
Atmos.LensFlare_SetLensParams 3, GetTex("flare3"), 4, 15, RGBA(1, 1, 1, 0.5), RGBA(0.7, 1, 1, 0.5)
Atmos.LensFlare_SetLensParams 4, GetTex("flare4"), 3, 6, RGBA(1, 0.1, 0, 0.5), RGBA(0.5, 1, 1, 0.5)
If 设置(0) > 1 And 设置(4) > 30 Then Atmos.LensFlare_Enable True '使光晕生效
'光影
LE.CreateDirectionalLight Vector(1, -1, 1), 1, 1, 1, "sun", 1 '添加一个平行光
If 设置(0) > 0 Then
  LE.SetSpecularLighting True '高光开关
  LE.SetLightProperties 0, True, True, False '灯光开启影子
End If
'=====场景=====
'带碰撞检测实体
Set Mesh(0) = Scene.CreateLandscape
Mesh(0).GenerateTerrain "Map\海滩\height.jpg", TV_PRECISION_AVERAGE, 8, 8, -256 * 4, 0, -256 * 4, True
Mesh(0).SetTexture GetTex("main")
Mesh(0).SetMaterial GetMat("solid")
Mesh(0).SetTextureScale 4, 4
Mesh(0).SetLightingMode TV_LIGHTING_NORMAL
Mesh(0).SetPosition 0, 0, 0

'带骨骼地图TVA实体
For i = 0 To UBound(MeshTVA): Set MeshTVA(i) = Scene.CreateActor: Next
MeshTVA(0).LoadTVA "Model\M1坦克\M1坦克.tva", True, True
MeshTVA(1).LoadTVA "Model\悍马\悍马.tva", True, True
MeshTVA(2).LoadTVA "Model\M1坦克\M1坦克.tva", True, True
With MeshTVA(0)
  .SetScale 0.33, 0.33, 0.33
  .SetPosition 1650, 物理计算高度(Vector(1650, 0, 1200), 999, 999) + 10, 1200
  .RotateY -120
  .SetMaterial GetMat("map")
End With
With MeshTVA(1)
  .SetScale 0.33, 0.33, 0.33
  .SetPosition 1400, 物理计算高度(Vector(1400, 0, 1500), 999, 999) + 12, 1500
  .SetRotation -90, -120, 0
  .SetMaterial GetMat("solid")
End With
With MeshTVA(2)
  .SetScale 0.33, 0.33, 0.33
  .SetPosition 1200, 物理计算高度(Vector(1200, 0, 1300), 999, 999) + 8, 1300
  .SetRotation 120, -120, 0
  .SetMaterial GetMat("solid")
End With
For i = 0 To UBound(MeshTVA): With MeshTVA(i) '光影
  .SetLightingMode TV_LIGHTING_NORMAL
End With: Next

'=====角色=====
For i = 0 To UBound(Player): Set Player(i) = Scene.CreateActor: Next
If UBound(Player) > 0 Then
For i = 1 To UBound(Player)
With Player(i)
.SetMaterial GetMat("solid") '设定材质
If 设置(0) > 0 Then .SetLightingMode TV_LIGHTING_NORMAL '设定光照模式
.SetAnimationByName ("run") '执行的动作名称
.PlayAnimation 游戏速度 '播放动作速度
.SetScale 0.31, 0.31, 0.31 '设定大小
End With
Next
End If

For i = 1 To UBound(Enemy)
Set Enemy(i) = Scene.CreateActor '角色初始化
With Enemy(i)
.LoadTVA ("Player\重甲\重甲.tva") '读取模型
.SetMaterial GetMat("solid") '设定材质
If 设置(0) > 0 Then .SetLightingMode TV_LIGHTING_NORMAL '设定光照模式
.SetScale 0.28, 0.28, 0.28 '设定大小
.SetPosition Rnd * 300 + 1200, 0, Rnd * 300 + 1200 '设定模型位置
.SetPosition .GetPosition.X, 物理计算高度(Vector(.GetPosition.X, .GetPosition.Y, .GetPosition.z), 999, 999) + 8, .GetPosition.z '设定模型位置
'.SetRotation -90, 0, 180
End With
执行动作 1, i, "death" & (1 + Round(Rnd * 2)) & "_die", 9, False
EnmHP(i) = 0
Next

For i = 1 To UBound(Enemy)
Set EnemyGun(i) = Scene.CreateActor '网格物体初始化，必加
With EnemyGun(i)
'.LoadTVA "Weapon\M16\v_M16.tva", True, True
If 设置(0) > 0 Then .SetLightingMode TV_LIGHTING_NORMAL
.SetMaterial GetMat("solid")
.SetScale 0.075, 0.075, 0.075 '设定大小
End With
Next
'=====特效====
建立水面 800, 1650, 2000, 2000, 72
If 设置(3) >= 50 Then
Set tPart = Scene.CreateParticleSystem '粒子
tPart.Load "Part\火焰1\火焰1.tvp"
tPart.SetGlobalPosition MeshTVA(2).GetPosition.X, MeshTVA(2).GetPosition.Y, MeshTVA(2).GetPosition.z
tPart.SetGlobalScale 0.05, 0.05, 0.05
Set tPart = Scene.CreateParticleSystem '粒子
tPart.Load "Part\烟雾1\烟雾1.tvp"
tPart.SetGlobalPosition MeshTVA(1).GetPosition.X - 1, MeshTVA(1).GetPosition.Y - 3, MeshTVA(1).GetPosition.z - 4
tPart.SetGlobalScale 4, 4, 4
End If
'=====参数=====
Lrc.NormalFont_Create "", "宋体", 25, False, False, False
初始化视角参数 1380, 物理计算高度(Vector(1380, 0, 1680), 999, 999), 1680, 180, 99
PlayerHeight = 2: 游戏速度 = 1
GunLoad "AKS-74U", 1, True
GunLoad "M16", 0, True
BASSplay "Audio\BGM\music_ingame_disaster_and_rescue.mp3", 0, 1, Me.hWnd
BASSplay "Audio\BGS\Battle-far1.mp3", 1, 1, Me.hWnd
开关隐藏武器 = True
Do Until 剧情进度 > 0
  Tv.Clear
  DrawCLRC (RGBA(1, 1, 1, 1)): DrawLRC LRCname, LRCtext, LRCcolor
  Tv.RenderToScreen
  DoEvents
Loop
'================================主循环=======================================
Do Until 剧情进度 > 5
 VoiceT = VoiceT + Tv.TimeElapsed
 物理引擎(3) = 0: 物理引擎(4) = 0
 If PlayerHP(0) = 0 Then GoTo 移动代码结束
 Inp.GetMouseState Mx, My, B1, B2, , , Roll           '接收鼠标信息
 CameraAngX = CameraAngX + 0.08 * Mx
 CameraAngY = CameraAngY + 0.08 * My
If CameraAngY > 50 Then CameraAngY = 50
If CameraAngY < -60 Then CameraAngY = -60
开关移动 = False
    If Inp.IsKeyPressed(TV_KEY_W) Then '前
      开关移动 = True
      物理引擎(3) = 物理引擎(3) + Cos(Math.Deg2Rad(CameraAngX - 90)) * 0.002 * Tv.TimeElapsed * 游戏速度
      物理引擎(4) = 物理引擎(4) - Sin(Math.Deg2Rad(CameraAngX - 90)) * 0.002 * Tv.TimeElapsed * 游戏速度
    End If
    If Inp.IsKeyPressed(TV_KEY_S) Then '后
     物理引擎(3) = 物理引擎(3) + Cos(Math.Deg2Rad(CameraAngX + 90)) * 0.001 * Tv.TimeElapsed * 游戏速度
     物理引擎(4) = 物理引擎(4) - Sin(Math.Deg2Rad(CameraAngX + 90)) * 0.001 * Tv.TimeElapsed * 游戏速度
     开关移动 = True
    End If
    If Inp.IsKeyPressed(TV_KEY_A) Then '左
     物理引擎(3) = 物理引擎(3) + Cos(Math.Deg2Rad(CameraAngX + 180)) * 0.002 * Tv.TimeElapsed * 游戏速度
     物理引擎(4) = 物理引擎(4) - Sin(Math.Deg2Rad(CameraAngX + 180)) * 0.002 * Tv.TimeElapsed * 游戏速度
     开关移动 = True
    End If
    If Inp.IsKeyPressed(TV_KEY_D) Then '右
     物理引擎(3) = 物理引擎(3) + Cos(Math.Deg2Rad(CameraAngX)) * 0.002 * Tv.TimeElapsed * 游戏速度
     物理引擎(4) = 物理引擎(4) - Sin(Math.Deg2Rad(CameraAngX)) * 0.002 * Tv.TimeElapsed * 游戏速度
     开关移动 = True
    End If
移动代码结束:
    If 开关移动 = True And Inp.IsKeyPressed(TV_KEY_LEFTSHIFT) = False Then
      CameraPoz偏移(0) = 0.1 + CameraPoz偏移(0)
      CameraPoz偏移(2) = Sin(CameraPoz偏移(0)) * 游戏速度 / 8
    End If
'=====物理碰撞=====
If 开关移动 = False Then GoTo 结束玩家移动
Dim 碰撞临时 As Boolean
CameraPozX = CameraPozX + 物理引擎(3)
CameraPozZ = CameraPozZ + 物理引擎(4)
CameraPozY = 物理计算高度(Vector(CameraPozX, CameraPozY, CameraPozZ), 20, 3)
For i = 0 To UBound(MeshTVA)
  If MeshTVA(i).Collision(Vector(CameraPozX - 2, CameraPozY + PlayerHeight, CameraPozZ), Vector(CameraPozX + 2, CameraPozY + PlayerHeight, CameraPozZ)) Or _
  MeshTVA(i).Collision(Vector(CameraPozX, CameraPozY + PlayerHeight, CameraPozZ - 2), Vector(CameraPozX, CameraPozY + PlayerHeight, CameraPozZ + 2)) Then
  视角坐标更新 False: GoTo 结束玩家移动
  End If
Next
If CameraPozZ > 1680 Then 视角坐标更新 False: GoTo 结束玩家移动
视角坐标更新 True
结束玩家移动:
'===============瞄准射击===============
If Inp.IsKeyPressed(TV_KEY_1) And 武器编号 <> 0 Then GunLoad 武器名(0), 0, False
If Inp.IsKeyPressed(TV_KEY_2) And 武器编号 <> 1 Then GunLoad 武器名(1), 1, False
If B1 = True Then
  If 武器状态 = 0 Then 武器状态 = 1: 射击间隔 = 0
End If
If Inp.IsKeyPressed(TV_KEY_R) Then '换弹匣
If 武器状态 <> 2 Then
  If 武器弹匣数(武器编号) > 0 And 武器弹匣(武器编号) < 限定武器弹匣(武器编号) And 开关隐藏武器 = False Then
  GunSE 武器名(武器编号), "reload.wav", 1
  武器状态 = 2
  射击间隔 = 0
  End If
End If
End If
If 开关隐藏武器 = True Then 武器状态 = 0
Select Case 武器状态
Case -1 '===装备===
  If Player(0).IsAnimationFinished Then 武器状态 = 0
Case 0 '===静止===
  If 开关移动 = True Then
   武器位置偏移(0) = 0.01 * Tv.TimeElapsed + 武器位置偏移(0)
   武器位置偏移(2) = Sin(武器位置偏移(0)) * 游戏速度 / 18
  Else
   武器位置偏移(0) = 0.001 * Tv.TimeElapsed + 武器位置偏移(0)
   武器位置偏移(2) = Sin(武器位置偏移(0)) * 游戏速度 / 18
  End If
  If 武器位置偏移(0) >= 6.28 Then 武器位置偏移(0) = 0
Case 1 '===开火===
  If 射击间隔 = 0 Then
  If 武器弹匣(武器编号) <= 0 Then '空膛
    GunSE "通用", "Empty.wav", 1
    武器状态 = 0
  Else
    GunSE 武器名(武器编号), "shoot.wav", 1
    i = 1
    Do Until InStr(Player(0).GetAnimationName(i), "shoot") > 0
     i = 1 + i
    Loop
    Player(0).SetAnimationID i: Player(0).SetAnimationLoop False: Player(0).PlayAnimation
    武器弹匣(武器编号) = 武器弹匣(武器编号) - 1
    后坐力(1) = 限定武器后坐力(1)
    For i = 1 To UBound(Enemy)
If EnmHP(i) < 0 Then GoTo 跳过此目标 '跳过此目标A
    With Enemy(i)
    If .Collision(Vector(CameraPozX, CameraPozY + PlayerHeight + CameraPoz偏移(2), CameraPozZ), Vector(CameraPozX + Cos(Math.Deg2Rad(CameraAngX - 90)) * Cos(Math.Deg2Rad(CameraAngY)) * 武器射程(武器编号), CameraPozY + PlayerHeight - Sin(Math.Deg2Rad(CameraAngY)) * 武器射程(武器编号), CameraPozZ - Sin(Math.Deg2Rad(CameraAngX - 90)) * Cos(Math.Deg2Rad(CameraAngY)) * 武器射程(武器编号)), TV_TESTTYPE_HITBOXES) Then
      EnmHP(i) = EnmHP(i) - 武器伤害(武器编号)
      If EnmHP(i) <= 0 Then
        Select Case EnmType(i)
        Case 0: 执行动作 1, i, "death" & (1 + Round(Rnd * 2)), 游戏速度, False
        Case 1: 执行动作 1, i, "death" & (1 + Round(Rnd * 2)) & "_die", 游戏速度, False
        Case 2: 执行动作 1, i, "die_simple", 游戏速度, False
        End Select
      End If
    End If
    End With
跳过此目标: '跳过此目标B
    Next
  End If
  End If
  射击间隔 = 射击间隔 + Tv.TimeElapsed
Case 2 '===换弹匣===
  If 射击间隔 = 0 Then 执行动作 0, 0, "reload", 游戏速度, False
  If Player(0).IsAnimationFinished Then
    If 武器弹匣(武器编号) > 0 Then
      武器弹匣(武器编号) = 1 + 限定武器弹匣(武器编号)
    Else
      武器弹匣(武器编号) = 限定武器弹匣(武器编号)
    End If
    武器弹匣数(武器编号) = 武器弹匣数(武器编号) - 1
    执行动作 0, 0, "idle", 游戏速度, False
    武器状态 = 0
  End If
  射击间隔 = 射击间隔 + Tv.TimeElapsed
End Select
'===============主角武器===============
If 后坐力(1) > 0 Then 后坐力(1) = 后坐力(1) - 0.02 * Tv.TimeElapsed
If 后坐力(1) < 0 Then 后坐力(1) = 后坐力(1) + 0.02 * Tv.TimeElapsed
If Abs(后坐力(1)) < 0.1 Then 后坐力(1) = 0
Player(0).SetPosition CameraPozX, CameraPozY + PlayerHeight + CameraPoz偏移(2) + 武器位置偏移(2), CameraPozZ
Player(0).SetRotation 0, CameraAngX + 90 + 后坐力(0), CameraAngY - 后坐力(1)
'设定摄像机
Camera.SetPosition CameraPozX, CameraPozY + PlayerHeight + CameraPoz偏移(2), CameraPozZ
Camera.SetRotation CameraAngY - 后坐力(1), CameraAngX + 后坐力(0), 0
'=====敌人武器=====
For i = 1 To UBound(Enemy): With EnemyGun(i)
  .SetMatrix Enemy(i).GetBoneMatrix(Enemy(i).GetBoneID("Bip01 R Hand")) '绑定到手
  .RotateY -70
  .RotateZ -95
  .MoveRelative 0.2, -1.2, 0.17
End With: Next
'===============清屏与渲染===============
Atmos.Fog_Enable False
ReflectRS.StartRender '反射部分渲染
 Atmos.Atmosphere_Render
ReflectRS.EndRender

RefractRS.StartRender '折射部分渲染
 Atmos.Atmosphere_Render
RefractRS.EndRender

Tv.Clear '清屏
Atmos.Atmosphere_Render '渲染大气
Atmos.Fog_Enable True
For i = 1 To UBound(Enemy): EnemyGun(i).Render: Next
For i = 1 To UBound(Enemy): Enemy(i).Render: Next
For i = 0 To UBound(MeshTVA): MeshTVA(i).Render: Next
Scene.RenderAllMeshes
Mesh(0).Render
Floor.Render
Scene.FinalizeShadows '渲染影子
Scene.RenderAllParticleSystems
'===============角色事件===============
For i = 1 To UBound(Enemy)
'LE.DeleteLight LE.GetLightFromName("EnmGunFire" & i)
With Enemy(i)
  If EnmHP(i) <= 0 Then GoTo 跳过此对象
  Select Case EnmType(i)
  Case 0: AIadv 1, i, 2, 0.008
  Case 1: AIadv 1, i, 2, 0.006
  Case 2: AIadv 1, i, 1, 0.01
  End Select
End With
跳过此对象:
Next
Select Case 剧情进度
Case 2: If 距离平方(Player(0).GetPosition, MeshTVA(1).GetPosition) < 2000 Then 剧情进度 = 3 '扶着悍马站起来
Case 3: If PlayerHeight < 8 Then PlayerHeight = 0.2 + PlayerHeight Else 游戏速度 = 2
Case 4
  If PlayerHeight > 2 Then
    PlayerHeight = PlayerHeight - 0.2
  Else
    剧情进度 = 5
    BASSset 0, 0: BASSset 1, 0: BASSset 2, 0
  End If
End Select
'===============最终渲染===============
If 武器状态 = 1 And 射击间隔 < 60 Then Scr图形.Draw_Sprite GetTex("flash" & 枪火类型 & Round(Rnd)), 枪火X * 适配X, 枪火Y * 适配Y
If 开关隐藏武器 = False Then Player(0).Render
If 血污残留时间 > 0 Or PlayerHP(0) < 20 Then Scr图形.Draw_SpriteScaled GetTex("血污0"), 0, 0, -1, 适配X, 适配Y: 血污残留时间 = 血污残留时间 - 1
If 开关隐藏武器 = False Then
Select Case 武器状态 '绘制准星
Case 0
  If B2 = True Then
  Scr图形.Draw_Line 准星X - 60, 准星Y, 准星X + 60, 准星Y, RGBA(0, 0.8, 0, 2)
  Scr图形.Draw_Line 准星X, 准星Y, 准星X, 准星Y + 30, RGBA(0, 0.8, 0, 2)
  Else
  Scr图形.Draw_Line 准星X - 40, 准星Y, 准星X - 40, 准星Y + 30, RGBA(0, 0.8, 0, 2)
  Scr图形.Draw_Line 准星X + 40, 准星Y, 准星X + 40, 准星Y + 30, RGBA(0, 0.8, 0, 2)
  Scr图形.Draw_Line 准星X - 100, 准星Y, 准星X - 40, 准星Y, RGBA(0, 0.8, 0, 2)
  Scr图形.Draw_Line 准星X + 100, 准星Y, 准星X + 40, 准星Y, RGBA(0, 0.8, 0, 2)
  End If
Case 1
  If 射击间隔 > 限定射击间隔(武器编号) Then
    射击间隔 = 0
    If B1 = False Then 武器状态 = 0: 执行动作 0, 0, "idle1", 游戏速度, False
  End If
  Scr图形.Draw_Line 准星X - 120, 准星Y, 准星X - 70, 准星Y, RGBA(0, 0.8, 0, 2)
  Scr图形.Draw_Line 准星X + 120, 准星Y, 准星X + 70, 准星Y, RGBA(0, 0.8, 0, 2)
End Select
End If
'===============文字渲染===============
Scr图形.Draw_Sprite GetTex("UI武器属性"), 15, Me.Height / 15 - 135 '武器属性UI
Select Case PlayerHP(0)
Case Is > 60: UI字体颜色 = RGBA(1, 1, 1, 1)
Case Is > 30: UI字体颜色 = RGBA(1, 1, 0.5, 0.6)
Case Else: UI字体颜色 = RGBA(1, 0.5, 0.5, 0.6)
End Select
If 武器弹匣(武器编号) > 限定武器弹匣(武器编号) Then
  Lrc.NormalFont_DrawText "1+" & (武器弹匣(武器编号) - 1) & "/" & 限定武器弹匣(武器编号), 50, Me.Height \ 15 - 100, UI字体颜色, 1
Else
  Lrc.NormalFont_DrawText 武器弹匣(武器编号) & "/" & 限定武器弹匣(武器编号), 50, Me.Height \ 15 - 100, UI字体颜色, 1
End If
Lrc.NormalFont_DrawText 武器名(武器编号) & " " & 武器弹匣数(武器编号) & "/" & 限定武器弹匣数(武器编号), 50, Me.Height \ 15 - 60, UI字体颜色, 1
If LRC消失时间 > 0 Then
  DrawLRC LRCname, LRCtext, LRCcolor
  LRC消失时间 = LRC消失时间 - Tv.TimeElapsed
End If

DrawCLRC RGBA(0.2, 1, 0.2, 1)
Tv.RenderToScreen
DoEvents
Loop
Do Until 剧情进度 > 6
  Tv.Clear
  DrawCLRC (RGBA(1, 1, 1, 1)): DrawLRC LRCname, LRCtext, LRCcolor
  Tv.RenderToScreen
  DoEvents
Loop
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
End
End Sub
Private Sub Tim敌人重生_Timer()
For i = 1 To UBound(Enemy)
With Enemy(i)
  If EnmHP(i) > 0 And Enemy(i).GetPosition.Y > -500 Then GoTo 跳过此对象
  .SetPosition Rnd * 40 + 80, 60, Rnd * 40 + 80
  Select Case EnmType(i)
  Case 0
    执行动作 1, i, "run", 游戏速度, True
    EnmHP(i) = 难度 * 10
  Case 1
    执行动作 1, i, "run", 游戏速度, True
    EnmHP(i) = 难度 * 20
  Case 2
    执行动作 1, i, "run2", 游戏速度, True
    EnmHP(i) = 难度 * 15
  End Select
  EnmState(i) = 0
  Exit Sub
End With
跳过此对象:
Next
End Sub
Private Sub Timer1_Timer()
T总 = 1 + T总
Select Case T总
Case 1: If 调试模式 Then 剧情进度 = 1
Case 6: SoundMp3.Load "Audio\SE\Type0.wav": CLRC "", "", "", "2XXX年9月18日 法国加莱"
Case 10: CLRCt(0) = 0
Case 11: 剧情进度 = 1: If 调试模式 = False Then GF.FadeIn 2000
Case 14: CreatLRC "W、A、S、D键爬行 扶着悍马站起", "", RGBA(1, 1, 0, 1)
Case 17: CLRC "—策划—", "佐亚_洛克k上尉", "木龙华易", "吃桃的叫天子"
Case 22: CLRC "—主程序—", "佐亚_洛克k上尉", "           Drzzm32(移植)", ""
Case 27: CLRC "—美术—", "木龙华易", "佐亚_洛克k上尉", ""
Case 32: CLRC "—军事顾问—", "            wen832238", "555十个九坑啊", "零点上将"
Case 37: CLRC "—测评监督—", "木龙华易", "灰鳍鲨", "          Archeb"
Case 42: CLRC "—制作团队—", "九州烽火工作室", "喵玉殿官方技术团", "        DeseCity工作室（已坑）"
Case 47: CLRC "—Powered By—", "           TrueVision 3D Engine", "        Visual Basic 6", "           BASS.dll Sound Engine"
Case 52: 剧情进度 = 2
Case 120: GF.Flash 1, 1, 1, 8000
Case 150: GF.Flash 1, 1, 1, 8000
Case 180: GF.Flash 1, 1, 1, 8000: 剧情进度 = 4: GF.FadeOut 3000
Case 183
  CLRC "", "", "东方现代战争", "月落幻想": CreatLRC "", "Touhou Modern War:Dreaming Fallen Of Moon", RGBA(1, 1, 1, 1)
  剧情进度 = 6
Case 191: Unload Me
End Select
End Sub

