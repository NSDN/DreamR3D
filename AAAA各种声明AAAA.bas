Attribute VB_Name = "AAAA各种声明AAAA"
Public Tv As New TVEngine '调用tv3d所必需的
Public Scene As New TVScene '调用tv3d所必需的
Public Inp As New TVInputEngine '按键检测
Public TF As New TVTextureFactory '添加一个贴图库
Public MF As New TVMaterialFactory ''添加一个材质库
Public SE As New TVSoundEngine '音频引擎
Public Sounds(0 To 25) As TVSounds, SEchan As Long
Public SoundMp3 As New TVSoundMP3
Public Sound3DNum As Integer '3D声音
Public Sound3D() As New TVSoundWave3D
Public Listener As TVListener
Public LE As New TVLightEngine '添加一个灯光库
Public GF As New TVGraphicEffect
Public DofRS As TVRenderSurface
Public 水面 As TV_PLANE, 水反(1 To 2) As TVRenderSurface
Public Camera As New TVCamera '定义一个摄像机，相当于人的眼睛
Public Atmos  As New TVAtmosphere '添加大气系统
Public Math As New TVMathLibrary '添加tv3d数学运算库
Public Scr图形 As New TVScreen2DImmediate '2d处理库
Public Lrc As New TVScreen2DText
Public Player(0 To 0) As TVActor  '添加主角\队友
'——水面反射——
Public ReflectRS As TVRenderSurface '添加反射层
Public RefractRS As TVRenderSurface '添加折射层
Public WaterPlane As TV_PLANE '添加水平面
'——普通变量——
Public 已阅读说明 As Boolean
Public 武器名(0 To 9) As String, 武器编号 As Long, 武器状态 As Long '武器状态 -1装备 0静止 1开枪 2装填
Public 射击间隔 As Long, 武器弹匣(0 To 9) As Long, 武器弹匣数(0 To 9) As Long
Public 限定射击间隔(0 To 9) As Long, 限定武器弹匣(0 To 9) As Long, 限定武器弹匣数(0 To 9) As Long, 限定武器后坐力(0 To 1) As Single
Public 武器射程(0 To 9) As Long, 武器伤害(0 To 9) As Long
Public 准星X As Long, 准星Y As Long
Public 枪火类型 As Long, 枪火X As Long, 枪火Y As Long, 后坐力(0 To 1) As Single, UI字体颜色 As Single
Public 武器位置偏移(0 To 3) As Single
Public 开关隐藏武器 As Boolean

Public 适配X As Single, 适配Y As Single '屏幕适配
Public i As Long, j As Long '循环专用
Public Mx As Long, My As Long, B1 As Boolean, B2 As Boolean, Roll As Long   '接收鼠标信息
Public CameraPozX As Single, CameraPozY As Single, CameraPozZ As Single '摄像机位置坐标
Public CameraAngX As Single, CameraAngY As Single '摄像机角度
Public CameraPozXOld As Single, CameraPozYOld As Single, CameraPozZOld As Single '摄像机上一帧位置坐标
Public PlayerHeight As Single '玩家身高
Public CameraPoz偏移(0 To 3) As Single '摄像机偏移位置

Public 游戏速度 As Single
Public PlayerHP(0 To 99) As Single, 血污残留时间 As Long, 物理引擎(0 To 6) As Single
Public EnmState(1 To 99) As Long, EnmHP(1 To 99) As Long, EnmT(1 To 99) As Long
Public 临时(0 To 9) As Long
Public 开关移动 As Boolean
'——各种方法——
Public Floor As TVMesh
Public Function 建立水面(X1 As Single, Z1 As Single, X2 As Single, Z2 As Single, Height As Single)
TF.LoadDUDVTexture "Pic\Stage\Water.jpg", "water", -1, -1, 25 '读取水面法线贴图
Set ReflectRS = Scene.CreateRenderSurfaceEx(-1, -1, TV_TEXTUREFORMAT_DEFAULT, True, True, 1)  '建立反射图层
Set RefractRS = Scene.CreateRenderSurfaceEx(-1, -1, TV_TEXTUREFORMAT_DEFAULT, True, True, 1) '建立折射图层
Set Floor = Scene.CreateMeshBuilder
Floor.AddFloor GetTex("water"), X1, Z1, X2, Z2, Height, 20, 20  '建立水平面
Floor.SetLightingMode TV_LIGHTING_BUMPMAPPING_TANGENTSPACE '灯光模式设为法线贴图模式
WaterPlane.Normal = Vector(0, 1, 0) '水面的法线
WaterPlane.Dist = -0.2 '水面反射高度，为水面高度的负数
GF.SetWaterReflection Floor, ReflectRS, RefractRS, 0, WaterPlane '初始化，当第三个值设为2时，水面不会有波纹，其他参数没什么用
GF.SetWaterReflectionBumpAnimation Floor, True, 0.2, 0.2 '水面的波纹速度
End Function
Public Function SEplay(文件 As String, 是否临时读取 As Boolean)
SEchan = 1 + SEchan: If SEchan > 25 Then SEchan = 1
Set Sounds(SEchan) = Nothing '某些音频似乎会出错
Set Sounds(SEchan) = SE.CreateSounds
Set Listener = Nothing
Set Listener = SE.Get3DListener
If 是否临时读取 = True Then SoundMp3.Load "Audio\SE\" & 文件
Sounds(SEchan).AddFile "Audio\SE\" & 文件
Sounds(SEchan).Item(Left(文件, InStrRev(文件, ".") - 1)).Play
End Function

Public Function GunLoad(武器名称 As String, 替换武器编号 As Long, 重置 As Boolean)
Dim wuqicanshuduqu As Long, NR(1 To 11) As String
With Player(0)
.LoadTVA ("Weapon\" & 武器名称 & "\v_" & 武器名称 & ".tva")
.SetMaterial GetMat("solid") '设定材质
.SetLightingMode TV_LIGHTING_NORMAL '设定光照模式
.SetAnimationByName ("draw") '执行的动作名称
.SetAnimationLoop False '动作不循环
.PlayAnimation 1.3 * 游戏速度 '播放动作速度
.SetScale 0.3, 0.3, 0.3 '设定大小
End With

Open App.Path & "\Weapon\" & 武器名称 & "\属性.ini" For Input As #91
For wuqicanshuduqu = 1 To 11
  Line Input #91, NR(wuqicanshuduqu)
Next
Close #91
武器名(替换武器编号) = 武器名称
武器状态 = -1: 枪火X = 准星X + Val(NR(7)) - 240: 枪火Y = 准星Y - Val(NR(8)) + 240
射击间隔 = 0: 限定射击间隔(替换武器编号) = Val(NR(3))
武器伤害(替换武器编号) = Val(NR(2)): 武器射程(替换武器编号) = Val(NR(11))
限定武器后坐力(0) = NR(4): 限定武器后坐力(1) = NR(5)
If 重置 = True Then
  武器弹匣数(替换武器编号) = Val(NR(10)): 限定武器弹匣数(替换武器编号) = Val(NR(10))
  武器弹匣(替换武器编号) = Val(NR(9)): 限定武器弹匣(替换武器编号) = Val(NR(9))
End If
GunSE 武器名称, "shoot.wav", 0
GunSE 武器名称, "reload.wav", 0
武器编号 = 替换武器编号
End Function
Public Function 初始化视角参数(X As Single, Y As Single, z As Single, 方向Y As Single, 初始HP As Long)
CameraPozX = X
CameraPozY = Y
CameraPozZ = z
CameraPozXOld = X
CameraPozYOld = Y
CameraPozZOld = z
CameraAngY = 0
CameraAngX = 方向Y
PlayerHeight = 17
PlayerHP(0) = 初始HP
游戏速度 = 1
SoundMp3.Load "Weapon\通用\Empty.wav"
SoundMp3.Load "Audio\SE\掩体着弹0.wav"
SoundMp3.Load "Audio\SE\AI_reload.wav"
For i = 0 To 7: SoundMp3.Load "Audio\SE\擦弹" & i & ".wav": Next
End Function
Public Function 视角坐标更新(是否更新 As Boolean)
If 是否更新 = False Then
  CameraPozX = CameraPozXOld
  CameraPozY = CameraPozYOld
  CameraPozZ = CameraPozZOld
Else
  CameraPozXOld = CameraPozX
  CameraPozYOld = CameraPozY
  CameraPozZOld = CameraPozZ
End If
End Function
Public Function GunSE(武器名 As String, 文件 As String, 命令编号 As Long)
Set Sounds(0) = Nothing
Set Sounds(0) = SE.CreateSounds
Set Listener = Nothing
Set Listener = SE.Get3DListener
Select Case 命令编号
Case 0 '读取
  SoundMp3.Load "Weapon\" & 武器名 & "\" & 文件
Case 1 '播放
  Sounds(0).AddFile "Weapon\" & 武器名 & "\" & 文件
  Sounds(0).Item(Left(文件, InStrRev(文件, ".") - 1)).Play
End Select
End Function
Public Function 距离平方(角色类1 As TV_3DVECTOR, 角色类2 As TV_3DVECTOR) As Long
距离平方 = (角色类1.X - 角色类2.X) ^ 2 + (角色类1.z - 角色类2.z) ^ 2
End Function
