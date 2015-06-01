Attribute VB_Name = "AAAA��������AAAA"
Public Tv As New TVEngine '����tv3d�������
Public Scene As New TVScene '����tv3d�������
Public Inp As New TVInputEngine '�������
Public TF As New TVTextureFactory '���һ����ͼ��
Public MF As New TVMaterialFactory ''���һ�����ʿ�
Public SE As New TVSoundEngine '��Ƶ����
Public Sounds(0 To 25) As TVSounds, SEchan As Long
Public SoundMp3 As New TVSoundMP3
Public Sound3DNum As Integer '3D����
Public Sound3D() As New TVSoundWave3D
Public Listener As TVListener
Public LE As New TVLightEngine '���һ���ƹ��
Public GF As New TVGraphicEffect
Public DofRS As TVRenderSurface
Public ˮ�� As TV_PLANE, ˮ��(1 To 2) As TVRenderSurface
Public Camera As New TVCamera '����һ����������൱���˵��۾�
Public Atmos  As New TVAtmosphere '��Ӵ���ϵͳ
Public Math As New TVMathLibrary '���tv3d��ѧ�����
Public Scrͼ�� As New TVScreen2DImmediate '2d�����
Public Lrc As New TVScreen2DText
Public Player(0 To 0) As TVActor  '�������\����
'����ˮ�淴�䡪��
Public ReflectRS As TVRenderSurface '��ӷ����
Public RefractRS As TVRenderSurface '��������
Public WaterPlane As TV_PLANE '���ˮƽ��
'������ͨ��������
Public ���Ķ�˵�� As Boolean
Public ������(0 To 9) As String, ������� As Long, ����״̬ As Long '����״̬ -1װ�� 0��ֹ 1��ǹ 2װ��
Public ������ As Long, ������ϻ(0 To 9) As Long, ������ϻ��(0 To 9) As Long
Public �޶�������(0 To 9) As Long, �޶�������ϻ(0 To 9) As Long, �޶�������ϻ��(0 To 9) As Long, �޶�����������(0 To 1) As Single
Public �������(0 To 9) As Long, �����˺�(0 To 9) As Long
Public ׼��X As Long, ׼��Y As Long
Public ǹ������ As Long, ǹ��X As Long, ǹ��Y As Long, ������(0 To 1) As Single, UI������ɫ As Single
Public ����λ��ƫ��(0 To 3) As Single
Public ������������ As Boolean

Public ����X As Single, ����Y As Single '��Ļ����
Public i As Long, j As Long 'ѭ��ר��
Public Mx As Long, My As Long, B1 As Boolean, B2 As Boolean, Roll As Long   '���������Ϣ
Public CameraPozX As Single, CameraPozY As Single, CameraPozZ As Single '�����λ������
Public CameraAngX As Single, CameraAngY As Single '������Ƕ�
Public CameraPozXOld As Single, CameraPozYOld As Single, CameraPozZOld As Single '�������һ֡λ������
Public PlayerHeight As Single '������
Public CameraPozƫ��(0 To 3) As Single '�����ƫ��λ��

Public ��Ϸ�ٶ� As Single
Public PlayerHP(0 To 99) As Single, Ѫ�۲���ʱ�� As Long, ��������(0 To 6) As Single
Public EnmState(1 To 99) As Long, EnmHP(1 To 99) As Long, EnmT(1 To 99) As Long
Public ��ʱ(0 To 9) As Long
Public �����ƶ� As Boolean
'�������ַ�������
Public Floor As TVMesh
Public Function ����ˮ��(X1 As Single, Z1 As Single, X2 As Single, Z2 As Single, Height As Single)
TF.LoadDUDVTexture "Pic\Stage\Water.jpg", "water", -1, -1, 25 '��ȡˮ�淨����ͼ
Set ReflectRS = Scene.CreateRenderSurfaceEx(-1, -1, TV_TEXTUREFORMAT_DEFAULT, True, True, 1)  '��������ͼ��
Set RefractRS = Scene.CreateRenderSurfaceEx(-1, -1, TV_TEXTUREFORMAT_DEFAULT, True, True, 1) '��������ͼ��
Set Floor = Scene.CreateMeshBuilder
Floor.AddFloor GetTex("water"), X1, Z1, X2, Z2, Height, 20, 20  '����ˮƽ��
Floor.SetLightingMode TV_LIGHTING_BUMPMAPPING_TANGENTSPACE '�ƹ�ģʽ��Ϊ������ͼģʽ
WaterPlane.Normal = Vector(0, 1, 0) 'ˮ��ķ���
WaterPlane.Dist = -0.2 'ˮ�淴��߶ȣ�Ϊˮ��߶ȵĸ���
GF.SetWaterReflection Floor, ReflectRS, RefractRS, 0, WaterPlane '��ʼ������������ֵ��Ϊ2ʱ��ˮ�治���в��ƣ���������ûʲô��
GF.SetWaterReflectionBumpAnimation Floor, True, 0.2, 0.2 'ˮ��Ĳ����ٶ�
End Function
Public Function SEplay(�ļ� As String, �Ƿ���ʱ��ȡ As Boolean)
SEchan = 1 + SEchan: If SEchan > 25 Then SEchan = 1
Set Sounds(SEchan) = Nothing 'ĳЩ��Ƶ�ƺ������
Set Sounds(SEchan) = SE.CreateSounds
Set Listener = Nothing
Set Listener = SE.Get3DListener
If �Ƿ���ʱ��ȡ = True Then SoundMp3.Load "Audio\SE\" & �ļ�
Sounds(SEchan).AddFile "Audio\SE\" & �ļ�
Sounds(SEchan).Item(Left(�ļ�, InStrRev(�ļ�, ".") - 1)).Play
End Function

Public Function GunLoad(�������� As String, �滻������� As Long, ���� As Boolean)
Dim wuqicanshuduqu As Long, NR(1 To 11) As String
With Player(0)
.LoadTVA ("Weapon\" & �������� & "\v_" & �������� & ".tva")
.SetMaterial GetMat("solid") '�趨����
.SetLightingMode TV_LIGHTING_NORMAL '�趨����ģʽ
.SetAnimationByName ("draw") 'ִ�еĶ�������
.SetAnimationLoop False '������ѭ��
.PlayAnimation 1.3 * ��Ϸ�ٶ� '���Ŷ����ٶ�
.SetScale 0.3, 0.3, 0.3 '�趨��С
End With

Open App.Path & "\Weapon\" & �������� & "\����.ini" For Input As #91
For wuqicanshuduqu = 1 To 11
  Line Input #91, NR(wuqicanshuduqu)
Next
Close #91
������(�滻�������) = ��������
����״̬ = -1: ǹ��X = ׼��X + Val(NR(7)) - 240: ǹ��Y = ׼��Y - Val(NR(8)) + 240
������ = 0: �޶�������(�滻�������) = Val(NR(3))
�����˺�(�滻�������) = Val(NR(2)): �������(�滻�������) = Val(NR(11))
�޶�����������(0) = NR(4): �޶�����������(1) = NR(5)
If ���� = True Then
  ������ϻ��(�滻�������) = Val(NR(10)): �޶�������ϻ��(�滻�������) = Val(NR(10))
  ������ϻ(�滻�������) = Val(NR(9)): �޶�������ϻ(�滻�������) = Val(NR(9))
End If
GunSE ��������, "shoot.wav", 0
GunSE ��������, "reload.wav", 0
������� = �滻�������
End Function
Public Function ��ʼ���ӽǲ���(X As Single, Y As Single, z As Single, ����Y As Single, ��ʼHP As Long)
CameraPozX = X
CameraPozY = Y
CameraPozZ = z
CameraPozXOld = X
CameraPozYOld = Y
CameraPozZOld = z
CameraAngY = 0
CameraAngX = ����Y
PlayerHeight = 17
PlayerHP(0) = ��ʼHP
��Ϸ�ٶ� = 1
SoundMp3.Load "Weapon\ͨ��\Empty.wav"
SoundMp3.Load "Audio\SE\�����ŵ�0.wav"
SoundMp3.Load "Audio\SE\AI_reload.wav"
For i = 0 To 7: SoundMp3.Load "Audio\SE\����" & i & ".wav": Next
End Function
Public Function �ӽ��������(�Ƿ���� As Boolean)
If �Ƿ���� = False Then
  CameraPozX = CameraPozXOld
  CameraPozY = CameraPozYOld
  CameraPozZ = CameraPozZOld
Else
  CameraPozXOld = CameraPozX
  CameraPozYOld = CameraPozY
  CameraPozZOld = CameraPozZ
End If
End Function
Public Function GunSE(������ As String, �ļ� As String, ������ As Long)
Set Sounds(0) = Nothing
Set Sounds(0) = SE.CreateSounds
Set Listener = Nothing
Set Listener = SE.Get3DListener
Select Case ������
Case 0 '��ȡ
  SoundMp3.Load "Weapon\" & ������ & "\" & �ļ�
Case 1 '����
  Sounds(0).AddFile "Weapon\" & ������ & "\" & �ļ�
  Sounds(0).Item(Left(�ļ�, InStrRev(�ļ�, ".") - 1)).Play
End Select
End Function
Public Function ����ƽ��(��ɫ��1 As TV_3DVECTOR, ��ɫ��2 As TV_3DVECTOR) As Long
����ƽ�� = (��ɫ��1.X - ��ɫ��2.X) ^ 2 + (��ɫ��1.z - ��ɫ��2.z) ^ 2
End Function
