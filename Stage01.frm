VERSION 5.00
Begin VB.Form Stage01 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "�����ִ�ս����糺�����5  �������ݷ������"
   ClientHeight    =   8970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Tim���� 
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
Dim Mesh(0 To 2) As TVMesh: Dim MeshTVA(0 To 0) As TVActor: Dim MeshSin(0 To 0) As TVMesh
Dim Enemy(1 To 9) As TVActor: Dim EnemyGun(1 To 9) As TVActor
Dim EnmLastView(1 To 15) As TV_3DVECTOR
Dim ��ɫλ��(0 To 3) As TV_3DVECTOR '0���12��ɫ3��ת
Dim ��ɫ����(0 To 3) As TV_3DVECTOR
Dim VoiceT As Long
'������ͨ������������
Dim �Ѷ� As Long, ������� As Long, ʱ���� As Long
Dim ����ģʽ As Boolean
Dim VoiName As String, Word As String
Private Function CreatWord(���� As String, ���� As String, ������ɫ As Single)

End Function
Private Function DrawText(���� As String, ���� As String, ������ɫ As Single)
Dim LSX As Long
LSX = Me.Width \ 30 - Len(���� & ����) * 10 - 8
Lrc.NormalFont_DrawText ����, LSX, Me.Height \ 15 - 100, ������ɫ, 1
Lrc.NormalFont_DrawText ����, LSX + Len(����) * 20 + 8, Me.Height \ 15 - 100, RGBA(1, 1, 1, 1), 1
End Function
Public Function �������߶�(λ�� As TV_3DVECTOR, ԽҰ�߶� As Single, �����ٶ� As Single) As Single
��������(1) = λ��.Y + 200
��������(2) = λ��.Y - �����ٶ�
For �ڲ�ѭ����ʱ = 1 To 15
  ��������(0) = (��������(1) + ��������(2)) / 2
  If Mesh(0).Collision(Vector(λ��.X, ��������(1), λ��.z), Vector(λ��.X, ��������(0), λ��.z), TV_TESTTYPE_HITBOXES) Then
    ��������(2) = ��������(0)
  Else
    ��������(1) = ��������(0)
  End If
Next
If ��������(0) > λ��.Y + ԽҰ�߶� Then ��������(0) = λ��.Y
�������߶� = ��������(0)
End Function
Private Function ִ�ж���(�������� As Long, ������ As Long, ������ As String, �����ٶ� As Single, �Ƿ�ѭ�� As Boolean)
Select Case ��������
Case 0 '��ҺͶ���
  Player(������).SetAnimationByName ������
  Player(������).SetAnimationLoop �Ƿ�ѭ��
  Player(������).PlayAnimation �����ٶ�
Case 1 '����
  Enemy(������).SetAnimationByName ������
  Enemy(������).SetAnimationLoop �Ƿ�ѭ��
  Enemy(������).PlayAnimation �����ٶ�
End Select
End Function
Private Function AImoveTo(������ As Long, �ƶ����� As Long, Ŀ��λ�� As TV_3DVECTOR, �ƶ��ٶ� As Single)
Dim ��ʱ����(0 To 6) As TV_3DVECTOR: Dim ����(0 To 6) As Single
Select Case �ƶ�����
Case 0 '===���У��������===
    ��ʱ����(0) = Ŀ��λ��
    ��ʱ����(3) = Enemy(������).GetRotation
    Enemy(������).LookAtPoint ��ʱ����(0)
    ��ʱ����(1) = Enemy(������).GetRotation
    Enemy(������).SetRotation ��ʱ����(3).X, ��ʱ����(1).Y + 90, ��ɫ����(3).z
    ��ʱ����(3) = Enemy(������).GetRotation
   If �ƶ��ٶ� > 360 Or Abs(��ɫ����(1).Y - ��ɫ����(3).Y) <= �ƶ��ٶ� Then Exit Function
Case 1 '===���У�ֱ���ƶ�===
    ��ʱ����(0) = Enemy(������).GetPosition: ��ʱ����(0).Y = ��ʱ����(0).Y - 12
    ����(1) = �ƶ��ٶ� * (Ŀ��λ��.X - ��ʱ����(0).X)
    ����(2) = �ƶ��ٶ� * (Ŀ��λ��.z - ��ʱ����(0).z)
    Enemy(������).SetPosition ��ʱ����(0).X + ����(1), �������߶�(��ʱ����(0), 1, 3) + 12, ��ʱ����(0).z + ����(2)
    ��ʱ����(1) = Enemy(������).GetPosition
    If Mesh(0).Collision(Vector(��ʱ����(1).X + 4, ��ʱ����(1).Y, ��ʱ����(1).z), Vector(��ʱ����(1).X - 4, ��ʱ����(1).Y, ��ʱ����(1).z)) Then Enemy(������).SetPosition ��ʱ����(1).X - ����(1), ��ʱ����(1).Y, ��ʱ����(1).z
    If Mesh(0).Collision(Vector(��ʱ����(1).X, ��ʱ����(1).Y, ��ʱ����(1).z + 4), Vector(��ʱ����(1).X, ��ʱ����(1).Y, ��ʱ����(1).z - 4)) Then Enemy(������).SetPosition ��ʱ����(1).X, ��ʱ����(1).Y, ��ʱ����(1).z - ����(2)
End Select
End Function
Private Function AImove(�������� As Long, ������ As Long, �¼����� As Long, �ƶ��ٶ� As Single)
Select Case �¼�����
Case 1 '===��������===
  If �������� = 0 Then

  Else '����
    ��ɫλ��(0) = Player(0).GetPosition
    ��ɫ����(3) = Enemy(������).GetRotation
    Enemy(������).LookAtPoint ��ɫλ��(0)
    ��ɫ����(1) = Enemy(������).GetRotation
    Enemy(������).SetRotation ��ɫ����(3).X, ��ɫ����(1).Y + 90, ��ɫ����(3).z
    ��ɫ����(1) = Enemy(������).GetRotation
   If �ƶ��ٶ� > 360 Or Abs(��ɫ����(1).Y - ��ɫ����(3).Y) <= �ƶ��ٶ� Then Exit Function
 'ֱ���������
    If Abs(��ɫ����(1).Y - ��ɫ����(3).Y) > 180 Then GoTo ���˷�ת�� '������ת�����
    If ��ɫ����(3).Y < ��ɫ����(1).Y Then
      Enemy(������).RotateY �ƶ��ٶ�
    Else
      Enemy(������).RotateY -�ƶ��ٶ�
    End If
    Exit Function
���˷�ת��:
    If ��ɫ����(3).Y < ��ɫ����(1).Y Then
      Enemy(������).RotateY -�ƶ��ٶ�
    Else
      Enemy(������).RotateY �ƶ��ٶ�
    End If
  End If
Case 2 '===��е�ӽ����===
  If �������� = 1 Then
    ��ɫλ��(1) = Enemy(������).GetPosition: ��ɫλ��(1).Y = ��ɫλ��(1).Y - 12
    ��������(5) = �ƶ��ٶ� * (CameraPozX - ��ɫλ��(1).X)
    ��������(6) = �ƶ��ٶ� * (CameraPozZ - ��ɫλ��(1).z)
    Enemy(������).SetPosition ��ɫλ��(1).X + ��������(5), �������߶�(��ɫλ��(1), 1, 3) + 12, ��ɫλ��(1).z + ��������(6)
    ��ɫλ��(3) = Enemy(������).GetPosition
    If Mesh(0).Collision(Vector(��ɫλ��(3).X + 4, ��ɫλ��(3).Y, ��ɫλ��(3).z), Vector(��ɫλ��(3).X - 4, ��ɫλ��(3).Y, ��ɫλ��(3).z)) Then Enemy(������).SetPosition ��ɫλ��(3).X - ��������(5), ��ɫλ��(3).Y, ��ɫλ��(3).z
    If Mesh(0).Collision(Vector(��ɫλ��(3).X, ��ɫλ��(3).Y, ��ɫλ��(3).z + 4), Vector(��ɫλ��(3).X, ��ɫλ��(3).Y, ��ɫλ��(3).z - 4)) Then Enemy(������).SetPosition ��ɫλ��(3).X, ��ɫλ��(3).Y, ��ɫλ��(3).z - ��������(6)
  End If
End Select
End Function
Private Function AIadv(�������� As Long, ������ As Long, AI���� As Long, �ƶ��ٶ� As Single)
Dim ��ײ(0 To 2) As Boolean
Select Case ��������
Case 0 '����
Case 1 '����
  ��ɫλ��(0) = Player(0).GetPosition
  ��ɫλ��(1) = Enemy(������).GetPosition
  Select Case AI����
  Case 1 '==��ʬ==
      If Abs(��ɫλ��(1).X - ��ɫλ��(0).X) < 12 Then
      If Abs(��ɫλ��(1).z - ��ɫλ��(0).z) < 12 Then
       If Ѫ�۲���ʱ�� <= 20 Then ������� 0, 20
       Exit Function
      End If
      End If
      AImove 1, ������, 1, 999
      AImove 1, ������, 2, �ƶ��ٶ�
  Case 2 '==����==
      AImove 1, ������, 2, 0
      ��ײ(0) = Mesh(0).Collision(Enemy(������).GetPosition, Player(0).GetPosition)
      If ��ײ(0) = False Then EnmLastView(������) = Vector(CameraPozX + 20 * ��������(3), CameraPozY, CameraPozZ + 20 * ��������(4))
      Select Case EnmState(������)
      Case 0 '�����ӽ���ҡ���
          If ��ײ(0) = False Then '��ҿ���
            AImoveTo ������, 0, Player(0).GetPosition, 999
            If ����ƽ��(Enemy(������).GetPosition, Player(0).GetPosition) > 6000 Then
              AImoveTo ������, 1, Player(0).GetPosition, �ƶ��ٶ�
            Else
              If VoiceT > 4000 Then SEplay "AI-fire" & CInt(2 * Rnd) & ".wav", True: VoiceT = 0 'voice
              ִ�ж��� 1, ������, "ref_shoot_onehanded", ��Ϸ�ٶ�, True
              EnmState(������) = 1
            End If
          Else '��Ҳ�����
            If EnmLastView(i).X <> 0 Then AImoveTo i, 1, EnmLastView(������), 2 * �ƶ��ٶ�
          End If
      Case 1 '���������������
          If ��ײ(0) = False Then
              AImove 1, ������, 1, 999
              If ����ƽ��(Enemy(������).GetPosition, Player(0).GetPosition) < 9000 Then
                If EnmT(������) Mod 300 < 20 Then '����
                    SEplay "�����ŵ�0.wav", False
                    SEplay "����" & CInt(Rnd * 6) & ".wav", False
                    If Ѫ�۲���ʱ�� <= 40 Then ������� 0, 2
                End If '��������
                EnmT(������) = EnmT(������) + Tv.TimeElapsed
                If EnmT(������) > 4000 Then '����
                    If VoiceT > 3000 Then SEplay "AI-reload.wav", True: VoiceT = 0 'voice
                    SEplay "AI_reload.wav", False
                    ִ�ж��� 1, ������, "ref_reload_onehanded", ��Ϸ�ٶ�, False
                    EnmState(i) = 2
                End If
              Else '����Զ���л�׷��
                  ִ�ж��� 1, ������, "run", ��Ϸ�ٶ�, True
                  EnmState(������) = 0
              End If
          Else '��Ҳ�����
              If VoiceT > 4000 Then SEplay "AI-go" & CInt(2 * Rnd) & ".wav", True: VoiceT = 0 'voice
              ִ�ж��� 1, ������, "run", ��Ϸ�ٶ�, True
              EnmLastView(������) = Player(0).GetPosition
              EnmState(������) = 0
          End If
      Case 2 '��������ϻ����
          If Enemy(������).IsAnimationFinished Then
            ִ�ж��� 1, ������, "ref_shoot_onehanded", ��Ϸ�ٶ�, True
            EnmT(������) = 0
            EnmState(i) = 1
          End If
      End Select
  End Select
End Select
End Function
Public Function �������(��� As Long, �˺� As Long)
If ��� = 0 Then
  PlayerHP(0) = PlayerHP(0) - �˺�
  Ѫ�۲���ʱ�� = 80
Else
End If
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 27 'ESC�˳�
  Tv.ShowWinCursor True: If MsgBox("��ͣ�У�Ҫ���ڷ���ս����", vbYesNo) = vbYes Then Unload Me Else Tv.ShowWinCursor False
Case Asc("h") Or Asc("H") '������
  MsgBox Player(0).GetPosition.X & "//" & Player(0).GetPosition.Y & "//" & Player(0).GetPosition.z & vbCrLf
  CameraPozY = �������߶�(Vector(CameraPozX, CameraPozY, CameraPozZ), 1000, 1000) + PlayerHeight: �ӽ�������� True '����������
Case Asc("f") Or Asc("F")
  For j = 0 To UBound(MeshTVA)
    If ����ƽ��(MeshTVA(j).GetPosition, Player(0).GetPosition) < 2500 Then '������ҩ
    ������ϻ��(�������) = �޶�������ϻ��(�������): SEplay "��ǹ˨.wav", True
    End If
  Next
End Select
End Sub

'============================��������===============================
Private Sub Form_Load()
Randomize
BASSplay "\Audio\BGM\EOF-ս�����.mp3", 0, 1, Me.hWnd
'=====��������=====
SE.Init Me.hWnd
'=====��������=====
����ģʽ = True
��ʼ���ӽǲ��� 0, 0, 0, 0, 100
With Me
  .Width = Screen.Width
  .Height = Screen.Height
  .Left = 0
  .top = 0
  .Show '��ʾ��ǰ���ڣ�ÿ�ζ����ϴ���
End With
����X = Me.Width / 15360: ����Y = Me.Height / 11520
Tv.SetSearchDirectory App.Path & "\" '�趨��ͼ��ȡĿ¼Ϊ��ǰĿ¼
Tv.SetVSync True '��ֱͬ������
Tv.Init3DWindowed Me.hWnd '�ô���ģʽ����tv3d
Tv.ShowWinCursor False '�������
Inp.Initialize '��ʼ���������
Tv.SetAngleSystem TV_ANGLE_DEGREE
Scene.SetViewFrustum 45, 0 '���ӷ�Χ���ޣ����ӽǶ�45
'=====��ͼ=====
TF.LoadTexture "Pic\System\Ѫ��0.png", "Ѫ��0", , , TV_COLORKEY_USE_ALPHA_CHANNEL  '��ȡ��ΪXX����ͼ����ΪXX������͸��Alphaͨ��
TF.LoadTexture "Pic\BG\Back.jpg", "SKYBOX_Back" '��պ�
TF.LoadTexture "Pic\BG\Front.jpg", "SKYBOX_Front"
TF.LoadTexture "Pic\BG\Left.jpg", "SKYBOX_Left"
TF.LoadTexture "Pic\BG\Right.jpg", "SKYBOX_Right"
TF.LoadTexture "Pic\BG\Up.jpg", "SKYBOX_Up"
TF.LoadTexture "Pic\BG\Down.jpg", "SKYBOX_DOWN"
TF.LoadTexture "Pic\Flash\flash00.png", "flash00", , , TV_COLORKEY_USE_ALPHA_CHANNEL 'ǹ��
TF.LoadTexture "Pic\Flash\flash01.png", "flash01", , , TV_COLORKEY_USE_ALPHA_CHANNEL
TF.LoadTexture "Pic\System\UI��������.png", "UI��������", , , TV_COLORKEY_USE_ALPHA_CHANNEL 'UI��������

'=====�����Ч=====
Atmos.SkyBox_Enable True '������պ�
Atmos.SkyBox_SetTexture GetTex("SKYBOX_Front"), GetTex("SKYBOX_Back"), GetTex("SKYBOX_Left"), GetTex("SKYBOX_Right"), GetTex("SKYBOX_Up"), GetTex("SKYBOX_DOWN") '�趨��ͼ
Atmos.Fog_Enable True                              '������
Atmos.Fog_SetColor 1, 1, 1                         '��ɫRGBA�������
Atmos.Fog_SetParameters 4000, 4500, 0              '������룬��Զ���룬Ũ��
Atmos.Fog_SetType TV_FOG_LINEAR, TV_FOGTYPE_PIXEL  '�������
'=====����=====
MF.CreateMaterial "solid" '������Ϊsolid�Ĳ���
MF.SetAmbient GetMat("solid"), 0.1, 0.1, 0.1, 1       '������
MF.SetDiffuse GetMat("solid"), 1, 1, 1, 1 '��ɢ�⣬������Ĺ�����ɫ
MF.SetEmissive GetMat("solid"), 0, 0, 0, 1  '�Է���
MF.SetOpacity GetMat("solid"), 1 '��͸����
MF.SetSpecular GetMat("solid"), 0, 0, 0, 0   '�߹�ɫ
MF.SetPower GetMat("solid"), 60 'ɢ��ǿ��
'=====��Ӱ=====
LE.CreateDirectionalLight Vector(1, -1, 1), 1, 1, 1, , 1  '���һ��ƽ�й�
LE.SetSpecularLighting True '�߹⿪��
LE.SetLightProperties 0, True, True, False '�ƹ⿪��Ӱ��
'=====����=====
'����ײ���ʵ��
For i = 0 To UBound(Mesh): Set Mesh(i) = Scene.CreateMeshBuilder: Next
Mesh(0).LoadTVM "Map\�ձ�С��\�ձ�С��.tvm", True, True '��ȡ
Mesh(0).SetScale 1.1, 1.1, 1.1: Mesh(0).SetPosition 0, 0, 0

Mesh(1).LoadTVM "Map\�ձ�С��\�ǽ�.tvm", True, True '��ȡ
Mesh(1).SetScale 1.1, 1.1, 1.1: Mesh(2).SetPosition 0, 0, 0

Mesh(2).LoadTVM "Model\ZBD05\ZBD05.tvm", True, True
With Mesh(2): .SetScale 7.6, 7.6, 7.6: .SetPosition 360, -55, -870: .RotateY -15: End With
For i = 0 To UBound(Mesh) '��Ӱ
With Mesh(i)
.SetAlphaTest
.SetMaterial GetMat("solid")
.SetLightingMode TV_LIGHTING_NORMAL
End With
Next
Mesh(2).SetShadowCast True, True '����Ӱ��
'��������ͼTVAʵ��
For i = 0 To UBound(MeshTVA): Set MeshTVA(i) = Scene.CreateActor: Next
MeshTVA(0).LoadTVA "Model\car1\car1.tva", True, True '����
With MeshTVA(0): .SetScale 0.33, 0.33, 0.33: .SetPosition 858, 0, -400: .RotateY 90: End With

For i = 0 To UBound(MeshTVA): With MeshTVA(i) '��Ӱ
  .SetMaterial GetMat("solid")
  .SetLightingMode TV_LIGHTING_NORMAL
End With: Next
'������ײװ��ʵ��
For i = 0 To UBound(MeshSin): Set MeshSin(i) = Scene.CreateMeshBuilder: Next
MeshSin(0).LoadTVM "Map\�ձ�С��\ӣ��.tvm", True, True '��ȡ
MeshSin(0).SetScale 1.1, 1.1, 1.1: Mesh(0).SetPosition 0, 0, 0
For i = 0 To UBound(MeshSin): With MeshSin(i) '��Ӱ
  .SetAlphaTest
  .SetMaterial GetMat("solid")
  .SetLightingMode TV_LIGHTING_NORMAL
End With: Next
'=====��ɫ=====
For i = 0 To UBound(Player): Set Player(i) = Scene.CreateActor: Next
If UBound(Player) > 0 Then
For i = 1 To UBound(Player)
With Player(i)
.SetMaterial GetMat("solid") '�趨����
.SetLightingMode TV_LIGHTING_NORMAL '�趨����ģʽ
.SetAnimationByName ("run") 'ִ�еĶ�������
.PlayAnimation ��Ϸ�ٶ� '���Ŷ����ٶ�
.SetScale 0.31, 0.31, 0.31 '�趨��С
End With
Next
End If

For i = 1 To UBound(Enemy)
Set Enemy(i) = Scene.CreateActor '��ɫ��ʼ��
With Enemy(i)
.LoadTVA ("Player\����\����.tva") '��ȡ��ɫ
.SetMaterial GetMat("solid") '�趨����
.SetLightingMode TV_LIGHTING_NORMAL '�趨����ģʽ
.SetScale 0.31, 0.31, 0.31 '�趨��С
.SetPosition Rnd * 40 + 80, 90, Rnd * 40 + 80 '�趨ģ��λ��
'.SetRotation -90, 0, 180
.SetAnimationByName ("run") 'ִ�еĶ�������
.PlayAnimation ��Ϸ�ٶ� '���Ŷ����ٶ�
End With
Next

For i = 1 To UBound(Enemy)
Set EnemyGun(i) = Scene.CreateActor '���������ʼ�����ؼ�
With EnemyGun(i)
'.LoadTVA "Weapon\M16\v_M16.tva", True, True
.SetLightingMode TV_LIGHTING_NORMAL
.SetMaterial GetMat("solid")
.SetScale 0.075, 0.075, 0.075 '�趨��С
End With
Next
'=====��Ч====
'=====����=====
Lrc.NormalFont_Create "", "����", 20, True, False, False
��ʼ���ӽǲ��� 852, -2, -400, 0, 100

'================================��ѭ��=======================================
Do
 VoiceT = VoiceT + Tv.TimeElapsed
 ��������(3) = 0: ��������(4) = 0
 If PlayerHP(0) = 0 Then GoTo �ƶ��������
 Inp.GetMouseState Mx, My, B1, B2, , , Roll           '���������Ϣ
 CameraAngX = CameraAngX + 0.11 * Mx * ��Ϸ�ٶ�
 CameraAngY = CameraAngY + 0.11 * My * ��Ϸ�ٶ�
If CameraAngY > 40 Then CameraAngY = 40
If CameraAngY < -60 Then CameraAngY = -60
�����ƶ� = False: If ������� < 20 Then GoTo ��������ƶ�
    If Inp.IsKeyPressed(TV_KEY_LEFTSHIFT) Then 'ǰ
      If Inp.IsKeyPressed(TV_KEY_W) Then
      �����ƶ� = True: ������������ = True
      ִ�ж��� 0, 0, "idle1", ��Ϸ�ٶ�, False
      CameraPozƫ��(0) = 0.4 + CameraPozƫ��(0): If CameraPozƫ��(0) >= 6.28 Then CameraPozƫ��(0) = 0
      CameraPozƫ��(2) = Sin(CameraPozƫ��(0)) * ��Ϸ�ٶ� * 1.5
      ��������(3) = ��������(3) + Cos(Math.Deg2Rad(CameraAngX - 90)) * 2 * ��Ϸ�ٶ�
      ��������(4) = ��������(4) - Sin(Math.Deg2Rad(CameraAngX - 90)) * 2 * ��Ϸ�ٶ�
      GoTo �ƶ��������
      End If
    Else
      ������������ = False
    End If
    If Inp.IsKeyPressed(TV_KEY_W) Then 'ǰ
      �����ƶ� = True
      ������������ = False
      ��������(3) = ��������(3) + Cos(Math.Deg2Rad(CameraAngX - 90)) * 1 * ��Ϸ�ٶ�
      ��������(4) = ��������(4) - Sin(Math.Deg2Rad(CameraAngX - 90)) * 1 * ��Ϸ�ٶ�
    End If
    If Inp.IsKeyPressed(TV_KEY_S) Then '��
     ��������(3) = ��������(3) + Cos(Math.Deg2Rad(CameraAngX + 90)) * 0.5 * ��Ϸ�ٶ�
     ��������(4) = ��������(4) - Sin(Math.Deg2Rad(CameraAngX + 90)) * 0.5 * ��Ϸ�ٶ�
     �����ƶ� = True
    End If
    If Inp.IsKeyPressed(TV_KEY_A) Then '��
     ��������(3) = ��������(3) + Cos(Math.Deg2Rad(CameraAngX + 180)) * 1 * ��Ϸ�ٶ�
     ��������(4) = ��������(4) - Sin(Math.Deg2Rad(CameraAngX + 180)) * 1 * ��Ϸ�ٶ�
     �����ƶ� = True
    End If
    If Inp.IsKeyPressed(TV_KEY_D) Then '��
     ��������(3) = ��������(3) + Cos(Math.Deg2Rad(CameraAngX)) * 1 * ��Ϸ�ٶ�
     ��������(4) = ��������(4) - Sin(Math.Deg2Rad(CameraAngX)) * 1 * ��Ϸ�ٶ�
     �����ƶ� = True
    End If
�ƶ��������:
    If Inp.IsKeyPressed(TV_KEY_ALT_LEFT) Or Inp.IsKeyPressed(TV_KEY_LEFTCONTROL) Then  '�׷�
      If PlayerHeight > 11 Then PlayerHeight = PlayerHeight - 1.5
    Else
      If PlayerHeight < 17 Then If PlayerHP(0) > 0 Then PlayerHeight = PlayerHeight + 1
    End If
    If �����ƶ� = True And Inp.IsKeyPressed(TV_KEY_LEFTSHIFT) = False Then
      CameraPozƫ��(0) = 0.2 + CameraPozƫ��(0)
      CameraPozƫ��(2) = Sin(CameraPozƫ��(0)) * ��Ϸ�ٶ�
    End If
'=====������ײ=====
If �����ƶ� = False Then GoTo ��������ƶ�
Dim ��ײ��ʱ As Boolean
CameraPozX = CameraPozX + ��������(3)
CameraPozZ = CameraPozZ + ��������(4)
CameraPozY = �������߶�(Vector(CameraPozX, CameraPozY, CameraPozZ), 5, 3)
For i = 0 To UBound(Mesh)
  If Mesh(i).Collision(Vector(CameraPozX - 10, CameraPozY + PlayerHeight, CameraPozZ), Vector(CameraPozX + 10, CameraPozY + PlayerHeight, CameraPozZ)) Or _
  Mesh(i).Collision(Vector(CameraPozX, CameraPozY + PlayerHeight, CameraPozZ - 10), Vector(CameraPozX, CameraPozY + PlayerHeight, CameraPozZ + 10)) Then
  �ӽ�������� False: GoTo ��������ƶ�
  End If
Next
For i = 0 To UBound(MeshTVA)
  If MeshTVA(i).Collision(Vector(CameraPozX - 10, CameraPozY + PlayerHeight, CameraPozZ), Vector(CameraPozX + 10, CameraPozY + PlayerHeight, CameraPozZ)) Or _
  MeshTVA(i).Collision(Vector(CameraPozX, CameraPozY + PlayerHeight, CameraPozZ - 10), Vector(CameraPozX, CameraPozY + PlayerHeight, CameraPozZ + 10)) Then
  �ӽ�������� False: GoTo ��������ƶ�
  End If
Next
�ӽ�������� True
��������ƶ�:
'�趨�����
Camera.SetPosition CameraPozX, CameraPozY + PlayerHeight + CameraPozƫ��(2), CameraPozZ
Camera.SetRotation CameraAngY, CameraAngX, 0
'===============��ɫ�¼�===============
For i = 1 To UBound(Enemy)
With Enemy(i)
  If EnmHP(i) < 0 Then GoTo �����˶���
    AIadv 1, i, 2, 0.01
End With
�����˶���:
Next
If PlayerHP(0) <= 0 Then '���������¼�
  ������������ = True
  If PlayerHP(0) < 0 Then PlayerHP(0) = 0
  If PlayerHeight > 0.1 Then PlayerHeight = PlayerHeight - 0.3
  If PlayerHeight <= 0.1 Then
    If MsgBox("�������ˡ��Ƿ�ѡ��������ϯ��", vbYesNo) = vbYes Then
    ��ʼ���ӽǲ��� 0, 0, -3, 0, 300
    Else
    Unload Me
    End If
  End If
End If
'===============��������===============
Player(0).SetPosition CameraPozX, CameraPozY + PlayerHeight + CameraPozƫ��(2) + ����λ��ƫ��(2), CameraPozZ
Player(0).SetRotation 0, CameraAngX + 90, CameraAngY
'=====��������=====
For i = 1 To UBound(Enemy): With EnemyGun(i)
  .SetMatrix Enemy(i).GetBoneMatrix(Enemy(i).GetBoneID("Bip01 R Hand")) '�󶨵���
  .RotateY -70
  .RotateZ -95
  .MoveRelative 0.2, -1.2, 0.17
End With: Next
'===============��������Ⱦ===============
Tv.Clear '����
Atmos.Fog_Enable False
Atmos.Atmosphere_Render '��Ⱦ����
Atmos.Fog_Enable True
'For i = 1 To UBound(Enemy): EnemyGun(i).Render: Next
For i = 1 To UBound(Enemy): Enemy(i).Render: Next
For i = 0 To UBound(MeshTVA): MeshTVA(i).Render: Next
For i = 0 To UBound(Mesh): Mesh(i).Render: Next
For i = 0 To UBound(MeshSin): MeshSin(i).Render: Next
Scene.FinalizeShadows '��ȾӰ��
'===============������Ⱦ===============
Scrͼ��.Draw_Sprite GetTex("UI��������"), 15, Me.Height / 15 - 135 '��������UI
Select Case PlayerHP(0)
Case Is > 50: UI������ɫ = RGBA(0.5, 1, 0.5, 0.6)
Case Is > 20: UI������ɫ = RGBA(1, 1, 0.5, 0.6)
Case Else: UI������ɫ = RGBA(1, 0.5, 0.5, 0.6)
End Select
Lrc.Action_BeginText
If ������ϻ(�������) > �޶�������ϻ(�������) Then
  Lrc.NormalFont_DrawText "1+" & (������ϻ(�������) - 1) & "/" & �޶�������ϻ(�������), 50, Me.Height \ 15 - 100, UI������ɫ, 1
Else
  Lrc.NormalFont_DrawText ������ϻ(�������) & "/" & �޶�������ϻ(�������), 50, Me.Height \ 15 - 100, UI������ɫ, 1
End If
Lrc.NormalFont_DrawText ������(�������) & " " & ������ϻ��(�������) & "/" & �޶�������ϻ��(�������), 50, Me.Height \ 15 - 60, UI������ɫ, 1
DrawText "���ȳ���", Word, RGBA(0, 1, 0, 1)
Lrc.Action_EndText
Tv.RenderToScreen
DoEvents
If ����ģʽ = True Then
End If
Loop
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub Tim����_Timer()
ʱ���� = 1 + ʱ����
Select Case ʱ����
Case 1: Word = "��ӭ����ɵ�ƻ��ػ�ӭ����ɵ��"
End Select
End Sub
