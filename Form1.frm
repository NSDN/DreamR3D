VERSION 5.00
Begin VB.Form Stage���� 
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   360
   End
   Begin VB.Timer Tim�������� 
      Interval        =   4000
      Left            =   0
      Top             =   360
   End
End
Attribute VB_Name = "Stage����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mesh(0 To 2) As TVMesh: Dim MeshTVA(0 To 3) As TVActor: Dim MeshSin(0 To 0) As TVMesh
Dim Enemy(1 To 9) As TVActor, EnemyGun(1 To 25) As TVActor, EnmType(1 To 25) As Long, EnmGunFire(1 To 25) As Long
Dim EnmLastView(1 To 25) As TV_3DVECTOR
Dim ��ɫλ��(0 To 3) As TV_3DVECTOR '0���12��ɫ3��ת
Dim ��ɫ����(0 To 3) As TV_3DVECTOR
Dim VoiceT As Long, T�� As Long
'������ͨ������������
Dim LRCname As String, LRCtext As String, LRCcolor As Single, LRC��ʧʱ�� As Long
Dim �Ѷ� As Long
Dim ����ģʽ As Boolean
Private Function CreatLRC(���� As String, ���� As String, ������ɫ As Single)
LRCname = ����: LRCtext = ����: LRCcolor = ������ɫ: LRC��ʧʱ�� = 3000
End Function
Private Function DrawLRC(���� As String, ���� As String, ������ɫ As Single)
Dim LSX As Long
LSX = Me.Width \ 30 - Len(���� & ����) * 10 - 10
Lrc.NormalFont_DrawText ����, LSX, Me.Height \ 15 - 100, ������ɫ, 1
Lrc.NormalFont_DrawText ����, LSX + Len(����) * 20 + 10, Me.Height \ 15 - 100, RGBA(1, 1, 1, 1), 1
End Function
Public Function �������߶�(λ�� As TV_3DVECTOR, ԽҰ�߶� As Single, �����ٶ� As Single) As Single
Dim Result As TV_COLLISIONRESULT
If Mesh(0).AdvancedCollision(Vector(λ��.X, λ��.Y + ԽҰ�߶�, λ��.z), Vector(λ��.X, λ��.Y - �����ٶ�, λ��.z), Result) Then
  �������߶� = Result.vCollisionImpact.Y
Else
  �������߶� = λ��.Y - �����ٶ�
End If
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
    If Mesh(0).Collision(Vector(��ʱ����(1).X + 4, ��ʱ����(1).Y, ��ʱ����(1).z), Vector(��ʱ����(1).X - 4, ��ʱ����(1).Y, ��ʱ����(1).z)) Then Enemy(������).SetPosition ��ʱ����(1).X - 2 * ����(1), ��ʱ����(1).Y, ��ʱ����(1).z
    If Mesh(0).Collision(Vector(��ʱ����(1).X, ��ʱ����(1).Y, ��ʱ����(1).z + 4), Vector(��ʱ����(1).X, ��ʱ����(1).Y, ��ʱ����(1).z - 4)) Then Enemy(������).SetPosition ��ʱ����(1).X, ��ʱ����(1).Y, ��ʱ����(1).z - 2 * ����(2)
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
      Select Case EnmState(������)
      Case 0 '������������
        AImoveTo ������, 0, Player(0).GetPosition, 999
        AImoveTo ������, 1, Player(0).GetPosition, �ƶ��ٶ�
        If ����ƽ��(Player(0).GetPosition, Enemy(������).GetPosition) < 400 Then
          ִ�ж��� 1, ������, "ref_shoot_wrench", ��Ϸ�ٶ�, False
          If Ѫ�۲���ʱ�� <= 20 Then ������� 0, �Ѷ�
          EnmState(������) = 1
        End If
      Case 1 '������������
        If Enemy(������).IsAnimationFinished Then
          ִ�ж��� 1, ������, "run2", ��Ϸ�ٶ�, False
          EnmState(������) = 0
        End If
      End Select
  Case 2 '==����==
      AImoveTo ������, 1, Player(0).GetPosition, 0
      ��ײ(0) = Mesh(0).Collision(Enemy(������).GetPosition, Player(0).GetPosition)
      If ��ײ(0) = False Then EnmLastView(������) = Vector(CameraPozX + 10 * ��������(3), CameraPozY, CameraPozZ + 10 * ��������(4))
      Select Case EnmState(������)
      Case 0 '�����ӽ���ҡ���
          If ��ײ(0) = False Then '��ҿ���
            AImoveTo ������, 0, Player(0).GetPosition, 999
            If ����ƽ��(Enemy(������).GetPosition, Player(0).GetPosition) > 6000 Then
              AImoveTo ������, 1, Player(0).GetPosition, �ƶ��ٶ�
            Else
              If VoiceT > 4000 Then SEplay "AI-fire" & CInt(2 * Rnd) & ".wav", True: VoiceT = 0 'voice
              Select Case EnmType(������)
              Case 0: ִ�ж��� 1, ������, "ref_shoot_onehanded", ��Ϸ�ٶ�, True
              Case 1: ִ�ж��� 1, ������, "lmg_fire", ��Ϸ�ٶ�, True
              End Select
              EnmState(������) = 1
            End If
          Else '��Ҳ�����
            If EnmLastView(i).X <> 0 Then AImoveTo i, 1, EnmLastView(������), 2 * �ƶ��ٶ�
          End If
      Case 1 '���������������
          If ��ײ(0) = False Then
              AImoveTo ������, 0, Player(0).GetPosition, 999
              If ����ƽ��(Enemy(������).GetPosition, Player(0).GetPosition) < 9000 Then
                If EnmT(������) Mod 300 < 20 Then '����
                    SEplay "�����ŵ�0.wav", False
                    SEplay "����" & CInt(Rnd * 6) & ".wav", False
                    If Ѫ�۲���ʱ�� <= 60 Then ������� 0, �Ѷ�
                End If '��������
                If EnmT(������) Mod 300 < 90 Then
                  Scrͼ��.Draw_Line3D Enemy(������).GetPosition.X, Enemy(������).GetPosition.Y + 4, Enemy(������).GetPosition.z, CameraPozX + Rnd * 6 - 3, CameraPozY + Rnd * 6 - 3 + PlayerHeight, CameraPozZ + Rnd * 6 - 3, RGBA(1, 1, 0, 0.3)
                  If ����(0) > 1 Then LE.CreatePointLight Enemy(������).GetPosition, 1, 1, 0, 50, "EnmGunFire" & ������
                End If
                EnmT(������) = EnmT(������) + Tv.TimeElapsed
                If EnmT(������) > 4000 Then '����
                    If VoiceT > 3000 Then SEplay "AI-reload.wav", True: VoiceT = 0 'voice
                    SEplay "AI_reload.wav", False
                    Select Case EnmType(������)
                    Case 0: ִ�ж��� 1, ������, "ref_reload_onehanded", ��Ϸ�ٶ�, False
                    Case 1: ִ�ж��� 1, ������, "lmg_reload", ��Ϸ�ٶ�, False
                    End Select
                    EnmState(i) = 2
                End If
              Else '����Զ���л�׷��
                  Select Case EnmType(������)
                  Case 0: ִ�ж��� 1, ������, "run", ��Ϸ�ٶ�, True
                  Case 1: ִ�ж��� 1, ������, "walk", ��Ϸ�ٶ�, True
                  End Select
                  EnmState(������) = 0
              End If
          Else '��Ҳ�����
              If VoiceT > 4000 Then SEplay "AI-go" & CInt(2 * Rnd) & ".wav", True: VoiceT = 0 'voice
              Select Case EnmType(������)
              Case 0: ִ�ж��� 1, ������, "run", ��Ϸ�ٶ�, True
              Case 1: ִ�ж��� 1, ������, "walk", ��Ϸ�ٶ�, True
              End Select
              EnmLastView(������) = Player(0).GetPosition
              EnmState(������) = 0
          End If
      Case 2 '��������ϻ����
          If Enemy(������).IsAnimationFinished Then
            Select Case EnmType(������)
            Case 0: ִ�ж��� 1, ������, "run", ��Ϸ�ٶ�, True
            Case 1: ִ�ж��� 1, ������, "walk", ��Ϸ�ٶ�, True
            End Select
            EnmT(������) = 0
            EnmState(������) = 1
          End If
      End Select
  End Select
End Select
End Function
Public Function �������(��� As Long, �˺� As Long)
If ��� = 0 Then
  PlayerHP(0) = PlayerHP(0) - �˺�
  GF.Flash 1, 0, 0, 500
  If PlayerHeight > 10 Then PlayerHeight = PlayerHeight - 5
  If ������(1) < 2 Then ������(1) = ������(1) - 6
  Ѫ�۲���ʱ�� = 80
Else
End If
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 27 'ESC�˳�
  Tv.ShowWinCursor True: If MsgBox("��ͣ�У�Ҫ���ڷ���ս����", vbYesNo) = vbYes Then Unload Me Else Tv.ShowWinCursor False
Case Asc("h") Or Asc("H") '������
  MsgBox "��ǰλ�ã�" & Player(0).GetPosition.X & "//" & Player(0).GetPosition.Y & "//" & Player(0).GetPosition.z & vbCrLf
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
'=====��������=====
BASSready Me.hWnd
SE.Init Me.hWnd
'=====��������=====
����ģʽ = True
��ʼ���ӽǲ��� 0, 0, 0, 0, 100
With Me
  .Width = 15360
  .Height = 11520
  .Left = 0
  .top = 0
  .Show '��ʾ��ǰ���ڣ�ÿ�ζ����ϴ���
End With
׼��X = Me.Width \ 30: ׼��Y = Me.Height \ 30
����X = Me.Width / 15360: ����Y = Me.Height / 11520
Tv.SetSearchDirectory App.Path & "\" '�趨��ͼ��ȡĿ¼Ϊ��ǰĿ¼
Tv.SetVSync True '��ֱͬ������
Tv.Init3DWindowed Me.hWnd '�ô���ģʽ����tv3d
Tv.ShowWinCursor False '�������
Inp.Initialize '��ʼ���������
Tv.SetAngleSystem TV_ANGLE_DEGREE
Scene.SetViewFrustum 45, 1200  '���ӷ�Χ�����ӽǶ�45
If ����(2) > 0 Then Scene.SetTextureFilter TV_FILTER_BILINEAR
'=====·��=====
'=====��ͼ=====
Dim BGname As String: BGname = "nuke"
TF.LoadTexture "Pic\System\Ѫ��0.png", "Ѫ��0", , , TV_COLORKEY_USE_ALPHA_CHANNEL  '��ȡ��ΪXX����ͼ����ΪXX������͸��Alphaͨ��
TF.LoadTexture "Pic\BG\" & BGname & "_Back.jpg", "SKYBOX_Back" '��պ�
TF.LoadTexture "Pic\BG\" & BGname & "_Front.jpg", "SKYBOX_Front"
TF.LoadTexture "Pic\BG\" & BGname & "_Left.jpg", "SKYBOX_Left"
TF.LoadTexture "Pic\BG\" & BGname & "_Right.jpg", "SKYBOX_Right"
TF.LoadTexture "Pic\BG\" & BGname & "_Top.jpg", "SKYBOX_Up"
TF.LoadTexture "Pic\BG\" & BGname & "_Down.jpg", "SKYBOX_DOWN"
TF.LoadTexture "Pic\Flash\flash00.png", "flash00", , , TV_COLORKEY_USE_ALPHA_CHANNEL 'ǹ��
TF.LoadTexture "Pic\Flash\flash01.png", "flash01", , , TV_COLORKEY_USE_ALPHA_CHANNEL
TF.LoadTexture "Pic\System\UI��������.png", "UI��������", , , TV_COLORKEY_USE_ALPHA_CHANNEL 'UI��������
TF.LoadTexture "Pic\Flash\sun.jpg", "sun", , , TV_COLORKEY_USE_ALPHA_CHANNEL '̫��
TF.LoadTexture "Pic\Flash\flare1.jpg", "flare1" '̫������
TF.LoadTexture "Pic\Flash\flare2.jpg", "flare2"
TF.LoadTexture "Pic\Flash\flare3.jpg", "flare3"
TF.LoadTexture "Pic\Flash\flare4.jpg", "flare4"

'=====�����Ч=====
Atmos.SkyBox_Enable True '������պ�
Atmos.SkyBox_SetTexture GetTex("SKYBOX_Front"), GetTex("SKYBOX_Back"), GetTex("SKYBOX_Left"), GetTex("SKYBOX_Right"), GetTex("SKYBOX_Up"), GetTex("SKYBOX_DOWN") '�趨��ͼ
Atmos.Fog_Enable True                              '������
Atmos.Fog_SetColor 0.1, 0.1, 0.1                         '��ɫRGBA�������
Atmos.Fog_SetParameters 50, 1000, 0              '������룬��Զ���룬Ũ��
Atmos.Fog_SetType TV_FOG_LINEAR, TV_FOGTYPE_PIXEL  '�������
'=====����=====
MF.CreateMaterial "solid" '����ģ�����ʲ���
MF.SetAmbient GetMat("solid"), 0.08, 0.04, 0.02, 1       '������
MF.SetDiffuse GetMat("solid"), 1, 1, 1, 1 '��ɢ�⣬������Ĺ�����ɫ
MF.SetEmissive GetMat("solid"), 0, 0, 0, 0 '�Է���
MF.SetOpacity GetMat("solid"), 1  '��͸����
MF.SetSpecular GetMat("solid"), 0, 0, 0, 0 '�߹�ɫ
MF.SetPower GetMat("solid"), 60 'ɢ��ǿ��

MF.CreateMaterial "map" '������ͼ�߹����
MF.SetAmbient GetMat("map"), 0.08, 0.04, 0.02, 1       '������
MF.SetDiffuse GetMat("map"), 1, 1, 1, 1 '��ɢ�⣬������Ĺ�����ɫ
MF.SetEmissive GetMat("map"), 0.08, 0.04, 0.02, 1 '�Է���
MF.SetOpacity GetMat("map"), 1  '��͸����
MF.SetSpecular GetMat("map"), 1, 1, 1, 1 '�߹�ɫ
MF.SetPower GetMat("map"), 15 'ɢ��ǿ��
'=====��Ӱ=====
'����
Atmos.Sun_SetBillboardSize 0.7 '����̫����ͼ��С
Atmos.Sun_SetPosition -400, 400, -400  '����̫��λ��
Atmos.Sun_SetTexture GetTex("sun") '����̫����ͼ
If ����(0) > 1 And ����(4) >= 30 Then Atmos.Sun_Enable True 'ʹ̫����ͼ��Ч
Atmos.LensFlare_SetLensNumber 4 '���β���
Atmos.LensFlare_SetLensParams 1, GetTex("flare1"), 7.5, 40, RGBA(1, 1, 1, 0.5), RGBA(1, 1, 1, 0.5)
Atmos.LensFlare_SetLensParams 2, GetTex("flare2"), 3, 18, RGBA(1, 1, 1, 0.5), RGBA(1, 1, 1, 0.5)
Atmos.LensFlare_SetLensParams 3, GetTex("flare3"), 4, 15, RGBA(1, 1, 1, 0.5), RGBA(0.7, 1, 1, 0.5)
Atmos.LensFlare_SetLensParams 4, GetTex("flare4"), 3, 6, RGBA(1, 0.1, 0, 0.5), RGBA(0.5, 1, 1, 0.5)
If ����(0) > 1 And ����(4) > 30 Then Atmos.LensFlare_Enable True 'ʹ������Ч
'��Ӱ
LE.CreateDirectionalLight Vector(1, -1, 1), 1, 1, 1, "sun", 1 '���һ��ƽ�й�
If ����(0) > 0 Then LE.SetSpecularLighting True '�߹⿪��
If ����(0) > 0 Then LE.SetLightProperties 0, True, True, False '�ƹ⿪��Ӱ��
'=====����=====
'����ײ���ʵ��
For i = 0 To UBound(Mesh): Set Mesh(i) = Scene.CreateMeshBuilder: Next
Mesh(0).LoadTVM "Map\�ձ�С��\�ձ�С��.tvm", True, True '��ȡ
Mesh(0).SetScale 1.1, 1.1, 1.1: Mesh(2).SetPosition 0, 0, 0: Mesh(0).SetRotation 0, 0, 0

Mesh(1).LoadTVM "Map\�ձ�С��\�ǽ�.tvm", True, True '��ȡ
Mesh(1).SetScale 1.1, 1.1, 1.1: Mesh(2).SetPosition 0, 0, 0

Mesh(2).LoadTVM "Model\ZBD05\ZBD05.tvm", True, True
With Mesh(2): .SetScale 7.6, 7.6, 7.6: .SetPosition 360, -55, -870: .RotateY -15: End With
For i = 0 To UBound(Mesh) '��Ӱ
With Mesh(i)
.SetAlphaTest True
.SetMaterial GetMat("map")
If ����(0) > 0 Then .SetLightingMode TV_LIGHTING_NORMAL
End With
Next
'Mesh(1).SetShadowCast True, True '����Ӱ��
'��������ͼTVAʵ��
For i = 0 To UBound(MeshTVA): Set MeshTVA(i) = Scene.CreateActor: Next
MeshTVA(0).LoadTVA "Model\����\����.tva", True, True '����
With MeshTVA(0): .SetScale 0.33, 0.33, 0.33: .SetPosition -900, 81, 345: .RotateY -120: End With

MeshTVA(1).LoadTVA "Model\���ÿ���\���ÿ���.tva", True, True '������
With MeshTVA(1): .SetScale 0.31, 0.31, 0.31: .SetPosition 847, -2, 98: .RotateY 100: End With

MeshTVA(2).LoadTVA "Model\����\����.tva", True, True '·����
With MeshTVA(2): .SetScale 0.33, 0.33, 0.33: .SetPosition -150, -3, 727: .RotateY -90: End With

MeshTVA(3).LoadTVA "Model\���ÿ���\���ÿ���.tva", True, True '·����
With MeshTVA(3): .SetScale 0.31, 0.31, 0.31: .SetPosition 220, -3, 595: .RotateY 140: End With
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
  .SetMaterial GetMat("map")
  If ����(0) > 0 Then .SetLightingMode TV_LIGHTING_NORMAL
End With: Next
'=====��ɫ=====
For i = 0 To UBound(Player): Set Player(i) = Scene.CreateActor: Next
If UBound(Player) > 0 Then
For i = 1 To UBound(Player)
With Player(i)
.SetMaterial GetMat("solid") '�趨����
If ����(0) > 0 Then .SetLightingMode TV_LIGHTING_NORMAL '�趨����ģʽ
.SetAnimationByName ("run") 'ִ�еĶ�������
.PlayAnimation ��Ϸ�ٶ� '���Ŷ����ٶ�
.SetScale 0.31, 0.31, 0.31 '�趨��С
End With
Next
End If

For i = 1 To UBound(Enemy)
Set Enemy(i) = Scene.CreateActor '��ɫ��ʼ��
With Enemy(i)
.LoadTVA ("Player\����\����.tva") '��ȡģ��
.SetMaterial GetMat("solid") '�趨����
If ����(0) > 0 Then .SetLightingMode TV_LIGHTING_NORMAL '�趨����ģʽ
.SetScale 0.31, 0.31, 0.31 '�趨��С
.SetPosition Rnd * 40 + 80, 60, Rnd * 40 + 80 '�趨ģ��λ��
'.SetRotation -90, 0, 180
.SetAnimationByName ("run") 'ִ�еĶ�������
.PlayAnimation ��Ϸ�ٶ� '���Ŷ����ٶ�
End With
EnmHP(i) = 5
Next

For i = 1 To UBound(Enemy)
Set EnemyGun(i) = Scene.CreateActor '���������ʼ�����ؼ�
With EnemyGun(i)
'.LoadTVA "Weapon\M16\v_M16.tva", True, True
If ����(0) > 0 Then .SetLightingMode TV_LIGHTING_NORMAL
.SetMaterial GetMat("solid")
.SetScale 0.075, 0.075, 0.075 '�趨��С
End With
Next
'=====��Ч====
GF.FadeIn 1000
'=====����=====
Lrc.NormalFont_Create "", "����", 25, False, False, False
��ʼ���ӽǲ��� 0, 0, 0, 0, 100
GunLoad "AKS-74U", 1, True
GunLoad "M16", 0, True
BASSplay "Audio\BGM\霺��ꥰ��å�.mp3", 0, 1, Me.hWnd
'================================��ѭ��=======================================
Do
 VoiceT = VoiceT + Tv.TimeElapsed
 ��������(3) = 0: ��������(4) = 0
 If PlayerHP(0) = 0 Then GoTo �ƶ��������
 Inp.GetMouseState Mx, My, B1, B2, , , Roll           '���������Ϣ
 CameraAngX = CameraAngX + 0.11 * Mx * ��Ϸ�ٶ�
 CameraAngY = CameraAngY + 0.11 * My * ��Ϸ�ٶ�
If CameraAngY > 50 Then CameraAngY = 50
If CameraAngY < -60 Then CameraAngY = -60
�����ƶ� = False
    If Inp.IsKeyPressed(TV_KEY_LEFTSHIFT) Then 'ǰ
      If Inp.IsKeyPressed(TV_KEY_W) Then
      �����ƶ� = True: ������������ = True
      ִ�ж��� 0, 0, "idle1", ��Ϸ�ٶ�, False
      CameraPozƫ��(0) = 0.02 * Tv.TimeElapsed + CameraPozƫ��(0): If CameraPozƫ��(0) >= 6.28 Then CameraPozƫ��(0) = 0
      CameraPozƫ��(2) = Sin(CameraPozƫ��(0)) * ��Ϸ�ٶ� * 1.5
      ��������(3) = ��������(3) + Cos(Math.Deg2Rad(CameraAngX - 90)) * 0.08 * Tv.TimeElapsed * ��Ϸ�ٶ�
      ��������(4) = ��������(4) - Sin(Math.Deg2Rad(CameraAngX - 90)) * 0.08 * Tv.TimeElapsed * ��Ϸ�ٶ�
      GoTo �ƶ��������
      End If
    Else
      ������������ = False
    End If
    If Inp.IsKeyPressed(TV_KEY_W) Then 'ǰ
      �����ƶ� = True
      ������������ = False
      ��������(3) = ��������(3) + Cos(Math.Deg2Rad(CameraAngX - 90)) * 0.04 * Tv.TimeElapsed * ��Ϸ�ٶ�
      ��������(4) = ��������(4) - Sin(Math.Deg2Rad(CameraAngX - 90)) * 0.04 * Tv.TimeElapsed * ��Ϸ�ٶ�
    End If
    If Inp.IsKeyPressed(TV_KEY_S) Then '��
     ��������(3) = ��������(3) + Cos(Math.Deg2Rad(CameraAngX + 90)) * 0.02 * Tv.TimeElapsed * ��Ϸ�ٶ�
     ��������(4) = ��������(4) - Sin(Math.Deg2Rad(CameraAngX + 90)) * 0.02 * Tv.TimeElapsed * ��Ϸ�ٶ�
     �����ƶ� = True
    End If
    If Inp.IsKeyPressed(TV_KEY_A) Then '��
     ��������(3) = ��������(3) + Cos(Math.Deg2Rad(CameraAngX + 180)) * 0.04 * Tv.TimeElapsed * ��Ϸ�ٶ�
     ��������(4) = ��������(4) - Sin(Math.Deg2Rad(CameraAngX + 180)) * 0.04 * Tv.TimeElapsed * ��Ϸ�ٶ�
     �����ƶ� = True
    End If
    If Inp.IsKeyPressed(TV_KEY_D) Then '��
     ��������(3) = ��������(3) + Cos(Math.Deg2Rad(CameraAngX)) * 0.04 * Tv.TimeElapsed * ��Ϸ�ٶ�
     ��������(4) = ��������(4) - Sin(Math.Deg2Rad(CameraAngX)) * 0.04 * Tv.TimeElapsed * ��Ϸ�ٶ�
     �����ƶ� = True
    End If
�ƶ��������:
    If Inp.IsKeyPressed(TV_KEY_ALT_LEFT) Or Inp.IsKeyPressed(TV_KEY_LEFTCONTROL) Then  '�׷�
      If PlayerHeight > 11 Then PlayerHeight = PlayerHeight - 1.2
    Else
      If PlayerHeight < 17 Then If PlayerHP(0) > 0 Then PlayerHeight = PlayerHeight + 0.8
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
CameraPozY = �������߶�(Vector(CameraPozX, CameraPozY, CameraPozZ), 20, 3)
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
For i = 1 To UBound(Enemy)
  If EnmHP(i) > 0 And ����ƽ��(Vector(CameraPozX, CameraPozY + PlayerHeight, CameraPozZ), Enemy(i).GetPosition) < 400 Then
  �ӽ�������� False: GoTo ��������ƶ�
  End If
Next
�ӽ�������� True
��������ƶ�:
'===============��׼���===============
If Inp.IsKeyPressed(TV_KEY_1) And ������� <> 0 Then GunLoad ������(0), 0, False
If Inp.IsKeyPressed(TV_KEY_2) And ������� <> 1 Then GunLoad ������(1), 1, False
If B1 = True Then
  If ����״̬ = 0 Then ����״̬ = 1: ������ = 0
End If
If Inp.IsKeyPressed(TV_KEY_R) Then '����ϻ
If ����״̬ <> 2 Then
  If ������ϻ��(�������) > 0 And ������ϻ(�������) < �޶�������ϻ(�������) And ������������ = False Then
  GunSE ������(�������), "reload.wav", 1
  ����״̬ = 2
  ������ = 0
  End If
End If
End If
If ������������ = True Then ����״̬ = 0
Select Case ����״̬
Case -1 '===װ��===
  If Player(0).IsAnimationFinished Then ����״̬ = 0
Case 0 '===��ֹ===
  If �����ƶ� = True Then
   ����λ��ƫ��(0) = 0.01 * Tv.TimeElapsed + ����λ��ƫ��(0)
   ����λ��ƫ��(2) = Sin(����λ��ƫ��(0)) * ��Ϸ�ٶ� / 18
  Else
   ����λ��ƫ��(0) = 0.001 * Tv.TimeElapsed + ����λ��ƫ��(0)
   ����λ��ƫ��(2) = Sin(����λ��ƫ��(0)) * ��Ϸ�ٶ� / 18
  End If
  If ����λ��ƫ��(0) >= 6.28 Then ����λ��ƫ��(0) = 0
Case 1 '===����===
  If ������ = 0 Then
  If ������ϻ(�������) <= 0 Then '����
    GunSE "ͨ��", "Empty.wav", 1
    ����״̬ = 0
  Else
    GunSE ������(�������), "shoot.wav", 1
    i = 1
    Do Until InStr(Player(0).GetAnimationName(i), "shoot") > 0
     i = 1 + i
    Loop
    Player(0).SetAnimationID i: Player(0).SetAnimationLoop False: Player(0).PlayAnimation
    ������ϻ(�������) = ������ϻ(�������) - 1
    ������(1) = �޶�����������(1)
    For i = 1 To UBound(Enemy)
If EnmHP(i) < 0 Then GoTo ������Ŀ�� '������Ŀ��A
    With Enemy(i)
    If .Collision(Vector(CameraPozX, CameraPozY + PlayerHeight + CameraPozƫ��(2), CameraPozZ), Vector(CameraPozX + Cos(Math.Deg2Rad(CameraAngX - 90)) * Cos(Math.Deg2Rad(CameraAngY)) * �������(�������), CameraPozY + PlayerHeight - Sin(Math.Deg2Rad(CameraAngY)) * �������(�������), CameraPozZ - Sin(Math.Deg2Rad(CameraAngX - 90)) * Cos(Math.Deg2Rad(CameraAngY)) * �������(�������)), TV_TESTTYPE_HITBOXES) Then
      EnmHP(i) = EnmHP(i) - �����˺�(�������)
      If EnmHP(i) <= 0 Then
        Select Case EnmType(i)
        Case 0: ִ�ж��� 1, i, "death" & (1 + Round(Rnd * 2)), ��Ϸ�ٶ�, False
        Case 1: ִ�ж��� 1, i, "death" & (1 + Round(Rnd * 2)) & "_die", ��Ϸ�ٶ�, False
        Case 2: ִ�ж��� 1, i, "die_simple", ��Ϸ�ٶ�, False
        End Select
      End If
    End If
    End With
������Ŀ��: '������Ŀ��B
    Next
  End If
  End If
  ������ = ������ + Tv.TimeElapsed
Case 2 '===����ϻ===
  If ������ = 0 Then ִ�ж��� 0, 0, "reload", ��Ϸ�ٶ�, False
  If Player(0).IsAnimationFinished Then
    If ������ϻ(�������) > 0 Then
      ������ϻ(�������) = 1 + �޶�������ϻ(�������)
    Else
      ������ϻ(�������) = �޶�������ϻ(�������)
    End If
    ������ϻ��(�������) = ������ϻ��(�������) - 1
    ִ�ж��� 0, 0, "idle", ��Ϸ�ٶ�, False
    ����״̬ = 0
  End If
  ������ = ������ + Tv.TimeElapsed
End Select
'===============��������===============
If ������(1) > 0 Then ������(1) = ������(1) - 0.02 * Tv.TimeElapsed
If ������(1) < 0 Then ������(1) = ������(1) + 0.02 * Tv.TimeElapsed
If Abs(������(1)) < 0.1 Then ������(1) = 0
Player(0).SetPosition CameraPozX, CameraPozY + PlayerHeight + CameraPozƫ��(2) + ����λ��ƫ��(2), CameraPozZ
Player(0).SetRotation 0, CameraAngX + 90 + ������(0), CameraAngY - ������(1)
'�趨�����
Camera.SetPosition CameraPozX, CameraPozY + PlayerHeight + CameraPozƫ��(2), CameraPozZ
Camera.SetRotation CameraAngY - ������(1), CameraAngX + ������(0), 0
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
For i = 0 To UBound(Mesh): Mesh(i).Render: Next
For i = 0 To UBound(MeshSin): MeshSin(i).Render: Next
For i = 1 To UBound(Enemy): EnemyGun(i).Render: Next
For i = 1 To UBound(Enemy): Enemy(i).Render: Next
For i = 0 To UBound(MeshTVA): MeshTVA(i).Render: Next
Scene.FinalizeShadows '��ȾӰ��
'===============��ɫ�¼�===============
For i = 1 To UBound(Enemy)
LE.DeleteLight LE.GetLightFromName("EnmGunFire" & i)
With Enemy(i)
  If EnmHP(i) <= 0 Then GoTo �����˶���
  Select Case EnmType(i)
  Case 0: AIadv 1, i, 2, 0.008
  Case 1: AIadv 1, i, 2, 0.006
  Case 2: AIadv 1, i, 1, 0.01
  End Select
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
'===============������Ⱦ===============
If ����״̬ = 1 And ������ < 60 Then Scrͼ��.Draw_Sprite GetTex("flash" & ǹ������ & Round(Rnd)), ǹ��X * ����X, ǹ��Y * ����Y
If ������������ = False Then Player(0).Render
If Ѫ�۲���ʱ�� > 0 Or PlayerHP(0) < 20 Then Scrͼ��.Draw_SpriteScaled GetTex("Ѫ��0"), 0, 0, -1, ����X, ����Y: Ѫ�۲���ʱ�� = Ѫ�۲���ʱ�� - 1
If ������������ = False Then
Select Case ����״̬ '����׼��
Case 0
  If B2 = True Then
  Scrͼ��.Draw_Line ׼��X - 60, ׼��Y, ׼��X + 60, ׼��Y, RGBA(0, 0.8, 0, 2)
  Scrͼ��.Draw_Line ׼��X, ׼��Y, ׼��X, ׼��Y + 30, RGBA(0, 0.8, 0, 2)
  Else
  Scrͼ��.Draw_Line ׼��X - 40, ׼��Y, ׼��X - 40, ׼��Y + 30, RGBA(0, 0.8, 0, 2)
  Scrͼ��.Draw_Line ׼��X + 40, ׼��Y, ׼��X + 40, ׼��Y + 30, RGBA(0, 0.8, 0, 2)
  Scrͼ��.Draw_Line ׼��X - 100, ׼��Y, ׼��X - 40, ׼��Y, RGBA(0, 0.8, 0, 2)
  Scrͼ��.Draw_Line ׼��X + 100, ׼��Y, ׼��X + 40, ׼��Y, RGBA(0, 0.8, 0, 2)
  End If
Case 1
  If ������ > �޶�������(�������) Then
    ������ = 0
    If B1 = False Then ����״̬ = 0: ִ�ж��� 0, 0, "idle1", ��Ϸ�ٶ�, False
  End If
  Scrͼ��.Draw_Line ׼��X - 120, ׼��Y, ׼��X - 70, ׼��Y, RGBA(0, 0.8, 0, 2)
  Scrͼ��.Draw_Line ׼��X + 120, ׼��Y, ׼��X + 70, ׼��Y, RGBA(0, 0.8, 0, 2)
End Select
End If
'===============������Ⱦ===============
Scrͼ��.Draw_Sprite GetTex("UI��������"), 15, Me.Height / 15 - 135 '��������UI
Select Case PlayerHP(0)
Case Is > 60: UI������ɫ = RGBA(1, 1, 1, 1)
Case Is > 30: UI������ɫ = RGBA(1, 1, 0.5, 0.6)
Case Else: UI������ɫ = RGBA(1, 0.5, 0.5, 0.6)
End Select
If ������ϻ(�������) > �޶�������ϻ(�������) Then
  Lrc.NormalFont_DrawText "1+" & (������ϻ(�������) - 1) & "/" & �޶�������ϻ(�������), 50, Me.Height \ 15 - 100, UI������ɫ, 1
Else
  Lrc.NormalFont_DrawText ������ϻ(�������) & "/" & �޶�������ϻ(�������), 50, Me.Height \ 15 - 100, UI������ɫ, 1
End If
Lrc.NormalFont_DrawText ������(�������) & " " & ������ϻ��(�������) & "/" & �޶�������ϻ��(�������), 50, Me.Height \ 15 - 60, UI������ɫ, 1
For i = 0 To UBound(MeshTVA)
  If ����ƽ��(MeshTVA(i).GetPosition, Player(0).GetPosition) < 2500 Then Lrc.NormalFont_DrawText "��F����ȡ��ҩ", Me.Width / 30 - 105, Me.Height \ 30, RGBA(1, 1, 1, 2), 1
Next
If LRC��ʧʱ�� > 0 Then
  DrawLRC LRCname, LRCtext, LRCcolor
  LRC��ʧʱ�� = LRC��ʧʱ�� - Tv.TimeElapsed
End If
Tv.RenderToScreen
DoEvents
Loop
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Shell App.Path & "\" & App.EXEName & ".exe 0", vbNormalFocus
End
End Sub
Private Sub Tim��������_Timer()
For i = 1 To UBound(Enemy)
With Enemy(i)
  If EnmHP(i) > 0 And Enemy(i).GetPosition.Y > -500 Then GoTo �����˶���
  .SetPosition Rnd * 40 + 80, 60, Rnd * 40 + 80
  Select Case EnmType(i)
  Case 0
    ִ�ж��� 1, i, "run", ��Ϸ�ٶ�, True
    EnmHP(i) = �Ѷ� * 10
  Case 1
    ִ�ж��� 1, i, "run", ��Ϸ�ٶ�, True
    EnmHP(i) = �Ѷ� * 20
  Case 2
    ִ�ж��� 1, i, "run2", ��Ϸ�ٶ�, True
    EnmHP(i) = �Ѷ� * 15
  End Select
  EnmState(i) = 0
  Exit Sub
End With
�����˶���:
Next
End Sub
Private Sub Timer1_Timer()
T�� = 1 + T��
Select Case T��
Case 1: �Ѷ� = 4
Case 3: CreatLRC "��", "ι�����������û��", RGBA(1, 1, 0.2, 1)
Case 6: CreatLRC "��", "������ֻ�ܿ��Լ��ˣ�", RGBA(1, 1, 0.2, 1)
Case 9: CreatLRC "��ʾ", "����·�ڳ������е�ҩ��", RGBA(1, 0, 0, 1)
Case 60: �Ѷ� = 3 + �Ѷ�
Case 240: CreatLRC "�¶�", "����������һ�ţ�ȫ������Χ���λ���", RGBA(1, 0.2, 0.2, 0.8)
Case 480: CreatLRC "�¾�", "�������£��ؼ�С�ӽӽ�����", RGBA(1, 0.2, 0.2, 0.8):  Tim��������.Enabled = False
Case 500
  For i = 1 To 7 Step 3: With Enemy(i)
  EnmType(i) = 1
  .LoadTVA ("Player\�ؼ�\�ؼ�.tva") '��ȡģ��
  .SetScale 0.31, 0.31, 0.31 '�趨��С
  .SetPosition Rnd * 40 + 80, 400, Rnd * 40 + 80 '�趨ģ��λ��
  .SetAnimationByName ("walk") 'ִ�еĶ�������
  .PlayAnimation ��Ϸ�ٶ� '���Ŷ����ٶ�
  EnmState(i) = 0
  EnmHP(i) = �Ѷ� * 20
  End With: Next
  �Ѷ� = 8: Tim��������.Enabled = True
Case 720: CreatLRC "�¾�", "�������£������������λ����ش����ڳ��ˣ�", RGBA(1, 0.2, 0.2, 0.8): Tim��������.Enabled = False
Case 740
  CreatLRC "�¶�", "׼������������Ԯ", RGBA(1, 0.2, 0.2, 0.8)
  For i = 2 To 8 Step 3: With Enemy(i)
  EnmType(i) = 2
  .LoadTVA ("Player\ܽ��\ܽ��.tva") '��ȡģ��
  .SetScale 0.28, 0.28, 0.28 '�趨��С
  .SetPosition Rnd * 40 + 80, 400, Rnd * 40 + 80 '�趨ģ��λ��
  .SetAnimationByName ("run2") 'ִ�еĶ�������
  .PlayAnimation ��Ϸ�ٶ� '���Ŷ����ٶ�
  EnmState(i) = 0
  EnmHP(i) = �Ѷ� * 15
  End With: Next
  �Ѷ� = 12: Tim��������.Enabled = True
End Select
End Sub
