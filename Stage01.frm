VERSION 5.00
Begin VB.Form Stage01 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "�����ִ�ս����糺�����5  �������ݷ������"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Tim�������� 
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
Dim ��ɫλ��(0 To 3) As TV_3DVECTOR '0���12��ɫ3��ת
Dim ��ɫ����(0 To 3) As TV_3DVECTOR
Dim VoiceT As Long, T�� As Long
'������ͨ������������
Dim LRCname As String, LRCtext As String, LRCcolor As Single, LRC��ʧʱ�� As Long
Dim ������� As Long: Dim �Ѷ� As Long
Dim ����ģʽ As Boolean
Dim CLRC����(1 To 4) As String, CLRCʵ��(1 To 4) As String, CLRCt(0 To 1) As Long
Private Function CLRC(���� As String, ��� As String, ���� As String, β�� As String)
For j = 1 To 4: CLRCʵ��(j) = "": Next
CLRC����(1) = ����: CLRC����(2) = ���: CLRC����(3) = ����: CLRC����(4) = β��
CLRCt(0) = 5000
End Function
Private Function DrawCLRC(��ɫ As Single)
Dim HavePlayed As Boolean
If CLRCt(0) <= 0 Then Exit Function
CLRCt(1) = CLRCt(1) + Tv.TimeElapsed
CLRCt(0) = CLRCt(0) - Tv.TimeElapsed
If CLRCt(1) > 100 Then
For j = 1 To 4
  If Len(CLRCʵ��(j)) < Len(CLRC����(j)) Then
    CLRCʵ��(j) = Left(CLRC����(j), 1 + Len(CLRCʵ��(j)))
    If HavePlayed = False Then
    SEplay "Type0.wav", False
    HavePlayed = True
    End If
  End If
Next
CLRCt(1) = 0
End If
Lrc.Action_BeginText
Lrc.NormalFont_DrawText CLRCʵ��(1), ׼��X - Len(CLRCʵ��(1)) * 10, ׼��Y - 200, ��ɫ, 1
Lrc.NormalFont_DrawText CLRCʵ��(2), ׼��X - Len(CLRCʵ��(2)) * 10, ׼��Y - 150, ��ɫ, 1
Lrc.NormalFont_DrawText CLRCʵ��(3), ׼��X - Len(CLRCʵ��(3)) * 10, ׼��Y - 100, ��ɫ, 1
Lrc.NormalFont_DrawText CLRCʵ��(4), ׼��X - Len(CLRCʵ��(4)) * 10, ׼��Y - 50, ��ɫ, 1
Lrc.Action_EndText
End Function
Private Function CreatLRC(���� As String, ���� As String, ������ɫ As Single)
LRCname = ����: LRCtext = ����: LRCcolor = ������ɫ: LRC��ʧʱ�� = 3000
End Function
Private Function DrawLRC(���� As String, ���� As String, ������ɫ As Single)
Dim LSX As Long
LSX = ׼��X - Len(���� & ����) * 10 - 10
Lrc.NormalFont_DrawText ����, LSX, Me.Height \ 15 - 100, ������ɫ, 1
Lrc.NormalFont_DrawText ����, LSX + Len(����) * 20 + 10, Me.Height \ 15 - 100, RGBA(1, 1, 1, 1), 1
End Function
Public Function �������߶�(λ�� As TV_3DVECTOR, ԽҰ�߶� As Single, �����ٶ� As Single) As Single
Dim Result As TVCollisionResult
�������߶� = Mesh(0).GetHeight(λ��.X, λ��.z)
If �������߶� > λ��.Y + ԽҰ�߶� Then �������߶� = λ��.Y
If �������߶� < λ��.Y - �����ٶ� Then �������߶� = λ��.Y - �����ٶ�
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
    If Mesh(0).AdvancedCollide(Vector(��ʱ����(1).X + 4, ��ʱ����(1).Y, ��ʱ����(1).z), Vector(��ʱ����(1).X - 4, ��ʱ����(1).Y, ��ʱ����(1).z)).IsCollision Then Enemy(������).SetPosition ��ʱ����(1).X - ����(1), ��ʱ����(1).Y, ��ʱ����(1).z
    If Mesh(0).AdvancedCollide(Vector(��ʱ����(1).X, ��ʱ����(1).Y, ��ʱ����(1).z + 4), Vector(��ʱ����(1).X, ��ʱ����(1).Y, ��ʱ����(1).z - 4)).IsCollision Then Enemy(������).SetPosition ��ʱ����(1).X, ��ʱ����(1).Y, ��ʱ����(1).z - ����(2)
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
      ��ײ(0) = Mesh(0).AdvancedCollide(Enemy(������).GetPosition, Player(0).GetPosition).IsCollision
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
  CameraPozY = �������߶�(Vector(CameraPozX, CameraPozY, CameraPozZ), 2000, 3) + PlayerHeight
  �ӽ�������� True '����������
Case Asc("f") Or Asc("F")

End Select
End Sub

'============================��������===============================
Private Sub Form_Load()
On Error Resume Next
Randomize
'=====��������=====
BASSready Me.hWnd
SE.Init Me.hWnd
'=====��������=====
����ģʽ = False
If ����ģʽ = True Then
  Tv.SetDebugFile "TV3D������־.txt"
  Tv.SetDebugMode True, True
End If
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
Scene.SetViewFrustum 45, 0    '���ӷ�Χ�����ӽǶ�45
If ����(2) > 0 Then Scene.SetTextureFilter TV_FILTER_BILINEAR
'=====·��=====
'=====��ͼ=====
Dim BGname As String: BGname = "sunny"
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
TF.LoadTexture "Pic\Flash\flare4.jpg", "flare4"
TF.LoadTexture "Map\��̲\height.jpg", "height" '�߶�ͼ
TF.LoadTexture "Map\��̲\mud.jpg", "main"
TF.LoadAlphaTexture "Map\��̲\mask.jpg", "mask" '����ͼ
TF.LoadTexture "Map\��̲\ggrass.dds", "main2"
TF.LoadTexture "Pic\Stage\Fire_smoke.jpg", "fire", , , TV_COLORKEY_MAGENTA '���������

'=====�����Ч=====
Atmos.SkyBox_Enable True '������պ�
Atmos.SkyBox_SetTexture GetTex("SKYBOX_Front"), GetTex("SKYBOX_Back"), GetTex("SKYBOX_Left"), GetTex("SKYBOX_Right"), GetTex("SKYBOX_Up"), GetTex("SKYBOX_DOWN") '�趨��ͼ
Atmos.Fog_Enable True                              '������
Atmos.Fog_SetColor 0.1, 0.1, 0.1                         '��ɫRGBA�������
Atmos.Fog_SetParameters 500, 2000, 0              '������룬��Զ���룬Ũ��
Atmos.Fog_SetType TV_FOG_LINEAR, TV_FOGTYPE_PIXEL  '�������
'=====����=====
MF.CreateMaterial "solid" '������Ϊsolid�Ĳ���
MF.SetAmbient GetMat("solid"), 0.8, 0.8, 0.8, 1    '������
MF.SetDiffuse GetMat("solid"), 1, 1, 1, 1 '��ɢ�⣬������Ĺ�����ɫ
MF.SetEmissive GetMat("solid"), 0, 0, 0, 1  '�Է���
MF.SetOpacity GetMat("solid"), 1  '��͸����
MF.SetSpecular GetMat("solid"), 0, 0, 0, 0 '�߹�ɫ
MF.SetPower GetMat("solid"), 60 'ɢ��ǿ��

MF.CreateMaterial "map" '������ͼ�߹����
MF.SetAmbient GetMat("map"), 0.8, 0.8, 0.8, 1   '������
MF.SetDiffuse GetMat("map"), 1, 1, 1, 1 '��ɢ�⣬������Ĺ�����ɫ
MF.SetEmissive GetMat("map"), 1, 1, 1, 1 '�Է���
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
If ����(0) > 0 Then
  LE.SetSpecularLighting True '�߹⿪��
  LE.SetLightProperties 0, True, True, False '�ƹ⿪��Ӱ��
End If
'=====����=====
'����ײ���ʵ��
Set Mesh(0) = Scene.CreateLandscape
Mesh(0).GenerateTerrain "Map\��̲\height.jpg", TV_PRECISION_AVERAGE, 8, 8, -256 * 4, 0, -256 * 4, True
Mesh(0).SetTexture GetTex("main")
Mesh(0).SetMaterial GetMat("solid")
Mesh(0).SetTextureScale 4, 4
Mesh(0).SetLightingMode TV_LIGHTING_NORMAL
Mesh(0).SetPosition 0, 0, 0

'��������ͼTVAʵ��
For i = 0 To UBound(MeshTVA): Set MeshTVA(i) = Scene.CreateActor: Next
MeshTVA(0).LoadTVA "Model\M1̹��\M1̹��.tva", True, True
MeshTVA(1).LoadTVA "Model\����\����.tva", True, True
MeshTVA(2).LoadTVA "Model\M1̹��\M1̹��.tva", True, True
With MeshTVA(0)
  .SetScale 0.33, 0.33, 0.33
  .SetPosition 1650, �������߶�(Vector(1650, 0, 1200), 999, 999) + 10, 1200
  .RotateY -120
  .SetMaterial GetMat("map")
End With
With MeshTVA(1)
  .SetScale 0.33, 0.33, 0.33
  .SetPosition 1400, �������߶�(Vector(1400, 0, 1500), 999, 999) + 12, 1500
  .SetRotation -90, -120, 0
  .SetMaterial GetMat("solid")
End With
With MeshTVA(2)
  .SetScale 0.33, 0.33, 0.33
  .SetPosition 1200, �������߶�(Vector(1200, 0, 1300), 999, 999) + 8, 1300
  .SetRotation 120, -120, 0
  .SetMaterial GetMat("solid")
End With
For i = 0 To UBound(MeshTVA): With MeshTVA(i) '��Ӱ
  .SetLightingMode TV_LIGHTING_NORMAL
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
.LoadTVA ("Player\�ؼ�\�ؼ�.tva") '��ȡģ��
.SetMaterial GetMat("solid") '�趨����
If ����(0) > 0 Then .SetLightingMode TV_LIGHTING_NORMAL '�趨����ģʽ
.SetScale 0.28, 0.28, 0.28 '�趨��С
.SetPosition Rnd * 300 + 1200, 0, Rnd * 300 + 1200 '�趨ģ��λ��
.SetPosition .GetPosition.X, �������߶�(Vector(.GetPosition.X, .GetPosition.Y, .GetPosition.z), 999, 999) + 8, .GetPosition.z '�趨ģ��λ��
'.SetRotation -90, 0, 180
End With
ִ�ж��� 1, i, "death" & (1 + Round(Rnd * 2)) & "_die", 9, False
EnmHP(i) = 0
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
����ˮ�� 800, 1650, 2000, 2000, 72
If ����(3) >= 50 Then
Set tPart = Scene.CreateParticleSystem '����
tPart.Load "Part\����1\����1.tvp"
tPart.SetGlobalPosition MeshTVA(2).GetPosition.X, MeshTVA(2).GetPosition.Y, MeshTVA(2).GetPosition.z
tPart.SetGlobalScale 0.05, 0.05, 0.05
Set tPart = Scene.CreateParticleSystem '����
tPart.Load "Part\����1\����1.tvp"
tPart.SetGlobalPosition MeshTVA(1).GetPosition.X - 1, MeshTVA(1).GetPosition.Y - 3, MeshTVA(1).GetPosition.z - 4
tPart.SetGlobalScale 4, 4, 4
End If
'=====����=====
Lrc.NormalFont_Create "", "����", 25, False, False, False
��ʼ���ӽǲ��� 1380, �������߶�(Vector(1380, 0, 1680), 999, 999), 1680, 180, 99
PlayerHeight = 2: ��Ϸ�ٶ� = 1
GunLoad "AKS-74U", 1, True
GunLoad "M16", 0, True
BASSplay "Audio\BGM\music_ingame_disaster_and_rescue.mp3", 0, 1, Me.hWnd
BASSplay "Audio\BGS\Battle-far1.mp3", 1, 1, Me.hWnd
������������ = True
Do Until ������� > 0
  Tv.Clear
  DrawCLRC (RGBA(1, 1, 1, 1)): DrawLRC LRCname, LRCtext, LRCcolor
  Tv.RenderToScreen
  DoEvents
Loop
'================================��ѭ��=======================================
Do Until ������� > 5
 VoiceT = VoiceT + Tv.TimeElapsed
 ��������(3) = 0: ��������(4) = 0
 If PlayerHP(0) = 0 Then GoTo �ƶ��������
 Inp.GetMouseState Mx, My, B1, B2, , , Roll           '���������Ϣ
 CameraAngX = CameraAngX + 0.08 * Mx
 CameraAngY = CameraAngY + 0.08 * My
If CameraAngY > 50 Then CameraAngY = 50
If CameraAngY < -60 Then CameraAngY = -60
�����ƶ� = False
    If Inp.IsKeyPressed(TV_KEY_W) Then 'ǰ
      �����ƶ� = True
      ��������(3) = ��������(3) + Cos(Math.Deg2Rad(CameraAngX - 90)) * 0.002 * Tv.TimeElapsed * ��Ϸ�ٶ�
      ��������(4) = ��������(4) - Sin(Math.Deg2Rad(CameraAngX - 90)) * 0.002 * Tv.TimeElapsed * ��Ϸ�ٶ�
    End If
    If Inp.IsKeyPressed(TV_KEY_S) Then '��
     ��������(3) = ��������(3) + Cos(Math.Deg2Rad(CameraAngX + 90)) * 0.001 * Tv.TimeElapsed * ��Ϸ�ٶ�
     ��������(4) = ��������(4) - Sin(Math.Deg2Rad(CameraAngX + 90)) * 0.001 * Tv.TimeElapsed * ��Ϸ�ٶ�
     �����ƶ� = True
    End If
    If Inp.IsKeyPressed(TV_KEY_A) Then '��
     ��������(3) = ��������(3) + Cos(Math.Deg2Rad(CameraAngX + 180)) * 0.002 * Tv.TimeElapsed * ��Ϸ�ٶ�
     ��������(4) = ��������(4) - Sin(Math.Deg2Rad(CameraAngX + 180)) * 0.002 * Tv.TimeElapsed * ��Ϸ�ٶ�
     �����ƶ� = True
    End If
    If Inp.IsKeyPressed(TV_KEY_D) Then '��
     ��������(3) = ��������(3) + Cos(Math.Deg2Rad(CameraAngX)) * 0.002 * Tv.TimeElapsed * ��Ϸ�ٶ�
     ��������(4) = ��������(4) - Sin(Math.Deg2Rad(CameraAngX)) * 0.002 * Tv.TimeElapsed * ��Ϸ�ٶ�
     �����ƶ� = True
    End If
�ƶ��������:
    If �����ƶ� = True And Inp.IsKeyPressed(TV_KEY_LEFTSHIFT) = False Then
      CameraPozƫ��(0) = 0.1 + CameraPozƫ��(0)
      CameraPozƫ��(2) = Sin(CameraPozƫ��(0)) * ��Ϸ�ٶ� / 8
    End If
'=====������ײ=====
If �����ƶ� = False Then GoTo ��������ƶ�
Dim ��ײ��ʱ As Boolean
CameraPozX = CameraPozX + ��������(3)
CameraPozZ = CameraPozZ + ��������(4)
CameraPozY = �������߶�(Vector(CameraPozX, CameraPozY, CameraPozZ), 20, 3)
For i = 0 To UBound(MeshTVA)
  If MeshTVA(i).Collision(Vector(CameraPozX - 2, CameraPozY + PlayerHeight, CameraPozZ), Vector(CameraPozX + 2, CameraPozY + PlayerHeight, CameraPozZ)) Or _
  MeshTVA(i).Collision(Vector(CameraPozX, CameraPozY + PlayerHeight, CameraPozZ - 2), Vector(CameraPozX, CameraPozY + PlayerHeight, CameraPozZ + 2)) Then
  �ӽ�������� False: GoTo ��������ƶ�
  End If
Next
If CameraPozZ > 1680 Then �ӽ�������� False: GoTo ��������ƶ�
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
Atmos.Fog_Enable False
ReflectRS.StartRender '���䲿����Ⱦ
 Atmos.Atmosphere_Render
ReflectRS.EndRender

RefractRS.StartRender '���䲿����Ⱦ
 Atmos.Atmosphere_Render
RefractRS.EndRender

Tv.Clear '����
Atmos.Atmosphere_Render '��Ⱦ����
Atmos.Fog_Enable True
For i = 1 To UBound(Enemy): EnemyGun(i).Render: Next
For i = 1 To UBound(Enemy): Enemy(i).Render: Next
For i = 0 To UBound(MeshTVA): MeshTVA(i).Render: Next
Scene.RenderAllMeshes
Mesh(0).Render
Floor.Render
Scene.FinalizeShadows '��ȾӰ��
Scene.RenderAllParticleSystems
'===============��ɫ�¼�===============
For i = 1 To UBound(Enemy)
'LE.DeleteLight LE.GetLightFromName("EnmGunFire" & i)
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
Select Case �������
Case 2: If ����ƽ��(Player(0).GetPosition, MeshTVA(1).GetPosition) < 2000 Then ������� = 3 '���ź���վ����
Case 3: If PlayerHeight < 8 Then PlayerHeight = 0.2 + PlayerHeight Else ��Ϸ�ٶ� = 2
Case 4
  If PlayerHeight > 2 Then
    PlayerHeight = PlayerHeight - 0.2
  Else
    ������� = 5
    BASSset 0, 0: BASSset 1, 0: BASSset 2, 0
  End If
End Select
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
If LRC��ʧʱ�� > 0 Then
  DrawLRC LRCname, LRCtext, LRCcolor
  LRC��ʧʱ�� = LRC��ʧʱ�� - Tv.TimeElapsed
End If

DrawCLRC RGBA(0.2, 1, 0.2, 1)
Tv.RenderToScreen
DoEvents
Loop
Do Until ������� > 6
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
Case 1: If ����ģʽ Then ������� = 1
Case 6: SoundMp3.Load "Audio\SE\Type0.wav": CLRC "", "", "", "2XXX��9��18�� ��������"
Case 10: CLRCt(0) = 0
Case 11: ������� = 1: If ����ģʽ = False Then GF.FadeIn 2000
Case 14: CreatLRC "W��A��S��D������ ���ź���վ��", "", RGBA(1, 1, 0, 1)
Case 17: CLRC "���߻���", "����_���k��ξ", "ľ������", "���ҵĽ�����"
Case 22: CLRC "��������", "����_���k��ξ", "           Drzzm32(��ֲ)", ""
Case 27: CLRC "��������", "ľ������", "����_���k��ξ", ""
Case 32: CLRC "�����¹��ʡ�", "            wen832238", "555ʮ���ſӰ�", "����Ͻ�"
Case 37: CLRC "�������ල��", "ľ������", "������", "          Archeb"
Case 42: CLRC "�������Ŷӡ�", "���ݷ������", "�����ٷ�������", "        DeseCity�����ң��ѿӣ�"
Case 47: CLRC "��Powered By��", "           TrueVision 3D Engine", "        Visual Basic 6", "           BASS.dll Sound Engine"
Case 52: ������� = 2
Case 120: GF.Flash 1, 1, 1, 8000
Case 150: GF.Flash 1, 1, 1, 8000
Case 180: GF.Flash 1, 1, 1, 8000: ������� = 4: GF.FadeOut 3000
Case 183
  CLRC "", "", "�����ִ�ս��", "�������": CreatLRC "", "Touhou Modern War:Dreaming Fallen Of Moon", RGBA(1, 1, 1, 1)
  ������� = 6
Case 191: Unload Me
End Select
End Sub

