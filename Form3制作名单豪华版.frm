VERSION 5.00
Begin VB.Form Form3�������������� 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "��������������"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form3��������������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mesh(0 To 10) As TVMesh, MeshVisible(0 To 99) As Boolean
Dim Enemy(0 To 9) As TVActor, EnmVisible(0 To 99) As Boolean
'������ͨ������������
Dim T�� As Long
Dim LRCname As String, LRCtext As String, LRCcolor As Single, LRC��ʧʱ�� As Long
Dim ������� As Long: Dim �Ѷ� As Long
Dim ����ģʽ As Boolean
Dim CLRC����(1 To 5) As String, CLRCʵ��(1 To 5) As String
Dim CLRCx As Long, CLRCy As Long, CLRCcolor As Single
Private Function CameraMoveTo(X As Long, Y As Long, z As Long, AngX As Long, AngY As Long, Speed As Single)
CameraPozX = CameraPozX + Speed * (X - CameraPozX)
CameraPozY = CameraPozY + Speed * (Y - CameraPozY)
CameraPozZ = CameraPozZ + Speed * (z - CameraPozZ)
CameraAngX = CameraAngX Mod 360: CameraAngY = CameraAngY Mod 360
CameraAngX = CameraAngX + Speed * (AngX - CameraAngX)
CameraAngY = CameraAngY + Speed * (AngY - CameraAngY)
End Function
Private Function CLRC(X As Long, Y As Long, ���� As String, ��� As String, ���� As String, β�� As String, ĩ�� As String)
For j = 1 To 5: CLRCʵ��(j) = "": Next
CLRCx = X: CLRCy = Y
CLRC����(1) = ����: CLRC����(2) = ���: CLRC����(3) = ����: CLRC����(4) = β��: CLRC����(5) = ĩ��
End Function
Private Function DrawCLRC(��ɫ As Single)
Dim HavePlayed As Boolean
For j = 1 To 5
  If Len(CLRCʵ��(j)) < Len(CLRC����(j)) Then
    CLRCʵ��(j) = Left(CLRC����(j), 1 + Len(CLRCʵ��(j)))
    If HavePlayed = False Then
    HavePlayed = True
    End If
  End If
Next
Lrc.Action_BeginText
Lrc.NormalFont_DrawText CLRCʵ��(1), CLRCx - Len(CLRCʵ��(1)) * 14, CLRCy, ��ɫ, 1
Lrc.NormalFont_DrawText CLRCʵ��(2), CLRCx - Len(CLRCʵ��(2)) * 14, CLRCy + 50, ��ɫ, 1
Lrc.NormalFont_DrawText CLRCʵ��(3), CLRCx - Len(CLRCʵ��(3)) * 14, CLRCy + 100, ��ɫ, 1
Lrc.NormalFont_DrawText CLRCʵ��(4), CLRCx - Len(CLRCʵ��(4)) * 14, CLRCy + 150, ��ɫ, 1
Lrc.NormalFont_DrawText CLRCʵ��(5), CLRCx - Len(CLRCʵ��(5)) * 14, CLRCy + 200, ��ɫ, 1
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

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 27 'ESC�˳�
  Unload Me
Case Asc("f") Or Asc("F")
  Open App.Path & "\ʱ����.ini" For Append As #1
    Print #1, T��
  Close #1
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
����ģʽ = True
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
'=====·��=====
'=====��ͼ=====
TF.LoadTexture "Pic\Flash\white.jpg", "w" '��պ�
TF.LoadTexture "Pic\Flash\black.jpg", "b"
TF.LoadTexture "Pic\Flash\flash00.png", "flash00", , , TV_COLORKEY_USE_ALPHA_CHANNEL 'ǹ��
TF.LoadTexture "Pic\Flash\flash01.png", "flash01", , , TV_COLORKEY_USE_ALPHA_CHANNEL
TF.LoadTexture "Pic\Flash\black.jpg", "height" '�߶�ͼ

'=====�����Ч=====
Atmos.SkyBox_Enable True '������պ�
Atmos.SkyBox_SetTexture GetTex("w"), GetTex("w"), GetTex("w"), GetTex("w"), GetTex("w"), GetTex("w") '�趨��ͼ
Atmos.Fog_Enable False
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
'��Ӱ
LE.CreateDirectionalLight Vector(1, -1, 1), 1, 1, 1, "sun", 1 '���һ��ƽ�й�
If ����(0) > 0 Then
  LE.SetSpecularLighting True '�߹⿪��
  LE.SetLightProperties 0, True, False, False '�ƹ⿪��Ӱ��
End If
'=====��ɫ=====
For i = 0 To UBound(Mesh): Set Mesh(i) = Scene.CreateMeshBuilder: Next
With Mesh(0)
.LoadTVM "Model\ZBD05\ZBD05.tvm", True, True
.SetScale 1, 1, 1
.SetPosition 10, 2, 15
.SetRotation 0, 180, 0
.SetLightingMode TV_LIGHTING_NORMAL
.SetTexture GetTex("b")
End With

For i = 0 To UBound(Enemy): Set Enemy(i) = Scene.CreateActor: Next '��ɫ��ʼ��
With Enemy(0)
.LoadTVA "Model\M1̹��\M1̹��.tva", True, True
.SetScale 0.03, 0.03, 0.03
.SetPosition 1, -5, 5
.SetRotation 0, -55, 0
End With
EnmVisible(0) = True

For i = 2 To 3
With Enemy(i)
.LoadTVA "Model\��ӥ\��ӥ.tva", True, True
.SetScale 0.02, 0.02, 0.02
.SetRotation 0, 90, 5
End With
ִ�ж��� 1, i, "idle", 1, True
Next
Enemy(2).SetPosition 16, 6, 5
Enemy(3).SetPosition 12, 2, 9

For i = 0 To UBound(Enemy)
With Enemy(i)
.SetLightingMode TV_LIGHTING_NORMAL
.SetTexture GetTex("b")
End With
Next
'=====��Ч====
'=====����=====
Lrc.NormalFont_Create "", "����", 35, False, False, False
��ʼ���ӽǲ��� 0, 0, 0, 0, 9999
PlayerHeight = 2: ��Ϸ�ٶ� = 1
BG.BGM(0).Controls.stop
BG.BGM(8).url = App.Path & "\Audio\BGM\Senya-Bad Apple.mp3"
BG.BGM(8).Controls.currentPosition = 0
CLRCcolor = RGBA(0, 0, 0, 1)
'================================��ѭ��=======================================
Do
 Inp.GetMouseState Mx, My, B1, B2, , , Roll           '���������Ϣ
 If ����ģʽ = True Then
   CameraAngX = CameraAngX + 0.1 * Mx * ��Ϸ�ٶ�
   CameraAngY = CameraAngY + 0.1 * My * ��Ϸ�ٶ�
 End If
�ӽ�������� True
'�趨�����
Camera.SetPosition CameraPozX, CameraPozY + PlayerHeight + CameraPozƫ��(2), CameraPozZ
Camera.SetRotation CameraAngY - ������(1), CameraAngX + ������(0), 0
'===============��������Ⱦ===============
Tv.Clear '����
Atmos.Fog_Enable False
Atmos.Atmosphere_Render '��Ⱦ����
Atmos.Fog_Enable True
For i = 0 To UBound(Enemy)
  If EnmVisible(i) = True Then Enemy(i).Render
Next
For i = 0 To UBound(Mesh)
  If MeshVisible(i) = True Then Mesh(i).Render
Next
Scene.FinalizeShadows '��ȾӰ��
'===============��ɫ�¼�===============

�����˶���:
Select Case �������
End Select
'===============������Ⱦ===============
Lrc.NormalFont_DrawText T��, 10, 10, RGBA(1, 0, 0, 1), 1
If LRC��ʧʱ�� > 0 Then
  DrawLRC LRCname, LRCtext, LRCcolor
  LRC��ʧʱ�� = LRC��ʧʱ�� - Tv.TimeElapsed
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
T�� = 10 * BG.BGM(8).Controls.currentPosition
Select Case T��
Case 22: CLRC Me.Width \ 30, Me.Height \ 30 - 150, "-PV����������-", "����_���k��ξ", "", "", ""
Case 90
  CLRC Me.Width \ 30, Me.Height \ 30 - 150, "-PV����������-", "����_���k��ξ", "    ����Bad Apple�¾���", "-BGM-", "�ı��ǹ⣨Senya��"
  For j = 1 To 2
    CLRCʵ��(j) = CLRC����(j)
  Next
Case 165:: GF.Flash 0, 0, 0, 500
Case 173:: GF.Flash 0, 0, 0, 500
Case 178:: GF.Flash 0, 0, 0, 500
Case 165: GF.Flash 0, 0, 0, 500
Case 200: GF.Flash 0, 0, 0, 500
Case 235: GF.Flash 0, 0, 0, 500
Case 270: GF.Flash 0, 0, 0, 500
Case 282: GF.Flash 0, 0, 0, 500
Case 298: CLRC -9999, 0, "", "", "", "", "": ������� = 1
Case 368: ������� = 3
Case 430: CLRC -9999, 0, "", "", "", "", "": ������� = 4
Case 440: CLRC Me.Width \ 30 + 300, 80, "-������-", "����_���k��ξ", "ľ������", "        Drzzm32", "       Reity": EnmVisible(2) = True: EnmVisible(3) = True
Case 510: ������� = 5: EnmVisible(0) = False: CLRC -9999, 0, "", "", "", "", ""
Case 580: ������� = 6: CLRC Me.Width \ 30, Me.Height \ 30 - 150, "  -3D����-", "ľ������", "  -2D����-", "����_���k��ξ", "ľ������"
Case 718: GF.Flash 0, 0, 0, 500
Case 788: GF.Flash 0, 0, 0, 500
Case 848: GF.Flash 0, 0, 0, 500
Case 855: ������� = 7: GF.Flash 0, 0, 0, 500: CLRC -9999, 0, "", "", "", "", "": MeshVisible(0) = True: EnmVisible(2) = False: EnmVisible(3) = False
Case 870: ������� = 8
Case 3780: Unload Me
End Select
'����������������
Select Case �������
Case 1 '̹�˸���
  If Enemy(0).GetPosition.Y < 2 Then
    Enemy(0).SetPosition Enemy(0).GetPosition.X, Enemy(0).GetPosition.Y + 0.3, Enemy(0).GetPosition.z
  Else
    CLRC Me.Width \ 30 - 300, Me.Width \ 30 - 200, "-�߻�-", "����_���k��ξ", "ľ������", "���ҵĽ�����", ""
    ������� = 2
  End If
Case 2: Enemy(0).SetPosition 0.01 + Enemy(0).GetPosition.X, Enemy(0).GetPosition.Y, Enemy(0).GetPosition.z
Case 3: CameraMoveTo 0, 3, 10, 180, 40, 0.019
Case 4: Enemy(0).SetPosition Enemy(0).GetPosition.X - 0.05, Enemy(0).GetPosition.Y, Enemy(0).GetPosition.z - 0.048
Case 5: CameraMoveTo 0, 3, 10, 85, 0, 0.08
Case 6 'ֱ��������
  Enemy(2).SetPosition Enemy(2).GetPosition.X, Enemy(2).GetPosition.Y, Enemy(2).GetPosition.z + 0.2
  Enemy(3).SetPosition Enemy(3).GetPosition.X, Enemy(3).GetPosition.Y, Enemy(3).GetPosition.z + 0.2
Case 7 'װ�׳�ʻ��
  Mesh(0).SetPosition Mesh(0).GetPosition.X, Mesh(0).GetPosition.Y, Mesh(0).GetPosition.z - 0.02
Case 8: Mesh(0).RotateY 0.01
End Select
End Sub

