Attribute VB_Name = "BGM����"
Public ����״̬(0 To 5) As Long     ' 0ֹͣ1����2��ͣ
Public chan(0 To 9) As Long         ' Channel Handle
Dim sldEQR(0 To 5) As Long
Dim fx(0 To 5) As Long        ' 3 EQ + 3D
'=========================BASS����===============================
Public Function BASSplay(��Ƶ���·�� As String, ���� As Long, �������� As Long, hWnd As Long)
On Error Resume Next
Dim ��Ƶ·�� As String
��Ƶ·�� = App.Path & "\" & ��Ƶ���·��
If Dir(��Ƶ·��) = "" Then MsgBox "��Ƶ�ļ�" & ��Ƶ·�� & vbCrLf & "�����ڣ�": Exit Function
If ����(6) = 0 Then '===WMP����===
  BG.BGM(����).url = ��Ƶ·��
  Exit Function
End If
If ����״̬(����) > 0 Then BASSset ����, 0
Select Case ��������
Case 0 'ͨ������
 sldEQR(0) = 10000
 sldEQR(3) = 18000
 sldEQR(1) = 20000
 sldEQR(4) = 7
 sldEQR(2) = 6
 fx(0) = 10000
 fx(3) = 18000
 fx(1) = 20000
 fx(4) = 7
 fx(2) = 6
Case 1 '�ص���
 sldEQR(0) = 6
 sldEQR(3) = 9
 sldEQR(1) = 27000
 sldEQR(4) = 8
 sldEQR(2) = 6
 fx(0) = 6
 fx(3) = 9
 fx(1) = 27000
 fx(4) = 8
 fx(2) = 6
End Select
For i = LBound(fx) To UBound(fx)
 Call UpdateFX(i)
Next
'=====================================================================================
    ' free both MOD and stream, it must be one of them! :)
If ����״̬(����) <> 0 Then
    Call BASS_MusicFree(chan(����))
    Call BASS_StreamFree(chan(����))
End If
'If 1 Then ' with FX flag
    chan(����) = BASS_StreamCreateFile(BASSFALSE, StrPtr(��Ƶ·��), 0, 0, BASS_SAMPLE_LOOP Or BASS_SAMPLE_FX)
    If (chan(����) = 0) Then chan(����) = BASS_MusicLoad(BASSFALSE, StrPtr(��Ƶ·��), 0, 0, BASS_MUSIC_LOOP Or BASS_MUSIC_RAMP Or BASS_SAMPLE_FX, 1)
'Else   ' without FX flag
    'chan(����) = BASS_StreamCreateFile(BASSFALSE, StrPtr(��Ƶ·��), 0, 0, BASS_SAMPLE_LOOP)
    'If (chan(����) = 0) Then chan(����) = BASS_MusicLoad(BASSFALSE, StrPtr(��Ƶ·��), 0, 0, BASS_MUSIC_LOOP Or BASS_MUSIC_RAMP, 1)
'End If
    If (chan(����) = 0) Then  '�޷�����
        Call Error_("�޷������ļ���" & ��Ƶ·��)
        Exit Function
    End If
    '��������Ч
    Dim P As BASS_DX8_PARAMEQ
    fx(0) = BASS_ChannelSetFX(chan(����), BASS_FX_DX8_PARAMEQ, 0) '����
    fx(3) = BASS_ChannelSetFX(chan(����), BASS_FX_DX8_PARAMEQ, 0) '�����ν�
    fx(1) = BASS_ChannelSetFX(chan(����), BASS_FX_DX8_PARAMEQ, 0) '����
    fx(4) = BASS_ChannelSetFX(chan(����), BASS_FX_DX8_PARAMEQ, 0) '�����ν�
    fx(2) = BASS_ChannelSetFX(chan(����), BASS_FX_DX8_PARAMEQ, 0) '����
    'fx(5) = BASS_chan(����)nelSetFX(chan(����), BASS_FX_DX8_REVERB, 0)  '3D����

    P.fGain = 0
    P.fBandwidth = 18

    P.fCenter = 80 '����
    Call BASS_FXSetParameters(fx(0), P)
    P.fCenter = 150 '�����ν�
    Call BASS_FXSetParameters(fx(3), P)
    P.fCenter = 910 '����
    Call BASS_FXSetParameters(fx(1), P)
    P.fCenter = 3000 '�����ν�
    Call BASS_FXSetParameters(fx(4), P)
    P.fCenter = 14000 '����
    Call BASS_FXSetParameters(fx(2), P)
    ' you can add more EQ bands with chan(����)ging:
    ' p.fCenter = N [Hz] N>=80 and N<=16000
    Call UpdateFX(0) '����
    Call UpdateFX(3)
    Call UpdateFX(1) '����
    Call UpdateFX(4)
    Call UpdateFX(2) '����
    'Call UpdateFX(5) '3D
    Call BASS_ChannelPlay(chan(����), BASSTRUE)
    ����״̬(����) = 1
End Function
Public Function BASSready(hWnd As Long)
On Error Resume Next
   If ����(6) = 0 Then Exit Function
   Call BASS_Free
   ' change and set the current path, to prevent from VB not finding BASS.DLL
    ChDrive App.Path
    ChDir App.Path
    ' check the correct BASS was loaded
    If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
        Call MsgBox("BASS.dll�汾����", vbCritical)
        End
    End If
    ' setup output - default device
    If (BASS_Init(-1, 44100, 0, hWnd, 0) = 0) Then
        Call Error_("�޷���ʼ������")
        End
    End If
    ' check that DX8 features are available
    Dim bi As BASS_INFO
    Call BASS_GetInfo(bi)
    If (bi.dsver < 8) Then
        Call BASS_Free
        Call Error_("���ĵ���δ��װDirectX 8")
        End
    End If
End Function
' display error messages
Public Sub Error_(ByVal es As String)
  Call MsgBox(es & vbCrLf & vbCrLf & "�Ҳ�����Ƶ�ļ� ������룺" & BASS_ErrorGetCode, vbExclamation, "����")
  If MsgBox("�Ƿ���ʱ�л����ݲ��ź��ģ��������������������ø��ģ�", vbYesNo) = vbYes Then ����(6) = 0
End Sub
' get file name from file path
Public Function GetFileName(ByVal fp As String) As String
  GetFileName = Mid(fp, InStrRev(fp, "\") + 1)
End Function
Public Sub UpdateFX(ByVal b As Integer)
On Error Resume Next
    Dim v As Integer
    v = sldEQR(b)
    If (b < 5) Then
        Dim P As BASS_DX8_PARAMEQ
        Call BASS_FXGetParameters(fx(b), P)
        P.fGain = 10# - v
        Call BASS_FXSetParameters(fx(b), P)
    Else
        Dim p1 As BASS_DX8_REVERB
        Call BASS_FXGetParameters(fx(b), p1)
        p1.fReverbMix = -0.012 * v * v * v
        Call BASS_FXSetParameters(fx(b), p1)
    End If
End Sub
Public Function BASSset(���� As Long, ״̬ As Long)
On Error Resume Next
If ����(6) = 0 Then '===WMP����===
  Select Case ״̬
  Case 0 '��Ϊ��ʼ��
    BG.BGM(����).Controls.stop
    ����״̬(����) = 0
  Case 1 '��Ϊ����
    BG.BGM(����).Controls.Play
    ����״̬(����) = 1
  Case 2 '��Ϊ��ͣ
    BG.BGM(����).Controls.Pause
    ����״̬(����) = 2
  End Select
  Exit Function
End If
Select Case ״̬
Case 0 '��Ϊ��ʼ��
  Call BASS_ChannelStop(chan(����))
  ����״̬(����) = 0
Case 1 '��Ϊ����
  Call BASS_ChannelPlay(chan(����), BASSFALSE)
  ����״̬(����) = 1
Case 2 '��Ϊ��ͣ
  Call BASS_ChannelPause(chan(����))
  ����״̬(����) = 2
End Select
End Function
