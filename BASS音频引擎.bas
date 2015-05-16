Attribute VB_Name = "BGM引擎"
Public 播放状态(0 To 5) As Long     ' 0停止1播放2暂停
Public chan(0 To 9) As Long         ' Channel Handle
Dim sldEQR(0 To 5) As Long
Dim fx(0 To 5) As Long        ' 3 EQ + 3D
'=========================BASS核心===============================
Public Function BASSplay(音频相对路径 As String, 音轨 As Long, 均衡类型 As Long, hWnd As Long)
On Error Resume Next
Dim 音频路径 As String
音频路径 = App.Path & "\" & 音频相对路径
If Dir(音频路径) = "" Then MsgBox "音频文件" & 音频路径 & vbCrLf & "不存在！": Exit Function
If 设置(6) = 0 Then '===WMP核心===
  BG.BGM(音轨).url = 音频路径
  Exit Function
End If
If 播放状态(音轨) > 0 Then BASSset 音轨, 0
Select Case 均衡类型
Case 0 '通用丽音
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
Case 1 '重低音
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
If 播放状态(音轨) <> 0 Then
    Call BASS_MusicFree(chan(音轨))
    Call BASS_StreamFree(chan(音轨))
End If
'If 1 Then ' with FX flag
    chan(音轨) = BASS_StreamCreateFile(BASSFALSE, StrPtr(音频路径), 0, 0, BASS_SAMPLE_LOOP Or BASS_SAMPLE_FX)
    If (chan(音轨) = 0) Then chan(音轨) = BASS_MusicLoad(BASSFALSE, StrPtr(音频路径), 0, 0, BASS_MUSIC_LOOP Or BASS_MUSIC_RAMP Or BASS_SAMPLE_FX, 1)
'Else   ' without FX flag
    'chan(音轨) = BASS_StreamCreateFile(BASSFALSE, StrPtr(音频路径), 0, 0, BASS_SAMPLE_LOOP)
    'If (chan(音轨) = 0) Then chan(音轨) = BASS_MusicLoad(BASSFALSE, StrPtr(音频路径), 0, 0, BASS_MUSIC_LOOP Or BASS_MUSIC_RAMP, 1)
'End If
    If (chan(音轨) = 0) Then  '无法播放
        Call Error_("无法播放文件：" & 音频路径)
        Exit Function
    End If
    '均衡器生效
    Dim P As BASS_DX8_PARAMEQ
    fx(0) = BASS_ChannelSetFX(chan(音轨), BASS_FX_DX8_PARAMEQ, 0) '低音
    fx(3) = BASS_ChannelSetFX(chan(音轨), BASS_FX_DX8_PARAMEQ, 0) '低音衔接
    fx(1) = BASS_ChannelSetFX(chan(音轨), BASS_FX_DX8_PARAMEQ, 0) '主唱
    fx(4) = BASS_ChannelSetFX(chan(音轨), BASS_FX_DX8_PARAMEQ, 0) '高音衔接
    fx(2) = BASS_ChannelSetFX(chan(音轨), BASS_FX_DX8_PARAMEQ, 0) '高音
    'fx(5) = BASS_chan(音轨)nelSetFX(chan(音轨), BASS_FX_DX8_REVERB, 0)  '3D混响

    P.fGain = 0
    P.fBandwidth = 18

    P.fCenter = 80 '低音
    Call BASS_FXSetParameters(fx(0), P)
    P.fCenter = 150 '低音衔接
    Call BASS_FXSetParameters(fx(3), P)
    P.fCenter = 910 '主唱
    Call BASS_FXSetParameters(fx(1), P)
    P.fCenter = 3000 '高音衔接
    Call BASS_FXSetParameters(fx(4), P)
    P.fCenter = 14000 '高音
    Call BASS_FXSetParameters(fx(2), P)
    ' you can add more EQ bands with chan(音轨)ging:
    ' p.fCenter = N [Hz] N>=80 and N<=16000
    Call UpdateFX(0) '低音
    Call UpdateFX(3)
    Call UpdateFX(1) '主唱
    Call UpdateFX(4)
    Call UpdateFX(2) '高音
    'Call UpdateFX(5) '3D
    Call BASS_ChannelPlay(chan(音轨), BASSTRUE)
    播放状态(音轨) = 1
End Function
Public Function BASSready(hWnd As Long)
On Error Resume Next
   If 设置(6) = 0 Then Exit Function
   Call BASS_Free
   ' change and set the current path, to prevent from VB not finding BASS.DLL
    ChDrive App.Path
    ChDir App.Path
    ' check the correct BASS was loaded
    If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
        Call MsgBox("BASS.dll版本错误", vbCritical)
        End
    End If
    ' setup output - default device
    If (BASS_Init(-1, 44100, 0, hWnd, 0) = 0) Then
        Call Error_("无法初始化服务")
        End
    End If
    ' check that DX8 features are available
    Dim bi As BASS_INFO
    Call BASS_GetInfo(bi)
    If (bi.dsver < 8) Then
        Call BASS_Free
        Call Error_("您的电脑未安装DirectX 8")
        End
    End If
End Function
' display error messages
Public Sub Error_(ByVal es As String)
  Call MsgBox(es & vbCrLf & vbCrLf & "找不到音频文件 错误代码：" & BASS_ErrorGetCode, vbExclamation, "错误")
  If MsgBox("是否临时切换兼容播放核心？（您可以在设置中永久更改）", vbYesNo) = vbYes Then 设置(6) = 0
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
Public Function BASSset(音轨 As Long, 状态 As Long)
On Error Resume Next
If 设置(6) = 0 Then '===WMP核心===
  Select Case 状态
  Case 0 '设为初始化
    BG.BGM(音轨).Controls.stop
    播放状态(音轨) = 0
  Case 1 '设为播放
    BG.BGM(音轨).Controls.Play
    播放状态(音轨) = 1
  Case 2 '设为暂停
    BG.BGM(音轨).Controls.Pause
    播放状态(音轨) = 2
  End Select
  Exit Function
End If
Select Case 状态
Case 0 '设为初始化
  Call BASS_ChannelStop(chan(音轨))
  播放状态(音轨) = 0
Case 1 '设为播放
  Call BASS_ChannelPlay(chan(音轨), BASSFALSE)
  播放状态(音轨) = 1
Case 2 '设为暂停
  Call BASS_ChannelPause(chan(音轨))
  播放状态(音轨) = 2
End Select
End Function
