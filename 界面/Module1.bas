Attribute VB_Name = "ooooooo前台变量方法"
Public 设置(0 To 6) As Long
Public Function 读取设置()
On Error Resume Next
Dim NR As String
Open App.Path & "\Save\Options.ini" For Input As #2
For i = 0 To UBound(设置)
  Line Input #2, NR
  设置(i) = Val(NR)
Next
Close #2
End Function
Public Function Asc_to_Key(编号 As Long) As String
Dim NR As String
Open App.Path & "\Pic\System\Asc_to_Key.ini" For Input As #91
Do Until EOF(91)
  Line Input #91, NR
  If 编号 = Val(Left(NR, InStr(NR, "/") - 1)) Then
    Asc_to_Key = Right(NR, Len(NR) - InStr(NR, "/"))
    Close #91
    Exit Function
  End If
Loop
Close #91
Asc_to_Key = "None"
End Function
Public Function Key_to_Asc(按键名 As String) As Long
Dim NR As String
Open App.Path & "\Pic\System\Asc_to_Key.ini" For Input As #91
Do Until EOF(91)
  Line Input #91, NR
  If 按键名 = Right(NR, Len(NR) - InStr(NR, "/")) Then
    Key_to_Asc = Left(NR, InStr(NR, "/") - 1)
    Close #91
    Exit Function
  End If
Loop
Close #91
Key_to_Asc = "-1"
End Function
Public Function 随机取行(文本相对路径 As String) As String
Dim 临时计数行数 As Long, 临时计数循环 As Long
Randomize
Open App.Path & "\" & 文本相对路径 For Binary As #91
Do Until EOF(91)
  Line Input #91, 随机取行
  临时计数行数 = 1 + 临时计数行数
Loop
Close #91
Open App.Path & "\" & 文本相对路径 For Input As #91
For 临时计数循环 = 0 To Round(Rnd * 临时计数行数)
  Line Input #91, 随机取行
Next
Close #91
End Function
