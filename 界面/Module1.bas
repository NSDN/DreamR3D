Attribute VB_Name = "oooooooǰ̨��������"
Public ����(0 To 6) As Long
Public Function ��ȡ����()
On Error Resume Next
Dim NR As String
Open App.Path & "\Save\Options.ini" For Input As #2
For i = 0 To UBound(����)
  Line Input #2, NR
  ����(i) = Val(NR)
Next
Close #2
End Function
Public Function Asc_to_Key(��� As Long) As String
Dim NR As String
Open App.Path & "\Pic\System\Asc_to_Key.ini" For Input As #91
Do Until EOF(91)
  Line Input #91, NR
  If ��� = Val(Left(NR, InStr(NR, "/") - 1)) Then
    Asc_to_Key = Right(NR, Len(NR) - InStr(NR, "/"))
    Close #91
    Exit Function
  End If
Loop
Close #91
Asc_to_Key = "None"
End Function
Public Function Key_to_Asc(������ As String) As Long
Dim NR As String
Open App.Path & "\Pic\System\Asc_to_Key.ini" For Input As #91
Do Until EOF(91)
  Line Input #91, NR
  If ������ = Right(NR, Len(NR) - InStr(NR, "/")) Then
    Key_to_Asc = Left(NR, InStr(NR, "/") - 1)
    Close #91
    Exit Function
  End If
Loop
Close #91
Key_to_Asc = "-1"
End Function
Public Function ���ȡ��(�ı����·�� As String) As String
Dim ��ʱ�������� As Long, ��ʱ����ѭ�� As Long
Randomize
Open App.Path & "\" & �ı����·�� For Binary As #91
Do Until EOF(91)
  Line Input #91, ���ȡ��
  ��ʱ�������� = 1 + ��ʱ��������
Loop
Close #91
Open App.Path & "\" & �ı����·�� For Input As #91
For ��ʱ����ѭ�� = 0 To Round(Rnd * ��ʱ��������)
  Line Input #91, ���ȡ��
Next
Close #91
End Function
