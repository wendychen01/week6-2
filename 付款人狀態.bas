Attribute VB_Name = "�I�ڤH���A"
Option Explicit

Sub �d�ߥI�ڤH���A()
Dim people As String '�ŧi�d�ߤH���A
people = Range("G1").Value
Dim rowNum As Integer '�ŧi�d�ߦC�ƪ��ܼƬ����
Dim content As String '�ŧi�������e�ܼ�
Dim paySatus As Boolean '�ŧi�I�ڪ��A��Boolean

For rowNum = 2 To 7
    If (Cells(rowNum, "A").Value = people) Then
        Range("G2").Value = Cells(rowNum, "B").Value
        If (Cells(rowNum, "C").Value = 0) Then
        paySatus = False
        Else
        paySatus = True
        End If
    content = people & "�I�ڪ��A = " & paySatus
    MsgBox (content)
    Exit For
    Else
    End If
Next
End Sub
