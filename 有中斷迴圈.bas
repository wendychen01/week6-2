Attribute VB_Name = "�����_�j��"
Option Explicit

Sub exitForDemo1()
Dim i, sum As Integer '�ŧii,sum����ƫ��A���ܼ�
sum = 0 'sum��l�Ȭ�0
    For i = 0 To 10
        MsgBox "�ثei = " & i '��ܥثe����
        sum = sum + i
        MsgBox "�ثe�`�M = " & sum '��ܥ����[�`
    If i >= 4 Then '�p�G�ŦX����
    MsgBox "�ŦX>=4" '��ܴ���
    Exit For '���_�j��
    End If '���_�P�_
    Next '�����j��
MsgBox "�ثe�`�M = " & sum '�̲׵��G
End Sub

 
