Attribute VB_Name = "�d�ߤu�@��"
Option Explicit

Sub demo()
Dim wSheet As Worksheet '�ŧiwSheet���u�@���A�ܼ�
For Each wSheet In Worksheets '�q�u�@���X�v�@����
MsgBox "���u�@�� : " & wSheet.Name
Next '���y�̫ܳ�,�����j��
MsgBox "�����j�鱽�y" '�������y
End Sub
