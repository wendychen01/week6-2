Attribute VB_Name = "�d�ߤu�@��i����"
Option Explicit

Sub demoAdv()
Dim usrStr As String '�ŧiusrStr����r�ܼ�,�x�s�ϥΪ̿�J�ت����e
usrStr = InputBox("�п�J�n�Ұʪ��u�@��W��")

Dim wSheetNew As Worksheet '�ŧiwSheetNew���u�@���A
    For Each wSheetNew In Worksheets '�q�Ҧ��u�@���X�v�@����
        If (wSheetNew.Name = usrStr) Then
        MsgBox "���u�@�� : " & wSheetNew.Name '����u������
        wSheetNew.Activate '�Ұʤu�@��
        Else
        MsgBox "Sorry�|�����"
        End If
    Next
MsgBox "�����j�鱽�y"

End Sub
