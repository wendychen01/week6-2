VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormInsert 
   Caption         =   "�ڪ������D"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "��J�������.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "UserFormInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnInsert_Click()
Dim supplyName As String
supplyName = txbName.Text '�����ӦW�� = ��J�ئ^�������e
Cells(2, 1).Value = supplyName

Dim supplyPhone As String
supplyPhone = txbPhone.Text '�����ӹq�� = ��J�ئ^�������e
Cells(2, 2).Value = supplyPhone

Dim price As Integer
price = txbPrice.Text '�X����� = ��J�ئ^�������e
Cells(2, 3).Value = CInt(price)

Dim finalprice As Integer
finalprice = txbFinalPrice.Text '�X������� = ��J�ئ^�������e
Cells(2, 4).Value = CInt(finalprice)

Dim totalDiscount As Single
totalDiscount = (price - finalprice) / price
Cells(2, 5).Value = totalDiscount

If (totalDiscount > 0.8) Then
    Cells(2, 6).Value = "���`"
    Else
    Cells(2, 6).Value = "���`"
End If

End Sub
