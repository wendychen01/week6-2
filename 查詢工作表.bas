Attribute VB_Name = "查詢工作表"
Option Explicit

Sub demo()
Dim wSheet As Worksheet '宣告wSheet為工作表型態變數
For Each wSheet In Worksheets '從工作表集合逐一掃瞄
MsgBox "找到工作表 : " & wSheet.Name
Next '掃描至最後,結束迴圈
MsgBox "完成迴圈掃描" '結束掃描
End Sub
