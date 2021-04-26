Attribute VB_Name = "有中斷迴圈"
Option Explicit

Sub exitForDemo1()
Dim i, sum As Integer '宣告i,sum為整數型態的變數
sum = 0 'sum初始值為0
    For i = 0 To 10
        MsgBox "目前i = " & i '顯示目前次數
        sum = sum + i
        MsgBox "目前總和 = " & sum '顯示本次加總
    If i >= 4 Then '如果符合條件
    MsgBox "符合>=4" '顯示提示
    Exit For '中斷迴圈
    End If '中斷判斷
    Next '結束迴圈
MsgBox "目前總和 = " & sum '最終結果
End Sub

 
