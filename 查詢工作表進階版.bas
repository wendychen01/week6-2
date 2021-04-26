Attribute VB_Name = "查詢工作表進階版"
Option Explicit

Sub demoAdv()
Dim usrStr As String '宣告usrStr為文字變數,儲存使用者輸入框的內容
usrStr = InputBox("請輸入要啟動的工作表名稱")

Dim wSheetNew As Worksheet '宣告wSheetNew為工作表型態
    For Each wSheetNew In Worksheets '從所有工作表集合逐一掃瞄
        If (wSheetNew.Name = usrStr) Then
        MsgBox "找到工作表 : " & wSheetNew.Name '執行彈跳視窗
        wSheetNew.Activate '啟動工作表
        Else
        MsgBox "Sorry尚未找到"
        End If
    Next
MsgBox "完成迴圈掃描"

End Sub
