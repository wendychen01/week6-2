Attribute VB_Name = "付款人狀態"
Option Explicit

Sub 查詢付款人狀態()
Dim people As String '宣告查詢人型態
people = Range("G1").Value
Dim rowNum As Integer '宣告查詢列數的變數為整數
Dim content As String '宣告視窗內容變數
Dim paySatus As Boolean '宣告付款狀態為Boolean

For rowNum = 2 To 7
    If (Cells(rowNum, "A").Value = people) Then
        Range("G2").Value = Cells(rowNum, "B").Value
        If (Cells(rowNum, "C").Value = 0) Then
        paySatus = False
        Else
        paySatus = True
        End If
    content = people & "付款狀態 = " & paySatus
    MsgBox (content)
    Exit For
    Else
    End If
Next
End Sub
