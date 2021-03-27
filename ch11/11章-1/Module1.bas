Attribute VB_Name = "Module1"
Option Explicit


'------------------------------------
'範例79
'變更系統日期
'
'(變更後請回復現在的日期)
'------------------------------------

Sub ChangeDate()
    Date = #3/26/1985#
End Sub


'-----------------------------------
'範例80
'使用Int函數捨去小數位數
'-----------------------------------

Sub IntA1()
    Range("A1").Value = 123.456
    Range("A2").Value = Int(Range("A1").Value)
End Sub


'----------------------------------------
'範例81
'Excel VBA對工作表函數的運用
'----------------------------------------

Sub SearchMax()
    Dim myMax As Long
    
    myMax = Application.WorksheetFunction.Max(Range("B1:D10").Value)
    MsgBox "最大值是：" & myMax & "。"
End Sub
