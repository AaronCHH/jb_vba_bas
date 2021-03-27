Attribute VB_Name = "Module2"
Option Explicit

'----------------------
'範例77
'計算工作表的數量
'----------------------

Sub DisplayWSCnt()
    Dim myWSCnt As Integer
    
    myWSCnt = ActiveWorkbook.Worksheets.Count
    MsgBox myWSCnt
End Sub


'----------------------------------------
'範例78
'使用物件變數的程序
'
'請先開啟Dummy.xls後再執行
'----------------------------------------


Sub SetObject()
    Dim myWSheet As Worksheet
        
    Set myWSheet = Workbooks("Dummy.xls").Worksheets("Sheet2")
    
    myWSheet.Range("A1:D10").Value = "ABC"
End Sub
