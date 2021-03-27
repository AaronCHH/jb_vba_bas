Attribute VB_Name = "Module1"

'------------------------------------------
'範例41
'不經剪貼簿複製資料
'------------------------------------------

Sub CopyData()
    Worksheets("Sheet1").Range("A1:B10").Copy _
        Destination:=Worksheets("Sheet2").Range("A1")
End Sub

'-------------------------
'範例42
'清除儲存格的計算式與值
'-------------------------
  
Sub ClearRange1()
    Range("A1").Select
    ActiveCell.Value = ""
End Sub
 
Sub ClearRange2()
    Range("A1:D5").Select
    Selection.ClearContents
End Sub


'---------------------
'範例43
'清除儲存格的格式
'---------------------

Sub ClearRange3()
    Worksheets("Sheet3").Range("A10:D12").ClearFormats
End Sub


'----------------------------
'範例44
'清除儲存格資料的資料與格式
'----------------------------
  
Sub ClearRange4()
    Worksheets("Sheet3").Range("A10:D12").Clear
End Sub
