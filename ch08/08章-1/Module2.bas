Attribute VB_Name = "Module2"

'------------------------------------
'範例33
'將儲存格內的資料顯示於對話方塊中
'------------------------------------

Sub ValueRange1()
    MsgBox Range("A1").Value
End Sub


'---------------------
'範例34
'在儲存格中輸入文字
'---------------------

Sub ValueRange2()
    Range("A1").Value = "XYZ"
    Worksheets("Sheet7").Cells(1, 1).Value = "XYZ"
    Worksheets("Sheet7").Range("B1:D5").Value = "XYZ"
End Sub


'------------------------------------
'範例35
'輸入儲存格各種格式
'------------------------------------
  
Sub ValueRange3()
    Range("A1").Value = 100.35          '通用格式
    Range("A2").Value = "-1,573,500"    '千分位
    Range("A3").Value = "2003/7/29"     '日期
    Range("A4").Value = "10:25:30"      '時間
    Range("A5").Value = "'0123"         '文字
End Sub


'--------------------------
'範例36
'將儲存格的值輸入到其他儲存格
'--------------------------
  
Sub ValueRange4()
    'Range("B10").Value = Range("A10").Value
    Range("B10") = Range("A10")
End Sub
