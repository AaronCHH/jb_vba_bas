
'---------------------
'範例21
'刪除工作表
'---------------------

Sub 刪除工作表()
    Application.DisplayAlerts = False
    Worksheets("Sheet6").Delete
    Application.DisplayAlerts = True
End Sub

Sub 刪除工作表2()
    Dim myChart As Chart
    
    '刪除所有圖表工作表
    For Each myChart In Charts
        myChart.Delete
    Next
End Sub


'------------------------
'範例22
'隱藏工作表
'------------------------

Sub 隱藏工作表()
    Worksheets("Sheet3").Visible = False
End Sub

Sub 隱藏工作表2()
    '不是使用「顯示」指令
    Worksheets("Sheet3").Visible = xlVeryHidden
End Sub


'-----------------------
'範例23
'顯示工作表
'-----------------------

Sub 顯示工作表()
    Worksheets("Sheet3").Visible = True
End Sub
