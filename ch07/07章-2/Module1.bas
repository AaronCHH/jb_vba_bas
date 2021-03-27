
'-------------------------
'範例13
'插入新的工作表
'-------------------------

Sub 插入工作表()
    Worksheets.Add
End Sub


'-------------------------------------------
'範例14
'指定插入位及數量
'-------------------------------------------

Sub 插入工作表2()
    Worksheets.Add After:=Worksheets(1), Count:=2
End Sub


'----------------------------
'範例15
'指定使用中工作表
'----------------------------

Sub 指定工作表()
    Worksheets("Sheet3").Activate
End Sub

Sub 指定工作表2()
    Worksheets("Sheet3").Select
End Sub


'----------------------------------------
'範例16
'選取多個工作表
'----------------------------------------
  
Sub 選取多個工作表()
    Worksheets.Select               '選取全部工作表
End Sub

Sub 選取多個工作表2()
    Worksheets(Array(1, 3)).Select  '選取第1、3個工作表
End Sub


'----------------------------------
'範例17
'在同一個活頁簿中移動工作表
'----------------------------------

Sub 移動工作表1()
    Worksheets("Sheet1").Move After:=Worksheets("Sheet3")
End Sub


'-----------------------------------------
'範例18
'將工作表移到其他活頁簿中
'
'(請開啟「Dummy.xls」後再執行)
'-----------------------------------------

Sub 移動工作表2()
    Worksheets("Sheet4").Move _
        Before:=Workbooks("Dummy.xls").Sheets(2)
End Sub


'----------------------------------
'範例19
'將工作表移到新建的活頁簿中
'----------------------------------

Sub 移動工作表3()
    Worksheets("Sheet5").Move
End Sub


'-----------------------
'範例20
'在同一個活頁簿中複製工作表
'-----------------------
  
Sub 複製工作表()
    '在同一個活頁簿中複製工作表
    Worksheets("Sheet1").Copy After:=Worksheets("Sheet3")

    '將工作表複製到其他的活頁簿中
    'Worksheets("Sheet1").Copy Before:=Workbooks("Dummy.xls").Sheets(2)

    '將工作表複製到新建的活頁簿中
    'Worksheets("Sheet1").Copy
End Sub

