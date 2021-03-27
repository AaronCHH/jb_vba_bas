Attribute VB_Name = "Module2"

'------------------------------------------
'範例8
'將活頁簿存檔後關閉
'
'(請開啟「Dummy.xls」後再執行)
'------------------------------------------

Sub 關閉活頁簿5()
    Workbooks("Dummy.xls").Close SaveChanges:=True '指定引數的名稱
End Sub

'------------------------------------------
'範例9
'活頁簿關閉時不存檔
'
'(請開啟「Dummy.xls」後再執行)
'------------------------------------------

Sub 關閉活頁簿6()
    Workbooks("Dummy.xls").Close False              '標準引數
End Sub

'------------------------------------------
'範例10
'指定使用中的活頁簿
'
'(請開啟「Dummy.xls」後再執行)
'------------------------------------------

Sub 指定活頁簿()
Attribute 指定活頁簿.VB_ProcData.VB_Invoke_Func = " \n14"
    Workbooks("Dummy.xls").Activate
End Sub

'----------------------------------
'範例11
'將活頁簿存檔
'
'----------------------------------
 
Sub 活頁簿存檔()
    ActiveWorkbook.Save             '儲存使用中活頁簿
End Sub

Sub 活頁簿存檔2()
    Workbooks("Dummy.xls").Save    '指定存檔名稱後儲存
End Sub

'--------------------------------------------
'範例12
'儲存活頁簿時另存新檔'
'
'---------------------------------------------

Sub 儲存活頁簿3()
Attribute 儲存活頁簿3.VB_ProcData.VB_Invoke_Func = " \n14"
    '指定儲存目標
    ActiveWorkbook.SaveAs Filename:="C:\Excel2003VBA基礎篇\Test.xls"
End Sub

Sub 儲存活頁簿4()
Attribute 儲存活頁簿4.VB_ProcData.VB_Invoke_Func = " \n14"
    '存放到目前資料夾中
    ActiveWorkbook.SaveAs Filename:="Test.xls"
End Sub
