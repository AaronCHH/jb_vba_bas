Attribute VB_Name = "Module1"

'------------------------------------
'範例76
'將活頁簿名稱顯示於對話方塊上
'------------------------------------

Sub DisplayWBName()
    myWBName = Workbooks(1).Name
    MsgBox "第一個開啟的活頁簿是：" & myWBName & "。"
End Sub

