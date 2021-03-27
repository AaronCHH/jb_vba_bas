Attribute VB_Name = "Module3"

'-------------------------
'範例72
'取得清單中資料筆數
'-------------------------

Sub CountListData()
    MsgBox ActiveSheet.ListObjects(1).ListRows.Count
    
End Sub

'-------------------------
'範例73
'選取清單中特定的行/列
'-------------------------

Sub SelListRow()
    ActiveSheet.ListObjects(1).ListRows(3).Range.Select
'    ActiveSheet.ListObjects(1).ListColumns(3).Range.Select
    
End Sub


'-------------------------
'範例74
'在清單中插入行
'-------------------------

Sub InsertListRow()
    ActiveSheet.ListObjects(1).ListRows.Add (2)
    
End Sub

'-------------------------
'範例75
'列印清單
'-------------------------

Sub PrintList()
'    ActiveSheet.ListObjects(1).Range.PrintOut
    ActiveSheet.ListObjects(1).Range.PrintPreview
    
End Sub


