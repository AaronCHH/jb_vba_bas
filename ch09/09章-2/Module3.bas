Attribute VB_Name = "Module3"

'-------------------------
'�d��72
'���o�M�椤��Ƶ���
'-------------------------

Sub CountListData()
    MsgBox ActiveSheet.ListObjects(1).ListRows.Count
    
End Sub

'-------------------------
'�d��73
'����M�椤�S�w����/�C
'-------------------------

Sub SelListRow()
    ActiveSheet.ListObjects(1).ListRows(3).Range.Select
'    ActiveSheet.ListObjects(1).ListColumns(3).Range.Select
    
End Sub


'-------------------------
'�d��74
'�b�M�椤���J��
'-------------------------

Sub InsertListRow()
    ActiveSheet.ListObjects(1).ListRows.Add (2)
    
End Sub

'-------------------------
'�d��75
'�C�L�M��
'-------------------------

Sub PrintList()
'    ActiveSheet.ListObjects(1).Range.PrintOut
    ActiveSheet.ListObjects(1).Range.PrintPreview
    
End Sub


