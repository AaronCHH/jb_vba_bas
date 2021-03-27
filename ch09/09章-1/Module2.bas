Attribute VB_Name = "Module2"

'-------------------------
'範例61
'取得末端儲存格
'-------------------------

Sub SelEndCell()
    Range("C1").End(xlDown).Select
    'Range("C2").End(xlDown).Select
    'Range("C3").End(xlDown).Select
End Sub


'--------------------------------------
'範例62
'移至最後一筆資料處
'--------------------------------------

Sub SelLastCell()
    Range("A3").End(xlDown).Select
    Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Select
End Sub


'--------------------------------
'範例63
'移至新建資料的目的儲存格
'--------------------------------

Sub SelNewCell()
    Range("A65536").End(xlUp).Offset(1).Select
End Sub


'-------------------------
'範例64
'選取特定資料
'-------------------------

Sub SelRecords()
    Range("A5", Range("A5").End(xlToRight)).Select
End Sub


