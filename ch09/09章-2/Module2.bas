Attribute VB_Name = "Module2"
'-------------------------
'範例68
'選取清單內所有資料
'-------------------------

Sub SelList()
    ActiveSheet.ListObjects(1).Range.Select
    
End Sub

'-------------------------
'範例69
'選取清單標籤所屬的行
'-------------------------

Sub SelListHeader()
    ActiveSheet.ListObjects(1).HeaderRowRange.Select
    
End Sub

'-------------------------
'範例70
'選取清單中資料的部份
'-------------------------

Sub SelListBody()
    ActiveSheet.ListObjects(1).DataBodyRange.Select
    
End Sub

'-------------------------
'範例71
'選取插入資料的行
'-------------------------

Sub SelListNewRow()
    Range("A3").Select
    Selection.ListObject.InsertRowRange.Select
End Sub

Sub SelListNewRow2()
    Dim myRowRnage As Range
    
    Set myRowRnage = ActiActiveSheet.ListObjects(1).InsertRowRangeveSheet.ListObjects(1).InsertRowRange
    If myRowRnage Is Nothing Then
        MsgBox "請選取清單"
    Else
        myRowRnage.Select
    End If
End Sub

