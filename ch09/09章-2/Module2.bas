Attribute VB_Name = "Module2"
'-------------------------
'�d��68
'����M�椺�Ҧ����
'-------------------------

Sub SelList()
    ActiveSheet.ListObjects(1).Range.Select
    
End Sub

'-------------------------
'�d��69
'����M����ҩ��ݪ���
'-------------------------

Sub SelListHeader()
    ActiveSheet.ListObjects(1).HeaderRowRange.Select
    
End Sub

'-------------------------
'�d��70
'����M�椤��ƪ�����
'-------------------------

Sub SelListBody()
    ActiveSheet.ListObjects(1).DataBodyRange.Select
    
End Sub

'-------------------------
'�d��71
'������J��ƪ���
'-------------------------

Sub SelListNewRow()
    Range("A3").Select
    Selection.ListObject.InsertRowRange.Select
End Sub

Sub SelListNewRow2()
    Dim myRowRnage As Range
    
    Set myRowRnage = ActiActiveSheet.ListObjects(1).InsertRowRangeveSheet.ListObjects(1).InsertRowRange
    If myRowRnage Is Nothing Then
        MsgBox "�п���M��"
    Else
        myRowRnage.Select
    End If
End Sub

