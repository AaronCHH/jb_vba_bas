Attribute VB_Name = "Module1"

'-------------------------
'�d��65
'�إ߸�Ʈw�M��
'-------------------------

Sub SetList()
    Range("A3").Select
    ActiveSheet.ListObjects.Add
    
End Sub

'-------------------------
'�d��66
'�]�w�M��W��
'-------------------------

Sub SetListName()
    Range("A3").Select
    ActiveSheet.ListObjects.Add.Name = "�|���W�U"
End Sub

'-------------------------
'�d��67
'�����M��
'-------------------------

Sub ChangeUnList()
    ActiveSheet.ListObjects(1).Unlist
'    ActiveSheet.ListObjects("�|���W�U").Unlist
    
End Sub

