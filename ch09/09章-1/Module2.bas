Attribute VB_Name = "Module2"

'-------------------------
'�d��61
'���o�����x�s��
'-------------------------

Sub SelEndCell()
    Range("C1").End(xlDown).Select
    'Range("C2").End(xlDown).Select
    'Range("C3").End(xlDown).Select
End Sub


'--------------------------------------
'�d��62
'���̫ܳ�@����ƳB
'--------------------------------------

Sub SelLastCell()
    Range("A3").End(xlDown).Select
    Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Select
End Sub


'--------------------------------
'�d��63
'���ܷs�ظ�ƪ��ت��x�s��
'--------------------------------

Sub SelNewCell()
    Range("A65536").End(xlUp).Offset(1).Select
End Sub


'-------------------------
'�d��64
'����S�w���
'-------------------------

Sub SelRecords()
    Range("A5", Range("A5").End(xlToRight)).Select
End Sub


