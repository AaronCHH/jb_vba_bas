Attribute VB_Name = "Module1"

'-------------------
'�d��24
'�����@�x�s��
'-------------------

Sub RangeSel1()
    Range("C5").Select
End Sub
 
 
'-------------------------
'�d��25
'����s���x�s��d��
'-------------------------
Sub RangeSel2()
    Range("B2:D4").Select
    'Range("B2", "D4").Select
End Sub


'-------------------------
'�d��26
'������s���x�s��d��
'-------------------------
  
Sub RangeSel3()
    'Range("B2,B4,D2,D4").Select
    Range("B2:D3,B5:D6").Select
End Sub


'----------------------------
'�d��27
'����w�q�W�٪��x�s��
'----------------------------
  
Sub RangeSel4()
    Range("��~�B�`�p").Select
End Sub


'-------------------
'�d��28
'�����/�C
'-------------------
  
Sub RangeSel5()
    Range("1:1").Select
    'Range("A:A").Select
    'Range("1:3").Select
    'Range("A:C").Select
    'Range("1:3,6:6").Select
    'Range("A:C,F:F").Select
End Sub


'-----------------------------------
'�d��29
'�ϥ�Cells�ݩʿ����@�x�s��
'-----------------------------------

Sub CellsSel1()
    Cells(5, 3).Activate
    'Cells(5, "C").Activate
End Sub

'-----------------------
'�d��30
'�H�s������x�s��
'-----------------------
  
Sub CellsSel2()
    Cells(1027).Activate
End Sub


'-------------------------------
'�d��31
'�ϥ�Cells�ݩʿ���Ҧ��x�s��
'-------------------------------
  
Sub CellsSel3()
    Cells.Select
End Sub


'---------------------------------
'�d��32
'�ϥ�Cells�ݩʿ���x�s��d��
'---------------------------------
  
Sub CellsSel4()
    Range(Cells(1, 2), Cells(5, 4)).Select
End Sub
