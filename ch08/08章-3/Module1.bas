Attribute VB_Name = "Module1"

'------------------------------------
'�d��45
'�ܧ��x�s��d�򪺦�m
'------------------------------------
  
Sub OffRange1()
    Selection.Offset(-1, 2).Select
End Sub


'-----------------------------
'�d��46
'�ܧ��x�s��d�򪺦��m
'-----------------------------
  
Sub OffRange2()
    'Selection.Offset(2).Select
    Selection.Offset(2, 0).Select
End Sub


'-----------------------------
'�d��47
'�ܧ��x�s��d�򪺦C��m
'-----------------------------

Sub OffRange3()
    'Selection.Offset(0, -1).Select
    Selection.Offset(, -1).Select
End Sub


'----------------
'�d��48
'���æ�
'----------------
 
Sub HideRows()
    Worksheets("Sheet2").Rows("5:7").Hidden = True
End Sub


'------------------------------------------
'�d��49
'���o�x�s��d�򪺦��
'
'(��ܲ�5����7���A����)
'------------------------------------------

Sub CountRows()
    Range("B5:D7").Select
    MsgBox Selection.Rows.Count
End Sub


'--------------------------------
'�d��50
'��Q����x�s���C�񺡸��
'--------------------------------

Sub ValueRows2()
    Range("B5:D7").Select
    Selection.EntireRow.Value = "VBA"
End Sub


'-----------------------------------
'�d��51
'���o�x�s��d�򪺦C��
'-----------------------------------

Sub CountColumns()
    Range("B2:C5").Select
    MsgBox Selection.Columns.Count
End Sub


'----------------------------
'�d��52
'�ܧ��x�s��d�򪺰ϰ�
'----------------------------

Sub ResizeRange1()
    Range("B2:C4").Select
    MsgBox "�ܧ��x�s��d�򪺰ϰ�"
    Selection.Resize(Selection.Rows.Count + 2, Selection.Columns.Count - 1).Select
End Sub


'------------------------------------------
'�d��53
'�NOffset��Resize�ݩʦX�֨ϥ�
'------------------------------------------

Sub ResizeRange2()
    Range("B2:C4").Select
    MsgBox "�ܧ��x�s��d�򪺰ϰ�"
    Selection.Offset(2).Resize(, Selection.Columns.Count + 2).Select
End Sub

'------------------------
'�d��54
'�N�ť��x�s�檺�I�����ܧ��Ŧ�
'------------------------

Sub BlankBlue()
    Range("A1").CurrentRegion.SpecialCells(xlCellTypeBlanks). _
        Interior.ColorIndex = 5
End Sub


'-------------------------------------
'�d��55
'�N�t���p�⦡���x�s��I�����ܧ��Ŧ�
'-------------------------------------

Sub FormulaBlue()
    Cells.SpecialCells(xlCellTypeFormulas). _
        Interior.ColorIndex = 5
End Sub


'----------------------------------
'�٭��x�s��A1��D10�I���⪺�{��
'----------------------------------

Sub CellBackClear()
    Range("A1:D10").Interior.ColorIndex = xlNone
End Sub

