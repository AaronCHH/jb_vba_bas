Attribute VB_Name = "Module2"

'------------------------------------
'�d��33
'�N�x�s�椺�������ܩ��ܤ����
'------------------------------------

Sub ValueRange1()
    MsgBox Range("A1").Value
End Sub


'---------------------
'�d��34
'�b�x�s�椤��J��r
'---------------------

Sub ValueRange2()
    Range("A1").Value = "XYZ"
    Worksheets("Sheet7").Cells(1, 1).Value = "XYZ"
    Worksheets("Sheet7").Range("B1:D5").Value = "XYZ"
End Sub


'------------------------------------
'�d��35
'��J�x�s��U�خ榡
'------------------------------------
  
Sub ValueRange3()
    Range("A1").Value = 100.35          '�q�ή榡
    Range("A2").Value = "-1,573,500"    '�d����
    Range("A3").Value = "2003/7/29"     '���
    Range("A4").Value = "10:25:30"      '�ɶ�
    Range("A5").Value = "'0123"         '��r
End Sub


'--------------------------
'�d��36
'�N�x�s�檺�ȿ�J���L�x�s��
'--------------------------
  
Sub ValueRange4()
    'Range("B10").Value = Range("A10").Value
    Range("B10") = Range("A10")
End Sub
