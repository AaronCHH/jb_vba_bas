Attribute VB_Name = "Module3"

'--------------------------
'�d��37
'�b�x�s�椤�HA1�榡��J�p�⦡
'--------------------------
  
Sub FormulaRange1()
    Range("A10").Formula = "=SUM(A1:A9)"
    Range("B10").Formula = "=AVERAGE(B1:B9)"
    Range("C10").Formula = "=MAX(C1:C9)"
    Range("D10").Formula = "=MIN(D1:D9)"
End Sub


'-----------------------------
'�d��38
'�b�x�s�椤�HR1C1�榡��J�p�⦡
'-----------------------------

Sub FormulaRange2()
    Worksheets("Sheet5").Range("E1:E10").FormulaR1C1 = "=RC[-2]+RC[-1]"
End Sub


'-----------------------------
'�d��39
'���o�x�s�檺��(�p�⦡�����G)
'-----------------------------
  
Sub GetValue()
    Range("F1").Value = Range("E1").Value
End Sub


'-------------------
'�d��40
'���o�x�s�檺�p�⦡
'-------------------
  
Sub GetFormula()
    Range("F1").Formula = Range("E1").Formula
    'Range("F1").FormulaR1C1 = Range("E1").FormulaR1C1
End Sub
