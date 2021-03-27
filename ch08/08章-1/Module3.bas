Attribute VB_Name = "Module3"

'--------------------------
'範例37
'在儲存格中以A1格式輸入計算式
'--------------------------
  
Sub FormulaRange1()
    Range("A10").Formula = "=SUM(A1:A9)"
    Range("B10").Formula = "=AVERAGE(B1:B9)"
    Range("C10").Formula = "=MAX(C1:C9)"
    Range("D10").Formula = "=MIN(D1:D9)"
End Sub


'-----------------------------
'範例38
'在儲存格中以R1C1格式輸入計算式
'-----------------------------

Sub FormulaRange2()
    Worksheets("Sheet5").Range("E1:E10").FormulaR1C1 = "=RC[-2]+RC[-1]"
End Sub


'-----------------------------
'範例39
'取得儲存格的值(計算式的結果)
'-----------------------------
  
Sub GetValue()
    Range("F1").Value = Range("E1").Value
End Sub


'-------------------
'範例40
'取得儲存格的計算式
'-------------------
  
Sub GetFormula()
    Range("F1").Formula = Range("E1").Formula
    'Range("F1").FormulaR1C1 = Range("E1").FormulaR1C1
End Sub
