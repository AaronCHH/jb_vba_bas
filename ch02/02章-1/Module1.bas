Attribute VB_Name = "Module1"

Sub ��J��r()
Attribute ��J��r.VB_Description = "��ʹ� �b 2006/2/28 ���s������"
Attribute ��J��r.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
' ��ʹ� �b 2006/2/28 ���s������
'

'
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Excel 2003"
    Range("B3").Select
    With Selection.Font
        .Name = "MS Gothic"
        .FontStyle = "����"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
End Sub
