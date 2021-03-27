Attribute VB_Name = "Module1"

Sub 鍵入文字()
Attribute 鍵入文字.VB_Description = "梁銘鼎 在 2006/2/28 錄製的巨集"
Attribute 鍵入文字.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
' 梁銘鼎 在 2006/2/28 錄製的巨集
'

'
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Excel 2003"
    Range("B3").Select
    With Selection.Font
        .Name = "MS Gothic"
        .FontStyle = "粗體"
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
