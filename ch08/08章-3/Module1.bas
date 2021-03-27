Attribute VB_Name = "Module1"

'------------------------------------
'範例45
'變更儲存格範圍的位置
'------------------------------------
  
Sub OffRange1()
    Selection.Offset(-1, 2).Select
End Sub


'-----------------------------
'範例46
'變更儲存格範圍的行位置
'-----------------------------
  
Sub OffRange2()
    'Selection.Offset(2).Select
    Selection.Offset(2, 0).Select
End Sub


'-----------------------------
'範例47
'變更儲存格範圍的列位置
'-----------------------------

Sub OffRange3()
    'Selection.Offset(0, -1).Select
    Selection.Offset(, -1).Select
End Sub


'----------------
'範例48
'隱藏行
'----------------
 
Sub HideRows()
    Worksheets("Sheet2").Rows("5:7").Hidden = True
End Sub


'------------------------------------------
'範例49
'取得儲存格範圍的行數
'
'(顯示第5行到第7行後再執行)
'------------------------------------------

Sub CountRows()
    Range("B5:D7").Select
    MsgBox Selection.Rows.Count
End Sub


'--------------------------------
'範例50
'對被選取儲存格整列填滿資料
'--------------------------------

Sub ValueRows2()
    Range("B5:D7").Select
    Selection.EntireRow.Value = "VBA"
End Sub


'-----------------------------------
'範例51
'取得儲存格範圍的列數
'-----------------------------------

Sub CountColumns()
    Range("B2:C5").Select
    MsgBox Selection.Columns.Count
End Sub


'----------------------------
'範例52
'變更儲存格範圍的區域
'----------------------------

Sub ResizeRange1()
    Range("B2:C4").Select
    MsgBox "變更儲存格範圍的區域"
    Selection.Resize(Selection.Rows.Count + 2, Selection.Columns.Count - 1).Select
End Sub


'------------------------------------------
'範例53
'將Offset及Resize屬性合併使用
'------------------------------------------

Sub ResizeRange2()
    Range("B2:C4").Select
    MsgBox "變更儲存格範圍的區域"
    Selection.Offset(2).Resize(, Selection.Columns.Count + 2).Select
End Sub

'------------------------
'範例54
'將空白儲存格的背景色變更為藍色
'------------------------

Sub BlankBlue()
    Range("A1").CurrentRegion.SpecialCells(xlCellTypeBlanks). _
        Interior.ColorIndex = 5
End Sub


'-------------------------------------
'範例55
'將含有計算式的儲存格背景色變更為藍色
'-------------------------------------

Sub FormulaBlue()
    Cells.SpecialCells(xlCellTypeFormulas). _
        Interior.ColorIndex = 5
End Sub


'----------------------------------
'還原儲存格A1到D10背景色的程序
'----------------------------------

Sub CellBackClear()
    Range("A1:D10").Interior.ColorIndex = xlNone
End Sub

