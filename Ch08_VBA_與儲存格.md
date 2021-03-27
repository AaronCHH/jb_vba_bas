# Ch08 VBA 與儲存格
## 08章-1
## 範例24 選取單一儲存格
```
Attribute VB_Name = "Module1"
'-------------------
'範例24 選取單一儲存格
'-------------------

Sub RangeSel1()
    Range("C5").Select
End Sub
 
``` 
## 範例25 選取連續的儲存格範圍
```
'-------------------------
'範例25 選取連續的儲存格範圍
'-------------------------
Sub RangeSel2()
    Range("B2:D4").Select
    'Range("B2", "D4").Select
End Sub

```
## 範例26 選取不連續的儲存格範圍
```
'-------------------------
'範例26 選取不連續的儲存格範圍
'-------------------------
  
Sub RangeSel3()
    'Range("B2,B4,D2,D4").Select
    Range("B2:D3,B5:D6").Select
End Sub

```
## 範例27 選取定義名稱的儲存格
```
'----------------------------
'範例27 選取定義名稱的儲存格
'----------------------------
  
Sub RangeSel4()
    Range("營業額總計").Select
End Sub

```
## 範例28 選取行/列
```
'-------------------
'範例28 選取行/列
'-------------------
  
Sub RangeSel5()
    Range("1:1").Select
    'Range("A:A").Select
    'Range("1:3").Select
    'Range("A:C").Select
    'Range("1:3,6:6").Select
    'Range("A:C,F:F").Select
End Sub

```
## 範例29 使用Cells屬性選取單一儲存格
```
'-----------------------------------
'範例29 使用Cells屬性選取單一儲存格
'-----------------------------------

Sub CellsSel1()
    Cells(5, 3).Activate
    'Cells(5, "C").Activate
End Sub
```
## 範例30 以編號選取儲存格
```
'-----------------------
'範例30 以編號選取儲存格
'-----------------------
  
Sub CellsSel2()
    Cells(1027).Activate
End Sub

```
## 範例31 使用Cells屬性選取所有儲存格
```
'-------------------------------
'範例31 使用Cells屬性選取所有儲存格
'-------------------------------
  
Sub CellsSel3()
    Cells.Select
End Sub

```
## 範例32 使用Cells屬性選取儲存格範圍
```
'---------------------------------
'範例32 使用Cells屬性選取儲存格範圍
'---------------------------------
  
Sub CellsSel4()
    Range(Cells(1, 2), Cells(5, 4)).Select
End Sub

```
## 範例33 將儲存格內的資料顯示於對話方塊中
```
Attribute VB_Name = "Module2"
'------------------------------------
'範例33 將儲存格內的資料顯示於對話方塊中
'------------------------------------

Sub ValueRange1()
    MsgBox Range("A1").Value
End Sub

```
## 範例34 在儲存格中輸入文字
```
'---------------------
'範例34 在儲存格中輸入文字
'---------------------

Sub ValueRange2()
    Range("A1").Value = "XYZ"
    Worksheets("Sheet7").Cells(1, 1).Value = "XYZ"
    Worksheets("Sheet7").Range("B1:D5").Value = "XYZ"
End Sub

```
## 範例35 輸入儲存格各種格式
```
'------------------------------------
'範例35 輸入儲存格各種格式
'------------------------------------
  
Sub ValueRange3()
    Range("A1").Value = 100.35          '通用格式
    Range("A2").Value = "-1,573,500"    '千分位
    Range("A3").Value = "2003/7/29"     '日期
    Range("A4").Value = "10:25:30"      '時間
    Range("A5").Value = "'0123"         '文字
End Sub

```
## 範例36 將儲存格的值輸入到其他儲存格
```
'--------------------------
'範例36 將儲存格的值輸入到其他儲存格
'--------------------------
  
Sub ValueRange4()
    'Range("B10").Value = Range("A10").Value
    Range("B10") = Range("A10")
End Sub

Attribute VB_Name = "Module3"
```
## 範例37 在儲存格中以A1格式輸入計算式
```
'--------------------------
'範例37 在儲存格中以A1格式輸入計算式
'--------------------------
  
Sub FormulaRange1()
    Range("A10").Formula = "=SUM(A1:A9)"
    Range("B10").Formula = "=AVERAGE(B1:B9)"
    Range("C10").Formula = "=MAX(C1:C9)"
    Range("D10").Formula = "=MIN(D1:D9)"
End Sub

```
## 範例38 在儲存格中以R1C1格式輸入計算式
```
'-----------------------------
'範例38 在儲存格中以R1C1格式輸入計算式
'-----------------------------

Sub FormulaRange2()
    Worksheets("Sheet5").Range("E1:E10").FormulaR1C1 = "=RC[-2]+RC[-1]"
End Sub

```
## 範例39 取得儲存格的值(計算式的結果)
```
'-----------------------------
'範例39 取得儲存格的值(計算式的結果)
'-----------------------------
  
Sub GetValue()
    Range("F1").Value = Range("E1").Value
End Sub

```
## 範例40 取得儲存格的計算式
```
'-------------------
'範例40 取得儲存格的計算式
'-------------------
  
Sub GetFormula()
    Range("F1").Formula = Range("E1").Formula
    'Range("F1").FormulaR1C1 = Range("E1").FormulaR1C1
End Sub


## 08章-2

Attribute VB_Name = "Module1"
```
## 範例41 不經剪貼簿複製資料
```
'------------------------------------------
'範例41 不經剪貼簿複製資料
'------------------------------------------

Sub CopyData()
    Worksheets("Sheet1").Range("A1:B10").Copy _
        Destination:=Worksheets("Sheet2").Range("A1")
End Sub
```
## 範例42 清除儲存格的計算式與值
```
'-------------------------
'範例42 清除儲存格的計算式與值
'-------------------------
  
Sub ClearRange1()
    Range("A1").Select
    ActiveCell.Value = ""
End Sub
 
Sub ClearRange2()
    Range("A1:D5").Select
    Selection.ClearContents
End Sub

```
## 範例43 清除儲存格的格式
```
'---------------------
'範例43 清除儲存格的格式
'---------------------

Sub ClearRange3()
    Worksheets("Sheet3").Range("A10:D12").ClearFormats
End Sub

```
## 範例44 清除儲存格資料的資料與格式
```
'----------------------------
'範例44 清除儲存格資料的資料與格式
'----------------------------
  
Sub ClearRange4()
    Worksheets("Sheet3").Range("A10:D12").Clear
End Sub


## 08章-3

Attribute VB_Name = "Module1"
```
## 範例45 變更儲存格範圍的位置
```
'------------------------------------
'範例45 變更儲存格範圍的位置
'------------------------------------
  
Sub OffRange1()
    Selection.Offset(-1, 2).Select
End Sub

```
## 範例46 變更儲存格範圍的行位置
```
'-----------------------------
'範例46 變更儲存格範圍的行位置
'-----------------------------
  
Sub OffRange2()
    'Selection.Offset(2).Select
    Selection.Offset(2, 0).Select
End Sub

```
## 範例47 變更儲存格範圍的列位置
```
'-----------------------------
'範例47 變更儲存格範圍的列位置
'-----------------------------

Sub OffRange3()
    'Selection.Offset(0, -1).Select
    Selection.Offset(, -1).Select
End Sub

```
## 範例48 隱藏行
```
'----------------
'範例48 隱藏行
'----------------
 
Sub HideRows()
    Worksheets("Sheet2").Rows("5:7").Hidden = True
End Sub

```
## 範例49 取得儲存格範圍的行數
```
'------------------------------------------
'範例49 取得儲存格範圍的行數
'
'(顯示第5行到第7行後再執行)
'------------------------------------------

Sub CountRows()
    Range("B5:D7").Select
    MsgBox Selection.Rows.Count
End Sub

```
## 範例50 對被選取儲存格整列填滿資料
```
'--------------------------------
'範例50 對被選取儲存格整列填滿資料
'--------------------------------

Sub ValueRows2()
    Range("B5:D7").Select
    Selection.EntireRow.Value = "VBA"
End Sub

```
## 範例51 取得儲存格範圍的列數
```
'-----------------------------------
'範例51 取得儲存格範圍的列數
'-----------------------------------

Sub CountColumns()
    Range("B2:C5").Select
    MsgBox Selection.Columns.Count
End Sub

```
## 範例52 變更儲存格範圍的區域
```
'----------------------------
'範例52 變更儲存格範圍的區域
'----------------------------

Sub ResizeRange1()
    Range("B2:C4").Select
    MsgBox "變更儲存格範圍的區域"
    Selection.Resize(Selection.Rows.Count + 2, Selection.Columns.Count - 1).Select
End Sub

```
## 範例53 將Offset及Resize屬性合併使用
```
'------------------------------------------
'範例53 將Offset及Resize屬性合併使用
'------------------------------------------

Sub ResizeRange2()
    Range("B2:C4").Select
    MsgBox "變更儲存格範圍的區域"
    Selection.Offset(2).Resize(, Selection.Columns.Count + 2).Select
End Sub
```
## 範例54 將空白儲存格的背景色變更為藍色
```
'------------------------
'範例54 將空白儲存格的背景色變更為藍色
'------------------------

Sub BlankBlue()
    Range("A1").CurrentRegion.SpecialCells(xlCellTypeBlanks). _
        Interior.ColorIndex = 5
End Sub

```
## 範例55 將含有計算式的儲存格背景色變更為藍色
```
'-------------------------------------
'範例55 將含有計算式的儲存格背景色變更為藍色
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
```