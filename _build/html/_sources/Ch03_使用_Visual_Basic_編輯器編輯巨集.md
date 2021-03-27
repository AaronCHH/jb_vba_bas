# Ch03 使用 Visual Basic 編輯器編輯巨集

## 變更儲存格背景色的巨集
```
Attribute VB_Name = "Module1"
'--------------------------
'變更儲存格背景色的巨集
'--------------------------
Sub 變更顏色()
    Range("A1:B5").Select           '選取儲存格
    With Selection.Interior         '設定背景色
        .ColorIndex = 6
        .Pattern = xlSolid
    End With
End Sub

```

## 複製儲存格資料的巨集
```
Sub 複製資料()
'----------------------------
'複製儲存格資料的巨集
'----------------------------
    Range("A1:A5").Select           '選取儲存格
    Selection.Copy                  '複製選取範圍中的資料
    Range("B1:B5").Select           '選取儲存格
    ActiveSheet.Paste               '將複製的資料貼上選取的儲存格中
    Application.CutCopyMode = False '解除複製的狀態
End Sub
```


## 在儲存格外繪製格線
```
Attribute VB_Name = "Module2"
'---------------------
'在儲存格外繪製格線
'---------------------
Sub 繪製格線()
    Range("A1:D10").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
```