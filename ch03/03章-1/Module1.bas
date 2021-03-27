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
