Attribute VB_Name = "Module1"

'-------------------
'範例24
'選取單一儲存格
'-------------------

Sub RangeSel1()
    Range("C5").Select
End Sub
 
 
'-------------------------
'範例25
'選取連續的儲存格範圍
'-------------------------
Sub RangeSel2()
    Range("B2:D4").Select
    'Range("B2", "D4").Select
End Sub


'-------------------------
'範例26
'選取不連續的儲存格範圍
'-------------------------
  
Sub RangeSel3()
    'Range("B2,B4,D2,D4").Select
    Range("B2:D3,B5:D6").Select
End Sub


'----------------------------
'範例27
'選取定義名稱的儲存格
'----------------------------
  
Sub RangeSel4()
    Range("營業額總計").Select
End Sub


'-------------------
'範例28
'選取行/列
'-------------------
  
Sub RangeSel5()
    Range("1:1").Select
    'Range("A:A").Select
    'Range("1:3").Select
    'Range("A:C").Select
    'Range("1:3,6:6").Select
    'Range("A:C,F:F").Select
End Sub


'-----------------------------------
'範例29
'使用Cells屬性選取單一儲存格
'-----------------------------------

Sub CellsSel1()
    Cells(5, 3).Activate
    'Cells(5, "C").Activate
End Sub

'-----------------------
'範例30
'以編號選取儲存格
'-----------------------
  
Sub CellsSel2()
    Cells(1027).Activate
End Sub


'-------------------------------
'範例31
'使用Cells屬性選取所有儲存格
'-------------------------------
  
Sub CellsSel3()
    Cells.Select
End Sub


'---------------------------------
'範例32
'使用Cells屬性選取儲存格範圍
'---------------------------------
  
Sub CellsSel4()
    Range(Cells(1, 2), Cells(5, 4)).Select
End Sub
