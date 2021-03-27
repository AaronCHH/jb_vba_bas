Attribute VB_Name = "Module1"

'--------------------------
'範例56
'選取儲存格區域
'--------------------------

Sub SelActRange()
    Range("C4").CurrentRegion.Select
End Sub


'-------------------------
'範例57
'選取資料庫
'-------------------------

Sub SelDatabase()
    Range("A3").CurrentRegion.Select
End Sub


'-------------------------
'範例58
'列印資料庫
'-------------------------

Sub PrintDatabase()
    Range("A3").CurrentRegion.Select
    ActiveWorkbook.Names.Add Name:="會員", RefersToR1C1:=Selection
    ActiveSheet.PageSetup.PrintArea = "會員"
    'ActiveSheet.PrintOut
    ActiveSheet.PrintPreview
End Sub


'------------------------
'範例59
'查詢資料筆數
'------------------------

Sub CountDatabase()
    MsgBox Range("A3").CurrentRegion.Rows.Count - 1
End Sub


'--------------------------------
'範例60
'設定儲存格區域外框
'--------------------------------

Sub LineDatabase()
    Range("A3").CurrentRegion.BorderAround Weight:=xlThick
End Sub

