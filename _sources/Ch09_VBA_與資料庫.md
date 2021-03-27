# Ch09 VBA 與資料庫
## 09章-1
## 範例56 選取儲存格區域
```
Attribute VB_Name = "Module1"
'--------------------------
'範例56 選取儲存格區域
'--------------------------

Sub SelActRange()
    Range("C4").CurrentRegion.Select
End Sub

```
## 範例57 選取資料庫
```
'-------------------------
'範例57 選取資料庫
'-------------------------

Sub SelDatabase()
    Range("A3").CurrentRegion.Select
End Sub

```
## 範例58 列印資料庫
```
'-------------------------
'範例58 列印資料庫
'-------------------------

Sub PrintDatabase()
    Range("A3").CurrentRegion.Select
    ActiveWorkbook.Names.Add Name:="會員", RefersToR1C1:=Selection
    ActiveSheet.PageSetup.PrintArea = "會員"
    'ActiveSheet.PrintOut
    ActiveSheet.PrintPreview
End Sub

```
## 範例59 查詢資料筆數
```
'------------------------
'範例59 查詢資料筆數
'------------------------

Sub CountDatabase()
    MsgBox Range("A3").CurrentRegion.Rows.Count - 1
End Sub

```
## 範例60 設定儲存格區域外框
```
'--------------------------------
'範例60 設定儲存格區域外框
'--------------------------------

Sub LineDatabase()
    Range("A3").CurrentRegion.BorderAround Weight:=xlThick
End Sub


Attribute VB_Name = "Module2"
```
## 範例61 取得末端儲存格
```
'-------------------------
'範例61 取得末端儲存格
'-------------------------

Sub SelEndCell()
    Range("C1").End(xlDown).Select
    'Range("C2").End(xlDown).Select
    'Range("C3").End(xlDown).Select
End Sub

```
## 範例62 移至最後一筆資料處
```
'--------------------------------------
'範例62 移至最後一筆資料處
'--------------------------------------

Sub SelLastCell()
    Range("A3").End(xlDown).Select
    Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Select
End Sub

```
## 範例63 移至新建資料的目的儲存格
```
'--------------------------------
'範例63 移至新建資料的目的儲存格
'--------------------------------

Sub SelNewCell()
    Range("A65536").End(xlUp).Offset(1).Select
End Sub

```
## 範例64 選取特定資料
```
'-------------------------
'範例64 選取特定資料
'-------------------------

Sub SelRecords()
    Range("A5", Range("A5").End(xlToRight)).Select
End Sub



## 09章-2

Attribute VB_Name = "Module1"
```
## 範例65 建立資料庫清單
```
'-------------------------
'範例65 建立資料庫清單
'-------------------------

Sub SetList()
    Range("A3").Select
    ActiveSheet.ListObjects.Add
    
End Sub
```
## 範例66 設定清單名稱
```
'-------------------------
'範例66 設定清單名稱
'-------------------------

Sub SetListName()
    Range("A3").Select
    ActiveSheet.ListObjects.Add.Name = "會員名冊"
End Sub
```
## 範例67 移除清單
```
'-------------------------
'範例67 移除清單
'-------------------------

Sub ChangeUnList()
    ActiveSheet.ListObjects(1).Unlist
'    ActiveSheet.ListObjects("會員名冊").Unlist
    
End Sub

```
## 範例68 選取清單內所有資料
```
Attribute VB_Name = "Module2"
'-------------------------
'範例68 選取清單內所有資料
'-------------------------

Sub SelList()
    ActiveSheet.ListObjects(1).Range.Select
    
End Sub
```
## 範例69 選取清單標籤所屬的行
```
'-------------------------
'範例69 選取清單標籤所屬的行
'-------------------------

Sub SelListHeader()
    ActiveSheet.ListObjects(1).HeaderRowRange.Select
    
End Sub
```
## 範例70 選取清單中資料的部份
```
'-------------------------
'範例70 選取清單中資料的部份
'-------------------------

Sub SelListBody()
    ActiveSheet.ListObjects(1).DataBodyRange.Select
    
End Sub
```
## 範例71 選取插入資料的行
```
'-------------------------
'範例71 選取插入資料的行
'-------------------------

Sub SelListNewRow()
    Range("A3").Select
    Selection.ListObject.InsertRowRange.Select
End Sub

Sub SelListNewRow2()
    Dim myRowRnage As Range
    
    Set myRowRnage = ActiActiveSheet.ListObjects(1).InsertRowRangeveSheet.ListObjects(1).InsertRowRange
    If myRowRnage Is Nothing Then
        MsgBox "請選取清單"
    Else
        myRowRnage.Select
    End If
End Sub

Attribute VB_Name = "Module3"
```
## 範例72 取得清單中資料筆數
```
'-------------------------
'範例72 取得清單中資料筆數
'-------------------------

Sub CountListData()
    MsgBox ActiveSheet.ListObjects(1).ListRows.Count
    
End Sub
```
## 範例73 選取清單中特定的行/列
```
'-------------------------
'範例73 選取清單中特定的行/列
'-------------------------

Sub SelListRow()
    ActiveSheet.ListObjects(1).ListRows(3).Range.Select
'    ActiveSheet.ListObjects(1).ListColumns(3).Range.Select
    
End Sub

```
## 範例74 在清單中插入行
```
'-------------------------
'範例74 在清單中插入行
'-------------------------

Sub InsertListRow()
    ActiveSheet.ListObjects(1).ListRows.Add (2)
    
End Sub
```
## 範例75 列印清單
```
'-------------------------
'範例75 列印清單
'-------------------------

Sub PrintList()
'    ActiveSheet.ListObjects(1).Range.PrintOut
    ActiveSheet.ListObjects(1).Range.PrintPreview
    
End Sub

