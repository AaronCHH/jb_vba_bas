# Ch11 VBA 的控制功能
## 11章-1

## 範例79 變更系統日期
```
Attribute VB_Name = "Module1"
Option Explicit
'------------------------------------
'範例79 變更系統日期
'
'(變更後請回復現在的日期)
'------------------------------------

Sub ChangeDate()
    Date = #3/26/1985#
End Sub

```
## 範例80 使用Int函數捨去小數位數
```
'-----------------------------------
'範例80 使用Int函數捨去小數位數
'-----------------------------------

Sub IntA1()
    Range("A1").Value = 123.456
    Range("A2").Value = Int(Range("A1").Value)
End Sub

```
## 範例81 Excel VBA對工作表函數的運用
```
'----------------------------------------
'範例81 Excel VBA對工作表函數的運用
'----------------------------------------

Sub SearchMax()
    Dim myMax As Long
    
    myMax = Application.WorksheetFunction.Max(Range("B1:D10").Value)
    MsgBox "最大值是：" & myMax & "。"
End Sub


```
## 範例82 巢狀With陳述式
```
Attribute VB_Name = "Module2"
Option Explicit
'-----------------------------
'範例82 巢狀With陳述式
'-----------------------------

Sub TestWith()
    With Range("B2")
        .Value = "客戶編號"
        With .Font
            .Name = "標楷體"
            .Bold = True
            .Size = 12
        End With
    End With
End Sub

```
## 範例83 顯示/隱藏列
```
'--------------------------
'範例83 顯示/隱藏列
'--------------------------

Sub ToggleColumn()
    With Columns("C")
        .Hidden = Not .Hidden
    End With
End Sub

```
## 範例84 使用IF陳述式判斷多個條件
```
'----------------------------------
'範例84 使用IF陳述式判斷多個條件
'----------------------------------
  
Sub TestIf()
    If Range("A1") = "特" Then
        MsgBox "您是高級會員"
    ElseIf Range("A1") = "正" Then
        MsgBox "您是普通會員"
    ElseIf Range("A1") = "準" Then
        MsgBox "您是預備會員"
    Else
        MsgBox "請鍵入會員類別"
    End If
End Sub
```


## 11章-2
## 範例85 使用Select Case陳述式簡化條件判斷式
```
Attribute VB_Name = "Module1"
Option Explicit
'---------------------------------------------
'範例85 使用Select Case陳述式簡化條件判斷式
'---------------------------------------------
  
Sub TestResult()
    Select Case Range("A1")
        Case Is > 80
            Range("B1").Value = "特優"
        Case Is > 60
            Range("B1").Value = "良"
        Case Is > 40
            Range("B1").Value = "不及格"
        Case Else
            Range("B1").Value = "重修"
    End Select
End Sub

```
## 範例86 使用For...Next陳述式顯示十次訊息
```
'-----------------------------------------------
'範例86 使用For...Next陳述式顯示十次訊息
'-----------------------------------------------

Sub TenMessages()
    Dim i As Integer
    
    For i = 1 To 10
        MsgBox "顯示十次訊息"
    Next i
End Sub

```
## 範例87 使用For...Next陳述式變更工作表名稱
```
'-----------------------------------------------
'範例87 使用For...Next陳述式變更工作表名稱
'-----------------------------------------------

Sub NameWorkSheets()
    Dim i As Integer
    Dim myWSCnt As Integer
    
    myWSCnt = Worksheets.Count
    
    For i = 1 To myWSCnt
        Worksheets(i).Name = "WORK" & i
    Next i
End Sub


'----------------------------------
'回復工作表名稱
'----------------------------------

Sub NameWorkSheets2()
    Dim i As Integer
    Dim myWSCnt As Integer
    
    myWSCnt = Worksheets.Count
    
    For i = 1 To myWSCnt
        Worksheets(i).Name = "Sheet" & i
    Next i
End Sub

```
## 範例88 使用For Each...Next陳述式刪除特定工作表
```
'--------------------------------------------------------
'範例88 使用For Each...Next陳述式刪除特定工作表
'--------------------------------------------------------

Sub DeleteSheet()
    Dim mySheet As Worksheet
    
    For Each mySheet In Worksheets
        If mySheet.Name = "Sheet4" Then
            mySheet.Delete
            Exit For
        End If
    Next mySheet
End Sub
```
## 範例89 使用For Each...Next陳述式變更特定儲存格的背景色
```
'----------------------------------------------------------
'範例89 使用For Each...Next陳述式變更特定儲存格的背景色
'----------------------------------------------------------
  
Sub InteriorBlue()
    Dim myRange As Range

    For Each myRange In Worksheets(2).Range("A1:D10")
        If myRange.Value >= 70 Then myRange.Interior.ColorIndex = 5
    Next
End Sub

Attribute VB_Name = "Module2"
Option Explicit

```
## 範例90 將空白儲存格前的儲存格文字變更為粗體字
```
'------------------------------------------
'範例90 將空白儲存格前的儲存格文字變更為粗體字
'
'(若儲存格A1是空白的則無法執行)
'------------------------------------------
  
Sub FontBold()
    Range("A1").Select
    
    Do Until ActiveCell.Value = ""
        ActiveCell.Font.Bold = True
        ActiveCell.Offset(1).Select
    Loop
End Sub

```
## 範例91 將空白儲存格前的儲存格文字變更為斜體字
```
'------------------------------------------
'範例91 將空白儲存格前的儲存格文字變更為斜體字
'
'(若儲存格A1是空白的則無法執行)
'------------------------------------------
  
Sub FontItalic()
    Range("A1").Select
    
    Do
        ActiveCell.Font.Italic = True
        ActiveCell.Offset(1).Select
    Loop Until ActiveCell.Value = ""
End Sub

'----------------------------------
'還原儲存格A1到A7的格式
'----------------------------------

Sub FontFormatClear()
    Range("A1").Select
    
    Do
        ActiveCell.ClearFormats
        ActiveCell.Offset(1).Select
    Loop Until ActiveCell.Value = ""
End Sub
```
## 範例92 在空白儲存格間輸入「ABC」
```
'---------------------------------------
'範例92 在空白儲存格間輸入「ABC」
'
'(若儲存格B1不是空白則無法執行)
'---------------------------------------
  
Sub WriteABC()
    Range("B1").Select
    
    Do While ActiveCell.Value = ""
        ActiveCell.Value = "ABC"
        ActiveCell.Offset(1).Select
    Loop
End Sub

```
## 範例93 在空白儲存格間輸入「DEF」
```
'---------------------------------------
'範例93 在空白儲存格間輸入「DEF」
'
'(儲存格B1不是空白亦可執行)
'---------------------------------------
  
Sub WriteDEF()
    Range("B1").Select
    
    Do
        ActiveCell.Value = "DEF"
        ActiveCell.Offset(1).Select
    Loop While ActiveCell.Value = ""
End Sub
```