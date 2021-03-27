Attribute VB_Name = "Module1"
Option Explicit

'---------------------------------------------
'範例85
'使用Select Case陳述式簡化條件判斷式
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


'-----------------------------------------------
'範例86
'使用For...Next陳述式顯示十次訊息
'-----------------------------------------------

Sub TenMessages()
    Dim i As Integer
    
    For i = 1 To 10
        MsgBox "顯示十次訊息"
    Next i
End Sub


'-----------------------------------------------
'範例87
'使用For...Next陳述式變更工作表名稱
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


'--------------------------------------------------------
'範例88
'使用For Each...Next陳述式刪除特定工作表
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

'----------------------------------------------------------
'範例89
'使用For Each...Next陳述式變更特定儲存格的背景色
'----------------------------------------------------------
  
Sub InteriorBlue()
    Dim myRange As Range

    For Each myRange In Worksheets(2).Range("A1:D10")
        If myRange.Value >= 70 Then myRange.Interior.ColorIndex = 5
    Next
End Sub
