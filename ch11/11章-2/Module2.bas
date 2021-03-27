Attribute VB_Name = "Module2"
Option Explicit


'------------------------------------------
'範例90
'將空白儲存格前的儲存格文字變更為粗體字
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


'------------------------------------------
'範例91
'將空白儲存格前的儲存格文字變更為斜體字
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

'---------------------------------------
'範例92
'在空白儲存格間輸入「ABC」
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


'---------------------------------------
'範例93
'在空白儲存格間輸入「DEF」
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
