Attribute VB_Name = "Module2"
Option Explicit


'------------------------------------------
'�d��90
'�N�ť��x�s��e���x�s���r�ܧ󬰲���r
'
'(�Y�x�s��A1�O�ťժ��h�L�k����)
'------------------------------------------
  
Sub FontBold()
    Range("A1").Select
    
    Do Until ActiveCell.Value = ""
        ActiveCell.Font.Bold = True
        ActiveCell.Offset(1).Select
    Loop
End Sub


'------------------------------------------
'�d��91
'�N�ť��x�s��e���x�s���r�ܧ󬰱���r
'
'(�Y�x�s��A1�O�ťժ��h�L�k����)
'------------------------------------------
  
Sub FontItalic()
    Range("A1").Select
    
    Do
        ActiveCell.Font.Italic = True
        ActiveCell.Offset(1).Select
    Loop Until ActiveCell.Value = ""
End Sub

'----------------------------------
'�٭��x�s��A1��A7���榡
'----------------------------------

Sub FontFormatClear()
    Range("A1").Select
    
    Do
        ActiveCell.ClearFormats
        ActiveCell.Offset(1).Select
    Loop Until ActiveCell.Value = ""
End Sub

'---------------------------------------
'�d��92
'�b�ť��x�s�涡��J�uABC�v
'
'(�Y�x�s��B1���O�ťիh�L�k����)
'---------------------------------------
  
Sub WriteABC()
    Range("B1").Select
    
    Do While ActiveCell.Value = ""
        ActiveCell.Value = "ABC"
        ActiveCell.Offset(1).Select
    Loop
End Sub


'---------------------------------------
'�d��93
'�b�ť��x�s�涡��J�uDEF�v
'
'(�x�s��B1���O�ťե�i����)
'---------------------------------------
  
Sub WriteDEF()
    Range("B1").Select
    
    Do
        ActiveCell.Value = "DEF"
        ActiveCell.Offset(1).Select
    Loop While ActiveCell.Value = ""
End Sub
