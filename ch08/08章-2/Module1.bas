Attribute VB_Name = "Module1"

'------------------------------------------
'�d��41
'���g�ŶKï�ƻs���
'------------------------------------------

Sub CopyData()
    Worksheets("Sheet1").Range("A1:B10").Copy _
        Destination:=Worksheets("Sheet2").Range("A1")
End Sub

'-------------------------
'�d��42
'�M���x�s�檺�p�⦡�P��
'-------------------------
  
Sub ClearRange1()
    Range("A1").Select
    ActiveCell.Value = ""
End Sub
 
Sub ClearRange2()
    Range("A1:D5").Select
    Selection.ClearContents
End Sub


'---------------------
'�d��43
'�M���x�s�檺�榡
'---------------------

Sub ClearRange3()
    Worksheets("Sheet3").Range("A10:D12").ClearFormats
End Sub


'----------------------------
'�d��44
'�M���x�s���ƪ���ƻP�榡
'----------------------------
  
Sub ClearRange4()
    Worksheets("Sheet3").Range("A10:D12").Clear
End Sub
