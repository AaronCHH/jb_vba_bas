Attribute VB_Name = "Module2"
Option Explicit

'----------------------
'�d��77
'�p��u�@���ƶq
'----------------------

Sub DisplayWSCnt()
    Dim myWSCnt As Integer
    
    myWSCnt = ActiveWorkbook.Worksheets.Count
    MsgBox myWSCnt
End Sub


'----------------------------------------
'�d��78
'�ϥΪ����ܼƪ��{��
'
'�Х��}��Dummy.xls��A����
'----------------------------------------


Sub SetObject()
    Dim myWSheet As Worksheet
        
    Set myWSheet = Workbooks("Dummy.xls").Worksheets("Sheet2")
    
    myWSheet.Range("A1:D10").Value = "ABC"
End Sub
