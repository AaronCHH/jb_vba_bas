Attribute VB_Name = "Module1"
Option Explicit

'---------------------------------------------
'�d��85
'�ϥ�Select Case���z��²�Ʊ���P�_��
'---------------------------------------------
  
Sub TestResult()
    Select Case Range("A1")
        Case Is > 80
            Range("B1").Value = "�S�u"
        Case Is > 60
            Range("B1").Value = "�}"
        Case Is > 40
            Range("B1").Value = "���ή�"
        Case Else
            Range("B1").Value = "����"
    End Select
End Sub


'-----------------------------------------------
'�d��86
'�ϥ�For...Next���z����ܤQ���T��
'-----------------------------------------------

Sub TenMessages()
    Dim i As Integer
    
    For i = 1 To 10
        MsgBox "��ܤQ���T��"
    Next i
End Sub


'-----------------------------------------------
'�d��87
'�ϥ�For...Next���z���ܧ�u�@��W��
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
'�^�_�u�@��W��
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
'�d��88
'�ϥ�For Each...Next���z���R���S�w�u�@��
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
'�d��89
'�ϥ�For Each...Next���z���ܧ�S�w�x�s�檺�I����
'----------------------------------------------------------
  
Sub InteriorBlue()
    Dim myRange As Range

    For Each myRange In Worksheets(2).Range("A1:D10")
        If myRange.Value >= 70 Then myRange.Interior.ColorIndex = 5
    Next
End Sub
