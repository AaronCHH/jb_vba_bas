Attribute VB_Name = "Module1"
Option Explicit


'------------------------------------
'�d��79
'�ܧ�t�Τ��
'
'(�ܧ��Ц^�_�{�b�����)
'------------------------------------

Sub ChangeDate()
    Date = #3/26/1985#
End Sub


'-----------------------------------
'�d��80
'�ϥ�Int��Ʊ˥h�p�Ʀ��
'-----------------------------------

Sub IntA1()
    Range("A1").Value = 123.456
    Range("A2").Value = Int(Range("A1").Value)
End Sub


'----------------------------------------
'�d��81
'Excel VBA��u�@���ƪ��B��
'----------------------------------------

Sub SearchMax()
    Dim myMax As Long
    
    myMax = Application.WorksheetFunction.Max(Range("B1:D10").Value)
    MsgBox "�̤j�ȬO�G" & myMax & "�C"
End Sub
