Attribute VB_Name = "Module2"
Option Explicit

'-----------------------------
'�d��82
'�_��With���z��
'-----------------------------

Sub TestWith()
    With Range("B2")
        .Value = "�Ȥ�s��"
        With .Font
            .Name = "�з���"
            .Bold = True
            .Size = 12
        End With
    End With
End Sub


'--------------------------
'�d��83
'���/���æC
'--------------------------

Sub ToggleColumn()
    With Columns("C")
        .Hidden = Not .Hidden
    End With
End Sub


'----------------------------------
'�d��84
'�ϥ�IF���z���P�_�h�ӱ���
'----------------------------------
  
Sub TestIf()
    If Range("A1") = "�S" Then
        MsgBox "�z�O���ŷ|��"
    ElseIf Range("A1") = "��" Then
        MsgBox "�z�O���q�|��"
    ElseIf Range("A1") = "��" Then
        MsgBox "�z�O�w�Ʒ|��"
    Else
        MsgBox "����J�|�����O"
    End If
End Sub
