Attribute VB_Name = "Module1"
'--------------------------
'�ܧ��x�s��I���⪺����
'--------------------------
Sub �ܧ��C��()
    Range("A1:B5").Select           '����x�s��
    With Selection.Interior         '�]�w�I����
        .ColorIndex = 6
        .Pattern = xlSolid
    End With
End Sub


Sub �ƻs���()
'----------------------------
'�ƻs�x�s���ƪ�����
'----------------------------
    Range("A1:A5").Select           '����x�s��
    Selection.Copy                  '�ƻs����d�򤤪����
    Range("B1:B5").Select           '����x�s��
    ActiveSheet.Paste               '�N�ƻs����ƶK�W������x�s�椤
    Application.CutCopyMode = False '�Ѱ��ƻs�����A
End Sub
