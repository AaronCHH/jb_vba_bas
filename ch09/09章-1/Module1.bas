Attribute VB_Name = "Module1"

'--------------------------
'�d��56
'����x�s��ϰ�
'--------------------------

Sub SelActRange()
    Range("C4").CurrentRegion.Select
End Sub


'-------------------------
'�d��57
'�����Ʈw
'-------------------------

Sub SelDatabase()
    Range("A3").CurrentRegion.Select
End Sub


'-------------------------
'�d��58
'�C�L��Ʈw
'-------------------------

Sub PrintDatabase()
    Range("A3").CurrentRegion.Select
    ActiveWorkbook.Names.Add Name:="�|��", RefersToR1C1:=Selection
    ActiveSheet.PageSetup.PrintArea = "�|��"
    'ActiveSheet.PrintOut
    ActiveSheet.PrintPreview
End Sub


'------------------------
'�d��59
'�d�߸�Ƶ���
'------------------------

Sub CountDatabase()
    MsgBox Range("A3").CurrentRegion.Rows.Count - 1
End Sub


'--------------------------------
'�d��60
'�]�w�x�s��ϰ�~��
'--------------------------------

Sub LineDatabase()
    Range("A3").CurrentRegion.BorderAround Weight:=xlThick
End Sub

