Attribute VB_Name = "Module2"

'------------------------------------------
'�d��8
'�N����ï�s�ɫ�����
'
'(�ж}�ҡuDummy.xls�v��A����)
'------------------------------------------

Sub ��������ï5()
    Workbooks("Dummy.xls").Close SaveChanges:=True '���w�޼ƪ��W��
End Sub

'------------------------------------------
'�d��9
'����ï�����ɤ��s��
'
'(�ж}�ҡuDummy.xls�v��A����)
'------------------------------------------

Sub ��������ï6()
    Workbooks("Dummy.xls").Close False              '�зǤ޼�
End Sub

'------------------------------------------
'�d��10
'���w�ϥΤ�������ï
'
'(�ж}�ҡuDummy.xls�v��A����)
'------------------------------------------

Sub ���w����ï()
Attribute ���w����ï.VB_ProcData.VB_Invoke_Func = " \n14"
    Workbooks("Dummy.xls").Activate
End Sub

'----------------------------------
'�d��11
'�N����ï�s��
'
'----------------------------------
 
Sub ����ï�s��()
    ActiveWorkbook.Save             '�x�s�ϥΤ�����ï
End Sub

Sub ����ï�s��2()
    Workbooks("Dummy.xls").Save    '���w�s�ɦW�٫��x�s
End Sub

'--------------------------------------------
'�d��12
'�x�s����ï�ɥt�s�s��'
'
'---------------------------------------------

Sub �x�s����ï3()
Attribute �x�s����ï3.VB_ProcData.VB_Invoke_Func = " \n14"
    '���w�x�s�ؼ�
    ActiveWorkbook.SaveAs Filename:="C:\Excel2003VBA��¦�g\Test.xls"
End Sub

Sub �x�s����ï4()
Attribute �x�s����ï4.VB_ProcData.VB_Invoke_Func = " \n14"
    '�s���ثe��Ƨ���
    ActiveWorkbook.SaveAs Filename:="Test.xls"
End Sub
