Attribute VB_Name = "Module1"
'------------------------------
'�̾ڷ|���s���ƧǦC�L�|���W�U
'------------------------------
Sub �C�L�|���W�U()
Attribute �C�L�|���W�U.VB_Description = "�N�s���Ƨǫ�A��ܦC�L���w�������C"
Attribute �C�L�|���W�U.VB_ProcData.VB_Invoke_Func = "e\n14"
    Application.ScreenUpdating = False
    Worksheets("�|���W�U").Activate
    If Worksheets("�|���W�U").AutoFilterMode = True Then
        Range("A3").AutoFilter
    End If
        
    Range("A3").SortSpecial Key1:=Range("A3"), Header:=xlGuess
    'ActiveSheet.PrintOut
    ActiveSheet.PrintPreview
End Sub

'---------------------
'�C�X�k�ʷ|���W��
'---------------------
Sub �C�X�k�ʷ|���W��()
    Application.ScreenUpdating = False
    Worksheets("�|���W�U").Activate
    Range("A3").AutoFilter Field:=5, Criteria1:="1"
    Range("A3").SortSpecial Key1:=Range("A3"), Header:=xlGuess
End Sub

'---------------------
'�C�X�k�ʷ|���W��
'---------------------
Sub �C�X�k�ʷ|���W��()
    Application.ScreenUpdating = False
    Worksheets("�|���W�U").Activate
    Range("A3").AutoFilter Field:=5, Criteria1:="2"
    Range("A3").SortSpecial Key1:=Range("A3"), Header:=xlGuess
End Sub
