Attribute VB_Name = "Module1"
'------------------------------
'依據會員編號排序列印會員名冊
'------------------------------
Sub 列印會員名冊()
Attribute 列印會員名冊.VB_Description = "將編號排序後，顯示列印的預覽視窗。"
Attribute 列印會員名冊.VB_ProcData.VB_Invoke_Func = "e\n14"
    Application.ScreenUpdating = False
    Worksheets("會員名冊").Activate
    If Worksheets("會員名冊").AutoFilterMode = True Then
        Range("A3").AutoFilter
    End If
        
    Range("A3").SortSpecial Key1:=Range("A3"), Header:=xlGuess
    'ActiveSheet.PrintOut
    ActiveSheet.PrintPreview
End Sub

'---------------------
'列出男性會員名單
'---------------------
Sub 列出男性會員名單()
    Application.ScreenUpdating = False
    Worksheets("會員名冊").Activate
    Range("A3").AutoFilter Field:=5, Criteria1:="1"
    Range("A3").SortSpecial Key1:=Range("A3"), Header:=xlGuess
End Sub

'---------------------
'列出女性會員名單
'---------------------
Sub 列出女性會員名單()
    Application.ScreenUpdating = False
    Worksheets("會員名冊").Activate
    Range("A3").AutoFilter Field:=5, Criteria1:="2"
    Range("A3").SortSpecial Key1:=Range("A3"), Header:=xlGuess
End Sub
