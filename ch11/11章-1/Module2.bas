Attribute VB_Name = "Module2"
Option Explicit

'-----------------------------
'範例82
'巢狀With陳述式
'-----------------------------

Sub TestWith()
    With Range("B2")
        .Value = "客戶編號"
        With .Font
            .Name = "標楷體"
            .Bold = True
            .Size = 12
        End With
    End With
End Sub


'--------------------------
'範例83
'顯示/隱藏列
'--------------------------

Sub ToggleColumn()
    With Columns("C")
        .Hidden = Not .Hidden
    End With
End Sub


'----------------------------------
'範例84
'使用IF陳述式判斷多個條件
'----------------------------------
  
Sub TestIf()
    If Range("A1") = "特" Then
        MsgBox "您是高級會員"
    ElseIf Range("A1") = "正" Then
        MsgBox "您是普通會員"
    ElseIf Range("A1") = "準" Then
        MsgBox "您是預備會員"
    Else
        MsgBox "請鍵入會員類別"
    End If
End Sub
