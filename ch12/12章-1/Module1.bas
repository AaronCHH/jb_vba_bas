Attribute VB_Name = "Module1"
Option Explicit

'------------------------------------
'範例94
'使用MsgBox函數確認刪除資料
'------------------------------------

Sub ClearAllData()
    Dim myBtn As Integer
    Dim myMsg As String, myTitle As String

    myMsg = "刪除所有資料？"
    myTitle = "確認刪除資料"

    myBtn = MsgBox(myMsg, vbYesNo + vbExclamation, myTitle)
                            
    If myBtn = vbYes Then
        Worksheets("Sheet1").Activate
        Cells.ClearContents
        Range("E1") = "會員名冊"
        Range("A2") = "編號"
        Range("B2") = "會員姓名"
        Range("C2") = "住址"
        Range("D2") = "TEL"
        Range("E2") = "性別"
        Range("F2") = "入會日"
    End If
End Sub


'------------------------------------
'範例95
'使用Input方法鍵入列印數量
'------------------------------------
  
Sub PrintMember()
    Dim myCopy As Integer
    Dim myMsg As String, myTitle As String
    
    myMsg = "請設定列印數量"
    myTitle = "列印會員名冊"
    myCopy = Application.InputBox(Prompt:=myMsg, Title:=myTitle, _
        Default:=1, Type:=1)

    If myCopy <> 0 Then
        Worksheets("Sheet2").PrintOut Copies:=myCopy
    Else
        MsgBox "列印取消"
    End If
End Sub


'--------------------------------------
'範例96
'使用InputBox方法鍵入會員編號
'--------------------------------------
  
Sub SearchMember()
    Dim myCode As Variant
    
    myCode = Application.InputBox("請鍵入客戶編號", "查詢客戶編號")
    
    If myCode <> False Then
        Worksheets("Sheet2").Activate
        Range("A1").AutoFilter Field:=1, Criteria1:=myCode
    End If
End Sub


'----------------------------------
'範例97
'列印滑鼠指定的儲存格範圍
'----------------------------------

Sub PrintRange()
    Dim myCell As Range
    Dim myMsg As String, myTitle As String
    
    Worksheets("Sheet3").Activate
    myMsg = "請拖曳滑鼠，指定列印範圍"
    myTitle = "設定列印範圍"
    
    On Error Resume Next
    Set myCell = Application.InputBox(Prompt:=myMsg, Title:=myTitle, _
        Type:=8)
    If myCell Is Nothing Then Exit Sub
    
    With ActiveSheet
        .PageSetup.PrintArea = myCell.Address
        .PrintOut
    End With
End Sub


'----------------------------------
'範例98
'InputBox函數的範例
'----------------------------------

Sub VBInputBox()
    Dim myNo As Integer
    Dim myMsg As String, myTitle As String
    
    myMsg = "請鍵入欲刪除的編號"
    myTitle = "刪除銷售記錄"
    myNo = Val(InputBox(Prompt:=myMsg, Title:=myTitle))

    If myNo <> 0 Then
        MsgBox myNo & "號銷售記錄將刪除"
    Else
        MsgBox "終止處理程序"
    End If
End Sub


'----------------------------------
'範例99
'InputBox方法的範例
'----------------------------------

Sub ExcelVBAInputBox()
    Dim myNo As Integer
    Dim myMsg As String, myTitle As String
    
    myMsg = "請鍵入欲刪除的編號"
    myTitle = "刪除銷售記錄"
    myNo = Application.InputBox(Prompt:=myMsg, Title:=myTitle, Type:=1)

    If myNo <> 0 Then
        MsgBox myNo & "號銷售記錄將刪除"
    Else
        MsgBox "終止處理程序"
    End If
End Sub

