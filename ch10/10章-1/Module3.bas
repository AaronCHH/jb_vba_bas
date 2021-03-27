Attribute VB_Name = "Module3"
Option Explicit

    Dim myModuleNo As Integer

'------------------------------------
'程序等級變數有效範圍測試
'
'(請重覆執行)
'------------------------------------

Sub NumberAdd1()
    Dim myProcedureNo As Integer

    myProcedureNo = myProcedureNo + 10
    MsgBox myProcedureNo
End Sub

'---------------------------------
'模組等級變數有效範圍測試
'
'(請重覆執行)
'---------------------------------

Sub NumberAdd2()
    myModuleNo = myModuleNo + 10
    MsgBox myModuleNo
End Sub

