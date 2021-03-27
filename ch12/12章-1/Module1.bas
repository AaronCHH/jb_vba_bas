Attribute VB_Name = "Module1"
Option Explicit

'------------------------------------
'�d��94
'�ϥ�MsgBox��ƽT�{�R�����
'------------------------------------

Sub ClearAllData()
    Dim myBtn As Integer
    Dim myMsg As String, myTitle As String

    myMsg = "�R���Ҧ���ơH"
    myTitle = "�T�{�R�����"

    myBtn = MsgBox(myMsg, vbYesNo + vbExclamation, myTitle)
                            
    If myBtn = vbYes Then
        Worksheets("Sheet1").Activate
        Cells.ClearContents
        Range("E1") = "�|���W�U"
        Range("A2") = "�s��"
        Range("B2") = "�|���m�W"
        Range("C2") = "��}"
        Range("D2") = "TEL"
        Range("E2") = "�ʧO"
        Range("F2") = "�J�|��"
    End If
End Sub


'------------------------------------
'�d��95
'�ϥ�Input��k��J�C�L�ƶq
'------------------------------------
  
Sub PrintMember()
    Dim myCopy As Integer
    Dim myMsg As String, myTitle As String
    
    myMsg = "�г]�w�C�L�ƶq"
    myTitle = "�C�L�|���W�U"
    myCopy = Application.InputBox(Prompt:=myMsg, Title:=myTitle, _
        Default:=1, Type:=1)

    If myCopy <> 0 Then
        Worksheets("Sheet2").PrintOut Copies:=myCopy
    Else
        MsgBox "�C�L����"
    End If
End Sub


'--------------------------------------
'�d��96
'�ϥ�InputBox��k��J�|���s��
'--------------------------------------
  
Sub SearchMember()
    Dim myCode As Variant
    
    myCode = Application.InputBox("����J�Ȥ�s��", "�d�߫Ȥ�s��")
    
    If myCode <> False Then
        Worksheets("Sheet2").Activate
        Range("A1").AutoFilter Field:=1, Criteria1:=myCode
    End If
End Sub


'----------------------------------
'�d��97
'�C�L�ƹ����w���x�s��d��
'----------------------------------

Sub PrintRange()
    Dim myCell As Range
    Dim myMsg As String, myTitle As String
    
    Worksheets("Sheet3").Activate
    myMsg = "�Щ즲�ƹ��A���w�C�L�d��"
    myTitle = "�]�w�C�L�d��"
    
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
'�d��98
'InputBox��ƪ��d��
'----------------------------------

Sub VBInputBox()
    Dim myNo As Integer
    Dim myMsg As String, myTitle As String
    
    myMsg = "����J���R�����s��"
    myTitle = "�R���P��O��"
    myNo = Val(InputBox(Prompt:=myMsg, Title:=myTitle))

    If myNo <> 0 Then
        MsgBox myNo & "���P��O���N�R��"
    Else
        MsgBox "�פ�B�z�{��"
    End If
End Sub


'----------------------------------
'�d��99
'InputBox��k���d��
'----------------------------------

Sub ExcelVBAInputBox()
    Dim myNo As Integer
    Dim myMsg As String, myTitle As String
    
    myMsg = "����J���R�����s��"
    myTitle = "�R���P��O��"
    myNo = Application.InputBox(Prompt:=myMsg, Title:=myTitle, Type:=1)

    If myNo <> 0 Then
        MsgBox myNo & "���P��O���N�R��"
    Else
        MsgBox "�פ�B�z�{��"
    End If
End Sub

