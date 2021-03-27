Attribute VB_Name = "Module1"

'-------------------------
'範例65
'建立資料庫清單
'-------------------------

Sub SetList()
    Range("A3").Select
    ActiveSheet.ListObjects.Add
    
End Sub

'-------------------------
'範例66
'設定清單名稱
'-------------------------

Sub SetListName()
    Range("A3").Select
    ActiveSheet.ListObjects.Add.Name = "會員名冊"
End Sub

'-------------------------
'範例67
'移除清單
'-------------------------

Sub ChangeUnList()
    ActiveSheet.ListObjects(1).Unlist
'    ActiveSheet.ListObjects("會員名冊").Unlist
    
End Sub

