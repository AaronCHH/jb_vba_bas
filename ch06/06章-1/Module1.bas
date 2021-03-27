Sub ±qVBE°õ¦æµ{§Ç()
    Range("A1:A5").Select
    Selection.Copy
    Range("A7").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub
