Attribute VB_Name = "Modul5"
Sub insert_row()
Attribute insert_row.VB_ProcData.VB_Invoke_Func = " \n14"
'
' insert_row Makro
'

'
    Rows("19:19").Select
    Range("B19:HV19").Select
    Selection.Copy
    Range("A19").Select
    Application.CutCopyMode = False
    Rows("19:19").Select
    Rows("19:19").Select
    Range("BM5").Select
End Sub
