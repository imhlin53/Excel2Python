Attribute VB_Name = "Modul4"
Sub move()
Attribute move.VB_ProcData.VB_Invoke_Func = " \n14"
'
' move Makro
'

'
    ActiveSheet.Shapes.Range(Array("DelSteckplz")).Select
    ActiveSheet.Shapes("DelSteckplz").IncrementLeft 432
    ActiveSheet.Shapes("DelSteckplz").IncrementTop -125.25
    Range("CE7").Select
    ActiveSheet.Shapes.Range(Array("DelSteckplz")).Select
    ActiveSheet.Shapes("DelSteckplz").IncrementLeft 157.5
    ActiveSheet.Shapes("DelSteckplz").IncrementTop 51.75
    Range("CO5").Select
End Sub
