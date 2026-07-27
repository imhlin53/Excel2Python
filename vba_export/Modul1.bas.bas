Attribute VB_Name = "Modul1"
' Initialisiert die Zellen der Stromwerte
Public Sub Set_Current_Cells(ByVal o_Target As Range, b_FormatOnly As Boolean)
    Dim o_PinSheet As Worksheet
    Dim s_PinChar As String
    Dim o_PinCharCell As Range

    ' Ist der Pintyp leer?
    If o_Target.Value = "" Then
        ' Das Sheet mit den Pintypinformationen ermitteln
        Set o_PinSheet = Sheets("Pinchar0")
    Else
        ' Das Sheet mit den Pintypinformationen ermitteln
        Set o_PinSheet = Get_Pin_Sheet(o_Target.Value)
    End If
    
    ' Pincharakteristik lesen
    Set o_PinCharCell = Cells(o_Target.Row, 65)
    s_PinChar = o_PinCharCell.Value
    
    ' Feld für die Pincharakteristik bei Bedarf ausgrauen
    If (o_PinSheet.Cells(5, 1).Value = "") Then
        o_PinCharCell.Interior.Color = RGB(216, 216, 216)
    Else
        o_PinCharCell.Interior.Color = RGB(204, 255, 255)
    End If
    
    ' Initialisiere alle Stromwert Zellen
    For i_Index = 0 To 7
        Call Set_Current_Cell(s_PinChar, o_Target.Row, i_Index, b_FormatOnly, o_PinSheet)
    Next i_Index
End Sub

' Initialisiert eine Stromwert Zelle
Private Sub Set_Current_Cell(ByVal s_PinChar As String, l_Row As Long, ByVal i_Index As Integer, b_FormatOnly As Boolean, ByVal o_PinSheet As Worksheet)
    Dim l_Column As Long
    Dim o_WorkingCell As Range
    Dim b_Deactivate As Boolean
    
    ' Spalte berechnen (7 Spalten pro Stromfeld) und Zelle ermitteln
    l_Column = 78 + i_Index * 7
    Set o_WorkingCell = Cells(l_Row, l_Column)
    
    '
    ' Zelle initialisieren
    '
    Dim s_Formula As String
    Dim s_Comment As String
            
    ' Kommentar lesen
    s_Comment = o_PinSheet.Cells(2, 2 + i_Index)
    b_Deactivate = (s_Comment = "")
    
    ' Formel lesen
    s_Formula = Get_Formula(s_PinChar, i_Index, o_PinSheet, l_Row)
    
    ' Zelle zurücksetzen
    Call Reset_Cell(o_WorkingCell, b_Deactivate, b_FormatOnly)
    
    ' Wenn ein Kommentar existiert, dann muss die Zelle befüllt werden
    If Not b_Deactivate Then
        ' Kommentar setzen
        If o_WorkingCell.Comment Is Nothing Then
            o_WorkingCell.AddComment
        End If
        o_WorkingCell.Comment.Visible = False
        o_WorkingCell.Comment.Text (s_Comment)
        If Len(s_Comment) > 60 Then
            o_WorkingCell.Comment.Shape.Height = 100
            o_WorkingCell.Comment.Shape.Width = 100
        End If
    End If
    
    ' Wenn eine Formel existiert, dann muss diese ggf. gesetzt werden
    ' (sofern nicht bereits ein Wert in der Zelle steht)
    If (s_Formula <> "") Then
        If (b_FormatOnly) Then
            ' Wenn nur formatiert werden soll (beim Öffnen des Excel Dokuments), dann wird die Formel
            ' nur gesetzt, wenn kein Wert in der Zelle enthalten ist. Damit werden mit Formeln belegte
            ' Felder nach einem Export aus e42 repariert. Ansonsten passiert nichts.
            If (o_WorkingCell.Value = "") Then
                o_WorkingCell.Formula = s_Formula
            End If
        Else
            ' Beim Bearbeiten des Formblatts wird die Formel in jedem Fall gesetzt.
            o_WorkingCell.Formula = s_Formula
        End If
    End If
End Sub

' Sheet mit den Pintypinformationen ermitteln
Private Function Get_Pin_Sheet(ByVal s_PinTyp As String) As Worksheet
    Dim s_PinSheetName As String
    Dim o_PinSheet As Worksheet

    ' Suche Pintyp in "Daten"-Blatt
    For Each o_Cell In Worksheets("Daten").Range("A2", "A32")
        If o_Cell.Value = s_PinTyp Then
            s_PinSheetName = Worksheets("Daten").Cells(o_Cell.Row, 2).Value
            Set Get_Pin_Sheet = Worksheets(s_PinSheetName)
            Exit For
        End If
    Next o_Cell
    'Urban CheckMe zugefügt, 21.03.2023
    If WB_open = False Then
       Tabelle1.CommandButton2.Caption = "CheckMe"
       Tabelle1.CommandButton2.ForeColor = blue
       Range("DataChecked").Value = "ungeprueft"
       Range("State").Value = ""
    End If
End Function

' Setzt eine Stromzelle zurück
Private Sub Reset_Cell(ByVal o_Cell As Range, b_Deactivate As Boolean, b_FormatOnly As Boolean)
    If Not b_FormatOnly Then
        ' Wenn nur formatiert werden soll (beim Öffnen des Excel Dokuments), dann werden keine Werte
        ' und Formeln gelöscht. Ansonsten wird der Wert nur gelöscht, wenn die Zelle deaktiviert wird
        ' oder wenn der Wert auf einer Formel basierte (dann wird auch die Formel gelöscht).
        If (b_Deactivate Or o_Cell.HasFormula) Then
            Call Clear_Connected_Cells(o_Cell)
        End If
    End If
    
    ' Kommentar löschen
    If Not o_Cell.Comment Is Nothing Then
        o_Cell.Comment.Delete
    End If
    
    ' Farbe setzen
    If (b_Deactivate) Then
        o_Cell.Interior.Color = RGB(216, 216, 216)
    Else
        o_Cell.Interior.Color = RGB(204, 255, 255)
    End If
End Sub

' Löscht eine verbundene Zelle
Private Sub Clear_Connected_Cells(ByVal o_Target As Range)
    Dim o_Cell As Range
    
    For Each o_Cell In o_Target
        o_Cell.MergeArea.ClearContents
    Next o_Cell
End Sub

' Prüft, ob das Target in einer Pin-Zeile liegt
Public Function Is_Pins_Row(ByVal o_Target As Range)
    If (o_Target.Row > Range("PinsUpper").Row) And (o_Target.Row < Range("PinsLower").Row) Then
        Is_Pins_Row = True
    Else
        Is_Pins_Row = False
    End If
End Function
' Gibt die Formel für einen Stromwert zurück
Private Function Get_Formula(ByVal s_PinChar As String, ByVal i_Index As Integer, ByVal o_PinSheet As Worksheet, l_Row As Long) As String
    Dim o_Range As Range
    Set o_Range = o_PinSheet.Range(o_PinSheet.Name).Find(s_PinChar, LookIn:=xlValues, LookAt:=xlWhole)
    If o_Range Is Nothing Then
        Get_Formula = ""
    Else
        Dim s_Formula As String
        s_Formula = o_PinSheet.Cells(o_Range.Row, 2 + i_Index)
        s_Formula = Replace(s_Formula, "<row>", l_Row)
        Get_Formula = s_Formula
    End If
End Function

