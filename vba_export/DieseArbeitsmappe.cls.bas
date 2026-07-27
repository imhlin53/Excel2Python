Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
'Public WB_open As Boolean
Public PinBereich

' Wird beim Öffnen der Exceldatei aufgerufen
Private Sub Workbook_Open()
    WB_open = True
    ' Es werden keine Pins generiert
    b_IsGeneratingPins = False
   
  chkAppl = Application.Visible
    If chkAppl = False Then
     MsgBox "EXCEL Workbook ist: " & chkAppl
     Application.Visible = True
      For Each win In Application.Windows
        If win.Caption <> ThisWorkbook.Name Then
          win.WindowState = xlMaximized
        End If
      Next
    End If
    ' Bereinigen des Formblatts 11.03.2020
   ActiveWorkbook.Worksheets("Formblatt").Visible = True
   ActiveWorkbook.Worksheets("Formblatt").Activate
   Call Fix_After_Open
'UGU CSV Import erlauben ' Anwender ist der meinung der Funktionsumfang ist zu komplex 04.09.2019
BackStop = Cells(1, 33).Value
If BackStop = "x" Or BackStop = "X" Then
        Smooth = MsgBox("Ihre Entscheidung! Der Funktionsumfang ist Ihnen zu komplex" & vbCrLf & vbCrLf & _
        "Begrenzung auf Version 1.4.1 startet jetzt" & vbCrLf & "Der Informationsgehalt der Klemmeninformation reicht Ihnen", vbInformation)
        ReadyDB = "closed"
     Else
      ReadyDB = "open"
      Application.StatusBar = "Lieber Anwender [ " & Environ("Username") & "] " & "Das Öffnen des Formblattes startet den automatischen Dialog zum" & _
      "Import der Datenbasis." & " . . .  Dear User [" & Environ("Username") & "] " & _
      "VOBES Formblatt will start import database dialogue automatically."
      Call Datenbasen_aus_archiv_entpacken
End If
      
      'Urban 25012023 Range("BC12").Value = "initial"
      'Call ExChange 25.01.2022 Urban
      Range("StartPin").Value = "m"
      Range("EndPin").Value = "n"
      Range("GER").Value = "n"
      StartPinS = "AH" & Range("PinsUpper").Row + 1    'erster Pin
      EndPinS = "AH" & Range("PinsLower").Row - 1    'lezter Pin
      PinBereich = StartPinS & ":" & EndPinS
     Cells(5, 10).Select ' Zelle Teilenummer auswählen
     ActiveWorkbook.Names.Add Name:="Pins", RefersToR1C1:="=Formblatt!R17C2:R17C230"
     Call ModulKl1.ExChange
     Call ModulKl1.ExChangeGER
     Call Tabelle1.Button1_Standard
     Call Tabelle1.Button2_Standard
    'Call Tabelle1.Button3_Standard
     Call Tabelle1.Button4_Standard
     Call Tabelle1.Button5_Standard
     Call Tabelle1.Button6_Standard
     Call Tabelle1.Button7_Standard
     Call Tabelle1.Button8_Standard
     Call Tabelle1.Button9_Standard
    ' 17.03.2023 Urban
    If Range("DataChecked").Value = "" Then Range("DataChecked").Value = "ungeprueft"
     chkval = Range("DataChecked").Value
     TestChkVal = InStr(1, chkval, "Fehlerhaft", 1)
     If Range("DataChecked").Value <> "ungeprueft" And TestChkVal = 0 Then
       Tabelle1.CommandButton2.Caption = "Eingaben Abgeschlossen"
       Tabelle1.CommandButton2.ForeColor = black
       Else: Tabelle1.CommandButton2.Caption = "CheckMe"
             Range("State").Value = ""
     End If
   WB_open = False
   
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
  Call Kill_local_Cl_InfoDB
  Application.UseSystemSeparators = True 'Urban Gebietsschema zurücksetzen
  
End Sub
