Attribute VB_Name = "DB_Import"
'mit 7Zip Inhalt von *.zip-Dateien auslesen
Public userTmpVerz
Public DBHeadLine ' Datenbasen erkennen
Public dbCnt, ReadyDB, ZSB_KLziffer, manipulated, BGER_msg
Public Fehlerzähler As Integer
Public Fehlermerker As Integer
Public SigRichtDisaster As String
Public Itxt
Public cnt As Integer
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STILL_ACTIVE = &H103

'Public clampSpeicher(20000, 12) ' Cl_Info
'Public UtiliSpeicher(20000, 4) ' Örtlickeit
'Public FunctionsSpeicher(20000, 6) ' Funktionen

Public Sub StarteImport_CSVdatei()

'  Exit Sub ' Stopp bei Fehler aktivieren
On Error GoTo ThatsMist
userTmpVerz = Environ("USERPROFILE") & "\AppData\Local\Temp\"
Dim ZeilenWerte()
Filename = userTmpVerz & "CL_InfoDB.csv"
Open Filename For Input As #1    ' Datei zum Einlesen öffnen.
Do While Not EOF(1)    ' Schleife bis Dateiende.
   i = i + 1
   dbCnt = dbCnt + 1
   ReDim Preserve ZeilenWerte(i)   ' Größe des Elemente ändern
   Line Input #1, ZeilenWerte(i)   ' ZeilenWerte(i, 1)   ' Daten in zwei Variablen einlesen, CStr
   Mtext = (ZeilenWerte(i))        ' msoEncodingUTF8 65001 UTF-8-Codierung
   'Mtext = ZeilenWerte(i)
   'DBHeadLine = "#CL_Clamp###"
   If Mtext = "#CL_Utilization###" Then DBHeadLine = "UTILIZ"
   If Mtext = "#CL_Functions###" Then DBHeadLine = "FUNCT"
   If Mtext = "#CL_Info###" Then
      DBHeadLine = "CL_INFO"
    End If
   If Mtext = "#CL_Clamp###" Then DBHeadLine = "CLAMP"
   If Mtext = "#CL_Voltage###" Then DBHeadLine = "VOLTA"
   If Mtext = "#CL_VoltageType###" Then DBHeadLine = "VTYPE"
   If Mtext = "#CL_BGER_Pins###" Then DBHeadLine = "BGER"
   If Mtext = "#CL_SignalAbbreviation###" Then DBHeadLine = "SABBREV"
   If Mtext = "#CL_SignalCharacteristic###" Then DBHeadLine = "SChAR"
   If Mtext = "#CL_SignalDirection###" Then DBHeadLine = "SDIRECT"
   If Mtext = "#Create_Date###" Then DBHeadLine = "CRDAT"
   If Mtext = "#Autor###" Then DBHeadLine = "END"
   'If Mtext = "#CL_FunctBasePart###" Then DBHeadLine = "BaseFnc"
   If Mtext <> "" Then
      untersuch Mtext, DBHeadLine ' Urban 20.08.2019, Leerzeilen vermeiden und Daten auswerten
   End If
Loop
ReadyDB = "imported"
Close #1    ' Datei schließen.
Application.StatusBar = "Cl_Info DataBasesImport: OK"
BGER_Id_ermitteln ' Liste der vorhandener Komponentenbeschreibungen "BGER" in Zelle BC12 bereitstellen
ThatsMist:
If ReadyDB = "open" Then
   Close #1
   MsgBox ("Es fehlt dem VOBES Formblatt eine CLInfo-Datenbasis" & vbCrLf _
   & "You need for your activities a CLInfo database, is missing or incorrect file..." _
   & vbCrLf & vbCrLf & "Bitte drücken Sie.. please Select a PIN" & vbCrLf & "and push the button( PinBeschreibung )")
   ReadyDB = "crashed"
End If
End Sub
Sub Datenbasen_aus_archiv_entpacken()
  Dim objShell As Object, fso As Object, objExec As Object
  Dim str7z As String, strName As String, msg As String
  Dim strArchiv, strOut, i As Integer, k As Long
  Dim existCl_Info_db As Boolean 'Database existiert neben Formblatt 2.0
  explore = ""
  Pfad = ActiveWorkbook.Path
  Name = ActiveWorkbook.Name
  userTmpVerz = Environ("USERPROFILE") & "\AppData\Local\Temp\"  'Environ("USERPROFILE") & "\AppData\Local\Temp\CL_InfoDB.zip
'Sicherheitsprüfung: Es darf kein Formblatt aus 7Zip (UserTmp) heraus geöffnet werden,
'da die darin befindliche DB nicht zu Aktualisierung genutzt werden kann
  Pos1_ImChkTxt = InStr(1, Pfad, "\AppData\Local\Temp", 1) ' lokales UserHome
  'Urban 23.02.2023
  Pos2_ImChkTxt = InStr(1, Pfad, "MicrosoftEdgeDownloads", 1) 'lokales UserHome mit MS EDGE
  If Pos1_ImChkTxt >= 1 And Pos2_ImChkTxt <= 1 Then
    MsgBox ("Das Formblatt darf nicht aus 7Zip oder WinZip heraus gestartet werden," & vbLf & _
    "bitte voher den Datencontainer vollständig entpacken!" & vbLf & vbLf & _
    "Don´t use and start Formblatt out of 7zip or WinZip file browser" & vbLf & "Closing this Workbook now!")
    Workbooks(Name).Close SaveChanges:=False
  End If
  'Urban23.02.2023 Umgang mit SharePoint
  Pos3_ImChkTxt = InStr(1, Pfad, "http", 1) ' BrowserDaten (GroupShare)erkennen
  If Pos3_ImChkTxt >= 1 Then ' keine BrowserDaten (GroupShare) erkannt
  'neue neu neuer
     While Filename = False Or Chk_Filename <> "CL_InfoDB.zip"
     ChDir (Environ("USERPROFILE") & "\documents")
     'ChDir (ActiveWorkbook.Path)
     strArchiv = "CL_InfoDB.zip"
     strArchiv = Application.GetOpenFilename("CL_InfoDB.zip (*.zip), *.zip", Title:="Bitte jetzt die komprimierte Datenbasis.. CL_InfoDB.zip auswählen")
     Filename = strArchiv
       If Filename = False Then
           MsgBox "Leider haben Sie kein VOBES Formblatt Datenfile ausgewählt" & vbCrLf & "You haven`t choose CL_InfoDB zipfile"
           ReadyDB = "crashed"
           Exit Sub
       End If
       Chk_Filename = Right(Filename, 13)
       If Chk_Filename <> "CL_InfoDB.zip" And Chk_Filename = True Then
            MsgBox "Falsches ZipFile, Bitte eine VOBES Formblatt Datendatei auswählen!" _
            & vbCrLf & "wrong zip file" & vbCrLf & "please choose CL_InfoDB.zip file"
       End If
       GoTo unzip
     Wend
    'Else 'keine Browserdaten, die Datenbasis CL_InfoDB.zip werden als
    'strArchiv = Pfad & "\" & "CL_InfoDB.zip"
  End If
  
  Set objShell = CreateObject("WScript.Shell")
  'userTmpVerz = Environ("USERPROFILE") & "\AppData\Local\Temp\"  'Environ("USERPROFILE") & "\AppData\Local\Temp\CL_InfoDB.zip
  'CL_InfoDB.csv
   If Dir(userTmpVerz & "CL_InfoDB.csv") <> "" Then GoTo importieren
   str7z = "C:\Program Files\7-Zip\7z.exe" ' 7z deklarieren
   strWzUn = "C:\Program Files\WinZip\WZUNZIP.EXE" '  UNZIP von WinZip64 deklarieren
   TstWUz = Dir(strWzUn)
   Tst7z = Dir(str7z)
  If Tst7z = "" And TstWUz = "" Then
      atom = MsgBox("WIN Default UnZip wird verwendet..." & vbCrLf & vbCrLf & "Hintergrund: Eine ausführbare Datei der Datenkompressionsprogramme" & _
      vbCrLf & "7-Zip oder WinZip + Command Line" & vbCrLf & "sind nicht vorhanden" & vbCrLf & "Bitte installieren Sie bevorzugt das Programm 7Zip!", vbInformation, "CL_Info database: automatic import")
      'Anpassung 19.01.2019 auf Win10 entpacken
      'Exit Sub
  End If
'Urban 26.06.2019
If Dir(Pfad & "\" & "CL_InfoDB.zip") = "" Then ' die Datenbasis CL_InfoDB.zip ist nicht vorhanden, der User soll einen anderen Pfad benennen
  While Filename = False Or Chk_Filename <> "CL_InfoDB.zip"
     ChDir (ActiveWorkbook.Path)
     strArchiv = "CL_InfoDB.zip"
     strArchiv = Application.GetOpenFilename("CL_InfoDB.zip (*.zip), *.zip", Title:="Bitte jetzt die komprimierte Datenbasis.. CL_InfoDB.zip auswählen")
     Filename = strArchiv
       If Filename = False Then
           MsgBox "Leider haben Sie kein VOBES Formblatt Datenfile ausgewählt" & vbCrLf & "You haven`t choose CL_InfoDB zipfile"
           ReadyDB = "crashed"
           Exit Sub
       End If
       Chk_Filename = Right(Filename, 13)
       If Chk_Filename <> "CL_InfoDB.zip" And Chk_Filename = True Then
            MsgBox "Falsches ZipFile, Bitte eine VOBES Formblatt Datendatei auswählen!" _
            & vbCrLf & "wrong zip file" & vbCrLf & "please choose CL_InfoDB.zip file"
       End If
 Wend
 Else 'die Datenbasis CL_InfoDB.zip ist vorhanden, strArchiv zusammensetzen
    strArchiv = Pfad & "\" & "CL_InfoDB.zip"
 End If  'Urban 26.08.2019 wenn 7z nicht arbeitet: WIN10 UnZip als zusätzliche Alternative eingebaut
'Urban 23.02.2022
unzip:
  If strArchiv <> False And Tst7z <> "" And TstWUz = "" Then
      Set fso = CreateObject("Scripting.FileSystemObject")
      Set objExec = objShell.Exec(str7z & " e " & fso.GetFile(strArchiv).ShortPath & " -y -aoa" & userTmpVerz)
  End If
  If strArchiv <> False And TstWUz <> "" And Tst7z = "" Then
      Set fso = CreateObject("Scripting.FileSystemObject")
      Set objExec = objShell.Exec(strWzUn & " -e " & fso.GetFile(strArchiv).ShortPath & " " & userTmpVerz)
  End If
  'Schleife warten auf unzip Prozess
  checkFile = ""
  While checkFile = ""
      cntFailure = cntFailure + 1
      checkFile = Dir(userTmpVerz & "CL_InfoDB.csv") 'das db File auf Vorhandensein prüfen, Urban 26.08.2019
      Application.Wait (Now + TimeValue("0:00:01"))  ' eine Sekunde auf 7Zip warten
      If cntFailure = 3 Then
        'mit WIN Bordmitteln versuchen zu entpacken, Urban 26.08.2019
          Application.StatusBar = "UNZIP process: internal UnZip is unpacking ..."
          Set oApp = CreateObject("Shell.Application")
          oApp.Namespace(userTmpVerz).CopyHere oApp.Namespace(strArchiv).items.Item("CL_InfoDB.csv")
      End If
  Wend
  If Dir(userTmpVerz & "CL_InfoDB.csv") = "" Then
     ReadyDB = "crashed"
     Exit Sub
  End If
'der Vorgang ist gelaufen
importieren:
ReadyDB = "open"
Application.StatusBar = "Datenbasis ist vorhanden, der Import der Daten läuft... " & vbCrLf & " database exist, import is running..."
StarteImport_CSVdatei
End Sub

Public Sub Kill_local_Cl_InfoDB() 'für Testzwecke wird die Datenbasis gelöscht
ReadyDB = "open"
Application.StatusBar = "Reset Import Cl_Info : " & ReadyDB
userTmpVerz = Environ("USERPROFILE") & "\AppData\Local\Temp\"  'Environ("USERPROFILE") & "\AppData\Local\Temp\CL_InfoDB.csv
If Dir(userTmpVerz & "CL_InfoDB.csv") <> "" Then
     Kill (Environ("USERPROFILE") & "\AppData\Local\Temp\CL_InfoDB.csv")
     Application.StatusBar = "Datenbais des VOBES Formblattes wurde nicht gelöscht, killed Database File: Cl_InfoDB.csv, please get a new Database File by e42"
End If
'ReadyDB = "open" ' zwingt da neue einlesen der CSV Datei
ReadyDB = "crashed" ' zwingt da neue einlesen der CSV Datei
End Sub

Public Function ManipulationZiffer(ZSB_KLAttribute)
ZSB_KLziffer = 0
For i = 1 To Len(ZSB_KLAttribute)
  ZWwert = Mid(ZSB_KLAttribute, i, 1)
  ZSB_KLziffer = ZSB_KLziffer + Asc(ZWwert)
Next
'MsgBox ZSB_KLziffer
End Function

Public Function EndCheck_Pin()
Dim PIntypen(6) As String
Dim PInChar(5) As String
Dim PInOfl(6) As String

Fehlerzähler = 0 ' Startwert
StromErr = 0
PIntypen(0) = "Senke (SG Last)" '         VEC: IN
PIntypen(1) = "Quelle (SG Ausgang)" ' VEC: OUT
PIntypen(2) = "Senke (Komponente)" ' VEC: IN
PIntypen(3) = "Signal (CAN, Messein/-ausgang)"
PIntypen(4) = "Schaltkontakt (Relais/Schalter)"
PIntypen(5) = "Spule (Ansteuerung)"
PIntypen(6) = "Nicht verbunden"

PInChar(0) = "Glühlampe"
PInChar(1) = "Kapazität"
PInChar(2) = "Motor"
PInChar(3) = "Ohmsch"
PInChar(4) = "Spule"

Worksheets("Formblatt").Select
  OrRu = Range("PinsUpper").Row + 1 'Oberere PinZeile
  OrRl = Range("PinsLower").Row - 1 'Untere PinZeile
  PNumber = Range("Teilenummer").Value ' Teilenummer des BT
'Pinreihen
RowCnt4Chk = OrRl - OrRu
For Cloop = 0 To RowCnt4Chk ' Schleife Pinreihen
        TmpKLinfo = ""     'tmep. Klemmeninformation
        TmpLAH = ""        'temp. LAH Verweis
        TmpKLz = ""         'temp. erstes Zeichen der Klemmeninformation
        TmpAnBAnfg = "" 'temp. Anbindungsanforderung
  aktive_Reihe = OrRu + Cloop
  origCHKZiff = Cells(aktive_Reihe, 33).Value
' tempBGERid Korrekturerkennung Schleife Urban 28.06.2019
    For it = 0 To RowCnt4Chk
        ChkBGERid = Cells(aktive_Reihe, 226).Value  'BGERid holen
        IDschnip = Left(ChkBGERid, 3)                          'tmp Schnipsel schneiden
        If IDschnip = "tmp" Then
            'temporäre Id erkannt, für alle Steckplätze jetzt neu generieren
            'Pins Schleife zum generieren der temp. BGER_ID je Steckplatz
                For P = 0 To OrRl - OrRu
                  ZSBziffer = 0
                  StckPlz = Cells(OrRu + P, 2).Value
                  ZSBSTPN = StckPlz & PNumber
                      For i = 1 To Len(ZSBSTPN)
                           ZWwert = Mid(ZSBSTPN, i, 1)
                           ZSBziffer = ZSBziffer + Asc(ZWwert)
                      Next i
                  tempId = "tmp_" & ZSBziffer
                  Cells(OrRu + P, 226).Value = tempId
                  Cells(OrRu + P, 226).Interior.Color = RGB(204, 255, 255)
                  Cells(OrRu + P, 226).Font.Color = RGB(0, 0, 0)
                Next P
          'Pins Schleifen Ende
          it = RowCnt4Chk ' grosse Schleifen Abbruch
        End If
    Next it
' tempBGERid
  orig_BGERid = Cells(aktive_Reihe, 226).Value     'Urban 20.06.19, BGERid holen
  
'Alte Netzklasse hat ND erkannt oder der Nutzer wendet V1.4.1 an
BackStop = Cells(1, 33).Value
If BackStop <> "x" And BackStop <> "X" Then
    If origCHKZiff = "ND" And orig_BGERid = "" Then  '
    'Zellen für einen nicht angeschlossenen Pin befüllen
        Cells(aktive_Reihe, 15).Value = "n.c.#" 'Klemmenbeschreibung
        Btxt = "n.c."                           ' Klemme
        Cells(aktive_Reihe, 136).Value = Btxt   ' Klemme
        Ctxt = ""                               ' ZSB_Funktion
        Cells(aktive_Reihe, 143).Value = Ctxt   ' ZSB_Funktion
        Dtxt = ""                               ' Oertlichkeit
        Cells(aktive_Reihe, 152).Value = Dtxt   ' Oertlichkeit
        Etxt = "CRX"                            ' Signalrichtung
        Cells(aktive_Reihe, 158).Value = Etxt   '
        Ftxt = "12"                             ' Spannung
        Cells(aktive_Reihe, 163).Value = Ftxt   ' Spannung
        Cells(aktive_Reihe, 167).Value = "DC"
        Gtxt = "V"                              ' PotentialTyp
        Cells(aktive_Reihe, 171).Value = Gtxt   ' PotentialTyp
        Htxt = "DC"                             ' SignalArt
        Cells(aktive_Reihe, 176).Value = Htxt   ' SignalArt
        Cells(aktive_Reihe, 183).Value = "nicht angeschlossen" ' Freitxt hier eintragen
        ZSB_KLAttribute = Btxt & Ctxt & Dtxt & Etxt & Ftxt & Gtxt & Htxt
        ManipulationZiffer (ZSB_KLAttribute)  ' Wert 0 = keine BGER Werte vorhanden 'Funktion Prüfsumme bilden
        Cells(aktive_Reihe, 33).Value = ZSB_KLziffer
    'die berechnete Prüfziffer als Original festlegen
        origCHKZiff = ZSB_KLziffer
    End If
End If
If BackStop = "x" Or BackStop = "X" Then  'Urban 04.09.19, Reduzierter SW Umfang ist aktiv
    'Zellen für einen nicht angeschlossenen Pin befüllen
        Atxt = Cells(aktive_Reihe, 134).Value ' PInName
        Btxt = "nCL" 'Klemmenbeschreibung
        Btxt = Cells(aktive_Reihe, 136).Value ' Klemme
        Ctxt = "Redncy"                               ' ZSB_Funktion
        Cells(aktive_Reihe, 143).Value = Ctxt   ' ZSB_Funktion
        Dtxt = "nVal"                               ' Oertlichkeit
        Cells(aktive_Reihe, 152).Value = Dtxt   ' Oertlichkeit
        Etxt = "CRX"                            ' Signalrichtung
        Cells(aktive_Reihe, 158).Value = Etxt   ' Signalrichtung
        Ftxt = "12"                             ' Spannung
        Cells(aktive_Reihe, 163).Value = Ftxt   ' Spannung
        Cells(aktive_Reihe, 167).Value = "DC"
        Gtxt = "V"                              ' PotentialTyp
        Cells(aktive_Reihe, 171).Value = Gtxt   ' PotentialTyp
        Htxt = "DC"                             ' SignalArt
        Cells(aktive_Reihe, 176).Value = Htxt   ' SignalArt
        Cells(aktive_Reihe, 183).Value = "CL_Info deaktiv" ' Freitxt hier eintragen
        'ZSB_KLAttribute = Btxt & Ctxt & Dtxt & Etxt & Ftxt & Gtxt & Htxt
        'UGU 27.05.2021
        ZSB_KLAttribute = Cells(aktive_Reihe, 15).Value ' die Klemmeninformation ist manuell enstanden
        ManipulationZiffer (ZSB_KLAttribute)  ' Wert 0 = keine BGER Werte vorhanden 'Funktion Prüfsumme bilden
        Cells(aktive_Reihe, 33).Value = ZSB_KLziffer
    'die berechnete Prüfziffer als Original festlegen
        origCHKZiff = ZSB_KLziffer
        'GoTo smart
End If

'das Formblatt wurde gefüllt aktiviert
'eine BGER kann es geben
If Cloop > RowCnt4Chk Then Exit For       ' Abbruch  der Schleife
If BackStop <> "x" And BackStop <> "X" Then
        Atxt = Cells(OrRu + Cloop, 134).Value ' PInName
        Btxt = Cells(OrRu + Cloop, 136).Value ' Klemme
           If Btxt = "nCL" Then Btxt = ""
        Ctxt = Cells(OrRu + Cloop, 143).Value ' ZSB_Funktion
           If Ctxt = "nFct" Then Ctxt = ""
        Dtxt = Cells(OrRu + Cloop, 152).Value ' Oertlichkeit
           If Dtxt = "nVal" Then Dtxt = ""
        Etxt = Cells(OrRu + Cloop, 158).Value ' Signalrichtung
        Ftxt = Cells(OrRu + Cloop, 163).Value ' Spannung
        Gtxt = Cells(OrRu + Cloop, 171).Value ' PotentialTyp
        Htxt = Cells(OrRu + Cloop, 176).Value ' SignalArt
           If Btxt = "" And Ctxt <> "" And Dtxt <> "" Then
               Btxt = ""        ' keine Klemme
               ZSB_Klemme = Ctxt & "." & Dtxt
           End If
        If Btxt = "nCL" And Dtxt = "nVal" Then
             Btxt = ""         ' keine Klemme
             Dtxt = ""         ' keine Örtlichkeit
             ZSB_Klemme = Ctxt ' nur Funktion
        End If
        If Btxt <> "" And Ctxt = "nFct" Then
             Ctxt = ""                            ' keine Funktion
             ZSB_Klemme = Btxt & "#" & "." & Dtxt ' Klemme mit Örtlichkeit
        End If
        If Btxt <> "" And Ctxt = "" And Dtxt = "" Then
            ZSB_Klemme = Btxt & "#"              ' nur Klemme
        End If
        ZSB_KLAttribute = Btxt & Ctxt & Dtxt & Etxt & Ftxt & Gtxt & Htxt
        ManipulationZiffer (ZSB_KLAttribute)    ' Wert 0 = keine BGER Werte vorhanden
End If 'Backstop
If ZSB_KLziffer = 0 And orig_BGERid = "" Then ZSB_KLziffer = 301 'Wert 301 entspricht einem n.c.
'Urban 01.07.2019 , BGERid ist vorhanden damit OK, oder eine Manipulation von Hand eingetragene BGER anzeigen
If ZSB_KLziffer > 0 And orig_BGERid <> "" Then
   origCHKZiff = ZSB_KLziffer 'Urban 20.06.19, BGERid holen
   Cells(aktive_Reihe, 33).Value = ZSB_KLziffer
   Cells(aktive_Reihe, 33).Interior.Color = RGB(204, 255, 255)  'Farbe der Zelle zurücksetzten
End If
'Urban 20.06.19
If origCHKZiff = "" And ZSB_KLziffer > 0 Then
   Cells(aktive_Reihe, 33).Value = ZSB_KLziffer
   Cells(aktive_Reihe, 15).Value = ZSB_Klemme ' BGER Attribute übernehmen
   Cells(aktive_Reihe, 15).Interior.Color = RGB(204, 255, 255)  'Farbe der Zelle zurücksetzten
   origCHKZiff = ZSB_KLziffer ' da es eine gültige ZSB_KLziffer gibt, die  auf BGER daten basiert
   
End If
If origCHKZiff <> ZSB_KLziffer Then
    Cells(OrRu + Cloop, 15).Interior.Color = RGB(255, 150, 150)
    manipulated = True
    MsgBox "Sie verwenden zum Pin: " & Atxt & " keine standardisierten Daten aus den CL_Info Datenbasen!"
     Else:
      Cells(OrRu + Cloop, 15).Font.Color = RGB(0, 0, 255) 'Farbe der Zelle zurücksetzten
End If
   cnt = OrRu + Cloop

'BGER ID neu erzeugen je Pin da zwischendurch manipuliert sein konnte
   Steckplz = Sheets("Formblatt").Cells(cnt, 2).Value 'Steckplatz holen
   ZSBziffer = 0
   ZSBSTPN = Steckplz & PNumber
     For i = 1 To Len(ZSBSTPN)
        ZWwert = Mid(ZSBSTPN, i, 1)
        ZSBziffer = ZSBziffer + Asc(ZWwert)
     Next
'Urban 20.06.2019,
If orig_BGERid = "" Then
   tempId = "tmp_" & ZSBziffer
   Cells(cnt, 226).Value = tempId  ' temporäre BGER ID vergeben"
   Cells(cnt, 226).Interior.Color = RGB(255, 100, 100)
End If
'Urban20.06.2019
'NEU Pins prüfen
  zwtxt = Asc(Steckplz)
  'check Steckplatz Name, erstes Zeichen
     If zwtxt < 65 Or zwtxt > 90 Then
        MsgBox ("Zeile" & cnt & vbCrLf & ": Steckplatz Name [ " & Steckplz & " ] entspricht nicht der Vorgabe (A-Z)")
        Sheets("Formblatt").Cells(cnt, 2).Font.Color = RGB(255, 0, 0)
        Fehlerzähler = Fehlerzähler + 1
        manipulated = True
        Else:
         Cells(cnt, 2).Font.Color = RGB(0, 0, 255) 'Farbe der Zelle zurücksetzten
     End If
     If Len(Steckplz) > 1 Then
        ZweiteSt = Right(Steckplz, 1)
        zwtxt = Asc(ZweiteSt)
       If zwtxt < 65 Or zwtxt > 90 Then
          MsgBox ("Zeile" & cnt & vbCrLf & ": Steckplatz Name [ " & Steckplz & " ] Fehler nur (A-Z) erlaubt!")
          Cells(cnt, 2).Font.Color = RGB(255, 0, 0)
          manipulated = True
        Else:
           If manipulated = False Then Cells(cnt, 2).Font.Color = RGB(0, 0, 255) 'Farbe der Zelle zurücksetzten
       End If
     End If
  'Check Pin Nummer, Urban 20.04.2023, Pin Namen X zulassen
  PinNr = Sheets("Formblatt").Cells(cnt, 7).Value 'Pin-Nr. holen
  ZWtxt1 = Asc(PinNr)           'die linke Ziffer
  ZWtxt2 = Asc(Right(PinNr, 1)) 'die rechte Ziffer
  'Ausnahme in den Pins, Urban 10.02.2021
  If ZWtxt2 = 65 Or ZWtxt2 = 66 Or ZWtxt2 = 83 Or ZWtxt2 = 88 Then ZWtxt2 = 48 ' A, B, S oder X ist erlaubt
     If ZWtxt1 < 48 Or ZWtxt2 < 48 Or ZWtxt1 > 57 Or ZWtxt2 > 57 Then
        MsgBox ("Zeile" & cnt & vbCrLf & ": eine Ziffer im Pin Namen [ " & PinNr & " ] entspricht nicht der Vorgabe (0-9/A/S/X)")
        Sheets("Formblatt").Cells(cnt, 7).Font.Color = RGB(255, 0, 0)
        Fehlerzähler = Fehlerzähler + 1 'Formblatt wurde manipuliert
        manipulated = True
        Else:
         Cells(cnt, 7).Font.Color = RGB(0, 0, 255) 'Farbe der Zelle zurücksetzten
     End If
  'Ende Check Pin Nummer
  'Check Anbindungsanforderung 23.07.20 Urban
  ' * kennzeichnet es ist eine Bordnetzanbindung/ LAH Info vorhanden
   TmpKLinfo = Cells(cnt, 15).Value 'Klemmeninformation holen
   TmpKLz = Left(TmpKLinfo, 1)       'erstes Zeichen ermitteln, zum Feststellen eines Sternchen
   TmpLAH = Cells(cnt, 41).Value    'Lastenheft Referenz holen
   TmpAnBAnfg = Cells(cnt, 34).Value
    'Aktiv eine Sternchen setzen oder entfernen
     If TmpAnBAnfg <> "" And TmpKLz <> "*" Or TmpLAH <> "" And TmpKLz <> "*" Then
        Cells(cnt, 15).Value = "*" & TmpKLinfo
     End If
     If TmpAnBAnfg = "" And TmpLAH = "" And TmpKLz = "*" Then
        LAnB = Len(TmpKLinfo)
        TmpKLinfo = Mid(TmpKLinfo, 2, LAnB) 'Sternchen aus Anbindungsanforderung entfernen
        Cells(cnt, 15).Value = TmpKLinfo
     End If
  
  'Check Pintyp
  
   zwtxt = Sheets("Formblatt").Cells(cnt, 52).Value 'Pintyp holen
       If zwtxt <> PIntypen(0) And zwtxt <> PIntypen(1) And zwtxt <> PIntypen(2) And zwtxt <> PIntypen(3) And zwtxt <> PIntypen(4) And zwtxt <> PIntypen(5) And zwtxt <> PIntypen(6) Then
          MsgBox ("Zeile" & cnt & vbCrLf & ": Pintyp > " & zwtxt & " < entspricht nicht einem Original Typ")
          Sheets("Formblatt").Cells(cnt, 52).Font.Color = RGB(255, 0, 0)
          Fehlerzähler = Fehlerzähler + 1 'Exit Function 'Formblatt wurde manipuliert
          manipulated = True
          Else:
           Cells(cnt, 52).Font.Color = RGB(0, 0, 255) 'Farbe der Zelle zurücksetzten
       End If
           ' Signalrichtung muss gepüft werden, da VEC Attribute die Richtung enthalten
           ' PowerDistribution und SignalDirection „IN“     = PIntypen(0) = "Senke (SG Last)"
           ' PowerSupply und SignalDirection „IN“            = PIntypen(2) = "Senke (Komponente)"
           ' PowerDistribution und SignalDirection „OUT“ = PIntypen(1) = "Quelle (SG Ausgang)"
           ' Signalrichtung gleichgesetzt OUT = OUT/CRX/BI oder IN = OUT/CRX/BI
          If zwtxt = PIntypen(0) And Etxt = "OUT" Or zwtxt = PIntypen(1) And Etxt = "IN" Then   ' parking: Or ZWtxt = PIntypen(2) And Etxt = "OUT"
                Sheets("Formblatt").Cells(cnt, 52).Font.Color = RGB(255, 0, 0)
                If BackStop <> "x" And BackStop <> "X" Then
                   Fehlerzähler = Fehlerzähler + 1 'Exit Function 'Formblatt wurde manipuliert
                   SigRichtDisaster = "Pintyp: " & zwtxt & "< ungleich zu >  Sig.Dir. " & Etxt
                   manipulated = True
                End If
          End If
     ZWtxt2 = Sheets("Formblatt").Cells(cnt, 65).Value 'Pinchar Zellenwerte holen
     
  'check PinCharacter
          If zwtxt = "Senke (Komponente)" Then
            If ZWtxt2 <> PInChar(0) And ZWtxt2 <> PInChar(1) And ZWtxt2 <> PInChar(2) And ZWtxt2 <> PInChar(3) And ZWtxt2 <> PInChar(4) And ZWtxt2 = "" Then
               MsgBox ("Zeile" & cnt & vbCrLf & "Der Pintyp: " & zwtxt & ": zusammen mit PinCharacter > " & ZWtxt2 & " < entsprechen nicht den möglichen Werten")
               Cells(cnt, 65).Interior.Color = RGB(255, 240, 0)
               Fehlerzähler = Fehlerzähler + 1 'Exit Function 'Formblatt wurde manipuliert 'Or ZWtxt = "Senke (Komponente)" And ZWtxt2 = ""
               manipulated = True
               Else:
                Cells(cnt, 65).Interior.Color = RGB(204, 255, 255)  'Farbe der Zelle zurücksetzten
            End If
          End If
'Stromwerte prüfen
'Gebietsschema Dezimalzeichen
    'Selection.NumberFormat = "0.0000"
    'ActiveCell.FormulaR1C1 = "0.012"
With Application
       .DecimalSeparator = ","
      .UseSystemSeparators = False
End With
  If Sheets("Formblatt").Cells(cnt, 52).Value <> "Nicht verbunden" Then
     'Check Strom 1
     Itxt = Sheets("Formblatt").Cells(cnt, 78).Value 'I1 Zellenwerte holen
     Itxt = Replace(Itxt, ",", ".")
     With Sheets("Formblatt").Cells(cnt, 78)
       .NumberFormat = "0.0000"
       .FormulaR1C1 = Itxt
     End With
        If Itxt = "" Or Itxt = 0 Then
           Sheets("Formblatt").Cells(cnt, 78).Font.Color = RGB(255, 0, 0)
           Sheets("Formblatt").Cells(cnt, 78).Value = 0
           StromErr = StromErr + 1
           Else:
            Cells(cnt, 78).Font.Color = RGB(0, 0, 255) 'Farbe der Zelle zurücksetzten
        End If
        CurrentChk ("I1")
        'Itxt = Replace(Itxt, ".", ",") 'Gebietsschema egalisieren Punkt 46 oder Komma 44
        Sheets("Formblatt").Cells(cnt, 78).Value = Itxt
        If Fehlermerker > 0 Then Sheets("Formblatt").Cells(cnt, 78).Value = 0
     'Check Strom 2
     Itxt = Sheets("Formblatt").Cells(cnt, 85).Value 'I2 Zellenwerte holen
     Itxt = Replace(Itxt, ",", ".")
     With Sheets("Formblatt").Cells(cnt, 85)
       .NumberFormat = "0.0000"
       .FormulaR1C1 = Itxt
     End With
     Sheets("Formblatt").Cells(cnt, 85).FormulaR1C1 = Itxt
       If Itxt = "" Or Itxt = 0 Then
          Sheets("Formblatt").Cells(cnt, 85).Font.Color = RGB(255, 0, 0)
          Sheets("Formblatt").Cells(cnt, 85).Value = 0
          StromErr = StromErr + 1
          Else:
            Cells(cnt, 85).Font.Color = RGB(0, 0, 255) 'Farbe der Zelle zurücksetzten
       End If
       CurrentChk ("I2")
       'Itxt = Replace(Itxt, ".", ",") 'Gebietsschema egalisieren Punkt 46 oder Komma 44
       Sheets("Formblatt").Cells(cnt, 85).Value = Itxt
       If Fehlermerker > 0 Then Sheets("Formblatt").Cells(cnt, 85).Value = 0
     'check Strom 3
     Pintyp = Sheets("Formblatt").Cells(cnt, 52).Value 'Pintyp holen
     Itxt = Sheets("Formblatt").Cells(cnt, 92).Value 'I3 Zellenwerte holen
     Itxt = Replace(Itxt, ",", ".")
     With Sheets("Formblatt").Cells(cnt, 92)
       .NumberFormat = "0.0000"
       .FormulaR1C1 = Itxt
     End With
      
        If Pintyp <> "Signal (CAN, Messein/-ausgang)" Then
           If Itxt = "" Or Itxt = 0 Then
             Sheets("Formblatt").Cells(cnt, 92).Font.Color = RGB(255, 0, 0)
             Sheets("Formblatt").Cells(cnt, 92).Value = 0
             StromErr = StromErr + 1
             Else:
                Cells(cnt, 92).Font.Color = RGB(0, 0, 255) 'Farbe der Zelle zurücksetzten
           End If
        End If
       If Pintyp <> "Signal (CAN, Messein/-ausgang)" Then
          CurrentChk ("I3")
          'Itxt = Replace(Itxt, ".", ",") 'Gebietsschema egalisieren Punkt 46 oder Komma 44
          
          Sheets("Formblatt").Cells(cnt, 92).Value = Itxt
          If Fehlermerker > 0 Then
            Sheets("Formblatt").Cells(cnt, 92).Value = 0
            Cells(cnt, 92).Font.Color = RGB(255, 0, 0)
          Else:
            Cells(cnt, 92).Font.Color = RGB(0, 0, 255) 'Farbe der Zelle zurücksetzten
       End If

       End If
     'check Strom 4
     Itxt = Sheets("Formblatt").Cells(cnt, 99).Value 'I4 Zellenwerte holen
     Itxt = Replace(Itxt, ",", ".")
     With Sheets("Formblatt").Cells(cnt, 99)
       .NumberFormat = "0.0000"
       .FormulaR1C1 = Itxt
     End With
     CurrentChk ("I4")
     'Itxt = Replace(Itxt, ".", ",") 'Gebietsschema egalisieren Punkt 46 oder Komma 44
     Sheets("Formblatt").Cells(cnt, 99).Value = Itxt
       If Fehlermerker > 0 Then
          Sheets("Formblatt").Cells(cnt, 99).Value = 0
          Cells(cnt, 99).Font.Color = RGB(255, 0, 0)
          Else:
            Cells(cnt, 99).Font.Color = RGB(0, 0, 255) 'Farbe der Zelle zurücksetzten
       End If
     'check Strom 5
       Itxt = Sheets("Formblatt").Cells(cnt, 106).Value 'I5 Zellenwerte holen
       Itxt = Replace(Itxt, ",", ".")
       With Sheets("Formblatt").Cells(cnt, 106)
         .NumberFormat = "0.0000"
         .FormulaR1C1 = Itxt
       End With
       CurrentChk ("I5")
       'Itxt = Replace(Itxt, ".", ",") 'Gebietsschema egalisieren Punkt 46 oder Komma 44
       Sheets("Formblatt").Cells(cnt, 106).Value = Itxt
       If Fehlermerker > 0 Then
          Sheets("Formblatt").Cells(cnt, 106).Value = 0
          Cells(cnt, 106).Font.Color = RGB(255, 0, 0)
         Else:
            Cells(cnt, 106).Font.Color = RGB(0, 0, 255) 'Farbe der Zelle zurücksetzten
       End If
  End If
Next
   s_UserName = Application.UserName
   If manipulated = False And Fehlerzähler = 0 Then
      s_StampText = "Checked: " & s_UserName & " - " & datum & "/" & zeit
      Range("DataChecked").Value = s_StampText
      Range("DataChecked").Interior.Color = RGB(204, 255, 255)
      Tabelle1.CommandButton2.Caption = "Eingaben Abgeschlossen" ' Urban 17.03.2023
   End If
   If manipulated = True And Fehlerzähler > 0 Then
      s_StampText = "Fehlerhafte Eingaben: " & s_UserName & " - " & datum & "/" & zeit
      Range("DataChecked").Value = s_StampText
      Range("DataChecked").Interior.Color = RGB(255, 240, 0)
      Range("State").Value = ""
      Tabelle1.CommandButton2.Caption = "CheckMe" ' Urban 21.03.2023
      Tabelle1.CommandButton2.ForeColor = blue
   End If

   If Fehlermerker > 0 Then
      MsgBox ("Bitte Ströme überprüfen")
      manipulated = True
      Tabelle1.CommandButton2.Caption = "CheckMe" ' Urban 17.03.2023
      Tabelle1.CommandButton2.ForeColor = blue
   End If


End Function
Public Function CurrentChk(ICell)
Fehlermerker = 0
curr_len = Len(Itxt)
'Itxt = Replace(Itxt, ".", ",") 'Gebietsschema egalisieren Punkt 46 oder Komma 44
For ii = 1 To curr_len
       ZifferCurrent = Asc(Mid(Itxt, ii, 1))
       If ZifferCurrent = 44 Or ZifferCurrent = 46 Then ' ist ein Komma oder Punkt
          PktKomaCnt = PktKomaCnt + 1
            If PktKomaCnt > 1 Then Fehlermerker = Fehlermerker + 1
       End If
       If ZifferCurrent <> 44 And ZifferCurrent <> 46 And ZifferCurrent < 48 Or ZifferCurrent > 57 Then   'Tiefenpruefung
          Fehlermerker = Fehlermerker + 1
        End If
Next ii
   If Fehlermerker > 0 Then MsgBox ("Zeile" & cnt & vbCrLf & ": Korrektur(en) im " & "Strom > " & ICell & " < erforderlich!" _
    & vbCrLf & "Im Strom-Datenfeldbereich sind nicht alle Stromwerte " & vbCrLf & "in Ampere mit 4 Stellen nach" _
    & vbCrLf & "dem Dezimaltrennzeichen (Komma) je Pin eingegeben!" & vbCrLf & "Korrekturwert ist 0,0000")
End Function

Public Function BGER_PinTab_UserFill(BGERIdUser) ' BGER Pintabelle einlesen
    tempId = "tmp_" & Int((99000 * Rnd) + 1)
    BGERIdUser_ln = Len(BGERIdUser) ' Länge der BGER id
    BGERIdUser = Mid(BGERIdUser, 16, BGERIdUser_ln) ' SteckplzA1: den Steckplatz und Pin Nummmer abschneiden)
  If ReadyDB = "open" Then ' es ist keine CL Info Datenbasis eingelesen
     Application.Cursor = xlWait
     StarteImport_CSVdatei 'UGU, neue Funktion des Datenimports
     Application.Cursor = xlDefault
  End If
 If ReadyDB <> "imported" Then Exit Function ' zuerst muss eine Datenbasis vorhanden sein, ReadyDB = "imported"
  'MsgBox ("Datenbasis ist importiert und jetzt ..."),vor dem Einlesen der Daten Pins säubern
  'ModulKl1.Klemmeninfo_delete ("NoGui") 'für Testzwecke kommentieren,der Nutzer sollte eine BGER ausgewählt haben
  Worksheets("Formblatt").Select
  OrgRu = Range("PinsUpper").Row + 1 'Zeile des oberen und unteren Bereichs der Pintabelle feststellen
  OrgRl = Range("PinsLower").Row - 1
  RowCnt2FillInn = OrgRl - OrgRu   'Pinreihen 0-n
For BGERcnt = 0 To RowCnt2FillInn
       BGERcheckID = Cells(OrgRu + BGERcnt, 226).Value
     If BGERcheckID = "" Then
         Exit Function 'Urban 13.06.2019, es ist keine BGER_Id (Wert ="") in der PinTabelle eingetragen
     End If
     If BGERIdUser = BGERcheckID Then ' es kann in der Zeile der Inhalt gelöscht werden
       Cells(OrgRu + BGERcnt, 15).Value = "n.c." ' aktive Zeile leeren
       Cells(OrgRu + BGERcnt, 15).Font.Color = RGB(0, 0, 0)
       Cells(OrgRu + BGERcnt, 33).Value = "ND" ' aktive Zeile leeren
       Cells(OrgRu + BGERcnt, 33).Font.Color = RGB(0, 0, 0)
       Range(Cells(OrgRu + BGERcnt, 136), Cells(OrgRu + BGERcnt, 183)).Value = "" ' aktive Zeile leeren
       Range(Cells(OrgRu + BGERcnt, 136), Cells(OrgRu + BGERcnt, 183)).Font.Color = RGB(0, 0, 0)
       Range(Cells(OrgRu + BGERcnt, 226), Cells(OrgRu + BGERcnt, 226)).Value = tempId ' temp.BGER Id festlegen
       Range(Cells(OrgRu + BGERcnt, 226), Cells(OrgRu + BGERcnt, 226)).Interior.Color = RGB(255, 255, 255)
     End If
Next

' BGER ID in jeder Zeile feststellen
  For i = 1 To 1000 ' entspricht einem Bauteil mit 16 Steckplz und 3 BGER je Steckplz
    BgerIdiRow = BGERspeicher(i, 0)  'BGerID
    BGERpinName = BGERspeicher(i, 1) 'PInName
    While BgerIdiRow <> BGERIdUser And i < 1000
        i = i + 1
        BgerIdiRow = BGERspeicher(i, 0)  'BGerID
        BGERpinName = BGERspeicher(i, 1) 'PInName
    Wend
'BGER Id der Zeile mit der vom Anwender geforderten BGER ID vergleichen
'0-BGerID; 1-Pin; 2-Clamp; 3-Lnk; 4-Function; 5-Utilization; 6-Sig.Dir.; 7-Volt;
'8-VoltType; 9-SigAbbrev.; 10-Sig.Char.; 11-Description(Freitext)
           For pLoop = 0 To RowCnt2FillInn 'Pinliste abarbeiten
             PinTabName = Cells(OrgRu + pLoop, 134).Value
            ' zuerst prüfen ob der 1. PinName gleich dem PinNamen im BGER Speicher ist
             'If PinTabName <> BGERpinName Then
                 'FehlerStatus = " ! Achtung: PinMissMatch (ungleiche Pinanzahl)" & vbCrLf & "Die in electric42 freigegebene Pinanzahl stimmt hier nicht mehr überein."
                 'Cells(OrgRu + dLooP, 226).Value = tempId '"tmp_Id"
                 'Cells(OrgRu + dLooP, 226).Interior.Color = RGB(255, 100, 100)
                 'Cells(OrgRu + dLooP, 15).Interior.Color = RGB(255, 100, 100)
             'End If
             If PinTabName = BGERpinName Then
              Cells(OrgRu + pLoop, 226).Value = BGERspeicher(dLooP + i, 0)  'BGerID
              If BGERspeicher(dLooP + i, 2) = "nCL" Then                    'Klemme
                 Cells(OrgRu + pLoop, 136).Value = ""
                Else: Cells(OrgRu + pLoop, 136).Value = BGERspeicher(dLooP + i, 2)
              End If
              Cells(OrgRu + pLoop, 139).Value = BGERspeicher(dLooP + i, 3)  'link
              If BGERspeicher(dLooP + i, 4) = "nFct" Then                   'Funktion
                 Cells(OrgRu + pLoop, 143).Value = ""
                 Else: Cells(OrgRu + pLoop, 143).Value = BGERspeicher(dLooP + i, 4)
              End If
              If BGERspeicher(dLooP + i, 5) = "nVal" Then                   'Oertlichkeit
                 Cells(OrgRu + pLoop, 152).Value = ""
                 Else: Cells(OrgRu + pLoop, 152).Value = BGERspeicher(dLooP + i, 5)
              End If
              Cells(OrgRu + pLoop, 158).Value = BGERspeicher(dLooP + i, 6)  'Signalrichtung
              Cells(OrgRu + pLoop, 163).Value = BGERspeicher(dLooP + i, 7)  'Spannung
              Cells(OrgRu + pLoop, 167).Value = BGERspeicher(dLooP + i, 8)  'Spannungs Typ
              Cells(OrgRu + pLoop, 171).Value = BGERspeicher(dLooP + i, 9)  'Signal typ
              Cells(OrgRu + pLoop, 176).Value = BGERspeicher(dLooP + i, 10) 'Signal Char.
              Cells(OrgRu + pLoop, 183).Value = BGERspeicher(dLooP + i, 11) 'Description (Freitext)
              If BGERspeicher(dLooP + i, 2) = "nCL" Then
                 ZSB_KLemmeninfo = Cells(OrgRu + pLoop, 136).Value & Cells(OrgRu + pLoop, 143).Value & "." & Cells(OrgRu + pLoop, 152).Value
              Cells(OrgRu + pLoop, 15).Value = ZSB_KLemmeninfo
              Else: ZSB_KLemmeninfo = Cells(OrgRu + pLoop, 136).Value & "#" & Cells(OrgRu + pLoop, 143).Value & "." & Cells(OrgRu + pLoop, 152).Value
              End If
              Cells(OrgRu + pLoop, 15).Value = ZSB_KLemmeninfo
             'ZSB_KLAttribute = Klemme & ZSB_Funktion & Oertlichkeit & PotentialTyp & Signalrichtung & Spannung & SignalArt
              ZSB_KLAttribute = Cells(OrgRu + pLoop, 136).Value & Cells(OrgRu + pLoop, 143).Value & Cells(OrgRu + pLoop, 152).Value _
              & BGERspeicher(dLooP + i, 9) & BGERspeicher(dLooP + i, 6) & BGERspeicher(dLooP + i, 7) & BGERspeicher(dLooP + i, 10)
              ManipulationZiffer (ZSB_KLAttribute)
              Cells(OrgRu + pLoop, 33).Value = ZSB_KLziffer
              Cells(OrgRu + pLoop, 15).Interior.Color = RGB(204, 255, 255)
              pLoop = RowCnt2FillInn
              FehlerStatus = ""
             End If
           Next pLoop 'bestellte Pinliste abarbeiten
    'End If
  Application.StatusBar = "BGER Pins imported:  " & FehlerStatus
  Next i
  'Exit Function ' Pins wurden gefüllt
End Function

Public Sub BGER_Id_ermitteln()
' es werden alle BGER IDs aud der CSV Datei eingelesen und in der Zelle B12 als Auswahlliste dargestellt
' nach der Auswahl soll eine BGER-Pinliste in das Formblatt hinter den Pins geladen werden
  '
  If Range("BC12").Value <> "initial" Then
    Exit Sub ' mehrmaliges laden der BGERid Liste verhindern
  End If
  If ReadyDB <> "imported" Then
     MsgBox "Please try later, because CL-Info Databases are missing"
     Exit Sub
  End If
  Worksheets("Formblatt").Select
  BGERidnt_act = "" 'aktuelle BGER ermitteln
  BGERidnt_pre = "" 'vorherige BGER ermitteln
  
  For i = 0 To 50 'UBound(BGERspeicher, 1) ' hier stehen die BGER ID´s
        BGERident_act = "IdNr_Steckplz" & Left(BGERspeicher(i, 1), 1) & ":" & BGERspeicher(i, 0) 'BGerID
        If i > 0 Then BGERidnt_pre = "IdNr_Steckplz" & Left(BGERspeicher(i - 1, 1), 1) & ":" & BGERspeicher(i - 1, 0) 'BGerID vorher
        If BGERidnt_pre <> BGERident_act And BGERident_act <> "" And BGERident_act <> "IdNr_Steckplz:" Then
            BGER_msg = BGER_msg & BGERident_act & "," ' Frage integrieren
            'MsgBox BGER_msg
        End If
  Next i
  If BGER_msg = "" Then
     Application.StatusBar = "VOBES INFO: Es ist keine alternative Beschaltung (Neuer Prozess) in der CL_Info Datenbasis gefunden"
     'Exit Sub
  End If
    With Range("BC12").Validation
    'With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=BGER_msg
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = "database must be corrupt"
        .ShowInput = True
        .ShowError = True
    End With
End Sub
Public Function untersuch(Mtext, DBHeadLine)
Dim SuchZeichen, PosImSuchTxt, test, Zähler, SWLen, Stopp
'Dim CLdb_Datum As Date 'Datum des C#CL_InfoDB ExportFiles
'Dim DatumAct As Date   'aktuelles Datum
Dim USuchtxt, USuchtxt1, USuchtxt2, USuchtxt3, USuchtxt4, USuchtxt5, USuchtxt6, USuchtxt7 ' Umlaute wandeln

Pfad = ActiveWorkbook.Path
Name = ActiveWorkbook.Name
'Sicherheitsprüfung: Es darf kein Formblatt aus 7Zip (UserTmp) heraus geöffnet werden,
'da die darin befindliche DB nicht zu Aktualisierung genutzt werden kann
'PosImChkTxt = InStr(1, Pfad, "\AppData\Local\Temp", 1) ' Reiner Textvergleich
'If PosImChkTxt >= 1 Then
 '  MsgBox ("Bitte diese Datei entpacken, das Formblatt darf nicht aus 7Zip heraus gestartet werden!" & vbLf & vbLf & "Don´t use and start Formblatt out of 7Z file browser" & vbLf & "Closing this Workbook now!")
  ' Workbooks(Name).Close SaveChanges:=False
'End If
'Ende

Filename = Pfad & "\" & Name
'Urban 23.02.2023
PosImChkTxt = InStr(1, Pfad, "sharepoint", 1) ' Sharepoint Ausnahme Reiner Textvergleich
If PosImChkTxt < 1 Then
  'todo
     DatumAct = FileDateTime(Filename)
     lnDtAct = Len(DatumAct)
     If lnDtAct > 19 Then
       DatumAct = Mid(DatumAct, 1, lnDtAct - 6) ' Urban 14.08.2023 Datum zu lang a/ m/
     End If
     'FileDateTime liefert "18.10.2018 12:44:21"
     'Datenbasis liefert   "19.08.2019 17:12:04"
     'Gebietsschema Spain "22/08/2019 12:23:34" Urban 22.08.2019
     
    DatumAct = Replace(DatumAct, ".", "/")
    If DBHeadLine = "CRDAT" Then
         If Mtext <> "#Create_Date###" Then 'Datum untersuchen
           CLdb_Datum = Mtext               'Datum der Datenbasis "19.08.2019 17:12:04"
           CLdb_Datum = Replace(CLdb_Datum, ".", "/") 'Gebietsschema Spain "22/08/2019 12:23:34"
           CLdb_Diff = DateDiff("d", CLdb_Datum, DatumAct) 'Tage
              If CLdb_Diff > 10 Then
                Range("CLdat").Font.Color = RGB(255, 0, 0)
                Else: Range("CLdat").Font.Color = RGB(0, 0, 0)
              End If ' DB ist älter als 5 Tage
              Range("CLdat").Value = "#CL_InfoDB " & CLdb_Datum & "###" '#CL_InfoDB 02.08.2018###
            Exit Function ' Urban 21.07.2023 bleibt auch in VWMX erhalten
         End If
    End If
     
End If

If DBHeadLine <> "UTILIZ" And DBHeadLine <> "FUNCT" And DBHeadLine <> "CL_INFO" And DBHeadLine <> "BGER" _
And DBHeadLine <> "CLAMP" And DBHeadLine <> "VOLTA" And DBHeadLine <> "SABBREV" And DBHeadLine <> "SChAR" _
And DBHeadLine <> "SDIRECT" Then Exit Function

Stopp = Right(Mtext, 3)
 If Stopp = "###" Then
   dbCnt = 0 ' Datenspeicher je Datenbasis auf 0 setzen
   Exit Function
 End If
If DBHeadLine = "SDIRECT" Then
    k = 0
End If
'USuchtxt = Mtext       ' "9;Blinkgeber;Flashing unit;V;DC;IN;12"   ' Zu durchsuchende Zeichenfolge.
 USuchtxt1 = Replace(Mtext, "ü", "ue")
 USuchtxt2 = Replace(USuchtxt1, "Ü", "Ue")
 USuchtxt3 = Replace(USuchtxt2, "ä", "ae")
 USuchtxt4 = Replace(USuchtxt3, "Ä", "Ae")
 USuchtxt5 = Replace(USuchtxt4, "ö", "oe")
 USuchtxt6 = Replace(USuchtxt5, "Ö", "Oe")
 USuchtxt7 = Replace(USuchtxt6, "ß", "ss")
 USuchtxt = USuchtxt7
If DBHeadLine = "BGER" Then
    If Mtext = "" Then
       'Range("BC12").Select
     ActiveSheet.ClearCircles
     Exit Function
   End If
End If
SWLen = Len(USuchtxt)
SuchZeichen = ";"    ' nach dem Zeichen ";" suchen.
test = True: Zähler = 0  ' Variablen initialisieren.
Do    ' Äußere Schleife.
    Do While Zähler < SWLen    ' Innere Schleife.
        Zähler = Zähler + 1 ' Zähler hochzählen.
        PosImSuchTxt = InStr(Zähler, USuchtxt, SuchZeichen, 1) ' Reiner Textvergleich
        'Ausnahme Endposition
         If Zähler > 1 And PosImSuchTxt = 0 Then
          'Enposition bestimmen
          DB_ValueTxt = Mid(USuchtxt, Zähler, SWLen)
          Else: DB_ValueTxt = Mid(USuchtxt, Zähler, PosImSuchTxt - Zähler)
         End If
         
         If DBHeadLine = "UTILIZ" Then
             UtiliSpeicher(dbCnt, Counter) = DB_ValueTxt
             Counter = Counter + 1
             'Application.StatusBar = "Import UTilization: " & dbCnt
         End If
         If DBHeadLine = "FUNCT" Then
             FunctionsSpeicher(dbCnt, Counter) = DB_ValueTxt
             Counter = Counter + 1
             'Application.StatusBar = "Import Function: " & dbCnt
         End If
         If DBHeadLine = "BGER" Then
             BGERspeicher(dbCnt, Counter) = DB_ValueTxt
             Counter = Counter + 1 'BGERspeicher(1000, 11)
             'Application.StatusBar = "Import BGERpins: " & dbCnt
         End If
         If DBHeadLine = "CL_INFO" Then
             clampSpeicher(dbCnt, Counter) = DB_ValueTxt
             'If DB_ValueTxt = "MidrangeV8A" Then
              '  hugo = 1
             'End If
             Counter = Counter + 1 'clampSpeicher(20000, 12)
             'Application.StatusBar = "Import Cl_Info: " & dbCnt
         End If
         If DBHeadLine = "CLAMP" Then
             clamps(dbCnt, Counter) = DB_ValueTxt
             Counter = Counter + 1  ' Low;Low-Side Schalter;low side switch;M;DC;IN;12
             'Application.StatusBar = "Import Clamps: " & dbCnt
         End If
         If DBHeadLine = "VOLTA" Then
             Volt(dbCnt, Counter) = DB_ValueTxt
             Counter = Counter + 1  ' 12;car voltage supply; PKW Batterie Versorg. I; DC
             'Application.StatusBar = "Import Voltages: " & dbCnt
         End If
         'SigAbbrev(10, 1)
         If DBHeadLine = "SABBREV" Then
             SigAbbrev(dbCnt, Counter) = DB_ValueTxt
             Counter = Counter + 1  ' mixed; mixed signal; gemischtes Signal
             'Application.StatusBar = "Import Signal Abbreviation: " & dbCnt
         End If
         'SigChar(100, 2) 'Signalcharakter
         If DBHeadLine = "SChAR" Then
             SigChar(dbCnt, Counter) = DB_ValueTxt
             Counter = Counter + 1  ' mixed; mixed signal; gemischtes Signal
             'Application.StatusBar = "Import Signal Character: " & dbCnt
         End If
         If DBHeadLine = "SDIRECT" Then
             IOXs(dbCnt, Counter) = DB_ValueTxt
             Counter = Counter + 1  ' In; Input; Eingang;
             'Application.StatusBar = "Import Direction: " & dbCnt
         End If
         
        If PosImSuchTxt = 0 Then    ' Wenn Bedingung = True,
            test = False    ' Attributwert auf False setzen.
            Exit Do         ' Innere Schleife verlassen.
        End If
        Zähler = PosImSuchTxt
    Loop
Loop Until test = False     ' Äußere Schleife sofort verlassen.
'Mexico: kein DatumProblem Urban 21.07.2023 ReadyDB = "imported" ist nicht notwendig
End Function


Public Function Tmp_BGR_Korrektur()
' tempBGERid Korrekturerkennung Schleife Urban 28.06.2019
   aktive_Reihe = ActiveCell.Row
   OrRu = Range("PinsUpper").Row + 1 'Oberere PinZeile
   OrRl = Range("PinsLower").Row - 1 'Untere PinZeile
   PNumber = Range("Teilenummer").Value ' Teilenummer des BT
' Pinreihen
  RowCnt4Chk = OrRl - OrRu
    For jt = 0 To RowCnt4Chk
        ChkBGERid = Cells(aktive_Reihe, 226).Value  'BGERid holen
        StckPlzNc = Cells(aktive_Reihe, 2).Value    'Steckplz erkennen
        '
        IDschnip = Left(ChkBGERid, 3)                          'tmp Schnipsel schneiden
        If IDschnip = "tmp" Or IDschnip = "" Then
            'temporäre Id erkannt, für alle Steckplätze jetzt neu generieren
            'Pins Schleife zum generieren der temp. BGER_ID je Steckplatz
                For P = 0 To OrRl - OrRu
                  ZSBziffer = 0
                  StckPlz = Cells(OrRu + P, 2).Value
                  If StckPlzNc = StckPlz Then
                    ZSBSTPN = StckPlz & PNumber
                      For i = 1 To Len(ZSBSTPN)
                           ZWwert = Mid(ZSBSTPN, i, 1)
                           ZSBziffer = ZSBziffer + Asc(ZWwert)
                      Next i
                     tempId = "tmp_" & ZSBziffer
                     Cells(OrRu + P, 226).Value = tempId
                     Cells(OrRu + P, 226).Interior.Color = RGB(255, 100, 100)
                     Cells(OrRu + P, 226).Font.Color = RGB(0, 0, 0)
                  End If
                Next P
          'Pins Schleifen Ende
          jt = RowCnt4Chk ' grosse Schleifen Abbruch
        End If
    Next jt
' tempBGERid
End Function
