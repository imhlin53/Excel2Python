Attribute VB_Name = "ModulKl1"
Public AnBAnfg ' Zusammenbau der Anbindungsanforderung
Public BN_LAH_ID ' BN/Q- Lastheftverweis Id Nummer
Public functN
Public FunctionsSpeicher(20000, 6) ' Funktionen
Public UtiliSpeicher(20000, 4) ' Örtlichkeiten
Public clampSpeicher(20000, 12) ' Kl_Info
Public clamps(100, 6) 'Klemmen
Public Volt(100, 3) 'Spannungen
Public SigAbbrev(10, 1)
Public SigChar(100, 2) 'Signalcharakter
Public IOXs(10, 2) ' Richtungen
Public BGERspeicher(1000, 11)
Public ShordWordsSpeicher(3000, 1)
Public PinIndexVec(4, 300)
Public mCurFild(8)
Public Fehler
Public SteckplzPin, SteckplzPinInfo, SteckplzPinTypInf, SteckplzPinTypI, Vortrag, BNAStat
Public defKLinhalt
Public frage
Public foundTxt
Public VFbl ' VOBES Formblatt
Public ReferenceDataSheet ' InputKomponentenDaten
Public Klemme, Beschreibung_DE, Beschreibung_EN, PotentialTyp, SignalArt, Signalrichtung, Spannung, SEarchVector, BGERIdUser_Now, BGERIdUser_Last
Dim cl_function, cl_descriptiveTextEN, cl_descriptiveTextDE, cl_utility, cl_util_descriptiveTextEN, cl_util_descriptiveTextDE, CL_SignalCharacteristic, CL_SignalAbbreviation, cl_direction, cl_manualTxt, cl_Freitext
Dim clamp, cl_funct, cl_descTextEN, cl_descTextDE, cl_util, cl_util_descTextEN, cl_util_descTextDE, cl_voltage, CL_SigChar, CL_SigAbbrev, cl_dir, cl_manTxt, cl_Frtxt

Sub Klemme_select()
BNAStat = False ' Bordnetzanbindung Status zurückgesetzt (nicht vorhanden)
 On Error GoTo Errorhandler
 If Tabelle1.NEWpins = "edit" Then
    Tabelle1.NEWpins = "gen"
    Exit Sub ' Pintabelle generiert PinZeilen
 End If
 VFbl = ActiveWorkbook.Name
 Workbooks(VFbl).Activate
 ActiveWorkbook.Worksheets("Formblatt").Visible = True
 ActiveWorkbook.Worksheets("Formblatt").Activate
  SteckplzPin = ""
  SteckplzPinInfo = ""
  aktive_Reihe = ActiveCell.Row
  aktive_Spalte = ActiveCell.Column
  origC = Range("PinsUpper").Column + 14 'ActiveCell.Column
   If aktive_Spalte < origC Or aktive_Spalte > 32 Then
     aktive_Spalte = 15
   End If
  OrigRu = Range("PinsUpper").Row + 1 'ActiveCell.Row
  OrigRl = Range("PinsLower").Row - 1 'ActiveCell.Row
  If aktive_Reihe < OrigRu Or aktive_Reihe > OrigRl Then
     MsgBox "Sie haben keinen Pin ausgewählt, You haven´t choose a Pin: Exit"
     Exit Sub
  End If
 frage = "yes" ' Default festgelegt
 SteckplzPin = Cells(aktive_Reihe, 2).Value & Cells(aktive_Reihe, 7).Value
 SteckplzPinInfo = Cells(aktive_Reihe, 15).Value ' Klemmeninformation alt oder zusammengebaut, Urban 27.06.19
 'Urban 20.07.20
 BNA = Cells(aktive_Reihe, 34).Value ' Bordnetzanbindung holen
 LAH = Cells(aktive_Reihe, 41).Value ' Lastenheft Referenz holen
 If BNA <> "" Or LAH <> "" Then
      BNAStat = True 'Status ist true wenn LAH oder BN Anbindung einen Wert hat
 End If
 'Urban 20.07.20
 SteckplzPinTypI = Cells(aktive_Reihe, 52).Value ' Pintyp als Info, Urban 27.06.19
 SteckplzPinTypInf = "UNI"
    If SteckplzPinTypI = "Quelle (SG Ausgang)" Then
       SteckplzPinTypInf = "OUT"
    End If
   If SteckplzPinTypI = "Senke (SG Last)" Then
        SteckplzPinTypInf = "IN"
   End If
 
 defKLinhalt = Cells(aktive_Reihe, 33).Value
 NoPInDflt = Cells(aktive_Reihe, 15).Value ' "n.c." 'Klemmenbeschreibung
 NoPIn2Use = Cells(aktive_Reihe, 52).Value ' "Nicht verbunden" darf keine KL_Info erhalten
  If NoPIn2Use = "Nicht verbunden" Then
   'Zellen für einen nicht angeschlossenen Pin befüllen
    Cells(aktive_Reihe, 15).Value = "n.c.#" 'Klemmenbeschreibung
         Cells(aktive_Reihe, 15).Font.Color = RGB(0, 0, 0) 'Urban 01.07.2019
         Cells(aktive_Reihe, 15).Interior.Color = RGB(204, 255, 255) 'auf Default Farbe zurücksetzen
    Cells(aktive_Reihe, 78).Value = "0"     ' I1 Strom
    Btxt = "n.c."                           ' Klemme
    Cells(aktive_Reihe, 136).Value = Btxt   ' Klemme
    Ctxt = ""                               ' ZSB_Funktion
    Cells(aktive_Reihe, 143).Value = Ctxt   ' ZSB_Funktion
    Dtxt = ""                               ' Oertlichkeit
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
    Cells(aktive_Reihe, 183).Value = "nicht angeschlossen" ' Freitxt hier eintragen
    ZSB_KLAttribute = Btxt & Ctxt & Dtxt & Etxt & Ftxt & Gtxt & Htxt
    ManipulationZiffer (ZSB_KLAttribute)  ' Wert 0 = keine BGER Werte vorhanden
    Call Tmp_BGR_Korrektur ' tmp BGER ID korriegieren Urban 02.03.2022
    Cells(aktive_Reihe, 33).Value = ZSB_KLziffer
    'Cells(aktive_Reihe, 226).Value = "tmp_" & ZSB_KLziffer 'Urban 01.07.2019
    'Urban 01.03.2022
    MsgBox ("Wichtig! Falls der Pin intern beschaltet ist" & vbCrLf & ".. bitte zuerst einen Pintyp festlegen," & vbCrLf & vbCrLf & "please at first define a pintype" & vbCrLf & "Exit! this Function")
    Exit Sub
  End If
  If NoPInDflt = "n.c.#" And NoPIn2Use <> "Nicht verbunden" Then
     Cells(aktive_Reihe, 136).Value = ""
     Cells(aktive_Reihe, 15).Value = "n.c."
  End If
     'If defKLinhalt <> "ND" Then
     frage = "yes" ' totgeschaltet
     'frage = Application.InputBox(Prompt:="Do you want to change actual CL-Info Definition (yes/no)?", Default:="yes", Left:=20, Top:=40, Type:=2)
     If frage = False Then Exit Sub
        If frage = "yes" Then
           Cells(aktive_Reihe, 33).Font.Color = RGB(0, 0, 255)            '.Value = "ND" ' aktive Zeile leeren
           Cells(aktive_Reihe, aktive_Spalte).Font.Color = RGB(0, 0, 255) '.Value = "n.c." ' aktive Zeile leeren
           Else: frage = "no"
        End If

 ActiveWorkbook.Sheets("Formblatt").Cells(aktive_Reihe, aktive_Spalte - 8).Select ' aktuellen Pin zur Orientierung selektieren
 'Urban 04.09.2019 Reduzierung der Funktion vom User gewollt
 BackStop = Cells(1, 33).Value
If BackStop = "x" Or BackStop = "X" Then
      ReadyDB = "closed"
      ' Urban 21.03.2023
      Range("State").Value = "" ' Eingaben geprüft zurücksetzen
      Range("State").Interior.Color = RGB(204, 255, 255)
      Range("DataChecked").Value = "ungeprüft"
      Range("DataChecked").Interior.Color = RGB(204, 255, 255)
      Tabelle1.CommandButton2.Caption = "CheckMe"
      Tabelle1.CommandButton2.ForeColor = blue
      MsgBox ("Data Changes are possible: Bitte erneut prüfen")
      Exit Sub
End If
 'urban
 If ReadyDB = "" Then ReadyDB = "crashed"
     If ReadyDB = "open" Then
        Application.Cursor = xlWait
         StarteImport_CSVdatei 'UGU, neue Funktion des Datenimports
         Application.Cursor = xlDefault
     End If
    If ReadyDB = "imported" Then
      Application.StatusBar = "CL-Info Databases OK"
    End If
'Import_CL_Function_Utiliti_list 'Funktion zum Daten einlesen starten
 If ReadyDB = "crashed" Then
    Application.StatusBar = "CL-Info Import Databases is crashed, ReImport/Extract VOBES Formblatt CL_InfoDB.zip file ... now"
'CL_InfoDB.zip entpacken und lokal speichern
    Datenbasen_aus_archiv_entpacken
 End If
 If ReadyDB = "crashed" Then Exit Sub ' das Entpacken ist Fehlgeschlagen
 manipulated = False ' Manipulations Status zurücksetzen
 UserForm2.Show 'Userform sichtbar schalten

Errorhandler:
    If Application.Visible = False Then Application.Visible = True
    Fehler = 0
 Exit Sub
End Sub

Public Sub Klemmeninfo_delete(delsingle)
 'tempId = "tmp_" & Int((99000 * Rnd) + 1)
If delsingle <> "NoGui" Then
delsingle = Application.InputBox(Prompt:="[all] Lösche/Delete Pin Informations, push the button <OK>" & vbCrLf & _
            "[aA-zZ] will delete <Single> Pin Informations", Default:="all", Top:=0, Left:=0, Type:=2)
End If
    
    origC = Range("PinsUpper").Column + 14 'ActiveCell.Column
    OrigRu = Range("PinsUpper").Row + 1 'ActiveCell.Row
    OrigRl = Range("PinsLower").Row - 1 'ActiveCell.Row
If delsingle = "all" Or delsingle = "NoGui" Then
     Range(Cells(OrigRu, origC), Cells(OrigRl, origC)).Value = "" ' Tabelle leeren
     Range(Cells(OrigRu, origC), Cells(OrigRl, origC)).Font.Color = RGB(0, 0, 0)
     Range(Cells(OrigRu, 15), Cells(OrigRl, 15)).Value = "n.c." ' aktive Zeile leeren
     Range(Cells(OrigRu, 15), Cells(OrigRl, 15)).Font.Color = RGB(0, 0, 0)
     Range(Cells(OrigRu, 15), Cells(OrigRl, 15)).Interior.Color = RGB(204, 255, 255)
     Range(Cells(OrigRu, 33), Cells(OrigRl, 33)).Value = "ND" ' aktive Zeile leeren
     Range(Cells(OrigRu, 33), Cells(OrigRl, 33)).Font.Color = RGB(0, 0, 0)
     Range(Cells(OrigRu, 34), Cells(OrigRl, 34)).Value = "" ' aktive Zeile leeren
     Range(Cells(OrigRu, 34), Cells(OrigRl, 34)).Font.Color = RGB(0, 0, 0)
     Range(Cells(OrigRu, 41), Cells(OrigRl, 41)).Value = "" ' aktive Zeile leeren
     Range(Cells(OrigRu, 41), Cells(OrigRl, 41)).Font.Color = RGB(0, 0, 0)
     Range(Cells(OrigRu, 52), Cells(OrigRl, 52)).Value = "Nicht verbunden" ' aktive Zeile leeren
     Range(Cells(OrigRu, 52), Cells(OrigRl, 52)).Font.Color = RGB(0, 0, 0)

     If delsingle <> "NoGui" Then
        Range(Cells(OrigRu, 33), Cells(OrigRl, 33)).Value = "ND" ' aktive Zeile leeren
        Range(Cells(OrigRu, 33), Cells(OrigRl, 33)).Font.Color = RGB(0, 0, 0)
        Range(Cells(OrigRu, 34), Cells(OrigRl, 34)).Value = "" ' aktive Zeile leeren
        Range(Cells(OrigRu, 34), Cells(OrigRl, 34)).Font.Color = RGB(0, 0, 0)
        Range(Cells(OrigRu, 41), Cells(OrigRl, 41)).Value = "" ' aktive Zeile leeren
        Range(Cells(OrigRu, 41), Cells(OrigRl, 41)).Font.Color = RGB(0, 0, 0)
        Range(Cells(OrigRu, 52), Cells(OrigRl, 52)).Value = "Nicht verbunden" ' aktive Zeile leeren
        Range(Cells(OrigRu, 52), Cells(OrigRl, 52)).Font.Color = RGB(0, 0, 0)
        Range(Cells(OrigRu, 78), Cells(OrigRl, 78)).Font.Color = RGB(0, 0, 0) ' I1 Strom
     End If
  Range("State").Value = "" ' Freigabestatus zurücksetzen
  Range("State").Interior.Color = RGB(204, 255, 255)
  Range("State").Font.Color = RGB(0, 0, 0)
'UGU, Werte der BGER Pin im gesamten Bereich löschen
  Range(Cells(OrigRu, 136), Cells(OrigRl, 183)).Value = "" ' aktive Zeile leeren
  Range(Cells(OrigRu, 136), Cells(OrigRl, 183)).Font.Color = RGB(0, 0, 0)
  Range(Cells(OrigRu, 226), Cells(OrigRl, 226)).Value = ""
  Range(Cells(OrigRu, 226), Cells(OrigRl, 226)).Interior.Color = RGB(255, 255, 255)
  Sheets("Formblatt").Cells(OrigRu, origC - 8).Select
  Sheets("Formblatt").Cells(OrigRu, origC).Interior.Color = RGB(204, 255, 255)
End If
If delsingle = False Then
       Call Tabelle1.Button1_Standard
       Call Tabelle1.Button2_Standard
       Call Tabelle1.Button3_Standard
       Call Tabelle1.Button4_Standard
       Call Tabelle1.Button5_Standard
       Call Tabelle1.Button6_Standard
       Exit Sub
End If
If delsingle <> "all" And delsingle <> "NoGui" Then
aktivCellR = ActiveCell.Row 'Reihe
   If aktivCellR < OrigRu Or aktivCellR > OrigRl Then ' Pin nicht im Pinbereich?
     MsgBox "Bitte einen vorhanden Pin selektieren/Please choose an E-Komp-Pin. Ihre Auswahl/selected cell, ist außerhalb des PinBereichs, Abbruch!"
     Call Tabelle1.Button1_Standard
     Call Tabelle1.Button2_Standard
     Call Tabelle1.Button3_Standard
     Call Tabelle1.Button4_Standard
     Call Tabelle1.Button5_Standard
     Call Tabelle1.Button6_Standard
     Exit Sub
   End If
    
    Cells(aktivCellR, 15).Value = "n.c." ' Zeile leeren
    Cells(aktivCellR, 15).Font.Color = RGB(0, 0, 0)
    Cells(aktivCellR, 15).Interior.Color = RGB(204, 255, 255)
    Cells(aktivCellR, 33).Value = "ND" ' aktive Zeile leeren
    Cells(aktivCellR, 33).Font.Color = RGB(0, 0, 0)
    Cells(aktivCellR, 34).Value = "" ' aktive Zeile leeren
    Cells(aktivCellR, 34).Font.Color = RGB(0, 0, 0)
    Cells(aktivCellR, 41).Value = "" ' aktive Zeile leeren
    Cells(aktivCellR, 41).Font.Color = RGB(0, 0, 0)
    Cells(aktivCellR, 52).Value = "Nicht verbunden" ' aktive Zeile leeren
    Cells(aktivCellR, 52).Font.Color = RGB(0, 0, 0)
    Cells(aktivCellR, 78).Font.Color = RGB(0, 0, 0) ' I1 Strom
    'UGU, Werte der BGER Pin im gesamten Bereich löschen Cells(OrgRu + dLooP, 226)
    Range(Cells(aktivCellR, 136), Cells(aktivCellR, 183)).Value = "" ' aktive Zeile leeren
    Range(Cells(aktivCellR, 136), Cells(aktivCellR, 183)).Font.Color = RGB(0, 0, 0)
    'Range(Cells(OrigRu, 226), Cells(OrigRl, 226)).Value = tempId ' temp.BGER Id festlegenD
    Range(Cells(OrigRu, 226), Cells(OrigRl, 226)).Value = ""
    Range(Cells(OrigRu, 226), Cells(OrigRl, 226)).Interior.Color = RGB(255, 255, 255)
    Cells(aktivCellR, 7).Select
    Cells(aktivCellR, 15).Interior.Color = RGB(204, 255, 255)
    
End If

'Einflüsse von außen verändern die Höhe der Buttons G.Urban EEKK1, deswegen Größen zurücksetzen
Call Tabelle1.Button1_Standard
Call Tabelle1.Button2_Standard
Call Tabelle1.Button3_Standard
Call Tabelle1.Button4_Standard
Call Tabelle1.Button5_Standard
Call Tabelle1.Button6_Standard

End Sub

Public Sub finde_Cl_info()
  ' Datenbasen sind eingelesen und befinden sich im Speicher
  CL_2searchTxt = Application.InputBox(Prompt:="Please enter your Text2Find: ", Default:="Searchtext", Left:=50, Top:=40, Type:=2)
  CL_2searchLen = Len(CL_2searchTxt)
  SpeicherTxt = FunctionsSpeicher ' Default Datenbasis
    If CL_2searchTxt <> "" Then
        For CL_S = 0 To UBound(SpeicherTxt, 1)
          Cl_Mem_function = SpeicherTxt(CL_S, 0) 'cl_function
          CL_Mem_functTxt = SpeicherTxt(CL_S, 1) 'cl_descriptiveTextEN
          CL_Mem_functTxt_len = Len(CL_Mem_functTxt)
             If CL_Mem_functTxt_len >= CL_2searchLen Then
                 For_len = CL_Mem_functTxt_len - CL_2searchLen + 1 '+1 damit bis zum Ende des Speicherwertes gesucht wird
              Else: CL_S = CL_S + 1
             End If
             For X = 1 To For_len ' Innere Schleife um im Datenwert zu suchen
                CuttedTxt = Mid(CL_Mem_functTxt, X, CL_2searchLen) 'Teilwort x=Zeichen Start, CL_2searchLen=Wortlänge
                If CL_2searchTxt = CuttedTxt Then
                  X = For_len ' Abbruch durch Schleifenende
                End If
             Next X
          If CL_2searchTxt = CL_Mem_functTxt Or CL_2searchTxt = CuttedTxt Then
            foundTxt = CL_Mem_functTxt & ":  CL_function= " & Cl_Mem_function
            frage = Application.InputBox(Prompt:="Is search result for CL_Function OK then" & vbCrLf & "  push button <OK>" & vbCrLf & _
            "    else <Abbrechen>" & vbCrLf & _
            "                      for next search result) ?", Default:=foundTxt, Top:=0, Left:=0, Type:=2)
            If frage = ":  CL_function= " Then Exit Sub
            frage_len = Len(frage)
              If frage_len > 6 Then
              UserForm2.CoBoFunktion.Text = Cl_Mem_function
              Funktion = Cl_Mem_function
              UserForm2.clampCoBo.Text = Cl_Mem_function
              frage = "yes"
                Exit Sub
              End If
              If Len(frage) = "" Then
                CL_S = CL_S + 1
              End If
          End If
        Next CL_S
    End If
    foundTxt = "Search function im database"
    'Application.Visible = True
    
End Sub

Public Sub MakeOpenDB_Import()
ReadyDB = "open"
End Sub

Public Sub DontDoThat()
If Tabelle1.NEWpins <> "gen" Then
  MsgBox "In diesem Bereich ist eine Eingabe von Werten nicht vorgesehen!" & vbCrLf _
  & "To change values, please push the button [ PinBeschreibung ]" & vbCrLf _
  & vbCrLf & "don´t try to change values in this cell / range", Title:="Forbidden Area"
End If
End Sub

Public Sub ExChange()
'erzeugt Pinlisten zum Verschieben von Pins
'Steckplz B17; Pin G17
 OrigRup = Range("PinsUpper").Row + 1 'ActiveCell.Row
 OrigRlp = Range("PinsLower").Row - 1 'ActiveCell.Row
 PinAnzahl = Range("SumPins")
   For i = 0 To PinAnzahl - 1
        PinName = Cells(OrigRup + i, 2) & Cells(OrigRup + i, 7)
        Pin_msg = Pin_msg & PinName & "," ' Fragewort
   Next i
    With Range("StartPin").Validation
    'With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=Pin_msg
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = "database must be corrupt"
        .ShowInput = True
        .ShowError = True
    End With
With Range("EndPin").Validation
    'With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=Pin_msg
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
Public Sub ExChangeGER()
'erzeugt Pinlisten zum Verschieben von Pins
'Steckplz B17; Pin G17
 OrigRup = Range("SlotsUpper").Row + 1 'ActiveCell.Row
 OrigRlp = Range("SlotsLower").Row - 1 'ActiveCell.Row
 GerAnzahl = Range("SumSlots")
   For i = 0 To GerAnzahl - 1
        GerName = Cells(OrigRup + i, 2) & Cells(OrigRup + i, 7)
        Ger_msg = Ger_msg & GerName & "," ' Fragewort
   Next i
    With Range("GER").Validation
    'With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=Ger_msg
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

Public Sub ExchangePins()
'verschiebt sicher von Pin nach Pin
'Steckplz B17; Pin G17
Dim Tmp_FromPDatas(25)
Dim Tmp_ToPinDatas(25)
 OrigRup = Range("PinsUpper").Row + 1 'ActiveCell.Row
 OrigRlp = Range("PinsLower").Row - 1 'ActiveCell.Row
 PNumber = Range("Teilenummer").Value ' Teilenummer des BT
 PinAnzahl = Range("SumPins")
 FromPin = Range("StartPin").Value
 ToPin = Range("EndPin").Value
 EndeC = Range("CmtLower").Row
 If FromPin = ToPin Or FromPin = "m" Or ToPin = "n" Then ' darf nur bei echten Pins stattfinden
   MsgBox ("Bitte (unterschiedliche) Pins wählen, please select different Pins: Abort")
     Exit Sub
 End If
   For i = 0 To PinAnzahl - 1
        PinName = Cells(OrigRup + i, 2) & Cells(OrigRup + i, 7)
        If PinName = FromPin Then
           X = i
           Range(Cells(OrigRup + X, 10), Cells(OrigRup + X, 230)).Copy
           'FromPinRstart = Range(Cells(OrigRup + x, 2), Cells(OrigRup + x, 230)).Copy
           Range(Cells(EndeC + 10, 10), Cells(EndeC + 10, 230)).PasteSpecial
           'FromPinRoff = Range(Cells(EndeC + 10, 2), Cells(EndeC + 10, 230)).PasteSpecial
           Cells(1, 1).Select
        End If

        If PinName = ToPin Then
          Y = i
          Range(Cells(OrigRup + Y, 10), Cells(OrigRup + Y, 230)).Copy
          Range(Cells(EndeC + 11, 10), Cells(EndeC + 11, 230)).PasteSpecial
          Cells(1, 2).Select
        End If
   Next i
          Range(Cells(EndeC + 10, 10), Cells(EndeC + 10, 230)).Copy            'FromPinRoff
          Tabelle1.NEWpins = "edit"
          Range(Cells(OrigRup + Y, 10), Cells(OrigRup + Y, 230)).PasteSpecial  'ToPinRstart
          'BGER ID neu erzeugen je Pin da zwischendurch manipulisert sein konnte
            Steckplz = Sheets("Formblatt").Cells(OrigRup + Y, 2).Value 'Steckplatz holen
            ZSBziffer = 0
            ZSBSTPN = Steckplz & PNumber
            For ai = 1 To Len(ZSBSTPN)
               ZWwert = Mid(ZSBSTPN, ai, 1)
               ZSBziffer = ZSBziffer + Asc(ZWwert)
            Next
            tempId = "tmp_" & ZSBziffer
            Cells(OrigRup + Y, 226).Value = tempId  ' temporäre BGER ID vergeben"
            Cells(OrigRup + Y, 226).Interior.Color = RGB(255, 255, 204)
          'Ende
          Range(Cells(EndeC + 11, 10), Cells(EndeC + 11, 230)).Copy            'ToPinRoff
          Tabelle1.NEWpins = "edit"
          Range(Cells(OrigRup + X, 10), Cells(OrigRup + X, 230)).PasteSpecial  'FromPinRstart
          'BGER ID neu erzeugen je Pin da zwischendurch manipulisert sein konnte
            Steckplz = Sheets("Formblatt").Cells(OrigRup + X, 2).Value 'Steckplatz holen
            ZSBziffer = 0
            ZSBSTPN = Steckplz & PNumber
            For bi = 1 To Len(ZSBSTPN)
               ZWwert = Mid(ZSBSTPN, bi, 1)
               ZSBziffer = ZSBziffer + Asc(ZWwert)
            Next
            tempId = "tmp_" & ZSBziffer
            Cells(OrigRup + X, 226).Value = tempId  ' temporäre BGER ID vergeben"
            Cells(OrigRup + X, 226).Interior.Color = RGB(255, 255, 204)
          'Ende
          'hier wird der Zwischenbereich gelöscht
          Range(Cells(EndeC + 10, 10), Cells(EndeC + 11, 230)).Delete
          Application.ActiveWorkbook.Worksheets("Formblatt").Cells(1, 1).Select
End Sub

Public Sub MakeIntCL_Link()
Dim ZellenBereich As Range
pos = 0
pos_vorher = 0
Treffer = 0
Dim PosSp(20, 2) ' 0 = Reihe; 1= LinkNamen; 2= Pin Direction (Spalte 158) ; Spalte ist statisch = 134
Dim PinsAdress As String
ROWlower = Range("PinsLower").Row - 1 'ActiveCell.Row
ROWuper = Range("PinsUpper").Row + 1
'PinAnzahl = Range("SumPins")
On Error GoTo Abgebrochen
ActiveWindow.ScrollColumn = 120 ' Scrollbar nach recht zu BGER verschhieben
Set ZellenBereich = Application.InputBox(Prompt:="Bitte markieren Sie / please select .. alle zu verbindenden / to link E-Komponenten Pins:", Type:=8)
PinsAdress = ZellenBereich.Address(RowAbsolute:=False, ColumnAbsolute:=False)
LnPinsAdr = Len(PinsAdress) ' Länge des Pinbereichs
'PinLink benennen und prüfen
  LinkN = Application.InputBox(Prompt:="Bitte eine LinkNamen [1-999] vergeben, please define a Pin LinkName:", Type:=2)
   LnLinkN = Len(LinkN) ' wenn größer 3 dann 4-Stellig
  If LinkN = "" Or LnLinkN > 3 Then
    MsgBox ("Bitte nur eine Ziffer bis max. 999 eingeben / wrong value: Abbruch")
    GoTo Abgebrochen
    End If
LinkNameNew = "#" & LinkN
For N = 0 To LnPinsAdr
'Analyse der Separatoren
  pos = InStr(pos_vorher + 1, PinsAdress, ",", 1) ' 1 = Komma als Suchtext
  If pos > 0 Then ' Treffer
     Treffer = Treffer + 1
     Start = 1 + pos_vorher
     Lweite = pos - Start
     PinsAdressCut = Mid(PinsAdress, Start, Lweite)
     PinsAdrCutRow = Range(PinsAdressCut).Row    ' die Reihe der Zelle  ermitteln
'Check auf eine Zelle oder einen Bereich
        pos_DP = InStr(1, PinsAdressCut, ":", 1) ' 1 = Doppelpunkt als Suchtext
        If pos_DP > 0 Then
            'MsgBox "Doppelpunkt in: " & PinsAdressCut & " an der Stelle " & pos_DP & " gefunden"
'Bereich durch einen ":" wurde erkannt, in Zellen schneiden
             LnPinsAdr2 = Len(PinsAdressCut) ' Länge des Bereichstextes
             Start = 1
             Lweite = pos_DP - Start
             Lweite2 = LnPinsAdr2 - pos_DP
           FromAdr = Mid(PinsAdressCut, Start, Lweite)
' Bereich zerlegen
           ToAdr = Mid(PinsAdressCut, pos_DP + 1, Lweite2)
              VonReihe = Range(FromAdr).Row
              VonSpalte = "134" 'Spalte Pin ind der BgerTabelle; Range(FromAdr).Column
              NachReihe = Range(ToAdr).Row
              NachSpalte = "134" 'Spalte Pin ind der BgerTabelle
'feststellen wieviel Zellen sich ergeben
             AnzahlZellen = NachReihe - VonReihe
             If pos > 0 Then Treffer = Treffer - 1
             For m = 0 To AnzahlZellen
'Zellen aufbereiten
                Treffer = Treffer + 1
                Reihe = VonReihe + m
                Spalte = "134"
                PosSp(Treffer, 0) = Reihe                    ' PinReihe
                PosSp(Treffer, 1) = Cells(Reihe, 139).Value  ' LinkName
                PosSp(Treffer, 2) = Cells(Reihe, 158).Value  ' Pin Richtung
                
             Next 'm

        End If 'pos_DP > 0
'Speichern wenn kein Bereich erkannt wurde
      If pos_DP = 0 And pos > 0 Then
         PosSp(Treffer, 0) = PinsAdrCutRow
         PosSp(Treffer, 1) = Cells(PinsAdrCutRow, 139).Value ' LinkName
         PosSp(Treffer, 2) = Cells(PinsAdrCutRow, 158).Value ' Pin Richtung
      End If
      pos_vorher = pos ' Merker vorherige Position
  End If 'pos > 0
Next 'n
'Restdaten ermitteln und auch auf Doppelpunkt prüfen
  Treffer = Treffer + 1
  PinsAdressCutEnd = Mid(PinsAdress, 1 + pos_vorher, LnPinsAdr - pos_vorher)
  PinsAdrCutEndRow = Range(PinsAdressCutEnd).Row ' dei Reihe der Zelle  ermitteln
'Check auf eine Zelle oder einen Bereich
        pos_DP = InStr(1, PinsAdressCutEnd, ":", 1) ' 1 = Doppelpunkt als Suchtext
        If pos_DP > 0 Then
            'MsgBox "Am Ende > Doppelpunkt in: " & PinsAdressCutEnd & " an der Stelle " & pos_DP & " gefunden"
'Bereich durch einen ":" wurde erkannt, in Zellen schneiden
             LnPinsAdr2 = Len(PinsAdressCutEnd) ' Länge des Bereichstextes
             Start = 1
             Lweite = pos_DP - Start
             Lweite2 = LnPinsAdr2 - pos_DP
           FromAdr = Mid(PinsAdressCutEnd, Start, Lweite)
           ' Bereich zerlegen
           ToAdr = Mid(PinsAdressCutEnd, pos_DP + 1, Lweite2)
              VonReihe = Range(FromAdr).Row
              VonSpalte = "134" 'Spalte Pin ind der BgerTabelle; Range(FromAdr).Column
              NachReihe = Range(ToAdr).Row
              NachSpalte = "134" 'Spalte Pin ind der BgerTabelle
             'feststellen wieviel Zellen sich ergeben
             AnzahlZellen = NachReihe - VonReihe
             For m = 0 To AnzahlZellen
              'Zellen aufbereiten
                Reihe = VonReihe + m
                Spalte = "134"
                PosSp(Treffer, 0) = Reihe                    ' PinReihe
                PosSp(Treffer, 1) = Cells(Reihe, 139).Value  ' LinkName
                PosSp(Treffer, 2) = Cells(Reihe, 158).Value  ' Pin Richtung
                If m < AnzahlZellen Then Treffer = Treffer + 1
             Next 'm
        End If 'pos_DP > 0
      If pos_DP = 0 And PinsAdrCutEndRow <> "" Then
        PosSp(Treffer, 0) = PinsAdrCutEndRow
        PosSp(Treffer, 1) = Cells(PinsAdrCutEndRow, 139).Value ' LinkName
        PosSp(Treffer, 2) = Cells(PinsAdrCutEndRow, 158).Value ' Pin Richtung
      End If
  
'CL_LINK analysieren
 PinDirOK = False                'Kein PinLink eines Pins mit der Direction "IN" enthalten
 For CLL = 1 To Treffer
   ChkDirIN = PosSp(CLL, 2)     ' Pin Richtung
   ChkDirLnk = PosSp(CLL, 1)    ' LinkName
   If ChkDirIN = "IN" Then
       PinDirOK = True          ' PinLink eines Pins mit der Direction "IN" enthalten
        LinkName = ChkDirLnk    ' LinkName merken, der letzte gewinnt
   End If
   'MsgBox ("Zeile: " & PosSp(CLL, 0) & " CL_LINK: " & LinkName & " Direction: " & PosSp(CLL, 2))
 Next
 
 'CL_Link an betroffene Pins in der BGER Tabelle eintragen
 If PinDirOK = True Then
    For CLM = 1 To Treffer
      WZeile = PosSp(CLM, 0)                      'aktuelle PinZeile
      If WZeile < ROWuper Then
         MsgBox ("Abbruch: Bin oberhalb der Pintabelle! Out of PinRange")
         GoTo Abgebrochen
      End If
      If WZeile > ROWlower Then
         MsgBox ("Abbruch: Bin unterhalb der Pintabelle, der neue PinLink" & vbCrLf & " bleibt innerhalb der Pintabelle! Out of PinRange")
         GoTo Abgebrochen  'ist keine Pinzeile
      End If
      WSpalte = 139                               'Zelle Spalte 139 ist CL_LINK
      Cells(WZeile, WSpalte).Value = LinkNameNew  'gewünschter LinkName
    Next
    MsgBox ("Es wurde die interne Verbindung [" & LinkName & "] mit dem Namen: " & vbCrLf & "CL_LINK: " & LinkNameNew & " vorgenommen (PinLink successful created).")
    Else: MsgBox ("Es konnte kein PinLink hergestellt werden," & vbCrLf & "es fehlte ein Pin mit der PinRichtung (IN)" & _
    vbCrLf & "Sorry can´t link selected pins, no pin with direction (IN)")
 
 End If
 
Abgebrochen:
ActiveWindow.ScrollColumn = 1

End Sub

Public Sub Datumfeststellen()
userTmpVerz = Environ("USERPROFILE") & "\AppData\Local\Temp\"
Pfad = ActiveWorkbook.Path
Name = ActiveWorkbook.Name
Filename = Pfad & "\" & Name
Zeitangabe = FileDateTime(Filename)    ' Liefert "12.02.1993 16:35:47".
End Sub

Public Sub Anbindung_select()
TmpKLinfo = ""
TmpLAH = ""
TmpKLz = ""
BN_LAH_ID = ""
AnBAnfg = ""
ASILevel = "no"
CL_SeC = False
CL_Stat = False

  AbgAktive_R = ActiveCell.Row
  AnBAnfg = Cells(AbgAktive_R, 34).Value 'ActiveCell.Value
  BN_LAH_ID = Cells(AbgAktive_R, 41).Value

  CL_Txt = Cells(AbgAktive_R, 136).Value   ' Klemme
  If CL_Txt <> "" Then CL_Stat = True      ' eine Klemme ist vorhanden
  CL_secure = Right(CL_Txt, 1) ' q oder s
  If CL_secure = "s" Or CL_secure = "q" Then CL_SeC = True     ' eine Klemme ist vorhanden
  FuncTxt = Cells(AbgAktive_R, 143).Value  ' Funktion
  If FuncTxt <> "" Then Fc_stat = True     ' eine Funktion ist vorhanden
  UtilTxt = Cells(AbgAktive_R, 152).Value  ' Örtlichkeit
  If UtilTxt <> "" Then Ut_stat = True     ' eine Örtlichkeit ist vorhanden
  UserForm1.Show

    ' Urban 20.07.20 * kennzeichnet es ist eine Bordnetzanbindung/ LAH Info vorhanden
   Cells(AbgAktive_R, 41).Value = BN_LAH_ID
   TmpKLinfo = Cells(AbgAktive_R, 15).Value
   TmpKLz = Left(TmpKLinfo, 1)
   TmpLAH = BN_LAH_ID
   'TmpLAH = Cells(AbgAktive_R, 41).Value ' Lastenheft Referenz holen
   Cells(AbgAktive_R, 34).Value = AnBAnfg

   'Suche ASIL Level
   'Urban 28.03.2020
    Len_AnBAnfg = Len(AnBAnfg) 'Länge des Kommentar feststellen
    TxtCut = Left(AnBAnfg, 4)
    TxtCut1 = Left(AnBAnfg, 3)
    TxtCut2 = Left(AnBAnfg, 2)
    TxtCut3 = Left(AnBAnfg, 6)
      'Schleife findet Trennzeichen
      AnB_pos = InStr(1, AnBAnfg, ";") 'Suche nach einem Trennzeichen ";" im String der Anbindungsanforderungen
      If AnB_pos = 0 And Len_AnBAnfg >= 1 Then
          ASILevel = "BNA"
        If TxtCut = "ASIL" Then
          ASILevel = "yes"
          ASIL_Val = TxtCut3
        End If
        If TxtCut = "ADAP" Then
          ASILevel = "yes"
          ASIL_Val = TxtCut
        End If
        If TxtCut1 = "HAF" Then
          ASILevel = "yes"
          ASIL_Val = TxtCut1
        End If
        If TxtCut2 = "QM" Then
          ASILevel = "yes"
          ASIL_Val = TxtCut2
        End If
        'Text_Auswertung (AnBAnfg)
      End If
      i = 1
      len_AnBtxt = 0
    While AnB_pos > 0
      'Text herausholen
      AnB_pos = InStr(i, AnBAnfg, ";")
      If AnB_pos > 0 Then AnB_pos_pre = AnB_pos

      If AnB_pos <> 0 Then
        len_AnBtxt = AnB_pos - i
        tmp_Txt = Mid(AnBAnfg, i, len_AnBtxt)
        If ASILevel <> "yes" Then ASILevel = "BNA"
      End If

      If AnB_pos = 0 And AnB_pos_pre > 0 Then
         'Ausnahme bei Länge Null
         len_AnBtxt = Len_AnBAnfg - AnB_pos_pre
         tmp_Txt = Mid(AnBAnfg, AnB_pos_pre + 1, len_AnBtxt)
      End If
       i = AnB_pos + 1
       'tmp_Txt auswerten LayL, Qmin LayL40mm Qmin-0.82 ASIL_CB.Value
       'Text_Auswertung (tmp_Txt) ' Sprung zur Funktion Text_Auswertung
       'ASIL_<A_D> feststellen
       ASIL_pos = InStr(1, tmp_Txt, "ASIL")
       If ASIL_pos > 0 Or tmp_Txt = "ADAP" Or tmp_Txt = "HAF" Or tmp_Txt = "QM" Then
          ASILevel = "yes"
          ASIL_Val = tmp_Txt
       End If
    Wend
   'Urban Backstop erkennen
   BackStop = Cells(1, 33).Value ' Urban 08.04.2022
   If BackStop = "x" Or BackStop = "X" Then
      SternChk = Left(Cells(AbgAktive_R, 15).Value, 1) ' Sternchen erkennen
      If SternChk = "*" Then IsStern = "ok"
      If ASILevel <> "no" And IsStern <> "ok" Or TmpLAH <> "" And IsStern <> "ok" Then
          zsb_value = "*" & Cells(AbgAktive_R, 15).Value
          Cells(AbgAktive_R, 15).Value = zsb_value
      End If
      If ASILevel <> "no" And IsStern = "ok" Or TmpLAH <> "" And IsStern = "ok" Then
          zsb_value = Cells(AbgAktive_R, 15).Value
          Cells(AbgAktive_R, 15).Value = zsb_value
      End If
      If ASILevel = "no" And TmpLAH = "" And IsStern = "ok" Then
          Cells(AbgAktive_R, 15).Value = Mid(Cells(AbgAktive_R, 15).Value, 2)
      End If
      If ASILevel = "no" And TmpLAH = "" And IsStern <> "ok" Then
          zsb_value = Cells(AbgAktive_R, 15).Value
          Cells(AbgAktive_R, 15).Value = zsb_value
      End If
      Cells(AbgAktive_R, 7).Select
      UserForm1.Hide
     Exit Sub
   End If
   
   'Urban 28.03.2020
     If ASILevel = "BNA" And TmpLAH <> "" And CL_Stat = True And Fc_stat = True And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & FuncTxt & "." & UtilTxt
     End If
     If ASILevel = "BNA" And TmpLAH = "" And CL_Stat = True And Fc_stat = True And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & FuncTxt & "." & UtilTxt
     End If
     If ASILevel = "BNA" And TmpLAH <> "" And CL_Stat = True And Fc_stat = True And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & FuncTxt
        If CL_SeC = True Then MsgBox ("Sichere Klemme " & CL_Txt & " definiert und keinen ASIL Level definiert!" & vbCrLf & "Secured Clamp and ASIL Level value undefined!")
     End If
     If ASILevel = "BNA" And TmpLAH = "" And CL_Stat = True And Fc_stat = True And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & FuncTxt
        If CL_SeC = True Then MsgBox ("Sichere Klemme " & CL_Txt & " definiert und keinen ASIL Level definiert!" & vbCrLf & "Secured Clamp and ASIL Level value undefined!")
     End If
     
     If ASILevel = "BNA" And TmpLAH <> "" And CL_Stat = True And Fc_stat = False And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & "." & UtilTxt
     End If
     
     If ASILevel = "BNA" And TmpLAH = "" And CL_Stat = True And Fc_stat = False And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & "." & UtilTxt
     End If
     
     If ASILevel = "BNA" And TmpLAH <> "" And CL_Stat = False And Fc_stat = True And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & FuncTxt & "." & UtilTxt
     End If
     If ASILevel = "BNA" And TmpLAH = "" And CL_Stat = False And Fc_stat = True And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & FuncTxt & "." & UtilTxt
     End If
     If ASILevel = "BNA" And TmpLAH = "" And CL_Stat = False And Fc_stat = True And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = "*" & FuncTxt & "." & UtilTxt
     End If
     If ASILevel = "yes" And TmpLAH = "" And CL_Stat = True And CL_SeC = False And Fc_stat = True And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & FuncTxt & "." & UtilTxt
     End If
     If ASILevel = "yes" And TmpLAH = "" And CL_Stat = True And CL_SeC = False And Fc_stat = True And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & FuncTxt
        MsgBox ("Sichere Klemme notwendig?! Sie haben einen ASIL Level definiert? ASIL Level value defined!")
     End If
     If ASILevel = "yes" And TmpLAH = "" And CL_Stat = True And CL_SeC = False And Fc_stat = False And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "." & UtilTxt
        MsgBox ("Sichere Klemme notwendig?! Sie haben einen ASIL Level definiert? ASIL Level value defined!")
     End If
     If ASILevel = "yes" And TmpLAH = "" And CL_Stat = True And CL_SeC = False And Fc_stat = False And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt
        MsgBox ("Sichere Klemme notwendig?! Sie haben einen ASIL Level definiert? ASIL Level value defined!")
     End If
     If ASILevel = "yes" And TmpLAH = "" And CL_Stat = True And CL_SeC = True And Fc_stat = False And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "_" & ASIL_Val
     End If
     If ASILevel = "yes" And TmpLAH <> "" And CL_Stat = True And CL_SeC = False And Fc_stat = True And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & FuncTxt & "." & UtilTxt
     End If
     If ASILevel = "yes" And TmpLAH <> "" And CL_Stat = True And CL_SeC = False And Fc_stat = True And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & FuncTxt
     End If
     If ASILevel = "yes" And TmpLAH <> "" And CL_Stat = True And CL_SeC = False And Fc_stat = False And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & "." & UtilTxt
     End If
     If ASILevel = "yes" And TmpLAH <> "" And CL_Stat = True And CL_SeC = False And Fc_stat = False And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#"
     End If
     If ASILevel = "yes" And TmpLAH = "" And CL_Stat = True And CL_SeC = True And Fc_stat = True And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & FuncTxt & "." & UtilTxt & "_" & ASIL_Val
     End If
     If ASILevel = "yes" And TmpLAH = "" And CL_Stat = True And CL_SeC = True And Fc_stat = True And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & FuncTxt & "_" & ASIL_Val
     End If
     If ASILevel = "yes" And TmpLAH = "" And CL_Stat = True And CL_SeC = True And Fc_stat = False And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "." & UtilTxt & "_" & ASIL_Val
     End If
     If ASILevel = "yes" And TmpLAH <> "" And CL_Stat = True And CL_SeC = True And Fc_stat = True And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & FuncTxt & "." & UtilTxt & "_" & ASIL_Val
     End If
     If ASILevel = "yes" And TmpLAH <> "" And CL_Stat = True And CL_SeC = True And Fc_stat = True And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & FuncTxt & "_" & ASIL_Val
     End If
    If ASILevel = "yes" And TmpLAH <> "" And CL_Stat = True And CL_SeC = True And Fc_stat = False And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "_" & ASIL_Val
     End If

     If ASILevel = "yes" And TmpLAH <> "" And CL_Stat = True And CL_SeC = True And Fc_stat = False And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "." & UtilTxt & "_" & ASIL_Val
     End If
     If ASILevel = "yes" And TmpLAH = "" And CL_Stat = False And CL_SeC = False And Fc_stat = True And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & FuncTxt & "." & UtilTxt
     End If
     If ASILevel = "yes" And TmpLAH = "" And CL_Stat = False And CL_SeC = False And Fc_stat = True And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & FuncTxt & "." & UtilTxt
     End If
     If ASILevel = "yes" And TmpLAH = "" And CL_Stat = False And CL_SeC = False And Fc_stat = True And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = "*" & FuncTxt
     End If
     
     If ASILevel = "yes" And TmpLAH <> "" And CL_Stat = False And CL_SeC = False And Fc_stat = True And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & FuncTxt & "." & UtilTxt
     End If
     
     If ASILevel = "no" And TmpLAH = "" And CL_Stat = True And CL_SeC = False And Fc_stat = True And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = CL_Txt & "#" & FuncTxt & "." & UtilTxt
     End If
     If ASILevel = "no" And TmpLAH = "" And CL_Stat = True And CL_SeC = False And Fc_stat = True And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = CL_Txt & "#" & FuncTxt
     End If
     If ASILevel = "no" And TmpLAH = "" And CL_Stat = True And CL_SeC = False And Fc_stat = False And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = CL_Txt & "#" & "." & UtilTxt
     End If
     If ASILevel = "no" And TmpLAH = "" And CL_Stat = True And CL_SeC = False And Fc_stat = False And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = CL_Txt & "#"
     End If
     If ASILevel = "no" And TmpLAH <> "" And CL_Stat = True And CL_SeC = False And Fc_stat = True And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & FuncTxt & "." & UtilTxt
     End If
     If ASILevel = "no" And TmpLAH <> "" And CL_Stat = True And CL_SeC = False And Fc_stat = True And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & FuncTxt
     End If
     If ASILevel = "no" And TmpLAH <> "" And CL_Stat = True And CL_SeC = False And Fc_stat = False And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & "." & UtilTxt
     End If
     If ASILevel = "no" And TmpLAH <> "" And CL_Stat = True And CL_SeC = False And Fc_stat = False And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#"
     End If
     If ASILevel = "no" And TmpLAH = "" And CL_Stat = True And CL_SeC = True And Fc_stat = True And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = CL_Txt & "#" & FuncTxt & "." & UtilTxt & "_" & ASIL_Val
     End If
     If ASILevel = "no" And TmpLAH = "" And CL_Stat = True And CL_SeC = True And Fc_stat = True And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = CL_Txt & "#" & FuncTxt
        MsgBox ("Sichere Klemme! Haben Sie den ASIL Level vergessen? No ASIL Level value!")
     End If
     If ASILevel = "no" And TmpLAH = "" And CL_Stat = True And CL_SeC = True And Fc_stat = False And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = CL_Txt & "#" & "." & UtilTxt
        MsgBox ("Sichere Klemme! Haben Sie den ASIL Level vergessen? No ASIL Level value!")
     End If
     If ASILevel = "no" And TmpLAH = "" And CL_Stat = True And CL_SeC = True And Fc_stat = False And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = CL_Txt & "#"
        MsgBox ("Sichere Klemme! Haben Sie den ASIL Level vergessen? No ASIL Level value!")
     End If
     If ASILevel = "no" And TmpLAH <> "" And CL_Stat = True And CL_SeC = True And Fc_stat = True And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & FuncTxt & "." & UtilTxt & "_" & ASIL_Val
     End If
     If ASILevel = "no" And TmpLAH <> "" And CL_Stat = True And CL_SeC = True And Fc_stat = True And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & FuncTxt & "_" & ASIL_Val
     End If
     If ASILevel = "no" And TmpLAH <> "" And CL_Stat = True And CL_SeC = True And Fc_stat = False And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#" & "." & UtilTxt
        MsgBox ("Sichere Klemme! Haben Sie den ASIL Level vergessen? No ASIL Level value!")
     End If
     If ASILevel = "no" And TmpLAH <> "" And CL_Stat = True And CL_SeC = True And Fc_stat = False And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = "*" & CL_Txt & "#"
        MsgBox ("Sichere Klemme! Haben Sie den ASIL Level vergessen? No ASIL Level value!")
     End If
     If ASILevel = "no" And TmpLAH = "" And CL_Stat = False And CL_SeC = False And Fc_stat = True And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = FuncTxt & "." & UtilTxt
     End If
     If ASILevel = "no" And TmpLAH = "" And CL_Stat = False And CL_SeC = False And Fc_stat = True And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = FuncTxt
     End If
     If ASILevel = "no" And TmpLAH <> "" And CL_Stat = False And CL_SeC = False And Fc_stat = True And Ut_stat = True Then
        Cells(AbgAktive_R, 15).Value = "*" & FuncTxt & "." & UtilTxt
     End If
     If ASILevel = "no" And TmpLAH <> "" And CL_Stat = False And CL_SeC = False And Fc_stat = True And Ut_stat = False Then
        Cells(AbgAktive_R, 15).Value = "*" & FuncTxt
     End If
     
     Cells(AbgAktive_R, 7).Select
 UserForm1.Hide
End Sub

Public Sub TmpCopyOfCT()
' Ströme und Zeiten vom aktiven Pin holen
frage = "no"
OrigRu = Range("PinsUpper").Row + 1 '   ActiveCell.Row
OrigRl = Range("PinsLower").Row - 1 '   ActiveCell.Row
cnt = ActiveCell.Row
PinNr = Sheets("Formblatt").Cells(cnt, 7).Value 'PinNummer
StckPlz = Sheets("Formblatt").Cells(cnt, 2).Value ' Steckplatz
'Fenster bei A13 einfrieren
Range("A13").Select
ActiveWindow.FreezePanes = True
Application.StatusBar = "-------------------V O B E S Message ---------------- Freeze Panes @True"
Sheets("Formblatt").Cells(cnt, 7).Select 'PinNummer
  If cnt > OrigRl Or cnt < OrigRu Then
       m = MsgBox("Sorry your Selection is out of Component Pin Area, Exit", vbInformation, "Fokus Pinfestlegungen")
       Application.StatusBar = "-------------------V O B E S Message ---------------- Freeze Panes @False"
     Exit Sub
  End If
  'prüfen was getan werden soll
  chkR = mCurFild(0) 'Von Zeile
  If chkR <> "" Then ' es wurde schon gespeichert
          frage = MsgBox("[Ja ]Paste content" & vbCrLf & vbCrLf & "[Nein ]Buffer new current values", vbYesNoCancel, "Speicher: Stromwerte vorhanden / Buffer: not empty (current/times)")
  End If
  If frage = "2" Then
       For i = 0 To 8
          mCurFild(i) = ""
       Next
      ActiveWindow.FreezePanes = False
      Worksheets("Formblatt").StromCopyBT.BackColor = RGB(220, 220, 220)
      Worksheets("Formblatt").StromCopyBT.ForeColor = RGB(0, 0, 0)
      Sheets("Formblatt").Cells(cnt, 7).Select 'PinNummer auswählen
      Application.StatusBar = "-------------------V O B E S Message ---------------- Freeze Panes @False"
      Exit Sub
  End If

  If frage = "7" Or chkR = "" Or chkR = cnt Then 'yes war die Antwort
             'aktive Zeile merken
             mCurFild(0) = ActiveCell.Row  ' Von Zeile
              ' Strom 1
             mCurFild(1) = Sheets("Formblatt").Cells(cnt, 78).Value 'I1 Zellenwerte holen
             Sheets("Formblatt").Cells(cnt, 78).Font.Color = RGB(0, 0, 255)
              ' Strom 2
             mCurFild(2) = Sheets("Formblatt").Cells(cnt, 85).Value 'I2 Zellenwerte holen
             Sheets("Formblatt").Cells(cnt, 85).Font.Color = RGB(0, 0, 255)
             ' Strom 3
             mCurFild(3) = Sheets("Formblatt").Cells(cnt, 92).Value 'I3 Zellenwerte holen
             Sheets("Formblatt").Cells(cnt, 92).Font.Color = RGB(0, 0, 255)
             ' Strom 4
             mCurFild(4) = Sheets("Formblatt").Cells(cnt, 99).Value 'I4 Zellenwerte holen
             Sheets("Formblatt").Cells(cnt, 99).Font.Color = RGB(0, 0, 255)
             ' Strom 5
             mCurFild(5) = Sheets("Formblatt").Cells(cnt, 106).Value 'I5 Zellenwerte holen
             ' Strom 6
             mCurFild(6) = Sheets("Formblatt").Cells(cnt, 113).Value 'I6 Zellenwerte holen
             Sheets("Formblatt").Cells(cnt, 113).Font.Color = RGB(0, 0, 255)
            'Zeit1
             mCurFild(7) = Sheets("Formblatt").Cells(cnt, 120).Value 't1 Zellenwerte holen
             Sheets("Formblatt").Cells(cnt, 120).Font.Color = RGB(0, 0, 255)
            'Zeit2
             mCurFild(8) = Sheets("Formblatt").Cells(cnt, 127).Value 't2 Zellenwerte holen
             Sheets("Formblatt").Cells(cnt, 127).Font.Color = RGB(0, 0, 255)
            With Worksheets("Formblatt").StromCopyBT
                       .ForeColor = RGB(0, 0, 0)
                       .BackColor = RGB(100, 255, 100)
            End With
            Application.StatusBar = "-------------------V O B E S Message ---------------- Current of Pin " & StckPlz & PinNr & " Copied  to > memory"
             Exit Sub
End If

If frage = "6" Or cnt <> chkR And chkR <> "" Then
             'aktive Zeile speichern
             vPinNr = Sheets("Formblatt").Cells(chkR, 7).Value 'PinNummer
             vStckPlz = Sheets("Formblatt").Cells(chkR, 2).Value ' Steckplatz
              ' Strom 1
             Sheets("Formblatt").Cells(cnt, 78).Value = mCurFild(1) 'I1 Zellenwerte speichern
             Sheets("Formblatt").Cells(cnt, 78).Font.Color = RGB(255, 0, 255)
              ' Strom 2
             Sheets("Formblatt").Cells(cnt, 85).Value = mCurFild(2) 'I2 Zellenwerte speichern
             Sheets("Formblatt").Cells(cnt, 85).Font.Color = RGB(255, 0, 255)
             ' Strom 3
             Sheets("Formblatt").Cells(cnt, 92).Value = mCurFild(3) 'I3 Zellenwerte speichern
             Sheets("Formblatt").Cells(cnt, 92).Font.Color = RGB(255, 0, 255)
             ' Strom 4
             Sheets("Formblatt").Cells(cnt, 99).Value = mCurFild(4) 'I4 Zellenwerte speichern
             Sheets("Formblatt").Cells(cnt, 99).Font.Color = RGB(255, 0, 255)
             ' Strom 5
             Sheets("Formblatt").Cells(cnt, 106).Value = mCurFild(5) 'I5 Zellenwerte speichern
             Sheets("Formblatt").Cells(cnt, 106).Font.Color = RGB(255, 0, 255)
             ' Strom 6
             Sheets("Formblatt").Cells(cnt, 113).Value = mCurFild(6) 'I6 Zellenwerte speichern
             Sheets("Formblatt").Cells(cnt, 113).Font.Color = RGB(255, 0, 255)
            'Zeit1
             Sheets("Formblatt").Cells(cnt, 120).Value = mCurFild(7) 't1 Zellenwerte speichern
             Sheets("Formblatt").Cells(cnt, 120).Font.Color = RGB(255, 0, 255)
            'Zeit2
             Sheets("Formblatt").Cells(cnt, 127).Value = mCurFild(8) 't2 Zellenwerte speichern
             Sheets("Formblatt").Cells(cnt, 127).Font.Color = RGB(255, 0, 255)
              With Worksheets("Formblatt").StromCopyBT
                       .ForeColor = RGB(255, 255, 255)
                       .BackColor = RGB(255, 255, 0)
               End With
             Application.StatusBar = " ---------------------------V O B E S Message --------------- Current of Pin " & vStckPlz & vPinNr & "  Pasted > " & StckPlz & PinNr
             'm = MsgBox("Current of Pin " & vStckPlz & vPinNr & "  Pasted > " & StckPlz & PinNr, vbInformation, "Stromwerte übernehmen / inherit values of current")
        End If
 
End Sub


