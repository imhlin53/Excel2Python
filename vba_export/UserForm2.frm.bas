Attribute VB_Name = "UserForm2"
Attribute VB_Base = "0{0A741943-3144-4B5E-98BF-67FA880187DC}{0F350BF8-F83C-480D-946E-AF8E1226C7F5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public LBFa, LBFb, LBFc, Funktion, Oertlichkeit, FunktZaehler, ActFctZaehler, FunktionenListe, FunktionenListe1, Spannungstyp, Newfct_state, tmp_FnctIdx
Public MaxLenTxtSp, MaxFctLenTxtSp
Public ActKlZaehler As Integer
Public go_thr, SO 'search option

Private Sub CB_Zoom_Click()
UserForm2.Zoom = UserForm2.Zoom - 5
End Sub

Private Sub Bild2_Click()
Bild2.BackStyle = fmBackStyleOpaque
Bild2.BackColor = red
End Sub


Private Sub An_Aus_CB_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
UserForm2.Label6.Visible = True
UserForm2.Label7.Visible = True
UserForm2.Label8.Visible = True
UserForm2.Label9.Visible = True
TextBox1.Visible = True
TextBox4.Visible = True
TextBox5.Visible = True
TextBox6.Visible = True
Label42.Visible = True
Label43.Visible = True
Label44.Visible = True
Label45.Visible = True
Label107.Visible = True
SigAbbrevCobo.Visible = True
SigCharCombo.Visible = True
IOX_ComBo.Visible = True
Voltage_CoBo.Visible = True
End Sub


Private Sub An_Aus_CB_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
UserForm2.Label6.Visible = False
UserForm2.Label7.Visible = False
UserForm2.Label8.Visible = False
UserForm2.Label9.Visible = False
TextBox1.Visible = False
TextBox4.Visible = False
TextBox5.Visible = False
TextBox6.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
Label107.Visible = False
SigAbbrevCobo.Visible = False
SigCharCombo.Visible = False
IOX_ComBo.Visible = False
Voltage_CoBo.Visible = False

End Sub

Private Sub CB_CLR1_Click()
'Alle Eingaben zurücksetzen ' Urban 19.03.2021
Cl_ListeBox.Text = "nCL"
CoBoFunktion.Text = "nFct"
CoBoFunktion_DropButtonClick
IdxText2.Value = ""
CoBoUtili.Text = "nVal"
   UserForm2.TextBox2.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox3.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox7.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox8.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox9.BackColor = RGB(255, 255, 245)
End Sub

Private Sub Cl_ListeBox_Change()
'Klemmen Daten füllen
'clamps(100, 6) Low;Low-Side Schalter;low side switch;M;DC;IN;Spalten = Me.Cl_ListeBox.ColumnCount
Zeile = Me.Cl_ListeBox.ListCount
Zindex = Me.Cl_ListeBox.ListIndex
If Zindex = -1 Then Exit Sub ' Abbruch durch Auswahl einer leeren Zeile
Klemme = Me.Cl_ListeBox.Column(0, Zindex) '          CL, Klemme, 30/31/75/plus
If Klemme = "nCL" Or Funktion <> "" And Funktion <> "nFct" Then
    'Signalart von der Funktion holen' die Klemme wird zurückgesetzt
    Me.CoBoFunktion.Text = Funktion
End If
'If Klemme <> "nCL" And Funktion = "" Or Funktion = "nFct" Then 'Urban 19.03.2021
If Klemme <> "nCL" And Funktion <> "" Or Funktion <> "nFct" Then
    
    PotentialTyp = Me.Cl_ListeBox.Column(3, Zindex) '    CL_SignalAbbreviation, V/M/S/D
End If
If Funktion <> "" Or Funktion <> "nFct" Then 'Urban 19.03.2021
    If Klemme = "nCL" Or Klemme = "" Then
       CoBoFunktion_DropButtonClick 'Signalart von der Funktion holen' die Klemme wird zurückgesetzt
       'PotentialTyp = Me.CoBoFunktion.Column(4, LBFa_Zindex) 'Signalabbreviation S,V,M
    End If
End If

Beschreibung_DE = Me.Cl_ListeBox.Column(1, Zindex) ' CL_DescriptiveText
Beschreibung_EN = Me.Cl_ListeBox.Column(2, Zindex) ' CL_DescriptiveText
'PotentialTyp = Me.Cl_ListeBox.Column(3, Zindex) '    CL_SignalAbbreviation, V/M/S/D
SignalArt_CL = Me.Cl_ListeBox.Column(4, Zindex) '    CL_SignalAbbreviation, AC, DC, HV, HF, NF
'todo
Signalrichtung = Me.Cl_ListeBox.Column(5, Zindex) '  CL_Direction, .Font.Color = RGB(0, 0, 0)
Spannung = Me.Cl_ListeBox.Column(6, Zindex) '        Cl_Voltage
Me.TextBox1.Value = Klemme

Me.TextBox2.Value = Beschreibung_DE
Me.TextBox3.Value = Beschreibung_EN
Me.TextBox4.Value = PotentialTyp
TextBox5.Value = SignalArt_CL ' Signal/Spannungs Art
Spannungstyp = SignalArt_CL 'neu
Me.Voltage_CoBo.Value = Spannung ' Neu
Me.TextBox6.Value = Signalrichtung ' Richtung
'Me.TextBox6.BackColor = RGB(256, 100, 100)
ActKlZaehler = ActKlZaehler + 1 ' Klemme war aktiviert
End Sub

Private Sub clampCoBo_Change()
UserForm2.ToggleButton1.BackColor = RGB(255, 255, 200)
UserForm2.TextBox2.BackColor = RGB(255, 255, 200)
UserForm2.TextBox3.BackColor = RGB(255, 255, 200)
UserForm2.TextBox8.BackColor = RGB(255, 255, 200)
UserForm2.OptionButton1.Value = False ' temp Funktion  auf CL_Info zurücksetzen
Newfct_state = False
UserForm2.Label32.Visible = False ' Textmarker unsichtbar
LBFc_Zindex = Me.clampCoBo.ListIndex
If LBFc_Zindex <= 0 Then
  LBFc_Zindex = 0
End If
Funktion = Me.clampCoBo.Column(0, LBFc_Zindex) 'UGU 07.10.2018
Klemme = Me.clampCoBo.Column(2, LBFc_Zindex) 'war 0
 'Klemme untersuchen
 KlAnfVal = Mid(Klemme, 1, 1)
 ActKlZaehler = ActKlZaehler + 1
If Klemme <> "nCL" Then
  ActKlZaehler = 0 'Klemme hat einen Wert"
    'Funktion untersuchen da eine leere Funktion im Speicher aus Klemme# und Clamp zusammengebaut wird 20.02.2020
      LenFunkt = Len(Funktion)
      ChkText = Left(Funktion, 7)
      If ChkText = "Klemme#" Then
         Funktion = "nFct"
      End If
End If
If Klemme = "CL" Or KlAnfVal = "J" Then Klemme = "n.c."
ActFctZaehler = ActFctZaehler + 1              'UGU 07.10.2018
Me.Cl_ListeBox.Value = Klemme
Me.TextBox1.Value = Klemme
'Funktion = Me.clampCoBo.Column(0, LBFc_Zindex) 'war 1, die Funktion muss früher eingelesen werden
Me.CoBoFunktion.Value = Funktion
UserForm2.TextBox7.BackColor = RGB(255, 255, 200)
UserForm2.TextBox9.BackColor = RGB(255, 255, 200)
Beschreibung_EN = Me.clampCoBo.Column(1, LBFc_Zindex) 'war 2
Beschreibung_DE = Me.clampCoBo.Column(9, LBFc_Zindex)
Spannung = Me.clampCoBo.Column(5, LBFc_Zindex)
SignalArt = Me.clampCoBo.Column(6, LBFc_Zindex)
PotentialTyp = Me.clampCoBo.Column(7, LBFc_Zindex)
Signalrichtung = Me.clampCoBo.Column(4, LBFc_Zindex) 'direction   CRX, IN, OUT
Oertlichkeit = Me.clampCoBo.Column(3, LBFc_Zindex) 'Utilization fL, rR, 3rD
'Signalrichtung = clampSpeicher(LBFc_Zindex + 1, 10)
'Oertlichkeit = clampSpeicher(LBFc_Zindex + 1, 4)
UserForm2.CoBoUtili.Value = Oertlichkeit
UserForm2.TextBox4.Value = PotentialTyp
UserForm2.TextBox5.Value = SignalArt
UserForm2.TextBox6.Value = Signalrichtung
UserForm2.TextBox7.Value = Beschreibung_EN ' "Description"
UserForm2.TextBox9.Value = Beschreibung_DE ' "Beschreibung"
UserForm2.Cl_ListeBox.Value = Klemme 'neu
End Sub

Private Sub CoBoFunktion_DropButtonClick()
'Funktions Daten erstellen
Newfct_state = False
UserForm2.OptionButton1.Value = False
UserForm2.CoBoFunktion.MatchEntry = fmMatchEntryComplete
'tmp_FnctIdx = 0 ' Urban 08.07.2021

UserForm2.Label32.Visible = False ' Textmarker unsichtbar
'If tmp_FnctIdx >= 1 Then
 '  UserForm2.CoBoFunktion.BackColor = RGB(255, 150, 150) 'temporäre Funktion anzeigen
  ' Else: UserForm2.CoBoFunktion.BackColor = RGB(255, 255, 255)
'End If
UserForm2.TextBox7.BackColor = RGB(255, 255, 255)
UserForm2.TextBox9.BackColor = RGB(255, 255, 255)
LBFa_Zindex = Me.CoBoFunktion.ListIndex  'Me = UserForm2.CoBoFunktion.ListIndex
'If LBFa_Zindex <= 0 Then LBFa_Zindex = 0
If LBFa_Zindex <= 0 Then Exit Sub
LBFa = Me.CoBoFunktion.Column(1, LBFa_Zindex)
DE_Comment = Me.CoBoFunktion.Column(2, LBFa_Zindex)
'If tmp_FnctIdx = 0 Then                 'temporäre Funktion bleibt erhalten 08.08.2021 klären!
   Funktion = Me.CoBoFunktion.Column(0, LBFa_Zindex)
'End If
SignalArt = Me.CoBoFunktion.Column(3, LBFa_Zindex)
'Urban 22.11.2018
Chk_Klemme = Me.Cl_ListeBox.Value
If Chk_Klemme = "" Or Chk_Klemme = "nCL" Then ' die Klemme gibt die Signalabbreviation vor
  PotentialTyp = Me.CoBoFunktion.Column(4, LBFa_Zindex) 'Signalabbreviation S,V,M
End If
Signalrichtung = Me.CoBoFunktion.Column(5, LBFa_Zindex) 'direction  CRX, IN, OUT
UserForm2.CoBoFunktion.BackColor = RGB(255, 255, 255)
Me.TextBox4.Value = PotentialTyp
Me.TextBox5.Value = SignalArt      'SigCharacter mixed, CAN, ND
Me.TextBox6.Value = Signalrichtung ' Richtung
Me.TextBox7.Value = LBFa           ' Description
Me.TextBox9.Value = DE_Comment     ' Beschreibung
ActFctZaehler = ActFctZaehler + 1
End Sub
Private Sub CoBoFunktion_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'DA 09.07.2019 Behandlung Sonderzeichen in Textboxen
Select Case KeyAscii
      Case 48 To 57   ' 0-9 zulassen
      Case 65 To 90   ' A-Z zulassen
      Case 97 To 122 ' a-z zulassen
      Case 47              ' / zulassen UGU 11.06.2020
      Case 43              ' + zulassen
      Case 44              ' , zulassen
      Case 45              ' - zulassen
      Case 95              ' _ zulassen
      Case Else           ' alles andere ablehnen
         KeyAscii = 0
   End Select
End Sub
Private Sub CoBoFunktion_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'DA 09.07.2019 Behandlung Sonderzeichen in Textboxen
If (Shift = 2 And KeyCode = 86) Then
    KeyCode = 0 ' CTRL+V ablehnen
ElseIf (Shift = 1 And KeyCode = 45) Then
    KeyCode = 0 ' SHIFT+INSERT ablehnen
ElseIf (Shift = 2 And KeyCode = 88) Then
    KeyCode = 0 ' CTRL+X ablehnen
ElseIf (Shift = 1 And KeyCode = 46) Then
    KeyCode = 0 ' SHIFT+ENTF ablehnen
End If
End Sub

Private Sub CoBoUtili_Change()
' Örtlichkeit bestimmen
LBFb_Zindex = Me.CoBoUtili.ListIndex
LBFb_Zcount = Me.CoBoUtili.ListCount
If LBFb_Zindex = -1 Then
    For suchIndex = 0 To LBFb_Zcount
      If suchIndex = LBFb_Zcount Then
        Me.TextBox8.Value = ""
        Exit Sub ' kein Index ermittelbar Fehler in der Datenbasis
      End If
      If Oertlichkeit = Me.CoBoUtili.Column(0, suchIndex) Then
          LBFb_Zindex = suchIndex
          suchIndex = LBFb_Zcount ' Schleife abbrechen weil Zaehler ermittelt wurde
       End If
    Next
End If
LBFb = Me.CoBoUtili.Column(1, LBFb_Zindex)
Oertlichkeit = Me.CoBoUtili.Column(0, LBFb_Zindex)
Me.TextBox8.Value = LBFb
End Sub

Private Sub CommandButton1_Click()
If Newfct_state = True Then GoTo TmpFunktion
Me.Hide
  Workbooks(VFbl).Activate
  Worksheets("Formblatt").Select
  origC = Range("PinsUpper").Column + 14 'ActiveCell.Column
  OrigRu = Range("PinsUpper").Row + 1 '   ActiveCell.Row
  OrigRl = Range("PinsLower").Row - 1 '   ActiveCell.Row

'Definitionen CL_Info 29.03.2020
 CL_Stat = False    ' eine Klemme nicht vorhanden
 Fc_stat = False    ' eine Funktion nicht vorhanden
 Ut_stat = False    ' eine Örtlichkeit nicht vorhanden
 TmpKLinfo = ""
 TmpLAH = ""
 TmpKLz = ""
 AnBAnfg = ""    ' BN Anbindungsanforderung
 ASILevel = "no" ' ASIL Level Status
 aktive_Reihe = ActiveCell.Row
 If aktive_Reihe < OrigRu Or aktive_Reihe > OrigRl Then ' Pinbereich prüfen
   MsgBox ("Sorry, leider haben Sie keinen KomponentenPin ausgewählt.")
   Me.Hide
   'Windows(1).WindowState = xlMaximized
   Set ACT = Application.InputBox(Prompt:="Bitte wählen Sie jetzt einen PIN von Steckplatz A oder Steckplatz B aus", Default:="$D$3", Left:=300, Top:=20, Type:=8)
    ACAdd = ACT.AddressLocal
    Range(ACAdd).Select
    aktive_Reihe = ActiveCell.Row
   If aktive_Reihe < OrigRu Then ' Zeile 2 oder kleiner ist nicht sinnvoll!
     Me.Hide
     MsgBox ("Sorry, Sie haben leider wiederholt keinen KomponentenPin ausgewählt, der Abbruch Funktion erfolgt jetzt.")
     Exit Sub
   End If
   Me.Show
 End If
 AnBAnfg = Cells(aktive_Reihe, 34).Value 'BN Anbindungsanforderung 29.03.2020
   Len_AnBAnfg = Len(AnBAnfg) 'Länge des Kommentar feststellen
   TxtCut = Left(AnBAnfg, 4)
   TxtCut1 = Left(AnBAnfg, 3)
   TxtCut2 = Left(AnBAnfg, 2)
   TxtCut3 = Left(AnBAnfg, 6)
    AnB_pos = InStr(1, AnBAnfg, ";") 'Suche nach einem Trennzeichen ";" im String der Anbindungsanforderungen
      'ASIL Level ermitteln
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
 
'7, Potential  9, Direction   10, Voltage 12, SignalArt
'Standard Defaultwerte setzen
If PotentialTyp = "" Then
   PotentialTyp = "VS"
End If

Signalrichtung_pre = Me.TextBox6.Value ' Signalrichtung bereits definiert feststellen
If Signalrichtung = "" Then
   Signalrichtung = "IO"
   ElseIf Signalrichtung <> Signalrichtung_pre Then Signalrichtung = Signalrichtung_pre
End If
If Spannung = "" Then
   Spannung = "12"
End If
If Spannungstyp = "" Then
    Spannungstyp = "DC"  ' 16.08.2019 Urban, bei tmp Funktionen notwendig
End If
SignalArt_pre = Me.TextBox5.Value ' Urban 19.03.2021
If SignalArt = "" Or SignalArt = "nVal" Then
   SignalArt = "multi"
   ElseIf SignalArt <> SignalArt_pre Then SignalArt = SignalArt_pre
End If
'Standard Defaultwerte setzen
ActiveWorkbook.Sheets("Formblatt").Activate
If Klemme = "nCL" Or Klemme = "n.c." Then Klemme = ""
If Klemme <> "nCL" And Klemme <> "" Then ' Leere Klemme abfangen
 'die KLemme erhält ein # nach ihrem Wert
   CL_secure = Right(Klemme, 1) ' q oder s
   If CL_secure = "s" Or CL_secure = "q" Then CL_Stat = True     ' eine Klemme ist vorhanden
   ZSB_Klemme = Klemme & "#" 'Cells(aktive_Reihe, 4).Value
  
 Else
    ZSB_Klemme = ""
End If

'FunktZaehler
If ActFctZaehler <> "" And Funktion <> "Funktion" And Funktion <> "nFct" Then
    ZSB_Funktion = Funktion 'Funktion ' Cells(aktive_Reihe, 5).Value
    Fc_stat = True
 Else
    ZSB_Funktion = ""
End If
If LBFb = "undefined" Then 'Oertlichkeit
    Oertlichkeit = "" ' Oertlichkeit
End If
If Oertlichkeit = "Örtlichkeit" Or Oertlichkeit = "nVal" Then
   Oertlichkeit = "" ' Anfangswert der Liste entfernen
End If
If Oertlichkeit <> "" Then
    ZsbOertlichkeit = "." & Oertlichkeit
    Ut_stat = True
End If
If ZSB_Klemme = "" And ZSB_Funktion = "" And ZsbOertlichkeit <> "" Then
     MakeNoSense = MsgBox("Eine Örtlichkeit reicht als PinBeschreibung nicht aus!" & vbCrLf & "Only a utilization isn´t enough, an uncomplete PinDefinition makes no sense", vbCritical, "Error: PinDefinition")
     Exit Sub
End If
If ZSB_Klemme = "" And ZSB_Funktion = "" And ZsbOertlichkeit = "" Then
     MakeNoSense = MsgBox("Eine leere PinBeschreibung reicht nicht aus!" & vbCrLf & "More details please, isn´t enough input", vbCritical, "Error: PinDefinition")
     Exit Sub
End If
' Urban 20.07.20 * kennzeichnet es ist eine Bordnetzanbindung/ LAH Info vorhanden
If BNAStat = True Then
'todo 29.03.2020
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
   
   'Urban 28.03.2020
     If ASILevel = "yes" And CL_Stat = True And Fc_stat = True And Ut_stat = True Then
        ZSB_KLemmeninfo = "*" & Klemme & "#" & Funktion & "." & Oertlichkeit & "_" & ASIL_Val
     End If
     If ASILevel = "yes" And CL_Stat = True And Fc_stat = True And Ut_stat = False Then
        ZSB_KLemmeninfo = "*" & Klemme & "#" & Funktion & "_" & ASIL_Val
     End If
     If ASILevel = "yes" And CL_Stat = True And Fc_stat = False Then
        ZSB_KLemmeninfo = "*" & Klemme & "_" & ASIL_Val
     End If
     If ASILevel = "yes" And CL_Stat = True And Fc_stat = False And Ut_stat = True Then
        ZSB_KLemmeninfo = "*" & Klemme & "." & Oertlichkeit & "_" & ASIL_Val
     End If
     If ASILevel = "yes" And CL_Stat = False And Klemme <> "" And Fc_stat = True And Ut_stat = True Then
        ZSB_KLemmeninfo = "*" & ZSB_Klemme & Funktion & "." & Oertlichkeit
     End If
     If ASILevel = "yes" And CL_Stat = False And Klemme = "" And Fc_stat = True And Ut_stat = True Then
        ZSB_KLemmeninfo = "*" & ZSB_Klemme & Funktion & "." & Oertlichkeit
     End If
     If ASILevel = "yes" And ZSB_Klemme = "" And CL_Stat = False And Fc_stat = True And Ut_stat = False Then
        ZSB_KLemmeninfo = "*" & Funktion
        MsgBox ("Sichere Klemme notwendig?! Sie haben einen ASIL Level definiert? ASIL Level value defined!")
     End If
     
    If ASILevel = "yes" And Klemme <> "" And CL_Stat = False And Fc_stat = True And Ut_stat = False Then
        ZSB_KLemmeninfo = "*" & ZSB_Klemme & Funktion
        MsgBox ("Sichere Klemme notwendig?! Sie haben einen ASIL Level definiert? ASIL Level value defined!")
     End If
     
     If ASILevel = "yes" And CL_Stat = False And Fc_stat = False And Ut_stat = True Then
        ZSB_KLemmeninfo = "*" & Klemme & "." & Oertlichkeit
     End If
     If ASILevel = "yes" And CL_Stat = False And Fc_stat = False And Ut_stat = False Then
        ZSB_KLemmeninfo = "*" & Klemme
     End If
      
End If 'BNAStat = true
     ' Standard
     If BNAStat = False Then
        ZSB_KLemmeninfo = ZSB_Klemme & ZSB_Funktion & ZsbOertlichkeit
     End If
     If BNAStat = False And ASILevel = "no" And CL_Stat = True Then 'Or ASILevel = "BNA"
        ZSB_KLemmeninfo = ZSB_Klemme & ZSB_Funktion & ZsbOertlichkeit
        MsgBox ("Sichere Klemme " & Klemme & " definiert und keinen ASIL Level definiert!" & vbCrLf & "Secured Clamp and ASIL Level value undefined!")
     End If
     
     If BNAStat = True And ASILevel = "BNA" And ZSB_Klemme = "" And CL_Stat = False And Fc_stat = True And Ut_stat = True Then
          ZSB_KLemmeninfo = "*" & ZSB_Funktion & ZsbOertlichkeit
     End If
     
     If BNAStat = True And ASILevel = "BNA" And ZSB_Klemme <> "" And CL_Stat = False And Fc_stat = True And Ut_stat = True Then
          ZSB_KLemmeninfo = "*" & ZSB_Klemme & ZSB_Funktion & ZsbOertlichkeit
     End If
     
     If BNAStat = True And ASILevel = "BNA" And ZSB_Klemme <> "" And CL_Stat = False And Fc_stat = True And Ut_stat = False Then
          ZSB_KLemmeninfo = "*" & ZSB_Klemme & ZSB_Funktion
     End If
     If BNAStat = True And ASILevel = "BNA" And ZSB_Klemme <> "" And CL_Stat = False And Fc_stat = False And Ut_stat = True Then
          ZSB_KLemmeninfo = "*" & ZSB_Klemme & ZsbOertlichkeit
     End If
     If BNAStat = True And ASILevel = "BNA" And ZSB_Klemme <> "" And CL_Stat = False And Fc_stat = False And Ut_stat = False Then
          ZSB_KLemmeninfo = "*" & ZSB_Klemme
     End If
     
     If BNAStat = True And ASILevel = "BNA" And ZSB_Klemme = "" And CL_Stat = False And Fc_stat = True And Ut_stat = False Then
          ZSB_KLemmeninfo = "*" & ZSB_Funktion
     End If
     
     If BNAStat = True And ASILevel = "no" And CL_Stat = False Then
        ZSB_KLemmeninfo = "*" & ZSB_Klemme & ZSB_Funktion & ZsbOertlichkeit
     End If
     If BNAStat = True And CL_Stat = True And ASILevel = "no" Or BNAStat = True And CL_Stat = True And ASILevel = "BNA" Then
        ZSB_KLemmeninfo = "*" & ZSB_Klemme & ZSB_Funktion & ZsbOertlichkeit
        MsgBox ("Sichere Klemme " & Klemme & " definiert und keinen ASIL Level definiert!" & vbCrLf & "Secured Clamp and ASIL Level value undefined!")
     End If
     'Standard Ende

DE_Comment = Me.TextBox9.Value
If ZSB_Funktion = "" Then
   DE_Comment = Me.TextBox3.Value
End If
If frage = "yes" Then
  Cells(aktive_Reihe, 15).Value = ZSB_KLemmeninfo
  Cells(aktive_Reihe, 136).Value = Klemme
  Cells(aktive_Reihe, 143).Value = ZSB_Funktion
  Cells(aktive_Reihe, 152).Value = Oertlichkeit
  Cells(aktive_Reihe, 158).Value = Signalrichtung
  Cells(aktive_Reihe, 163).Value = Spannung
  Cells(aktive_Reihe, 167).Value = Spannungstyp
  Cells(aktive_Reihe, 171).Value = PotentialTyp
  Cells(aktive_Reihe, 176).Value = SignalArt
  Cells(aktive_Reihe, 183).Value = DE_Comment ' Freitxt hier eintragen
  ZSB_KLAttribute = Klemme & ZSB_Funktion & Oertlichkeit & PotentialTyp & Signalrichtung & Spannung & SignalArt
  ManipulationZiffer (ZSB_KLAttribute)
  'temp. je BGER ID des Steckplatzes erzeugen
    StckPlz = Cells(aktive_Reihe, 2).Value
    PNumber = Range("Teilenummer").Value
    ZSBSTPN = StckPlz & PNumber
      For i = 1 To Len(ZSBSTPN)
        ZWwert = Mid(ZSBSTPN, i, 1)
        ZSBziffer = ZSBziffer + Asc(ZWwert)
      Next
    tempId = "tmp_" & ZSBziffer
    Cells(aktive_Reihe, 226).Value = tempId  ' temporäre BGER Id vergeben"
    Cells(aktive_Reihe, 226).Interior.Color = RGB(255, 100, 100)
  Cells(aktive_Reihe, 33).Value = ZSB_KLziffer 'ZSB_KLAttribute ' Kontrollausgabe
  Cells(aktive_Reihe, 15).Interior.Color = RGB(204, 255, 255)
End If ' Frage ist yes
BNAStat = False
'Urban 20.03.2023
Tabelle1.CommandButton2.Caption = "CheckMe"
Tabelle1.CommandButton2.ForeColor = blue
Range("DataChecked").Value = "ungeprueft"
Range("State").Value = ""
Unload Me ' UserForm2
TmpFunktion:
If Newfct_state = True Then NewFctLabel
End Sub

Private Sub CommandButton2_Click()
Cl_ListeBox.Text = "nCL"
End Sub


Private Sub CommandButton1_Exit(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub HelpCmBt_Click()
UserForm3.Show
'MsgBox ("Welche Wirkung hat der Pin nach außen?" & vbCrLf & "Ist der Pin erine versorgende Klemme?" & vbCrLf & "(PLUS/ 30/ 15) oder eine Masse(GND/ 31)" & vbCrLf & "--")
End Sub

Private Sub IdxText_Change()
If IdxText.Value = "" Then IdxText.Value = -1
UserForm2.TextBox2.BackColor = RGB(255, 255, 200)
UserForm2.TextBox3.BackColor = RGB(255, 255, 200)
UserForm2.TextBox7.BackColor = RGB(255, 255, 200)
UserForm2.TextBox8.BackColor = RGB(255, 255, 200)
UserForm2.TextBox9.BackColor = RGB(255, 255, 200)
testlen = Len(IdxText.Value)
For q = 1 To testlen
 TxtSnip = Mid(IdxText.Value, q, 1)
 If Asc(TxtSnip) < 48 Or Asc(TxtSnip) > 57 Then
    IdxText.Value = ""
    Exit For
 End If
Next
If UserForm2.clampCoBo.ListCount > IdxText.Value Then UserForm2.clampCoBo.ListIndex = IdxText.Value
End Sub


Private Sub IdxText2_Change()
'If IdxText2.Value = "" Then IdxText2.Value = 0
testlen = Len(IdxText2.Value)
For q = 1 To testlen
 TxtSnip = Mid(IdxText2.Value, q, 1)
 If Asc(TxtSnip) < 48 Or Asc(TxtSnip) > 57 Then
    IdxText.Value = ""
    Exit For
 End If
Next

If UserForm2.CoBoFunktion.ListCount > IdxText2.Value Then
      UserForm2.CoBoFunktion.ListIndex = IdxText2.Value
      CoBoFunktion_DropButtonClick
End If

End Sub

Private Sub Image27_Click()
Image27.BackStyle = fmBackStyleTransparent
End Sub


Private Sub Label108_Click()

End Sub

Private Sub RepListB_Click()
'Urban 30.03.2023 Selektion einer Zeile im Report für eine Index suche aus
If SO = "FKT" And ToggleButton1.Caption = "Edit NewDef" Then
  Call Label11_Click 'Leerwert setzen
  Call Label17_Click
  IdxText2 = UserForm2.RepListB.Value
End If
If SO = "CL_Info" Then
   Call Label10_Click   'Leerwert setzen
   Call Label11_Click
   Call Label17_Click
  IdxText = UserForm2.RepListB.Value
End If
End Sub

Private Sub SpinButton1_SpinDown()
UserForm2.Zoom = UserForm2.Zoom - 5
End Sub

Private Sub SpinButton1_SpinUp()
UserForm2.Zoom = UserForm2.Zoom + 5
End Sub


Private Sub TextBox7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'DA 09.07.2019 Behandlung Sonderzeichen in Textboxen
Select Case KeyAscii
      Case 48 To 57   ' 0-9 zulassen
      Case 65 To 90   ' A-Z zulassen
      Case 97 To 122 ' a-z zulassen
      Case 43              ' + zulassen
      Case 44              ' , zulassen
      Case 45              ' - zulassen
      Case 95              ' _ zulassen
      Case Else           ' alles andere ablehnen
         KeyAscii = 0
   End Select
End Sub

Private Sub TextBox8_Change()

End Sub

Private Sub TextBox9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'DA 09.07.2019 Behandlung Sonderzeichen in Textboxen
Select Case KeyAscii
      Case 48 To 57   ' 0-9 zulassen
      Case 65 To 90   ' A-Z zulassen
      Case 97 To 122 ' a-z zulassen
      Case 43              ' + zulassen
      Case 44              ' , zulassen
      Case 45              ' - zulassen
      Case 95              ' _ zulassen
      Case Else           ' alles andere ablehnen
         KeyAscii = 0
   End Select
End Sub

Private Sub Voltage_CoBo_Change()
'die ComboBox1 gibt bei Auswahl der Fzg NennSpannung die veränderten Werte zurück
'Voltage;EN Description; DE Description;VoltageType
ZLindex = Me.Voltage_CoBo.ListIndex
If ZLindex <= 0 Then ZLindex = 0
Spannung = Me.Voltage_CoBo.Column(0, ZLindex)
Spannungstyp = Voltage_CoBo.Column(3, ZLindex) ' der Spannungstyp wird gemerkt
UserForm2.SigCharCombo.Value = Spannungstyp
Funktion = Me.CoBoFunktion.Value 'Urban23.11.18
If Funktion = "" Or Funktion = "nFct" Then
   'eine Funktion belegt die Signalart und nur wenn keine
   'Funktion festgelegt ist dann setzt die Spanunng den Wert

   SignalArt = Me.Voltage_CoBo.Column(3, ZLindex)
   TextBox5.Value = SignalArt
End If
End Sub


Private Sub IOX_ComBo_Change()
 'Signalrichtung_pre = Me.TextBox6.Value
 'datenbasis einlesen UGU 26.06.2018
IO_Zindex = Me.IOX_ComBo.ListIndex
If IO_Zindex <= 0 Then IO_Zindex = 0
 Signalrichtung = IOX_ComBo.Column(0, IO_Zindex)
 'Me.IOX_ComBo.Text = Signalrichtung_pre
 Me.TextBox6.Value = Signalrichtung
End Sub



Private Sub Label10_Click()
CoBoFunktion.Text = "nFct"
CoBoFunktion_DropButtonClick
   UserForm2.TextBox2.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox3.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox7.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox8.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox9.BackColor = RGB(255, 255, 245)
End Sub

Private Sub Label11_Click()
CoBoUtili.Text = "nVal"
   UserForm2.TextBox2.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox3.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox7.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox8.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox9.BackColor = RGB(255, 255, 245)
End Sub

Private Sub Label17_Click()
Cl_ListeBox.Text = "nCL"
   UserForm2.TextBox2.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox3.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox7.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox8.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox9.BackColor = RGB(255, 255, 245)
End Sub

Private Sub Label22_Click()
UserForm2.RepListB.BackColor = RGB(255, 255, 180)
SO = "CL_Info"
Recherche (SO) ' Funktionsaufruf Suche nach einer Funktion in clampSpeicher
'SuchBOXausgabe.Text = FunktionenListe
End Sub

Private Sub Label24_Click()
' search for short word in Functions
UserForm2.RepListB.BackColor = RGB(180, 180, 255)
SO = "FKT"
Recherche (SO) ' Funktionsaufruf Suche nach einer Funktion in clampSpeicher
'SuchBOXausgabe.Text = FunktionenListe

End Sub

Private Sub NewFctLabel()
'UGU, Werte einer beantragten temporären Funktion festhalten
  Me.Hide
  Workbooks(VFbl).Activate
  Worksheets("Formblatt").Select
  ActRow = ActiveCell.Row
'28.06.19, Prüfung auf Länge und nicht erlaubten Sonderzeichen
    newFCT = CoBoFunktion.Value                                    'Funktion
       newFCT = ChknCorValue(newFCT, "F")                    'korrigierte Funktion
    newEN_Cmnt = UserForm2.TextBox7.Value               'Description"
       newEN_Cmnt = ChknCorValue(newEN_Cmnt, "C")  'korrigierter englische Description
    newDE_Cmnt = UserForm2.TextBox9.Value               'Beschreibung"
       newDE_Cmnt = ChknCorValue(newDE_Cmnt, "C")  'korrigierter deutsche Beschreibung
    If newEN_Cmnt = "must have a description" Or newDE_Cmnt = "Beschreibung ist erforderlich" Or newEN_Cmnt = "" Or newDE_Cmnt = "" Then
       MsgBox ("Leider haben Sie noch keine Beschreibung eingetragen, bitte nachholen!" & vbCrLf _
       & "I´m sorry, but please fill in german und english description!")
       Me.Show
       Exit Sub
    End If
    tmpFnct = "[TmpFnct " & newFCT & "#" & newEN_Cmnt & "*" & newDE_Cmnt & "]"  ' Urban 27.06.19 Leerzeichen entfernt
     Cells(ActRow, 183).Value = tmpFnct ' aktive Zeile leeren
     Cells(ActRow, 183).Font.Color = RGB(255, 100, 100)
'7, Potential  9, Direction   10, Voltage 12, SignalArt
'Standard Defaultwerte setzen
If PotentialTyp = "" Then
   PotentialTyp = "VS"
End If

Signalrichtung_pre = Me.TextBox6.Value ' Signalrichtung bereits definiert feststellen
If Signalrichtung = "" Then
   Signalrichtung = "IO"
   ElseIf Signalrichtung <> Signalrichtung_pre Then Signalrichtung = Signalrichtung_pre
End If
If Spannung = "" Then
   Spannung = "12"
End If
If Spannungstyp = "" Then
    Spannungstyp = "DC"  ' 16.08.2019 Urban, bei tmp Funktionen notwendig
End If
If SignalArt = "" Or SignalArt = "nVal" Then
   SignalArt = "multi"
End If
If Klemme <> "nCL" And Klemme <> "" Then ' Leere Klemme abfangen
 'die KLemme erhält ein # nach ihrem Wert
   ZSB_Klemme = Klemme '& "#" 'Cells(aktive_Reihe, 4).Value
 Else
    ZSB_Klemme = ""
End If
If Klemme = "nCL" Then Klemme = ""
If LBFb = "undefined" Then 'Oertlichkeit
    Oertlichkeit = "" ' Oertlichkeit
End If
If Oertlichkeit = "Örtlichkeit" Or Oertlichkeit = "nVal" Then
   Oertlichkeit = "" ' Anfangswert der Liste entfernen
End If
If Oertlichkeit <> "" Then
    ZsbOertlichkeit = "." & Oertlichkeit
    
End If
ZSB_KLemmeninfo = ZSB_Klemme & newFCT & ZsbOertlichkeit

  Cells(ActRow, 15).Value = ZSB_KLemmeninfo
  Cells(ActRow, 136).Value = Klemme
  Cells(ActRow, 143).Value = newFCT
  Cells(ActRow, 152).Value = Oertlichkeit
  Cells(ActRow, 158).Value = Signalrichtung
  Cells(ActRow, 163).Value = Spannung
  Cells(ActRow, 167).Value = Spannungstyp
  Cells(ActRow, 171).Value = PotentialTyp
  Cells(ActRow, 176).Value = SignalArt
  ZSB_KLAttribute = Klemme & newFCT & Oertlichkeit & PotentialTyp & Signalrichtung & Spannung & SignalArt
  ManipulationZiffer (ZSB_KLAttribute)
  StckPlz = Cells(ActRow, 2).Value
  PNumber = Range("Teilenummer").Value
  ZSBSTPN = StckPlz & PNumber
     For i = 1 To Len(ZSBSTPN)
        ZWwert = Mid(ZSBSTPN, i, 1)
        ZSBziffer = ZSBziffer + Asc(ZWwert)
     Next
  tempId = "tmp_" & ZSBziffer
  Cells(ActRow, 226).Value = tempId  ' temporäre BGER Id vergeben"
  Cells(ActRow, 226).Interior.Color = RGB(255, 100, 100)
  Cells(ActRow, 33).Value = ZSB_KLziffer 'ZSB_KLAttribute ' Kontrollausgabe
  Cells(ActRow, 15).Interior.Color = RGB(204, 255, 255)

Unload Me ' UserForm2
End Sub



Private Sub Label3_Click()
Call CommandButton1_Click
End Sub

Private Sub OptionButton1_Click()
'Prüfung
Newfct_state = True
    tmp_Fnct_Cmnt = Cells(ActiveCell.Row, 183).Value 'globale Var
    Ln_tmpFnctCmnt = Len(tmp_Fnct_Cmnt)              'Länge des Kommentar feststellen
    tmp_FnctIdx = InStr(1, tmp_Fnct_Cmnt, "TmpFnct") 'Suche nach temporärer Funktion im Kommentar
   
    UserForm2.Label32.Visible = True
    
    UserForm2.CoBoFunktion.MatchEntry = fmMatchEntryNone
    UserForm2.Cl_ListeBox.Value = "nCL"
    If tmp_FnctIdx = "" Then
       UserForm2.CoBoFunktion.Value = "nFct"
    End If
    UserForm2.CoBoUtili.Value = "nVal"
    If tmp_FnctIdx < 1 Then ' temp Definition ist nicht vorhanden
        UserForm2.SigCharCombo.Value = "mixed"
        UserForm2.IOX_ComBo.Value = "BI"
        UserForm2.Voltage_CoBo.Value = "12"
    End If
    UserForm2.SigAbbrevCobo.Value = "S"
    UserForm2.TextBox7.BackColor = RGB(255, 150, 150)
    'Urban 07.07.2021 beschreibende Texte erhalten
    If UserForm2.TextBox7.Value = "" And tmp_FnctIdx >= 1 Then
        UserForm2.TextBox7.Value = "ENG: Must have a description"
    End If
    UserForm2.TextBox8.Value = ""
    UserForm2.TextBox9.BackColor = RGB(255, 150, 150)
    If UserForm2.TextBox9.Value = "" And tmp_FnctIdx >= 1 Then
        UserForm2.TextBox9.Value = "DEU: Beschreibung ist erforderlich"
    End If
    UserForm2.CoBoFunktion.BackColor = RGB(255, 150, 150)

End Sub

Private Sub SigAbbrevCobo_Change()
'SigAbbrev(10, 1)#CL_SignalAbbreviation###
SigAbbrevCobo.ZOrder (fmBottom)
SigAbbrev_Zindex = Me.SigAbbrevCobo.ListIndex
If SigAbbrev_Zindex <= 0 Then SigAbbrev_Zindex = 0
  PotentialTyp = SigAbbrevCobo.Column(0, SigAbbrev_Zindex)
  Me.TextBox4.Value = PotentialTyp 'Signalrichtung

End Sub

Private Sub SigCharCombo_Change()
SigAbbrevCobo.ZOrder (fmBottom)
'SigChar(100, 2) 'Signalcharakter
SigChar_Zindex = Me.SigCharCombo.ListIndex
Funktion = Me.CoBoFunktion.Value 'Urban23.11.18
If SigChar_Zindex <= 0 Then SigChar_Zindex = 0
  If Funktion = "" Or Funktion = "nFct" Or Newfct_state = True Then
   'eine Funktion belegt die Signalart und nur wenn keine
   'Funktion festgelegt ist dann setzt die Spanunng den Wert
   SignalArt = SigCharCombo.Column(0, SigChar_Zindex)
   Me.TextBox5.Value = SignalArt 'Signalrichtung
  End If
End Sub


Private Sub Terminate_CB_Click()
  Unload Me
  If Application.Visible = False Then
     Application.Visible = True ' EXCEL wieder sichtbar machen
  End If
  Workbooks(VFbl).Activate
  If BNAStat = True Then
     ZSB_KLemmeninfo = "*" & ZSB_Klemme & ZSB_Funktion & ZsbOertlichkeit
     Else: ZSB_KLemmeninfo = ZSB_Klemme & ZSB_Funktion & ZsbOertlichkeit
  End If

  Worksheets("Formblatt").Visible = True

End Sub

Private Sub ToggleButton1_Click()
         Label6.Visible = True
         Label6.Visible = True
         Label7.Visible = True
         Label8.Visible = True
         Label9.Visible = True
         Label42.Visible = True
         Label43.Visible = True
         Label44.Visible = True
         Label45.Visible = True
         Label107.Visible = True
         TextBox1.Visible = True
         TextBox4.Visible = True
         TextBox5.Visible = True
         TextBox6.Visible = True
         SigAbbrevCobo.Visible = True
         SigCharCombo.Visible = True
         Voltage_CoBo.Visible = True
         IOX_ComBo.Visible = True
If ToggleButton1.Value = True Then
   ToggleButton1.BackColor = RGB(200, 255, 200)
   ToggleButton1.Caption = "Edit NewDef"
   UserForm2.TextBox2.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox3.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox7.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox8.BackColor = RGB(255, 255, 245)
   UserForm2.TextBox9.BackColor = RGB(255, 255, 245)
   'UserForm2.Label18.Enabled = False
   Workbooks(VFbl).Activate
   Worksheets("Formblatt").Activate
   UserForm2.Image1.Visible = False
   'UserForm2.TextBox2.Value = "Beschreibung"
   'UserForm2.TextBox3.Value = "Description"
   Signalrichtung_pre = Me.TextBox6.Value
   'SigCharCombo.DropButtonClick 'Urban 19.03.2021
   IOX_ComBo.Text = Signalrichtung_pre
   Else: UserForm2.Image1.Visible = True
         ToggleButton1.BackColor = RGB(255, 255, 200)
         ToggleButton1.Caption = "ActivateNewDef"

IOX_ComBo.Visible = True
Voltage_CoBo.Visible = True
End If
    If tmp_FnctIdx >= 1 Then
    OptionButton1.Value = 1
    End If
End Sub

Private Sub UserForm_Initialize()
' Die Daten werden im Code Zeilenweise eingetragen
  Dim Zeile As Long
  Dim Spalte As Long
  'UserForm2.Image1.BackStyle = fmBackStyleOpaque
  '4WD;Allrad;4 Wheel Drive;mixed;S;CRX
 With Me.CoBoFunktion
  Me.CoBoFunktion.Clear
   For Zeile = 1 To UBound(FunctionsSpeicher, 1)
         If FunctionsSpeicher(Zeile, Spalte) <> "" Then ' And Not FunctionsSpeicher(Zeile, Spalte) = "nFct" Then
           CCF_ActVal = FunctionsSpeicher(Zeile, Spalte)
           If CCF_PreVal = "" And CCF_ActVal = "" Then Exit For 'die Schleife abbrechen, da mehrmals kein Wert vorhanden
           UserForm2.CoBoFunktion.AddItem FunctionsSpeicher(Zeile, 0)                               'function
           UserForm2.CoBoFunktion.List(CoBoFunktion.ListCount - 1, 1) = FunctionsSpeicher(Zeile, 1) 'Beschreibung_EN
           UserForm2.CoBoFunktion.List(CoBoFunktion.ListCount - 1, 2) = FunctionsSpeicher(Zeile, 2) 'Beschreibung_DE
           UserForm2.CoBoFunktion.List(CoBoFunktion.ListCount - 1, 3) = FunctionsSpeicher(Zeile, 3) 'SigCharacter mixed, CAN, ND
           UserForm2.CoBoFunktion.List(CoBoFunktion.ListCount - 1, 4) = FunctionsSpeicher(Zeile, 4) 'Signalabbreviation S,V,M
           UserForm2.CoBoFunktion.List(CoBoFunktion.ListCount - 1, 5) = FunctionsSpeicher(Zeile, 5) 'direction  CRX, IN, OUT
           Else: CCF_PreVal = "" ' kein Wert in DB vorhanden
         End If
   Next
   
 End With
With Me.CoBoUtili
 Me.CoBoUtili.Clear
   For Zeile = 1 To UBound(UtiliSpeicher, 1)
         If UtiliSpeicher(Zeile, Spalte) <> "" Then 'And Not UtiliSpeicher(Zeile, Spalte) = "nVal" Then 'UtiliSpeicher
           UserForm2.CoBoUtili.AddItem UtiliSpeicher(Zeile, 0)
           UserForm2.CoBoUtili.List(CoBoUtili.ListCount - 1, 1) = UtiliSpeicher(Zeile, 1)
           UserForm2.CoBoUtili.List(CoBoUtili.ListCount - 1, 2) = UtiliSpeicher(Zeile, 2)
           UserForm2.CoBoUtili.List(CoBoUtili.ListCount - 1, 3) = UtiliSpeicher(Zeile, 3)
         End If
   Next
 End With
 With Me.clampCoBo
   Me.clampCoBo.Clear ' Inhalt leeren
   For Zeile = 1 To UBound(clampSpeicher, 1)
           CCB_ActVal = clampSpeicher(Zeile, Spalte)
           If CCB_PreVal = "" And CCB_ActVal = "" Then Exit For 'die Schleife abbrechen, da mehrmals kein Wert vorhanden
           If clampSpeicher(Zeile, Spalte) <> "" And clampSpeicher(Zeile, 1) <> "Redncy" Then    'Urban 02.Mrz20
           UserForm2.clampCoBo.AddItem clampSpeicher(Zeile, 0)  'eine Zeile zufügen und CL_Info Attribute füllen
           UserForm2.clampCoBo.List(clampCoBo.ListCount - 1, 2) = clampSpeicher(Zeile, 0)  '0 CL (Klemme)
           If clampSpeicher(Zeile, 1) <> "nFct" Then
              UserForm2.clampCoBo.List(clampCoBo.ListCount - 1, 0) = clampSpeicher(Zeile, 1)  '1 function (Funktion)
              Else: UserForm2.clampCoBo.List(clampCoBo.ListCount - 1, 0) = "Klemme# " & clampSpeicher(Zeile, 0) '1 function (Funktion)
           End If
           
           UserForm2.clampCoBo.List(clampCoBo.ListCount - 1, 1) = clampSpeicher(Zeile, 2)  'Funct desc text (Englisch)
           UserForm2.clampCoBo.List(clampCoBo.ListCount - 1, 3) = clampSpeicher(Zeile, 4)  'Utilization fL, rR, 3rD
           UserForm2.clampCoBo.List(clampCoBo.ListCount - 1, 4) = clampSpeicher(Zeile, 10) 'direction   CRX, IN, OUT
           UserForm2.clampCoBo.List(clampCoBo.ListCount - 1, 5) = clampSpeicher(Zeile, 7)  'voltage     12, 48
           UserForm2.clampCoBo.List(clampCoBo.ListCount - 1, 6) = clampSpeicher(Zeile, 8)  'SigCharacter CAN, ND
           UserForm2.clampCoBo.List(clampCoBo.ListCount - 1, 7) = clampSpeicher(Zeile, 9)  'Signalabbreviation S,V,M
           UserForm2.clampCoBo.List(clampCoBo.ListCount - 1, 8) = clampSpeicher(Zeile, 5)  'Util desc text
           UserForm2.clampCoBo.List(clampCoBo.ListCount - 1, 9) = clampSpeicher(Zeile, 11) 'Cl desc text
               Else: CCB_PreVal = "" ' kein Wert in DB vorhanden
         End If
   Next
 End With
 
'Cl_ListeBox KLemmen einlesen
 With Me.Cl_ListeBox
  Me.Cl_ListeBox.Clear ' Inhalt leeren
   For Zeile = 1 To UBound(clamps, 1) 'clamps(100, 6) Low;Low-Side Schalter;low side switch;M;DC;IN;12
         If clamps(Zeile, Spalte) <> "" Then
           UserForm2.Cl_ListeBox.AddItem clamps(Zeile, 0) 'KLemme
           UserForm2.Cl_ListeBox.List(Cl_ListeBox.ListCount - 1, 1) = clamps(Zeile, 1) 'EN Description
           UserForm2.Cl_ListeBox.List(Cl_ListeBox.ListCount - 1, 2) = clamps(Zeile, 2) 'DE Description
           UserForm2.Cl_ListeBox.List(Cl_ListeBox.ListCount - 1, 3) = clamps(Zeile, 3) ' SAbbrevation
           UserForm2.Cl_ListeBox.List(Cl_ListeBox.ListCount - 1, 4) = clamps(Zeile, 4) ' SCharacter
           UserForm2.Cl_ListeBox.List(Cl_ListeBox.ListCount - 1, 5) = clamps(Zeile, 5) ' Direction
           UserForm2.Cl_ListeBox.List(Cl_ListeBox.ListCount - 1, 6) = clamps(Zeile, 6) ' Voltage
         End If
   Next
   
 End With
'Voltage_CoBo = Spannung
With Me.Voltage_CoBo
  Me.Voltage_CoBo.Clear ' Inhalt leeren
   For Zeile = 1 To UBound(Volt, 1) '
         If Volt(Zeile, Spalte) <> "" Then
           UserForm2.Voltage_CoBo.AddItem Volt(Zeile, 0) 'Voltage, 7V5;Sensors Power Supply;Sensorspannung II;DC
           UserForm2.Voltage_CoBo.List(Voltage_CoBo.ListCount - 1, 1) = Volt(Zeile, 1) 'EN Description
           UserForm2.Voltage_CoBo.List(Voltage_CoBo.ListCount - 1, 2) = Volt(Zeile, 2) 'DE Description
           UserForm2.Voltage_CoBo.List(Voltage_CoBo.ListCount - 1, 3) = Volt(Zeile, 3) ' VoltageType
         End If
   Next
 End With
 
 'Richtungen
 With Me.IOX_ComBo
  Me.IOX_ComBo.Clear ' Inhalt leeren
   For Zeile = 1 To UBound(IOXs, 1) '
         If IOXs(Zeile, Spalte) <> "" Then
           UserForm2.IOX_ComBo.AddItem IOXs(Zeile, 0) 'IN, OUT, BI
           UserForm2.IOX_ComBo.List(IOX_ComBo.ListCount - 1, 1) = IOXs(Zeile, 1) 'EN Description
           UserForm2.IOX_ComBo.List(IOX_ComBo.ListCount - 1, 2) = IOXs(Zeile, 2) 'DE Description
         End If
   Next
 End With
'SigAbbrev(10, 1) 'CL_SignalAbbreviation
 With Me.SigAbbrevCobo
  Me.SigAbbrevCobo.Clear ' Inhalt leeren
   For Zeile = 1 To UBound(SigAbbrev, 1)
         If SigAbbrev(Zeile, Spalte) <> "" Then
           UserForm2.SigAbbrevCobo.AddItem SigAbbrev(Zeile, 0) 'A;Antennen,D,H,M,N,S,V;Versorgungsverbindungen
           UserForm2.SigAbbrevCobo.List(SigAbbrevCobo.ListCount - 1, 1) = SigAbbrev(Zeile, 1) 'DE Description
         End If
   Next
 End With
 
 'SigChar(100, 2) 'Signalcharakter
 With Me.SigCharCombo
  Me.SigCharCombo.Clear ' Inhalt leeren
   For Zeile = 1 To UBound(SigChar, 1)
         If SigChar(Zeile, Spalte) <> "" Then
           UserForm2.SigCharCombo.AddItem SigChar(Zeile, 0) 'IN, OUT, BI
           UserForm2.SigCharCombo.List(SigCharCombo.ListCount - 1, 1) = SigChar(Zeile, 1) 'EN Description
           UserForm2.SigCharCombo.List(SigCharCombo.ListCount - 1, 2) = SigChar(Zeile, 2) 'DE Description
         End If
   Next
 End With
 
With Me.CommandButton1
  .Visible = True
  .Locked = False
End With
If frage <> "yes" Then
  Me.CommandButton1.Visible = False ' Das Speichern ist nicht gewollt
  Me.CommandButton1.Locked = True
  
End If
 
Me.Label12.Caption = SteckplzPin
Me.Label35.Caption = SteckplzPinInfo
Me.Label37.Caption = SteckplzPinTypInf 'Urban 06.02.2020 Ausgabe Richtung und PinTyp zugefügt
Me.Label38.Caption = SteckplzPinTypI

'Me.Label15.Caption = defKLinhalt
UserForm2.Label32.Visible = False ' Textmarker unsichtbar
genPinDaten_actCell ' alte Daten holen und Felder befüllen
Workbooks(VFbl).Activate
Worksheets("Formblatt").Visible = True

Load Me
Me.Show
'dende = "das Ende"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then ' X wird gedrückt
        Application.Visible = True ' EXCEL wieder sichtbar machen
        Workbooks(VFbl).Activate
        Worksheets("Formblatt").Visible = True
        'ActiveWindow.WindowState = xlMaximized
        'End 'UGU?
    Else
      Application.Visible = True ' EXCEL wieder sichtbar machen
    End If
    
End Sub


Private Sub UserForm_Terminate()
  Unload Me
  If Application.Visible = False Then
     Application.Visible = True ' EXCEL wieder sichtbar machen
  End If
  Workbooks(VFbl).Activate
  Worksheets("Formblatt").Visible = True
End Sub


Public Function CheckConnectionIdx()
  PinZaehlerChk = UBound(PinIndexVec, 1)

End Function

Public Function genPinDaten_actCell()
' holt die BGER Pin definitionen aus dem Formblatt
aktiveRCelle = ActiveCell.Row
ZwSp = Cells(aktiveRCelle, 33).Value
ZwSp2 = Cells(aktiveRCelle, 226).Value  ' Urban 20.06.19, BGERid holen
ZwSpl = Len(ZwSp)
    Klemme = ""
    Beschreibung_DE = ""
    Beschreibung_EN = ""
    PotentialTyp = ""
    SignalArt = ""
    Signalrichtung = ""
    Funktion = ""
    Oertlichkeit = ""
If ZwSp = "ND" And ZwSp2 = "" Then ' Urban 20.06.19
    Klemme = ""
    Beschreibung_DE = ""
    Beschreibung_EN = ""
    PotentialTyp = ""
    SignalArt = ""
    Signalrichtung = ""
    Funktion = ""
    Oertlichkeit = ""
    Exit Function
End If
  KLinfoIdx = Cells(aktiveRCelle, 134).Value
  SignalArt_pre = Cells(aktiveRCelle, 176).Value 'SignalArt
  SignalArt = Cells(aktiveRCelle, 176).Value 'SignalArt
  SignalArt_pre = Cells(aktiveRCelle, 176).Value
  Signalrichtung = Cells(aktiveRCelle, 158).Value
  Signalrichtung_pre = Signalrichtung
  Spannung_pre = Cells(aktiveRCelle, 163).Value
  Klemme = Cells(aktiveRCelle, 136).Value
  Funktion = Cells(aktiveRCelle, 143).Value
  If Klemme = "nCL" Then
     Klemme = ""
  End If
  Spannungstyp = Cells(aktiveRCelle, 167).Value ' die Funktion gibt die den Signal Charakter in der UserForm vor
  Oertlichkeit = Cells(aktiveRCelle, 152).Value
  PotentialTyp = Cells(aktiveRCelle, 171).Value
'Urban 20.06.19
  ZSB_KLAttribute = Klemme & Funktion & Oertlichkeit & Signalrichtung & Spannung_pre & PotentialTyp & SignalArt
  ManipulationZiffer (ZSB_KLAttribute)  ' Wert 0 = keine BGER Werte vorhanden 'Funktion Prüfsumme bilden
  Cells(aktiveRCelle, 33).Value = ZSB_KLziffer 'fehlende Prüfsumme in Zelle schreiben
'Urban 20.06.19
'Cells(aktive_Reihe, 183).Value = ZSB_KLemme
'Felder der UserForm mit Werten füllen
    UserForm2.Cl_ListeBox.Value = Klemme
    UserForm2.TextBox1.Value = Klemme
    If SignalArt_pre <> SignalArt Then SignalArt = SignalArt_pre ' eine Temporäre Signalart gewinnt
'check temporäre Funktion vorhanden und Beschreibungen ermitteln
'Beispiel [TmpFnct: GruCkAnz# Showler* GangAnzeige]
    tmp_Fnct_Cmnt = Cells(aktiveRCelle, 183).Value 'globale Var
    Ln_tmpFnctCmnt = Len(tmp_Fnct_Cmnt)              'Länge des Kommentar feststellen
    tmp_FnctIdx = InStr(1, tmp_Fnct_Cmnt, "TmpFnct") 'Suche nach temporärer Funktion im Kommentar
UserForm2.CoBoFunktion.Text = Funktion ' Übergabe der Funktion an CoBoFunktion, tmp_FnctIdx <> "" behält die Funktion bei
CoBoFunktion_DropButtonClick           ' Funktion einlesen
     If tmp_FnctIdx >= 1 Then
       UserForm2.Label32.Visible = True ' Textmarker sichtbar
       RauteImtxt = InStr(1, tmp_Fnct_Cmnt, "#", vbTextCompare)
       SternImtxt = InStr(1, tmp_Fnct_Cmnt, "*", vbTextCompare)
       newEN_Cmnt = Mid(tmp_Fnct_Cmnt, RauteImtxt + 1, SternImtxt - RauteImtxt - 1) ' Urban 28.06.19
       newDE_Cmnt = Mid(tmp_Fnct_Cmnt, SternImtxt + 1, Ln_tmpFnctCmnt - SternImtxt - 1)
       UserForm2.OptionButton1.Value = 1 ' UGU
       UserForm2.SigCharCombo.Value = Cells(aktiveRCelle, 176).Value 'SignalArt, Urban 07.07.2021
       UserForm2.TextBox7.Value = newEN_Cmnt ' "Description"
       UserForm2.TextBox9.Value = newDE_Cmnt ' "Beschreibung"
     End If
    UserForm2.CoBoUtili.Text = Oertlichkeit
    UserForm2.Voltage_CoBo.Text = Spannung_pre
    UserForm2.TextBox4.Value = PotentialTyp
    UserForm2.TextBox5.Value = SignalArt_pre ' Urban 07.07.2021
    UserForm2.TextBox6.Value = Signalrichtung_pre
    UserForm2.TextBox6.BackColor = RGB(255, 255, 255)
    UserForm2.TextBox2.Value = Beschreibung_DE
    UserForm2.TextBox3.Value = Beschreibung_EN
End Function
'Urban 2021 März
Public Function Recherche(SO)
'Erase ReportSpeicher ' alter Speicher löschen
UserForm2.RepListB.Clear
PreCnt = 1
CutFnctCnt = 1
RepMemCnt = 1
If SO = "CL_Info" Then
  SuchSpeicher = clampSpeicher
  SuchIDX = 3  '0  wird benötigt, wenn keine Funktion "nFct" in der CL_Info
  UserForm2.Label26.Caption = " Index| Funktion                               Klemme  Utilization| Beschreibung in Deutsch"
End If
If SO = "FKT" Then
  SuchSpeicher = FunctionsSpeicher
  SuchIDX = 2
  UserForm2.Label26.Caption = " Index| Funktion                               Sigtyp      SigDir     | Beschreibung in Deutsch"
End If
'Suche in der Datenbasis FunctionsSpeicher
'FunctionsSpeicher(functN, 0) CL_Function
'FunctionsSpeicher(functN, 1) CL_descriptiveTextEN
'FunctionsSpeicher(functN, 2) CL_descriptiveTextDE
'FunctionsSpeicher(functN, 3) CL_SignalCharacteristic: PWM switched,DC,AC
'FunctionsSpeicher(functN, 4) CL_SignalAbbreviation: S,V,M,D
'FunctionsSpeicher(functN, 5) CL_direction: IN, OUT, IOX
'clampSpeicher(functN, 0)'clamp
'clampSpeicher(functN, 1)'cl_function
'clampSpeicher(functN, 2)'cl_descTextEN
'clampSpeicher(functN, 3)'cl_descTextDE
'clampSpeicher(functN, 4) 'Utilization fL, rR, 3rD
'clampSpeicher(functN, 11)'description

Suchtext = InputBox("Suche/ Search Function : ", "Suche nach einer Funktion in der Datenbasis", "sens")
   If Suchtext = "" Then Exit Function
   Suchtext = UCase("*" & Suchtext & "*") 'damit Treffer ermittelt werden können sind am Anfang und Ende des Suchtext ein * zugefügt'
For cntFsSr = 1 To 20000
  TxtSpeicherzelle = UCase(SuchSpeicher(cntFsSr, SuchIDX))   'Deutsche Textbeschreibung
  
    If SO = "CL_Info" Then
      TxtSpeicherzelle2 = UCase(SuchSpeicher(cntFsSr, 11)) 'Verwendungsbeschreibung
      'Urban 26.02.2021
      TxtSpeicherzelle3 = UCase(SuchSpeicher(cntFsSr, 1)) 'Funktion
    End If
  fnctSW = SuchSpeicher(cntFsSr, SuchIDX - 2)          ' Funktion ermitteln aus CL_Info/ Function
  
  While fnctSW = "Redncy" ' weiter im Schleifenzähler wenn keine Standards angewendet wurden
        cntFsSr = cntFsSr + 1
        If SO = "CL_Info" Then CutFnctCnt = CutFnctCnt + 1 'Urban 17.03.2021
        TxtSpeicherzelle = SuchSpeicher(cntFsSr, SuchIDX)
        fnctSW = SuchSpeicher(cntFsSr, SuchIDX - 2)
  Wend
  'Kl_Info hat keine Funktion nFct ist nur Klemme dann Suche im Beschreibungstext Urban 20.Feb 2020
  
  If SO = "FKT" Then
      TxtSpeicherzelle3 = UCase(SuchSpeicher(cntFsSr, 0)) 'Funktion
      test = TxtSpeicherzelle Like Suchtext 'testen ob Suchtext in deutscher Beschreibung enthalten ist
      test3 = TxtSpeicherzelle3 Like Suchtext       'testen ob Suchtext in deutscher Beschreibung enthalten ist
  End If
    
  If SO = "CL_Info" Then
      test = TxtSpeicherzelle Like Suchtext 'testen ob Suchtext in deutscher Beschreibung enthalten ist
      TxtSpeicherzelle2 = SuchSpeicher(cntFsSr, 11) 'Verwendungsbeschreibung
      TxtSpeicherzelle3 = UCase(SuchSpeicher(cntFsSr, 1)) 'Funktion
      test2 = TxtSpeicherzelle2 Like Suchtext       'testen ob Suchtext in deutscher Beschreibung enthalten ist
      test3 = TxtSpeicherzelle3 Like Suchtext       'testen ob Suchtext in deutscher Beschreibung enthalten ist
      Else: test2 = False
    End If
   If test = True Or test2 = True Or test3 = True Then
   RepMemCnt = RepMemCnt + 1 ' ReportSpeicher hochzählen
        If SO = "CL_Info" Then
           UtiTxt = SuchSpeicher(cntFsSr, 4)
              If UtiTxt = "nVal" Then UtiTxt = ""
           KlTxt = SuchSpeicher(cntFsSr, 0)
              If KlTxt = "nCL" Then KlTxt = ""
        End If
         If fnctSW <> "nFct" And SO = "CL_Info" Then 'wir haben eine Funktion in der CL_Info
            DescTxt = SuchSpeicher(cntFsSr, SuchIDX)
         End If
         If fnctSW = "nFct" And SO = "CL_Info" Then
            fnctSW = "Klemme# " & SuchSpeicher(cntFsSr, 0) 'echte Klemme wird für Ausgabe in Funktion umgebaut
            DescTxt = SuchSpeicher(cntFsSr, 11)
         End If
         If fnctSW <> "nFct" And SO = "FKT" Then ' wir haben eine reine Funktion
             KlTxt = SuchSpeicher(cntFsSr, 3)
             UtiTxt = SuchSpeicher(cntFsSr, 5)
             DescTxt = SuchSpeicher(cntFsSr, SuchIDX)
         End If
     IdxSuchSp = cntFsSr - CutFnctCnt
     UserForm2.RepListB.AddItem "empty"  'eine Zeile zufügen und CL_Info Attribute füllen
     UserForm2.RepListB.List(RepListB.ListCount - 1, 0) = IdxSuchSp
     UserForm2.RepListB.List(RepListB.ListCount - 1, 1) = fnctSW
     UserForm2.RepListB.List(RepListB.ListCount - 1, 2) = KlTxt
     UserForm2.RepListB.List(RepListB.ListCount - 1, 3) = UtiTxt
     UserForm2.RepListB.List(RepListB.ListCount - 1, 4) = DescTxt
     'FunktionenListe = FunktionenListe & IdxSuchSp & " [" & fnctSW & "]" & Fuellzeichen & vbTab & TxtSpeicherzelle & vbCrLf
   End If
Next
'MsgBox ("Standards in den Fnct: " & CutFnctCnt)
End Function

Public Function ChknCorValue(TT, WaT) As String
'WaT ist der Typ des Einganswertes F= Funktion, C = Beschreibung
'Urban 28.06.19, entfernt aus den teporären Funktionen und Beschreibungen die Zeichen # und *
  tmpW = ""
  cVlng = Len(TT)
  For k = 1 To cVlng
      ChkW = Mid(TT, k, 1)
      Ziffer = Asc(ChkW)
      If Ziffer = 94 Or Ziffer < 43 Or Ziff > 90 And Ziffer < 97 Or Ziffer > 122 Then ' / wir zugelassen UGU 11.06.2020
         ChkW = "_"
      End If
      
      If ChkW <> "#" And ChkW <> "*" And ChkW <> " " And WaT = "F" Then
           tmpW = tmpW & ChkW
      End If
      If ChkW <> "#" And ChkW <> "*" And WaT = "C" Then
           tmpW = tmpW & ChkW
      End If
  Next
ChknCorValue = tmpW

End Function

