Attribute VB_Name = "main_Modul"
Public strBestellnummer As String
Public strAuftragsnummer As String
Public strIDocNummer As String
Public strBestelljahr As String
Public strNameSheetEins As String

Dim strAuftrag As String
Dim xlBook As Workbook
Dim b_DisplayAlerts As Boolean 'zum Merken des alten Wertes
Dim intSpaltePosition As Integer
Dim intSpalteStatus As Integer
Dim intSpalteStatus_SAP As Integer
Dim inSpalteAlterWert As Integer
Dim intZeileIDocNummer As Integer
Dim intZeileBestellnummer As Integer
Dim intZeileAuftragsnummer As Integer
Dim intSpalteNummern As Integer
Dim intSpalteBestelljahr As Integer
Dim intZeileBestelljahr As Integer
Dim intZeileBestelldatum As Integer
Dim intSpalteBestelldatum As Integer
Dim strBestelldatum As String

#bei mir
Sub EDI_auslesen()

On Error GoTo SAP_Error

If Not IsObject(SAP_Application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set SAP_Application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = SAP_Application.Children(0)
End If
If Not IsObject(SAP_session) Then
   Set SAP_session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject SAP_session, "on"
   WScript.ConnectObject Application, "on"
End If

b_DisplayAlerts = Application.DisplayAlerts 'alten Wert merken
Application.DisplayAlerts = False 'zum Abschalten jeglicher Fehlermeldungen
intSpaltePosition = 2
intSpalteStatus = 3
intSpalteStatus_SAP = 5
intSpalteAlterWert = 4
intZeileBestelldatum = 6
intZeileIDocNummer = 5
intZeileBestellnummer = 3
intZeileAuftragsnummer = 4
intSpalteNummern = 6
intSpalteBestelljahr = 7
intZeileBestelljahr = 3
intSpalteEDI_Pos = 2
intSpalteEDI_Produkt = 3
intSpalteEDI_Menge = 4
intSpalteEDI_Betrag = 5
intSpalteEDI_Waehrung = 6
intSpalteEDI_PTGL = 7
intSpaltePartner = 10
intZeileRolle = 21
intZeileName = 22
intZeileName2 = 23
intZeileName3 = 24
intZeileStrasse = 25
intZeileStrasse2 = 26
intZeileStrasse3 = 27
intZeileOrt = 28
intZeileRegion = 29
intZeilePLZ = 30
intZeileLandschluessel = 31


Set xlBook = ActiveWorkbook
strNameSheetEins = "Start EDI"

strBestellnummer = xlBook.Sheets(strNameSheetEins).Cells(intZeileBestellnummer, intSpalteNummern).Value

If xlBook.Sheets(strNameSheetEins).Cells(intZeileBestelljahr, intSpalteBestelljahr).Value <> "" Then
    strBestelljahr = CStr(xlBook.Sheets(strNameSheetEins).Cells(intZeileBestelljahr, intSpalteBestelljahr).Value)
ElseIf xlBook.Sheets(strNameSheetEins).Cells(intZeileBestelljahr, intSpalteBestelljahr).Value = "" Then
    MyDate = Date
    strBestelljahr = CStr(Year(MyDate))
End If

'########################################################################################################################################################################
'SAP
'########################################################################################################################################################################
SAP_session.findById("wnd[0]").maximize
SAP_session.findById("wnd[0]/tbar[0]/okcd").Text = "/nydr3"
SAP_session.findById("wnd[0]").sendVKey 0
SAP_session.findById("wnd[0]/usr/ctxtMESTYP").Text = "ORDERS"
SAP_session.findById("wnd[0]/usr/ctxtDIRECT").Text = "2"
'###Auswahl richtiges Format für Datum:###############################################################################################################
SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-LOW").Text = "01.01." & strBestelljahr
SAP_session.findById("wnd[0]").sendVKey 0
If SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form __.__.____ ein" Then
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-LOW").Text = "01.01." & strBestelljahr
    SAP_session.findById("wnd[0]").sendVKey 0
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-HIGH").Text = "31.12." & strBestelljahr
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form __/__/____ ein" Then
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-LOW").Text = "01/01/" & strBestelljahr
    SAP_session.findById("wnd[0]").sendVKey 0
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-HIGH").Text = "12/31/" & strBestelljahr
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form __-__-____ ein" Then
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-LOW").Text = "01-01-" & strBestelljahr
    SAP_session.findById("wnd[0]").sendVKey 0
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-HIGH").Text = "12-31-" & strBestelljahr
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form ____.__.__ ein" Then
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-LOW").Text = strBestelljahr & ".01.01"
    SAP_session.findById("wnd[0]").sendVKey 0
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-HIGH").Text = strBestelljahr & ".12.31"
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form ____/__/__ ein" Then
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-LOW").Text = strBestelljahr & "/01/01"
    SAP_session.findById("wnd[0]").sendVKey 0
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-HIGH").Text = strBestelljahr & "/12/31"
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form ____-__-__ ein" Then
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-LOW").Text = strBestelljahr & "-01-01"
    SAP_session.findById("wnd[0]").sendVKey 0
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-HIGH").Text = strBestelljahr & "-12-31"
    SAP_session.findById("wnd[0]").sendVKey 0
Else
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-HIGH").Text = "31.12." & strBestelljahr
    SAP_session.findById("wnd[0]").sendVKey 0
End If

'########################################################################################################################################################
SAP_session.findById("wnd[0]/usr/btnBUT1").press
SAP_session.findById("wnd[1]/usr/txtBESTNR-LOW").Text = strBestellnummer
SAP_session.findById("wnd[1]/tbar[0]/btn[0]").press
SAP_session.findById("wnd[1]/tbar[0]/btn[8]").press
SAP_session.findById("wnd[0]/tbar[1]/btn[8]").press
strIDocNummer = SAP_session.findById("wnd[0]/usr/lbl[12,1]").Text
strBestelldatum = SAP_session.findById("wnd[0]/usr/lbl[12,5]").Text
'#######################################################################################################################################################################



bolErrorFlag = False

'In Excel eintragen:
xlBook.Sheets(strNameSheetEins).Cells(intZeileIDocNummer, intSpalteNummern).Value = strIDocNummer
xlBook.Sheets(strNameSheetEins).Cells(intZeileBestelldatum, intSpalteNummern).Value = strBestelldatum

On Error Resume Next

'###################################
' EDI auslesen
'###################################

dblRowCountEDI = 0
'strIDocNummer = xlBook.Sheets(strNameSheetEins).Cells(intZeileEDI_Nummer, intSpalteEDI_Nummer).Value
strTechName = ""
strBeschreibung = ""
strWert = ""
strWertNext = ""        '##############
strBestelldatum = ""    'Variable doppelt belegt????
strORG_ID_Besteller = ""
strIncoTerm = ""
strIncoTerm_Ort = ""

'Transaktion öffnen
SAP_session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
SAP_session.findById("wnd[0]").sendVKey 0
SAP_session.findById("wnd[0]/tbar[0]/okcd").Text = "we05"
SAP_session.findById("wnd[0]").sendVKey 0

'Maske für Auswahl ausfüllen
SAP_session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-LOW").Text = "2000-01-01"

'###Auswahl richtiges Format für Datum:###############################################################################################################
SAP_session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-LOW").Text = "01.01.2000"
SAP_session.findById("wnd[0]").sendVKey 0
If SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form __.__.____ ein" Then
    SAP_session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-LOW").Text = "01.01.2000"
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form __/__/____ ein" Then
    SAP_session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-LOW").Text = "01/01/2000"
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form __-__-____ ein" Then
    SAP_session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-LOW").Text = "01-01-2000"
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form ____.__.__ ein" Then
    SAP_session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-LOW").Text = "2000.01.01"
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form ____/__/__ ein" Then
    SAP_session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-LOW").Text = "2000/01/01"
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form ____-__-__ ein" Then
    SAP_session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-LOW").Text = "2000-01-01"
    SAP_session.findById("wnd[0]").sendVKey 0
End If

'########################################################################################################################################################

SAP_session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/txtDOCNUM-LOW").Text = strIDocNummer
SAP_session.findById("wnd[0]/tbar[1]/btn[8]").press

'Druck der EDI ausrufen
SAP_session.findById("wnd[0]/mbar/menu[0]/menu[0]").Select

'Gridview erfassen --> Grid enthält alle Infos der Bestellung
Set gridViewEDI_Daten = SAP_session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell") 'auslesen Tabelle / Grid aus SAP und speichern in Variable
dblRowCountEDI = gridViewEDI_Daten.RowCount
'gridViewEDI_Daten.getcellvalue(0, "TECHNAME")  'Name für Wert im EDI
'gridViewEDI_Daten.getcellvalue(0, "DESCRIPTION")  'Beschreibung der Position
'gridViewEDI_Daten.getcellvalue(0, "VALUE")  'Wert im EDI

dblPosWert = 0
dblBestellwert = 0
intAnzahlPos = 0
'Gridview auslesen, letzt Zeile wird nie gelesen

Dim arrEDI_PosDaten() As Variant  'Array für EDI-Daten


For i = 0 To dblRowCountEDI - 2 'Um Anzahl Positionen zu erfassen
    'scrollen um Werte richtig auszulesen
    SAP_session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = i
    
    strTechName = gridViewEDI_Daten.getcellvalue(i, "TECHNAME")  'Wert/Inhalt des Segmentes
    If strTechName Like "Y1082" Then    '"Y1082" = Positionsnummer (darauf folgen die Infos zu dieser Position)
        intAnzahlPos = intAnzahlPos + 1
    End If
Next i

SAP_session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 0


'define sequence of information in Array / dataset
intArrPos = 0
intArrPosProd = 1
intArrPosMenge = 2
intArrPosBetrag = 3
intArrPosWaehrung = 4
intArrPosPTGL = 5

intTabStartzeile = 15 'Startzeile für das Schreiben in die Tabelle

ReDim arrEDI_PosDaten(intAnzahlPos - 1, 6) 'Array für die Aufnahme aller Positionsdaten definieren, Länge Variabel auf die Anzahl der Positionen einstellen
                                           ' Bsp. 3 Pos. --> (0,1,2)*(0,1,2,3,4,5,6) Array

intPosWrite = -1 'define the Pos where to write the dataset
For i = 0 To dblRowCountEDI - 2
    
    
    'scrollen um Werte richtig auszulesen
    SAP_session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = i
    
    strTechName = gridViewEDI_Daten.getcellvalue(i, "TECHNAME")  'Wert/Inhalt des Segmentes
    strBeschreibung = gridViewEDI_Daten.getcellvalue(i, "DESCRIPTION")  'Wert/Inhalt des Segmentes
    strWert = gridViewEDI_Daten.getcellvalue(i, "VALUE")  'Wert/Inhalt des Segmentes
    strWertNext = gridViewEDI_Daten.getcellvalue(i + 1, "VALUE") 'Wert/Inhalt des Nachfolge-Segmentes
    
    'Partnerdaten zu Endkunde auslesen
    If strTechName Like "Y3035" And strWert Like "*Endkunde*" Then
        strRolle = strWert
        For n3 = 1 To 12
            If gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3036" Then        'Name
                strName = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            ElseIf gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3036A" Then   'Name 2
                strName2 = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            ElseIf gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3036B" Then   'Name 3
                strName3 = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            ElseIf gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3042" Then    'Strasse
                strStrasse = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            ElseIf gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3042A" Then   'Strasse2
                strStrasse2 = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            ElseIf gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3042B" Then   'Strasse3
                strStrasse3 = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            ElseIf gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3164" Then    'Ort
                strOrt = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            ElseIf gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3229" Then    'Region
                strRegion = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            ElseIf gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3251" Then    'PLZ
                strPLZ = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            ElseIf gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3207" Then    'Land
                strLandschluessel = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            
            End If
        Next n3
    End If
    
    'festlegen welcher Datensatz geschrieben werden soll, hochzählen wenn neues EDI-Segment beginnt
    If strTechName Like "Y1082" Then
        intPosWrite = intPosWrite + 1
        arrEDI_PosDaten(intPosWrite, intArrPos) = strWert
    End If
    
    'Produkt/Leistungsbeschreibung auslesen                         #################################################
    If strBeschreibung Like "Produkt/*" And strTechName Like "Y7008" Then   'Müsste nicht strBeschreibung Like Pordukt sein??
        arrEDI_PosDaten(intPosWrite, intArrPosProd) = strWert       '################################################
    End If
    
    'Menge auslesen
    If strTechName Like "Y6060" And strBeschreibung Like "*Menge*" Then
        arrEDI_PosDaten(intPosWrite, intArrPosMenge) = strWert
    End If
    
    'Positionswert und Währung auslesen
    If strTechName Like "SEGNAM" And strWert Like "*Z1LMOA_D01*" Then
        For n2 = 1 To 10
            If gridViewEDI_Daten.getcellvalue(i + n2, "DESCRIPTION") = "Geldbetrag" Then
                arrEDI_PosDaten(intPosWrite, intArrPosBetrag) = CDbl(Replace(gridViewEDI_Daten.getcellvalue(i + n2, "VALUE"), ".", ","))
            End If
            If gridViewEDI_Daten.getcellvalue(i + n2, "DESCRIPTION") = "Währung, codiert" Then
                arrEDI_PosDaten(intPosWrite, intArrPosWaehrung) = gridViewEDI_Daten.getcellvalue(i + n2, "VALUE")
                Exit For
            End If
        Next n2
    End If
    
    'PTGL auslesen
    If strTechName Like "Y1153" And strBeschreibung Like "*Qualifier*" And strWert = "ADT" Then
        arrEDI_PosDaten(intPosWrite, intArrPosPTGL) = strWertNext
    End If
    
    If i > 300 Then '####################
        abc = 123   'Überhaupt verwendet?
    End If          '####################
    n2 = 0
    
Next i

SAP_session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"

'Ergebnisse in Tabelle schreiben
For i = 0 To intAnzahlPos - 1
    xlBook.Sheets(strNameSheetEins).Cells(intTabStartzeile + i, intSpalteEDI_Pos).Value = arrEDI_PosDaten(i, intArrPos)
    xlBook.Sheets(strNameSheetEins).Cells(intTabStartzeile + i, intSpalteEDI_Produkt).Value = arrEDI_PosDaten(i, intArrPosProd)
    xlBook.Sheets(strNameSheetEins).Cells(intTabStartzeile + i, intSpalteEDI_Menge).Value = arrEDI_PosDaten(i, intArrPosMenge)
    xlBook.Sheets(strNameSheetEins).Cells(intTabStartzeile + i, intSpalteEDI_Betrag).Value = arrEDI_PosDaten(i, intArrPosBetrag)
    xlBook.Sheets(strNameSheetEins).Cells(intTabStartzeile + i, intSpalteEDI_Waehrung).Value = arrEDI_PosDaten(i, intArrPosWaehrung)
    xlBook.Sheets(strNameSheetEins).Cells(intTabStartzeile + i, intSpalteEDI_PTGL).Value = arrEDI_PosDaten(i, intArrPosPTGL)
Next i

'Partnerinformationen in Tabelle schreiben
xlBook.Sheets(strNameSheetEins).Cells(intZeileRolle, intSpaltePartner).Value = strRolle
xlBook.Sheets(strNameSheetEins).Cells(intZeileName, intSpaltePartner).Value = strName
xlBook.Sheets(strNameSheetEins).Cells(intZeileName2, intSpaltePartner).Value = strName2
xlBook.Sheets(strNameSheetEins).Cells(intZeileName3, intSpaltePartner).Value = strName3
xlBook.Sheets(strNameSheetEins).Cells(intZeileStrasse, intSpaltePartner).Value = strStrasse
xlBook.Sheets(strNameSheetEins).Cells(intZeileStrasse2, intSpaltePartner).Value = strStrasse2
xlBook.Sheets(strNameSheetEins).Cells(intZeileStrasse3, intSpaltePartner).Value = strStrasse3
xlBook.Sheets(strNameSheetEins).Cells(intZeileOrt, intSpaltePartner).Value = strOrt
xlBook.Sheets(strNameSheetEins).Cells(intZeileRegion, intSpaltePartner).Value = strRegion
xlBook.Sheets(strNameSheetEins).Cells(intZeilePLZ, intSpaltePartner).Value = strPLZ
xlBook.Sheets(strNameSheetEins).Cells(intZeileLandschluessel, intSpaltePartner).Value = strLandschluessel

Application.DisplayAlerts = b_DisplayAlerts

MsgBox "EDI ausgelesen. Bitte Daten in der Tabelle prüfen.", vbOKOnly

Exit Sub

SAP_Error:
   MsgBox "Fehler. Programm wird beendet.", vbCritical
   Exit Sub

End Sub

Sub EDI_Change_auslesen()

On Error GoTo SAP_Error

If Not IsObject(SAP_Application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set SAP_Application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = SAP_Application.Children(0)
End If
If Not IsObject(SAP_session) Then
   Set SAP_session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject SAP_session, "on"
   WScript.ConnectObject Application, "on"
End If

b_DisplayAlerts = Application.DisplayAlerts 'alten Wert merken
Application.DisplayAlerts = False 'zum Abschalten jeglicher Fehlermeldungen
intSpaltePosition = 2
intSpalteStatus = 3
intSpalteStatus_SAP = 5
intSpalteAlterWert = 4
intZeileBestelldatum = 6
intZeileIDocNummer = 5
intZeileBestellnummer = 3
intZeileAuftragsnummer = 4
intSpalteNummern = 6
intSpalteBestelljahr = 7
intZeileBestelljahr = 3
intSpalteEDI_Pos = 2
intSpalteEDI_Produkt = 3
intSpalteEDI_Menge = 4
intSpalteEDI_Betrag = 5
intSpalteEDI_Waehrung = 6
intSpalteEDI_PTGL = 7
intSpaltePartner = 10
intZeileRolle = 21
intZeileName = 22
intZeileName2 = 23
intZeileName3 = 24
intZeileStrasse = 25
intZeileStrasse2 = 26
intZeileStrasse3 = 27
intZeileOrt = 28
intZeileRegion = 29
intZeilePLZ = 30
intZeileLandschluessel = 31


Set xlBook = ActiveWorkbook
strNameSheetEins = "Start EDI"

strBestellnummer = xlBook.Sheets(strNameSheetEins).Cells(intZeileBestellnummer, intSpalteNummern).Value

If xlBook.Sheets(strNameSheetEins).Cells(intZeileBestelljahr, intSpalteBestelljahr).Value <> "" Then
    strBestelljahr = CStr(xlBook.Sheets(strNameSheetEins).Cells(intZeileBestelljahr, intSpalteBestelljahr).Value)
ElseIf xlBook.Sheets(strNameSheetEins).Cells(intZeileBestelljahr, intSpalteBestelljahr).Value = "" Then
    MyDate = Date
    strBestelljahr = CStr(Year(MyDate))
End If

'########################################################################################################################################################################
'SAP
'########################################################################################################################################################################
SAP_session.findById("wnd[0]").maximize
SAP_session.findById("wnd[0]/tbar[0]/okcd").Text = "/nydr3"
SAP_session.findById("wnd[0]").sendVKey 0
SAP_session.findById("wnd[0]/usr/ctxtMESTYP").Text = "ORDCHG"
SAP_session.findById("wnd[0]/usr/ctxtDIRECT").Text = "2"
'###Auswahl richtiges Format für Datum:###############################################################################################################
SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-LOW").Text = "01.01." & strBestelljahr
SAP_session.findById("wnd[0]").sendVKey 0
If SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form __.__.____ ein" Then
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-LOW").Text = "01.01." & strBestelljahr
    SAP_session.findById("wnd[0]").sendVKey 0
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-HIGH").Text = "31.12." & strBestelljahr
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form __/__/____ ein" Then
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-LOW").Text = "01/01/" & strBestelljahr
    SAP_session.findById("wnd[0]").sendVKey 0
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-HIGH").Text = "12/31/" & strBestelljahr
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form __-__-____ ein" Then
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-LOW").Text = "01-01-" & strBestelljahr
    SAP_session.findById("wnd[0]").sendVKey 0
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-HIGH").Text = "12-31-" & strBestelljahr
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form ____.__.__ ein" Then
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-LOW").Text = strBestelljahr & ".01.01"
    SAP_session.findById("wnd[0]").sendVKey 0
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-HIGH").Text = strBestelljahr & ".12.31"
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form ____/__/__ ein" Then
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-LOW").Text = strBestelljahr & "/01/01"
    SAP_session.findById("wnd[0]").sendVKey 0
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-HIGH").Text = strBestelljahr & "/12/31"
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form ____-__-__ ein" Then
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-LOW").Text = strBestelljahr & "-01-01"
    SAP_session.findById("wnd[0]").sendVKey 0
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-HIGH").Text = strBestelljahr & "-12-31"
    SAP_session.findById("wnd[0]").sendVKey 0
Else
    SAP_session.findById("wnd[0]/usr/ctxtERST_DAT-HIGH").Text = "31.12." & strBestelljahr
    SAP_session.findById("wnd[0]").sendVKey 0
End If

'########################################################################################################################################################
SAP_session.findById("wnd[0]/usr/btnBUT1").press
SAP_session.findById("wnd[1]/usr/txtBESTNR2-LOW").Text = strBestellnummer
SAP_session.findById("wnd[1]/tbar[0]/btn[0]").press
SAP_session.findById("wnd[1]/tbar[0]/btn[8]").press
SAP_session.findById("wnd[0]/tbar[1]/btn[8]").press
strIDocNummer = SAP_session.findById("wnd[0]/usr/lbl[12,1]").Text
strBestelldatum = SAP_session.findById("wnd[0]/usr/lbl[33,5]").Text
'#######################################################################################################################################################################



bolErrorFlag = False

'In Excel eintragen:
xlBook.Sheets(strNameSheetEins).Cells(intZeileIDocNummer, intSpalteNummern).Value = strIDocNummer
xlBook.Sheets(strNameSheetEins).Cells(intZeileBestelldatum, intSpalteNummern).Value = strBestelldatum

On Error Resume Next

'###################################
' EDI auslesen
'###################################

dblRowCountEDI = 0
'strIDocNummer = xlBook.Sheets(strNameSheetEins).Cells(intZeileEDI_Nummer, intSpalteEDI_Nummer).Value
strTechName = ""
strBeschreibung = ""
strWert = ""
strWertNext = ""        '##############
strBestelldatum = ""    'Variable doppelt belegt????
strORG_ID_Besteller = ""
strIncoTerm = ""
strIncoTerm_Ort = ""

'Transaktion öffnen
SAP_session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
SAP_session.findById("wnd[0]").sendVKey 0
SAP_session.findById("wnd[0]/tbar[0]/okcd").Text = "we05"
SAP_session.findById("wnd[0]").sendVKey 0

'Maske für Auswahl ausfüllen
SAP_session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-LOW").Text = "2000-01-01"

'###Auswahl richtiges Format für Datum:###############################################################################################################
SAP_session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-LOW").Text = "01.01.2000"
SAP_session.findById("wnd[0]").sendVKey 0
If SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form __.__.____ ein" Then
    SAP_session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-LOW").Text = "01.01.2000"
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form __/__/____ ein" Then
    SAP_session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-LOW").Text = "01/01/2000"
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form __-__-____ ein" Then
    SAP_session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-LOW").Text = "01-01-2000"
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form ____.__.__ ein" Then
    SAP_session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-LOW").Text = "2000.01.01"
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form ____/__/__ ein" Then
    SAP_session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-LOW").Text = "2000/01/01"
    SAP_session.findById("wnd[0]").sendVKey 0
ElseIf SAP_session.findById("wnd[0]/sbar").Text Like "Geben Sie das Datum in der Form ____-__-__ ein" Then
    SAP_session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-LOW").Text = "2000-01-01"
    SAP_session.findById("wnd[0]").sendVKey 0
End If

'########################################################################################################################################################

SAP_session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/txtDOCNUM-LOW").Text = strIDocNummer
SAP_session.findById("wnd[0]/tbar[1]/btn[8]").press

'Druck der EDI ausrufen
SAP_session.findById("wnd[0]/mbar/menu[0]/menu[0]").Select

'Gridview erfassen --> Grid enthält alle Infos der Bestellung
Set gridViewEDI_Daten = SAP_session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell") 'auslesen Tabelle / Grid aus SAP und speichern in Variable
dblRowCountEDI = gridViewEDI_Daten.RowCount
'gridViewEDI_Daten.getcellvalue(0, "TECHNAME")  'Name für Wert im EDI
'gridViewEDI_Daten.getcellvalue(0, "DESCRIPTION")  'Beschreibung der Position
'gridViewEDI_Daten.getcellvalue(0, "VALUE")  'Wert im EDI

dblPosWert = 0
dblBestellwert = 0
intAnzahlPos = 0
'Gridview auslesen, letzt Zeile wird nie gelesen

Dim arrEDI_PosDaten() As Variant  'Array für EDI-Daten


For i = 0 To dblRowCountEDI - 2 'Um Anzahl Positionen zu erfassen
    'scrollen um Werte richtig auszulesen
    SAP_session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = i
    
    strTechName = gridViewEDI_Daten.getcellvalue(i, "TECHNAME")  'Wert/Inhalt des Segmentes
    If strTechName Like "Y1082" Then    '"Y1082" = Positionsnummer (darauf folgen die Infos zu dieser Position)
        intAnzahlPos = intAnzahlPos + 1
    End If
Next i

SAP_session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 0


'define sequence of information in Array / dataset
intArrPos = 0
intArrPosProd = 1
intArrPosMenge = 2
intArrPosBetrag = 3
intArrPosWaehrung = 4
intArrPosPTGL = 5

intTabStartzeile = 15 'Startzeile für das Schreiben in die Tabelle

ReDim arrEDI_PosDaten(intAnzahlPos - 1, 6) 'Array für die Aufnahme aller Positionsdaten definieren, Länge Variabel auf die Anzahl der Positionen einstellen
                                           ' Bsp. 3 Pos. --> (0,1,2)*(0,1,2,3,4,5,6) Array

intPosWrite = -1 'define the Pos where to write the dataset
For i = 0 To dblRowCountEDI - 2
    
    
    'scrollen um Werte richtig auszulesen
    SAP_session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = i
    
    strTechName = gridViewEDI_Daten.getcellvalue(i, "TECHNAME")  'Wert/Inhalt des Segmentes
    strBeschreibung = gridViewEDI_Daten.getcellvalue(i, "DESCRIPTION")  'Wert/Inhalt des Segmentes
    strWert = gridViewEDI_Daten.getcellvalue(i, "VALUE")  'Wert/Inhalt des Segmentes
    strWertNext = gridViewEDI_Daten.getcellvalue(i + 1, "VALUE") 'Wert/Inhalt des Nachfolge-Segmentes
    
    'Partnerdaten zu Endkunde auslesen
    If strTechName Like "Y3035" And strWert Like "*Endkunde*" Then
        strRolle = strWert
        For n3 = 1 To 12
            If gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3036" Then        'Name
                strName = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            ElseIf gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3036A" Then   'Name 2
                strName2 = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            ElseIf gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3036B" Then   'Name 3
                strName3 = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            ElseIf gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3042" Then    'Strasse
                strStrasse = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            ElseIf gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3042A" Then   'Strasse2
                strStrasse2 = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            ElseIf gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3042B" Then   'Strasse3
                strStrasse3 = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            ElseIf gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3164" Then    'Ort
                strOrt = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            ElseIf gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3229" Then    'Region
                strRegion = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            ElseIf gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3251" Then    'PLZ
                strPLZ = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            ElseIf gridViewEDI_Daten.getcellvalue(i + n3, "TECHNAME") = "Y3207" Then    'Land
                strLandschluessel = gridViewEDI_Daten.getcellvalue(i + n3, "VALUE")
            
            End If
        Next n3
    End If
    
    'festlegen welcher Datensatz geschrieben werden soll, hochzählen wenn neues EDI-Segment beginnt
    If strTechName Like "Y1082" Then
        intPosWrite = intPosWrite + 1
        arrEDI_PosDaten(intPosWrite, intArrPos) = strWert
    End If
    
    'Produkt/Leistungsbeschreibung auslesen                         #################################################
    If strBeschreibung Like "Produkt/*" And strTechName Like "Y7008" Then   'Müsste nicht strBeschreibung Like Pordukt sein??
        arrEDI_PosDaten(intPosWrite, intArrPosProd) = strWert       '################################################
    End If
    
    'Menge auslesen
    If strTechName Like "Y6060" And strBeschreibung Like "*Menge*" Then
        arrEDI_PosDaten(intPosWrite, intArrPosMenge) = strWert
    End If
    
    'Positionswert und Währung auslesen
    If strTechName Like "SEGNAM" And strWert Like "*Z1LMOA_D01*" Then
        For n2 = 1 To 10
            If gridViewEDI_Daten.getcellvalue(i + n2, "DESCRIPTION") = "Geldbetrag" Then
                arrEDI_PosDaten(intPosWrite, intArrPosBetrag) = CDbl(Replace(gridViewEDI_Daten.getcellvalue(i + n2, "VALUE"), ".", ","))
            End If
            If gridViewEDI_Daten.getcellvalue(i + n2, "DESCRIPTION") = "Währung, codiert" Then
                arrEDI_PosDaten(intPosWrite, intArrPosWaehrung) = gridViewEDI_Daten.getcellvalue(i + n2, "VALUE")
                Exit For
            End If
        Next n2
    End If
    
    'PTGL auslesen
    If strTechName Like "Y1153" And strBeschreibung Like "*Qualifier*" And strWert = "ADT" Then
        arrEDI_PosDaten(intPosWrite, intArrPosPTGL) = strWertNext
    End If
    
    If i > 300 Then '####################
        abc = 123   'Überhaupt verwendet?
    End If          '####################
    n2 = 0
    
Next i

SAP_session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"

'Ergebnisse in Tabelle schreiben
For i = 0 To intAnzahlPos - 1
    xlBook.Sheets(strNameSheetEins).Cells(intTabStartzeile + i, intSpalteEDI_Pos).Value = arrEDI_PosDaten(i, intArrPos)
    xlBook.Sheets(strNameSheetEins).Cells(intTabStartzeile + i, intSpalteEDI_Produkt).Value = arrEDI_PosDaten(i, intArrPosProd)
    xlBook.Sheets(strNameSheetEins).Cells(intTabStartzeile + i, intSpalteEDI_Menge).Value = arrEDI_PosDaten(i, intArrPosMenge)
    xlBook.Sheets(strNameSheetEins).Cells(intTabStartzeile + i, intSpalteEDI_Betrag).Value = arrEDI_PosDaten(i, intArrPosBetrag)
    xlBook.Sheets(strNameSheetEins).Cells(intTabStartzeile + i, intSpalteEDI_Waehrung).Value = arrEDI_PosDaten(i, intArrPosWaehrung)
    xlBook.Sheets(strNameSheetEins).Cells(intTabStartzeile + i, intSpalteEDI_PTGL).Value = arrEDI_PosDaten(i, intArrPosPTGL)
Next i

'Partnerinformationen in Tabelle schreiben
xlBook.Sheets(strNameSheetEins).Cells(intZeileRolle, intSpaltePartner).Value = strRolle
xlBook.Sheets(strNameSheetEins).Cells(intZeileName, intSpaltePartner).Value = strName
xlBook.Sheets(strNameSheetEins).Cells(intZeileName2, intSpaltePartner).Value = strName2
xlBook.Sheets(strNameSheetEins).Cells(intZeileName3, intSpaltePartner).Value = strName3
xlBook.Sheets(strNameSheetEins).Cells(intZeileStrasse, intSpaltePartner).Value = strStrasse
xlBook.Sheets(strNameSheetEins).Cells(intZeileStrasse2, intSpaltePartner).Value = strStrasse2
xlBook.Sheets(strNameSheetEins).Cells(intZeileStrasse3, intSpaltePartner).Value = strStrasse3
xlBook.Sheets(strNameSheetEins).Cells(intZeileOrt, intSpaltePartner).Value = strOrt
xlBook.Sheets(strNameSheetEins).Cells(intZeileRegion, intSpaltePartner).Value = strRegion
xlBook.Sheets(strNameSheetEins).Cells(intZeilePLZ, intSpaltePartner).Value = strPLZ
xlBook.Sheets(strNameSheetEins).Cells(intZeileLandschluessel, intSpaltePartner).Value = strLandschluessel

Application.DisplayAlerts = b_DisplayAlerts

MsgBox "EDI ausgelesen. Bitte Daten in der Tabelle prüfen.", vbOKOnly

Exit Sub

SAP_Error:
   MsgBox "Fehler. Programm wird beendet.", vbCritical
   Exit Sub

End Sub

Sub in_KDA_schreiben()

'On Error GoTo SAP_Error

If Not IsObject(SAP_Application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set SAP_Application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = SAP_Application.Children(0)
End If
If Not IsObject(SAP_session) Then
   Set SAP_session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject SAP_session, "on"
   WScript.ConnectObject Application, "on"
End If

b_DisplayAlerts = Application.DisplayAlerts 'alten Wert merken
Application.DisplayAlerts = False 'zum Abschalten jeglicher Fehlermeldungen
Set xlBook = ActiveWorkbook
strNameSheetEins = "Start EDI"

intSpaltePosition = 2
intSpalteStatus = 3
intSpalteStatus_SAP = 5
intSpalteAlterWert = 4
intZeileBestelldatum = 6
intZeileIDocNummer = 5
intZeileBestellnummer = 3
intZeileAuftragsnummer = 4
intSpalteNummern = 6
intSpalteBestelljahr = 7
intZeileBestelljahr = 3
intSpalteEDI_Pos = 2
intSpalteEDI_Produkt = 3
intSpalteEDI_Menge = 4
intSpalteEDI_Betrag = 5
intSpalteEDI_Waehrung = 6
intSpalteEDI_PTGL = 7

strBestellnummer = xlBook.Sheets(strNameSheetEins).Cells(intZeileBestellnummer, intSpalteNummern).Value

SAP_session.findById("wnd[0]").maximize
SAP_session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva02"
SAP_session.findById("wnd[0]").sendVKey 0
SAP_session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = ""
SAP_session.findById("wnd[0]/usr/txtRV45S-BSTNK").Text = strBestellnummer
SAP_session.findById("wnd[0]/usr/btnBT_SUCH").press

'Auswahl falls mehrere: ################    Abfangen falls mehrere zur Auswahl      ###########################
If SAP_session.findById("wnd[1]").Text Like "*Trefferliste 1 Eintrag*" Then
    SAP_session.findById("wnd[1]/tbar[0]/btn[0]").press
Else
    MsgBox "Zur angegeben Bestellnummer wurde mehr als 1 Beleg gefunden. Programm wird beendet.", vbCritical
    Exit Sub
End If

'1. Schleife um die Summenzeile zu finden --> Index (Zeilennummer merken)
For intZeileSAP = 0 To 99
    If SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1," & intZeileSAP & "]").Text Like "5*" Then 'And SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtKOMV-KBETR[5, " & intZeileSAP & "]").Text = "0,00" Then
        intZeileStart = intZeileSAP + 1
        Exit For
    End If
Next intZeileSAP

'PSP-Element, Materialnr, Auftragsnr speichern
strPSP = SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-PS_PSP_PNR[11," & intZeileStart & "]").Text
strMaterialnummer = SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1," & intZeileStart & "]").Text
strAuftragsnummer = SAP_session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-VBELN").Text
xlBook.Sheets(strNameSheetEins).Cells(intZeileAuftragsnummer, intSpalteNummern).Value = strAuftragsnummer

'2. Zeile löschen
strPos_alt = SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0," & intZeileStart & "]").Text
intPos_neu = CInt(strPos_alt) + 10
SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").getAbsoluteRow(intZeileStart).Selected = True
SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POLO").press
SAP_session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

'2. Schleife (SAP und Excel)
intZeileExcel = 15
intZeileAnzeige = 0
SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").verticalScrollbar.Position = 0

For intZeile = intZeileStart To 99
    SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").verticalScrollbar.Position = intZeileAnzeige
    If xlBook.Sheets(strNameSheetEins).Cells(intZeileExcel, intSpalteEDI_Pos).Value <> "" Then
        
        'Position:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0," & intZeile - intZeileAnzeige & "]").Text = intPos_neu
        'Materialnr:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1," & intZeile - intZeileAnzeige & "]").Text = strMaterialnummer
        'Auftragsmenge:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2," & intZeile - intZeileAnzeige & "]").Text = "1"
        'Mengeneinheit:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-VRKME[3," & intZeile - intZeileAnzeige & "]").Text = "ST"
        'Bezeichnung / Beschreibung:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-ARKTX[4," & intZeile - intZeileAnzeige & "]").Text = xlBook.Sheets(strNameSheetEins).Cells(intZeileExcel, intSpalteEDI_Produkt).Value
        'Betrag:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtKOMV-KBETR[5," & intZeile - intZeileAnzeige & "]").Text = xlBook.Sheets(strNameSheetEins).Cells(intZeileExcel, intSpalteEDI_Betrag).Value
        'Währung:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KOEIN[6," & intZeile - intZeileAnzeige & "]").Text = xlBook.Sheets(strNameSheetEins).Cells(intZeileExcel, intSpalteEDI_Waehrung).Value
        'PSP-Element:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-PS_PSP_PNR[11," & intZeile - intZeileAnzeige & "]").Text = strPSP
        'Hauptposition:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-YYBCHPOS[12," & intZeile - intZeileAnzeige & "]").Text = xlBook.Sheets(strNameSheetEins).Cells(intZeileExcel, intSpalteEDI_Pos).Value
        'PTGL:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-ZZPARTTGLP[33," & intZeile - intZeileAnzeige & "]").Text = xlBook.Sheets(strNameSheetEins).Cells(intZeileExcel, intSpalteEDI_PTGL).Value
        
        'Bestätigen
        SAP_session.findById("wnd[0]").sendVKey 0
        
        'Fehlermeldung wegklicken
        If SAP_session.findById("wnd[0]/sbar").Text Like "Kombination Siemens-H/-Upos * schon vorhanden" Then
            SAP_session.findById("wnd[0]").sendVKey 0
        End If
        
        'LZB ausfüllen
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").getAbsoluteRow(intZeile).Selected = True
        SAP_session.findById("wnd[0]/mbar/menu[2]/menu[2]/menu[14]/menu[4]").Select
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/ctxtVBAP-ZZ_LZBKZ").Text = "000"
        SAP_session.findById("wnd[0]").sendVKey 0
        SAP_session.findById("wnd[0]/tbar[0]/btn[3]").press
        
        'Scrollen
        intZeileAnzeige = intZeileAnzeige + 1
        
        'Nächste Zeile aus Excel
        intZeileExcel = intZeileExcel + 1
        
        'Nächste Position in SAP
        intPos_neu = intPos_neu + 10
        
    Else
        Exit For '(wenn alle Excel-Positionen erfasst --> Abbruch)
    End If
    
Next intZeile

Application.DisplayAlerts = b_DisplayAlerts 'alten Wert wieder restaurieren

MsgBox "Alles bearbeitet. Bitte prüfen und manuell speichern.", vbOKOnly

Exit Sub

End Sub


# bei mir
Sub in_KDA_schreiben_change()

'On Error GoTo SAP_Error

If Not IsObject(SAP_Application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set SAP_Application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = SAP_Application.Children(0)
End If
If Not IsObject(SAP_session) Then
   Set SAP_session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject SAP_session, "on"
   WScript.ConnectObject Application, "on"
End If

b_DisplayAlerts = Application.DisplayAlerts 'alten Wert merken
Application.DisplayAlerts = False 'zum Abschalten jeglicher Fehlermeldungen
Set xlBook = ActiveWorkbook
strNameSheetEins = "Start EDI"

intSpaltePosition = 2
intSpalteStatus = 3
intSpalteStatus_SAP = 5
intSpalteAlterWert = 4
intZeileBestelldatum = 6
intZeileIDocNummer = 5
intZeileBestellnummer = 3
intZeileAuftragsnummer = 4
intSpalteNummern = 6
intSpalteBestelljahr = 7
intZeileBestelljahr = 3
intSpalteEDI_Pos = 2
intSpalteEDI_Produkt = 3
intSpalteEDI_Menge = 4
intSpalteEDI_Betrag = 5
intSpalteEDI_Waehrung = 6
intSpalteEDI_PTGL = 7

strBestellnummer = xlBook.Sheets(strNameSheetEins).Cells(intZeileBestellnummer, intSpalteNummern).Value

SAP_session.findById("wnd[0]").maximize
SAP_session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva02"
SAP_session.findById("wnd[0]").sendVKey 0
SAP_session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = ""
SAP_session.findById("wnd[0]/usr/txtRV45S-BSTNK").Text = strBestellnummer
SAP_session.findById("wnd[0]/usr/btnBT_SUCH").press

'Auswahl falls mehrere: ################    Abfangen falls mehrere zur Auswahl      ###########################
If SAP_session.findById("wnd[1]").Text Like "*Trefferliste 1 Eintrag*" Then
    SAP_session.findById("wnd[1]/tbar[0]/btn[0]").press
Else
    MsgBox "Zur angegeben Bestellnummer wurde mehr als 1 Beleg gefunden. Programm wird beendet.", vbCritical
    Exit Sub
End If
If SAP_session.findById("wnd[1]").Text Like "Information" Then
    SAP_session.findById("wnd[1]/tbar[0]/btn[0]").press
End If
'1. Schleife um die Summenzeile zu finden --> Index (Zeilennummer merken)
For intZeileSAP = 0 To 99
    If SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1," & intZeileSAP & "]").Text Like "5*" Then                  'And SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtKOMV-KBETR[5, " & intZeileSAP & "]").Text = "0,00" Then
        intZeileStart = intZeileSAP + 1
        Exit For
    End If
Next intZeileSAP

'PSP-Element, Materialnr, Auftragsnr speichern
strPSP = SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-PS_PSP_PNR[11," & intZeileStart & "]").Text
strMaterialnummer = SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1," & intZeileStart & "]").Text
strAuftragsnummer = SAP_session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-VBELN").Text
xlBook.Sheets(strNameSheetEins).Cells(intZeileAuftragsnummer, intSpalteNummern).Value = strAuftragsnummer
intZeileStart = 0
'letzteZeile
SAP_session.findById("wnd[0]").resizeWorkingPane 155, 31, False
SAP_session.findById("wnd[0]/tbar[0]/btn[83]").press
SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,1]").SetFocus
SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,1]").caretPosition = 6

strPos_alt = SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0," & intZeileStart & "]").Text
intPos_neu = CInt(strPos_alt) + 10
'SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").getAbsoluteRow("0").Selected = True
'SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POLO").press
'SAP_session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

'2. Schleife (SAP und Excel)
intZeileExcel = 15
intZeileAnzeige = 0
'SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").verticalScrollbar.Position = 0
intZeileStart = 1
For intZeile = intZeileStart To 99
    'SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").verticalScrollbar.Position = intZeileAnzeige
    If xlBook.Sheets(strNameSheetEins).Cells(intZeileExcel, intSpalteEDI_Pos).Value <> "" Then
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0," & intZeile - intZeileAnzeige & "]").SetFocus
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0," & intZeile - intZeileAnzeige & "]").caretPosition = 6
        'Position:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0," & intZeile - intZeileAnzeige & "]").Text = intPos_neu
        'Materialnr:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1," & intZeile - intZeileAnzeige & "]").Text = strMaterialnummer
        'Auftragsmenge:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2," & intZeile - intZeileAnzeige & "]").Text = "1"
        'Mengeneinheit:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-VRKME[3," & intZeile - intZeileAnzeige & "]").Text = "ST"
        'Bezeichnung / Beschreibung:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-ARKTX[4," & intZeile - intZeileAnzeige & "]").Text = xlBook.Sheets(strNameSheetEins).Cells(intZeileExcel, intSpalteEDI_Produkt).Value
        'Betrag:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtKOMV-KBETR[5," & intZeile - intZeileAnzeige & "]").Text = xlBook.Sheets(strNameSheetEins).Cells(intZeileExcel, intSpalteEDI_Betrag).Value
        'Währung:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KOEIN[6," & intZeile - intZeileAnzeige & "]").Text = xlBook.Sheets(strNameSheetEins).Cells(intZeileExcel, intSpalteEDI_Waehrung).Value
        'PSP-Element:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-PS_PSP_PNR[11," & intZeile - intZeileAnzeige & "]").Text = strPSP
        'Hauptposition:
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-YYBCHPOS[12," & intZeile - intZeileAnzeige & "]").Text = xlBook.Sheets(strNameSheetEins).Cells(intZeileExcel, intSpalteEDI_Pos).Value
        'PTGL:
        If (xlBook.Sheets(strNameSheetEins).Cells(intZeileExcel, intSpalteEDI_PTGL).Value) = "SPG0708301" Then
            SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-ZZPARTTGLP[33," & intZeile - intZeileAnzeige & "]").Text = "SPG17081"
        Else
            SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-ZZPARTTGLP[33," & intZeile - intZeileAnzeige & "]").Text = xlBook.Sheets(strNameSheetEins).Cells(intZeileExcel, intSpalteEDI_PTGL).Value
        End If
        'Bestätigen
        SAP_session.findById("wnd[0]").sendVKey 0
        
        'Fehlermeldung wegklicken
        If SAP_session.findById("wnd[0]/sbar").Text Like "Kombination Siemens-H/-Upos * schon vorhanden" Then
            SAP_session.findById("wnd[0]").sendVKey 0
        End If
        
        'LZB ausfüllen
        'SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").getAbsoluteRow(intZeile).Selected = True
        'SAP_session.findById("wnd[0]/mbar/menu[2]/menu[2]/menu[14]/menu[4]").Select
        'SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/ctxtVBAP-ZZ_LZBKZ").Text = "000"
        'SAP_session.findById("wnd[0]").sendVKey 0
        'SAP_session.findById("wnd[0]/tbar[0]/btn[3]").press
        'SAP_session.findById("wnd[0]").sendVKey 0
        'Scrollen
        intZeileAnzeige = intZeileAnzeige + 1
        
        'Nächste Zeile aus Excel
        intZeileExcel = intZeileExcel + 1
        
        'Nächste Position in SAP
        intPos_neu = intPos_neu + 10
        
    Else
        Exit For '(wenn alle Excel-Positionen erfasst --> Abbruch)
    End If
    intZeile = intZeile + 1
Next intZeile

Application.DisplayAlerts = b_DisplayAlerts 'alten Wert wieder restaurieren

MsgBox "Alles bearbeitet. Bitte prüfen und manuell speichern.", vbOKOnly

Exit Sub

End Sub

Sub Partnerrollen_übertragen()

If Not IsObject(SAP_Application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set SAP_Application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = SAP_Application.Children(0)
End If
If Not IsObject(SAP_session) Then
   Set SAP_session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject SAP_session, "on"
   WScript.ConnectObject Application, "on"
End If

b_DisplayAlerts = Application.DisplayAlerts 'alten Wert merken
Application.DisplayAlerts = False 'zum Abschalten jeglicher Fehlermeldungen
Set xlBook = ActiveWorkbook
strNameSheetEins = "Start EDI"

intZeileBestellnummer = 3
intSpalteNummern = 6
intSpaltePartner = 10
intZeileName = 22
intZeileName2 = 23
intZeileName3 = 24
intZeileStrasse = 25
intZeileStrasse2 = 26
intZeileStrasse3 = 27
intZeileOrt = 28
intZeileRegion = 29
intZeilePLZ = 30
intZeileLandschluessel = 31
strBestellnummer = xlBook.Sheets(strNameSheetEins).Cells(intZeileBestellnummer, intSpalteNummern).Value

strName = xlBook.Sheets(strNameSheetEins).Cells(intZeileName, intSpaltePartner).Value
strName2 = xlBook.Sheets(strNameSheetEins).Cells(intZeileName2, intSpaltePartner).Value
strName3 = xlBook.Sheets(strNameSheetEins).Cells(intZeileName3, intSpaltePartner).Value
strStrasse = xlBook.Sheets(strNameSheetEins).Cells(intZeileStrasse, intSpaltePartner).Value
strStrasse2 = xlBook.Sheets(strNameSheetEins).Cells(intZeileStrasse2, intSpaltePartner).Value
strStrasse3 = xlBook.Sheets(strNameSheetEins).Cells(intZeileStrasse3, intSpaltePartner).Value
strOrt = xlBook.Sheets(strNameSheetEins).Cells(intZeileOrt, intSpaltePartner).Value
strRegion = xlBook.Sheets(strNameSheetEins).Cells(intZeileRegion, intSpaltePartner).Value
strPLZ = xlBook.Sheets(strNameSheetEins).Cells(intZeilePLZ, intSpaltePartner).Value
strLandschluessel = xlBook.Sheets(strNameSheetEins).Cells(intZeileLandschluessel, intSpaltePartner).Value

SAP_session.findById("wnd[0]").maximize
SAP_session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva02"
SAP_session.findById("wnd[0]").sendVKey 0
SAP_session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = ""
SAP_session.findById("wnd[0]/usr/txtRV45S-BSTNK").Text = strBestellnummer
SAP_session.findById("wnd[0]/usr/btnBT_SUCH").press
'Auswahl falls mehrere: ################    Abfangen falls mehrere zur Auswahl      ###########################
If SAP_session.findById("wnd[1]").Text Like " Trefferliste 1 Eintrag*" Then
    SAP_session.findById("wnd[1]/tbar[0]/btn[0]").press
Else
    MsgBox "Zur angegeben Bestellnummer wurde mehr als 1 Beleg gefunden. Programm wird beendet.", vbCritical
    Exit Sub
End If
'SAP_session.findById("wnd[1]/tbar[0]/btn[0]").press
SAP_session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[9]").Select
For n4 = 0 To 11
    If SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & n4 & "]").Text Like "*Endkunde*" Then
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1," & n4 & "]").SetFocus
        SAP_session.findById("wnd[0]").sendVKey 2
        SAP_session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/btnG_D0100_DUMMY_NAME2").press
        SAP_session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/btnG_D0100_DUMMY_TIMEZONE").press
        SAP_session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME1").Text = strName
        SAP_session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME2").Text = strName2
        SAP_session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME3").Text = strName3
        SAP_session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME4").Text = ""
        SAP_session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-STREET").Text = strStrasse
        SAP_session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-STR_SUPPL1").Text = strStrasse2
        SAP_session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-STR_SUPPL2").Text = strStrasse3
        SAP_session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-POST_CODE1").Text = strPLZ
        SAP_session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-CITY1").Text = strOrt
        SAP_session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-COUNTRY").Text = strLandschluessel
        SAP_session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-REGION").Text = strRegion
        SAP_session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-TIME_ZONE").Text = ""
        MsgBox "Endkunde übertragen. Bitte prüfen und manuell speichern.", vbOKOnly
        Exit For
    ElseIf n4 = 11 Then
        MsgBox "Partnerrolle 'Endkunde' konnte nicht gefunden werden. Programm wird beendet."
        Exit For
    End If
Next n4

Application.DisplayAlerts = b_DisplayAlerts 'alten Wert wieder restaurieren

End Sub

