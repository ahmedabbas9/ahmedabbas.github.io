Attribute VB_Name = "Modul1"
Dim strNameSheetEins As String


Sub fakturaplan_auslesen()


On Error GoTo ErrorHandler

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

Dim strNummer As String
Dim xlBook As Workbook
Dim intLetzteZeile As Integer
Dim intZeileZeitEnde As Integer
Dim intSpalteZeitEnde As Integer
Dim intSpalteAuftrag As Integer
Dim ErstellDatum As Date
Dim strBezeichnung As String
Dim strProzent As String
Dim strWert As Double
Dim strRegel As String
Dim strTyp As String
Dim strArt As String


Set xlBook = ActiveWorkbook


Application.DisplayAlerts = False 'zum Abschalten jeglicher Fehlermeldungen
strNameSheetEins = "Start"

intZeileZeitStart = 6
intZeileZeitEnde = 7
intZeileNummer = 4


intSpalteZeit = 6
intSpalteNummer = 11
intSpalteErstellDatum = 2
intSpalteBezeichnung = 3
intSpalteProzent = 4
intSpalteWert = 5
intSpalteRegel = 6
intSpalteTyp = 7
intSpalteArt = 8
strStatus = "Fehler"


'datStart = Format(Now(), "hh:mm:ss")
'xlBook.Sheets(strNameSheetEins).Cells(intZeileZeitStart, intSpalteZeit).Value = datStart
'bolErrorFlag = False

strNummer = xlBook.Sheets(strNameSheetEins).Cells(intZeileNummer, intSpalteNummer).Value


SAP_session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
SAP_session.findById("wnd[0]").sendVKey 0
SAP_session.findById("wnd[0]/tbar[0]/okcd").Text = "va02"
SAP_session.findById("wnd[0]").sendVKey 0
SAP_session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = strNummer
SAP_session.findById("wnd[0]").sendVKey 0

If SAP_session.ActiveWindow.Name = "wnd[1]" Then
    If SAP_session.findById("wnd[1]").Text Like "Inform*" Then SAP_session.findById("wnd[1]/tbar[0]/btn[0]").press
End If

SAP_session.findById("wnd[0]").maximize
SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-ARKTX[4,1]").SetFocus
SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-ARKTX[4,1]").caretPosition = 13
SAP_session.findById("wnd[0]").sendVKey 2
SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05").Select
'SAP_session.FindById("wnd[0]/tbar[0]/btn[83]").press
b = 0
Max = SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-AFDAT[0,0]").Parent.verticalScrollbar.Maximum

SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA").verticalScrollbar.Position = Max - 1
On Error Resume Next
y = 15
i = 1
'Schleife zum Durchlaufen der Auftragsliste
Do Until xlBook.Sheets(strNameSheetEins).Cells(y, intSpalteErstellDatum).Value = ""
    If IsDate(xlBook.Sheets(strNameSheetEins).Cells(y, intSpalteErstellDatum).Value) Then

        ErstellDatum = xlBook.Sheets(strNameSheetEins).Cells(y, intSpalteErstellDatum).Value
        strBezeichnung = xlBook.Sheets(strNameSheetEins).Cells(y, intSpalteBezeichnung).Value
        strProzent = xlBook.Sheets(strNameSheetEins).Cells(y, intSpalteProzent).Value
        strWert = xlBook.Sheets(strNameSheetEins).Cells(y, intSpalteWert).Value
        strRegel = xlBook.Sheets(strNameSheetEins).Cells(y, intSpalteRegel).Value
        strTyp = xlBook.Sheets(strNameSheetEins).Cells(y, intSpalteTyp).Value
        strArt = xlBook.Sheets(strNameSheetEins).Cells(y, intSpalteArt).Value
        
    
        
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-AFDAT[0," & i & "]").Text = ErstellDatum
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-TETXT[1," & i & "]").Text = strBezeichnung
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/txtFPLT-FAKWR[5," & i & "]").Text = strWert
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FAREG[9," & i & "]").Text = strRegel
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/txtFPLT-FAKWR[5," & i & "]").SetFocus
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/txtFPLT-FAKWR[5," & i & "]").caretPosition = 12
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FPTTP[12," & i & "]").Text = strTyp
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FKARV[13," & i & "]").Text = strArt
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FAKSP[7," & i & "]").SetFocus
        SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FAKSP[7," & i & "]").caretPosition = 2
        SAP_session.findById("wnd[0]").sendVKey 0
        
        If SAP_session.ActiveWindow.Name = "wnd[1]" Then
            If SAP_session.findById("wnd[1]").Text Like "Inform*" Then SAP_session.findById("wnd[1]/tbar[0]/btn[0]").press
        End If
    End If
   ' i = i + 1
    y = y + 1
   pos = SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA").verticalScrollbar.Position
    pos = pos + 1
    SAP_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\05/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA").verticalScrollbar.Position = pos

Loop
  
MsgBox ("Alle Daten wurden in SAP eingetragen")

Exit Sub
ErrorHandler:
    MsgBox "Ein Fehler ist passiert: " & Err.Description
  
End Sub


