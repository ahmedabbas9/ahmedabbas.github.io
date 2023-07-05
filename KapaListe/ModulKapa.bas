Attribute VB_Name = "Modul1"
Option Compare Text
Dim today As Date
Dim data_kalenderwoche
Dim gesamtlieferungen As Integer
Sub transfer_data()
    read_data
    write_data
    status_message ("")
    MsgBox "Übertragung abgeschlossen!"
End Sub
Function status_message(text As String)
    If text <> "" Then
        With ThisWorkbook.ActiveSheet.Range("D2")
            .Value = text & "..."
            .Interior.Color = RGB(255, 255, 0)
        End With
    Else
        With ThisWorkbook.ActiveSheet.Range("D2")
            .Value = ""
            .Interior.Color = xlNone
        End With
    End If
    
End Function
Function read_data() 'Auslesen der Daten aus der Excelliste

Dim excel As excel.Application
Dim wb As Workbook
Dim ws As Worksheet
Dim kw As A_Kalenderwoche
Dim vs As B_Versandstelle
Dim lf As C_Lieferung
Dim row, col, row_max As Integer

Set wb = ThisWorkbook
Set ws = wb.ActiveSheet
Set kw = New A_Kalenderwoche

row_max = 25    'maxiamle Anzahl von Zeilen des Arbeitsblattes die vom Makro beachtet werden
today = Date
gesamtlieferungen = 0

status_message ("Lese Excelliste aus")
For col = 18 To ws.UsedRange.Columns.Count
        
    'Speicher Kalendertage, die nicht in der Vergangenheit liegen
        If CDate(ws.Cells(3, col - 14).Value) >= today Then
        
             For row = 5 To row_max
                On Error Resume Next
                 
                 With ws
                    '
            
                    'Beachte Zeilen, in denen eine Ladenlistennr. und Versandstelle eingetragen ist und die Spalte "GBS-PL7" leer ist
                     'If .Cells(row, col - 6).Value <> "" And .Cells(row, col - 1).Value <> ""
                     
                     If .Cells(row, col - 12).Value <> "" And Len(.Cells(row, col - 12).Value) >= 7 And _
                       .Cells(row, col - 1).Value <> "" And _
                     Not (.Cells(row, col - 13).Value Like "Block") And _
                     Not (.Cells(row, col - 13).Value Like "verschoben") Then 'col-12= Spalte: "Ladenr." ; col-1= Spalte: "Versandstelle" ; col=Spalte: "Kommentare"
                                    
                            'Speicher Daten für neue Liefeurng
                                    Set lf = New C_Lieferung
                                    lf.verladedatum = CDate(ws.Cells(3, col - 14).Value)                    '3,col-14=Feld: Datum
                                    lf.liefernummer = .Cells(row, col - 12)                                 'col-12= Spalte: "Liefernr."
                                    
                                    If IsNumeric(lf.liefernummer) = False Then
                                        MsgBox "Bitte in LadeListe Zeile " & row & " am Verladedatum " & lf.verladedatum & " nur ein Zahl eingeben!"
                                        End
                                    End If
                                    
                                    lf.liefergruppe = .Cells(row, col - 11)                                 'col-11= Spalte: "LKW Anzahl"
                                    
                                '   If InStr(lf.liefernummer, "-") > 0 Then
                                '       lf.liefergruppe = Right(lf.liefernummer, 1) + 2
                                '        lf.liefernummer = Left(lf.liefernummer, InStr(lf.liefernummer, "-") - 1)
                                '    ElseIf InStr(lf.liefernummer, "/") > 0 Then
                                '        lf.liefergruppe = Right(lf.liefernummer, 1) + 2
                                '        lf.liefernummer = Left(lf.liefernummer, InStr(lf.liefernummer, "/") - 1)
                                '    ElseIf InStr(lf.liefernummer, "_") > 0 Then
                                '        lf.liefergruppe = Right(lf.liefernummer, 1) + 2
                                '        lf.liefernummer = Left(lf.liefernummer, InStr(lf.liefernummer, "_") - 1)
                                '    Else: lf.liefernummer = .Cells(row, col - 12)
                                '    End If
                                
                                
                                    lf.beladung = .Cells(row, col - 9)                                     'col-9= Spalte: "Land"
                                        If lf.beladung Like "Cont*" Then
                                            lf.beladung = "container"
                                            lf.ladebeginn = CDate(.Cells(row, 3))                          'col=3= Spalte: "Ladebeginn"
                                        Else
                                            lf.beladung = "lkw"
                                        End If
                                    lf.versandstelle = .Cells(row, col - 1)                                'col-1= Spalte: "Versandstelle"
                                    
                                    If IsNumeric(lf.versandstelle) = False Then
                                        MsgBox "Bitte in versandstelle Zeile " & row & " am Verladedatum " & lf.verladedatum & " nur ein Zahl eingeben!"
                                        End
                                    End If
                                    
                                    lf.cell_address = .Cells(row, col).Address                             'col= Spalte: "Kommentar"
                                    lf.cell_address2 = .Cells(row, col - 1).Address                        'col-1= Spalte: "Versandstelle"
                                   
                                    
                                    'Bestimmung der Ladedauer: Kontrolliere nachfolgende Zeile ob Land nicht leer ist und eine Versandstelle angegeben ist
                                        If .Cells(row + 1, col - 1) = "" And .Cells(row + 1, col - 9) <> "" Then   'col-1 = Spalte: "Versandstelle" ; col-9= Spalte "Land"
                                            lf.ladedauer = 60
                                        Else
                                            lf.ladedauer = 30
                                        End If

                            'Überprüfe ob Versandstelle aus vorherigen Lieferungen angelegt ist
                                    If kw.logistik.Count <> 0 Then
                                            Dim entry As Boolean
                                            entry = False
                                            For i = 1 To kw.logistik.Count                             'Überprüfe die Versandstelle, ob als Eintrag in der Collection vorhanden ist
                                                If kw.logistik(i).gruppe = lf.versandstelle Then
                                                    kw.logistik(i).lieferungen.Add lf                               'Füge neue Lieferung in vorhandener Versandstelle ein
                                                    entry = True
                                                    Exit For
                                                End If
                                            Next i
                                            
                                            If entry = False Then GoTo Sprung
                                    'Lege neue Versandstelle an und speicher in dieser die neue Lieferung
                                    Else
Sprung:
                                            Set vs = New B_Versandstelle
                                            vs.lieferungen.Add lf                                             'Füge neuen Liefertag in neuer Versandstelle zu
                                            vs.gruppe = lf.versandstelle
                                            kw.logistik.Add vs  'Lege Datensatz als neuen Eintrag in Kalenderwoche ab
                                    End If
                                    
                    End If
                     
                     
                 End With
                 
            
             Next row

        
        End If
             
        col = col + 14                'col+5= Spalte: Versandstelle (Spalte: "GBS-PL7" -1, da mit "Next col" Sprung auf Spalte: "GBS-PL7")
Next col

'Gebe Error Message zurück, wenn kein Eintrag zum Übertragen gefunden
If kw.logistik.Count = 0 Then
    MsgBox "Keine Daten zum Übertragen gefunden!"
    status_message ("")
    End
    Exit Function
End If

Debug.Print "Function read_data"
Debug.Print "------------------"
Debug.Print "Heutiges Datum: " & today
For Each logistik In kw.logistik
    Debug.Print "Versandstelle: " & logistik.gruppe
        For Each lieferung In logistik.lieferungen
            gesamtlieferungen = gesamtlieferungen + 1
            Debug.Print "       Verladungsdatum: " & lieferung.verladedatum
            Debug.Print "       Ladelistennr.: " & lieferung.liefernummer
            Debug.Print "       Ladelistegruppe.: " & lieferung.liefergruppe
            Debug.Print "       Beladung: " & lieferung.beladung
            Debug.Print "       Ladedauer: " & lieferung.ladedauer
            Debug.Print "       Ladebeginn: " & lieferung.ladebeginn
            Debug.Print "       Excel-Zelle :" & lieferung.cell_address
            Debug.Print "       Excel-Zelle2 :" & lieferung.cell_address2
            
        Next lieferung
Next logistik
Debug.Print vbCrLf & "Gesamtlieferungen: " & gesamtlieferungen & vbCrLf


Set data_kalenderwoche = kw

Set lf = Nothing
Set vs = Nothing
Set kw = Nothing
Set ws = Nothing
Set wb = Nothing

End Function

Function write_data() 'Übertragen der Daten in SAP und Excellliste

Debug.Print "Function read_data"
Debug.Print "------------------"

copydata_to_SAP
'add_SAP_message      'deaktivieren mit kommentar Funktion ('),,,, aktivieren (') löschen
    
Set data_kalenderwoche = Nothing
status_message ("Datenübertragung abgeschlossen")
Application.Wait (Now + TimeValue("00:00:02"))
End Function

Function copydata_to_SAP() 'Übertragen der ausgelesenen Daten in SAP

Dim sp As SAP
Dim write_data As Boolean
Dim progress As Integer

On Error GoTo Error_registration

Set sp = New SAP
Set session = sp.session
write_data = True
progress = 0


status_message ("Starte SAP")
session.findById("wnd[0]").maximize
session.findById("wnd[0]").iconify
session.findById("wnd[0]/tbar[0]/okcd").text = "/n/SIE/EVN_ZVHB_LEKAPA"
session.findById("wnd[0]").sendVKey 0

On Error GoTo Error_transaction

session.findById("wnd[0]/usr/ctxtS_DALBG-LOW").text = Date
                    
                    
For Each logistik In data_kalenderwoche.logistik
        
        If logistik.gruppe <> 1103 Then '+++Löschen!!! Aus Testzwecken eingefügt, da keine Berechtigung für Versandstelle 1103
        
        session.findById("wnd[0]/usr/ctxtP_VSTEL").text = logistik.gruppe
        session.findById("wnd[0]/usr/ctxtP_VSTEL").SetFocus
        session.findById("wnd[0]/usr/ctxtP_VSTEL").caretPosition = 4
        session.findById("wnd[0]/tbar[1]/btn[8]").press

On Error GoTo Error_dilevery


        For Each lieferung In logistik.lieferungen
                
                    With lieferung
                    
                                progress = progress + 1
                                status_message ("Schreibe in SAP " & progress & "/" & gesamtlieferungen)
                                session.findById("wnd[0]/tbar[1]/btn[18]").press
                                session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[3,21]").text = .verladedatum
                                session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = .liefernummer
                                ' wenn liefernummer schon verhanden dann im gegebenen Datum auf verschieben klicken
                                ' Ladeliste 2702286 ist bereits für den 12.12.2022 eingeplant!
                                If (.liefergruppe <> "") Then
                                    session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[1,21]").SetFocus
                                    session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[1,21]").caretPosition = 0
                                    session.findById("wnd[1]").sendVKey 4
                                    session.findById("wnd[2]/usr/lbl[1," & (.liefergruppe + 2) & "]").SetFocus
                                    session.findById("wnd[2]/usr/lbl[1," & (.liefergruppe + 2) & "]").caretPosition = 1
                                    session.findById("wnd[2]").sendVKey 2
                                End If
                                If .beladung = "container" Then session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[2,21]").text = "C"
On Error GoTo Error_fillout
                                session.findById("wnd[1]").sendVKey 0
                                While InStr(1, session.ActiveWindow.text, "Information") <> 0
                                        write_data = False
                                        'Speicher Anzeigetext der SAP-Rückmeldung
                                        For Each obj In session.findById("wnd[1]/usr/").Children
                                                .sap_meldung = .sap_meldung & obj.text & " "
                                        Next
                                        sap_meldung = .sap_meldung & vbCrLf
                                        session.findById("wnd[1]").sendVKey 0
                                Wend
                                
                                If write_data = False Then
                                    session.findById("wnd[1]/tbar[0]/btn[12]").press
                                    write_data = True
                                    GoTo Next_Lieferung
                                End If
                                
                                If .beladung = "container" Then
                                    session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[7,21]").text = "C40"
                                    session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[4,21]").text = .ladebeginn
                                Else
                                    session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[7,21]").text = "T14"
                                End If
                                
                                session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[5,21]").text = .ladedauer
                                session.findById("wnd[1]").sendVKey 0
                                
                    End With
Next_Lieferung:
        Next lieferung

        session.findById("wnd[0]/tbar[0]/btn[3]").press
        
        End If '+++Löschen
        
Next logistik

session.StartTransaction ("SE38")
If sp.connection.sessions.Count > 1 Then session.findById("wnd[0]").Close

Set session = Nothing
Set sp = Nothing

Exit Function

Error_registration:
status_message ("")
Set session = Nothing
Set sp = Nothing
End

Error_transaction:
MsgBox "SAP: Keine Berechtigung für Zugriff auf Transaktion ""SIE/EVN_ZVHB_LEKAPA"" "
status_message ("")
Set session = Nothing
Set sp = Nothing
End

Error_dilevery:
MsgBox "SAP: Fehler beim Laden der Kalenderwoche!"
status_message ("")
Set session = Nothing
Set sp = Nothing
End

Error_fillout:
MsgBox "SAP: Fehler bei Datenübertragung in der Ladeliste!"
status_message ("")
Set session = Nothing
Set sp = Nothing
End
End Function

Function add_SAP_message()

Dim wb As Workbook
Dim ws As Worksheet

Set wb = ThisWorkbook
Set ws = wb.ActiveSheet

On Error GoTo Error_addresponse
For Each logistik In data_kalenderwoche.logistik
        If logistik.gruppe <> 1103 Then
        status_message ("Schreibe zurück in Excelliste")
        For Each lieferung In logistik.lieferungen
                
                With ws.Range(lieferung.cell_address)
                
                    If lieferung.sap_meldung = "" Then
                            If .CommentThreaded Is Nothing = False Then .CommentThreaded.Delete
                            .Interior.Color = RGB(34, 139, 34)
                            'ws.Range(lieferung.cell_address2).Value = "OK"
                    Else
                            .Interior.Color = RGB(255, 255, 255)
                            If .CommentThreaded Is Nothing Then
                                .AddCommentThreaded ("SAP-Meldung:" & vbCrLf & vbCrLf & lieferung.sap_meldung)
                            Else
                                .CommentThreaded.AddReply ("SAP-Meldung:" & vbCrLf & vbCrLf & lieferung.sap_meldung)
                            End If
                    End If
                
                End With
    
        Next lieferung
        End If
Next logistik

Set ws = Nothing
Set wb = Nothing

Exit Function

Error_addresponse:
MsgBox "Excel: Fehler beim Zurückschreiben in die Excel-Liste!"
status_message ("")
Set ws = Nothing
Set wb = Nothing
End

End Function









