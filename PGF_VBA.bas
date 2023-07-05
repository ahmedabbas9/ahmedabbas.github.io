Attribute VB_Name = "Modul1"
Dim wb As Workbook

Dim intSpalteBeleg As Integer
Dim intZeileBeleg As Integer
Dim intZeileTeil As Integer
Dim intSpalteTeil As Integer
Dim intSpalteProjekt As Integer
Dim intZeileProjekt As Integer
Dim intSpalteNummer As Integer
Dim intZeileNummer As Integer
Dim intSpaltePGF As Integer
Dim intZeilePGF As Integer
Dim intSpaltePGFDatum As Integer
Dim intZeilePGFDatum As Integer
Dim intSpalteP0 As Integer
Dim intZeileP0 As Integer
Dim intSpalteP0K As Integer      ' P0-KdA
Dim intZeileP0K As Integer
Dim intSpalteP1 As Integer
Dim intZeileP1 As Integer
Dim intSpalteP0Datum As Integer
Dim intZeileP0Datum As Integer
Dim intSpalteIndexName As Integer
Dim intSpalteBaswert As Integer
Dim intSpalteVarwert As Integer
Dim intSpalteAbrechnung As Integer
Dim intSpalteDelta As Integer
Dim intZeileDelta As Integer
Dim intSpalteTDelta As Integer
Dim intZeileTDelta As Integer
Dim intSpalteFix As Integer
Dim intZeileFix As Integer




Dim strProjekt() As String        ' ProjektName
Dim arrBeleg() As Variant        'Vertriebsbelege
Dim arrTeil() As Variant          ' Teillieferung
Dim strPGF() As String            'Formel
Dim strNummer() As String          'P-Nummer
Dim strP0() As String               'P0
Dim strP0Datum() As String        ' Basismonat / Basisjahr
Dim arrIndex() As String
Dim arrIndexName() As String      'index Kürzung
Dim strPGFDatum() As String       ' PGF Simulation
Dim arrBaswert() As Double      ' Basiswert aus Index & Basismonat/basisjahr
Dim arrVarwert() As Double       'Variablerwert aus Index & PGF Simulation
Dim arrAnteil() As Double
Dim arrAbrechnung() As Double
Dim arrDelta() As Double
Dim dbDelta() As Double
Dim dbP1() As Double            'P1
Dim TDelta() As Double
Dim strFix() As Double
Dim arrInfo() As Variant
Dim sum As Double


Sub daten_auslesen()
On Error GoTo ErrorHandler
On Error Resume Next
Set wb = ActiveWorkbook

Dim letzteZeile As Long
Dim strHilfsBeleg As String
Dim strHilfsDatum As String
Dim strHilfsBasDatum As String
Dim strHilfsPGFDatum As String
Dim strHilfsIndex As String
Dim j As Integer
Dim i As Long
Dim z As Long
Dim x As Integer

    
intSpalteBeleg = 4
intZeileBeleg = 13
intSpalteProjekt = 3
intZeileProjekt = 1
intSpalteNummer = 4
intZeileNummer = 12
intSpaltePGF = 3
intZeilePGF = 4
intSpaltePGFDatum = 8
intZeilePGFDatum = 19
intSpalteP0 = 4
intZeileP0 = 6
intSpalteP0K = 4
intZeileP0K = 6
intSpalteP0Datum = 6
intZeileP0Datum = 19
intSpalteIndex = 2
intZeileIndex = 19
intSpalteIndexName = 3
intZeileIndexName = 19
intSpalteBaswert = 5
intZeileBaswert = 19
intSpalteVarwert = 7
intZeileVarwert = 19
intSpalteAnteil = 4
intZeileAnteil = 19
intSpalteAbrechnung = 9
intZeileAbrechnung = 19
intSpalteDelta = 11
intZeileDelta = 19
intSpalteP0K = 4
intZeileP0K = 14
intSpalteP1 = 4
intZeileP1 = 7
intSpalteTDelta = 4
intZeileTDelta = 8
intZeileFix = 18
intSpalteFix = 4
intSpalteTeil = 5



letzteZeile = wb.Sheets("start").Cells(Rows.Count, "D").End(xlUp).Row
ReDim arrBeleg(0 To (letzteZeile - 7))
ReDim arrTeil(0 To (letzteZeile - 7))
ReDim strProjekt(0 To UBound(arrBeleg) - 1)
ReDim strNummer(0 To UBound(arrBeleg) - 1)
ReDim strPGF(0 To UBound(arrBeleg) - 1)
ReDim strPGFDatum(0 To UBound(arrBeleg) - 1)
ReDim strP0(0 To UBound(arrBeleg) - 1)
ReDim strP0Datum(0 To UBound(arrBeleg) - 1)
ReDim strFix(0 To UBound(arrBeleg) - 1)

x = 0
'die gesuchte Vertriebsbeleg

For j = 8 To letzteZeile
    arrBeleg(x) = wb.Sheets("start").Cells(j, intSpalteBeleg).Value
    arrTeil(x) = wb.Sheets("start").Cells(j, intSpalteTeil).Value
    x = x + 1
Next j


letzteZeile = wb.Sheets("PGF Controlling View").Cells(Rows.Count, "B").End(xlUp).Row  ' letzte belegte Zeile

' Daten aus PGF Controlling View
For x = 0 To UBound(arrBeleg) - 1
    For i = 3 To letzteZeile
        strHilfsBeleg = wb.Sheets("PGF Controlling View").Cells(i, 2).Value
        strHilfsTeil = wb.Sheets("PGF Controlling View").Cells(i, 18).Value
        If arrBeleg(x) = strHilfsBeleg And Replace(arrTeil(x), " ", "") = Replace(strHilfsTeil, " ", "") Then
            strProjekt(x) = wb.Sheets("PGF Controlling View").Cells(i, 4).Value
            strNummer(x) = wb.Sheets("PGF Controlling View").Cells(i, 3).Value
            strPGFDatum(x) = wb.Sheets("PGF Controlling View").Cells(i, 12).Value
            strP0(x) = wb.Sheets("PGF Controlling View").Cells(i, 9).Value
        End If
    Next i
Next x

'arrays für index und Anteil
ReDim arrIndex(0 To UBound(arrBeleg), 0 To 4)
ReDim arrAnteil(0 To UBound(arrBeleg), 0 To 4)
j = 0
x = 0
letzteZeile = wb.Sheets("Projekt-Stammdaten").Cells(Rows.Count, "A").End(xlUp).Row  ' letzte belegte Zeile
' Daten aus Projekt-Stammdaten
For x = 0 To UBound(arrBeleg) - 1
    For i = 4 To letzteZeile
        strHilfsBeleg = wb.Sheets("Projekt-Stammdaten").Cells(i, 1).Value
        If arrBeleg(x) = strHilfsBeleg Then
            strPGF(x) = wb.Sheets("Projekt-Stammdaten").Cells(i, 10).Value
            strP0Datum(x) = wb.Sheets("Projekt-Stammdaten").Cells(i, 11).Value
            strFix(x) = wb.Sheets("Projekt-Stammdaten").Cells(i, 13).Value
            For y = 14 To 20
                For j = 0 To 4
                    arrIndex(x, j) = wb.Sheets("Projekt-Stammdaten").Cells(i, y).Value
                    arrAnteil(x, j) = wb.Sheets("Projekt-Stammdaten").Cells(i, y + 1).Value
                    y = y + 2
                Next j
            Next y
          
        Exit For 'mehrere Daten
        End If
    Next i
Next x

'arrays für Basis- und Variablewerte
ReDim arrBaswert(0 To UBound(arrBeleg), 0 To 4)
ReDim arrVarwert(0 To UBound(arrBeleg), 0 To 4)
letzteZeile = wb.Sheets("Indize Werte").Cells(Rows.Count, "A").End(xlUp).Row  ' letzte belegte Zeile
letzteSpalte = wb.Sheets("Indize Werte").Cells(4, Columns.Count).End(xlToLeft).Column

' Daten aus Indize Werte
For x = 0 To UBound(arrBeleg)
    For i = 5 To letzteZeile
        strHilfsIndex = wb.Sheets("Indize Werte").Cells(i, 1).Value
        For j = 0 To 4
            If arrIndex(x, j) = strHilfsIndex Then
                For z = 1 To letzteSpalte
                    strHilfsBasDatum = wb.Sheets("Indize Werte").Cells(4, z).Value
                    strHilfsPGFDatum = wb.Sheets("Indize Werte").Cells(4, z).Value
                    If strHilfsBasDatum = strP0Datum(x) Then
                      arrBaswert(x, j) = wb.Sheets("Indize Werte").Cells(i, z).Value
                    End If
                    If strHilfsPGFDatum = strPGFDatum(x) Then
                      arrVarwert(x, j) = wb.Sheets("Indize Werte").Cells(i, z).Value
                    End If
                Next z
            End If
        Next j
    Next i
Next x


ReDim arrIndexName(0 To UBound(arrBeleg), 0 To 4)
letzteZeile = wb.Sheets("Indize-Stammdaten").Cells(Rows.Count, "B").End(xlUp).Row  ' letzte belegte Zeile

' Daten aus Indize-Stammdaten
For i = 4 To letzteZeile
    strHilfsIndex = wb.Sheets("Indize-Stammdaten").Cells(i, 2).Value
    For x = 0 To UBound(arrBeleg)
        For j = 0 To 4
            If arrIndex(x, j) = strHilfsIndex Then
                arrIndexName(x, j) = wb.Sheets("Indize-Stammdaten").Cells(i, 3).Value
        End If
        Next j
    Next x
Next i

'Anteile, Basis-und Variablewerte leere Werte mit nullen ersetzten
For x = 0 To UBound(arrBeleg)
    For j = 0 To 4
        If IsEmpty(arrAnteil(x, j)) Then arrAnteil(x) = "0,00"
        If IsEmpty(arrBaswert(x, j)) Then arrBaswert(x) = "0,00"
        If IsEmpty(arrVarwert(x, j)) Then arrVarwert(x) = "0,00"
    Next j
Next x


' Abrechnungen rechnen
ReDim arrAbrechnung(0 To UBound(arrBeleg), 0 To 4)
For x = 0 To UBound(arrBeleg)
    For j = 0 To 4
        If IsEmpty(arrIndex(x, j)) = False Then
            If arrBaswert(x, j) = 0 Or arrAnteil(x, j) = 0 Then  ' falls eine leer ist, dann wird es sich ein überlauf in der Formel unten geben
                arrAbrechnung(x, j) = 0
            Else
                arrAbrechnung(x, j) = (CDbl(arrVarwert(x, j)) / CDbl(arrBaswert(x, j)) * CDbl(arrAnteil(x, j)))
            End If
        End If
    Next j
Next x


'Delta rechnen
ReDim arrDelta(0 To UBound(arrBeleg), 0 To 4)
For x = 0 To UBound(arrBeleg) - 1
    For j = 0 To 4
            If IsEmpty(arrIndex(x, j)) = False Then
                If strP0(x) = 0 Then                 'falls P0 null ist, wird es sich ein Überlauf geben
                    arrDelta(x, j) = 0
                Else
                    arrDelta(x, j) = Round((arrAbrechnung(x, j) - arrAnteil(x, j)) / 100 * strP0(x), 2)
                End If
            End If
    Next j
Next x


ReDim TDelta(0 To UBound(arrBeleg))
ReDim dbP1(0 To UBound(arrBeleg))

sum = 0
For x = 0 To UBound(arrBeleg) - 1
    For j = 0 To 4
        sum = sum + arrAbrechnung(x, j)
        'P1 rechnen
        dbP1(x) = (strP0(x)) * (sum + CDbl(strFix(x)))
        
        'Delta in Reihe 8, Spalte D
        TDelta(x) = dbP1(x) - strP0(x)
        
        
    Next j
    sum = 0
Next x

Dim gefunden As Boolean
Dim invalidChars As String

invalidChars = "/\*[]{}!@#$%"

' ins Makro schreiben
For x = 0 To UBound(arrBeleg) - 1
    If x = 0 Then
        wb.Sheets("Makro").Select
        wb.Sheets("Makro").Copy After:=wb.Sheets("Makro")
        'check if the woksheet Name schon existiert
        sheetName = strProjekt(x)
        For Each ws In wb.Worksheets
            If ws.name = sheetName Then
                gefunden = True
                Exit For
            End If
        Next ws
        If gefunden = True Then
            MsgBox ("Der Arbeitsblatt (" & strProjekt(x) & ") ist schon Vorhanden")
            Exit Sub
        Else
            For i = 1 To Len(invalidChars)
            strProjekt(x) = Replace(strProjekt(x), Mid(invalidChars, i, 1), "")
            Next i
        If arrTeil(x) <> "" Then
            wb.ActiveSheet.name = strProjekt(x) & arrTeil(x)
        Else
            wb.ActiveSheet.name = strProjekt(x)
        End If
        
        End If
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileBeleg, intSpalteBeleg).Value = arrBeleg(x)             'strName hat sich geändert
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileProjekt, intSpalteProjekt).Value = strProjekt(x)
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeilePGF, intSpaltePGF).Value = strPGF(x)
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileNummer, intSpalteNummer).Value = strNummer(x)
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileP0, intSpalteP0).NumberFormat = "#,##0.00"
        'wb.Sheets(strProjekt(x)).Cells(intZeileP0, intSpalteP0).Value = CDbl(strP0(x))
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileP1, intSpalteP1).NumberFormat = "#,##0.00"
        'wb.Sheets(strProjekt(x)).Cells(intZeileP1, intSpalteP1).Value = CDbl(dbP1(x))
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileP0K, intSpalteP0K).NumberFormat = "#,##0.00"
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileP0K, intSpalteP0K).Value = CDbl(strP0(x))
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileTDelta, intSpalteTDelta).NumberFormat = "#,##0.00"
        'wb.Sheets(strProjekt(x)).Cells(intZeileTDelta, intSpalteTDelta).Value = CDbl(TDelta(x))
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileFix, intSpalteFix).Value = CDbl(strFix(x))
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileFix, intSpalteAbrechnung).Value = CDbl(strFix(x))
            
        
    Else
        wb.Sheets("Makro").Select
        wb.Sheets("Makro").Copy After:=wb.Sheets("Makro")
        For i = 1 To Len(invalidChars)
            strProjekt(x) = Replace(strProjekt(x), Mid(invalidChars, i, 1), "")
        Next i
        If arrTeil(x) <> "" Then
            wb.ActiveSheet.name = strProjekt(x) & arrTeil(x)
        Else
            wb.ActiveSheet.name = strProjekt(x)
        End If
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileBeleg, intSpalteBeleg).Value = arrBeleg(x)
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileProjekt, intSpalteProjekt).Value = strProjekt(x)
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeilePGF, intSpaltePGF).Value = strPGF(x)
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileNummer, intSpalteNummer).Value = strNummer(x)
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileP0, intSpalteP0).NumberFormat = "#,##0.00"
        'wb.Sheets(strProjekt(x)).Cells(intZeileP0, intSpalteP0).Value = CDbl(strP0(x))
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileP1, intSpalteP1).NumberFormat = "#,##0.00"
        'wb.Sheets(strProjekt(x)).Cells(intZeileP1, intSpalteP1).Value = CDbl(dbP1(x))
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileP0K, intSpalteP0K).NumberFormat = "#,##0.00"
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileP0K, intSpalteP0K).Value = CDbl(strP0(x))
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileTDelta, intSpalteTDelta).NumberFormat = "#,##0.00"
        'wb.Sheets(strProjekt(x)).Cells(intZeileTDelta, intSpalteTDelta).Value = CDbl(TDelta(x))
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileFix, intSpalteFix).Value = CDbl(strFix(x))
        wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(intZeileFix, intSpalteAbrechnung).Value = CDbl(strFix(x))
    End If
Next x

For x = 0 To UBound(arrBeleg) - 1
    For i = 19 To 23
        For y = 0 To 4
            If x = 0 Then
                If arrIndex(x, y) <> "0" Then
                    wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(i, intSpalteIndex).Value = arrIndex(x, y)
                    wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(i, intSpalteIndexName).Value = arrIndexName(x, y)
                    wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(i, intSpalteP0Datum).Value = DateValue(strP0Datum(x))
                    wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(i, intSpaltePGFDatum).Value = DateValue(strPGFDatum(x))
                    wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(i, intSpalteBaswert).NumberFormat = "#,##0.00"
                    'wb.Sheets(strProjekt(x)).Cells(i, intSpalteBaswert).Value = CDbl(arrBaswert(x, y))
                    wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(i, intSpalteVarwert).NumberFormat = "#,##0.00"
                    'wb.Sheets(strProjekt(x)).Cells(i, intSpalteVarwert).Value = CDbl(arrVarwert(x, y))
                    wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(i, intSpalteAnteil).Value = arrAnteil(x, y)
                    wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(i, intSpalteAbrechnung).NumberFormat = "#,##0.00"
                    'wb.Sheets(strProjekt(x)).Cells(i, intSpalteAbrechnung).Value = CDbl(arrAbrechnung(x, y))
                    wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(i, intSpalteDelta).NumberFormat = "#,##0.00"
                    'wb.Sheets(strProjekt(x)).Cells(i, intSpalteDelta).Value = CDbl(arrDelta(x, y))
                End If
                i = i + 1
            Else
                If arrIndex(x, y) <> "0" Then
                    wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(i, intSpalteIndex).Value = arrIndex(x, y)
                    wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(i, intSpalteIndexName).Value = arrIndexName(x, y)
                    wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(i, intSpalteP0Datum).Value = DateValue(strP0Datum(x))
                    wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(i, intSpaltePGFDatum).Value = DateValue(strPGFDatum(x))
                    wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(i, intSpalteBaswert).NumberFormat = "#,##0.00"
                    'wb.Sheets(strProjekt(x)).Cells(i, intSpalteBaswert).Value = CDbl(arrBaswert(x, y))
                    wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(i, intSpalteVarwert).NumberFormat = "#,##0.00"
                    'wb.Sheets(strProjekt(x)).Cells(i, intSpalteVarwert).Value = CDbl(arrVarwert(x, y))
                    If arrAnteil(x, y) <> "0" Then
                    wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(i, intSpalteAnteil).Value = arrAnteil(x, y)
                    End If
                    wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(i, intSpalteAbrechnung).NumberFormat = "#,##0.00"
                    'wb.Sheets(strProjekt(x)).Cells(i, intSpalteAbrechnung).Value = CDbl(arrAbrechnung(x, y))
                    wb.Sheets(strProjekt(x) & arrTeil(x)).Cells(i, intSpalteDelta).NumberFormat = "#,##0.00"
                    'wb.Sheets(strProjekt(x)).Cells(i, intSpalteDelta).Value = CDbl(arrDelta(x, y))
                End If
                i = i + 1
            End If
        Next y
    Next i
Next x

MsgBox "Die Bearbeitung ist abgeschlossen"
Exit Sub
ErrorHandler:
    MsgBox "Es gibt einen Fehler beim Beleg" & arrBeleg(x) & " : " & Err.Description

End Sub
Function IsValidWorksheetName(name As String) As Boolean
    Dim pattern As String
    
    ' Define a regular expression pattern to match allowed characters
    pattern = "^[a-zA-Z0-9_]+$"
    
    ' Validate the worksheet name using the pattern
    IsValidWorksheetName = RegExTest(pattern, name)
End Function

Function RegExTest(pattern As String, inputString As String) As Boolean
    Dim regex As Object
    
    ' Create a regular expression object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Set the pattern and ignore case
    With regex
        .pattern = pattern
        .IgnoreCase = True
    End With
    
    ' Test if the input string matches the pattern
    RegExTest = regex.Test(inputString)
End Function
