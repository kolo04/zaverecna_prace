Attribute VB_Name = "Module3"
Dim ws As Worksheet
Dim wsInput As Worksheet
Dim numOfCriteria As Integer
Dim numOfCandidates As Integer
Dim Shape As Shape

' Makro pro výpoèet metody WSA
Sub M3_metoda_WSA()
    
    ' Zobrazení výsledkù procedury až po kompletním naètení procedury
    Application.ScreenUpdating = False
    
    ' Volání kontroly vyplnìných hodnot
    Call CheckFilledCells

' Ovìøení existence listu "Metoda WSA"
    Dim wsExists As Boolean
    wsExists = False
    
    For Each ws In ThisWorkbook.Sheets
        If ws.name = "Metoda WSA" Then
            wsExists = True
            ' Pøesun na list a jeho vyèištìní
            ws.Activate
            ws.Unprotect "1234"
            ws.Cells.Clear
            
            For Each Shape In ws.Shapes
                Shape.Delete
            Next Shape
            Exit For
        End If
    Next ws
    
    ' Vytvoøení listu, pokud ještì neexistuje
    If Not wsExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = "Metoda WSA"
        ' Pøesun na novì vytvoøený list
        ws.Activate
        ws.Unprotect "1234"
    End If
    
' Práce s "Metoda WSA"
    ' Nahrání vstupních dat
    Dim lastRow As Long
    Dim i As Long
    Dim maxVal As Double
    
    ' Definice pracovních listù
    Set wsInput = ThisWorkbook.Sheets("Vstupní data")
    Set ws = ThisWorkbook.Sheets("Metoda WSA")
    
    numOfCriteria = wsInput.Range("C2").value
    numOfCandidates = wsInput.Range("F2").value
    
    With ws
        ' Kopírování záhlaví a kritérií (pøidání apostrofu pro pøevod na text)
        For i = 3 To 4 + numOfCriteria
            .Cells(i, 2).value = "'" & wsInput.Cells(i, 2).value
            .Cells(i, 3).value = wsInput.Cells(i, 3).value
        Next i
        
        ' Kopírování záhlaví a variant (pøidání apostrofu pro pøevod na text)
        For i = 4 To 4 + numOfCandidates
            .Cells(3, i).value = wsInput.Cells(3, i).value
            .Cells(4, i).value = "'" & wsInput.Cells(4, i).value
        Next i
    
        ' Kopírování èíselných hodnot bez zmìny formátu
        .Range(.Cells(5, 4), .Cells(4 + numOfCriteria, 4 + numOfCandidates)).value = _
            wsInput.Range(wsInput.Cells(5, 4), wsInput.Cells(4 + numOfCriteria, 4 + numOfCandidates)).value
        
        ' Formátování záhlaví B3 až sloupec 4 + numOfCandidates na øádku 3
        With .Range(.Cells(3, 2), .Cells(4, 4 + numOfCandidates))
            ' Tuènì a zarovnání na støed
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        
            ' Nastavení ohranièení
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
    
        ' Zarovnání bunìk cíle na støed
        .Range(.Cells(4, 3), .Cells(4 + numOfCriteria, 3)).HorizontalAlignment = xlCenter
        
        ' Nastavení stylu bunìk váhy jako na formát "0.0 %"
        .Range(.Cells(4, 4), .Cells(4 + numOfCriteria, 4)).NumberFormat = "0.0 %"
    End With

    
    ' Nastavení ohranièení pro sloupce B až D v øádcích 4 až poslední kritérium
    Dim column As Range
    Dim columnRange As Range
    
    ' Nastavení rozsahu sloupcù B až D
    Set columnRange = ws.Range(ws.Cells(4, 2), ws.Cells(4 + numOfCriteria, 4))
    
    ' Pro každý sloupec v rozsahu
    For Each column In columnRange.Columns
        ' Nastavení ohranièení pro pravý okraj
        With column.Columns(column.Columns.Count).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    Next column
    
    ' Úprava minimalizaèních kritérií a èíselný formát
    For i = 5 To (4 + numOfCriteria)
        ' Pokud je ve sloupci C hodnota "min"
        If ws.Cells(i, 3).value = "min" Then
            ' Najdi maximální hodnotu v celém øádku
            maxVal = Application.WorksheetFunction.Max(wsInput.Range(wsInput.Cells(i, 4), wsInput.Cells(i, 4 + numOfCandidates)))
            ' Pro všechny sloupce od E do poslední varianty
            For j = 5 To (5 + numOfCandidates - 1)
                ' Pøepoèet hodnoty na maximální hodnotu podle maximální hodnoty v øádku
                ws.Cells(i, j).value = maxVal - wsInput.Cells(i, j).value
            Next j
        End If
    Next i
    
    ' Úprava šíøky sloupcù (Autofit na minimálnì 80px)
    AdjustColumnWidth ws, ws.Range(ws.Columns(2), ws.Columns(4 + numOfCandidates))
    ws.Columns(5 + numOfCandidates).ColumnWidth = 10.33
    
    With ws
        ' Popisek tabulky po maximalizaci minimalizaèních kritérií
        .Range("A1:C2").Merge
        With .Range("A1")
            .value = "Maximalizovaná kritéria"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Italic = True
        End With

        ' Nastavení obsahu pro buòky cíle na "max"
        .Range(.Cells(5, 3), .Cells(5 + numOfCriteria - 1, 3)).FormulaR1C1 = "max"
        
        ' Pøeformátování èísla
        Dim cell As Range
        For Each cell In .Range(.Cells(5, 5), .Cells(5 + numOfCriteria - 1, 5 + numOfCandidates - 1)).Cells
            If cell.value = 0 Then
                cell.NumberFormat = "0"
            ElseIf Int(cell.value) = cell.value Then
                ' Pøeformátování èísla pomocí oddìlovaèe tisícù
                cell.NumberFormat = "#,##0"
            Else
                ' Pøeformátování èísla na dvì desetinná místa
                cell.NumberFormat = "0.0#"
            End If
        Next cell
    End With

' Normalizovaná matice
    With ws
        ' Kopírování obsahu první tabulky
        .Range(.Cells(3, 2), .Cells(4 + numOfCriteria + 1, 4 + numOfCandidates)).Copy Destination:=.Cells(6 + numOfCriteria + 3, 2)
        
        ' Slouèení bunìk pro popisek Normalizované matice
        .Range(.Cells(5 + numOfCriteria + 2, 1), .Cells(5 + numOfCriteria + 3, 3)).Merge
        
        With .Cells(5 + numOfCriteria + 2, 1)
            .value = "Normalizovaná matice"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Italic = True
            .EntireColumn.AutoFit
        End With
        
        ' Definice poèáteèní a koncové buòky pro vzorec
        Set startCell = .Cells(4, 5)
        Set endCell = .Cells(4, 5 + numOfCandidates - 1)
        
        ' Definice cílové buòky pro výsledek
        Dim targetCell As Range
        Set targetCell = .Cells(5 + numOfCriteria + 5, 5)
        
        ' Cyklus pro úpravu každé cílové buòky s odpovídajícím indexem øádku
        ' Aplikace vzorce pro celou matici normalizace
        For i = 1 To numOfCriteria
            For j = 1 To numOfCandidates
                ' Pøepoèítání hodnoty cílové buòky pro výpoèet normalizované matice
                targetCell.Offset(i, j - 1).formula = "=(" & startCell.Offset(i, j - 1).Address(False, False) & _
                    "-MIN(" & startCell.Offset(i, 0).Address(False, True) & ":" & endCell.Offset(i, 0).Address(False, True) & "))/(" & _
                    "MAX(" & startCell.Offset(i, 0).Address(False, True) & ":" & endCell.Offset(i, 0).Address(False, True) & _
                    ")-MIN(" & startCell.Offset(i, 0).Address(False, True) & ":" & endCell.Offset(i, 0).Address(False, True) & "))"
            Next j
        Next i
                
        ' Nastavení formátu èísla
        For Each cell In .Range(.Cells(11 + numOfCriteria, 5), .Cells(10 + (2 * numOfCriteria), 5 + numOfCandidates - 1)).Cells
            If Int(cell.value) = cell.value Then
                cell.NumberFormat = "0"
            Else
                ' Pokud hodnota není celé èíslo pak na dvì desetinná místa
                cell.NumberFormat = "0.00"
            End If
        Next cell
        
        ' Suma vah pro kontrolu
        .Cells(11 + (2 * numOfCriteria), 4).formula = "=SUM(" & .Range(.Cells(11 + numOfCriteria, 4), .Cells(10 + (2 * numOfCriteria), 4)).Address & ")"
        .Cells(11 + (2 * numOfCriteria), 4).Font.Bold = True
        
        
    ' Užitek jednotlivých variant
        ' Výpis variant pro pøehlednost užitku
        .Range(.Cells(4, 5), .Cells(4, 5 + numOfCandidates - 1)).Copy Destination:=.Cells(12 + (2 * numOfCriteria), 5)
        
        ' Nastavení popisku pro užitky
        With .Cells(13 + (2 * numOfCriteria), 4)
            .value = "Užitek"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Borders(xlEdgeRight).LineStyle = xlContinuous
        End With
        
        For j = 1 To numOfCandidates
            ' Výpoèet užitkù pro varianty
            .Cells(13 + (2 * numOfCriteria), 4 + j).formula = _
                "=SUMPRODUCT(" & .Range(.Cells(11 + numOfCriteria, 4), .Cells(10 + (2 * numOfCriteria), 4)).Address & "," & _
                        .Range(.Cells(11 + numOfCriteria, 4 + j), .Cells(10 + (2 * numOfCriteria), 4 + j)).Address & ")"
            
            ' Formátování na tøi desetinná místa
            .Cells(13 + (2 * numOfCriteria), 4 + j).NumberFormat = "0.000"
        Next j
        
        ' Pøidání podmínìného formátování barvou pro užitky (Zelený nejlepší, èervený nejhorší)
        .Range(.Cells(13 + (2 * numOfCriteria), 5), .Cells(13 + (2 * numOfCriteria), 4 + numOfCandidates)).FormatConditions.AddColorScale ColorScaleType:=3
        
        ' Popisek nejvyššího užitku
        With .Cells(12 + (2 * numOfCriteria), 6 + numOfCandidates)
            .formula = "Nejvyšší užitek:"
            .Font.Bold = True
            .EntireColumn.AutoFit
        End With
        
        ' Vyhledání a zobrazení užitku od nejvyššího po nejnižší
        Dim rowCounter As Long
        
        ' Nastavení poèáteèního øádku pro výstup
        rowCounter = 12 + (2 * numOfCriteria)
        
        ' Cyklus pro vypsání výsledkù od nejlepšího po nejhoršímu
        For i = 1 To numOfCandidates
            ' Vyhledání a zobrazení i-tého nejvyššího užitku
            .Cells(rowCounter, 7 + numOfCandidates).Formula2 = "=XLOOKUP(LARGE(" & .Range(.Cells(13 + (2 * numOfCriteria), 5), _
                    .Cells(13 + (2 * numOfCriteria), 4 + numOfCandidates)).Address & "," & i & ")" & "," & _
                    .Range(.Cells(13 + (2 * numOfCriteria), 5), .Cells(13 + (2 * numOfCriteria), 4 + numOfCandidates)).Address & "," & _
                    .Range(.Cells(12 + (2 * numOfCriteria), 5), .Cells(12 + (2 * numOfCriteria), 4 + numOfCandidates)).Address & ",,0,1)"
            
            .Cells(rowCounter, 7 + numOfCandidates).HorizontalAlignment = xlCenter
            
            ' Posunutí na další øádek pro další výsledek
            rowCounter = rowCounter + 1
        Next i

        ' Úprava formátování nejlepší varianty
        With .Cells(12 + (2 * numOfCriteria), 7 + numOfCandidates)
            .Font.Bold = False
            .Font.Italic = True
        End With

        ' Nastavení popisku pro vybrání správné varianty
        With .Range(.Cells(9 + numOfCriteria, 5 + numOfCandidates), .Cells(9 + numOfCriteria, 6 + numOfCandidates))
            .Merge
            .value = "Jaká varianta má být správná:"
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .WrapText = False
            '.EntireColumn.AutoFit
            .Select
        End With
        
        ' Úprava šíøky sloupcù (Autofit na minimálnì 80px)
        AdjustColumnWidth ws, .Range(.Columns(5 + numOfCandidates), .Columns(6 + numOfCandidates))
        
        ' Volání funkce, která vykreslí rozbalovací seznam
        ' Parametry jsou WorkSheet (ws), jméno (name), výstup (targetCell), možnosti (optionsRange) a Makro (macroName)
        AddComboBox ws, "newBestCandidateWSA", ws.Cells(9 + numOfCriteria, 7 + numOfCandidates), ws.Range(.Cells(12 + (2 * numOfCriteria), 5), .Cells(12 + (2 * numOfCriteria), 4 + numOfCandidates)), "newBestCandidateWSA_Change"
        
        ' Volání funkce pro pøidání tlaèítka na spuštìní Solveru
        AddButtonTo ws, ws.Cells(8 + (2 * numOfCriteria), 7 + numOfCandidates), "Vyøešit", "CallSolverWSA"
        
        ' Pøidání tlaèítka pro opìtovné nahrání vstupních dat
        AddButtonTo ws, ws.Cells(4, 7 + numOfCandidates), "Aktualizovat", "M3_metoda_WSA"
    End With
    
    ws.Protect "1234"

End Sub

' Procedura metody WSA volající proceduru obsluhující zmìnu hodnoty ComboBoxu
Private Sub newBestCandidateWSA_Change()
    Set ws = ThisWorkbook.Sheets("Metoda WSA")
    
    'Zavolání metody a pøedání parametrù worksheet a název ComboBoxu
    Call newBestCandidate_Change(ws, "newBestCandidateWSA")
End Sub

' Volání Solveru pro metodu WSA
Private Sub CallSolverWSA()
    Set ws = ThisWorkbook.Sheets("Metoda WSA")
    
    ' Pøedání proceduøe obsluhující Solver požadované parametry
    Call M5_Solver(ws, "newBestCandidateWSA")

End Sub
