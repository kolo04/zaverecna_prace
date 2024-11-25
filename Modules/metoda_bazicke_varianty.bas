Attribute VB_Name = "Module4"
Dim ws As Worksheet
Dim wsInput As Worksheet
Dim numOfCriteria As Integer
Dim numOfCandidates As Integer
Dim Shape As Shape

' Makro pro výpoèet metody Bazické varianty
Sub M4_metoda_Bazicke_varianty()

    ' Zobrazení výsledkù procedury až po kompletním naètení procedury
    Application.ScreenUpdating = False
    
    ' Volání kontroly vyplnìných hodnot
    Call CheckFilledData
    
    ' Kontrola unikátních hodnot v øádcích
    If CheckUniqueRowValues() Then
        Exit Sub
    End If
    
' Ovìøení existence listu "Metoda bazické varianty"
    Dim wsExists As Boolean
    wsExists = False
    
    For Each ws In ThisWorkbook.Sheets
        If ws.name = "Metoda bazické varianty" Then
            wsExists = True
            ' Pøesun na list a jeho vyèištìní
            ws.Activate
            ws.Unprotect "1234"
            ws.Cells.Clear
            ws.Columns.AutoFit
            
            For Each Shape In ws.Shapes
                Shape.Delete
            Next Shape
            Exit For
        End If
    Next ws
    
    ' Vytvoøení listu, pokud ještì neexistuje
    If Not wsExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = "Metoda bazické varianty"
        
        ' Pøesun na novì vytvoøený list
        ws.Activate
        ws.Unprotect "1234"
    End If
    
' Práce s "Metoda bazické varianty"
    ' Nahrání vstupních dat
    Dim lastRow As Long
    Dim i As Long
    Dim maxVal As Double
    
    ' Definice pracovních listù
    Set wsInput = ThisWorkbook.Sheets("Vstupní data")
    Set ws = ThisWorkbook.Sheets("Metoda bazické varianty")
    
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
        
        ' Formátování záhlaví B3 až sloupec 5 + numOfCandidates na øádku 4
        With .Range(.Cells(3, 2), .Cells(4, 5 + numOfCandidates))
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
    
    With ws
    ' Vytvoøení sloupce "Báze" a vložení hodnot
        .Cells(4, 5 + numOfCandidates).value = "Báze"
        
        ' Naplnìní sloupce "Báze" podle extrémù øádkù
        For i = 5 To 5 + numOfCriteria
            If .Cells(i, 3).value = "min" Then
                .Cells(i, 5 + numOfCandidates).formula = "=MIN(" & .Cells(i, 5).Address & ":" & .Cells(i, 4 + numOfCandidates).Address & ")"
            ElseIf .Cells(i, 3).value = "max" Then
                .Cells(i, 5 + numOfCandidates).formula = "=MAX(" & .Cells(i, 5).Address & ":" & .Cells(i, 4 + numOfCandidates).Address & ")"
            End If
        Next i

        ' Formátování sloupce "Báze"
        With Range(.Cells(4, 5 + numOfCandidates), .Cells(4 + numOfCriteria, 5 + numOfCandidates))
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlInsideVertical).LineStyle = xlContinuous
        End With
        
        ' Popisek tabulky po pøidání báze
        .Range("A1:C2").Merge
        With .Range("A1")
            .value = "Kritéria a jejich báze"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Italic = True
        End With
        
        ' Pøeformátování èísla
        Dim cell As Range
        For Each cell In .Range(.Cells(5, 5), .Cells(5 + numOfCriteria - 1, 5 + numOfCandidates)).Cells
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
        
        ' Úprava šíøky sloupcù (Autofit na minimálnì 80px)
        AdjustColumnWidth ws, .Range(.Columns(2), .Columns(4 + numOfCandidates))
        ws.Columns(5 + numOfCandidates).ColumnWidth = 10.33
        
    End With
    
' Normalizovaná matice
    With ws
        ' Kopírování obsahu první tabulky
        .Range(.Cells(3, 2), .Cells(5 + numOfCriteria + 1, 4 + numOfCandidates)).Copy Destination:=.Cells(6 + numOfCriteria + 3, 2)
        
        ' Slouèení bunìk pro popisek Normalizované matice
        .Range(.Cells(5 + numOfCriteria + 2, 1), .Cells(5 + numOfCriteria + 3, 3)).Merge
        
        With .Cells(5 + numOfCriteria + 2, 1)
            .value = "Normalizovaná matice"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Italic = True
            .EntireColumn.AutoFit
        End With
        
        Dim epsilon As Double
        epsilon = 0.0000000001 ' Malé èíslo pro ošetøení dìlení 0

        ' Cyklus pro vypoèítání normalizované hodnoty buòky
        For i = 1 To numOfCriteria
            For j = 1 To numOfCandidates
            
                ' Vypoèítání normalizované hodnoty
                If .Cells(4 + i, 3).value = "min" Then
                    If .Cells(4 + i, 4 + j).value = 0 Then
                        .Cells(4 + i, 4 + j).value = epsilon
                        .Cells(10 + numOfCriteria + i, 4 + j).formula = "=" & .Cells(4 + i, 5 + numOfCandidates).Address(False, True) & "/" & .Cells(4 + i, 4 + j).Address(False, False)
                    Else
                        .Cells(10 + numOfCriteria + i, 4 + j).formula = "=" & .Cells(4 + i, 5 + numOfCandidates).Address(False, True) & "/" & .Cells(4 + i, 4 + j).Address(False, False)
                    End If
                ElseIf .Cells(4 + i, 3).value = "max" Then
                    If .Cells(4 + i, 5 + numOfCandidates).value = 0 Then
                        .Cells(4 + i, 5 + numOfCandidates).value = epsilon
                        .Cells(10 + numOfCriteria + i, 4 + j).formula = "=" & .Cells(4 + i, 4 + j).Address(False, False) & "/" & .Cells(4 + i, 5 + numOfCandidates).Address(False, True)
                    Else
                        .Cells(10 + numOfCriteria + i, 4 + j).formula = "=" & .Cells(4 + i, 4 + j).Address(False, False) & "/" & .Cells(4 + i, 5 + numOfCandidates).Address(False, True)
                    End If
                End If
            Next j
        Next i
        
        ' Nastavení formátu èísla
        For Each cell In .Range(.Cells(11 + numOfCriteria, 5), .Cells(11 + (2 * numOfCriteria), 5 + numOfCandidates - 1)).Cells
            If Int(cell.value) = cell.value Then
                cell.NumberFormat = "0"
            Else
                ' Pokud hodnota není celé èíslo pak na dvì desetinná místa
                cell.NumberFormat = "0.00"
            End If
        Next cell

        ' Suma vah pro kontrolu
        .Cells(11 + (2 * numOfCriteria), 4).formula = "=SUM(" & .Range(.Cells(11 + numOfCriteria, 4), .Cells(11 + (2 * numOfCriteria) - 1, 4)).Address & ")"
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
        With .Cells(17 + numOfCriteria, 6 + numOfCandidates)
            .formula = "Nejvyšší užitek:"
            .Font.Bold = True
            .EntireColumn.AutoFit
        End With
        
        ' Cyklus pro vypsání výsledkù od nejlepšího po nejhoršímu
        For i = 1 To numOfCandidates
            ' Vyhledání a zobrazení i-tého nejvyššího užitku
            .Cells(16 + numOfCriteria + i, 7 + numOfCandidates).Formula2 = "=XLOOKUP(LARGE(" & .Range(.Cells(13 + (2 * numOfCriteria), 5), _
                    .Cells(13 + (2 * numOfCriteria), 4 + numOfCandidates)).Address & "," & i & ")" & "," & _
                    .Range(.Cells(13 + (2 * numOfCriteria), 5), .Cells(13 + (2 * numOfCriteria), 4 + numOfCandidates)).Address & "," & _
                    .Range(.Cells(12 + (2 * numOfCriteria), 5), .Cells(12 + (2 * numOfCriteria), 4 + numOfCandidates)).Address & ",,0,1)"
            
            .Cells(16 + numOfCriteria + i, 7 + numOfCandidates).HorizontalAlignment = xlCenter
            
        Next i
            
        ' Úprava formátování nejlepší varianty
        With .Cells(17 + numOfCriteria, 7 + numOfCandidates)
            .Font.Bold = False
            .Font.Italic = True
        End With

        ' Nastavení popisku pro vybrání testované varianty
        With .Range(.Cells(9 + numOfCriteria, 5 + numOfCandidates), .Cells(9 + numOfCriteria, 6 + numOfCandidates))
            .Merge
            .value = "Jaká varianta má být testována:"
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .WrapText = False
            '.EntireColumn.AutoFit
            .Select
        End With
        
        ' Zavolání procedury pro kontrolu, zda je nìkterá z variant dominována jinou
        Call FindDominatedCandidates(ws)
        
        ' Úprava šíøky sloupcù (Autofit na minimálnì 80px)
        AdjustColumnWidth ws, .Range(.Columns(5 + numOfCandidates), .Columns(7 + numOfCandidates))
        
        If .Columns(5 + numOfCandidates).ColumnWidth < 12 Then
            .Columns(5 + numOfCandidates).ColumnWidth = 12
        End If
        
        ' Volání funkce, která vykreslí rozbalovací seznam
        ' Parametry jsou WorkSheet (ws), výstup (targetCell) a možnosti (optionsRange)
        AddComboBox ws, "newBestCandidateBV", ws.Cells(9 + numOfCriteria, 7 + numOfCandidates), ws.Range(.Cells(12 + (2 * numOfCriteria), 5), .Cells(12 + (2 * numOfCriteria), 4 + numOfCandidates)), "newBestCandidateBV_Change"
        
        ' Volání funkce pro pøidání tlaèítka na spuštìní Solveru
        AddButtonTo ws, ws.Cells(14 + numOfCriteria, 7 + numOfCandidates), "Vyøešit", "CallSolverBV"
        
        ' Pøidání tlaèítka pro opìtovné nahrání vstupních dat
        AddButtonTo ws, ws.Cells(4, 7 + numOfCandidates), "Aktualizovat", "M4_metoda_Bazicke_varianty"
        
    End With
    
    ws.Protect "1234"
    
End Sub

' Procedura metody bazické varianty volající proceduru obsluhující zmìnu hodnoty ComboBoxu
Private Sub newBestCandidateBV_Change()
    Set ws = ThisWorkbook.Sheets("Metoda bazické varianty")
    
    'Zavolání metody a pøedání parametrù worksheet a název ComboBoxu
    Call newBestCandidate_Change(ws, "newBestCandidateBV")
End Sub

' Volání Solveru pro metodu bazické varianty
Private Sub CallSolverBV()

    Set ws = ThisWorkbook.Sheets("Metoda bazické varianty")
    
    ' Pøedání proceduøe obsluhující Solver požadované parametry
    Call M5_Solver(ws, "newBestCandidateBV")

End Sub
