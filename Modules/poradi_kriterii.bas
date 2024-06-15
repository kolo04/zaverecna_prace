Attribute VB_Name = "Module2"
Dim ws As Worksheet
Dim wsInput As Worksheet
Dim wsOutput As Worksheet
Dim numOfCriteria As Integer

' Procedura obsluhující metodu poøadí
Sub MoveToM2()

    ' Zobrazení výsledkù procedury až po kompletním naètení procedury
    Application.ScreenUpdating = False

    ' Získání poètu kritérií
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    numOfCriteria = ws.Range("C2").value
    
    If numOfCriteria < 2 Then
        MsgBox "Pøi rozhodování bychom mìli zohledòovat minimálnì 2 kritéria.", vbExclamation
        Exit Sub
        End
    End If
    
    ' Kontrola existence a vyèištìní listu "Poøadí kritérií"
    wsExists = False
    Set ws = Nothing
    
    For Each ws In ThisWorkbook.Sheets
        If ws.name = "Poøadí kritérií" Then
            wsExists = True
            ws.Activate
            ws.Protect "1234", UserInterfaceOnly:=True
            ws.Cells.Clear
            ActiveSheet.Buttons.Delete
            
            Exit For
        End If
    Next ws
    
    ' Vytvoøení listu, pokud neexistuje
    If Not wsExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = "Poøadí kritérií"
        ws.Activate
        ws.Unprotect "1234"
    End If
    
    Call M2_metoda_poradi

End Sub

' Procedura pracující s listem Poøadí kritérií
Sub M2_metoda_poradi()

    Application.ScreenUpdating = False
    
    ' Definice pracovního listu
    Set ws = ThisWorkbook.Sheets("Poøadí kritérií")
    
    With ws
        ' Vytvoøení záhlaví tabulky
        .Range("B2").value = "Kritérium"
        .Range("C2").value = "Poøadí"
        .Range("B2:C2").Font.Bold = True
    End With
    
    ' Možnost zaèít znovu
    AddButtonTo ws, ws.Range("G2"), "Aktualizovat", "OrderList"
    
    ' Zavolání skriptu OrderList pro poøadí
    Call OrderList
End Sub

' Procedura obsluhující vytvoøení rozevíracího seznamu,
' ve kterém uživatel vybírá poøadí svých priorit pro kritéria
Sub OrderList()

    Application.ScreenUpdating = False
    
    ' Výpis poøadí a tlaèítka pro výpoèet váhy
    Dim changedRows As Collection
    Dim i As Integer
    Dim rowIndex As Variant
    
    Set ws = ThisWorkbook.Sheets("Poøadí kritérií")
    Set wsInput = ThisWorkbook.Sheets("Vstupní data")
    numOfCriteria = wsInput.Range("C2").value
    Set changedRows = New Collection            ' Kolekce k uchování zmìnìných øádkù
    
    With ws
        .Unprotect "1234"
        ' Vyèištìní obsahu sloupcù D a E
        .Columns("D:E").ClearContents

        ' Kontrola zmìn a aktualizace hodnot ve sloupci B
        For i = 3 To 2 + numOfCriteria
            If .Cells(i, 2).value <> wsInput.Cells(i + 2, 2).value Then
                .Cells(i, 2).value = wsInput.Cells(i + 2, 2).value
                changedRows.Add i
            End If
        Next i
        
        ' Nastavení validaèního seznamu a vymazání odpovídajících bunìk ve sloupci C
        If changedRows.Count > 0 Then
            For Each rowIndex In changedRows
                .Cells(rowIndex, 3).value = "Vyberte"
                ' Deklarace dynamického pole øetezcù
                Dim validationArray() As String
                
                ' Pøizpùsobení velikosti pole podle poètu kritérií
                ReDim validationArray(numOfCriteria - 1)
                
                ' Cyklus pøidávající jednotlivá kritéria do pole
                For j = 1 To numOfCriteria
                    validationArray(j - 1) = j
                Next j
                
                ' Odstranìní jakékoliv dosavadní validace a pøidání validace pro výèet hodnot
                With .Cells(rowIndex, 3).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(validationArray, ",")
                End With
                
                ' Nastavení formátu buòky na èíselný formát
                .Cells(rowIndex, 3).NumberFormat = "General"
            Next rowIndex
        End If
        
        ' Úprava šíøky sloupcù (Autofit na minimálnì 80px)
        AdjustColumnWidth ws, ws.Range(ws.Columns(2), ws.Columns(3))
        
        HideButton ws, "Pokraèovat"
        
        ' Pøidání tlaèítka "Vypoèítat váhu"
        AddButtonTo ws, .Range("G9"), "Vypoèítat váhu", "CountWeight"
        
        ' Aktivace buòky s rozevíracím seznamem
        .Range(.Cells(3, 3), .Cells(2 + numOfCriteria, 3)).Locked = False
        .Cells(3, 3).Select
        
        .Protect "1234"
        
    End With
End Sub

' Procedura obsluhující výpoèet váhy
Sub CountWeight()
    Dim i As Integer, j As Integer
    Dim ranks As Object, allRanks As Object, filledRankPoints As Object
    Dim filledRanks As Collection
    Dim value As Variant, rankList() As Variant
    Dim rankSum As Double, rankPoints As Double
    Dim rankIndex As Integer, currentRank As Integer, totalRanks As Integer, rankPos As Integer
    Dim formula As String
    
    ' Inicializace listù
    Set ws = ThisWorkbook.Sheets("Poøadí kritérií")
    Set wsInput = ThisWorkbook.Sheets("Vstupní data")
    Set wsOutput = ThisWorkbook.Sheets("Vstupní data")
    
    ' Naètení poètu kritérií
    numOfCriteria = wsInput.Range("C2").value
    
    ' Inicializace slovníkù a kolekcí pro uchování hodnocení
    Set ranks = CreateObject("Scripting.Dictionary")        ' Slovník k uchování poètu výskytù poøadí
    Set allRanks = CreateObject("Scripting.Dictionary")     ' Slovník pøedávající index a hodnotu poøadí
    Set filledRanks = New Collection                        ' Kolekce uchovávající seznam obsazených pozic
    Set filledRankPoints = CreateObject("Scripting.Dictionary") ' Slovník uchovávající vzorec pro výpoèet bodù z poøadí
    
    ' Ovìøení, zda jsou všechna pole ve sloupci 3 vybrána správnì
    For i = 1 To numOfCriteria
        value = ws.Cells(2 + i, 3).value
        
        ' Pokud není vybrána varianta
        If value = "Vyberte" Or IsEmpty(value) Then
            ws.Cells(2 + i, 3).Select
            MsgBox "Vyplòte prosím všechna poøadí kritérií.", vbExclamation
            Exit Sub
            
        ' Neèíselná nebo nesprávná hodnota
        ElseIf Not IsNumeric(value) Or value < 1 Or value > numOfCriteria Then
            ws.Cells(2 + i, 3).Select
            MsgBox "Poøadí musí být èíslo mezi 1 a " & numOfCriteria & ".", vbExclamation
            Exit Sub
            
        ' Poøadí již evidáno
        ElseIf ranks.Exists(value) Then
            ranks(value) = ranks(value) + 1
            
        ' Pøidání nové hodnoty poøadí
        Else
            ranks.Add value, 1
        End If
        
        ' Pøidání dvojice index a hodnota do kolekce
        allRanks.Add i, value
    Next i
    
    ' Kontrola, zda je jednièka ve sloupci poøadí
    If Not ranks.Exists(1) Then
        MsgBox "Poøadí musí zaèínat od 1.", vbExclamation
        Exit Sub
    End If
    
    ' Vytvoøení seznamu obsazených poøadí
    For Each Key In ranks.Keys
        filledRanks.Add Key
    Next Key
    
    ' Pøidání chybìjících poøadí pro rozdìlení bodù
    For i = 1 To numOfCriteria
        If Not ranks.Exists(i) Then
            filledRanks.Add i
        End If
    Next i
    
    ' Pøevedení kolekce filledRanks na pole
    ReDim rankList(filledRanks.Count - 1)
    rankIndex = 0
    For Each rank In filledRanks
        rankList(rankIndex) = rank
        rankIndex = rankIndex + 1
    Next rank

    ' Výpoèet bodù a jejich pøiøazení
    With ws
        ' Nadpis pro sloupec Bodù
        .Unprotect "1234"
        .Cells(2, 4).value = "Bodù"

        ' Poèet všech možných poøadí
        totalRanks = UBound(rankList) + 1
        
        ' Poèáteèní pozice hodnocení
        rankPos = 1
        
        ' Cyklus procházející všechny hodnocené pozice
        For i = 0 To UBound(rankList)
            currentRank = rankList(i)
            
            ' Výpoèet bodù pro duplicitní poøadí
            If ranks(currentRank) > 1 Then
                rankSum = 0
                formula = ""
                
                ' Pro každé duplicitní poøadí vypoèítá celkový souèet bodù a pøipraví vzorec
                For j = 0 To ranks(currentRank) - 1
                    rankSum = rankSum + (totalRanks + 1 - (rankPos + j))
                    
                    ' Koøen vzorce pro výpoèet bodù
                    If j <> ranks(currentRank) - 1 Then
                        formula = formula & (totalRanks + 1 - (rankPos + j)) & " + "
                    Else
                        formula = formula & (totalRanks + 1 - (rankPos + j))
                    End If
                Next j
                
                ' Výpoèet prùmìrného bodového hodnocení pro duplicitní poøadí
                formula = "= (" & formula & ") / " & ranks(currentRank)
                filledRankPoints(currentRank) = formula
                rankPos = rankPos + ranks(currentRank)
                
            ' Výpoèet bodù pro jedineèné poøadí
            Else
                rankPoints = totalRanks + 1 - rankPos
                formula = "= (" & numOfCriteria & " + 1 - " & rankPos & ") / 1"
                filledRankPoints(currentRank) = formula
                rankPos = rankPos + 1
            End If
        Next i
        
        ' Pøiøazení vzorcù do bunìk
        For Each Key In allRanks.Keys
            .Cells(2 + Key, 4).formula = filledRankPoints(allRanks(Key))
        Next Key
    
        ' Nadpis pro sloupec Váha
        .Cells(2, 5).value = "Váha"
        
        ' Výpoèet váhy jako podíl bodù a celkového poètu bodù
        .Range("E3:E" & 2 + numOfCriteria).formula = "=$D3/(SUM($D$3:$D$" & (2 + numOfCriteria) & "))"
        
        ' Formátování procentuálního stylu s jedním desetinným místem
        .Range("E3:E" & 2 + numOfCriteria).Style = "Percent"
        .Range("E3:E" & 2 + numOfCriteria).NumberFormat = "0.0 %"
        
        ' Tuèné písmo pro záhlaví
        .Range("B2:E2").Font.Bold = True
        
        ' Ohranièení pro nadpisy sloupcù
        With .Range("B2:E2").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        ' Úprava šíøky sloupcù (Autofit na minimálnì 80px)
        AdjustColumnWidth ws, .Range(.Columns(4), .Columns(5))
    End With
    
    HideButton ws, "Vypoèítat váhu"
    
    ' Vložení popisku a vah kritérií z listu Poøadí kritérií do bunìk D5:D (5 + numOfCriteria)
    wsOutput.Unprotect "1234"
    wsOutput.Range("D4").value = "Váha"
    wsOutput.Range("D5:D" & 4 + numOfCriteria).value = ws.Range("E3:E" & 2 + numOfCriteria).value
    HideButton wsOutput, "Stanovit váhu"
    AdjustColumnWidth wsOutput, wsOutput.Range(wsOutput.Columns(2), wsOutput.Columns(4))
    wsOutput.Protect "1234"
    
    ' Pøidání tlaèítka pro návrat na vstupní data
    AddButtonTo ws, ws.Range("G9"), "Pokraèovat", "WeightedInputData"
    
    ws.Protect "1234"
End Sub
