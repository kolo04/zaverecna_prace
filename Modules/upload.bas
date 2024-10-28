Attribute VB_Name = "Module7"
Dim ws As Worksheet
Dim numOfCriteria As Integer
Dim numOfCandidates As Integer

Sub SetObjectives()

    ' Otevøe formuláø s možností výbìru metody pro zadání cílù
    ObjectivesForm.Show

End Sub

Sub UploadObjectives()
    Dim subject As String
    Dim Objectives As Range
        
    ' Odkaz na list "Vstupní data" a poèet kritérií
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    numOfCriteria = ws.Range("C2").value

    ' Nastavení cílové buòky (zaèátek v C5 + poèet kritérií)
    Set Objectives = ws.Range(ws.Cells(5, 3), ws.Cells(4 + numOfCriteria, 3))

    ' Pøedmìt pro zobrazení v InputBoxu
    subject = "cíle"
    
    ' Volání samostatné procedury pro nahrávání dat
    Call UploadData(Objectives, subject)
    
    Call CheckObjectives(Objectives, ws)
    
End Sub

' Procedura pro nahrání bloku dat (kritéria x varianty) z externího souboru do tabulky
Sub UploadDataBlock()
    Dim srcRange As Range
    Dim targetRange As Range
    Dim validSelection As Boolean

    ' Nastavení listu a zjištìní poètu kritérií a kandidátù
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    numOfCriteria = ws.Range("C2").value
    numOfCandidates = ws.Range("F2").value

    ' Ovìøení, že poèet kritérií a kandidátù je dostateèný
    If numOfCriteria < 2 Then
        MsgBox "Pøi rozhodování bychom mìli zohledòovat minimálnì 2 kritéria.", vbExclamation
        Exit Sub
    End If
    If numOfCandidates < 2 Then
        MsgBox "Pøi rozhodování bychom mìli zohledòovat minimálnì 2 varianty.", vbExclamation
        Exit Sub
    End If

    ' Výbìr oblasti dat k nahrání
RestartSelection:
    validSelection = False
    Set srcRange = Nothing

    ' Uživatel zvolí rozsah dat pomocí InputBoxu
    On Error Resume Next
    Set srcRange = Application.InputBox("Vyberte oblast dat o velikosti " & numOfCriteria & " øádkù a " & numOfCandidates & " sloupcù:", _
                                        "Nahrát data", Type:=8)
    On Error GoTo 0

    ' Kontrola, zda uživatel nìco vybral
    If srcRange Is Nothing Then
        MsgBox "Nebyla vybrána žádná oblast.", vbExclamation
        Exit Sub
    ElseIf srcRange.Rows.Count <> numOfCriteria Or srcRange.Columns.Count <> numOfCandidates Then
        MsgBox "Vybraný rozsah musí mít pøesnì " & numOfCriteria & " øádkù (kritérií) a " & numOfCandidates & " sloupcù (variant).", vbExclamation
        GoTo RestartSelection
    End If

    ' Nastavení cílového rozsahu pro vložení dat v listu
    Set targetRange = ws.Range(ws.Cells(5, 5), ws.Cells(4 + numOfCriteria, 4 + numOfCandidates))

    ' Odemknutí listu pro vložení dat
    ws.Unprotect "1234"

    ' Zkopírování dat z externího souboru do cílového rozsahu
    srcRange.Copy targetRange
    
    ' Pøeformátování èísla
    Dim cell As Range
    For Each cell In targetRange.Cells
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
    
    HideButton ws, "Vložit hodnoty"
    HideButton ws, "Nahrát hodnoty"
    
    ' Pøidání tlaèítka pro úpravu vyplnìných hodnot
    AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Upravit hodnoty", "EditCellValue"
    
    ' Pøidání tlaèítka pro spuštìní metody WSA
    AddButtonTo ws, ws.Range("B" & 9 + numOfCriteria), "Metoda WSA", "M3_metoda_WSA"
    
    ' Pøidání tlaèítka pro spuštìní metody bazické varianty s vìtší šíøkou
    AddButtonTo ws, ws.Range("D" & 9 + numOfCriteria, "E" & 9 + numOfCriteria), "Metoda bazické varianty", "M4_metoda_Bazicke_varianty", 4.5, 1
    
    ' Uzamknutí listu po vložení dat
    ws.Protect "1234"

    MsgBox "Data byla úspìšnì nahrána.", vbInformation
End Sub

' Skript pro kontrolu cílù
Sub CheckObjectives(Objectives As Range, ws As Worksheet)
    Dim validObjectives As Boolean
    Dim cell As Range
    
    ' Kontrola, zda jsou v rozsahu pouze hodnoty "min" nebo "max"
        validObjectives = True
        
        ws.Unprotect "1234"
        
        For Each cell In Objectives
            
            'Nastavení budoucí kontroly - výbìrové pole
            With cell
                options = Array("min", "max")
                
                .Locked = False
                .Validation.Delete
                .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(options, ",")
            End With
            
            'Kontrola hodnot
            If LCase(cell.value) <> "min" And LCase(cell.value) <> "max" Then
                validObjectives = False
                Exit For
            End If
        Next cell
        
        If validObjectives Then
            HideButton ws, "Stanovit cíle"
            
            ' Naètení poètu kritérií
            numOfCriteria = ws.Range("C2").value
        
            ' Pøidání tlaèítka pokraèovat
            AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Pokraèovat", "Candidates"
        Else
            MsgBox "Cílem funkce mùže být pouze minimalizace (min) nebo maximalizace (max)!", vbExclamation
        End If
        
        ws.Protect "1234"
End Sub

' Skript pro kontrolu vah
Sub CheckWeights(weights As Range, ws As Worksheet)
    Dim sumWeights As Float
    
    ' Kontrola, zda jsou všechny váhy vyplnìné
    If CheckFilledCells(weights, "number") Then
        ' Získání souètu vah
        sumWeights = Application.WorksheetFunction.Sum(weights)
        
        ' Zkontroluj, zda je souèet roven 1 (100 %)
        If Not Round(sumWeights, 4) = 1 Then ' Používáme zaokrouhlení pro pøesnost
            MsgBox "Souèet vah není roven 100%! Aktuální souèet: " & Format(sumWeights * 100, "0.00") & "%.", vbExclamation
        End If
    Else
        ' Pokud nejsou všechny váhy vyplnìné
        MsgBox "Nìkteré váhy nejsou vyplnìné.", vbExclamation
    End If
End Sub
