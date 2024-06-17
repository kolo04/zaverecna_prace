Attribute VB_Name = "Module1"
Dim ws As Worksheet
Dim wsInput As Worksheet
Dim wsExists As Boolean
Dim numOfCriteria As Integer
Dim numOfCandidates As Integer
Dim criteriaDone As Boolean

' Pøi otevøení souboru je automaticky spuštìna tato procedura
Sub auto_open()
    
    ' Zavolání procedury Vstupni_data
    Call InputData
End Sub

' Úvodní procedura, která je automaticky spuštìna po otevøení
Sub InputData()
    
    ' Zobrazení výsledkù procedury až po kompletním naètení
    ' Zrychluje proces a zabraòuje nepøijemné "blikání" pøed oèima uživatele
    Application.ScreenUpdating = False
    
    ' Ovìøení existence listu "Vstupní data"
    wsExists = False
    
    For Each ws In ThisWorkbook.Sheets
        If ws.name = "Vstupní data" Then
            wsExists = True
            ' Pøesun na list a jeho vyèištìní
            ws.Activate
            ws.Unprotect "1234"
            ws.Cells.Clear
            
            ' Deklarace promìnné, která je typu Shape
            ' Jakýkoliv objekt, který má tvar = tlaèítko, TextBox, ComboBox, ..
            Dim Shape As Shape
            'Cyklus, který projde všechny objekty typu Shape na listu a odstraní je
            For Each Shape In ws.Shapes
                Shape.Delete
            Next Shape
            
            Exit For
        End If
    Next ws
    
    ' Vytvoøení listu, pokud ještì neexistuje
    If Not wsExists Then
        ' Pøidání listu za poslední již existující list
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = "Vstupní data"
        ' Pøesun na novì vytvoøený list
        ws.Activate
        ws.Unprotect "1234"
    End If
        
    ' Nahrání vstupních dat
    With ws
    
        ' Vytvoøení záhlaví tabulky
        .Range("B2").value = "Poèet kritérií"
        .Range("C2").value = 0 ' Poèet kritérií na zaèátku bude nula
        .Range("B4").value = "Kritérium"
                
        ' Tuèné písmo pro poèet kritérií
        .Range("B2").Font.Bold = True
        .Range("B4").Font.Bold = True
        
        ' Úprava šíøky sloupcù (Autofit na minimálnì 80px)
        AdjustColumnWidth ws, 2
        
        '.Columns("B").EntireColumn.AutoFit
        .Cells(4, 2).Select
        
        Application.ScreenUpdating = True
        
        ' Pokud je criteriaDone Nepravda, pak
        If criteriaDone = False Then
            ' Zavolání/Vytvoøení UserFormu pro zadávání kritérií
            If Not AddCriteriaForm Is Nothing Then
                Unload AddCriteriaForm
                .Unprotect "1234"
                Set AddCriteriaForm = New AddCriteriaForm
                AddCriteriaForm.Show
            End If
            
            .Unprotect "1234"
            
            ' Získání poètu kritérií
            numOfCriteria = .Range("C2").value
            
            ' Kontrola splnìní podmínky pro minimálnì 2 kritéria
            If numOfCriteria < 2 Then
                MsgBox "Pøi rozhodování bychom mìli zohledòovat minimálnì 2 kritéria.", vbExclamation
                .Protect "1234"
                Exit Sub
            End If
        End If
    End With
    
End Sub

' Procedura obsluhující stanovení cílù úèelových funkcí pro jednotlivá kritéria
Sub WeightedInputData()

    Application.ScreenUpdating = False
    
    ' Definice pracovního listu pro vstupní data
    Set wsInput = ThisWorkbook.Sheets("Poøadí kritérií")
    Set ws = ThisWorkbook.Sheets("Vstupní data")
        
    With ws
        .Activate
        .Unprotect "1234"
        
        ' Získání poètu kritérií
        numOfCriteria = .Range("C2").value
        
        ' Nastavení textu "Cíl" do buòky C4
        .Range("C4").value = "Cíl"
        
        ' Vytvoøení rozevíracího seznamu s možnostmi "min" a "max" pro každou buòku v rozsahu C4 až C(4 + numOfCriteria)
        Dim criteriaRange As Range
        Dim options As Variant
        options = Array("min", "max")
        
        Set criteriaRange = .Range(.Cells(5, 3), .Cells(5 + numOfCriteria - 1, 3))
        For Each cell In criteriaRange
            With cell.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(options, ",")
            End With
            ' Nastavení popisku "Vyberte" pro každou buòku
            cell.value = "Vyberte"
            cell.Locked = False
        Next cell
        
        ' Formátování záhlaví B4:D4
        With .Range("B4:D4")
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
    
        ' Zarovnání bunìk C4:C (4 + numOfCriteria) na støed
        .Range("C4:C" & 4 + numOfCriteria).HorizontalAlignment = xlCenter
        
        ' Nastavení stylu bunìk D4:D (4 + numOfCriteria) jako "Percent" s formátem "0.0 %"
        .Range("D4:D" & 4 + numOfCriteria).NumberFormat = "0.0 %"
        
        ' Úprava šíøky sloupcù
        AdjustColumnWidth ws, .Range(.Columns(2), .Columns(3))
        
        .Cells(5, 3).Select
    End With
    
    HideButton ws, "Stanovit váhy"
    
    ' Pøidání tlaèítka pro návrat na vstupní data
    AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Pokraèovat", "Candidates"
    
    ws.Protect "1234"
    
End Sub

' Procedura obsluhující zavolání pøidávání variant a pøidání tlaèítka "Pokraèovat"
' pro pøechod na vyplnìní hodnot tabulky
Sub Candidates()
    Set ws = ThisWorkbook.Sheets("Vstupní data")

    ' Ovìøení, zda jsou všechny cíle vyplnìny
    Dim i As Integer
    For i = 1 To numOfCriteria
        If ws.Cells(4 + i, 3).value = "Vyberte" Then
            ws.Cells(4 + i, 3).Select
            MsgBox "Vyplòte prosím všechny cíle.", vbExclamation
            Exit Sub
        End If
    Next i
    
    With ws
        .Unprotect "1234"
        If IsEmpty(Range("E2")) Then
            ' Pøidávání variant
            .Range("E2").value = "Poèet variant"
            .Range("F2").value = 0 ' Poèet variant na zaèátku bude roven nule
            .Range("E3").value = "Varianta"
            
            ' Tuèné písmo pro poèet variant
            .Range("E2:E3").Font.Bold = True
            
            .Columns("E").EntireColumn.AutoFit
            .Cells(3, 5).Select
        End If
        
        If candidatesDone = False Then
            ' Otevøení UserFormu pro zadávání variant
            If Not AddCandidateForm Is Nothing Then
                Unload AddCandidateForm
                .Unprotect "1234"
                Set AddCandidateForm = New AddCandidateForm
                AddCandidateForm.Show
            End If

            .Unprotect "1234"

        End If
        
        ' Získání poètu variant
        numOfCandidates = .Range("F2").value
        
        ' Kontrola splnìní podmínky pro minimálnì 2 varianty
        If numOfCandidates < 2 Then
            MsgBox "Pøi rozhodování bychom mìli zohledòovat minimálnì 2 varianty.", vbExclamation
            .Protect "1234"
            Exit Sub
            End
        Else
            ' Pøidání tlaèítka pro vyplnìní dat
            ws.Protect "1234", UserInterfaceOnly
        End If
    End With
End Sub

' Procedura pro vyplnìní hodnot tabulky
Sub FillData()
    Dim cellRange As Range

    Set ws = ThisWorkbook.Sheets("Vstupní data")
    
    With ws
        numOfCriteria = .Range("C2").value
        numOfCandidates = .Range("F2").value
        
        ' Kontrola poètu kritérií
        If numOfCriteria < 2 Then
            MsgBox "Pøi rozhodování bychom mìli zohledòovat minimálnì 2 kritéria.", vbExclamation
            Exit Sub
        End If
        
        ' Kontrola poètu variant
        If numOfCandidates < 2 Then
            MsgBox "Pøi rozhodování bychom mìli zohledòovat minimálnì 2 varianty.", vbExclamation
            Exit Sub
        End If
        
        ' Pro každou zmìnìnou buòku kritéria
        For Each cell In Range("B5:B" & 5 + numOfCriteria - 1)
            ' Kontrola, zda jsou pole v D prázdné ve stejném øádku jako pole v sloupci B
            If IsEmpty(cell.Offset(0, 2).value) Then
                MsgBox "Vyplòte, prosím, váhu kritéria.", vbExclamation
                
                ' Oznaèení prázdné buòky
                cell.Offset(0, 2).Select
                
                Exit Sub
                
            ' Kontrola, zda jsou pole v C prázdné ve stejném øádku jako pole v sloupci B
            ElseIf IsEmpty(cell.Offset(0, 1).value) Then
                MsgBox "Vyplòte, prosím, cíl kritéria", vbExclamation
                ' Oznaèení prázdné buòky
                cell.Offset(0, 1).Select
                Exit Sub
            End If
        Next cell
        
        ' Nastavení rozsahu bunìk pro zadání hodnot kritérií a variant
        Set cellRange = ws.Range(ws.Cells(5, 5), ws.Cells(4 + numOfCriteria, 4 + numOfCandidates))
        
        ' Cyklus pro zadání hodnot kritérií a variant
        For Each cell In cellRange
            ' Kontrola, zda je buòka prázdná
            If IsEmpty(cell) Then
                ' Zavolání procedury FillDataForm pouze pro prázdné buòky
                FillDataForm cell
            End If
        Next cell

        ' Kontrola, zda jsou buòky prázdné
        ws.Unprotect "1234"
        
        HideButton ws, "Pokraèovat"
        
        If Not IsEmpty(ws.Range(ws.Cells(5, 5), ws.Cells(4 + numOfCriteria, 4 + numOfCandidates))) Then
            ' Pøidání tlaèítka pro úpravu vyplnìných hodnot
            AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Upravit hodnoty", "EditCellValue"
        Else
            ' Pøidání tlaèítka pro vyplnìní dat
            AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Pokraèovat", "FillData"
        End If
        
        ' Pøidání tlaèítka pro spuštìní metody WSA
        AddButtonTo ws, ws.Range("B" & 9 + numOfCriteria), "Metoda WSA", "M3_metoda_WSA"
        
        ' Pøidání tlaèítka pro spuštìní metody bazické varianty s vìtší šíøkou
        AddButtonTo ws, ws.Range("D" & 9 + numOfCriteria, "E" & 9 + numOfCriteria), "Metoda bazické varianty", "M4_metoda_Bazicke_varianty", 4.5, 1
        
        ws.Protect "1234"
    End With
End Sub

' Procedura pro naplnìní buòky, kterou procedura dostane formou parametru
Sub FillDataForm(cellRef As Variant)
    Dim cell As Range
    Dim criteriaNames() As Variant
    Dim variantNames() As Variant
    Dim inputVal As Variant
    Dim validInput As Boolean
    Dim convertedVal As Double

    Set ws = ThisWorkbook.Sheets("Vstupní data")
    
    ' Pøetypujeme referenci na buòku na objekt typu Range
    Set cell = cellRef

    ' Získání názvu kritéria
    criteriaName = ws.Cells(cell.Row, 2).value
    
    ' Získání názvu varianty
    variantName = ws.Cells(4, cell.column).value
    
    ' Oznaèení buòky pro zadání hodnoty
    cell.Select
    
    Do
        ' Kontrola, zda buòka obsahuje již hodnotu
        If Not IsEmpty(cell.value) Then
            ' Pokud buòka již obsahuje hodnotu, nabídne se možnost její úpravy
            inputVal = InputBox("Aktuální hodnota pro kritérium '" & criteriaName & "' a variantu '" & variantName & "' je: " & _
                        cell.value & vbCrLf & "Zadejte novou hodnotu nebo kliknìte na OK pro ponechání stávající hodnoty:", _
                        "Hodnota pro kritérium a variantu", cell.value)
    
        Else
            ' Pokud buòka neobsahuje hodnotu, standardní postup zadávání nové hodnoty
            inputVal = InputBox("Zadejte hodnotu pro kritérium '" & criteriaName & "' a variantu '" & variantName & "':", _
            "Hodnota pro kritérium a variantu")
        End If
        
        ' Kontrola, zda uživatel klikl na Cancel
        If inputVal = "" Then
            MsgBox "Zadání bylo zrušeno.", vbInformation
            ws.Protect "1234"
            End
        End If
        
        ' Pokus o pøevod zadané hodnoty na èíslo
        If IsNumeric(inputVal) Then
            ' Pøevod na èíslo (Double)
            convertedVal = CDbl(inputVal)
            
            ' Uložení hodnoty do buòky
            ws.Unprotect "1234"
            cell.value = convertedVal
            'Nastavení èíselného formátu buòky
            If cell.value = 0 Then
                cell.NumberFormat = "0"
            ElseIf Int(cell.value) = cell.value Then
                ' Pøeformátování èísla pomocí oddìlovaèe tisícù
                cell.NumberFormat = "#,##0"
            Else
                ' Pøeformátování èísla na dvì desetinná místa
                cell.NumberFormat = "0.0#"
            End If
            
            ws.Protect "1234"
            validInput = True ' Platný vstup
        Else
            ' Hodnota není èíslo, zobrazení chybové zprávy a cyklus pokraèuje
            MsgBox "Zadávejte, prosím, pouze èíselné hodnoty." & vbCrLf & _
            "V pøípadì kritéria 'ano/ne' vkládejte hodnoty 1 pro 'ano' a 0 pro 'ne'.", vbExclamation
            validInput = False
        End If
    Loop Until validInput
    
End Sub

' Procedura kontrolující, zda jsou hodnoty tabulky vyplnìny
Sub CheckFilledCells()
    ' Nastavení pracovního listu
    Set ws = ThisWorkbook.Sheets("Vstupní data")

    ' Získáme poèet kritérií a poèet variant
    numOfCriteria = ws.Range("C2").value
    numOfCandidates = ws.Range("F2").value
    
   ' Procházení všech bunìk v daném rozsahu
    For j = 1 To numOfCandidates
        For i = 1 To numOfCriteria
            ' Nastavení buòky
            Set cell = ws.Cells(4 + i, 4 + j)
            
            ' Kontrola, zda je buòka prázdná
            If IsEmpty(cell) Then
                ' Zavolání procedury FillDataForm pro prázdné buòky
                FillDataForm cell
                
                ' Pokud najde prázdnou buòku, ukonèí kontrolu
                Exit Sub
            End If
        Next i
    Next j
End Sub

' Kód pro vytvoøení formuláøe, který umožní uživateli vybrat buòku
Sub EditCellValue()
    Dim selectedRange As Range
    Dim cell As Range
    
    ' Umožní uživateli vybrat buòku/buòky
    On Error Resume Next
    Set selectedRange = Application.InputBox("Vyberte buòku (buòky), kterou (které) chcete upravit:", Type:=8)
    On Error GoTo 0
    
    ' Pokud uživatel klikl na Cancel, ukonèíme proceduru
    If selectedRange Is Nothing Then
        Exit Sub
    End If
    
    ' Projdeme každou vybranou buòku z rozsahu
    For Each cell In selectedRange
        ' Zavoláme proceduru FillDataForm pro každou buòku zvláš
        FillDataForm cell
    Next cell
    
End Sub

' Procedura volající formuláø pro pøidání dalších kritérií
Sub AddMoreCriteria()

' Nastavení hodnoty criteriaDone na False pro pøidání dalších kritérií
    criteriaDone = False
    
    ThisWorkbook.ActiveSheet.Unprotect "1234"
    
    ' Zavolání formuláøe
    AddCriteriaForm.Show
End Sub

' Procedura volající formuláø pro pøidání dalších variant
Sub AddMoreCandidates()

' Tlaèítko pro pøidání dalších variant
    candidatesDone = False
    
    ThisWorkbook.ActiveSheet.Unprotect "1234"
    
    AddCandidateForm.Show
End Sub

' Procedura volá a naplòuje formuláø pro odebrání kritéria
Sub RemoveCriteria()
    Dim criteriaList As Range
    Dim criteriaCell As Range
    
    ' Nastavení pracovního listu, kde jsou kritéria uložena
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    
    ' Získat poèet kritérií z listu
    numOfCriteria = ws.Range("C2").value
    
    ' Definuj rozsah obsahující kritéria
    Set criteriaList = ws.Range("B5:B" & 5 + numOfCriteria - 1)
    
    ' Vynuluj ListBox
    RemoveCriteriaForm.CriteriaListBox.Clear

    ' Naplò ListBox seznamem existujících kritérií
    For Each criteriaCell In criteriaList
        RemoveCriteriaForm.CriteriaListBox.AddItem criteriaCell.value
    Next criteriaCell

    ' Zavolání formuláøe pro odebrání kritérií
    RemoveCriteriaForm.Show

End Sub

' Procedura pro odebrání varianty
Sub RemoveCandidate()
    Dim candidateList As Range
    Dim candidateCell As Range
    
    ' Nastavení pracovního listu, kde jsou kritéria uložena
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    
    ' Získat poèet variant z listu
    numOfCandidates = ws.Range("F2").value
    
    ' Definování rozsahu obsahující varianty
    Set candidateList = ws.Range(ws.Cells(4, 5), ws.Cells(4, 5 + numOfCandidates - 1))

    ' Vyprázdìní ListBoxu
    RemoveCandidateForm.CandidateListBox.Clear
    
    ' Naplnìní ListBox seznamem existujících variant
    For Each candidateCell In candidateList
        RemoveCandidateForm.CandidateListBox.AddItem candidateCell.value
    Next candidateCell
    
    ' Zavolání formuláøe pro odebrání varianty
    RemoveCandidateForm.Show
End Sub

