Attribute VB_Name = "Module1"
Dim ws As Worksheet
Dim wsInput As Worksheet
Dim wsExists As Boolean
Dim numOfCriteria As Integer
Dim numOfCandidates As Integer
Dim criteriaDone As Boolean

' Globální promìnná pro sledování stavu ukonèení úprav hodnot
Dim cancelEditing As Boolean

' Pøi otevøení souboru je automaticky spuštìna tato procedura
Sub Auto_Open()
    
    ' Zobrazení výsledkù procedury až po kompletním naètení procedury
    Application.ScreenUpdating = False
    
    ' Zavolání formuláøe pro výbìr metody zadání vstupních dat
    EntryForm.Show
    
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
        ThisWorkbook.Unprotect "1234"
        
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
        .Range("B4:D4").Font.Bold = True
        
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

' Procedura obsluhující zavolání pøidávání variant a pøidání tlaèítka "Pokraèovat"
' pro pøechod na vyplnìní hodnot tabulky
Sub Candidates()
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    
    ' Získání poètu kritérií
    numOfCriteria = ws.Range("C2").value

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
        
        ' Inicializace promìnné pro sledování stavu
        cancelEditing = False
        
        ' Cyklus pro zadání hodnot kritérií a variant
        For Each cell In cellRange
            ' Kontrola, zda je buòka prázdná
            If IsEmpty(cell) Then
                ' Zavolání procedury FillDataForm pouze pro prázdné buòky
                FillDataForm cell
                
                ' Kontrola, zda došlo k zrušení procesu
                If cancelEditing Then
                    Exit Sub
                End If
            End If
        Next cell

        ' Kontrola, zda jsou buòky prázdné
        ws.Unprotect "1234"
        
        HideButton ws, "Pokraèovat"
        
        If Not IsEmpty(ws.Range(ws.Cells(5, 5), ws.Cells(4 + numOfCriteria, 4 + numOfCandidates))) Then
            ' Pokud jsou buòky již vyplnìny, není tøeba je znovu vkládat
            HideButton ws, "Vložit hodnoty"
            HideButton ws, "Nahrát hodnoty"
            
            ' Pøidání tlaèítka pro úpravu vyplnìných hodnot
            AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Upravit hodnoty", "EditCellValue"
        Else
            ' Pøidání tlaèítka pro vyplnìní dat
            AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Vložit hodnoty", "FillData"
            
            ' Pøidání tlaèítka pro nahrání dat
            AddButtonTo ws, ws.Range("F" & 9 + numOfCriteria), "Nahrát hodnoty", "UploadDataBlock"
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
    Dim criteriaName As String
    Dim variantName As String
    Dim inputVal As Variant
    Dim validInput As Boolean
    Dim convertedVal As Double
    
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    
    ' Pøetypujeme referenci na buòku na objekt typu Range
    Set cell = cellRef
    
    ' Získání poètu kritérií a variant pro urèení rozsahu oblastí
    numOfCriteria = ws.Range("C2").value
    numOfCandidates = ws.Range("F2").value
    
    ' Získání názvu kritéria a varianty
    criteriaName = ws.Cells(cell.row, 2).value
    variantName = ws.Cells(4, cell.column).value
    
    ' Oznaèení buòky pro zadání hodnoty a zobrazení aktuální hodnoty
    cell.Select
    Do
        ' Pokud má buòka již hodnotu, nabídneme ji uživateli ke zmìnì
        If Not IsEmpty(cell.value) Then
            inputVal = InputBox("Aktuální hodnota pro kritérium '" & criteriaName & "' a variantu '" & variantName & "' je: " & _
                        cell.value & vbCrLf & "Zadejte novou hodnotu nebo kliknìte na OK pro ponechání stávající hodnoty:", _
                        "Hodnota pro kritérium a variantu", cell.value)
        Else
            inputVal = InputBox("Zadejte hodnotu pro kritérium '" & criteriaName & "' a variantu '" & variantName & "':")
        End If

        ' Pokud uživatel klikne na Cancel, ukonèíme proceduru
        If inputVal = "" Then
            MsgBox "Zadání bylo zrušeno.", vbInformation
            cancelEditing = True  ' Nastavení Boolean promìnné pro možnost ukonèení zadávání
            ws.Protect "1234"
            Exit Sub
        End If

        ' Ovìøení, zda je zadaná hodnota èíselná
        If IsNumeric(inputVal) Then
            convertedVal = CDbl(inputVal)
            ws.Unprotect "1234"
            cell.value = convertedVal
            ' Nastavení èíselného formátu buòky
            If cell.value = 0 Then
                cell.NumberFormat = "0"
            ElseIf Int(cell.value) = cell.value Then
                cell.NumberFormat = "#,##0"
            Else
                cell.NumberFormat = "0.0#"
            End If
            ws.Protect "1234"
            validInput = True
        Else
            MsgBox "Zadávejte, prosím, pouze èíselné hodnoty." & vbCrLf & _
            "V pøípadì kritéria 'ano/ne' vkládejte hodnoty 1 pro 'ano' a 0 pro 'ne'.", vbExclamation
            validInput = False
        End If
        
    Loop Until validInput
End Sub

' Procedura kontrolující, zda jsou hodnoty tabulky vyplnìny
Sub CheckFilledData()
    Dim cell As Range

    ' Nastavení pracovního listu
    Set ws = ThisWorkbook.Sheets("Vstupní data")

    ' Získání poètu kritérií a poèet variant
    numOfCriteria = ws.Range("C2").value
    numOfCandidates = ws.Range("F2").value

    ' Procházení všech bunìk v daném rozsahu
    For j = 1 To numOfCandidates
        For i = 1 To numOfCriteria
            ' Nastavení buòky
            Set cell = ws.Cells(4 + i, 4 + j)
            
            ' Kontrola, zda je buòka prázdná
            If IsEmpty(cell) Then
                ' Upozornìní uživatele na prázdnou buòku
                MsgBox "Buòka " & cell.Address & " je prázdná. Prosím, vyplòte ji.", vbExclamation

                ' Zavolání procedury FillDataForm pro vyplnìní buòky
                FillDataForm cell
                
                ' Po nalezení chyby ukonèíme kontrolu
                Exit Sub
            ' Kontrola, zda buòka neobsahuje èíslo
            ElseIf Not IsNumeric(cell.value) Then
                ' Upozornìní uživatele na neèíselnou hodnotu
                MsgBox "Buòka " & cell.Address & " obsahuje neèíselnou hodnotu." & vbCrLf & _
                "V pøípadì kritéria 'ano/ne' vkládejte hodnoty 1 pro 'ano' a 0 pro 'ne'.", vbExclamation

                ' Zavolání procedury FillDataForm pro opravu hodnoty
                FillDataForm cell
                
                ' Po nalezení chyby ukonèíme kontrolu
                Exit Sub
            End If
        Next i
    Next j

End Sub

' Kód pro vytvoøení formuláøe, který umožní uživateli upravit buòku
Sub EditCellValue()
    Dim selectedRange As Range
    Dim cell As Range
    
    ' Nastavení pracovního listu
    Set ws = ThisWorkbook.Sheets("Vstupní data")

    ' Získání poètu kritérií a poètu variant
    numOfCriteria = ws.Range("C2").value
    numOfCandidates = ws.Range("F2").value
    
    ' Umožní uživateli vybrat buòku/buòky
    On Error Resume Next
    Set selectedRange = Application.InputBox("Vyberte buòku (buòky), kterou (které) chcete upravit:", Type:=8)
    On Error GoTo 0

    ' Pokud uživatel klikl na Cancel, ukonèíme proceduru
    If selectedRange Is Nothing Then
        Exit Sub
    End If

    ' Definování platného rozsahu, ve kterém lze mìnit hodnoty
    Dim validRange As Range
    Set validRange = ws.Range(ws.Cells(5, 5), ws.Cells(4 + numOfCriteria, 4 + numOfCandidates))
    
    ' Inicializace promìnné pro sledování stavu
    cancelEditing = False

    ' Kontrola pro každou vybranou buòku z rozsahu, zda je buòka v povoleném rozsahu
    For Each cell In selectedRange
        If Not Intersect(cell, validRange) Is Nothing Then
        
            ' Zavoláme proceduru FillDataForm pro každou povolenou buòku
            FillDataForm cell
            
            ' Kontrola, zda došlo k zrušení procesu
            If cancelEditing Then
                Exit Sub
            End If
        Else
            ' Pokud je buòka mimo povolený rozsah, zobrazíme varování a pøeskoèíme ji
            MsgBox "Buòku " & cell.Address & " nelze upravit.", vbExclamation
        End If
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
