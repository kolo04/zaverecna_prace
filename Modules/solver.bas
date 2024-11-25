Attribute VB_Name = "Module5"
Dim ws As Worksheet
Dim wsInput As Worksheet
Dim numOfCriteria As Integer
Dim numOfCandidates As Integer
Dim cbo As Object

Dim selectedCandidate As String
Dim selectedColumn As Integer

' Makro pro spouštìní Solveru
Sub M5_Solver(ws As Worksheet, cboName As String)
    Dim metoda As String
    Dim minSheetName As String
    Dim maxSheetName As String
    Dim wsMin As Worksheet
    Dim wsMax As Worksheet

    ' Získání referencí na listy
    Set ws = ws
    
    ' Získání hodnoty vybrané v ComboBoxu
    For Each cbo In ws.DropDowns
        If cbo.name = cboName Then
        
            ' Kontrola, zda je ComboBox obsahuje nìjaké varianty
            If cbo.ListCount > 0 Then
            
                ' Kontrola, zda je vybrána nìjaká varianta
                If cbo.ListIndex = 0 Then
                    MsgBox "Zvolte, prosím, testovanou variantu.", vbExclamation
                    Exit Sub
                Else
                    ' Získání hodnoty vybrané v ComboBoxu a její pøevod na øetìzec
                    selectedCandidate = CStr(cbo.List(cbo.ListIndex))
                    
                    ' Získání referencí na listy
                    Set wsInput = ThisWorkbook.Sheets("Vstupní data")
                    
                    numOfCriteria = wsInput.Range("C2").value
                    numOfCandidates = wsInput.Range("F2").value
                    
                    ' Výpoèet klíèové funkce: použití Match pro získání hodnoty varianty
                    selectedColumn = Application.WorksheetFunction.Match(selectedCandidate, _
                        ws.Range(ws.Cells(12 + (2 * numOfCriteria), 5), ws.Cells(12 + (2 * numOfCriteria), 4 + numOfCandidates)), 0)
                End If
            Else
                MsgBox "Není k dispozici žádná varianta k výbìru.", vbExclamation
                Exit Sub
            End If
        End If
    Next cbo

    ' Urèení metody na základì názvu listu
    Select Case ws.name
        Case "Metoda WSA"
            metoda = "WSA"
        Case "Metoda bazické varianty"
            metoda = "SAW"
        Case Else
            MsgBox "Neznámá metoda! Ujistìte se, že voláte Solver z platného listu.", vbCritical
            Exit Sub
    End Select

    ' Urèení názvù listù podle metody
    minSheetName = metoda & "_MIN"
    maxSheetName = metoda & "_MAX"

    ' Vytvoøení a úprava listu minimalizace
    CreateAndConfigureSheet ws, selectedCandidate, selectedColumn, minSheetName, True
    
    ' Spuštìní Solveru pro minimalizaci
    ThisWorkbook.Sheets(minSheetName).Activate
    SolverSolve UserFinish:=True
    SolverFinish KeepFinal:=1

    ' Vytvoøení a úprava listu maximalizace
    CreateAndConfigureSheet ws, selectedCandidate, selectedColumn, maxSheetName, False
    
    ' Spuštìní Solveru pro maximalizaci
    ThisWorkbook.Sheets(maxSheetName).Activate
    SolverSolve UserFinish:=True
    SolverFinish KeepFinal:=1
    
    Set wsMin = ThisWorkbook.Sheets(minSheetName)
    Set wsMax = ThisWorkbook.Sheets(maxSheetName)

    
    With wsMin
        .Range(.Cells(15 + (2 * numOfCriteria) + 1, 4), .Cells(15 + (3 * numOfCriteria), 4)).value _
                = .Range(.Cells(10 + numOfCriteria + 1, 4), .Cells(10 + (2 * numOfCriteria), 4)).value
        .Protect "1234"
    End With
    
    With wsMax
        .Range(.Cells(15 + (2 * numOfCriteria) + 1, 4), .Cells(15 + (3 * numOfCriteria), 4)).value _
                = .Range(.Cells(10 + numOfCriteria + 1, 4), .Cells(10 + (2 * numOfCriteria), 4)).value
        .Protect "1234"
    End With
    
    With ws
        .Activate
        
        With .Range(.Cells(11 + numOfCriteria, 7 + numOfCandidates), .Cells(12 + numOfCriteria, 7 + numOfCandidates))
            .NumberFormat = "0.0 %"
            .Select
        End With
        .Protect "1234"
    End With

    ' Oznámení dokonèení
    MsgBox "Výsledky jsou dostupné na listech: " & minSheetName & " a " & maxSheetName, vbInformation
    
End Sub

Sub CreateAndConfigureSheet(ws As Worksheet, selectedCandidate As String, selectedColumn As Integer, sheetName As String, isMin As Boolean)
    Dim wsOutput As Worksheet
    Dim objective As String
    Dim metoda As String

    ' Odstranìní existujícího listu, pokud již existuje
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets(sheetName)
    If Not wsOutput Is Nothing Then
        Application.DisplayAlerts = False
        wsOutput.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' Vytvoøení nového listu
    Set wsOutput = ThisWorkbook.Sheets.Add
    wsOutput.name = sheetName
    
    Set wsInput = ThisWorkbook.Sheets("Vstupní data")
    Set ws = ws
    numOfCriteria = wsInput.Range("C2").value
    numOfCandidates = wsInput.Range("F2").value
    
    ' Urèení metody na základì názvu listu
    If ws.name = "Metoda WSA" Then
            metoda = "WSA"
        Else
            metoda = "bazické varianty"
    End If
    
    ' Kopírování obsahu
    With wsOutput
        ws.Range(ws.Cells(3, 1), ws.Cells(13 + (2 * numOfCriteria), 4 + numOfCandidates)).Copy Destination:=.Cells(3, 1)
        
        If ws.name = "Metoda bazické varianty" Then
            ws.Range(ws.Cells(4, 4 + numOfCandidates + 1), ws.Cells(4 + numOfCriteria, 4 + numOfCandidates + 1)).Copy Destination:=.Cells(4, 4 + numOfCandidates + 1)
        End If
        
        ' Kopírování poøadí
        ws.Range(ws.Cells(16 + numOfCriteria + 1, 6 + numOfCandidates), ws.Cells(16 + numOfCriteria + numOfCandidates, 7 + numOfCandidates)).Copy _
            Destination:=.Cells(14 + numOfCriteria, 6 + numOfCandidates)
            
        ' Nastavení popisku pro klíèovou funkci
        With .Range(.Cells(10 + numOfCriteria, 5 + numOfCandidates), .Cells(10 + numOfCriteria, 6 + numOfCandidates))
            .Merge
            .value = "Užitek vybrané varianty:"
            .HorizontalAlignment = xlCenter
            .WrapText = False
            .Font.Bold = True
        End With

        If .Columns(5 + numOfCandidates).ColumnWidth < 12 Then
            .Columns(5 + numOfCandidates).ColumnWidth = 12
        End If
        
        ' Pøedání vzorce buòky s odkazem na vybranou variantu
        .Cells(10 + numOfCriteria, 7 + numOfCandidates).formula = ws.Cells(10 + numOfCriteria, 7 + numOfCandidates).formula
        
        ' Nastavení nadpisù v tabulce
        .Cells(15 + (2 * numOfCriteria), 2).value = "Kritérium"
        .Cells(15 + (2 * numOfCriteria), 3).value = "Wj"
        .Cells(15 + (2 * numOfCriteria), 4).value = "Wj'"
        .Cells(15 + (2 * numOfCriteria), 5).value = "Dj+"
        .Cells(15 + (2 * numOfCriteria), 6).value = "Dj-"
        .Cells(15 + (2 * numOfCriteria), 7).value = "Dj"
        
        With .Range(.Cells(15 + (2 * numOfCriteria), 2), .Cells(15 + (2 * numOfCriteria), 7))
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
        
        ' Cykly pro pøidání podmínek Solveru
        Dim j As Long
        For j = 1 To numOfCriteria
            Dim Wj As Range
            Dim Wj_new As Range
            Dim Dj_neg As Range
            Dim Dj_pos As Range
            Dim Dj As Range
            
            ' Definice jednotlivých bunìk
            Set Wj = .Cells(4 + j, 4)    ' pùvodní váha Wj
            Set Wj_new = .Cells(10 + numOfCriteria + j, 4)       ' nová váha Wj'
            Set Dj_pos = .Cells(15 + (2 * numOfCriteria) + j, 5) ' buòka pro Dj+
            Set Dj_neg = .Cells(15 + (2 * numOfCriteria) + j, 6) ' buòka pro Dj-
            Set Dj = .Cells(15 + (2 * numOfCriteria) + j, 7)     ' buòka pro Dj
            
            ' Kopírování kritérií
            .Cells(15 + (2 * numOfCriteria) + j, 2).value = .Cells(4 + j, 2).value
            
            ' Kopírování vah
            .Cells(15 + (2 * numOfCriteria) + j, 3).value = .Cells(4 + j, 4).value
            
            ' Nastavení vzorce pro odchylky
            Dj.formula = "=" & Wj.Address & " - " & Wj_new.Address & " - " & Dj_neg.Address & " + " & Dj_pos.Address
            
            .Range(.Cells(15 + (2 * numOfCriteria) + j, 3), .Cells(15 + (2 * numOfCriteria) + j, 7)).NumberFormat = "0.0 %"
        Next j
        
        ' Konfigurace Solveru:
        SolverReset
        
        ' Podmínky Solveru:

        ' Jednotlivé váhy musí být menší nebo rovny 1 (100%)
        SolverAdd cellRef:=.Range(.Cells(11 + numOfCriteria, 4), .Cells(10 + (2 * numOfCriteria), 4)).Address, _
                    Relation:=1, _
                    FormulaText:="1"

        ' Celkový souèet vah musí být roven 1 (100%)
        SolverAdd cellRef:=.Cells(11 + (2 * numOfCriteria), 4), _
                    Relation:=2, _
                    FormulaText:="=1"
        
        ' Jednotlivé odchylky Dj+ a Dj- musí menší nebo rovny 1
        SolverAdd cellRef:=.Range(.Cells(15 + (2 * numOfCriteria) + 1, 5), .Cells(14 + (3 * numOfCriteria) + 1, 6)).Address, _
                    Relation:=1, _
                    FormulaText:="1"
        
        ' Celkové odchylky Dj musí rovny 0
        SolverAdd cellRef:=.Range(.Cells(15 + (2 * numOfCriteria) + 1, 7), .Cells(14 + (3 * numOfCriteria) + 1, 7)).Address, _
                    Relation:=2, _
                    FormulaText:="0"


        ' Podmínky pro zajištìní, že vybraná varianta má nejvyšší užitek
        Dim epsilon As Double
        epsilon = 0.000001
        For i = 1 To numOfCandidates
            If i <> selectedColumn Then
                SolverAdd cellRef:=.Cells(10 + numOfCriteria, 7 + numOfCandidates).Address, _
                    Relation:=3, _
                    FormulaText:=.Cells(13 + (2 * numOfCriteria), 4 + i).Address & "+" & epsilon
            End If
        Next i
        
        ' Nastavení vzorce pro sumu odchylek (spoleèné)
        .Cells(11 + numOfCriteria, 7 + numOfCandidates).formula = _
            "=SUM(" & .Range(.Cells(15 + (2 * numOfCriteria) + 1, 5), .Cells(14 + (3 * numOfCriteria) + 1, 6)).Address & ")"
        
        ' Formátování na jedno desetinné místo
        .Range(.Cells(11 + numOfCriteria, 7 + numOfCandidates), .Cells(11 + numOfCriteria, 7 + numOfCandidates)).NumberFormat = "0.0 %"
        
        If isMin Then
            ' Minimalizaèní funkce
            objective = "Minimalizace"
            
            ' Nadpis pro klíèovou funkci
            .Cells(11 + numOfCriteria, 6 + numOfCandidates).value = "Co nejmenší:"
            
            ' Minimalizaèní Solver
            SolverOk SetCell:=.Cells(11 + numOfCriteria, 7 + numOfCandidates).Address, _
                     MaxMinVal:=2, _
                     ValueOf:=0, _
                     ByChange:=.Range(.Cells(10 + numOfCriteria + 1, 4), .Cells(10 + (2 * numOfCriteria), 4)).Address & ";" & _
                                .Range(.Cells(15 + (2 * numOfCriteria) + 1, 5), .Cells(15 + (3 * numOfCriteria), 6)).Address, _
                     Engine:=2, EngineDesc:="Simplex LP"
                     
        Else
            objective = "Maximalizace"
            
            ' Nadpis pro klíèovou funkci
            .Cells(11 + numOfCriteria, 6 + numOfCandidates).value = "Co nejvìtší:"
            
            ' Maximalizaèní Solver
            SolverOk SetCell:=.Cells(11 + numOfCriteria, 7 + numOfCandidates).Address, _
                MaxMinVal:=1, _
                ValueOf:=0, _
                ByChange:=Range(.Cells(11 + numOfCriteria, 4), .Cells(10 + (2 * numOfCriteria), 4)).Address & ";" & _
                            .Range(.Cells(15 + (2 * numOfCriteria) + 1, 5), .Cells(15 + (3 * numOfCriteria), 6)).Address & ";" & _
                            .Range(.Cells(17 + (3 * numOfCriteria) + 1, 3), .Cells(17 + (4 * numOfCriteria), 4)).Address, _
                Engine:=2, EngineDesc:="Simplex LP"
                     
            ' Pøidání nadpisù pro horní omezení
            .Cells(17 + (3 * numOfCriteria), 3).value = "Yj+"
            .Cells(17 + (3 * numOfCriteria), 4).value = "Yj-"
            .Cells(17 + (3 * numOfCriteria), 5).value = "2Yj+"
            .Cells(17 + (3 * numOfCriteria), 6).value = "2Yj-"
            .Cells(17 + (3 * numOfCriteria), 7).value = "Yj"
            
            With .Range(.Cells(17 + (3 * numOfCriteria), 3), .Cells(17 + (3 * numOfCriteria), 7))
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
            
            For j = 1 To numOfCriteria
                Dim Yj_pos As Range
                Dim Yj_neg As Range
                Dim Yj_pos_2 As Range
                Dim Yj_neg_2 As Range
                Dim Yj As Range
                
                ' Definice jednotlivých bunìk
                Set Yj_pos = .Cells(17 + (3 * numOfCriteria) + j, 3)    ' buòka pro Yj+
                Set Yj_neg = .Cells(17 + (3 * numOfCriteria) + j, 4)    ' buòka pro Yj-
                Set Yj_pos_2 = .Cells(17 + (3 * numOfCriteria) + j, 5)  ' buòka pro 2*Yj+
                Set Yj_neg_2 = .Cells(17 + (3 * numOfCriteria) + j, 6)  ' buòka pro 2*Yj-
                Set Yj = .Cells(17 + (3 * numOfCriteria) + j, 7)        ' buòka pro Yj
                
                Yj_pos.value = 0
                Yj_neg.value = 0
                
                ' Nastavení vzorcù pro pomocné promìnné
                Yj_pos_2.formula = "= 2 * " & Yj_pos.Address
                Yj_neg_2.formula = "= 2 * " & Yj_neg.Address
                Yj.formula = "=" & Yj_pos.Address & " + " & Yj_neg.Address
    
            Next j
            
            ' Yj+ a Yj- jsou binární
            SolverAdd cellRef:=.Range(.Cells(17 + (3 * numOfCriteria) + 1, 3), .Cells(17 + (4 * numOfCriteria), 4)).Address, _
                      Relation:=4
            
            ' Souèet Yj+ a Yj- je vždy roven 1
            SolverAdd cellRef:=.Range(.Cells(17 + (3 * numOfCriteria) + 1, 7), .Cells(17 + (4 * numOfCriteria), 7)).Address, _
                      Relation:=2, _
                      FormulaText:="1"
            
            ' Dj+ je menší nebo rovno 2*Yj+
            SolverAdd cellRef:=.Range(.Cells(15 + (2 * numOfCriteria) + 1, 5), .Cells(15 + (3 * numOfCriteria), 5)).Address, _
                    Relation:=1, _
                    FormulaText:=.Range(.Cells(17 + (3 * numOfCriteria) + 1, 5), .Cells(17 + (4 * numOfCriteria), 5)).Address
            
            ' Dj- je menší nebo rovno 2*Yj-
            SolverAdd cellRef:=.Range(.Cells(15 + (2 * numOfCriteria) + 1, 6), .Cells(15 + (3 * numOfCriteria), 6)).Address, _
                      Relation:=1, _
                      FormulaText:=.Range(.Cells(17 + (3 * numOfCriteria) + 1, 6), .Cells(17 + (4 * numOfCriteria), 6)).Address

        End If

        ' Popisek tabulky pro minimalizaci odchylek
        .Range("A1:J2").Merge
        With .Range("A1")
            .value = objective & " zmìny vah pro stanovení varianty '" & selectedCandidate & "' kompromisní použitím metody " & metoda
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Italic = True
            .Font.Bold = True
            
            ' Vypnutí tuèného písma pro specifické èásti textu
            .Characters(InStr(.value, "zmìny vah pro stanovení varianty"), Len("zmìny vah pro stanovení varianty")).Font.Bold = False
            .Characters(InStr(.value, "použitím metody"), Len("použitím metody")).Font.Bold = False

        End With

        ' Úprava šíøky sloupcù (aplikuje se vždy)
        AdjustColumnWidth wsOutput, .Range(.Columns(2), .Columns(4 + numOfCandidates + 3))
    End With
    
    ' Pøepsání promìnné metoda na základì názvu listu
    If ws.name = "Metoda WSA" Then
            metoda = "WSA"
        Else
            metoda = "SAW"
    End If
    
    With ws
        ws.Unprotect "1234"
        If isMin Then
            ' Minimalizaèní funkce
            .Cells(11 + numOfCriteria, 6 + numOfCandidates).value = "Co nejmenší:"
            
            ' Nastavení vzorce pro sumu odchylek (nejmenší)
            .Cells(11 + numOfCriteria, 7 + numOfCandidates).formula = _
                "=SUM(" & metoda & "_MIN!" & .Range(.Cells(15 + (2 * numOfCriteria) + 1, 5), _
                    .Cells(14 + (3 * numOfCriteria) + 1, 6)).Address & ")"
                
        Else
            ' Maximalizaèní funkce
            .Cells(12 + numOfCriteria, 6 + numOfCandidates).value = "Co nejvìtší:"
        
            ' Nastavení vzorce pro sumu odchylek (nejvìtší)
            .Cells(12 + numOfCriteria, 7 + numOfCandidates).formula = _
                "=SUM(" & metoda & "_MAX!" & .Range(.Cells(15 + (2 * numOfCriteria) + 1, 5), _
                    .Cells(14 + (3 * numOfCriteria) + 1, 6)).Address & ")"
        End If
    End With
End Sub

'Funkce pøidávající ComboBox pro výbìr varianty
Function AddComboBox(ws As Worksheet, name As String, targetCell As Range, optionsRange As Range, macroName As String) As Variant
    Dim cbo As Shape
    Dim itemCount As Long
    Dim cell As Range
    Dim maxWidth As Double


    ' Nastavení minimální šíøky
    minWidth = 80
    
    ' Urèení poètu položek v rozsahu
    itemCount = WorksheetFunction.CountA(optionsRange)
    
    ' Najít nejširší prvek v seznamu
    maxWidth = 0
    For Each cell In optionsRange
        If cell.Width > maxWidth Then
            maxWidth = cell.Width
        End If
    Next cell

    ' Nastavení šíøky na maximální hodnotu nebo minimální šíøku
    If maxWidth < minWidth Then
        maxWidth = minWidth
    End If

    ' Odstranìní existujícího ComboBoxu, pokud existuje
    For Each cbo In ws.Shapes
        If cbo.name = "MyComboBox" Then
            cbo.Delete
            Exit For
        End If
    Next cbo

    ' Vytvoøení nového ComboBoxu
    Set cbo = ws.Shapes.AddFormControl(Type:=xlDropDown, Left:=targetCell.Left, Top:=targetCell.Top, Width:=maxWidth, Height:=targetCell.Height)

    ' Nastavení názvu ComboBoxu
    cbo.name = name

    ' Pøidání položek do ComboBoxu z rozsahu
    For Each cell In optionsRange
        cbo.ControlFormat.AddItem cell.value
    Next cell

    ' Zobrazení ComboBoxu
    cbo.Visible = True
    
    ' Nastavení poètu øádkù v rozbalovacím seznamu
    cbo.ControlFormat.DropDownLines = itemCount
    
    ' Pøiøazení makra k tlaèítku
    cbo.OnAction = macroName
End Function

' Procedura obslující ComboBox na listu, pokud se zmìní vybraná varianta
Sub newBestCandidate_Change(ws As Worksheet, newBestCandidateName As String)

    Dim keyValue As String
    Dim selectedCandidate As String
    Dim dominatedCandidates As String
    Dim cbo As Object
    Dim selectedColumn As Integer
    Dim selectedAddress As String
    Dim dominatedArray() As String
    Dim i, j As Integer
    
    ' Získání referencí na listy
    Set wsInput = ThisWorkbook.Sheets("Vstupní data")
    
    numOfCriteria = wsInput.Range("C2").value
    numOfCandidates = wsInput.Range("F2").value
    
    ' Naètení seznamu dominovaných variant z buòky (oddìlené èárkami)
    dominatedCandidates = ws.Cells(6 + numOfCriteria, 7 + numOfCandidates).value
    
    ' Pøeveïte text do pole oddìleného èárkami
    dominatedArray = Split(dominatedCandidates, ", ")

    ' Získání hodnoty vybrané v ComboBoxu
    For Each cbo In ws.DropDowns
        If cbo.name = newBestCandidateName Then
        
            ' Kontrola, zda je ComboBox obsahuje nìjaké varianty
            If cbo.ListCount > 0 Then
            
                ' Kontrola, zda je vybrána nìjaká varianta
                If cbo.ListIndex <> 0 Then
                
                    ' Získání hodnoty vybrané v ComboBoxu a její pøevod na øetìzec
                    selectedCandidate = CStr(cbo.List(cbo.ListIndex))
                    
                    ' Zkontrolujeme, zda je vybraná varianta mezi dominovanými
                    For i = LBound(dominatedArray) To UBound(dominatedArray)
                        If Trim(dominatedArray(i)) = selectedCandidate Then
                            ' Pokud je varianta dominovaná, zobrazíme upozornìní
                            MsgBox "Vybraná varianta je dominovaná a žádnou zmìnou vah nemùže být stanovena kompromisní variantou.", vbOKOnly, "Upozornìní"
                        End If
                    Next i
                
                    With ws
                        .Unprotect "1234"
                        
                        ' Nastavení popisku pro klíèovou funkci
                        With .Range(.Cells(10 + numOfCriteria, 5 + numOfCandidates), .Cells(10 + numOfCriteria, 6 + numOfCandidates))
                            .Merge
                            .value = "Užitek vybrané varianty:"
                            .HorizontalAlignment = xlCenter
                            .WrapText = False
                        End With
                        
                        ' Výpoèet klíèové funkce: použití Match pro získání hodnoty varianty
                        selectedColumn = Application.WorksheetFunction.Match(selectedCandidate, _
                            .Range(.Cells(12 + (2 * numOfCriteria), 5), .Cells(12 + (2 * numOfCriteria), 4 + numOfCandidates)), 0)
                        
                        selectedAddress = .Cells(13 + (2 * numOfCriteria), 4 + selectedColumn).Address
                        .Cells(10 + numOfCriteria, 7 + numOfCandidates).formula = "=" & selectedAddress
                    
                        ' Formátování na tøi desetinná místa
                        .Cells(10 + numOfCriteria, 7 + numOfCandidates).NumberFormat = "0.000"
                        
                        ' Obnovení pùvodních vah kritérií
                        For j = 1 To numOfCriteria
                            .Cells(10 + numOfCriteria + j, 4).value = wsInput.Cells(4 + j, 4).value
                        Next j
                        
                        .Range(.Cells(11 + numOfCriteria, 7 + numOfCandidates), .Cells(12 + numOfCriteria, 7 + numOfCandidates)).Clear
                        
                        .Calculate
                        
                        .Protect "1234"
                        Exit Sub
                    End With
                End If
            Else
                MsgBox "Není k dispozici žádná varianta k výbìru.", vbExclamation
                Exit Sub
            End If
        End If
    Next cbo
End Sub


