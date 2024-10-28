Attribute VB_Name = "Module5"
Dim ws As Worksheet
Dim wsInput As Worksheet
Dim numOfCriteria As Integer
Dim numOfCandidates As Integer
Dim selectedVariant As String
Dim cbo As Object

' Makro pro spouštìní Solveru
Sub M5_Solver(ws As Worksheet, cboName As String)

    ' Získání referencí na listy
    Set ws = ws
    Set wsInput = ThisWorkbook.Sheets("Vstupní data")
    
    ' Získání hodnoty vybrané v ComboBoxu
    For Each cbo In ws.DropDowns
        If cbo.name = cboName Then
        
            ' Kontrola, zda je ComboBox obsahuje nìjaké varianty
            If cbo.ListCount > 0 Then
            
                ' Kontrola, zda je vybrána nìjaká varianta
                If cbo.ListIndex = 0 Then
                    MsgBox "Zvolte, prosím, požadované kompromisní øešení.", vbExclamation
                    Exit Sub
                End If
            Else
                MsgBox "Není k dispozici žádná varianta k výbìru.", vbExclamation
                Exit Sub
            End If
        End If
    Next cbo

    numOfCriteria = wsInput.Range("C2").value
    numOfCandidates = wsInput.Range("F2").value
    
    With ws
        .Unprotect "1234"
    
    ' Konfigurace Solveru:
        SolverReset
        
        SolverOk SetCell:=.Cells(11 + numOfCriteria, 7 + numOfCandidates).Address, _
                MaxMinVal:=2, _
                ValueOf:=0, _
                ByChange:=.Range(.Cells(11 + numOfCriteria, 4), .Cells(11 + (2 * numOfCriteria), 4)).Address, _
                Engine:=1, EngineDesc:="GRG Nonlinear"
                'Engine:=3, EngineDesc:="Evolutionary"
                'Engine:=2, EngineDesc:="Simplex LP"
                
        ' Nastavení maximálního èasu pro všechny metody na 180 sekund a využití více poèáteèních bodù gradientní metody
        SolverOptions MaxTime:=180, MultiStart:=True
            
    ' Podmínky Solveru:

        ' Jednotlivé váhy musí být menší nebo rovny 1 (100%)
        SolverAdd cellRef:=.Range(.Cells(11 + numOfCriteria, 4), .Cells(10 + (2 * numOfCriteria), 4)).Address, _
                    Relation:=1, _
                    FormulaText:="1"
                    
        ' Celkový souèet vah je roven souètu (100%)
        SolverAdd cellRef:=.Cells(11 + (2 * numOfCriteria), 4).Address, Relation:=2, _
                    FormulaText:="=SUM(" & .Range(.Cells(11 + numOfCriteria, 4), _
                                                  .Cells(10 + (2 * numOfCriteria), 4)).Address & ")"

        ' Celkový souèet vah musí být roven 1 (100%)
        SolverAdd cellRef:=.Cells(11 + (2 * numOfCriteria), 4), _
                    Relation:=2, _
                    FormulaText:="=1"
        
        ' Požadované øešení musí být stejné jako øešení metody
        SolverAdd cellRef:=.Cells(10 + numOfCriteria, 7 + numOfCandidates), _
                    Relation:=2, _
                    FormulaText:="=1"

        SolverSolve
        
        AdjustColumnWidth ws, 7 + numOfCandidates
        
        .Protect "1234"

    End With
End Sub

Function AddComboBox(ws As Worksheet, name As String, targetCell As Range, optionsRange As Range, macroName As String) As Variant
    Dim cbo As Shape
    Dim itemCount As Long
    Dim cell As Range
    Dim maxWidth As Double

    ' Urèení poètu položek v rozsahu
    itemCount = WorksheetFunction.CountA(optionsRange)
    
    ' Najít nejširší prvek v seznamu
    maxWidth = 0
    For Each cell In optionsRange
        If cell.Width > maxWidth Then
            maxWidth = cell.Width
        End If
    Next cell

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

' Procedura obnovující pùvodní váhy kritérií, pokud se zmìní hodnota ComboBoxu
Sub newBestCandidate_Change(ws As Worksheet, newBestCandidateName As String)

    Dim keyValue As String
    
    ' Získání referencí na listy
    Set ws = ws
    Set wsInput = ThisWorkbook.Sheets("Vstupní data")
    
    numOfCriteria = wsInput.Range("C2").value
    numOfCandidates = wsInput.Range("F2").value
    
    ' Získání hodnoty vybrané v ComboBoxu
    For Each cbo In ws.DropDowns
        If cbo.name = newBestCandidateName Then
        
            ' Kontrola, zda je ComboBox obsahuje nìjaké varianty
            If cbo.ListCount > 0 Then
            
                ' Kontrola, zda je vybrána nìjaká varianta
                If cbo.ListIndex <> 0 Then
                
                    ' Získání hodnoty vybrané v ComboBoxu a její pøevod na øetìzec
                    selectedVariant = CStr(cbo.List(cbo.ListIndex))
                
                    With ws
                        .Unprotect "1234"
                        .Range(ws.Cells(11 + numOfCriteria, 4), .Cells(10 + (2 * numOfCriteria), 4)).formula _
                            = wsInput.Range(wsInput.Cells(5, 4), wsInput.Cells(4 + numOfCriteria, 4)).value
                        
                        ' Nastavení popisku pro klíèovou funkci
                        .Cells(10 + numOfCriteria, 6 + numOfCandidates).value = "Klíèová funkce"
                        
                        ' Výpoèet klíèové funkce: požadované øešení musí být stejné jako øešení metody
                        keyValue = .Cells(12 + (2 * numOfCriteria), 7 + numOfCandidates).Address(True, True)
                        .Cells(10 + numOfCriteria, 7 + numOfCandidates).formula = "=IF(" & keyValue & "=""" & selectedVariant & """,1,0)"
                
                        ' Minimalizaèní funkce
                        .Cells(11 + numOfCriteria, 6 + numOfCandidates).value = "Co nejmenší:"
                        
                        .Cells(11 + numOfCriteria, 7 + numOfCandidates).Formula2 = _
                                    "=SUM(ABS(" & .Range(.Cells(5, 4), .Cells(4 + numOfCriteria, 4)).Address & _
                                    " - " & .Range(.Cells(11 + numOfCriteria, 4), .Cells(10 + (2 * numOfCriteria), 4)).Address & "))"
                        
                        .Cells(11 + numOfCriteria, 7 + numOfCandidates).NumberFormat = "0.0 %"
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
