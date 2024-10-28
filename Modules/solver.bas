Attribute VB_Name = "Module5"
Dim ws As Worksheet
Dim wsInput As Worksheet
Dim numOfCriteria As Integer
Dim numOfCandidates As Integer
Dim selectedVariant As String
Dim cbo As Object

' Makro pro spou�t�n� Solveru
Sub M5_Solver(ws As Worksheet, cboName As String)

    ' Z�sk�n� referenc� na listy
    Set ws = ws
    Set wsInput = ThisWorkbook.Sheets("Vstupn� data")
    
    ' Z�sk�n� hodnoty vybran� v ComboBoxu
    For Each cbo In ws.DropDowns
        If cbo.name = cboName Then
        
            ' Kontrola, zda je ComboBox obsahuje n�jak� varianty
            If cbo.ListCount > 0 Then
            
                ' Kontrola, zda je vybr�na n�jak� varianta
                If cbo.ListIndex = 0 Then
                    MsgBox "Zvolte, pros�m, po�adovan� kompromisn� �e�en�.", vbExclamation
                    Exit Sub
                End If
            Else
                MsgBox "Nen� k dispozici ��dn� varianta k v�b�ru.", vbExclamation
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
                
        ' Nastaven� maxim�ln�ho �asu pro v�echny metody na 180 sekund a vyu�it� v�ce po��te�n�ch bod� gradientn� metody
        SolverOptions MaxTime:=180, MultiStart:=True
            
    ' Podm�nky Solveru:

        ' Jednotliv� v�hy mus� b�t men�� nebo rovny 1 (100%)
        SolverAdd cellRef:=.Range(.Cells(11 + numOfCriteria, 4), .Cells(10 + (2 * numOfCriteria), 4)).Address, _
                    Relation:=1, _
                    FormulaText:="1"
                    
        ' Celkov� sou�et vah je roven sou�tu (100%)
        SolverAdd cellRef:=.Cells(11 + (2 * numOfCriteria), 4).Address, Relation:=2, _
                    FormulaText:="=SUM(" & .Range(.Cells(11 + numOfCriteria, 4), _
                                                  .Cells(10 + (2 * numOfCriteria), 4)).Address & ")"

        ' Celkov� sou�et vah mus� b�t roven 1 (100%)
        SolverAdd cellRef:=.Cells(11 + (2 * numOfCriteria), 4), _
                    Relation:=2, _
                    FormulaText:="=1"
        
        ' Po�adovan� �e�en� mus� b�t stejn� jako �e�en� metody
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

    ' Ur�en� po�tu polo�ek v rozsahu
    itemCount = WorksheetFunction.CountA(optionsRange)
    
    ' Naj�t nej�ir�� prvek v seznamu
    maxWidth = 0
    For Each cell In optionsRange
        If cell.Width > maxWidth Then
            maxWidth = cell.Width
        End If
    Next cell

    ' Odstran�n� existuj�c�ho ComboBoxu, pokud existuje
    For Each cbo In ws.Shapes
        If cbo.name = "MyComboBox" Then
            cbo.Delete
            Exit For
        End If
    Next cbo

    ' Vytvo�en� nov�ho ComboBoxu
    Set cbo = ws.Shapes.AddFormControl(Type:=xlDropDown, Left:=targetCell.Left, Top:=targetCell.Top, Width:=maxWidth, Height:=targetCell.Height)

    ' Nastaven� n�zvu ComboBoxu
    cbo.name = name

    ' P�id�n� polo�ek do ComboBoxu z rozsahu
    For Each cell In optionsRange
        cbo.ControlFormat.AddItem cell.value
    Next cell

    ' Zobrazen� ComboBoxu
    cbo.Visible = True
    
    ' Nastaven� po�tu ��dk� v rozbalovac�m seznamu
    cbo.ControlFormat.DropDownLines = itemCount
    
    ' P�i�azen� makra k tla��tku
    cbo.OnAction = macroName
End Function

' Procedura obnovuj�c� p�vodn� v�hy krit�ri�, pokud se zm�n� hodnota ComboBoxu
Sub newBestCandidate_Change(ws As Worksheet, newBestCandidateName As String)

    Dim keyValue As String
    
    ' Z�sk�n� referenc� na listy
    Set ws = ws
    Set wsInput = ThisWorkbook.Sheets("Vstupn� data")
    
    numOfCriteria = wsInput.Range("C2").value
    numOfCandidates = wsInput.Range("F2").value
    
    ' Z�sk�n� hodnoty vybran� v ComboBoxu
    For Each cbo In ws.DropDowns
        If cbo.name = newBestCandidateName Then
        
            ' Kontrola, zda je ComboBox obsahuje n�jak� varianty
            If cbo.ListCount > 0 Then
            
                ' Kontrola, zda je vybr�na n�jak� varianta
                If cbo.ListIndex <> 0 Then
                
                    ' Z�sk�n� hodnoty vybran� v ComboBoxu a jej� p�evod na �et�zec
                    selectedVariant = CStr(cbo.List(cbo.ListIndex))
                
                    With ws
                        .Unprotect "1234"
                        .Range(ws.Cells(11 + numOfCriteria, 4), .Cells(10 + (2 * numOfCriteria), 4)).formula _
                            = wsInput.Range(wsInput.Cells(5, 4), wsInput.Cells(4 + numOfCriteria, 4)).value
                        
                        ' Nastaven� popisku pro kl��ovou funkci
                        .Cells(10 + numOfCriteria, 6 + numOfCandidates).value = "Kl��ov� funkce"
                        
                        ' V�po�et kl��ov� funkce: po�adovan� �e�en� mus� b�t stejn� jako �e�en� metody
                        keyValue = .Cells(12 + (2 * numOfCriteria), 7 + numOfCandidates).Address(True, True)
                        .Cells(10 + numOfCriteria, 7 + numOfCandidates).formula = "=IF(" & keyValue & "=""" & selectedVariant & """,1,0)"
                
                        ' Minimaliza�n� funkce
                        .Cells(11 + numOfCriteria, 6 + numOfCandidates).value = "Co nejmen��:"
                        
                        .Cells(11 + numOfCriteria, 7 + numOfCandidates).Formula2 = _
                                    "=SUM(ABS(" & .Range(.Cells(5, 4), .Cells(4 + numOfCriteria, 4)).Address & _
                                    " - " & .Range(.Cells(11 + numOfCriteria, 4), .Cells(10 + (2 * numOfCriteria), 4)).Address & "))"
                        
                        .Cells(11 + numOfCriteria, 7 + numOfCandidates).NumberFormat = "0.0 %"
                        .Protect "1234"
                        Exit Sub
                    End With
                End If
            Else
                MsgBox "Nen� k dispozici ��dn� varianta k v�b�ru.", vbExclamation
                Exit Sub
            End If
        End If
    Next cbo
End Sub
