Attribute VB_Name = "Module5"
Dim ws As Worksheet
Dim wsInput As Worksheet
Dim numOfCriteria As Integer
Dim numOfCandidates As Integer
Dim cbo As Object

Dim selectedCandidate As String
Dim selectedColumn As Integer

' Makro pro spou�t�n� Solveru
Sub M5_Solver(ws As Worksheet, cboName As String)
    Dim metoda As String
    Dim minSheetName As String
    Dim maxSheetName As String
    Dim wsMin As Worksheet
    Dim wsMax As Worksheet

    ' Z�sk�n� referenc� na listy
    Set ws = ws
    
    ' Z�sk�n� hodnoty vybran� v ComboBoxu
    For Each cbo In ws.DropDowns
        If cbo.name = cboName Then
        
            ' Kontrola, zda je ComboBox obsahuje n�jak� varianty
            If cbo.ListCount > 0 Then
            
                ' Kontrola, zda je vybr�na n�jak� varianta
                If cbo.ListIndex = 0 Then
                    MsgBox "Zvolte, pros�m, testovanou variantu.", vbExclamation
                    Exit Sub
                Else
                    ' Z�sk�n� hodnoty vybran� v ComboBoxu a jej� p�evod na �et�zec
                    selectedCandidate = CStr(cbo.List(cbo.ListIndex))
                    
                    ' Z�sk�n� referenc� na listy
                    Set wsInput = ThisWorkbook.Sheets("Vstupn� data")
                    
                    numOfCriteria = wsInput.Range("C2").value
                    numOfCandidates = wsInput.Range("F2").value
                    
                    ' V�po�et kl��ov� funkce: pou�it� Match pro z�sk�n� hodnoty varianty
                    selectedColumn = Application.WorksheetFunction.Match(selectedCandidate, _
                        ws.Range(ws.Cells(12 + (2 * numOfCriteria), 5), ws.Cells(12 + (2 * numOfCriteria), 4 + numOfCandidates)), 0)
                End If
            Else
                MsgBox "Nen� k dispozici ��dn� varianta k v�b�ru.", vbExclamation
                Exit Sub
            End If
        End If
    Next cbo

    ' Ur�en� metody na z�klad� n�zvu listu
    Select Case ws.name
        Case "Metoda WSA"
            metoda = "WSA"
        Case "Metoda bazick� varianty"
            metoda = "SAW"
        Case Else
            MsgBox "Nezn�m� metoda! Ujist�te se, �e vol�te Solver z platn�ho listu.", vbCritical
            Exit Sub
    End Select

    ' Ur�en� n�zv� list� podle metody
    minSheetName = metoda & "_MIN"
    maxSheetName = metoda & "_MAX"

    ' Vytvo�en� a �prava listu minimalizace
    CreateAndConfigureSheet ws, selectedCandidate, selectedColumn, minSheetName, True
    
    ' Spu�t�n� Solveru pro minimalizaci
    ThisWorkbook.Sheets(minSheetName).Activate
    SolverSolve UserFinish:=True
    SolverFinish KeepFinal:=1

    ' Vytvo�en� a �prava listu maximalizace
    CreateAndConfigureSheet ws, selectedCandidate, selectedColumn, maxSheetName, False
    
    ' Spu�t�n� Solveru pro maximalizaci
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

    ' Ozn�men� dokon�en�
    MsgBox "V�sledky jsou dostupn� na listech: " & minSheetName & " a " & maxSheetName, vbInformation
    
End Sub

Sub CreateAndConfigureSheet(ws As Worksheet, selectedCandidate As String, selectedColumn As Integer, sheetName As String, isMin As Boolean)
    Dim wsOutput As Worksheet
    Dim objective As String
    Dim metoda As String

    ' Odstran�n� existuj�c�ho listu, pokud ji� existuje
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets(sheetName)
    If Not wsOutput Is Nothing Then
        Application.DisplayAlerts = False
        wsOutput.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' Vytvo�en� nov�ho listu
    Set wsOutput = ThisWorkbook.Sheets.Add
    wsOutput.name = sheetName
    
    Set wsInput = ThisWorkbook.Sheets("Vstupn� data")
    Set ws = ws
    numOfCriteria = wsInput.Range("C2").value
    numOfCandidates = wsInput.Range("F2").value
    
    ' Ur�en� metody na z�klad� n�zvu listu
    If ws.name = "Metoda WSA" Then
            metoda = "WSA"
        Else
            metoda = "bazick� varianty"
    End If
    
    ' Kop�rov�n� obsahu
    With wsOutput
        ws.Range(ws.Cells(3, 1), ws.Cells(13 + (2 * numOfCriteria), 4 + numOfCandidates)).Copy Destination:=.Cells(3, 1)
        
        If ws.name = "Metoda bazick� varianty" Then
            ws.Range(ws.Cells(4, 4 + numOfCandidates + 1), ws.Cells(4 + numOfCriteria, 4 + numOfCandidates + 1)).Copy Destination:=.Cells(4, 4 + numOfCandidates + 1)
        End If
        
        ' Kop�rov�n� po�ad�
        ws.Range(ws.Cells(16 + numOfCriteria + 1, 6 + numOfCandidates), ws.Cells(16 + numOfCriteria + numOfCandidates, 7 + numOfCandidates)).Copy _
            Destination:=.Cells(14 + numOfCriteria, 6 + numOfCandidates)
            
        ' Nastaven� popisku pro kl��ovou funkci
        With .Range(.Cells(10 + numOfCriteria, 5 + numOfCandidates), .Cells(10 + numOfCriteria, 6 + numOfCandidates))
            .Merge
            .value = "U�itek vybran� varianty:"
            .HorizontalAlignment = xlCenter
            .WrapText = False
            .Font.Bold = True
        End With

        If .Columns(5 + numOfCandidates).ColumnWidth < 12 Then
            .Columns(5 + numOfCandidates).ColumnWidth = 12
        End If
        
        ' P�ed�n� vzorce bu�ky s odkazem na vybranou variantu
        .Cells(10 + numOfCriteria, 7 + numOfCandidates).formula = ws.Cells(10 + numOfCriteria, 7 + numOfCandidates).formula
        
        ' Nastaven� nadpis� v tabulce
        .Cells(15 + (2 * numOfCriteria), 2).value = "Krit�rium"
        .Cells(15 + (2 * numOfCriteria), 3).value = "Wj"
        .Cells(15 + (2 * numOfCriteria), 4).value = "Wj'"
        .Cells(15 + (2 * numOfCriteria), 5).value = "Dj+"
        .Cells(15 + (2 * numOfCriteria), 6).value = "Dj-"
        .Cells(15 + (2 * numOfCriteria), 7).value = "Dj"
        
        With .Range(.Cells(15 + (2 * numOfCriteria), 2), .Cells(15 + (2 * numOfCriteria), 7))
            ' Tu�n� a zarovn�n� na st�ed
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
    
            ' Nastaven� ohrani�en�
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
        
        ' Cykly pro p�id�n� podm�nek Solveru
        Dim j As Long
        For j = 1 To numOfCriteria
            Dim Wj As Range
            Dim Wj_new As Range
            Dim Dj_neg As Range
            Dim Dj_pos As Range
            Dim Dj As Range
            
            ' Definice jednotliv�ch bun�k
            Set Wj = .Cells(4 + j, 4)    ' p�vodn� v�ha Wj
            Set Wj_new = .Cells(10 + numOfCriteria + j, 4)       ' nov� v�ha Wj'
            Set Dj_pos = .Cells(15 + (2 * numOfCriteria) + j, 5) ' bu�ka pro Dj+
            Set Dj_neg = .Cells(15 + (2 * numOfCriteria) + j, 6) ' bu�ka pro Dj-
            Set Dj = .Cells(15 + (2 * numOfCriteria) + j, 7)     ' bu�ka pro Dj
            
            ' Kop�rov�n� krit�ri�
            .Cells(15 + (2 * numOfCriteria) + j, 2).value = .Cells(4 + j, 2).value
            
            ' Kop�rov�n� vah
            .Cells(15 + (2 * numOfCriteria) + j, 3).value = .Cells(4 + j, 4).value
            
            ' Nastaven� vzorce pro odchylky
            Dj.formula = "=" & Wj.Address & " - " & Wj_new.Address & " - " & Dj_neg.Address & " + " & Dj_pos.Address
            
            .Range(.Cells(15 + (2 * numOfCriteria) + j, 3), .Cells(15 + (2 * numOfCriteria) + j, 7)).NumberFormat = "0.0 %"
        Next j
        
        ' Konfigurace Solveru:
        SolverReset
        
        ' Podm�nky Solveru:

        ' Jednotliv� v�hy mus� b�t men�� nebo rovny 1 (100%)
        SolverAdd cellRef:=.Range(.Cells(11 + numOfCriteria, 4), .Cells(10 + (2 * numOfCriteria), 4)).Address, _
                    Relation:=1, _
                    FormulaText:="1"

        ' Celkov� sou�et vah mus� b�t roven 1 (100%)
        SolverAdd cellRef:=.Cells(11 + (2 * numOfCriteria), 4), _
                    Relation:=2, _
                    FormulaText:="=1"
        
        ' Jednotliv� odchylky Dj+ a Dj- mus� men�� nebo rovny 1
        SolverAdd cellRef:=.Range(.Cells(15 + (2 * numOfCriteria) + 1, 5), .Cells(14 + (3 * numOfCriteria) + 1, 6)).Address, _
                    Relation:=1, _
                    FormulaText:="1"
        
        ' Celkov� odchylky Dj mus� rovny 0
        SolverAdd cellRef:=.Range(.Cells(15 + (2 * numOfCriteria) + 1, 7), .Cells(14 + (3 * numOfCriteria) + 1, 7)).Address, _
                    Relation:=2, _
                    FormulaText:="0"


        ' Podm�nky pro zaji�t�n�, �e vybran� varianta m� nejvy��� u�itek
        Dim epsilon As Double
        epsilon = 0.000001
        For i = 1 To numOfCandidates
            If i <> selectedColumn Then
                SolverAdd cellRef:=.Cells(10 + numOfCriteria, 7 + numOfCandidates).Address, _
                    Relation:=3, _
                    FormulaText:=.Cells(13 + (2 * numOfCriteria), 4 + i).Address & "+" & epsilon
            End If
        Next i
        
        ' Nastaven� vzorce pro sumu odchylek (spole�n�)
        .Cells(11 + numOfCriteria, 7 + numOfCandidates).formula = _
            "=SUM(" & .Range(.Cells(15 + (2 * numOfCriteria) + 1, 5), .Cells(14 + (3 * numOfCriteria) + 1, 6)).Address & ")"
        
        ' Form�tov�n� na jedno desetinn� m�sto
        .Range(.Cells(11 + numOfCriteria, 7 + numOfCandidates), .Cells(11 + numOfCriteria, 7 + numOfCandidates)).NumberFormat = "0.0 %"
        
        If isMin Then
            ' Minimaliza�n� funkce
            objective = "Minimalizace"
            
            ' Nadpis pro kl��ovou funkci
            .Cells(11 + numOfCriteria, 6 + numOfCandidates).value = "Co nejmen��:"
            
            ' Minimaliza�n� Solver
            SolverOk SetCell:=.Cells(11 + numOfCriteria, 7 + numOfCandidates).Address, _
                     MaxMinVal:=2, _
                     ValueOf:=0, _
                     ByChange:=.Range(.Cells(10 + numOfCriteria + 1, 4), .Cells(10 + (2 * numOfCriteria), 4)).Address & ";" & _
                                .Range(.Cells(15 + (2 * numOfCriteria) + 1, 5), .Cells(15 + (3 * numOfCriteria), 6)).Address, _
                     Engine:=2, EngineDesc:="Simplex LP"
                     
        Else
            objective = "Maximalizace"
            
            ' Nadpis pro kl��ovou funkci
            .Cells(11 + numOfCriteria, 6 + numOfCandidates).value = "Co nejv�t��:"
            
            ' Maximaliza�n� Solver
            SolverOk SetCell:=.Cells(11 + numOfCriteria, 7 + numOfCandidates).Address, _
                MaxMinVal:=1, _
                ValueOf:=0, _
                ByChange:=Range(.Cells(11 + numOfCriteria, 4), .Cells(10 + (2 * numOfCriteria), 4)).Address & ";" & _
                            .Range(.Cells(15 + (2 * numOfCriteria) + 1, 5), .Cells(15 + (3 * numOfCriteria), 6)).Address & ";" & _
                            .Range(.Cells(17 + (3 * numOfCriteria) + 1, 3), .Cells(17 + (4 * numOfCriteria), 4)).Address, _
                Engine:=2, EngineDesc:="Simplex LP"
                     
            ' P�id�n� nadpis� pro horn� omezen�
            .Cells(17 + (3 * numOfCriteria), 3).value = "Yj+"
            .Cells(17 + (3 * numOfCriteria), 4).value = "Yj-"
            .Cells(17 + (3 * numOfCriteria), 5).value = "2Yj+"
            .Cells(17 + (3 * numOfCriteria), 6).value = "2Yj-"
            .Cells(17 + (3 * numOfCriteria), 7).value = "Yj"
            
            With .Range(.Cells(17 + (3 * numOfCriteria), 3), .Cells(17 + (3 * numOfCriteria), 7))
                ' Tu�n� a zarovn�n� na st�ed
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
        
                ' Nastaven� ohrani�en�
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
                
                ' Definice jednotliv�ch bun�k
                Set Yj_pos = .Cells(17 + (3 * numOfCriteria) + j, 3)    ' bu�ka pro Yj+
                Set Yj_neg = .Cells(17 + (3 * numOfCriteria) + j, 4)    ' bu�ka pro Yj-
                Set Yj_pos_2 = .Cells(17 + (3 * numOfCriteria) + j, 5)  ' bu�ka pro 2*Yj+
                Set Yj_neg_2 = .Cells(17 + (3 * numOfCriteria) + j, 6)  ' bu�ka pro 2*Yj-
                Set Yj = .Cells(17 + (3 * numOfCriteria) + j, 7)        ' bu�ka pro Yj
                
                Yj_pos.value = 0
                Yj_neg.value = 0
                
                ' Nastaven� vzorc� pro pomocn� prom�nn�
                Yj_pos_2.formula = "= 2 * " & Yj_pos.Address
                Yj_neg_2.formula = "= 2 * " & Yj_neg.Address
                Yj.formula = "=" & Yj_pos.Address & " + " & Yj_neg.Address
    
            Next j
            
            ' Yj+ a Yj- jsou bin�rn�
            SolverAdd cellRef:=.Range(.Cells(17 + (3 * numOfCriteria) + 1, 3), .Cells(17 + (4 * numOfCriteria), 4)).Address, _
                      Relation:=4
            
            ' Sou�et Yj+ a Yj- je v�dy roven 1
            SolverAdd cellRef:=.Range(.Cells(17 + (3 * numOfCriteria) + 1, 7), .Cells(17 + (4 * numOfCriteria), 7)).Address, _
                      Relation:=2, _
                      FormulaText:="1"
            
            ' Dj+ je men�� nebo rovno 2*Yj+
            SolverAdd cellRef:=.Range(.Cells(15 + (2 * numOfCriteria) + 1, 5), .Cells(15 + (3 * numOfCriteria), 5)).Address, _
                    Relation:=1, _
                    FormulaText:=.Range(.Cells(17 + (3 * numOfCriteria) + 1, 5), .Cells(17 + (4 * numOfCriteria), 5)).Address
            
            ' Dj- je men�� nebo rovno 2*Yj-
            SolverAdd cellRef:=.Range(.Cells(15 + (2 * numOfCriteria) + 1, 6), .Cells(15 + (3 * numOfCriteria), 6)).Address, _
                      Relation:=1, _
                      FormulaText:=.Range(.Cells(17 + (3 * numOfCriteria) + 1, 6), .Cells(17 + (4 * numOfCriteria), 6)).Address

        End If

        ' Popisek tabulky pro minimalizaci odchylek
        .Range("A1:J2").Merge
        With .Range("A1")
            .value = objective & " zm�ny vah pro stanoven� varianty '" & selectedCandidate & "' kompromisn� pou�it�m metody " & metoda
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Italic = True
            .Font.Bold = True
            
            ' Vypnut� tu�n�ho p�sma pro specifick� ��sti textu
            .Characters(InStr(.value, "zm�ny vah pro stanoven� varianty"), Len("zm�ny vah pro stanoven� varianty")).Font.Bold = False
            .Characters(InStr(.value, "pou�it�m metody"), Len("pou�it�m metody")).Font.Bold = False

        End With

        ' �prava ���ky sloupc� (aplikuje se v�dy)
        AdjustColumnWidth wsOutput, .Range(.Columns(2), .Columns(4 + numOfCandidates + 3))
    End With
    
    ' P�eps�n� prom�nn� metoda na z�klad� n�zvu listu
    If ws.name = "Metoda WSA" Then
            metoda = "WSA"
        Else
            metoda = "SAW"
    End If
    
    With ws
        ws.Unprotect "1234"
        If isMin Then
            ' Minimaliza�n� funkce
            .Cells(11 + numOfCriteria, 6 + numOfCandidates).value = "Co nejmen��:"
            
            ' Nastaven� vzorce pro sumu odchylek (nejmen��)
            .Cells(11 + numOfCriteria, 7 + numOfCandidates).formula = _
                "=SUM(" & metoda & "_MIN!" & .Range(.Cells(15 + (2 * numOfCriteria) + 1, 5), _
                    .Cells(14 + (3 * numOfCriteria) + 1, 6)).Address & ")"
                
        Else
            ' Maximaliza�n� funkce
            .Cells(12 + numOfCriteria, 6 + numOfCandidates).value = "Co nejv�t��:"
        
            ' Nastaven� vzorce pro sumu odchylek (nejv�t��)
            .Cells(12 + numOfCriteria, 7 + numOfCandidates).formula = _
                "=SUM(" & metoda & "_MAX!" & .Range(.Cells(15 + (2 * numOfCriteria) + 1, 5), _
                    .Cells(14 + (3 * numOfCriteria) + 1, 6)).Address & ")"
        End If
    End With
End Sub

'Funkce p�id�vaj�c� ComboBox pro v�b�r varianty
Function AddComboBox(ws As Worksheet, name As String, targetCell As Range, optionsRange As Range, macroName As String) As Variant
    Dim cbo As Shape
    Dim itemCount As Long
    Dim cell As Range
    Dim maxWidth As Double


    ' Nastaven� minim�ln� ���ky
    minWidth = 80
    
    ' Ur�en� po�tu polo�ek v rozsahu
    itemCount = WorksheetFunction.CountA(optionsRange)
    
    ' Naj�t nej�ir�� prvek v seznamu
    maxWidth = 0
    For Each cell In optionsRange
        If cell.Width > maxWidth Then
            maxWidth = cell.Width
        End If
    Next cell

    ' Nastaven� ���ky na maxim�ln� hodnotu nebo minim�ln� ���ku
    If maxWidth < minWidth Then
        maxWidth = minWidth
    End If

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

' Procedura obsluj�c� ComboBox na listu, pokud se zm�n� vybran� varianta
Sub newBestCandidate_Change(ws As Worksheet, newBestCandidateName As String)

    Dim keyValue As String
    Dim selectedCandidate As String
    Dim dominatedCandidates As String
    Dim cbo As Object
    Dim selectedColumn As Integer
    Dim selectedAddress As String
    Dim dominatedArray() As String
    Dim i, j As Integer
    
    ' Z�sk�n� referenc� na listy
    Set wsInput = ThisWorkbook.Sheets("Vstupn� data")
    
    numOfCriteria = wsInput.Range("C2").value
    numOfCandidates = wsInput.Range("F2").value
    
    ' Na�ten� seznamu dominovan�ch variant z bu�ky (odd�len� ��rkami)
    dominatedCandidates = ws.Cells(6 + numOfCriteria, 7 + numOfCandidates).value
    
    ' P�eve�te text do pole odd�len�ho ��rkami
    dominatedArray = Split(dominatedCandidates, ", ")

    ' Z�sk�n� hodnoty vybran� v ComboBoxu
    For Each cbo In ws.DropDowns
        If cbo.name = newBestCandidateName Then
        
            ' Kontrola, zda je ComboBox obsahuje n�jak� varianty
            If cbo.ListCount > 0 Then
            
                ' Kontrola, zda je vybr�na n�jak� varianta
                If cbo.ListIndex <> 0 Then
                
                    ' Z�sk�n� hodnoty vybran� v ComboBoxu a jej� p�evod na �et�zec
                    selectedCandidate = CStr(cbo.List(cbo.ListIndex))
                    
                    ' Zkontrolujeme, zda je vybran� varianta mezi dominovan�mi
                    For i = LBound(dominatedArray) To UBound(dominatedArray)
                        If Trim(dominatedArray(i)) = selectedCandidate Then
                            ' Pokud je varianta dominovan�, zobraz�me upozorn�n�
                            MsgBox "Vybran� varianta je dominovan� a ��dnou zm�nou vah nem��e b�t stanovena kompromisn� variantou.", vbOKOnly, "Upozorn�n�"
                        End If
                    Next i
                
                    With ws
                        .Unprotect "1234"
                        
                        ' Nastaven� popisku pro kl��ovou funkci
                        With .Range(.Cells(10 + numOfCriteria, 5 + numOfCandidates), .Cells(10 + numOfCriteria, 6 + numOfCandidates))
                            .Merge
                            .value = "U�itek vybran� varianty:"
                            .HorizontalAlignment = xlCenter
                            .WrapText = False
                        End With
                        
                        ' V�po�et kl��ov� funkce: pou�it� Match pro z�sk�n� hodnoty varianty
                        selectedColumn = Application.WorksheetFunction.Match(selectedCandidate, _
                            .Range(.Cells(12 + (2 * numOfCriteria), 5), .Cells(12 + (2 * numOfCriteria), 4 + numOfCandidates)), 0)
                        
                        selectedAddress = .Cells(13 + (2 * numOfCriteria), 4 + selectedColumn).Address
                        .Cells(10 + numOfCriteria, 7 + numOfCandidates).formula = "=" & selectedAddress
                    
                        ' Form�tov�n� na t�i desetinn� m�sta
                        .Cells(10 + numOfCriteria, 7 + numOfCandidates).NumberFormat = "0.000"
                        
                        ' Obnoven� p�vodn�ch vah krit�ri�
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
                MsgBox "Nen� k dispozici ��dn� varianta k v�b�ru.", vbExclamation
                Exit Sub
            End If
        End If
    Next cbo
End Sub


