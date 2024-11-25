Attribute VB_Name = "Module3"
Dim ws As Worksheet
Dim wsInput As Worksheet
Dim numOfCriteria As Integer
Dim numOfCandidates As Integer
Dim Shape As Shape

' Makro pro v�po�et metody WSA
Sub M3_metoda_WSA()
    
    ' Zobrazen� v�sledk� procedury a� po kompletn�m na�ten� procedury
    Application.ScreenUpdating = False
    
    ' Vol�n� kontroly vypln�n�ch hodnot
    Call CheckFilledData
    
    ' Kontrola unik�tn�ch hodnot v ��dc�ch
    If CheckUniqueRowValues() Then
        Exit Sub
    End If

' Ov��en� existence listu "Metoda WSA"
    Dim wsExists As Boolean
    wsExists = False
    
    For Each ws In ThisWorkbook.Sheets
        If ws.name = "Metoda WSA" Then
            wsExists = True
            ' P�esun na list a jeho vy�i�t�n�
            ws.Activate
            ws.Unprotect "1234"
            ws.Cells.Clear
            
            For Each Shape In ws.Shapes
                Shape.Delete
            Next Shape
            Exit For
        End If
    Next ws
    
    ' Vytvo�en� listu, pokud je�t� neexistuje
    If Not wsExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = "Metoda WSA"
        ' P�esun na nov� vytvo�en� list
        ws.Activate
        ws.Unprotect "1234"
    End If
    
' Pr�ce s "Metoda WSA"
    ' Nahr�n� vstupn�ch dat
    Dim lastRow As Long
    Dim i As Long
    Dim maxVal As Double
    
    ' Definice pracovn�ch list�
    Set wsInput = ThisWorkbook.Sheets("Vstupn� data")
    Set ws = ThisWorkbook.Sheets("Metoda WSA")
    
    numOfCriteria = wsInput.Range("C2").value
    numOfCandidates = wsInput.Range("F2").value
    
    With ws
        ' Kop�rov�n� z�hlav� a krit�ri� (p�id�n� apostrofu pro p�evod na text)
        For i = 3 To 4 + numOfCriteria
            .Cells(i, 2).value = "'" & wsInput.Cells(i, 2).value
            .Cells(i, 3).value = wsInput.Cells(i, 3).value
        Next i
        
        ' Kop�rov�n� z�hlav� a variant (p�id�n� apostrofu pro p�evod na text)
        For i = 4 To 4 + numOfCandidates
            .Cells(3, i).value = wsInput.Cells(3, i).value
            .Cells(4, i).value = "'" & wsInput.Cells(4, i).value
        Next i
    
        ' Kop�rov�n� ��seln�ch hodnot bez zm�ny form�tu
        .Range(.Cells(5, 4), .Cells(4 + numOfCriteria, 4 + numOfCandidates)).value = _
            wsInput.Range(wsInput.Cells(5, 4), wsInput.Cells(4 + numOfCriteria, 4 + numOfCandidates)).value
        
        ' Form�tov�n� z�hlav� B3 a� sloupec 4 + numOfCandidates na ��dku 3
        With .Range(.Cells(3, 2), .Cells(4, 4 + numOfCandidates))
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
    
        ' Zarovn�n� bun�k c�le na st�ed
        .Range(.Cells(4, 3), .Cells(4 + numOfCriteria, 3)).HorizontalAlignment = xlCenter
        
        ' Nastaven� stylu bun�k v�hy jako na form�t "0.0 %"
        .Range(.Cells(4, 4), .Cells(4 + numOfCriteria, 4)).NumberFormat = "0.0 %"
    End With

    
    ' Nastaven� ohrani�en� pro sloupce B a� D v ��dc�ch 4 a� posledn� krit�rium
    Dim column As Range
    Dim columnRange As Range
    
    ' Nastaven� rozsahu sloupc� B a� D
    Set columnRange = ws.Range(ws.Cells(4, 2), ws.Cells(4 + numOfCriteria, 4))
    
    ' Pro ka�d� sloupec v rozsahu
    For Each column In columnRange.Columns
        ' Nastaven� ohrani�en� pro prav� okraj
        With column.Columns(column.Columns.Count).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    Next column
    
    ' �prava minimaliza�n�ch krit�ri� a ��seln� form�t
    For i = 5 To (4 + numOfCriteria)
        ' Pokud je ve sloupci C hodnota "min"
        If ws.Cells(i, 3).value = "min" Then
            ' Najdi maxim�ln� hodnotu v cel�m ��dku
            maxVal = Application.WorksheetFunction.Max(wsInput.Range(wsInput.Cells(i, 4), wsInput.Cells(i, 4 + numOfCandidates)))
            ' Pro v�echny sloupce od E do posledn� varianty
            For j = 5 To (5 + numOfCandidates - 1)
                ' P�epo�et hodnoty na maxim�ln� hodnotu podle maxim�ln� hodnoty v ��dku
                ws.Cells(i, j).value = maxVal - wsInput.Cells(i, j).value
            Next j
        End If
    Next i
    
    ' �prava ���ky sloupc� (Autofit na minim�ln� 80px)
    AdjustColumnWidth ws, ws.Range(ws.Columns(2), ws.Columns(4 + numOfCandidates))
    
    With ws
        ' Popisek tabulky po maximalizaci minimaliza�n�ch krit�ri�
        .Range("A1:C2").Merge
        With .Range("A1")
            .value = "Maximalizovan� krit�ria"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Italic = True
        End With

        ' Nastaven� obsahu pro bu�ky c�le na "max"
        .Range(.Cells(5, 3), .Cells(5 + numOfCriteria - 1, 3)).FormulaR1C1 = "max"
        
        ' P�eform�tov�n� ��sla
        Dim cell As Range
        For Each cell In .Range(.Cells(5, 5), .Cells(5 + numOfCriteria - 1, 5 + numOfCandidates - 1)).Cells
            If cell.value = 0 Then
                cell.NumberFormat = "0"
            ElseIf Int(cell.value) = cell.value Then
                ' P�eform�tov�n� ��sla pomoc� odd�lova�e tis�c�
                cell.NumberFormat = "#,##0"
            Else
                ' P�eform�tov�n� ��sla na dv� desetinn� m�sta
                cell.NumberFormat = "0.0#"
            End If
        Next cell
    End With

' Normalizovan� matice
    With ws
        ' Kop�rov�n� obsahu prvn� tabulky
        .Range(.Cells(3, 2), .Cells(4 + numOfCriteria + 1, 4 + numOfCandidates)).Copy Destination:=.Cells(6 + numOfCriteria + 3, 2)
        
        ' Slou�en� bun�k pro popisek Normalizovan� matice
        .Range(.Cells(5 + numOfCriteria + 2, 1), .Cells(5 + numOfCriteria + 3, 3)).Merge
        
        With .Cells(5 + numOfCriteria + 2, 1)
            .value = "Normalizovan� matice"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Italic = True
            .EntireColumn.AutoFit
        End With
        
        ' Definice po��te�n� a koncov� bu�ky pro vzorec
        Set startCell = .Cells(4, 5)
        Set endCell = .Cells(4, 5 + numOfCandidates - 1)
        
        ' Definice c�lov� bu�ky pro v�sledek
        Dim targetCell As Range
        Set targetCell = .Cells(5 + numOfCriteria + 5, 5)
        
        ' Cyklus pro �pravu ka�d� c�lov� bu�ky s odpov�daj�c�m indexem ��dku
        ' Aplikace vzorce pro celou matici normalizace
        For i = 1 To numOfCriteria
            For j = 1 To numOfCandidates
                ' P�epo��t�n� hodnoty c�lov� bu�ky pro v�po�et normalizovan� matice
                targetCell.Offset(i, j - 1).formula = "=(" & startCell.Offset(i, j - 1).Address(False, False) & _
                    "-MIN(" & startCell.Offset(i, 0).Address(False, True) & ":" & endCell.Offset(i, 0).Address(False, True) & "))/(" & _
                    "MAX(" & startCell.Offset(i, 0).Address(False, True) & ":" & endCell.Offset(i, 0).Address(False, True) & _
                    ")-MIN(" & startCell.Offset(i, 0).Address(False, True) & ":" & endCell.Offset(i, 0).Address(False, True) & "))"
            Next j
        Next i
                
        ' Nastaven� form�tu ��sla
        For Each cell In .Range(.Cells(11 + numOfCriteria, 5), .Cells(10 + (2 * numOfCriteria), 5 + numOfCandidates - 1)).Cells
            If Int(cell.value) = cell.value Then
                cell.NumberFormat = "0"
            Else
                ' Pokud hodnota nen� cel� ��slo pak na dv� desetinn� m�sta
                cell.NumberFormat = "0.00"
            End If
        Next cell
        
        ' Suma vah pro kontrolu
        .Cells(11 + (2 * numOfCriteria), 4).formula = "=SUM(" & .Range(.Cells(11 + numOfCriteria, 4), .Cells(10 + (2 * numOfCriteria), 4)).Address & ")"
        .Cells(11 + (2 * numOfCriteria), 4).Font.Bold = True
        
        
    ' U�itek jednotliv�ch variant
        ' V�pis variant pro p�ehlednost u�itku
        .Range(.Cells(4, 5), .Cells(4, 5 + numOfCandidates - 1)).Copy Destination:=.Cells(12 + (2 * numOfCriteria), 5)
        
        ' Nastaven� popisku pro u�itky
        With .Cells(13 + (2 * numOfCriteria), 4)
            .value = "U�itek"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Borders(xlEdgeRight).LineStyle = xlContinuous
        End With
        
        For j = 1 To numOfCandidates
            ' V�po�et u�itk� pro varianty
            .Cells(13 + (2 * numOfCriteria), 4 + j).formula = _
                "=SUMPRODUCT(" & .Range(.Cells(11 + numOfCriteria, 4), .Cells(10 + (2 * numOfCriteria), 4)).Address & "," & _
                        .Range(.Cells(11 + numOfCriteria, 4 + j), .Cells(10 + (2 * numOfCriteria), 4 + j)).Address & ")"
            
            ' Form�tov�n� na t�i desetinn� m�sta
            .Cells(13 + (2 * numOfCriteria), 4 + j).NumberFormat = "0.000"
        Next j
        
        ' P�id�n� podm�n�n�ho form�tov�n� barvou pro u�itky (Zelen� nejlep��, �erven� nejhor��)
        .Range(.Cells(13 + (2 * numOfCriteria), 5), .Cells(13 + (2 * numOfCriteria), 4 + numOfCandidates)).FormatConditions.AddColorScale ColorScaleType:=3
        
        ' Popisek nejvy���ho u�itku
        With .Cells(17 + numOfCriteria, 6 + numOfCandidates)
            .formula = "Nejvy��� u�itek:"
            .Font.Bold = True
            .EntireColumn.AutoFit
        End With
        
        ' Cyklus pro vyps�n� v�sledk� od nejlep��ho po nejhor��mu
        For i = 1 To numOfCandidates
            ' Vyhled�n� a zobrazen� i-t�ho nejvy���ho u�itku
            .Cells(16 + numOfCriteria + i, 7 + numOfCandidates).Formula2 = "=XLOOKUP(LARGE(" & .Range(.Cells(13 + (2 * numOfCriteria), 5), _
                    .Cells(13 + (2 * numOfCriteria), 4 + numOfCandidates)).Address & "," & i & ")" & "," & _
                    .Range(.Cells(13 + (2 * numOfCriteria), 5), .Cells(13 + (2 * numOfCriteria), 4 + numOfCandidates)).Address & "," & _
                    .Range(.Cells(12 + (2 * numOfCriteria), 5), .Cells(12 + (2 * numOfCriteria), 4 + numOfCandidates)).Address & ",,0,1)"
            
            .Cells(16 + numOfCriteria + i, 7 + numOfCandidates).HorizontalAlignment = xlCenter
        Next i

        ' �prava form�tov�n� nejlep�� varianty
        With .Cells(17 + numOfCriteria, 7 + numOfCandidates)
            .Font.Bold = False
            .Font.Italic = True
        End With

        ' Nastaven� popisku pro vybr�n� testovan� varianty
        With .Range(.Cells(9 + numOfCriteria, 5 + numOfCandidates), .Cells(9 + numOfCriteria, 6 + numOfCandidates))
            .Merge
            .value = "Jak� varianta m� b�t testov�na:"
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .WrapText = False
            .Select
        End With
        
        If .Columns(5 + numOfCandidates).ColumnWidth < 12 Then
            .Columns(5 + numOfCandidates).ColumnWidth = 12
        End If
        
        ' Zavol�n� procedury pro kontrolu, zda je n�kter� z variant dominov�na jinou
        Call FindDominatedCandidates(ws)
        
        ' �prava ���ky sloupc� (Autofit na minim�ln� 80px)
        AdjustColumnWidth ws, .Range(.Columns(5 + numOfCandidates), .Columns(6 + numOfCandidates))
        
        ' Vol�n� funkce, kter� vykresl� rozbalovac� seznam
        ' Parametry jsou WorkSheet (ws), jm�no (name), v�stup (targetCell), mo�nosti (optionsRange) a Makro (macroName)
        AddComboBox ws, "newBestCandidateWSA", ws.Cells(9 + numOfCriteria, 7 + numOfCandidates), ws.Range(.Cells(12 + (2 * numOfCriteria), 5), .Cells(12 + (2 * numOfCriteria), 4 + numOfCandidates)), "newBestCandidateWSA_Change"
        
        ' Vol�n� funkce pro p�id�n� tla��tka na spu�t�n� Solveru
        AddButtonTo ws, ws.Cells(14 + numOfCriteria, 7 + numOfCandidates), "Vy�e�it", "CallSolverWSA"
        
        ' P�id�n� tla��tka pro op�tovn� nahr�n� vstupn�ch dat
        AddButtonTo ws, ws.Cells(4, 7 + numOfCandidates), "Aktualizovat", "M3_metoda_WSA"
    End With
    
    ws.Protect "1234"

End Sub

' Procedura metody WSA volaj�c� proceduru obsluhuj�c� zm�nu hodnoty ComboBoxu
Private Sub newBestCandidateWSA_Change()
    Set ws = ThisWorkbook.Sheets("Metoda WSA")
    
    'Zavol�n� metody a p�ed�n� parametr� worksheet a n�zev ComboBoxu
    Call newBestCandidate_Change(ws, "newBestCandidateWSA")
End Sub

' Vol�n� Solveru pro metodu WSA
Private Sub CallSolverWSA()
    Set ws = ThisWorkbook.Sheets("Metoda WSA")
    
    ' P�ed�n� procedu�e obsluhuj�c� Solver po�adovan� parametry
    Call M5_Solver(ws, "newBestCandidateWSA")

End Sub
