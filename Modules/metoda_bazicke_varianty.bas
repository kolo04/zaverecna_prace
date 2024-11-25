Attribute VB_Name = "Module4"
Dim ws As Worksheet
Dim wsInput As Worksheet
Dim numOfCriteria As Integer
Dim numOfCandidates As Integer
Dim Shape As Shape

' Makro pro v�po�et metody Bazick� varianty
Sub M4_metoda_Bazicke_varianty()

    ' Zobrazen� v�sledk� procedury a� po kompletn�m na�ten� procedury
    Application.ScreenUpdating = False
    
    ' Vol�n� kontroly vypln�n�ch hodnot
    Call CheckFilledData
    
    ' Kontrola unik�tn�ch hodnot v ��dc�ch
    If CheckUniqueRowValues() Then
        Exit Sub
    End If
    
' Ov��en� existence listu "Metoda bazick� varianty"
    Dim wsExists As Boolean
    wsExists = False
    
    For Each ws In ThisWorkbook.Sheets
        If ws.name = "Metoda bazick� varianty" Then
            wsExists = True
            ' P�esun na list a jeho vy�i�t�n�
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
    
    ' Vytvo�en� listu, pokud je�t� neexistuje
    If Not wsExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = "Metoda bazick� varianty"
        
        ' P�esun na nov� vytvo�en� list
        ws.Activate
        ws.Unprotect "1234"
    End If
    
' Pr�ce s "Metoda bazick� varianty"
    ' Nahr�n� vstupn�ch dat
    Dim lastRow As Long
    Dim i As Long
    Dim maxVal As Double
    
    ' Definice pracovn�ch list�
    Set wsInput = ThisWorkbook.Sheets("Vstupn� data")
    Set ws = ThisWorkbook.Sheets("Metoda bazick� varianty")
    
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
        
        ' Form�tov�n� z�hlav� B3 a� sloupec 5 + numOfCandidates na ��dku 4
        With .Range(.Cells(3, 2), .Cells(4, 5 + numOfCandidates))
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
    
    With ws
    ' Vytvo�en� sloupce "B�ze" a vlo�en� hodnot
        .Cells(4, 5 + numOfCandidates).value = "B�ze"
        
        ' Napln�n� sloupce "B�ze" podle extr�m� ��dk�
        For i = 5 To 5 + numOfCriteria
            If .Cells(i, 3).value = "min" Then
                .Cells(i, 5 + numOfCandidates).formula = "=MIN(" & .Cells(i, 5).Address & ":" & .Cells(i, 4 + numOfCandidates).Address & ")"
            ElseIf .Cells(i, 3).value = "max" Then
                .Cells(i, 5 + numOfCandidates).formula = "=MAX(" & .Cells(i, 5).Address & ":" & .Cells(i, 4 + numOfCandidates).Address & ")"
            End If
        Next i

        ' Form�tov�n� sloupce "B�ze"
        With Range(.Cells(4, 5 + numOfCandidates), .Cells(4 + numOfCriteria, 5 + numOfCandidates))
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlInsideVertical).LineStyle = xlContinuous
        End With
        
        ' Popisek tabulky po p�id�n� b�ze
        .Range("A1:C2").Merge
        With .Range("A1")
            .value = "Krit�ria a jejich b�ze"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Italic = True
        End With
        
        ' P�eform�tov�n� ��sla
        Dim cell As Range
        For Each cell In .Range(.Cells(5, 5), .Cells(5 + numOfCriteria - 1, 5 + numOfCandidates)).Cells
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
        
        ' �prava ���ky sloupc� (Autofit na minim�ln� 80px)
        AdjustColumnWidth ws, .Range(.Columns(2), .Columns(4 + numOfCandidates))
        ws.Columns(5 + numOfCandidates).ColumnWidth = 10.33
        
    End With
    
' Normalizovan� matice
    With ws
        ' Kop�rov�n� obsahu prvn� tabulky
        .Range(.Cells(3, 2), .Cells(5 + numOfCriteria + 1, 4 + numOfCandidates)).Copy Destination:=.Cells(6 + numOfCriteria + 3, 2)
        
        ' Slou�en� bun�k pro popisek Normalizovan� matice
        .Range(.Cells(5 + numOfCriteria + 2, 1), .Cells(5 + numOfCriteria + 3, 3)).Merge
        
        With .Cells(5 + numOfCriteria + 2, 1)
            .value = "Normalizovan� matice"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Italic = True
            .EntireColumn.AutoFit
        End With
        
        Dim epsilon As Double
        epsilon = 0.0000000001 ' Mal� ��slo pro o�et�en� d�len� 0

        ' Cyklus pro vypo��t�n� normalizovan� hodnoty bu�ky
        For i = 1 To numOfCriteria
            For j = 1 To numOfCandidates
            
                ' Vypo��t�n� normalizovan� hodnoty
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
        
        ' Nastaven� form�tu ��sla
        For Each cell In .Range(.Cells(11 + numOfCriteria, 5), .Cells(11 + (2 * numOfCriteria), 5 + numOfCandidates - 1)).Cells
            If Int(cell.value) = cell.value Then
                cell.NumberFormat = "0"
            Else
                ' Pokud hodnota nen� cel� ��slo pak na dv� desetinn� m�sta
                cell.NumberFormat = "0.00"
            End If
        Next cell

        ' Suma vah pro kontrolu
        .Cells(11 + (2 * numOfCriteria), 4).formula = "=SUM(" & .Range(.Cells(11 + numOfCriteria, 4), .Cells(11 + (2 * numOfCriteria) - 1, 4)).Address & ")"
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
            '.EntireColumn.AutoFit
            .Select
        End With
        
        ' Zavol�n� procedury pro kontrolu, zda je n�kter� z variant dominov�na jinou
        Call FindDominatedCandidates(ws)
        
        ' �prava ���ky sloupc� (Autofit na minim�ln� 80px)
        AdjustColumnWidth ws, .Range(.Columns(5 + numOfCandidates), .Columns(7 + numOfCandidates))
        
        If .Columns(5 + numOfCandidates).ColumnWidth < 12 Then
            .Columns(5 + numOfCandidates).ColumnWidth = 12
        End If
        
        ' Vol�n� funkce, kter� vykresl� rozbalovac� seznam
        ' Parametry jsou WorkSheet (ws), v�stup (targetCell) a mo�nosti (optionsRange)
        AddComboBox ws, "newBestCandidateBV", ws.Cells(9 + numOfCriteria, 7 + numOfCandidates), ws.Range(.Cells(12 + (2 * numOfCriteria), 5), .Cells(12 + (2 * numOfCriteria), 4 + numOfCandidates)), "newBestCandidateBV_Change"
        
        ' Vol�n� funkce pro p�id�n� tla��tka na spu�t�n� Solveru
        AddButtonTo ws, ws.Cells(14 + numOfCriteria, 7 + numOfCandidates), "Vy�e�it", "CallSolverBV"
        
        ' P�id�n� tla��tka pro op�tovn� nahr�n� vstupn�ch dat
        AddButtonTo ws, ws.Cells(4, 7 + numOfCandidates), "Aktualizovat", "M4_metoda_Bazicke_varianty"
        
    End With
    
    ws.Protect "1234"
    
End Sub

' Procedura metody bazick� varianty volaj�c� proceduru obsluhuj�c� zm�nu hodnoty ComboBoxu
Private Sub newBestCandidateBV_Change()
    Set ws = ThisWorkbook.Sheets("Metoda bazick� varianty")
    
    'Zavol�n� metody a p�ed�n� parametr� worksheet a n�zev ComboBoxu
    Call newBestCandidate_Change(ws, "newBestCandidateBV")
End Sub

' Vol�n� Solveru pro metodu bazick� varianty
Private Sub CallSolverBV()

    Set ws = ThisWorkbook.Sheets("Metoda bazick� varianty")
    
    ' P�ed�n� procedu�e obsluhuj�c� Solver po�adovan� parametry
    Call M5_Solver(ws, "newBestCandidateBV")

End Sub
