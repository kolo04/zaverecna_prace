Attribute VB_Name = "Module1"
Dim ws As Worksheet
Dim wsInput As Worksheet
Dim wsExists As Boolean
Dim numOfCriteria As Integer
Dim numOfCandidates As Integer
Dim criteriaDone As Boolean

' P�i otev�en� souboru je automaticky spu�t�na tato procedura
Sub auto_open()
    
    ' Zavol�n� procedury Vstupni_data
    Call InputData
End Sub

' �vodn� procedura, kter� je automaticky spu�t�na po otev�en�
Sub InputData()
    
    ' Zobrazen� v�sledk� procedury a� po kompletn�m na�ten�
    ' Zrychluje proces a zabra�uje nep�ijemn� "blik�n�" p�ed o�ima u�ivatele
    Application.ScreenUpdating = False
    
    ' Ov��en� existence listu "Vstupn� data"
    wsExists = False
    
    For Each ws In ThisWorkbook.Sheets
        If ws.name = "Vstupn� data" Then
            wsExists = True
            ' P�esun na list a jeho vy�i�t�n�
            ws.Activate
            ws.Unprotect "1234"
            ws.Cells.Clear
            
            ' Deklarace prom�nn�, kter� je typu Shape
            ' Jak�koliv objekt, kter� m� tvar = tla��tko, TextBox, ComboBox, ..
            Dim Shape As Shape
            'Cyklus, kter� projde v�echny objekty typu Shape na listu a odstran� je
            For Each Shape In ws.Shapes
                Shape.Delete
            Next Shape
            
            Exit For
        End If
    Next ws
    
    ' Vytvo�en� listu, pokud je�t� neexistuje
    If Not wsExists Then
        ' P�id�n� listu za posledn� ji� existuj�c� list
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = "Vstupn� data"
        ' P�esun na nov� vytvo�en� list
        ws.Activate
        ws.Unprotect "1234"
    End If
        
    ' Nahr�n� vstupn�ch dat
    With ws
    
        ' Vytvo�en� z�hlav� tabulky
        .Range("B2").value = "Po�et krit�ri�"
        .Range("C2").value = 0 ' Po�et krit�ri� na za��tku bude nula
        .Range("B4").value = "Krit�rium"
                
        ' Tu�n� p�smo pro po�et krit�ri�
        .Range("B2").Font.Bold = True
        .Range("B4").Font.Bold = True
        
        ' �prava ���ky sloupc� (Autofit na minim�ln� 80px)
        AdjustColumnWidth ws, 2
        
        '.Columns("B").EntireColumn.AutoFit
        .Cells(4, 2).Select
        
        Application.ScreenUpdating = True
        
        ' Pokud je criteriaDone Nepravda, pak
        If criteriaDone = False Then
            ' Zavol�n�/Vytvo�en� UserFormu pro zad�v�n� krit�ri�
            If Not AddCriteriaForm Is Nothing Then
                Unload AddCriteriaForm
                .Unprotect "1234"
                Set AddCriteriaForm = New AddCriteriaForm
                AddCriteriaForm.Show
            End If
            
            .Unprotect "1234"
            
            ' Z�sk�n� po�tu krit�ri�
            numOfCriteria = .Range("C2").value
            
            ' Kontrola spln�n� podm�nky pro minim�ln� 2 krit�ria
            If numOfCriteria < 2 Then
                MsgBox "P�i rozhodov�n� bychom m�li zohled�ovat minim�ln� 2 krit�ria.", vbExclamation
                .Protect "1234"
                Exit Sub
            End If
        End If
    End With
    
End Sub

' Procedura obsluhuj�c� stanoven� c�l� ��elov�ch funkc� pro jednotliv� krit�ria
Sub WeightedInputData()

    Application.ScreenUpdating = False
    
    ' Definice pracovn�ho listu pro vstupn� data
    Set wsInput = ThisWorkbook.Sheets("Po�ad� krit�ri�")
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
        
    With ws
        .Activate
        .Unprotect "1234"
        
        ' Z�sk�n� po�tu krit�ri�
        numOfCriteria = .Range("C2").value
        
        ' Nastaven� textu "C�l" do bu�ky C4
        .Range("C4").value = "C�l"
        
        ' Vytvo�en� rozev�rac�ho seznamu s mo�nostmi "min" a "max" pro ka�dou bu�ku v rozsahu C4 a� C(4 + numOfCriteria)
        Dim criteriaRange As Range
        Dim options As Variant
        options = Array("min", "max")
        
        Set criteriaRange = .Range(.Cells(5, 3), .Cells(5 + numOfCriteria - 1, 3))
        For Each cell In criteriaRange
            With cell.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(options, ",")
            End With
            ' Nastaven� popisku "Vyberte" pro ka�dou bu�ku
            cell.value = "Vyberte"
            cell.Locked = False
        Next cell
        
        ' Form�tov�n� z�hlav� B4:D4
        With .Range("B4:D4")
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
    
        ' Zarovn�n� bun�k C4:C (4 + numOfCriteria) na st�ed
        .Range("C4:C" & 4 + numOfCriteria).HorizontalAlignment = xlCenter
        
        ' Nastaven� stylu bun�k D4:D (4 + numOfCriteria) jako "Percent" s form�tem "0.0 %"
        .Range("D4:D" & 4 + numOfCriteria).NumberFormat = "0.0 %"
        
        ' �prava ���ky sloupc�
        AdjustColumnWidth ws, .Range(.Columns(2), .Columns(3))
        
        .Cells(5, 3).Select
    End With
    
    HideButton ws, "Stanovit v�hy"
    
    ' P�id�n� tla��tka pro n�vrat na vstupn� data
    AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Pokra�ovat", "Candidates"
    
    ws.Protect "1234"
    
End Sub

' Procedura obsluhuj�c� zavol�n� p�id�v�n� variant a p�id�n� tla��tka "Pokra�ovat"
' pro p�echod na vypln�n� hodnot tabulky
Sub Candidates()
    Set ws = ThisWorkbook.Sheets("Vstupn� data")

    ' Ov��en�, zda jsou v�echny c�le vypln�ny
    Dim i As Integer
    For i = 1 To numOfCriteria
        If ws.Cells(4 + i, 3).value = "Vyberte" Then
            ws.Cells(4 + i, 3).Select
            MsgBox "Vypl�te pros�m v�echny c�le.", vbExclamation
            Exit Sub
        End If
    Next i
    
    With ws
        .Unprotect "1234"
        If IsEmpty(Range("E2")) Then
            ' P�id�v�n� variant
            .Range("E2").value = "Po�et variant"
            .Range("F2").value = 0 ' Po�et variant na za��tku bude roven nule
            .Range("E3").value = "Varianta"
            
            ' Tu�n� p�smo pro po�et variant
            .Range("E2:E3").Font.Bold = True
            
            .Columns("E").EntireColumn.AutoFit
            .Cells(3, 5).Select
        End If
        
        If candidatesDone = False Then
            ' Otev�en� UserFormu pro zad�v�n� variant
            If Not AddCandidateForm Is Nothing Then
                Unload AddCandidateForm
                .Unprotect "1234"
                Set AddCandidateForm = New AddCandidateForm
                AddCandidateForm.Show
            End If

            .Unprotect "1234"

        End If
        
        ' Z�sk�n� po�tu variant
        numOfCandidates = .Range("F2").value
        
        ' Kontrola spln�n� podm�nky pro minim�ln� 2 varianty
        If numOfCandidates < 2 Then
            MsgBox "P�i rozhodov�n� bychom m�li zohled�ovat minim�ln� 2 varianty.", vbExclamation
            .Protect "1234"
            Exit Sub
            End
        Else
            ' P�id�n� tla��tka pro vypln�n� dat
            ws.Protect "1234", UserInterfaceOnly
        End If
    End With
End Sub

' Procedura pro vypln�n� hodnot tabulky
Sub FillData()
    Dim cellRange As Range

    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    
    With ws
        numOfCriteria = .Range("C2").value
        numOfCandidates = .Range("F2").value
        
        ' Kontrola po�tu krit�ri�
        If numOfCriteria < 2 Then
            MsgBox "P�i rozhodov�n� bychom m�li zohled�ovat minim�ln� 2 krit�ria.", vbExclamation
            Exit Sub
        End If
        
        ' Kontrola po�tu variant
        If numOfCandidates < 2 Then
            MsgBox "P�i rozhodov�n� bychom m�li zohled�ovat minim�ln� 2 varianty.", vbExclamation
            Exit Sub
        End If
        
        ' Pro ka�dou zm�n�nou bu�ku krit�ria
        For Each cell In Range("B5:B" & 5 + numOfCriteria - 1)
            ' Kontrola, zda jsou pole v D pr�zdn� ve stejn�m ��dku jako pole v sloupci B
            If IsEmpty(cell.Offset(0, 2).value) Then
                MsgBox "Vypl�te, pros�m, v�hu krit�ria.", vbExclamation
                
                ' Ozna�en� pr�zdn� bu�ky
                cell.Offset(0, 2).Select
                
                Exit Sub
                
            ' Kontrola, zda jsou pole v C pr�zdn� ve stejn�m ��dku jako pole v sloupci B
            ElseIf IsEmpty(cell.Offset(0, 1).value) Then
                MsgBox "Vypl�te, pros�m, c�l krit�ria", vbExclamation
                ' Ozna�en� pr�zdn� bu�ky
                cell.Offset(0, 1).Select
                Exit Sub
            End If
        Next cell
        
        ' Nastaven� rozsahu bun�k pro zad�n� hodnot krit�ri� a variant
        Set cellRange = ws.Range(ws.Cells(5, 5), ws.Cells(4 + numOfCriteria, 4 + numOfCandidates))
        
        ' Cyklus pro zad�n� hodnot krit�ri� a variant
        For Each cell In cellRange
            ' Kontrola, zda je bu�ka pr�zdn�
            If IsEmpty(cell) Then
                ' Zavol�n� procedury FillDataForm pouze pro pr�zdn� bu�ky
                FillDataForm cell
            End If
        Next cell

        ' Kontrola, zda jsou bu�ky pr�zdn�
        ws.Unprotect "1234"
        
        HideButton ws, "Pokra�ovat"
        
        If Not IsEmpty(ws.Range(ws.Cells(5, 5), ws.Cells(4 + numOfCriteria, 4 + numOfCandidates))) Then
            ' P�id�n� tla��tka pro �pravu vypln�n�ch hodnot
            AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Upravit hodnoty", "EditCellValue"
        Else
            ' P�id�n� tla��tka pro vypln�n� dat
            AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Pokra�ovat", "FillData"
        End If
        
        ' P�id�n� tla��tka pro spu�t�n� metody WSA
        AddButtonTo ws, ws.Range("B" & 9 + numOfCriteria), "Metoda WSA", "M3_metoda_WSA"
        
        ' P�id�n� tla��tka pro spu�t�n� metody bazick� varianty s v�t�� ���kou
        AddButtonTo ws, ws.Range("D" & 9 + numOfCriteria, "E" & 9 + numOfCriteria), "Metoda bazick� varianty", "M4_metoda_Bazicke_varianty", 4.5, 1
        
        ws.Protect "1234"
    End With
End Sub

' Procedura pro napln�n� bu�ky, kterou procedura dostane formou parametru
Sub FillDataForm(cellRef As Variant)
    Dim cell As Range
    Dim criteriaNames() As Variant
    Dim variantNames() As Variant
    Dim inputVal As Variant
    Dim validInput As Boolean
    Dim convertedVal As Double

    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    
    ' P�etypujeme referenci na bu�ku na objekt typu Range
    Set cell = cellRef

    ' Z�sk�n� n�zvu krit�ria
    criteriaName = ws.Cells(cell.Row, 2).value
    
    ' Z�sk�n� n�zvu varianty
    variantName = ws.Cells(4, cell.column).value
    
    ' Ozna�en� bu�ky pro zad�n� hodnoty
    cell.Select
    
    Do
        ' Kontrola, zda bu�ka obsahuje ji� hodnotu
        If Not IsEmpty(cell.value) Then
            ' Pokud bu�ka ji� obsahuje hodnotu, nab�dne se mo�nost jej� �pravy
            inputVal = InputBox("Aktu�ln� hodnota pro krit�rium '" & criteriaName & "' a variantu '" & variantName & "' je: " & _
                        cell.value & vbCrLf & "Zadejte novou hodnotu nebo klikn�te na OK pro ponech�n� st�vaj�c� hodnoty:", _
                        "Hodnota pro krit�rium a variantu", cell.value)
    
        Else
            ' Pokud bu�ka neobsahuje hodnotu, standardn� postup zad�v�n� nov� hodnoty
            inputVal = InputBox("Zadejte hodnotu pro krit�rium '" & criteriaName & "' a variantu '" & variantName & "':", _
            "Hodnota pro krit�rium a variantu")
        End If
        
        ' Kontrola, zda u�ivatel klikl na Cancel
        If inputVal = "" Then
            MsgBox "Zad�n� bylo zru�eno.", vbInformation
            ws.Protect "1234"
            End
        End If
        
        ' Pokus o p�evod zadan� hodnoty na ��slo
        If IsNumeric(inputVal) Then
            ' P�evod na ��slo (Double)
            convertedVal = CDbl(inputVal)
            
            ' Ulo�en� hodnoty do bu�ky
            ws.Unprotect "1234"
            cell.value = convertedVal
            'Nastaven� ��seln�ho form�tu bu�ky
            If cell.value = 0 Then
                cell.NumberFormat = "0"
            ElseIf Int(cell.value) = cell.value Then
                ' P�eform�tov�n� ��sla pomoc� odd�lova�e tis�c�
                cell.NumberFormat = "#,##0"
            Else
                ' P�eform�tov�n� ��sla na dv� desetinn� m�sta
                cell.NumberFormat = "0.0#"
            End If
            
            ws.Protect "1234"
            validInput = True ' Platn� vstup
        Else
            ' Hodnota nen� ��slo, zobrazen� chybov� zpr�vy a cyklus pokra�uje
            MsgBox "Zad�vejte, pros�m, pouze ��seln� hodnoty." & vbCrLf & _
            "V p��pad� krit�ria 'ano/ne' vkl�dejte hodnoty 1 pro 'ano' a 0 pro 'ne'.", vbExclamation
            validInput = False
        End If
    Loop Until validInput
    
End Sub

' Procedura kontroluj�c�, zda jsou hodnoty tabulky vypln�ny
Sub CheckFilledCells()
    ' Nastaven� pracovn�ho listu
    Set ws = ThisWorkbook.Sheets("Vstupn� data")

    ' Z�sk�me po�et krit�ri� a po�et variant
    numOfCriteria = ws.Range("C2").value
    numOfCandidates = ws.Range("F2").value
    
   ' Proch�zen� v�ech bun�k v dan�m rozsahu
    For j = 1 To numOfCandidates
        For i = 1 To numOfCriteria
            ' Nastaven� bu�ky
            Set cell = ws.Cells(4 + i, 4 + j)
            
            ' Kontrola, zda je bu�ka pr�zdn�
            If IsEmpty(cell) Then
                ' Zavol�n� procedury FillDataForm pro pr�zdn� bu�ky
                FillDataForm cell
                
                ' Pokud najde pr�zdnou bu�ku, ukon�� kontrolu
                Exit Sub
            End If
        Next i
    Next j
End Sub

' K�d pro vytvo�en� formul��e, kter� umo�n� u�ivateli vybrat bu�ku
Sub EditCellValue()
    Dim selectedRange As Range
    Dim cell As Range
    
    ' Umo�n� u�ivateli vybrat bu�ku/bu�ky
    On Error Resume Next
    Set selectedRange = Application.InputBox("Vyberte bu�ku (bu�ky), kterou (kter�) chcete upravit:", Type:=8)
    On Error GoTo 0
    
    ' Pokud u�ivatel klikl na Cancel, ukon��me proceduru
    If selectedRange Is Nothing Then
        Exit Sub
    End If
    
    ' Projdeme ka�dou vybranou bu�ku z rozsahu
    For Each cell In selectedRange
        ' Zavol�me proceduru FillDataForm pro ka�dou bu�ku zvl᚝
        FillDataForm cell
    Next cell
    
End Sub

' Procedura volaj�c� formul�� pro p�id�n� dal��ch krit�ri�
Sub AddMoreCriteria()

' Nastaven� hodnoty criteriaDone na False pro p�id�n� dal��ch krit�ri�
    criteriaDone = False
    
    ThisWorkbook.ActiveSheet.Unprotect "1234"
    
    ' Zavol�n� formul��e
    AddCriteriaForm.Show
End Sub

' Procedura volaj�c� formul�� pro p�id�n� dal��ch variant
Sub AddMoreCandidates()

' Tla��tko pro p�id�n� dal��ch variant
    candidatesDone = False
    
    ThisWorkbook.ActiveSheet.Unprotect "1234"
    
    AddCandidateForm.Show
End Sub

' Procedura vol� a napl�uje formul�� pro odebr�n� krit�ria
Sub RemoveCriteria()
    Dim criteriaList As Range
    Dim criteriaCell As Range
    
    ' Nastaven� pracovn�ho listu, kde jsou krit�ria ulo�ena
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    
    ' Z�skat po�et krit�ri� z listu
    numOfCriteria = ws.Range("C2").value
    
    ' Definuj rozsah obsahuj�c� krit�ria
    Set criteriaList = ws.Range("B5:B" & 5 + numOfCriteria - 1)
    
    ' Vynuluj ListBox
    RemoveCriteriaForm.CriteriaListBox.Clear

    ' Napl� ListBox seznamem existuj�c�ch krit�ri�
    For Each criteriaCell In criteriaList
        RemoveCriteriaForm.CriteriaListBox.AddItem criteriaCell.value
    Next criteriaCell

    ' Zavol�n� formul��e pro odebr�n� krit�ri�
    RemoveCriteriaForm.Show

End Sub

' Procedura pro odebr�n� varianty
Sub RemoveCandidate()
    Dim candidateList As Range
    Dim candidateCell As Range
    
    ' Nastaven� pracovn�ho listu, kde jsou krit�ria ulo�ena
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    
    ' Z�skat po�et variant z listu
    numOfCandidates = ws.Range("F2").value
    
    ' Definov�n� rozsahu obsahuj�c� varianty
    Set candidateList = ws.Range(ws.Cells(4, 5), ws.Cells(4, 5 + numOfCandidates - 1))

    ' Vypr�zd�n� ListBoxu
    RemoveCandidateForm.CandidateListBox.Clear
    
    ' Napln�n� ListBox seznamem existuj�c�ch variant
    For Each candidateCell In candidateList
        RemoveCandidateForm.CandidateListBox.AddItem candidateCell.value
    Next candidateCell
    
    ' Zavol�n� formul��e pro odebr�n� varianty
    RemoveCandidateForm.Show
End Sub

