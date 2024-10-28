Attribute VB_Name = "Module7"
Dim ws As Worksheet
Dim numOfCriteria As Integer
Dim numOfCandidates As Integer

Sub SetObjectives()

    ' Otev�e formul�� s mo�nost� v�b�ru metody pro zad�n� c�l�
    ObjectivesForm.Show

End Sub

Sub UploadObjectives()
    Dim subject As String
    Dim Objectives As Range
        
    ' Odkaz na list "Vstupn� data" a po�et krit�ri�
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    numOfCriteria = ws.Range("C2").value

    ' Nastaven� c�lov� bu�ky (za��tek v C5 + po�et krit�ri�)
    Set Objectives = ws.Range(ws.Cells(5, 3), ws.Cells(4 + numOfCriteria, 3))

    ' P�edm�t pro zobrazen� v InputBoxu
    subject = "c�le"
    
    ' Vol�n� samostatn� procedury pro nahr�v�n� dat
    Call UploadData(Objectives, subject)
    
    Call CheckObjectives(Objectives, ws)
    
End Sub

' Procedura pro nahr�n� bloku dat (krit�ria x varianty) z extern�ho souboru do tabulky
Sub UploadDataBlock()
    Dim srcRange As Range
    Dim targetRange As Range
    Dim validSelection As Boolean

    ' Nastaven� listu a zji�t�n� po�tu krit�ri� a kandid�t�
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    numOfCriteria = ws.Range("C2").value
    numOfCandidates = ws.Range("F2").value

    ' Ov��en�, �e po�et krit�ri� a kandid�t� je dostate�n�
    If numOfCriteria < 2 Then
        MsgBox "P�i rozhodov�n� bychom m�li zohled�ovat minim�ln� 2 krit�ria.", vbExclamation
        Exit Sub
    End If
    If numOfCandidates < 2 Then
        MsgBox "P�i rozhodov�n� bychom m�li zohled�ovat minim�ln� 2 varianty.", vbExclamation
        Exit Sub
    End If

    ' V�b�r oblasti dat k nahr�n�
RestartSelection:
    validSelection = False
    Set srcRange = Nothing

    ' U�ivatel zvol� rozsah dat pomoc� InputBoxu
    On Error Resume Next
    Set srcRange = Application.InputBox("Vyberte oblast dat o velikosti " & numOfCriteria & " ��dk� a " & numOfCandidates & " sloupc�:", _
                                        "Nahr�t data", Type:=8)
    On Error GoTo 0

    ' Kontrola, zda u�ivatel n�co vybral
    If srcRange Is Nothing Then
        MsgBox "Nebyla vybr�na ��dn� oblast.", vbExclamation
        Exit Sub
    ElseIf srcRange.Rows.Count <> numOfCriteria Or srcRange.Columns.Count <> numOfCandidates Then
        MsgBox "Vybran� rozsah mus� m�t p�esn� " & numOfCriteria & " ��dk� (krit�ri�) a " & numOfCandidates & " sloupc� (variant).", vbExclamation
        GoTo RestartSelection
    End If

    ' Nastaven� c�lov�ho rozsahu pro vlo�en� dat v listu
    Set targetRange = ws.Range(ws.Cells(5, 5), ws.Cells(4 + numOfCriteria, 4 + numOfCandidates))

    ' Odemknut� listu pro vlo�en� dat
    ws.Unprotect "1234"

    ' Zkop�rov�n� dat z extern�ho souboru do c�lov�ho rozsahu
    srcRange.Copy targetRange
    
    ' P�eform�tov�n� ��sla
    Dim cell As Range
    For Each cell In targetRange.Cells
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
    
    HideButton ws, "Vlo�it hodnoty"
    HideButton ws, "Nahr�t hodnoty"
    
    ' P�id�n� tla��tka pro �pravu vypln�n�ch hodnot
    AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Upravit hodnoty", "EditCellValue"
    
    ' P�id�n� tla��tka pro spu�t�n� metody WSA
    AddButtonTo ws, ws.Range("B" & 9 + numOfCriteria), "Metoda WSA", "M3_metoda_WSA"
    
    ' P�id�n� tla��tka pro spu�t�n� metody bazick� varianty s v�t�� ���kou
    AddButtonTo ws, ws.Range("D" & 9 + numOfCriteria, "E" & 9 + numOfCriteria), "Metoda bazick� varianty", "M4_metoda_Bazicke_varianty", 4.5, 1
    
    ' Uzamknut� listu po vlo�en� dat
    ws.Protect "1234"

    MsgBox "Data byla �sp�n� nahr�na.", vbInformation
End Sub

' Skript pro kontrolu c�l�
Sub CheckObjectives(Objectives As Range, ws As Worksheet)
    Dim validObjectives As Boolean
    Dim cell As Range
    
    ' Kontrola, zda jsou v rozsahu pouze hodnoty "min" nebo "max"
        validObjectives = True
        
        ws.Unprotect "1234"
        
        For Each cell In Objectives
            
            'Nastaven� budouc� kontroly - v�b�rov� pole
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
            HideButton ws, "Stanovit c�le"
            
            ' Na�ten� po�tu krit�ri�
            numOfCriteria = ws.Range("C2").value
        
            ' P�id�n� tla��tka pokra�ovat
            AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Pokra�ovat", "Candidates"
        Else
            MsgBox "C�lem funkce m��e b�t pouze minimalizace (min) nebo maximalizace (max)!", vbExclamation
        End If
        
        ws.Protect "1234"
End Sub

' Skript pro kontrolu vah
Sub CheckWeights(weights As Range, ws As Worksheet)
    Dim sumWeights As Float
    
    ' Kontrola, zda jsou v�echny v�hy vypln�n�
    If CheckFilledCells(weights, "number") Then
        ' Z�sk�n� sou�tu vah
        sumWeights = Application.WorksheetFunction.Sum(weights)
        
        ' Zkontroluj, zda je sou�et roven 1 (100 %)
        If Not Round(sumWeights, 4) = 1 Then ' Pou��v�me zaokrouhlen� pro p�esnost
            MsgBox "Sou�et vah nen� roven 100%! Aktu�ln� sou�et: " & Format(sumWeights * 100, "0.00") & "%.", vbExclamation
        End If
    Else
        ' Pokud nejsou v�echny v�hy vypln�n�
        MsgBox "N�kter� v�hy nejsou vypln�n�.", vbExclamation
    End If
End Sub
