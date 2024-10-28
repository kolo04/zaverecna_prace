Attribute VB_Name = "Module2"
Dim ws As Worksheet
Dim wsInput As Worksheet
Dim wsOutput As Worksheet
Dim numOfCriteria As Integer

Sub SetWeights()

    ' Otev�e formul�� s mo�nost� v�b�ru metody pro zad�n� vah
    WeightsForm.Show
    
End Sub

' Procedura obsluhuj�c� metodu po�ad�
Sub MoveToM2()

    ' Zobrazen� v�sledk� procedury a� po kompletn�m na�ten� procedury
    Application.ScreenUpdating = False

    ' Z�sk�n� po�tu krit�ri�
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    numOfCriteria = ws.Range("C2").value
    
    If numOfCriteria < 2 Then
        MsgBox "P�i rozhodov�n� bychom m�li zohled�ovat minim�ln� 2 krit�ria.", vbExclamation
        Exit Sub
        End
    End If
    
    ' Kontrola existence a vy�i�t�n� listu "Po�ad� krit�ri�"
    wsExists = False
    Set ws = Nothing
    
    For Each ws In ThisWorkbook.Sheets
        If ws.name = "Po�ad� krit�ri�" Then
            wsExists = True
            ws.Activate
            ws.Protect "1234", UserInterfaceOnly:=True
            ws.Cells.Clear
            ActiveSheet.Buttons.Delete
            
            Exit For
        End If
    Next ws
    
    ' Vytvo�en� listu, pokud neexistuje
    If Not wsExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = "Po�ad� krit�ri�"
        ws.Activate
        ws.Unprotect "1234"
    End If
    
    Call M2_metoda_poradi

End Sub

' Procedura pracuj�c� s listem Po�ad� krit�ri�
Sub M2_metoda_poradi()

    Application.ScreenUpdating = False
    
    ' Definice pracovn�ho listu
    Set ws = ThisWorkbook.Sheets("Po�ad� krit�ri�")
    
    With ws
        ' Vytvo�en� z�hlav� tabulky
        .Range("B2").value = "Krit�rium"
        .Range("C2").value = "Po�ad�"
        .Range("B2:C2").Font.Bold = True
    End With
    
    ' Mo�nost za��t znovu
    AddButtonTo ws, ws.Range("G2"), "Aktualizovat", "OrderList"
    
    ' Zavol�n� skriptu OrderList pro po�ad�
    Call OrderList
End Sub

' Procedura obsluhuj�c� vytvo�en� rozev�rac�ho seznamu,
' ve kter�m u�ivatel vyb�r� po�ad� sv�ch priorit pro krit�ria
Sub OrderList()

    Application.ScreenUpdating = False
    
    ' V�pis po�ad� a tla��tka pro v�po�et v�hy
    Dim changedRows As Collection
    Dim i As Integer
    Dim rowIndex As Variant
    
    Set ws = ThisWorkbook.Sheets("Po�ad� krit�ri�")
    Set wsInput = ThisWorkbook.Sheets("Vstupn� data")
    numOfCriteria = wsInput.Range("C2").value
    Set changedRows = New Collection            ' Kolekce k uchov�n� zm�n�n�ch ��dk�
    
    With ws
        .Unprotect "1234"
        ' Vy�i�t�n� obsahu sloupc� D a E
        .Columns("D:E").ClearContents

        ' Kontrola zm�n a aktualizace hodnot ve sloupci B
        For i = 3 To 2 + numOfCriteria
            If .Cells(i, 2).value <> wsInput.Cells(i + 2, 2).value Then
                .Cells(i, 2).value = wsInput.Cells(i + 2, 2).value
                changedRows.Add i
            End If
        Next i
        
        ' Nastaven� valida�n�ho seznamu a vymaz�n� odpov�daj�c�ch bun�k ve sloupci C
        If changedRows.Count > 0 Then
            For Each rowIndex In changedRows
                .Cells(rowIndex, 3).value = "Vyberte"
                ' Deklarace dynamick�ho pole �etezc�
                Dim validationArray() As String
                
                ' P�izp�soben� velikosti pole podle po�tu krit�ri�
                ReDim validationArray(numOfCriteria - 1)
                
                ' Cyklus p�id�vaj�c� jednotliv� krit�ria do pole
                For j = 1 To numOfCriteria
                    validationArray(j - 1) = j
                Next j
                
                ' Odstran�n� jak�koliv dosavadn� validace a p�id�n� validace pro v��et hodnot
                With .Cells(rowIndex, 3).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(validationArray, ",")
                End With
                
                ' Nastaven� form�tu bu�ky na ��seln� form�t
                .Cells(rowIndex, 3).NumberFormat = "General"
            Next rowIndex
        End If
        
        ' �prava ���ky sloupc� (Autofit na minim�ln� 80px)
        AdjustColumnWidth ws, ws.Range(ws.Columns(2), ws.Columns(3))
        
        HideButton ws, "Pokra�ovat"
        
        ' P�id�n� tla��tka "Vypo��tat v�hu"
        AddButtonTo ws, .Range("G9"), "Vypo��tat v�hy", "CountWeight"
        
        ' Aktivace bu�ky s rozev�rac�m seznamem
        .Range(.Cells(3, 3), .Cells(2 + numOfCriteria, 3)).Locked = False
        .Cells(3, 3).Select
        
        .Protect "1234"
        
    End With
End Sub

' Procedura obsluhuj�c� v�po�et v�hy
Sub CountWeight()
    Dim i As Integer, j As Integer
    Dim ranks As Object, allRanks As Object, filledRankPoints As Object
    Dim filledRanks As Collection
    Dim value As Variant, rankList() As Variant
    Dim rankSum As Double, rankPoints As Double
    Dim rankIndex As Integer, currentRank As Integer, totalRanks As Integer, rankPos As Integer
    Dim formula As String
    
    ' Inicializace list�
    Set ws = ThisWorkbook.Sheets("Po�ad� krit�ri�")
    Set wsInput = ThisWorkbook.Sheets("Vstupn� data")
    Set wsOutput = ThisWorkbook.Sheets("Vstupn� data")
    
    ' Na�ten� po�tu krit�ri�
    numOfCriteria = wsInput.Range("C2").value
    
    ' Inicializace slovn�k� a kolekc� pro uchov�n� hodnocen�
    Set ranks = CreateObject("Scripting.Dictionary")        ' Slovn�k k uchov�n� po�tu v�skyt� po�ad�
    Set allRanks = CreateObject("Scripting.Dictionary")     ' Slovn�k p�ed�vaj�c� index a hodnotu po�ad�
    Set filledRanks = New Collection                        ' Kolekce uchov�vaj�c� seznam obsazen�ch pozic
    Set filledRankPoints = CreateObject("Scripting.Dictionary") ' Slovn�k uchov�vaj�c� vzorec pro v�po�et bod� z po�ad�
    
    ' Ov��en�, zda jsou v�echna pole ve sloupci 3 vybr�na spr�vn�
    For i = 1 To numOfCriteria
        value = ws.Cells(2 + i, 3).value
        
        ' Pokud nen� vybr�na varianta
        If value = "Vyberte" Or IsEmpty(value) Then
            ws.Cells(2 + i, 3).Select
            MsgBox "Vypl�te pros�m v�echna po�ad� krit�ri�.", vbExclamation
            Exit Sub
            
        ' Ne��seln� nebo nespr�vn� hodnota
        ElseIf Not IsNumeric(value) Or value < 1 Or value > numOfCriteria Then
            ws.Cells(2 + i, 3).Select
            MsgBox "Po�ad� mus� b�t ��slo mezi 1 a " & numOfCriteria & ".", vbExclamation
            Exit Sub
            
        ' Po�ad� ji� evid�no
        ElseIf ranks.exists(value) Then
            ranks(value) = ranks(value) + 1
            
        ' P�id�n� nov� hodnoty po�ad�
        Else
            ranks.Add value, 1
        End If
        
        ' P�id�n� dvojice index a hodnota do kolekce
        allRanks.Add i, value
    Next i
    
    ' Kontrola, zda je jedni�ka ve sloupci po�ad�
    If Not ranks.exists(1) Then
        MsgBox "Po�ad� mus� za��nat od 1.", vbExclamation
        Exit Sub
    End If
    
    ' Vytvo�en� seznamu obsazen�ch po�ad�
    For Each Key In ranks.Keys
        filledRanks.Add Key
    Next Key
    
    ' P�id�n� chyb�j�c�ch po�ad� pro rozd�len� bod�
    For i = 1 To numOfCriteria
        If Not ranks.exists(i) Then
            filledRanks.Add i
        End If
    Next i
    
    ' P�eveden� kolekce filledRanks na pole
    ReDim rankList(filledRanks.Count - 1)
    rankIndex = 0
    For Each rank In filledRanks
        rankList(rankIndex) = rank
        rankIndex = rankIndex + 1
    Next rank

    ' V�po�et bod� a jejich p�i�azen�
    With ws
        ' Nadpis pro sloupec Bod�
        .Unprotect "1234"
        .Cells(2, 4).value = "Bod�"

        ' Po�et v�ech mo�n�ch po�ad�
        totalRanks = UBound(rankList) + 1
        
        ' Po��te�n� pozice hodnocen�
        rankPos = 1
        
        ' Cyklus proch�zej�c� v�echny hodnocen� pozice
        For i = 0 To UBound(rankList)
            currentRank = rankList(i)
            
            ' V�po�et bod� pro duplicitn� po�ad�
            If ranks(currentRank) > 1 Then
                rankSum = 0
                formula = ""
                
                ' Pro ka�d� duplicitn� po�ad� vypo��t� celkov� sou�et bod� a p�iprav� vzorec
                For j = 0 To ranks(currentRank) - 1
                    rankSum = rankSum + (totalRanks + 1 - (rankPos + j))
                    
                    ' Ko�en vzorce pro v�po�et bod�
                    If j <> ranks(currentRank) - 1 Then
                        formula = formula & (totalRanks + 1 - (rankPos + j)) & " + "
                    Else
                        formula = formula & (totalRanks + 1 - (rankPos + j))
                    End If
                Next j
                
                ' V�po�et pr�m�rn�ho bodov�ho hodnocen� pro duplicitn� po�ad�
                formula = "= (" & formula & ") / " & ranks(currentRank)
                filledRankPoints(currentRank) = formula
                rankPos = rankPos + ranks(currentRank)
                
            ' V�po�et bod� pro jedine�n� po�ad�
            Else
                rankPoints = totalRanks + 1 - rankPos
                formula = "= (" & numOfCriteria & " + 1 - " & rankPos & ") / 1"
                filledRankPoints(currentRank) = formula
                rankPos = rankPos + 1
            End If
        Next i
        
        ' P�i�azen� vzorc� do bun�k
        For Each Key In allRanks.Keys
            .Cells(2 + Key, 4).formula = filledRankPoints(allRanks(Key))
        Next Key
    
        ' Nadpis pro sloupec V�ha
        .Cells(2, 5).value = "V�ha"
        
        ' V�po�et v�hy jako pod�l bod� a celkov�ho po�tu bod�
        .Range("E3:E" & 2 + numOfCriteria).formula = "=$D3/(SUM($D$3:$D$" & (2 + numOfCriteria) & "))"
        
        ' Form�tov�n� procentu�ln�ho stylu s jedn�m desetinn�m m�stem
        .Range("E3:E" & 2 + numOfCriteria).Style = "Percent"
        .Range("E3:E" & 2 + numOfCriteria).NumberFormat = "0.0 %"
        
        ' Tu�n� p�smo pro z�hlav�
        .Range("B2:E2").Font.Bold = True
        
        ' Ohrani�en� pro nadpisy sloupc�
        With .Range("B2:E2").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        ' �prava ���ky sloupc� (Autofit na minim�ln� 80px)
        AdjustColumnWidth ws, .Range(.Columns(4), .Columns(5))
    End With
    
    HideButton ws, "Vypo��tat v�hy"
    
    ' Vlo�en� popisku a vah krit�ri� z listu Po�ad� krit�ri� do bun�k D5:D (5 + numOfCriteria)
    wsOutput.Unprotect "1234"
    wsOutput.Range("D5:D" & 4 + numOfCriteria).value = ws.Range("E3:E" & 2 + numOfCriteria).value
    HideButton wsOutput, "Stanovit v�hy"
    AdjustColumnWidth wsOutput, wsOutput.Range(wsOutput.Columns(2), wsOutput.Columns(4))
    wsOutput.Protect "1234"
    
    ' P�id�n� tla��tka pro n�vrat na vstupn� data
    AddButtonTo ws, ws.Range("G9"), "Pokra�ovat", "SetObjectives"
    
    ws.Protect "1234"
End Sub

Sub UploadWeights()
    Dim weights As Range
    Dim sumWeights As Double
    Dim Objectives As Range
    Dim subject As String

    ' Odkaz na list "Vstupn� data"
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    
    ' Na�ten� po�tu krit�ri�
    numOfCriteria = ws.Range("C2").value

    ' Nastaven� c�lov� bu�ky (za��tek v D5 + po�et krit�ri�)
    Set weights = ws.Range(ws.Cells(5, 4), ws.Cells(4 + numOfCriteria, 4))

    ' P�edm�t pro zobrazen� v InputBoxu
    subject = "v�hy"
    
    ' Vol�n� samostatn� procedury pro nahr�v�n� dat
    Call UploadData(weights, subject)
    
 ' Kontrola, zda jsou v�echny v�hy vypln�n�
    If CheckFilledCells(weights, "number") Then
        
        ' Kontrola, zda sou�et vah je roven 100%
        ' Z�sk�n� sou�tu vah
        sumWeights = Application.WorksheetFunction.Sum(weights)

        ' Zkontroluj, zda je sou�et roven 1 (100 %)
        If Round(sumWeights, 4) = 1 Then ' Pou��v�me zaokrouhlen� pro p�esnost
        ' Schov�n� tla��tka, pokud existuje
            HideButton ws, "Stanovit v�hy"
            
            'Kontrola, zda jsou vypln�ny c�le
            Set Objectives = ws.Range(ws.Cells(5, 3), ws.Cells(4 + numOfCriteria, 3))
            
            ws.Unprotect "1234"
            
            If CheckFilledCells(Objectives, "text") Then
                Call CheckObjectives(Objectives, ws)
            Else:
                ' P�id�n� tla��tka pro stanoven� c�l�
                AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Stanovit c�le", "SetObjectives"
            End If
        Else
            ' Pokud sou�et vah nen� roven 100%, tla��tko z�st�v�
            MsgBox "Sou�et vah nen� roven 100%! Aktu�ln� sou�et: " & Format(sumWeights * 100, "0.00") & "%.", vbExclamation
        End If
    Else
        ' Pokud nejsou v�echny v�hy vypln�n�, tla��tko z�st�v�
        MsgBox "Vkl�dan� v�hy nejsou vypln�n� nebo nejsou ve tvaru ��sla." & vbCrLf & "Nahr�v�n� bylo zru�eno.", vbExclamation
        ws.Unprotect "1234"
        ws.Range(ws.Cells(5, 4), ws.Cells(4 + numOfCriteria, 4)).Clear
    End If
    
    With ws
    
        .Unprotect "1234"
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
        
        .Protect "1234"
    End With

End Sub
