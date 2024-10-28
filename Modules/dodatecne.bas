Attribute VB_Name = "Module6"
Function IsUniqueValue(rng As Range, value As Variant) As Boolean
' Funkce pro ovìøení unikátnosti hodnoty
' Parametry jsou rozsah dat a hodnota buòky
' Návratová hodnota je Boolean
'
    ' Deklarace dimenze, datovým typem je rozsah
    Dim cell As Range
    
    ' Nastavení hodnoty funkce
    IsUniqueValue = True
    
    ' Cyklus pro prohledání všech bunìk v rozsahu
    For Each cell In rng
        
        ' Podmínka, zda se hodnota vybrané buòky z rozsahu rovná zkoumané hodnotì
        If cell.value = value Then
            
            ' Pokud ano, nastavení návratové hodnoty funkce na Nepravda
            IsUniqueValue = False
            
            ' Ukonèení funkce
            Exit Function
            
        ' Pokud ne, pøechod na další buòku v rozsahu
        End If
    Next cell
End Function

Sub AdjustColumnWidth(ByVal ws As Worksheet, ByVal columnRange As Variant)
' Skript pro upravení šíøky sloupce na minimální hodnotu 80 pixelù (Excel default) nebo Autofit
' Parametry jsou list výstupu a rozsah sloupcù

    Dim rng As Range
    Dim minColumnWidth As Double
    Dim column As Range
    
    ' Kontrola typu columnRange
    If TypeName(columnRange) = "Range" Then
        ' Pokud je columnRange typu Range, nastavím ho jako rozsah sloupcù
        Set rng = columnRange
    Else
        ' Pokud není columnRange typu Range, pøevedu ho na rozsah sloupcù na základì èísla sloupce
        Set rng = ws.Columns(columnRange)
    End If
    
    ' Autofit pro sloupce v rozsahu
    rng.Columns.AutoFit
    
    ' Nastavení minimální šíøky sloupce na 80 pixelù (8.11 cm)
     minColumnWidth = 8.11 ' Pøepoèet na šíøku sloupce v jednotkách Excelu
    
    ' Nastavení minimální šíøky sloupce
    For Each column In rng.Columns
        ' Reálná šíøka sloupce (cm) po Autofit
        If column.ColumnWidth < minColumnWidth Then
            column.ColumnWidth = minColumnWidth
        End If
    Next column
    
End Sub

Sub AddButtonTo(ws As Worksheet, position As Range, buttonText As String, macroName As String, Optional buttonWidth As Double = 3.75, Optional buttonHeight As Double = 1)
' Skript pro pøidání tlaèítka
' Parametry jsou list výstupu, pozice (a už absolutní nebo buòka), popisek a pøiøazené makro
'
    Dim btn As Button
    Dim btnExists As Boolean
    btnExists = False
    
    ' Cyklus pro všechna tlaèítka na listu
    For Each btn In ws.Buttons
        ' Pokud tlaèítko existuje na stejné pozici, oznaèí ho jako existující
        If btn.Top = position.Top And btn.Left = position.Left Then
            btnExists = True
            Exit For
        End If
    Next btn
    
    ' Pokud tlaèítko existuje, smaže ho
    If btnExists Then
        btn.Delete
    End If
    
    ' Vytvoøí nové tlaèítko, rozmìry jsou 3.5 cm x 1 cm
    ws.Unprotect "1234"
    Set btn = ws.Buttons.Add(position.Left, position.Top, buttonWidth * 28.35, buttonHeight * 28.35)
    
    ' Nastavení popisku tlaèítka
    btn.Text = buttonText
    
    ' Pøiøazení makra k tlaèítku
    btn.OnAction = macroName
End Sub

' Skript pro schování tlaèítka
Sub HideButton(ws As Worksheet, ByVal buttonText As String)

    Dim btn As Button
    
    ' Cyklus pro všechna tlaèítka na listu
    For Each btn In ws.Buttons
        ' Pokud text tlaèítka odpovídá hledanému textu
        If btn.Text = buttonText Then
            ' Skryje tlaèítko
            btn.Visible = False
             ' Ukonèení funkce po nalezení prvního tlaèítka se shodným textem
            Exit Sub
        End If
    Next btn
End Sub

' Skript obsluhující pøidání tlaèítka pro vytvoøení nového pøíkladu
Sub AddRestartButton()
    Dim ws As Worksheet
    Dim btn As Shape
    Dim buttonText As String
    Dim macroName As String
    Dim buttonWidth As Double
    Dim buttonHeight As Double
    Dim buttonTop As Double
    Dim buttonLeft As Double

    ' Nastavení pracovního listu
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    
    ' Parametry tlaèítka
    buttonText = "Nový" & vbCrLf & "pøíklad" ' Text rozdìlený na dva øádky
    macroName = "auto_open"
    buttonWidth = 2.069 * 28.35 ' Rozmìry tlaèítka v pixelech (pøevod na cm)
    buttonHeight = 1.69 * 28.35

    ' Nastavení pozice tlaèítka na listu
    buttonTop = ws.Cells(1, 1).Top + 10 ' 10 pixelù od vrchu buòky
    buttonLeft = ws.Cells(1, 1).Left + 14 ' 14 pixelù od levého okraje buòky

    ' Smazání existujícího tlaèítka, pokud již existuje
    On Error Resume Next
    ws.Shapes("RestartButton").Delete
    On Error GoTo 0

    ' Vytvoøení tlaèítka s urèenými parametry
    Set btn = ws.Shapes.AddShape(msoShapeBevel, buttonLeft, buttonTop, buttonWidth, buttonHeight)

    ' Pojmenování tlaèítka pro pozdìjší odstranìní
    btn.name = "RestartButton"

    ' Nastavení textu tlaèítka
    btn.TextFrame2.TextRange.Text = buttonText

    ' Formátování tlaèítka
    With btn.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2 ' Použití barev podle tématu
        .Solid
    End With

    ' Nastavení barvy obrysu
    With btn.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorLight1 ' Použití barev podle tématu
        .Weight = 0.5
    End With

    ' Nastavení stylu písma ve tlaèítku
    With btn.TextFrame2.TextRange.Font
        .Size = 11  ' Velikost textu
        .Bold = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1 ' Použití barev podle tématu pro text
    End With

    ' Vertikální zarovnání textu na støed
    btn.TextFrame2.VerticalAnchor = msoAnchorMiddle
    btn.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

    ' Pøiøazení makra k tlaèítku
    btn.OnAction = macroName

    ' Úprava šíøky sloupce A
    ws.Columns("A").ColumnWidth = 15
End Sub

' Skript pro nahrávání dat z vybrané oblasti (z libovolného sešitu)
Public Function UploadData(rng As Range, subject As String, Optional insertAsRow As Boolean = False) As Integer
    Dim validSelection As Boolean
    Dim srcRange As Range
    Dim transposedData As Variant
    Dim numOfUnits As Integer
    Dim ws As Worksheet

' Smyèka pro opakovaný výbìr, dokud nebude platný
RestartLoop:
    validSelection = False ' Inicializace promìnné pro platnost výbìru
    Set srcRange = Nothing
    
    ' Smyèka pro opakovaný výbìr, dokud nebude platný
    Do While Not validSelection
        ' Získání vstupu od uživatele pomocí InputBoxu s možností výbìru oblasti myší
        On Error Resume Next
        Set srcRange = Application.InputBox("Vyberte oblast dat, odkud chcete " & subject & " nahrát:", "Vyberte rozsah dat", Type:=8)
        On Error GoTo 0

        ' Kontrola, zda uživatel nìco vybral
        If srcRange Is Nothing Then
            MsgBox "Nebyla vybrána žádná oblast.", vbExclamation
            UploadData = 0 ' V pøípadì, že uživatel nevybral oblast, vrátí 0
            Exit Function
        Else
            ' Kontrola, zda uživatel vybral pouze jeden øádek nebo jeden sloupec
            If srcRange.Rows.Count > 1 And srcRange.Columns.Count > 1 Then
                MsgBox "Vyberte pouze jeden øádek nebo jeden sloupec dat, odkud chcete " & subject & " nahrát!", vbExclamation
                GoTo RestartLoop
            Else
                ' Kontrola prázdných bunìk
                hasEmpty = False
                For Each cell In srcRange
                    If IsEmpty(cell.value) Then
                        hasEmpty = True
                        Exit For
                    End If
                Next cell
                
                If hasEmpty Then
                    MsgBox "Vybraný rozsah obsahuje prázdné buòky. Vyberte, prosím, jiný rozsah.", vbExclamation
                    GoTo RestartLoop
                Else
                    ' Pokud jsou výše uvedené podmínky splnìné, nastavení výbìru jako platného
                    validSelection = True
                End If
            End If
        End If
    Loop
    
    ' Získání informací o listu, kam se data vkládají
    Set ws = rng.Worksheet
    
    ' Naètení poètu kritérií
    numOfCriteria = ws.Range("C2").value
    
    ' Kontrola poètu vložených øádkù pro "cíle" nebo "váhy" proti poètu kritérií
    If subject = "cíle" Or subject = "váhy" Then
        If srcRange.Rows.Count <> numOfCriteria Then
            MsgBox "Poèet vložených øádkù musí odpovídat poètu kritérií (" & numOfCriteria & "). Vyberte, prosím, správný rozsah.", vbExclamation
            GoTo RestartLoop
        End If
    End If
    
    ' Odemknutí listu pro kopírování dat
    ws.Unprotect "1234"

    ' Pokud vkládáme data jako øádek, ale uživatel zadal data ve sloupci, pøevedeme je na øádek a naopak
    If insertAsRow And srcRange.Columns.Count = 1 Then
        ' Data zadána ve sloupci pøevedena na øádek
        transposedData = Application.WorksheetFunction.Transpose(srcRange.value)
        
        ' Ošetøení podle množství vkládaných bunìk
        If IsArray(transposedData) Then
            ' Pøidání apostrofu, pokud jde o varianty
            If subject = "varianty" Then
                For i = LBound(transposedData) To UBound(transposedData)
                    transposedData(i) = "'" & transposedData(i)
                Next i
            End If
            
            ' Úprava cílového rozsahu pro více bunìk
            Set rng = rng.Resize(1, UBound(transposedData, 1))
            rng.value = transposedData ' Zápis transponovaných dat do cílového rozsahu
            numOfUnits = UBound(transposedData, 1) ' Poèet transponovaných jednotek (nový poèet sloupcù)
        Else
            ' Ošetøení, pokud je vkládána jen jedna buòka
            If subject = "varianty" Then
                transposedData = "'" & transposedData ' Pøidání apostrofu pro jednu buòku
            End If
            
            ' Pøiøazení hodnoty do cílové buòky
            rng.value = transposedData
            numOfUnits = 1
        End If
        
    ElseIf Not insertAsRow And srcRange.Rows.Count = 1 Then
        ' Data zadána v øádku pøevedena na sloupec
        transposedData = Application.WorksheetFunction.Transpose(srcRange.value)
        
        ' Ošetøení podle množství vkládaných bunìk
        If IsArray(transposedData) Then
            ' Pokud jde o více bunìk, upravíme hodnoty a pøidáme apostrof, pokud jde o kritéria
            If subject = "kritéria" Then
                For i = LBound(transposedData) To UBound(transposedData)
                    transposedData(i, 1) = "'" & transposedData(i, 1)
                Next i
            End If
            
            ' Úprava cílového rozsahu pro více bunìk
            Set rng = rng.Resize(UBound(transposedData, 1), 1)
            rng.value = transposedData
            numOfUnits = UBound(transposedData, 1)
        Else
            ' Ošetøení, pokud je vkládána jen jedna buòka
            If subject = "kritéria" Then
                transposedData = "'" & transposedData ' Pøidání apostrofu pro jednu buòku
            End If
            
            ' Pøiøazení hodnoty do cílové buòky
            rng.value = transposedData
            numOfUnits = 1
        End If
    Else
        ' Pokud není potøeba transpozice, upravíme hodnoty pøed pøímým vložením
        If subject = "kritéria" Or subject = "varianty" Then
            For Each cell In srcRange
                cell.value = "'" & cell.value
            Next cell
        End If
        
        srcRange.Copy rng
        
        If subject = "varianty" Then
            ' Poèet øádkù v pøípadì vkládání do øádku (pro varianty)
            numOfUnits = srcRange.Columns.Count
        Else
            ' Poèet øádkù v pøípadì vkládání jako sloupec
            numOfUnits = srcRange.Rows.Count
        End If
        
    End If

    ' Uzamknutí listu po dokonèení
    ws.Protect "1234"

    ' Vrácení poètu vložených jednotek (buï øádkù nebo sloupcù)
    UploadData = numOfUnits
End Function

' Skript pro kontrolu (ne)vyplnìných bunìk
' Parametry jsou rozsah a typ dat
Function CheckFilledCells(rng As Range, dataType As String) As Boolean
    Dim cell As Range
    Dim isFilled As Boolean
    isFilled = True ' Pøedpokládáme, že všechny buòky jsou vyplnìné

    ' Procházíme všechny buòky v zadaném rozsahu
    For Each cell In rng
        ' Kontrola na základì oèekávaného typu dat
        Select Case dataType
            Case "number"
                ' Pokud je buòka prázdná nebo neobsahuje èíslo, nastavíme isFilled na False
                If IsEmpty(cell) Or Not IsNumeric(cell.value) Then
                    isFilled = False
                    Exit For
                End If
            Case "text"
                ' Pokud je buòka prázdná nebo neobsahuje text, nastavíme isFilled na False
                If IsEmpty(cell) Or VarType(cell.value) <> vbString Then
                    isFilled = False
                    Exit For
                End If
            Case Else
                ' Neoèekávaný typ dat
                MsgBox "Neplatný typ dat: " & dataType, vbExclamation
                isFilled = False
                Exit Function
        End Select
    Next cell

    ' Vrátíme výsledek, zda jsou všechny buòky vyplnìné
    CheckFilledCells = isFilled
End Function
