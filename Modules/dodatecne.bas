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
