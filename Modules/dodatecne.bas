Attribute VB_Name = "Module6"
Function IsUniqueValue(rng As Range, value As Variant) As Boolean
' Funkce pro ov��en� unik�tnosti hodnoty
' Parametry jsou rozsah dat a hodnota bu�ky
' N�vratov� hodnota je Boolean
'
    ' Deklarace dimenze, datov�m typem je rozsah
    Dim cell As Range
    
    ' Nastaven� hodnoty funkce
    IsUniqueValue = True
    
    ' Cyklus pro prohled�n� v�ech bun�k v rozsahu
    For Each cell In rng
        
        ' Podm�nka, zda se hodnota vybran� bu�ky z rozsahu rovn� zkouman� hodnot�
        If cell.value = value Then
            
            ' Pokud ano, nastaven� n�vratov� hodnoty funkce na Nepravda
            IsUniqueValue = False
            
            ' Ukon�en� funkce
            Exit Function
            
        ' Pokud ne, p�echod na dal�� bu�ku v rozsahu
        End If
    Next cell
End Function

Sub AdjustColumnWidth(ByVal ws As Worksheet, ByVal columnRange As Variant)
' Skript pro upraven� ���ky sloupce na minim�ln� hodnotu 80 pixel� (Excel default) nebo Autofit
' Parametry jsou list v�stupu a rozsah sloupc�

    Dim rng As Range
    Dim minColumnWidth As Double
    Dim column As Range
    
    ' Kontrola typu columnRange
    If TypeName(columnRange) = "Range" Then
        ' Pokud je columnRange typu Range, nastav�m ho jako rozsah sloupc�
        Set rng = columnRange
    Else
        ' Pokud nen� columnRange typu Range, p�evedu ho na rozsah sloupc� na z�klad� ��sla sloupce
        Set rng = ws.Columns(columnRange)
    End If
    
    ' Autofit pro sloupce v rozsahu
    rng.Columns.AutoFit
    
    ' Nastaven� minim�ln� ���ky sloupce na 80 pixel� (8.11 cm)
     minColumnWidth = 8.11 ' P�epo�et na ���ku sloupce v jednotk�ch Excelu
    
    ' Nastaven� minim�ln� ���ky sloupce
    For Each column In rng.Columns
        ' Re�ln� ���ka sloupce (cm) po Autofit
        If column.ColumnWidth < minColumnWidth Then
            column.ColumnWidth = minColumnWidth
        End If
    Next column
    
End Sub

Sub AddButtonTo(ws As Worksheet, position As Range, buttonText As String, macroName As String, Optional buttonWidth As Double = 3.75, Optional buttonHeight As Double = 1)
' Skript pro p�id�n� tla��tka
' Parametry jsou list v�stupu, pozice (a� u� absolutn� nebo bu�ka), popisek a p�i�azen� makro
'
    Dim btn As Button
    Dim btnExists As Boolean
    btnExists = False
    
    ' Cyklus pro v�echna tla��tka na listu
    For Each btn In ws.Buttons
        ' Pokud tla��tko existuje na stejn� pozici, ozna�� ho jako existuj�c�
        If btn.Top = position.Top And btn.Left = position.Left Then
            btnExists = True
            Exit For
        End If
    Next btn
    
    ' Pokud tla��tko existuje, sma�e ho
    If btnExists Then
        btn.Delete
    End If
    
    ' Vytvo�� nov� tla��tko, rozm�ry jsou 3.5 cm x 1 cm
    Set btn = ws.Buttons.Add(position.Left, position.Top, buttonWidth * 28.35, buttonHeight * 28.35)
    
    ' Nastaven� popisku tla��tka
    btn.Text = buttonText
    
    ' P�i�azen� makra k tla��tku
    btn.OnAction = macroName
End Sub

' Skript pro schov�n� tla��tka
Sub HideButton(ws As Worksheet, ByVal buttonText As String)

    Dim btn As Button
    
    ' Cyklus pro v�echna tla��tka na listu
    For Each btn In ws.Buttons
        ' Pokud text tla��tka odpov�d� hledan�mu textu
        If btn.Text = buttonText Then
            ' Skryje tla��tko
            btn.Visible = False
             ' Ukon�en� funkce po nalezen� prvn�ho tla��tka se shodn�m textem
            Exit Sub
        End If
    Next btn
End Sub
