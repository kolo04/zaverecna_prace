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
    ws.Unprotect "1234"
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

' Skript obsluhuj�c� p�id�n� tla��tka pro vytvo�en� nov�ho p��kladu
Sub AddRestartButton()
    Dim ws As Worksheet
    Dim btn As Shape
    Dim buttonText As String
    Dim macroName As String
    Dim buttonWidth As Double
    Dim buttonHeight As Double
    Dim buttonTop As Double
    Dim buttonLeft As Double

    ' Nastaven� pracovn�ho listu
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    
    ' Parametry tla��tka
    buttonText = "Nov�" & vbCrLf & "p��klad" ' Text rozd�len� na dva ��dky
    macroName = "auto_open"
    buttonWidth = 2.069 * 28.35 ' Rozm�ry tla��tka v pixelech (p�evod na cm)
    buttonHeight = 1.69 * 28.35

    ' Nastaven� pozice tla��tka na listu
    buttonTop = ws.Cells(1, 1).Top + 10 ' 10 pixel� od vrchu bu�ky
    buttonLeft = ws.Cells(1, 1).Left + 14 ' 14 pixel� od lev�ho okraje bu�ky

    ' Smaz�n� existuj�c�ho tla��tka, pokud ji� existuje
    On Error Resume Next
    ws.Shapes("RestartButton").Delete
    On Error GoTo 0

    ' Vytvo�en� tla��tka s ur�en�mi parametry
    Set btn = ws.Shapes.AddShape(msoShapeBevel, buttonLeft, buttonTop, buttonWidth, buttonHeight)

    ' Pojmenov�n� tla��tka pro pozd�j�� odstran�n�
    btn.name = "RestartButton"

    ' Nastaven� textu tla��tka
    btn.TextFrame2.TextRange.Text = buttonText

    ' Form�tov�n� tla��tka
    With btn.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2 ' Pou�it� barev podle t�matu
        .Solid
    End With

    ' Nastaven� barvy obrysu
    With btn.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorLight1 ' Pou�it� barev podle t�matu
        .Weight = 0.5
    End With

    ' Nastaven� stylu p�sma ve tla��tku
    With btn.TextFrame2.TextRange.Font
        .Size = 11  ' Velikost textu
        .Bold = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1 ' Pou�it� barev podle t�matu pro text
    End With

    ' Vertik�ln� zarovn�n� textu na st�ed
    btn.TextFrame2.VerticalAnchor = msoAnchorMiddle
    btn.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

    ' P�i�azen� makra k tla��tku
    btn.OnAction = macroName

    ' �prava ���ky sloupce A
    ws.Columns("A").ColumnWidth = 15
End Sub

' Skript pro nahr�v�n� dat z vybran� oblasti (z libovoln�ho se�itu)
Public Function UploadData(rng As Range, subject As String, Optional insertAsRow As Boolean = False) As Integer
    Dim validSelection As Boolean
    Dim srcRange As Range
    Dim transposedData As Variant
    Dim numOfUnits As Integer
    Dim ws As Worksheet

' Smy�ka pro opakovan� v�b�r, dokud nebude platn�
RestartLoop:
    validSelection = False ' Inicializace prom�nn� pro platnost v�b�ru
    Set srcRange = Nothing
    
    ' Smy�ka pro opakovan� v�b�r, dokud nebude platn�
    Do While Not validSelection
        ' Z�sk�n� vstupu od u�ivatele pomoc� InputBoxu s mo�nost� v�b�ru oblasti my��
        On Error Resume Next
        Set srcRange = Application.InputBox("Vyberte oblast dat, odkud chcete " & subject & " nahr�t:", "Vyberte rozsah dat", Type:=8)
        On Error GoTo 0

        ' Kontrola, zda u�ivatel n�co vybral
        If srcRange Is Nothing Then
            MsgBox "Nebyla vybr�na ��dn� oblast.", vbExclamation
            UploadData = 0 ' V p��pad�, �e u�ivatel nevybral oblast, vr�t� 0
            Exit Function
        Else
            ' Kontrola, zda u�ivatel vybral pouze jeden ��dek nebo jeden sloupec
            If srcRange.Rows.Count > 1 And srcRange.Columns.Count > 1 Then
                MsgBox "Vyberte pouze jeden ��dek nebo jeden sloupec dat, odkud chcete " & subject & " nahr�t!", vbExclamation
                GoTo RestartLoop
            Else
                ' Kontrola pr�zdn�ch bun�k
                hasEmpty = False
                For Each cell In srcRange
                    If IsEmpty(cell.value) Then
                        hasEmpty = True
                        Exit For
                    End If
                Next cell
                
                If hasEmpty Then
                    MsgBox "Vybran� rozsah obsahuje pr�zdn� bu�ky. Vyberte, pros�m, jin� rozsah.", vbExclamation
                    GoTo RestartLoop
                Else
                    ' Pokud jsou v��e uveden� podm�nky spln�n�, nastaven� v�b�ru jako platn�ho
                    validSelection = True
                End If
            End If
        End If
    Loop
    
    ' Z�sk�n� informac� o listu, kam se data vkl�daj�
    Set ws = rng.Worksheet
    
    ' Na�ten� po�tu krit�ri�
    numOfCriteria = ws.Range("C2").value
    
    ' Kontrola po�tu vlo�en�ch ��dk� pro "c�le" nebo "v�hy" proti po�tu krit�ri�
    If subject = "c�le" Or subject = "v�hy" Then
        If srcRange.Rows.Count <> numOfCriteria Then
            MsgBox "Po�et vlo�en�ch ��dk� mus� odpov�dat po�tu krit�ri� (" & numOfCriteria & "). Vyberte, pros�m, spr�vn� rozsah.", vbExclamation
            GoTo RestartLoop
        End If
    End If
    
    ' Odemknut� listu pro kop�rov�n� dat
    ws.Unprotect "1234"

    ' Pokud vkl�d�me data jako ��dek, ale u�ivatel zadal data ve sloupci, p�evedeme je na ��dek a naopak
    If insertAsRow And srcRange.Columns.Count = 1 Then
        ' Data zad�na ve sloupci p�evedena na ��dek
        transposedData = Application.WorksheetFunction.Transpose(srcRange.value)
        
        ' O�et�en� podle mno�stv� vkl�dan�ch bun�k
        If IsArray(transposedData) Then
            ' P�id�n� apostrofu, pokud jde o varianty
            If subject = "varianty" Then
                For i = LBound(transposedData) To UBound(transposedData)
                    transposedData(i) = "'" & transposedData(i)
                Next i
            End If
            
            ' �prava c�lov�ho rozsahu pro v�ce bun�k
            Set rng = rng.Resize(1, UBound(transposedData, 1))
            rng.value = transposedData ' Z�pis transponovan�ch dat do c�lov�ho rozsahu
            numOfUnits = UBound(transposedData, 1) ' Po�et transponovan�ch jednotek (nov� po�et sloupc�)
        Else
            ' O�et�en�, pokud je vkl�d�na jen jedna bu�ka
            If subject = "varianty" Then
                transposedData = "'" & transposedData ' P�id�n� apostrofu pro jednu bu�ku
            End If
            
            ' P�i�azen� hodnoty do c�lov� bu�ky
            rng.value = transposedData
            numOfUnits = 1
        End If
        
    ElseIf Not insertAsRow And srcRange.Rows.Count = 1 Then
        ' Data zad�na v ��dku p�evedena na sloupec
        transposedData = Application.WorksheetFunction.Transpose(srcRange.value)
        
        ' O�et�en� podle mno�stv� vkl�dan�ch bun�k
        If IsArray(transposedData) Then
            ' Pokud jde o v�ce bun�k, uprav�me hodnoty a p�id�me apostrof, pokud jde o krit�ria
            If subject = "krit�ria" Then
                For i = LBound(transposedData) To UBound(transposedData)
                    transposedData(i, 1) = "'" & transposedData(i, 1)
                Next i
            End If
            
            ' �prava c�lov�ho rozsahu pro v�ce bun�k
            Set rng = rng.Resize(UBound(transposedData, 1), 1)
            rng.value = transposedData
            numOfUnits = UBound(transposedData, 1)
        Else
            ' O�et�en�, pokud je vkl�d�na jen jedna bu�ka
            If subject = "krit�ria" Then
                transposedData = "'" & transposedData ' P�id�n� apostrofu pro jednu bu�ku
            End If
            
            ' P�i�azen� hodnoty do c�lov� bu�ky
            rng.value = transposedData
            numOfUnits = 1
        End If
    Else
        ' Pokud nen� pot�eba transpozice, uprav�me hodnoty p�ed p��m�m vlo�en�m
        If subject = "krit�ria" Or subject = "varianty" Then
            For Each cell In srcRange
                cell.value = "'" & cell.value
            Next cell
        End If
        
        srcRange.Copy rng
        
        If subject = "varianty" Then
            ' Po�et ��dk� v p��pad� vkl�d�n� do ��dku (pro varianty)
            numOfUnits = srcRange.Columns.Count
        Else
            ' Po�et ��dk� v p��pad� vkl�d�n� jako sloupec
            numOfUnits = srcRange.Rows.Count
        End If
        
    End If

    ' Uzamknut� listu po dokon�en�
    ws.Protect "1234"

    ' Vr�cen� po�tu vlo�en�ch jednotek (bu� ��dk� nebo sloupc�)
    UploadData = numOfUnits
End Function

' Skript pro kontrolu (ne)vypln�n�ch bun�k
' Parametry jsou rozsah a typ dat
Function CheckFilledCells(rng As Range, dataType As String) As Boolean
    Dim cell As Range
    Dim isFilled As Boolean
    isFilled = True ' P�edpokl�d�me, �e v�echny bu�ky jsou vypln�n�

    ' Proch�z�me v�echny bu�ky v zadan�m rozsahu
    For Each cell In rng
        ' Kontrola na z�klad� o�ek�van�ho typu dat
        Select Case dataType
            Case "number"
                ' Pokud je bu�ka pr�zdn� nebo neobsahuje ��slo, nastav�me isFilled na False
                If IsEmpty(cell) Or Not IsNumeric(cell.value) Then
                    isFilled = False
                    Exit For
                End If
            Case "text"
                ' Pokud je bu�ka pr�zdn� nebo neobsahuje text, nastav�me isFilled na False
                If IsEmpty(cell) Or VarType(cell.value) <> vbString Then
                    isFilled = False
                    Exit For
                End If
            Case Else
                ' Neo�ek�van� typ dat
                MsgBox "Neplatn� typ dat: " & dataType, vbExclamation
                isFilled = False
                Exit Function
        End Select
    Next cell

    ' Vr�t�me v�sledek, zda jsou v�echny bu�ky vypln�n�
    CheckFilledCells = isFilled
End Function
