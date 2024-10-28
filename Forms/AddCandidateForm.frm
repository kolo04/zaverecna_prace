VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddCandidateForm 
   Caption         =   "Formuláø pro pøidání variant"
   ClientHeight    =   3036
   ClientLeft      =   144
   ClientTop       =   408
   ClientWidth     =   5160
   OleObjectBlob   =   "AddCandidateForm.frx":0000
End
Attribute VB_Name = "AddCandidateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ws As Worksheet
Dim numOfCriteria As Integer
Dim numOfCandidates As Integer
Dim candidatesDone As Boolean

Private Sub UserForm_Initialize()

    ' Pøi inicializaci formuláøe bude TextBox1 aktivní pro vstup uživatele
    TextBox1.SetFocus
    
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    
    ' Schování tlaèítka, pokud existuje
    HideButton ws, "Pøidat variantu"
    
    ' Pøidání tlaèítka pro pøidání dalších variant
    AddButtonTo ws, ws.Cells(2, 8), "Pøidat variantu", "AddMoreCandidates"
    
    ThisWorkbook.Sheets("Vstupní data").Protect "1234"
    
    ' Nastavení velikosti (pùvodnì 160x269)
    With frm
        Height = 185
        Width = 269
    End With
    
End Sub

' Procedura ovládající tlaèítko Pøidat variantu, reaguje na stisknutí tlaèítka
Private Sub Add_Click()

' Pøidání nové varianty na list "Vstupní data"
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    
    ' Urèení sloupce pro novou variantu
    Dim lastColumn As Long
    lastColumn = ws.Cells(4, ws.Columns.Count).End(xlToLeft).column + 1
    
    Dim validInput As Boolean
    
    Do
        ' Zobrazení chybové zprávy, pokud je TextBox prázdný,
        If TextBox1.Text = "" Then
            MsgBox "Název varianty nesmí být prázdný.", vbExclamation
            
            ' Ukonèit proceduru, ale nechat formuláø otevøený
            TextBox1.SetFocus
            ThisWorkbook.Sheets("Vstupní data").Protect "1234"
            Exit Sub
        Else
            ' Znovunaètení aktuálního listu
            Set ws = ThisWorkbook.Sheets("Vstupní data")
            
            ' Získání aktuálního poètu variant
            numOfCandidates = ws.Range("F2").value
        
            ' Kontrola, zda se varianta se stejným názvem již nevyskytuje
            If Not IsUniqueValue(ws.Range(ws.Cells(4, 5), ws.Cells(4, 4 + numOfCandidates - 1)), TextBox1.Text) Then
                MsgBox "Varianty musí být unikátní!", vbExclamation
                
                ' Ukonèit proceduru, ale nechat formuláø otevøený
                TextBox1.SetFocus
                ThisWorkbook.Sheets("Vstupní data").Protect "1234"
               Exit Sub
            Else
                ' Platný vstup
                validInput = True
            End If
        End If
    Loop Until validInput
    
    ' Zapsání názvu varianty na list
    ws.Unprotect "1234"
    ws.Cells(4, lastColumn).value = "'" & TextBox1.Text
    
    ' Aktualizace poètu variant na listu
    ws.Range("F2").value = numOfCandidates + 1
    numOfCandidates = numOfCandidates + 1
    
    ' Vyprázdnìní pole pro název kritéria
    TextBox1.Text = ""
    
    Call Update

End Sub

' Skript umožòující nahrát varianty z vybrané oblasti
Private Sub Upload_Click()
    Dim rng As Range
    Dim subject As String
    Dim candidatesRange As Range
    Dim numOfColumns As Integer
    Dim duplicateFound As Boolean
    Dim cell As Range
    Dim UniqueValues As Object

    ' Odkaz na list "Vstupní data"
    Set ws = ThisWorkbook.Sheets("Vstupní data")

    ' Získání aktuálního poètu variant z buòky F2
    Set candidatesRange = ws.Range("F2")
    numOfCandidates = candidatesRange.value

    ' Nastavení cílové buòky (zaèátek v E4 + poèet variant)
    Set rng = ws.Cells(4, 5 + numOfCandidates)

    ' Pøedmìt pro zobrazení v InputBoxu
    subject = "varianty"

    ' Volání samostatné procedury pro nahrávání dat a získání poètu vložených sloupcù
    numOfColumns = UploadData(rng, subject, True)

    ' Ovìøení, zda došlo k úspìšnému nahrání dat
    If numOfColumns > 0 Then
        ' Slovník pro kontrolu duplicit v rámci novì nahraných dat
        Set UniqueValues = CreateObject("Scripting.Dictionary")

        ' Pomocná promìnná pro kontrolu duplicit
        duplicateFound = False

        ' Kontrola unikátnosti novì nahraných variant
        For Each cell In ws.Range(ws.Cells(4, 5 + numOfCandidates), ws.Cells(4, 4 + numOfCandidates + numOfColumns))
            If cell.value <> "" Then
                ' Kontrola existujících variant (pokud nìjaké existují)
                If numOfCandidates > 0 Then
                    If Not IsUniqueValue(ws.Range(ws.Cells(4, 5), ws.Cells(4, 4 + numOfCandidates)), cell.value) Then
                        duplicateFound = True
                        Exit For
                    End If
                End If

                ' Kontrola duplicit v rámci novì vkládaných dat
                If UniqueValues.exists(cell.value) Then
                    duplicateFound = True
                    Exit For
                Else
                    ' Pøidání hodnoty do slovníku novì vkládaných hodnot
                    UniqueValues.Add cell.value, True
                End If
            End If
        Next cell

        ' Zpracování výsledkù kontroly duplicit
        If duplicateFound Then
            MsgBox "Vkládané varianty musí být unikátní! Nahrávání bylo zrušeno.", vbExclamation
            ws.Unprotect "1234"
            ws.Range(ws.Cells(4, 5 + numOfCandidates), ws.Cells(4, 4 + numOfCandidates + numOfColumns)).Clear
        Else
            ' Aktualizace poètu variant o poèet vložených sloupcù
            ws.Unprotect "1234"
            candidatesRange.value = numOfCandidates + numOfColumns
        End If
    End If

    ' Uzamknutí listu na konci procedury
    ws.Protect "1234"

    Call Update
    
    ' Pøidání tlaèítka pro nový pøíklad
    Call AddRestartButton
    
    Unload Me
End Sub

' Spoleèný skript pro oba zpùsoby vkládání variant, obslouží potøebné úpravy formuláøe i listu
Private Sub Update()
    
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    ws.Activate
    
    ws.Unprotect "1234"
    
    ' Zrušení všech tlaèítek na listu
    ws.Buttons.Delete
    
    ' Získání aktuálního poètu kritérií a variant z listu
    numOfCriteria = ws.Range("C2").value
    numOfCandidates = ws.Range("F2").value
    
    ' Pøidání tlaèítka pro pøidání dalších kritérií
    AddButtonTo ws, ws.Range("B" & 6 + numOfCriteria), "Pøidat kritérium", "AddMoreCriteria"
    
    'Pøi jednom a více kritériu pøidat tlaèítko pro odebrání kritéria
    If numOfCriteria > 0 Then
        AddButtonTo ws, ws.Range("D" & 6 + numOfCriteria), "Odebrat kritérium", "RemoveCriteria"
    End If
    
    ' Pokud není vyplnìna váha u posledního pøidaného kritéria, pak pøidat tlaèítko pro Stanovení vah
    If IsEmpty(ws.Cells(4 + numOfCriteria, 4)) Then
        AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Stanovit váhy", "MoveToM2"
    Else
        ' Kontrola, zda jsou buòky prázdné pomocí funkce CountA (výpoèet prázdných bunìk)
        If WorksheetFunction.CountBlank(ws.Range(ws.Cells(5, 5), ws.Cells(4 + numOfCriteria, 4 + numOfCandidates))) > 0 Then
            ' Pøidání tlaèítka pro vyplnìní dat
            AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Vložit hodnoty", "FillData"
            
            ' Pøidání tlaèítka pro vyplnìní dat
            AddButtonTo ws, ws.Range("F" & 9 + numOfCriteria), "Nahrát hodnoty", "UploadDataBlock"
        Else
            ' Pøidání tlaèítka pro úpravu vyplnìných hodnot
            AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Upravit hodnoty", "EditCellValue"
            
            ' Pøidání tlaèítka pro spuštìní metody WSA
            AddButtonTo ws, ws.Range("B" & 9 + numOfCriteria), "Metoda WSA", "M3_metoda_WSA"
        
            ' Pøidání tlaèítka pro spuštìní metody bazické varianty s vìtší šíøkou
            AddButtonTo ws, ws.Range("D" & 9 + numOfCriteria, "E" & 9 + numOfCriteria), "Metoda bazické varianty", "M4_metoda_Bazicke_varianty", 4.5, 1

        End If
    End If
            
    ' Získání aktuálního poètu kritérií
    numOfCandidates = ws.Range("F2").value
    
    If Not IsEmpty(numOfCandidates) Then
        ' Pøidání tlaèítka pro pøidání dalších variant
        AddButtonTo ws, ws.Cells(2, 8), "Pøidat variantu", "AddMoreCandidates"
        
        ' Pøidání tlaèítka pro odebrání kritérií, pokud je poèet variant > 0
        If numOfCandidates > 0 Then
            AddButtonTo ws, ws.Cells(2, 10), "Odebrat variantu", "RemoveCandidate"
        End If
    End If
    
    ' Nastavení nové varianty na tuènì a podtržení
    With Range(ws.Cells(4, 5), ws.Cells(4, 5 + numOfCandidates - 1))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    ' Úprava šíøky novì pøidaného sloupce
    AdjustColumnWidth ws, 4 + numOfCandidates
    
    ThisWorkbook.Sheets("Vstupní data").Protect "1234"
    
    ' Aktivace TextBoxu pro další vstup
    TextBox1.SetFocus
    
End Sub

' Procedura obsluhující stisknutí tlaèítka pokraèovat
Private Sub Continue_Click()
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    
    ' Kontrola poètu variant, spodní hranice 2
    If ws.Range("F2").value < 2 Then
        MsgBox "Pøi rozhodování bychom mìli zohledòovat minimálnì 2 varianty.", vbExclamation
        Me.Hide
        AddCandidateForm.Show
    End If
    
    Call Update
    
    ' Zavøení UserFormu
    Unload Me
    
    ' Pøechod zpìt do Vstupní data pomocí boolean podmínky candidatesDone
    candidatesDone = True
    
    ThisWorkbook.Sheets("Vstupní data").Protect "1234"
    
End Sub
