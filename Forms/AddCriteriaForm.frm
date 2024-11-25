VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddCriteriaForm 
   Caption         =   "Formuláø pro pøidání kritérií"
   ClientHeight    =   3048
   ClientLeft      =   96
   ClientTop       =   228
   ClientWidth     =   5160
   OleObjectBlob   =   "AddCriteriaForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddCriteriaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Mimo proceduru => deklarace promìných globálnì pro celý modul
Dim ws As Worksheet

' Promìnná poètu kritérií
Dim numOfCriteria As Integer

' Promìnná poètu variant
Dim numOfCandidates As Integer

' Pøíprava True/False promìnné pro obsluhu otevírání UserFormu - volána v metodì InputData
Dim criteriaDone As Boolean

Private Sub UserForm_Initialize()
    
    ' Pøi inicializaci formuláøe bude TextBox1 aktivní pro vstup uživatele
    TextBox1.SetFocus
    
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    
    ' Získání aktuálního poètu kritérií (pokud už nìjaká jsou)
    numOfCriteria = ws.Range("C2").value
    
    ' Schování tlaèítka, pokud existuje
    HideButton ws, "Pøidat kritérium"
    
    ' Pøidání tlaèítka pro pøidání dalších kritérií
    AddButtonTo ws, ws.Range("B" & 6 + numOfCriteria), "Pøidat kritérium", "AddMoreCriteria"
    
    ThisWorkbook.Sheets("Vstupní data").Protect "1234"
    
    ' Nastavení velikosti (pùvodnì 160x269)
    With frm
        Height = 185
        Width = 269
    End With
    
End Sub

' Procedura ovládající tlaèítko Pøidat kritérium, reaguje na stisknutí tlaèítka
Private Sub Add_Click()

' Pøidání nového kritéria na list Vstupní data
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    
    ' Urèení øádku pro nové kritérium
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row + 1
    
    Dim validInput As Boolean
    
    ' Cyklus, který bude kontrolovat vstup proti všem podmínkám, dokud nebude validní
    Do
        ' Pokud TextBox je prázdný, zobrazí se chybová zpráva
        If TextBox1.Text = "" Then
            MsgBox "Název kritéria nesmí být prázdný.", vbExclamation
            
            ' Ukonèit proceduru, ale nechat formuláø otevøený
            TextBox1.SetFocus
            ThisWorkbook.Sheets("Vstupní data").Protect "1234"
            Exit Sub
            
        Else
            ' Znovunaètení aktuálního listu
            Set ws = ThisWorkbook.Sheets("Vstupní data")
            
            ' Získání aktuálního poètu kritérií
            numOfCriteria = ws.Range("C2").value
        
            ' Kontrola, zda se kritérium se stejným názvem již nevyskytuje
            If Not IsUniqueValue(ws.Range(ws.Cells(5, 2), ws.Cells(4 + numOfCriteria, 2)), TextBox1.Text) Then
                MsgBox "Kritéria musí být unikátní!", vbExclamation
                
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
        
    ' Zapsání názvu kritéria na list jako text
    ws.Unprotect "1234"
    ws.Cells(lastRow, 2).value = "'" & TextBox1.Text
    
    ' Aktualizace poètu kritérií z listu
    ws.Range("C2").value = numOfCriteria + 1
    numOfCriteria = numOfCriteria + 1
    
    ' Vyprázdnìní pole pro název kritéria
    TextBox1.Text = ""
    
    Call Update
End Sub

' Skript umožòující nahrát kritéria z vybrané oblasti
Private Sub Upload_Click()
    Dim rng As Range
    Dim subject As String
    Dim criteriaRange As Range
    Dim numOfRows As Integer
    Dim duplicateFound As Boolean
    Dim cell As Range
    Dim UniqueValues As Object

    ' Odkaz na list "Vstupní data"
    Set ws = ThisWorkbook.Sheets("Vstupní data")

    ' Získání aktuálního poètu kritérií z buòky C2
    Set criteriaRange = ws.Range("C2")
    numOfCriteria = criteriaRange.value

    ' Nastavení cílové buòky (zaèátek v B5 + poèet kritérií)
    Set rng = ws.Cells(5 + numOfCriteria, 2)

    ' Pøedmìt pro zobrazení v InputBoxu
    subject = "kritéria"
    
    ' Volání samostatné procedury pro nahrávání dat a získání poètu vložených øádkù
    numOfRows = UploadData(rng, subject)

    ' Ovìøení, zda došlo k úspìšnému nahrání dat
    If numOfRows > 0 Then
        ' Slovník pro kontrolu unikátních hodnot
        Set UniqueValues = CreateObject("Scripting.Dictionary")
        
        ' Pomocná promìnná pro kontrolu duplicit
        duplicateFound = False
        
        ' Kontrola unikátnosti novì nahraných kritérií
        For Each cell In ws.Range(ws.Cells(5 + numOfCriteria, 2), ws.Cells(4 + numOfCriteria + numOfRows, 2))
            If cell.value <> "" Then
                
                ' Kontrola existujících kritérií (pokud nìjaká existují)
                If numOfCriteria > 0 Then
                    If Not IsUniqueValue(ws.Range("B5:B" & 4 + numOfCriteria), cell.value) Then
                        duplicateFound = True
                        Exit For
                    End If
                End If
                
                ' Kontrola duplicit v aktuálním slovníku
                If UniqueValues.Exists(cell.value) Then
                    duplicateFound = True
                    Exit For
                End If
                
                ' Pøidání hodnoty do slovníku
                UniqueValues.Add cell.value, True
            End If
        Next cell

        ' Zpracování výsledkù kontroly duplicit
        If duplicateFound Then
            MsgBox "Vkládaná kritéria musí být unikátní! Nahrávání bylo zrušeno.", vbExclamation
            ws.Unprotect "1234"
            ws.Range(ws.Cells(5 + numOfCriteria, 2), ws.Cells(4 + numOfCriteria + numOfRows, 2)).Clear
        Else
            ' Aktualizace poètu kritérií o poèet vložených øádkù
            ws.Unprotect "1234"
            criteriaRange.value = numOfCriteria + numOfRows
        End If
    End If

    ' Uzamknutí listu na konci procedury
    ThisWorkbook.Sheets("Vstupní data").Protect "1234"

    Call Update
    
    Unload Me
End Sub

' Spoleèný skript pro oba zpùsoby vkládání kritérií, obslouží potøebné úpravy formuláøe i listu
Private Sub Update()
    
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    ws.Activate
    
    ws.Unprotect "1234"
    
    ' Aktivace TextBoxu pro další vstup
    TextBox1.SetFocus
    
    ' Zrušení všech tlaèítek na listu
    ws.Buttons.Delete
    
    ' Získání aktuálního poètu kritérií
    numOfCriteria = ws.Range("C2").value
        
    ' Pøidání tlaèítka pro pøidání dalších kritérií
    AddButtonTo ws, ws.Range("B" & 6 + numOfCriteria), "Pøidat kritérium", "AddMoreCriteria"
    
    'Pøi jednom a více kritériu pøidat tlaèítko pro odebrání kritéria
    If numOfCriteria > 0 Then
        AddButtonTo ws, ws.Range("D" & 6 + numOfCriteria), "Odebrat kritérium", "RemoveCriteria"
    End If
    
    ' Stanovit váhy lze pouze, když jsou pøítomna aspoò dvì kritéria
    If numOfCriteria > 1 Then
        AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Stanovit váhy", "SetWeights"
    End If
    
    ' Získání aktuálního poètu variant
    numOfCandidates = ws.Range("F2").value
    
    If Not IsEmpty(ws.Range("E2")) Then
        ' Pøidání tlaèítka pro pøidání dalších variant
        AddButtonTo ws, ws.Cells(2, 8), "Pøidat variantu", "AddMoreCandidates"
        
        ' Pøidání tlaèítka pro odebrání kritérií, pokud je poèet variant > 0
        If numOfCandidates > 0 Then
            AddButtonTo ws, ws.Cells(2, 10), "Odebrat variantu", "RemoveCandidate"
        End If
    End If
    
    ' Úprava šíøky sloupce kritérií
    AdjustColumnWidth ws, 2
    
    ' Pøidání tlaèítka pro nový pøíklad
    Call AddRestartButton
    
    ThisWorkbook.Sheets("Vstupní data").Protect "1234"
    
    ' Aktivace TextBoxu pro další vstup
    TextBox1.SetFocus

End Sub

' Procedura obsluhující stisknutí tlaèítka pokraèovat
Private Sub Continue_Click()
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    
    ' Kontrola poètu kritérií, spodní hranice 2
    If ws.Range("C2").value < 2 Then
        MsgBox "Pøi rozhodování bychom mìli zohledòovat minimálnì 2 kritéria.", vbExclamation
        Me.Hide
        AddCriteriaForm.Show
    End If
    
    ' Zavøení UserFormu
    Unload Me
    
    ' Pøechod zpìt do Vstupní data pomocí boolean podmínky criteriaDone
    criteriaDone = True
    
    ThisWorkbook.Sheets("Vstupní data").Protect "1234"
    
End Sub
