VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddCandidateForm 
   Caption         =   "Formuláø pro pøidání variant"
   ClientHeight    =   2112
   ClientLeft      =   144
   ClientTop       =   408
   ClientWidth     =   4128
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
    
    ' Reset the size
    With frm
        ' Set the form size
        Height = 160
        Width = 269
    End With
    
End Sub

Private Sub AddButton_Click()

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
    
    ' Aktivace TextBoxu pro další vstup
    TextBox1.SetFocus
    
    ' Zrušení všech tlaèítek na listu
    ActiveSheet.Buttons.Delete
    
    ' Získání poètu kritérií z listu
    numOfCriteria = ws.Range("C2").value
    
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
            AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Pokraèovat", "FillData"
        Else
            ' Pøidání tlaèítka pro úpravu vyplnìných hodnot
            AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Upravit hodnoty", "EditCellValue"
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
    
    AdjustColumnWidth ws, 4 + numOfCandidates
    
    ThisWorkbook.Sheets("Vstupní data").Protect "1234"

End Sub

Private Sub Continue_Click()
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    ' Kontrola poètu variant, spodní hranice 2
    If ws.Range("F2").value < 2 Then
        MsgBox "Pøi rozhodování bychom mìli zohledòovat minimálnì 2 varianty.", vbExclamation
        Me.Hide
        AddCandidateForm.Show
    End If
    
    ' Zavøení UserFormu
    Unload Me
    ' Pøechod zpìt do Vstupní data pomocí boolean podmínky candidatesDone
    candidatesDone = True
    
    ThisWorkbook.Sheets("Vstupní data").Protect "1234"
    
End Sub

