VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoveCandidateForm 
   Caption         =   "Formuláø pro odebírání variant"
   ClientHeight    =   3432
   ClientLeft      =   84
   ClientTop       =   360
   ClientWidth     =   4128
   OleObjectBlob   =   "RemoveCandidateForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RemoveCandidateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    ' Nastavení focus na ListBox
    CandidateListBox.SetFocus
    
    ' Nastavení velikosti formuláøe
    With frm
        Height = 200
        Width = 269
    End With
    
End Sub

' Skript reagující na stisk tlaèítka
Private Sub RemoveButton_Click()
    Dim selectedCandidateIndex As Integer
    Dim selectedCandidate As String
    Dim lastColumn As Long
    Dim ws As Worksheet
    Dim numOfCandidates As Integer
    Dim numOfCriteria As Integer
    
    ' Nastavení pracovního listu, kde jsou varianty uloženy
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    
    ' Kontrola, zda je vybrána varianta
    If CandidateListBox.ListIndex = -1 Then
        MsgBox "Vyberte prosím variantu k odebrání.", vbExclamation
        Exit Sub
    End If
    
    ' Získání indexu vybraného kritéria v ListBoxu
    selectedCandidateIndex = CandidateListBox.ListIndex
    
    ' Získání názvu vybraného kritéria
    selectedCandidate = CandidateListBox.List(selectedCandidateIndex)
    
    ' Získat poèet kritérií z listu
    numOfCriteria = ws.Range("C2").value
    
    ' Získat poèet variant z listu
    numOfCandidates = ws.Range("F2").value
    
    ' Odebrání vybrané varianty z listu "Vstupní data"
    With ws
        .Unprotect "1234"
        ' Vymazání obsahu bunìk od øádku 4 po (4 + numOfCriteria) v sloupci (5 + selectedCandidateIndex)
        .Range(.Cells(4, 5 + selectedCandidateIndex), .Cells(4 + numOfCriteria, 5 + selectedCandidateIndex)).ClearContents
        
        ' Pøesunutí obsah bunìk vpravo od vybrané varianty o jeden sloupec doleva
        .Range(.Cells(4, 6 + selectedCandidateIndex), .Cells(4 + numOfCriteria, 6 + numOfCandidates)).Cut Destination:=.Cells(4, 5 + selectedCandidateIndex)
        
        ' Snížení hodnoty v buòce F2 o 1
        .Range("F2").value = numOfCandidates - 1
    End With
    
    ' Odebrání vybrané varianty z ListBoxu
    CandidateListBox.RemoveItem selectedCandidateIndex
    
    ' Kontrola, zda zùstal ještì nìjaký prvek v ListBoxu
    If CandidateListBox.ListCount = 0 Then
        MsgBox "Není žádná varianta k odebrání.", vbInformation
        Me.Hide
        HideButton ws, "Odebrat variantu"
        Exit Sub
    End If
    
    If CandidateListBox.ListCount < 2 Then
        HideButton ws, "Upravit hodnoty"
        HideButton ws, "Metoda WSA"
        HideButton ws, "Metoda bazické varianty"
    End If
    
    ' Zpráva potvrzující odebrání varianty
    MsgBox "Varianta '" & selectedCandidate & "' byla úspìšnì odebrána.", vbInformation
    
    ws.Columns(4 + numOfCandidates).ColumnWidth = 8.11
    ws.Protect "1234"
End Sub
