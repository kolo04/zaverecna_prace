VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoveCandidateForm 
   Caption         =   "Formul�� pro odeb�r�n� variant"
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

    ' Nastaven� focus na ListBox
    CandidateListBox.SetFocus
    
    ' Nastaven� velikosti formul��e
    With frm
        Height = 200
        Width = 269
    End With
    
End Sub

' Skript reaguj�c� na stisk tla��tka
Private Sub RemoveButton_Click()
    Dim selectedCandidateIndex As Integer
    Dim selectedCandidate As String
    Dim lastColumn As Long
    Dim ws As Worksheet
    Dim numOfCandidates As Integer
    Dim numOfCriteria As Integer
    
    ' Nastaven� pracovn�ho listu, kde jsou varianty ulo�eny
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    
    ' Kontrola, zda je vybr�na varianta
    If CandidateListBox.ListIndex = -1 Then
        MsgBox "Vyberte pros�m variantu k odebr�n�.", vbExclamation
        Exit Sub
    End If
    
    ' Z�sk�n� indexu vybran�ho krit�ria v ListBoxu
    selectedCandidateIndex = CandidateListBox.ListIndex
    
    ' Z�sk�n� n�zvu vybran�ho krit�ria
    selectedCandidate = CandidateListBox.List(selectedCandidateIndex)
    
    ' Z�skat po�et krit�ri� z listu
    numOfCriteria = ws.Range("C2").value
    
    ' Z�skat po�et variant z listu
    numOfCandidates = ws.Range("F2").value
    
    ' Odebr�n� vybran� varianty z listu "Vstupn� data"
    With ws
        .Unprotect "1234"
        ' Vymaz�n� obsahu bun�k od ��dku 4 po (4 + numOfCriteria) v sloupci (5 + selectedCandidateIndex)
        .Range(.Cells(4, 5 + selectedCandidateIndex), .Cells(4 + numOfCriteria, 5 + selectedCandidateIndex)).ClearContents
        
        ' P�esunut� obsah bun�k vpravo od vybran� varianty o jeden sloupec doleva
        .Range(.Cells(4, 6 + selectedCandidateIndex), .Cells(4 + numOfCriteria, 6 + numOfCandidates)).Cut Destination:=.Cells(4, 5 + selectedCandidateIndex)
        
        ' Sn�en� hodnoty v bu�ce F2 o 1
        .Range("F2").value = numOfCandidates - 1
    End With
    
    ' Odebr�n� vybran� varianty z ListBoxu
    CandidateListBox.RemoveItem selectedCandidateIndex
    
    ' Kontrola, zda z�stal je�t� n�jak� prvek v ListBoxu
    If CandidateListBox.ListCount = 0 Then
        MsgBox "Nen� ��dn� varianta k odebr�n�.", vbInformation
        Me.Hide
        HideButton ws, "Odebrat variantu"
        Exit Sub
    End If
    
    If CandidateListBox.ListCount < 2 Then
        HideButton ws, "Upravit hodnoty"
        HideButton ws, "Metoda WSA"
        HideButton ws, "Metoda bazick� varianty"
    End If
    
    ' Zpr�va potvrzuj�c� odebr�n� varianty
    MsgBox "Varianta '" & selectedCandidate & "' byla �sp�n� odebr�na.", vbInformation
    
    ws.Columns(4 + numOfCandidates).ColumnWidth = 8.11
    ws.Protect "1234"
End Sub
