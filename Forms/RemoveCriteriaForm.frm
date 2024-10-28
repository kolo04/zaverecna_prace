VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoveCriteriaForm 
   Caption         =   "Formuláø pro odebírání kritérií"
   ClientHeight    =   3456
   ClientLeft      =   120
   ClientTop       =   396
   ClientWidth     =   5172
   OleObjectBlob   =   "RemoveCriteriaForm.frx":0000
End
Attribute VB_Name = "RemoveCriteriaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    
    ' Nastavení focus na ListBox
    CriteriaListBox.SetFocus
    
    ' Nastavení velikosti formuláøe
    With frm
        Height = 200
        Width = 269
    End With
    
End Sub

' Skript reagující na stisk tlaèítka
Private Sub RemoveButton_Click()
    Dim selectedCriteriaIndex As Integer
    Dim selectedCriteria As String
    Dim ws As Worksheet
    Dim numOfCriteria As Integer
    
    ' Nastavení pracovního listu, kde jsou kritéria uložena
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    
    ' Kontrola, zda je vybráno kritérium
    If CriteriaListBox.ListIndex = -1 Then
        MsgBox "Vyberte prosím kritérium k odebrání.", vbExclamation
        Exit Sub
    End If
    
    ' Získání indexu vybraného kritéria v ListBoxu
    selectedCriteriaIndex = CriteriaListBox.ListIndex
    
    ' Získání názvu vybraného kritéria
    selectedCriteria = CriteriaListBox.List(selectedCriteriaIndex)
    
    ' Odebrání vybraného kritéria z listu "Vstupní data"
    With ws
        .Unprotect "1234"
        
        ' Vymazání øádku s indexem kritéria
        .Rows(5 + selectedCriteriaIndex).Delete
        
        ' Odebrání vybraného kritéria z ListBoxu
        CriteriaListBox.RemoveItem selectedCriteriaIndex
        
        ' Snížení hodnoty v buòce C2 o 1
        .Range("C2").value = .Range("C2").value - 1
        
        HideButton ws, "Pokraèovat"
        HideButton ws, "Nahrát cíle"
        HideButton ws, "Metoda WSA"
        HideButton ws, "Metoda bazické varianty"
        
        ' Získání poètu kritérií
        numOfCriteria = ws.Range("C2").value
        
        ' Stanovit váhy lze pouze, když jsou pøítomna aspoò dvì kritéria
        If numOfCriteria > 1 Then
            HideButton ws, "Stanovit váhy"
            AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Stanovit váhy", "SetWeights"
        End If
    End With
    
    ' Pokud bude poèet položek v ListBoxu < 2, schování tlaèítka
    If CriteriaListBox.ListCount < 2 Then
        HideButton ws, "Stanovit váhy"
    End If
    
    ' Kontrola, zda zùstal ještì nìjaký prvek v ListBoxu
    If CriteriaListBox.ListCount = 0 Then
        MsgBox "Není žádné kritérium k odebrání.", vbInformation
        Me.Hide
        HideButton ws, "Odebrat kritérium"
        Exit Sub
    End If
    
    ' Zpráva potvrzující odebrání kritéria
    MsgBox "Kritérium '" & selectedCriteria & "' bylo úspìšnì odebráno.", vbInformation
    
    ws.Protect "1234"
End Sub
