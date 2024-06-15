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
    
    ' Nastav focus na ListBox
    CriteriaListBox.SetFocus
    
    ' Reset the size
    With frm
        ' Set the form size
        Height = 200
        Width = 269
    End With
    
End Sub

' Skript reagující na stisk tlaèítka
Private Sub RemoveButton_Click()
    Dim selectedCriteriaIndex As Integer
    Dim selectedCriteria As String
    Dim ws As Worksheet
    
    ' Nastav pracovní list, kde jsou kritéria uložena
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    
    ' Zkontroluj, zda je vybráno kritérium
    If CriteriaListBox.ListIndex = -1 Then
        MsgBox "Vyberte prosím kritérium k odebrání.", vbExclamation
        Exit Sub
    End If
    
    ' Získání indexu vybraného kritéria v ListBoxu
    selectedCriteriaIndex = CriteriaListBox.ListIndex
    
    ' Získání názvu vybraného kritéria
    selectedCriteria = CriteriaListBox.List(selectedCriteriaIndex)
    
    ' Odebrání vybraného kritéria z listu "Vstupní data"
    ws.Unprotect "1234"
    ws.Rows(5 + selectedCriteriaIndex).Delete
    
    ' Odebrání vybraného kritéria z ListBoxu
    CriteriaListBox.RemoveItem selectedCriteriaIndex
    
    ' Snížení hodnoty v buòce C2 o 1
    ws.Range("C2").value = ws.Range("C2").value - 1
    ws.Protect "1234"
    
    'Pokud bude poèet položek v ListBoxu < 2, pak schovej tlaèítko
    If CriteriaListBox.ListCount < 2 Then
        HideButton ws, "Stanovit váhy"
    End If
    
    ' Zkontroluj, zda zùstal ještì nìjaký prvek v ListBoxu
    If CriteriaListBox.ListCount = 0 Then
        MsgBox "Není žádné kritérium k odebrání.", vbInformation
        Me.Hide
        HideButton ws, "Odebrat kritérium"
        Exit Sub
    End If
    
    ' Zpráva potvrzující odebrání kritéria
    MsgBox "Kritérium '" & selectedCriteria & "' bylo úspìšnì odebráno.", vbInformation
    
End Sub

