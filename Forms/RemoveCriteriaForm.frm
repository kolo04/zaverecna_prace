VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoveCriteriaForm 
   Caption         =   "Formul�� pro odeb�r�n� krit�ri�"
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
    
    ' Nastaven� focus na ListBox
    CriteriaListBox.SetFocus
    
    ' Nastaven� velikosti formul��e
    With frm
        Height = 200
        Width = 269
    End With
    
End Sub

' Skript reaguj�c� na stisk tla��tka
Private Sub RemoveButton_Click()
    Dim selectedCriteriaIndex As Integer
    Dim selectedCriteria As String
    Dim ws As Worksheet
    Dim numOfCriteria As Integer
    
    ' Nastaven� pracovn�ho listu, kde jsou krit�ria ulo�ena
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    
    ' Kontrola, zda je vybr�no krit�rium
    If CriteriaListBox.ListIndex = -1 Then
        MsgBox "Vyberte pros�m krit�rium k odebr�n�.", vbExclamation
        Exit Sub
    End If
    
    ' Z�sk�n� indexu vybran�ho krit�ria v ListBoxu
    selectedCriteriaIndex = CriteriaListBox.ListIndex
    
    ' Z�sk�n� n�zvu vybran�ho krit�ria
    selectedCriteria = CriteriaListBox.List(selectedCriteriaIndex)
    
    ' Odebr�n� vybran�ho krit�ria z listu "Vstupn� data"
    With ws
        .Unprotect "1234"
        
        ' Vymaz�n� ��dku s indexem krit�ria
        .Rows(5 + selectedCriteriaIndex).Delete
        
        ' Odebr�n� vybran�ho krit�ria z ListBoxu
        CriteriaListBox.RemoveItem selectedCriteriaIndex
        
        ' Sn�en� hodnoty v bu�ce C2 o 1
        .Range("C2").value = .Range("C2").value - 1
        
        HideButton ws, "Pokra�ovat"
        HideButton ws, "Nahr�t c�le"
        HideButton ws, "Metoda WSA"
        HideButton ws, "Metoda bazick� varianty"
        
        ' Z�sk�n� po�tu krit�ri�
        numOfCriteria = ws.Range("C2").value
        
        ' Stanovit v�hy lze pouze, kdy� jsou p��tomna aspo� dv� krit�ria
        If numOfCriteria > 1 Then
            HideButton ws, "Stanovit v�hy"
            AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Stanovit v�hy", "SetWeights"
        End If
    End With
    
    ' Pokud bude po�et polo�ek v ListBoxu < 2, schov�n� tla��tka
    If CriteriaListBox.ListCount < 2 Then
        HideButton ws, "Stanovit v�hy"
    End If
    
    ' Kontrola, zda z�stal je�t� n�jak� prvek v ListBoxu
    If CriteriaListBox.ListCount = 0 Then
        MsgBox "Nen� ��dn� krit�rium k odebr�n�.", vbInformation
        Me.Hide
        HideButton ws, "Odebrat krit�rium"
        Exit Sub
    End If
    
    ' Zpr�va potvrzuj�c� odebr�n� krit�ria
    MsgBox "Krit�rium '" & selectedCriteria & "' bylo �sp�n� odebr�no.", vbInformation
    
    ws.Protect "1234"
End Sub
