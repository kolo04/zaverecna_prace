VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddCriteriaForm 
   Caption         =   "Formul�� pro p�id�n� krit�ri�"
   ClientHeight    =   2640
   ClientLeft      =   96
   ClientTop       =   228
   ClientWidth     =   4128
   OleObjectBlob   =   "AddCriteriaForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddCriteriaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Mimo proceduru => deklarace prom�n�ch glob�ln� pro cel� modul
Dim ws As Worksheet

' Prom�nn� po�tu krit�ri�
Dim numOfCriteria As Integer

' Prom�nn� po�tu variant
Dim numOfCandidates As Integer

' P��prava True/False prom�nn� pro obsluhu otev�r�n� UserFormu - vol�na v metod� InputData
Dim criteriaDone As Boolean

Private Sub UserForm_Initialize()
    
    ' P�i inicializaci formul��e bude TextBox1 aktivn� pro vstup u�ivatele
    TextBox1.SetFocus
    
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    
    ' Z�sk�n� aktu�ln�ho po�tu krit�ri� (pokud u� n�jak� jsou)
    numOfCriteria = ws.Range("C2").value
    
    ' Schov�n� tla��tka, pokud existuje
    HideButton ws, "P�idat krit�rium"
    
    ' P�id�n� tla��tka pro p�id�n� dal��ch krit�ri�
    AddButtonTo ws, ws.Range("B" & 6 + numOfCriteria), "P�idat krit�rium", "AddMoreCriteria"
    
    ThisWorkbook.Sheets("Vstupn� data").Protect "1234"
    
    ' Reset the size
    With frm
        ' Set the form size
        Height = 160
        Width = 269
    End With
    
End Sub

' Procedura ovl�daj�c� tla��tko P�idat krit�rium, reaguje na stisknut� tla��tka
Private Sub AddButton_Click()

' P�id�n� nov�ho krit�ria na list Vstupn� data
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    
    ' Ur�en� ��dku pro nov� krit�rium
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row + 1
    
    Dim validInput As Boolean
    
    ' Cyklus, kter� bude kontrolovat vstup proti v�em podm�nk�m, dokud nebude validn�
    Do
        ' Pokud TextBox je pr�zdn�, zobraz� se chybov� zpr�va
        If TextBox1.Text = "" Then
            MsgBox "N�zev krit�ria nesm� b�t pr�zdn�.", vbExclamation
            
            ' Ukon�it proceduru, ale nechat formul�� otev�en�
            TextBox1.SetFocus
            ThisWorkbook.Sheets("Vstupn� data").Protect "1234"
            Exit Sub
            
        Else
            ' Znovuna�ten� aktu�ln�ho listu
            Set ws = ThisWorkbook.Sheets("Vstupn� data")
            
            ' Z�sk�n� aktu�ln�ho po�tu krit�ri�
            numOfCriteria = ws.Range("C2").value
        
            ' Kontrola, zda se krit�rium se stejn�m n�zvem ji� nevyskytuje
            If Not IsUniqueValue(ws.Range(ws.Cells(5, 2), ws.Cells(4 + numOfCriteria, 2)), TextBox1.Text) Then
                MsgBox "Krit�ria mus� b�t unik�tn�!", vbExclamation
                
                ' Ukon�it proceduru, ale nechat formul�� otev�en�
                TextBox1.SetFocus
                ThisWorkbook.Sheets("Vstupn� data").Protect "1234"
                Exit Sub
            Else
                ' Platn� vstup
                validInput = True
            End If
        End If
    Loop Until validInput
        
    ' Zaps�n� n�zvu krit�ria na list jako text
    ws.Unprotect "1234"
    ws.Cells(lastRow, 2).value = "'" & TextBox1.Text
    
    ' Aktualizace po�tu krit�ri� z listu
    ws.Range("C2").value = numOfCriteria + 1
    numOfCriteria = numOfCriteria + 1
    
    ' Vypr�zdn�n� pole pro n�zev krit�ria
    TextBox1.Text = ""
    
    ' Aktivace TextBoxu pro dal�� vstup
    TextBox1.SetFocus
    
    ws.Buttons.Delete
        
    ' P�id�n� tla��tka pro p�id�n� dal��ch krit�ri�
    AddButtonTo ws, ws.Range("B" & 6 + numOfCriteria), "P�idat krit�rium", "AddMoreCriteria"
    
    'P�i jednom a v�ce krit�riu p�idat tla��tko pro odebr�n� krit�ria
    If numOfCriteria > 0 Then
        AddButtonTo ws, ws.Range("D" & 6 + numOfCriteria), "Odebrat krit�rium", "RemoveCriteria"
    End If
    
    ' Stanovit v�hy lze pouze, kdy� jsou p��tomna aspo� dv� krit�ria
    If numOfCriteria > 1 Then
        AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Stanovit v�hy", "MoveToM2"
    End If
    
    ' Z�sk�n� aktu�ln�ho po�tu variant
    numOfCandidates = ws.Range("F2").value
    
    If Not IsEmpty(ws.Range("E2")) Then
        ' P�id�n� tla��tka pro p�id�n� dal��ch variant
        AddButtonTo ws, ws.Cells(2, 8), "P�idat variantu", "AddMoreCandidates"
        
        ' P�id�n� tla��tka pro odebr�n� krit�ri�, pokud je po�et variant > 0
        If numOfCandidates > 0 Then
            AddButtonTo ws, ws.Cells(2, 10), "Odebrat variantu", "RemoveCandidate"
        End If
    End If
    
    ' �prava ���ky sloupce krit�ri�
    AdjustColumnWidth ws, 2
    
    ThisWorkbook.Sheets("Vstupn� data").Protect "1234"

End Sub

' Procedura obsluhuj�c� stisknut� tla��tka pokra�ovat
Private Sub Continue_Click()
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    
    ' Kontrola po�tu krit�ri�, spodn� hranice 2
    If ws.Range("C2").value < 2 Then
        MsgBox "P�i rozhodov�n� bychom m�li zohled�ovat minim�ln� 2 krit�ria.", vbExclamation
        Me.Hide
        AddCriteriaForm.Show
    End If
    
    ' Zav�en� UserFormu
    Unload Me
    
    ' P�echod zp�t do Vstupn� data pomoc� boolean podm�nky criteriaDone
    criteriaDone = True
    
    ThisWorkbook.Sheets("Vstupn� data").Protect "1234"
    
End Sub
