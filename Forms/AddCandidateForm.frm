VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddCandidateForm 
   Caption         =   "Formul�� pro p�id�n� variant"
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
    ' P�i inicializaci formul��e bude TextBox1 aktivn� pro vstup u�ivatele
    TextBox1.SetFocus
    
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    
    ' Schov�n� tla��tka, pokud existuje
    HideButton ws, "P�idat variantu"
    
    ' P�id�n� tla��tka pro p�id�n� dal��ch variant
    AddButtonTo ws, ws.Cells(2, 8), "P�idat variantu", "AddMoreCandidates"
    
    ThisWorkbook.Sheets("Vstupn� data").Protect "1234"
    
    ' Reset the size
    With frm
        ' Set the form size
        Height = 160
        Width = 269
    End With
    
End Sub

' Procedura ovl�daj�c� tla��tko P�idat variantu, reaguje na stisknut� tla��tka
Private Sub AddButton_Click()

' P�id�n� nov� varianty na list "Vstupn� data"
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    
    ' Ur�en� sloupce pro novou variantu
    Dim lastColumn As Long
    lastColumn = ws.Cells(4, ws.Columns.Count).End(xlToLeft).column + 1
    
    Dim validInput As Boolean
    
    Do
        ' Zobrazen� chybov� zpr�vy, pokud je TextBox pr�zdn�,
        If TextBox1.Text = "" Then
            MsgBox "N�zev varianty nesm� b�t pr�zdn�.", vbExclamation
            
            ' Ukon�it proceduru, ale nechat formul�� otev�en�
            TextBox1.SetFocus
            ThisWorkbook.Sheets("Vstupn� data").Protect "1234"
            Exit Sub
        Else
            ' Znovuna�ten� aktu�ln�ho listu
            Set ws = ThisWorkbook.Sheets("Vstupn� data")
            
            ' Z�sk�n� aktu�ln�ho po�tu variant
            numOfCandidates = ws.Range("F2").value
        
            ' Kontrola, zda se varianta se stejn�m n�zvem ji� nevyskytuje
            If Not IsUniqueValue(ws.Range(ws.Cells(4, 5), ws.Cells(4, 4 + numOfCandidates - 1)), TextBox1.Text) Then
                MsgBox "Varianty mus� b�t unik�tn�!", vbExclamation
                
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
    
    ' Zaps�n� n�zvu varianty na list
    ws.Unprotect "1234"
    ws.Cells(4, lastColumn).value = "'" & TextBox1.Text
    
    ' Aktualizace po�tu variant na listu
    ws.Range("F2").value = numOfCandidates + 1
    numOfCandidates = numOfCandidates + 1
    
    ' Vypr�zdn�n� pole pro n�zev krit�ria
    TextBox1.Text = ""
    
    ' Aktivace TextBoxu pro dal�� vstup
    TextBox1.SetFocus
    
    ' Zru�en� v�ech tla��tek na listu
    ActiveSheet.Buttons.Delete
    
    ' Z�sk�n� po�tu krit�ri� z listu
    numOfCriteria = ws.Range("C2").value
    
    ' P�id�n� tla��tka pro p�id�n� dal��ch krit�ri�
    AddButtonTo ws, ws.Range("B" & 6 + numOfCriteria), "P�idat krit�rium", "AddMoreCriteria"
    
    'P�i jednom a v�ce krit�riu p�idat tla��tko pro odebr�n� krit�ria
    If numOfCriteria > 0 Then
        AddButtonTo ws, ws.Range("D" & 6 + numOfCriteria), "Odebrat krit�rium", "RemoveCriteria"
    End If
    
    ' Pokud nen� vypln�na v�ha u posledn�ho p�idan�ho krit�ria, pak p�idat tla��tko pro Stanoven� vah
    If IsEmpty(ws.Cells(4 + numOfCriteria, 4)) Then
        AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Stanovit v�hy", "MoveToM2"
    Else
        ' Kontrola, zda jsou bu�ky pr�zdn� pomoc� funkce CountA (v�po�et pr�zdn�ch bun�k)
        If WorksheetFunction.CountBlank(ws.Range(ws.Cells(5, 5), ws.Cells(4 + numOfCriteria, 4 + numOfCandidates))) > 0 Then
            ' P�id�n� tla��tka pro vypln�n� dat
            AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Pokra�ovat", "FillData"
        Else
            ' P�id�n� tla��tka pro �pravu vypln�n�ch hodnot
            AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Upravit hodnoty", "EditCellValue"
        End If
    End If
            
    ' Z�sk�n� aktu�ln�ho po�tu krit�ri�
    numOfCandidates = ws.Range("F2").value
    
    If Not IsEmpty(numOfCandidates) Then
        ' P�id�n� tla��tka pro p�id�n� dal��ch variant
        AddButtonTo ws, ws.Cells(2, 8), "P�idat variantu", "AddMoreCandidates"
        
        ' P�id�n� tla��tka pro odebr�n� krit�ri�, pokud je po�et variant > 0
        If numOfCandidates > 0 Then
            AddButtonTo ws, ws.Cells(2, 10), "Odebrat variantu", "RemoveCandidate"
        End If
    End If
    
    ' Nastaven� nov� varianty na tu�n� a podtr�en�
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
    
    ' �prava ���ky nov� p�idan�ho sloupce
    AdjustColumnWidth ws, 4 + numOfCandidates
    
    ThisWorkbook.Sheets("Vstupn� data").Protect "1234"

End Sub

' Procedura obsluhuj�c� stisknut� tla��tka pokra�ovat
Private Sub Continue_Click()
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    
    ' Kontrola po�tu variant, spodn� hranice 2
    If ws.Range("F2").value < 2 Then
        MsgBox "P�i rozhodov�n� bychom m�li zohled�ovat minim�ln� 2 varianty.", vbExclamation
        Me.Hide
        AddCandidateForm.Show
    End If
    
    ' Zav�en� UserFormu
    Unload Me
    
    ' P�echod zp�t do Vstupn� data pomoc� boolean podm�nky candidatesDone
    candidatesDone = True
    
    ThisWorkbook.Sheets("Vstupn� data").Protect "1234"
    
End Sub
