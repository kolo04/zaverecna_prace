VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddCriteriaForm 
   Caption         =   "Formul�� pro p�id�n� krit�ri�"
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
    
    ' Nastaven� velikosti (p�vodn� 160x269)
    With frm
        Height = 185
        Width = 269
    End With
    
End Sub

' Procedura ovl�daj�c� tla��tko P�idat krit�rium, reaguje na stisknut� tla��tka
Private Sub Add_Click()

' P�id�n� nov�ho krit�ria na list Vstupn� data
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    
    ' Ur�en� ��dku pro nov� krit�rium
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row + 1
    
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
    
    Call Update
End Sub

' Skript umo��uj�c� nahr�t krit�ria z vybran� oblasti
Private Sub Upload_Click()
    Dim rng As Range
    Dim subject As String
    Dim criteriaRange As Range
    Dim numOfRows As Integer
    Dim duplicateFound As Boolean
    Dim cell As Range
    Dim UniqueValues As Object

    ' Odkaz na list "Vstupn� data"
    Set ws = ThisWorkbook.Sheets("Vstupn� data")

    ' Z�sk�n� aktu�ln�ho po�tu krit�ri� z bu�ky C2
    Set criteriaRange = ws.Range("C2")
    numOfCriteria = criteriaRange.value

    ' Nastaven� c�lov� bu�ky (za��tek v B5 + po�et krit�ri�)
    Set rng = ws.Cells(5 + numOfCriteria, 2)

    ' P�edm�t pro zobrazen� v InputBoxu
    subject = "krit�ria"
    
    ' Vol�n� samostatn� procedury pro nahr�v�n� dat a z�sk�n� po�tu vlo�en�ch ��dk�
    numOfRows = UploadData(rng, subject)

    ' Ov��en�, zda do�lo k �sp�n�mu nahr�n� dat
    If numOfRows > 0 Then
        ' Slovn�k pro kontrolu unik�tn�ch hodnot
        Set UniqueValues = CreateObject("Scripting.Dictionary")
        
        ' Pomocn� prom�nn� pro kontrolu duplicit
        duplicateFound = False
        
        ' Kontrola unik�tnosti nov� nahran�ch krit�ri�
        For Each cell In ws.Range(ws.Cells(5 + numOfCriteria, 2), ws.Cells(4 + numOfCriteria + numOfRows, 2))
            If cell.value <> "" Then
                
                ' Kontrola existuj�c�ch krit�ri� (pokud n�jak� existuj�)
                If numOfCriteria > 0 Then
                    If Not IsUniqueValue(ws.Range("B5:B" & 4 + numOfCriteria), cell.value) Then
                        duplicateFound = True
                        Exit For
                    End If
                End If
                
                ' Kontrola duplicit v aktu�ln�m slovn�ku
                If UniqueValues.Exists(cell.value) Then
                    duplicateFound = True
                    Exit For
                End If
                
                ' P�id�n� hodnoty do slovn�ku
                UniqueValues.Add cell.value, True
            End If
        Next cell

        ' Zpracov�n� v�sledk� kontroly duplicit
        If duplicateFound Then
            MsgBox "Vkl�dan� krit�ria mus� b�t unik�tn�! Nahr�v�n� bylo zru�eno.", vbExclamation
            ws.Unprotect "1234"
            ws.Range(ws.Cells(5 + numOfCriteria, 2), ws.Cells(4 + numOfCriteria + numOfRows, 2)).Clear
        Else
            ' Aktualizace po�tu krit�ri� o po�et vlo�en�ch ��dk�
            ws.Unprotect "1234"
            criteriaRange.value = numOfCriteria + numOfRows
        End If
    End If

    ' Uzamknut� listu na konci procedury
    ThisWorkbook.Sheets("Vstupn� data").Protect "1234"

    Call Update
    
    Unload Me
End Sub

' Spole�n� skript pro oba zp�soby vkl�d�n� krit�ri�, obslou�� pot�ebn� �pravy formul��e i listu
Private Sub Update()
    
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    ws.Activate
    
    ws.Unprotect "1234"
    
    ' Aktivace TextBoxu pro dal�� vstup
    TextBox1.SetFocus
    
    ' Zru�en� v�ech tla��tek na listu
    ws.Buttons.Delete
    
    ' Z�sk�n� aktu�ln�ho po�tu krit�ri�
    numOfCriteria = ws.Range("C2").value
        
    ' P�id�n� tla��tka pro p�id�n� dal��ch krit�ri�
    AddButtonTo ws, ws.Range("B" & 6 + numOfCriteria), "P�idat krit�rium", "AddMoreCriteria"
    
    'P�i jednom a v�ce krit�riu p�idat tla��tko pro odebr�n� krit�ria
    If numOfCriteria > 0 Then
        AddButtonTo ws, ws.Range("D" & 6 + numOfCriteria), "Odebrat krit�rium", "RemoveCriteria"
    End If
    
    ' Stanovit v�hy lze pouze, kdy� jsou p��tomna aspo� dv� krit�ria
    If numOfCriteria > 1 Then
        AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Stanovit v�hy", "SetWeights"
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
    
    ' P�id�n� tla��tka pro nov� p��klad
    Call AddRestartButton
    
    ThisWorkbook.Sheets("Vstupn� data").Protect "1234"
    
    ' Aktivace TextBoxu pro dal�� vstup
    TextBox1.SetFocus

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
