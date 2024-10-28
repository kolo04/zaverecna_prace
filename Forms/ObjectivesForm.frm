VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ObjectivesForm 
   Caption         =   "Formuláø pro stanovení cílù"
   ClientHeight    =   1704
   ClientLeft      =   96
   ClientTop       =   372
   ClientWidth     =   4128
   OleObjectBlob   =   "ObjectivesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ObjectivesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numOfCriteria As Integer
Dim ws As Worksheet

Private Sub UserForm_Initialize()
    
' Nastavení velikosti
    With frm
        Height = 135
        Width = 269
    End With
    
' Nastavení validace pro buòky
    Set ws = ThisWorkbook.Sheets("Vstupní data")

    ' Metoda pro stanovení cílù manuálnì
    With ws
        .Activate
        .Unprotect "1234"
        
        ' Získání aktuálního poètu kritérií
        numOfCriteria = .Range("C2").value
        
        ' Nastavení textu "Cíl" do buòky C4
        .Range("C4").value = "Cíl"
        
        ' Vytvoøení rozevíracího seznamu s možnostmi "min" a "max" pro každou buòku v rozsahu C4 až C(4 + numOfCriteria)
        Dim criteriaRange As Range
        Dim options As Variant
        options = Array("min", "max")
        
        Set criteriaRange = .Range(.Cells(5, 3), .Cells(5 + numOfCriteria - 1, 3))
        For Each cell In criteriaRange
            With cell.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(options, ",")
            End With
            ' Nastavení popisku "Vyberte" pro každou buòku
            cell.value = "Vyberte"
            cell.Locked = False
        Next cell
        
        ' Formátování záhlaví B4:D4
        With .Range("B4:D4")
            ' Tuènì a zarovnání na støed
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        
            ' Nastavení ohranièení
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
    
        ' Zarovnání bunìk C4:C (4 + numOfCriteria) na støed
        .Range("C4:C" & 4 + numOfCriteria).HorizontalAlignment = xlCenter
        
        ' Nastavení stylu bunìk D4:D (4 + numOfCriteria) jako "Percent" s formátem "0.0 %"
        .Range("D4:D" & 4 + numOfCriteria).NumberFormat = "0.0 %"
        
        ' Úprava šíøky sloupcù
        AdjustColumnWidth ws, .Range(.Columns(2), .Columns(3))
        
        .Cells(5, 3).Select
        
        .Protect "1234"
    End With
    
End Sub

Private Sub Manual_Click()

    Set ws = ThisWorkbook.Sheets("Vstupní data")
    
    ' Získání aktuálního poètu kritérií
    numOfCriteria = ws.Range("C2").value
    
    ws.Unprotect "1234"

    HideButton ws, "Stanovit cíle"
    
    ' Pøidání tlaèítka pro nahrání cílù
    AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Nahrát cíle", "UploadObjectives"
    
    ' Pøidání tlaèítka pro zadání variant
    AddButtonTo ws, ws.Range("F" & 9 + numOfCriteria), "Pokraèovat", "Candidates"
    
    ws.Protect "1234"

    ' Zavøení formuláøe po dokonèení
    Unload Me
End Sub

Private Sub Upload_Click()
    ' Volání metody pro nahrání cílù
    Call UploadObjectives
    
    ' Zavøení formuláøe po dokonèení
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Kontrola, zda byl formuláø zavøen pomocí tlaèítka "X" (CloseMode = 0)
    If CloseMode = vbFormControlMenu Then
        MsgBox "Výbìr metody zadávání byl zrušen.", vbExclamation
        Unload Me
    End If
End Sub
