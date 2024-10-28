VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ObjectivesForm 
   Caption         =   "Formul�� pro stanoven� c�l�"
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
    
' Nastaven� velikosti
    With frm
        Height = 135
        Width = 269
    End With
    
' Nastaven� validace pro bu�ky
    Set ws = ThisWorkbook.Sheets("Vstupn� data")

    ' Metoda pro stanoven� c�l� manu�ln�
    With ws
        .Activate
        .Unprotect "1234"
        
        ' Z�sk�n� aktu�ln�ho po�tu krit�ri�
        numOfCriteria = .Range("C2").value
        
        ' Nastaven� textu "C�l" do bu�ky C4
        .Range("C4").value = "C�l"
        
        ' Vytvo�en� rozev�rac�ho seznamu s mo�nostmi "min" a "max" pro ka�dou bu�ku v rozsahu C4 a� C(4 + numOfCriteria)
        Dim criteriaRange As Range
        Dim options As Variant
        options = Array("min", "max")
        
        Set criteriaRange = .Range(.Cells(5, 3), .Cells(5 + numOfCriteria - 1, 3))
        For Each cell In criteriaRange
            With cell.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(options, ",")
            End With
            ' Nastaven� popisku "Vyberte" pro ka�dou bu�ku
            cell.value = "Vyberte"
            cell.Locked = False
        Next cell
        
        ' Form�tov�n� z�hlav� B4:D4
        With .Range("B4:D4")
            ' Tu�n� a zarovn�n� na st�ed
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        
            ' Nastaven� ohrani�en�
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
    
        ' Zarovn�n� bun�k C4:C (4 + numOfCriteria) na st�ed
        .Range("C4:C" & 4 + numOfCriteria).HorizontalAlignment = xlCenter
        
        ' Nastaven� stylu bun�k D4:D (4 + numOfCriteria) jako "Percent" s form�tem "0.0 %"
        .Range("D4:D" & 4 + numOfCriteria).NumberFormat = "0.0 %"
        
        ' �prava ���ky sloupc�
        AdjustColumnWidth ws, .Range(.Columns(2), .Columns(3))
        
        .Cells(5, 3).Select
        
        .Protect "1234"
    End With
    
End Sub

Private Sub Manual_Click()

    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    
    ' Z�sk�n� aktu�ln�ho po�tu krit�ri�
    numOfCriteria = ws.Range("C2").value
    
    ws.Unprotect "1234"

    HideButton ws, "Stanovit c�le"
    
    ' P�id�n� tla��tka pro nahr�n� c�l�
    AddButtonTo ws, ws.Range("F" & 6 + numOfCriteria), "Nahr�t c�le", "UploadObjectives"
    
    ' P�id�n� tla��tka pro zad�n� variant
    AddButtonTo ws, ws.Range("F" & 9 + numOfCriteria), "Pokra�ovat", "Candidates"
    
    ws.Protect "1234"

    ' Zav�en� formul��e po dokon�en�
    Unload Me
End Sub

Private Sub Upload_Click()
    ' Vol�n� metody pro nahr�n� c�l�
    Call UploadObjectives
    
    ' Zav�en� formul��e po dokon�en�
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Kontrola, zda byl formul�� zav�en pomoc� tla��tka "X" (CloseMode = 0)
    If CloseMode = vbFormControlMenu Then
        MsgBox "V�b�r metody zad�v�n� byl zru�en.", vbExclamation
        Unload Me
    End If
End Sub
