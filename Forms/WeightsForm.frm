VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WeightsForm 
   Caption         =   "Formuláø pro stanovení vah"
   ClientHeight    =   1608
   ClientLeft      =   96
   ClientTop       =   396
   ClientWidth     =   4128
   OleObjectBlob   =   "WeightsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WeightsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ws As Worksheet

Private Sub UserForm_Initialize()
    
    ' Nastavení velikosti
    With frm
        Height = 135
        Width = 269
    End With
    
    Set ws = ThisWorkbook.Sheets("Vstupní data")
    With ws
        .Unprotect "1234"
        .Range("D4").value = "Váha"
        
        ' Tuèné písmo pro záhlaví
        .Range("B4:D4").Font.Bold = True
        
        .Protect "1234"
    End With
    
End Sub

Private Sub Manual_Click()
    ' Volání metody pro stanovení vah výpoètem
    Call MoveToM2
    Unload Me ' Zavøe formuláø po dokonèení
End Sub

Private Sub Upload_Click()
    ' Volání metody pro nahrání vah
    Call UploadWeights
    Unload Me ' Zavøe formuláø po dokonèení
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Kontrola, zda byl formuláø zavøen pomocí tlaèítka "X" (CloseMode = 0)
    If CloseMode = vbFormControlMenu Then
        MsgBox "Výbìr metody zadávání byl zrušen.", vbExclamation
        Unload Me
    End If
End Sub
