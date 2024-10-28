VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EntryForm 
   Caption         =   "Formuláø pro vkládání dat"
   ClientHeight    =   2112
   ClientLeft      =   168
   ClientTop       =   696
   ClientWidth     =   5292
   OleObjectBlob   =   "EntryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EntryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    
    Dim ws As Worksheet
    Dim wsExists As Boolean
    
    wsExists = False
    
    For Each ws In ThisWorkbook.Sheets
        If ws.name = "Vstupní data" Then
            wsExists = True
        End If
    Next ws
    
    With frm
        If wsExists Then
            ' Nastavení velikosti vèetnì tlaèítka pro zachování pùvodních dat
            Height = 135
            Width = 269
        Else
            Call InputData
            Unload Me
        End If
    End With
    
End Sub

Private Sub Manual_Click()
    ' Volání metody pro stanovení vah výpoètem
    Call InputData
    Unload Me ' Zavøe formuláø po dokonèení
End Sub


Private Sub KeepCurrent_Click()
    
    'Zavøení formuláøe a pøesun na list Vstupní data
    Sheets("Vstupní data").Activate
    Unload Me

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Kontrola, zda byl formuláø zavøen pomocí tlaèítka "X" (CloseMode = 0)
    If CloseMode = vbFormControlMenu Then
        Sheets("Vstupní data").Activate
        Unload Me ' Zavíráme formuláø
    End If
End Sub

