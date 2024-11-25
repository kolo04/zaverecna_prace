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
    
    ThisWorkbook.Activate
    
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
            Unload Me
            Call InputData
            End
        End If
    End With
    
End Sub

Private Sub Manual_Click()
    Unload Me
    
    ' Volání metody pro vkládání nového pøíkladu
    Call InputData
End Sub


Private Sub KeepCurrent_Click()
    'Zavøení formuláøe a pøesun na list Vstupní data
    Unload Me
    Sheets("Vstupní data").Activate

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Kontrola, zda byl formuláø zavøen pomocí tlaèítka "X" (CloseMode = 0)
    If CloseMode = vbFormControlMenu Then
        Unload Me ' Zavíráme formuláø
        Sheets("Vstupní data").Activate
    End If
End Sub

