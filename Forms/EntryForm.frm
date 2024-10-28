VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EntryForm 
   Caption         =   "Formul�� pro vkl�d�n� dat"
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
        If ws.name = "Vstupn� data" Then
            wsExists = True
        End If
    Next ws
    
    With frm
        If wsExists Then
            ' Nastaven� velikosti v�etn� tla��tka pro zachov�n� p�vodn�ch dat
            Height = 135
            Width = 269
        Else
            Call InputData
            Unload Me
        End If
    End With
    
End Sub

Private Sub Manual_Click()
    ' Vol�n� metody pro stanoven� vah v�po�tem
    Call InputData
    Unload Me ' Zav�e formul�� po dokon�en�
End Sub


Private Sub KeepCurrent_Click()
    
    'Zav�en� formul��e a p�esun na list Vstupn� data
    Sheets("Vstupn� data").Activate
    Unload Me

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Kontrola, zda byl formul�� zav�en pomoc� tla��tka "X" (CloseMode = 0)
    If CloseMode = vbFormControlMenu Then
        Sheets("Vstupn� data").Activate
        Unload Me ' Zav�r�me formul��
    End If
End Sub

