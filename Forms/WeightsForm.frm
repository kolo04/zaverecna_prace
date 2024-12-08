VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WeightsForm 
   Caption         =   "Formul�� pro stanoven� vah"
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
    
    ' Nastaven� velikosti
    With frm
        Height = 135
        Width = 269
    End With
    
    Set ws = ThisWorkbook.Sheets("Vstupn� data")
    With ws
        .Unprotect "1234"
        .Range("D4").value = "V�ha"
        
        ' Tu�n� p�smo pro z�hlav�
        .Range("B4:D4").Font.Bold = True
        
        .Protect "1234"
    End With
    
End Sub

Private Sub Manual_Click()
    ' Vol�n� metody pro stanoven� vah v�po�tem
    Call MoveToM2
    Unload Me ' Zav�e formul�� po dokon�en�
End Sub

Private Sub Upload_Click()
    ' Vol�n� metody pro nahr�n� vah
    Call UploadWeights
    Unload Me ' Zav�e formul�� po dokon�en�
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Kontrola, zda byl formul�� zav�en pomoc� tla��tka "X" (CloseMode = 0)
    If CloseMode = vbFormControlMenu Then
        MsgBox "V�b�r metody zad�v�n� byl zru�en.", vbExclamation
        Unload Me
    End If
End Sub
