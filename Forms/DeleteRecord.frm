VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteRecord 
   Caption         =   "Delete Record"
   ClientHeight    =   2085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5355
   OleObjectBlob   =   "DeleteRecord.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DeleteRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()
Unload Me
End Sub

Private Sub Delete_Click()
Dim a As String
a = MsgBox("Are you sure, you want to delete this record?", vbYesNo)

If a = vbYes Then
Range("A1").Offset(Me.Record.ListIndex + 1, 0).EntireRow.Delete
Unload Me
Exit Sub
End If
End Sub


Private Sub UserForm_Activate()
Dim ws As Worksheet
    Dim cell As Range
    Dim lastColumn As Long
    
    ' Set the worksheet (change "Sheet1" to the name of your sheet)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Loop through each cell in the first row until the last non-empty cell
    For Each cell In ws.Range("A2", Range("A2").End(xlDown))
        ' Add the cell value to the ComboBox
        Me.Record.AddItem cell.Value
    Next cell
    
Me.Record.ListIndex = 0
    
End Sub

