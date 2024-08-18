VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchReplace 
   Caption         =   "Search & Replace Tool"
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4920
   OleObjectBlob   =   "SearchReplace.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SearchReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()
Unload Me
End Sub

Private Sub Replace_Click()
Dim a As String
a = InputBox("Replace with what?")
If Len(a) <= 0 Then Exit Sub

ActiveCell.Offset(Me.Category.ListIndex + 1, Me.ComboBox1.ListIndex).Value = a
Me.TextBox1 = a


End Sub

Private Sub Search_Click()
Range("A1").Select

Me.TextBox1 = ActiveCell.Offset(Me.Category.ListIndex + 1, Me.ComboBox1.ListIndex).Value


End Sub

Private Sub UserForm_Activate()
Dim ws As Worksheet
    Dim cell As Range
    Dim lastColumn As Long
    
    ' Set the worksheet (change "Sheet1" to the name of your sheet)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Loop through each cell in the first row until the last non-empty cell
    For Each cell In ws.Range("A1", Range("A1").End(xlToRight))
        ' Add the cell value to the ComboBox
        Me.ComboBox1.AddItem cell.Value
    Next cell
    Me.ComboBox1.ListIndex = 0
    
       
    
    ' Set the worksheet (change "Sheet1" to the name of your sheet)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Loop through each cell in the first row until the last non-empty cell
    For Each cell In ws.Range("A2", Range("A2").End(xlDown))
        ' Add the cell value to the ComboBox
        Me.Category.AddItem cell.Value
    Next cell
    
    Me.Category.ListIndex = 0
    
    
    
    
End Sub


