VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddCategory 
   Caption         =   "Add a category"
   ClientHeight    =   2370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5505
   OleObjectBlob   =   "AddCategory.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Add_Click()
Sheet1.Activate
Dim i As Integer, cell As Range

For Each cell In Range("A1", Range("A1").End(xlToRight)).Cells
    If StrConv(cell.Value, vbProperCase) = StrConv(NewCategoryName, vbProperCase) Then
        MsgBox "New category name already exists"
        NewCategoryName.Text = ""
        NewCategoryName.SetFocus
        Exit Sub
    End If
Next cell


' Check if the Category input is empty
If Len(Me.NewCategoryName.Text) <= 0 Then
      Me.NewCategoryName.BackColor = vbRed
      Me.NewCategoryName.SetFocus
      MsgBox "New category name cannot be empty"
Else
' Add the new category in proper case to the next empty cell in the first row
Sheet1.Range("A1").End(xlToRight).Offset(0, 1).Value = WorksheetFunction.Proper(Me.NewCategoryName.Text)
End If

UserForm_Activate

'Range("A1").Select
'Do
'    i = i + 1
'   If IsEmpty(ActiveCell.Offset(0, i - 1)) Then Exit Do
    
'Loop
'ActiveCell.Offset(0, i - 1) = Me.NewCategoryName.Text

End Sub

Private Sub NewCategoryName_Change()
Me.NewCategoryName.BackColor = vbWhite
End Sub

Private Sub Quit_Click()
Unload Me
End Sub

Private Sub UserForm_Activate()
Me.NewCategoryName = ""
Me.NewCategoryName.SetFocus
End Sub


