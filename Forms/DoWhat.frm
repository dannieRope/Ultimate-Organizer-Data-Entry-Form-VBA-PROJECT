VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DoWhat 
   Caption         =   "What do you want to do?"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6195
   OleObjectBlob   =   "DoWhat.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DoWhat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AddCat_Click()
Call addCategorysub
End Sub

Private Sub AddRecord_Click()

 Dim AddRecord As New AddRecord
AddRecord.Show
End Sub

Private Sub DeleteCategory_Click()
Dim DeleteCategory As New DeleteCategory
DeleteCategory.Show

End Sub

Private Sub DeleteRecord_Click()
Dim DeleteRecord As New DeleteRecord
DeleteRecord.Show

End Sub

Private Sub Quit_Click()
Unload Me
End Sub

Private Sub SearchReplace_Click()
Dim SearchReplace As New SearchReplace
SearchReplace.Show

End Sub

Sub AddtoComboBox()
    

End Sub

Sub addCategorysub()
Dim AddCategory As New AddCategory
 AddCategory.Show
End Sub


Private Sub UserForm_Click()

End Sub
