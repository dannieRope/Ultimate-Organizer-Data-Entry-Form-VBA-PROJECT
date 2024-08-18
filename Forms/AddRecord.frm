VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddRecord 
   Caption         =   "Add Record"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5895
   OleObjectBlob   =   "AddRecord.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1


Private Sub Add_Click()

Range("A1").End(xlDown).Offset(1, 0).Value = StrConv(TextBox1.Text, vbProperCase)
Range("A1").End(xlDown).Offset(0, 1).Value = StrConv(TextBox2.Text, vbProperCase)
Range("A1").End(xlDown).Offset(0, 2).Value = TextBox3.Text
Range("A1").End(xlDown).Offset(0, 3).Value = StrConv(TextBox4.Text, vbLowerCase)
Range("A1").End(xlDown).Offset(0, 4).Value = TextBox5.Text
Range("A1").End(xlDown).Offset(0, 5).Value = TextBox6.Text
Range("A1").End(xlDown).Offset(0, 6).Value = TextBox12.Text
Range("A1").End(xlDown).Offset(0, 7).Value = TextBox11.Text
Range("A1").End(xlDown).Offset(0, 8).Value = TextBox10.Text
Range("A1").End(xlDown).Offset(0, 9).Value = TextBox9.Text
Range("A1").End(xlDown).Offset(0, 10).Value = TextBox8.Text
Range("A1").End(xlDown).Offset(0, 11).Value = TextBox7.Text

Unload Me
    

End Sub

Private Sub Cancel_Click()
Unload Me
End Sub



Private Sub Label2_Click()

End Sub

Private Sub UserForm_Activate()

    Dim Caption() As String
    Dim i As Integer
    Dim nc As Integer
    Dim ws As Worksheet
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Determine the number of columns in the first row
    nc = ws.Range("A1", ws.Range("A1").End(xlToRight)).Cells.Count
    
    ' Resize the Caption array to hold the number of columns
    ReDim Caption(1 To nc)
    
    ' Populate the Caption array with the values from the first row
    For i = 1 To nc
        Caption(i) = ws.Range("A1", ws.Range("A1").End(xlToRight)).Cells(1, i).Value
    Next i
    
    With AddRecord
        TextBox1.Visible = False
        TextBox2.Visible = False
        TextBox3.Visible = False
        TextBox4.Visible = False
        TextBox5.Visible = False
        TextBox6.Visible = False
        TextBox7.Visible = False
        TextBox8.Visible = False
        TextBox9.Visible = False
        TextBox10.Visible = False
        TextBox11.Visible = False
        TextBox12.Visible = False
    End With
    ' Set the Label1 caption to a combined string of all values in the Caption array
    If nc >= 1 Then
        Me.Label1.Caption = Caption(1)
        Me.TextBox1.Visible = True
    End If
    
     If nc >= 2 Then
        Me.Label2.Caption = Caption(2)
        Me.TextBox2.Visible = True
    End If
    
     If nc >= 3 Then
        Me.Label3.Caption = Caption(3)
        Me.TextBox3.Visible = True
    End If
    
     If nc >= 4 Then
        Me.Label4.Caption = Caption(4)
        Me.TextBox4.Visible = True
    End If
    
     If nc >= 5 Then
        Me.Label5.Caption = Caption(5)
        Me.TextBox5.Visible = True
    End If
    
     If nc >= 6 Then
        Me.Label6.Caption = Caption(6)
        Me.TextBox6.Visible = True
    End If
    
     If nc >= 7 Then
        Me.Label12.Caption = Caption(7)
        Me.TextBox12.Visible = True
        Me.Width = 497
        Me.Cancel.Left = 366
    End If
    
     If nc >= 8 Then
        Me.Label11.Caption = Caption(8)
        Me.TextBox11.Visible = True
    End If
    
     If nc >= 9 Then
        Me.Label10.Caption = Caption(9)
        Me.TextBox10.Visible = True
    End If
    
    If nc >= 10 Then
        Me.Label9.Caption = Caption(10)
        Me.TextBox9.Visible = True
    End If
    
    If nc >= 11 Then
        Me.Label8.Caption = Caption(11)
        Me.TextBox8.Visible = True
    End If
    
    If nc = 12 Then
        Me.Label7.Caption = Caption(12)
        Me.TextBox7.Visible = True
    End If
    
End Sub


