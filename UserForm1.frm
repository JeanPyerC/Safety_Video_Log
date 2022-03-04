VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "User Info"
   ClientHeight    =   2860
   ClientLeft      =   -30
   ClientTop       =   -150
   ClientWidth     =   5550
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComboBox3_Change()

End Sub

Private Sub CommandButton4_Click()

Application.DisplayAlerts = False
ThisWorkbook.Save
Application.DisplayAlerts = True

If TextBox1.Value = "" Or TextBox2.Value = "" Or ComboBox3.Value = "" Then
    MsgBox "Form is not complete. Please Fill-out form."
    'If MsgBox("Form is not complete. Do you want to contine?", vbQuestion + vbYesNo) <> vbYes Then
    Exit Sub
End If

Application.ScreenUpdating = False
Sheets("DATA RECORD").Visible = True
Sheets("DATA RECORD").Select

Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select

ActiveCell = TextBox1.Value
ActiveCell.Offset(0, 1) = TextBox2.Value
ActiveCell.Offset(0, 2) = ComboBox3.Value

Application.DisplayAlerts = False
ThisWorkbook.Save
Application.DisplayAlerts = True

ActiveCell.Offset(1, 0).Select

Sheets("SAFETY VIDEO").Select

Sheets("DATA RECORD").Visible = False

Application.DisplayAlerts = False
ThisWorkbook.Save
Application.DisplayAlerts = True

Call ResetForm

CreateObject("Shell.Application").Open ("C:\Users\Owner\OneDrive\EXCEL\PERSONAL PROJECTS\SAFETY VIDEO RECORDER\COVID19_Planning_Ahead.mp4")

End Sub

Sub ResetForm()

TextBox1.Value = ""
TextBox2.Value = ""
ComboBox3.Value = ""

End Sub


Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_Click()

End Sub
