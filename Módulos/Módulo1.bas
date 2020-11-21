Attribute VB_Name = "Módulo1"
Option Explicit

Sub CommandButtonAbrirForm()
    On Error Resume Next
        If UserForm1.Visible = False Then
            UserForm1.Show
        End If
End Sub
