VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12225
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Sub CommandButton1_Click()

    Dim pasta As String
    
    pasta = PickFolder
    
    If pasta = "" Then
        MsgBox "É NECESSÁRIO SELECIONAR UMA PASTA!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    ThisWorkbook.Sheets("Planilha1").Cells(200, 70) = pasta
    UserForm1.TextBox1 = ThisWorkbook.Sheets("Planilha1").Cells(200, 70)
    
End Sub

Public Function PickFolder() As String

    If Application.FileDialog(4).Show Then
        PickFolder = Application.FileDialog(4).SelectedItems(1)
    End If

End Function

Private Sub CommandButton3_Click()

On Error GoTo err_handler

    If UserForm1.TextBox1 = "" Or UserForm1.TextBox2 = "" Then
        MsgBox "É NECESSÁRIO SELECIONAR AS DUAS PASTAS!", vbExclamation, "ATENÇÃO!"
        Exit Sub
    End If
  
    If Dir(UserForm1.TextBox1, vbDirectory) = "" Then
        MsgBox "A PASTA DE ARQUIVOS EXCEL NÃO EXISTE!", vbExclamation, "ATENÇÃO!"
        Exit Sub
    End If
  
    If Dir(UserForm1.TextBox2, vbDirectory) = "" Then
        MsgBox "A PASTA DE DESTINO NÃO EXISTE!", vbExclamation, "ATENÇÃO!"
        Exit Sub
    End If
  

    Dim fso As Object
    Dim fl As Object
    Dim folder As Object
    Dim exc As Object
    Dim wbSaida As Object
    Dim wb As Object
    Dim st As Object

    Dim total As Long: total = 0
    Dim Executados As Long: Executados = 0
    Dim nFl As Long
    Dim contador As Long: contador = 0
  
    Dim barraTotal As Long
    Dim qtdAndar As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(UserForm1.TextBox1)

    For Each fl In folder.Files
        If Right(fl.Name, 3) = "xls" Or Right(fl.Name, 4) = "xlsx" Or Right(fl.Name, 4) = "xlsm" Then
        total = total + 1
    End If
    Next fl
  
    If total = 0 Then
        MsgBox "NÃO EXISTEM ARQUIVOS EXCEL NA PASTA SELECIONADA!", vbExclamation, "ATENÇÃO!"
        Exit Sub
    End If
  
    barraTotal = Label4.Width - 50
    qtdAndar = barraTotal / total
    Label5.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    Label3.Caption = ""
    Label5.Width = "24,05"
  
    Set exc = CreateObject("Excel.Application")
    exc.DisplayAlerts = False
    exc.Visible = False

    Set wbSaida = exc.Workbooks.Add

    For Each fl In folder.Files

        If Right(fl.Name, 3) = "xls" Or Right(fl.Name, 4) = "xlsx" Or Right(fl.Name, 4) = "xlsm" Then
        nFl = nFl + 1
        Label3.Caption = "Copiando Arquivos " & nFl & " de " & total

        Set wb = exc.Workbooks.Open(fl.Path)
        

        For Each st In wb.Sheets
            st.Copy after:=wbSaida.Sheets(wbSaida.Sheets.Count)
            contador = contador + 1
            Next st

            Executados = Executados + 1
            wb.Close
            Set wb = Nothing

            Label5.Width = Label5.Width + qtdAndar
        End If
    Next fl
  

    wbSaida.SaveAs (TextBox2 & "\NovaPlanilha " & Format(Now, "dd\-mm\-yyyy\ hh\Hmm\Mss\S") & ".xlsx")
  

err_handler:

    If Err.Number <> 0 Then
        MsgBox Err.Number & "  " & Err.Description, vbCritical, "HOUVE UM ERRO!"
        exc.Quit
        Set exc = Nothing
    Else
  
        Call MsgBox("PROCESSO FINALIZADO COM SUCESSO!!!" & vbNewLine & "ARQUIVO: " & wbSaida.Path & vbNewLine & "ARQUIVOS COMBINADOS: " & Executados & " ARQUIVOS DE " & total & vbNewLine & "TOTAL DE PLANILHAS: " & contador, vbInformation, "PROCESSO FINALIZADO!")

        If MsgBox("DESEJA ABRIR O ARQUIVO?", vbInformation + vbYesNo, "PROCESSO FINALIZADO!") = vbYes Then
            exc.Visible = True
        Else

        wbSaida.Close
        exc.Quit
        
        Set exc = Nothing
        
        End If
    End If
  
    Label5.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    UserForm1.Hide
    Exit Sub
    Resume
    
End Sub

Private Sub CommandButton2_Click()

    Dim pasta As String
    
    pasta = PickFolder
    
    If pasta = "" Then
        MsgBox "É NECESSÁRIO SELECIONAR UMA PASTA!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    ThisWorkbook.Sheets("Planilha1").Cells(200, 71) = pasta
    UserForm1.TextBox2 = ThisWorkbook.Sheets("Planilha1").Cells(200, 71)

End Sub

Private Sub UserForm1_Initialize()

    Principal.TextBox1 = ThisWorkbook.Sheets("Planilha1").Cells(200, 70)
    Principal.TextBox2 = ThisWorkbook.Sheets("Planilha1").Cells(200, 71)
    
End Sub
