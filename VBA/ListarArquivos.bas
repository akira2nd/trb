Attribute VB_Name = "Módulo1"
Sub ResgistrarEntrada()

    Dim l As Long
    
    l = 2
    While Sheets("Timesheet").Cells(l, 2).Value <> ""
        l = l + 1
    Wend
    
    Sheets("Timesheet").Cells(l, 2).Value = Now()
    
End Sub
Sub RegistrarSaida()
    
    Dim l As Long
    
    l = 2
    While Sheets("Timesheet").Cells(l, 3).Value <> ""
        l = l + 1
    Wend
    
    Sheets("Timesheet").Cells(l, 3).Value = Now()
   
End Sub

Public Function ListaArquivos(ByVal Caminho As String) As String()
    'Atenção: Faça referência à biblioteca Micrsoft Scripting Runtime
    Dim FSO As New FileSystemObject
    Dim result() As String
    Dim Pasta As Folder
    Dim Arquivo As File
    Dim Indice As Long
 
 
    ReDim result(0) As String
    If FSO.FolderExists(Caminho) Then
        Set Pasta = FSO.GetFolder(Caminho)
 
        For Each Arquivo In Pasta.Files
            Indice = IIf(result(0) = "", 0, Indice + 1)
            ReDim Preserve result(Indice) As String
            result(Indice) = Arquivo.Name
        Next
    End If
 
    ListaArquivos = result
ErrHandler:
    Set FSO = Nothing
    Set Pasta = Nothing
    Set Arquivo = Nothing
End Function

Private Sub ListarArquivos()
    Dim arquivos() As String
    Dim lCtr As Long
    arquivos = ListaArquivos("C:\temp")
    For lCtr = 0 To UBound(arquivos)
      Debug.Print arquivos(lCtr)
    Next
End Sub
