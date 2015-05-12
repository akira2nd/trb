Attribute VB_Name = "Módulo2"
Sub CMS()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim caminho, L, PASTA, ORIGEM, DESTINO

    Dim DIA As String
    Dim col As Integer
    Dim ramal As String
    
        Dim periodo As String
        Dim flg As Boolean
        Dim b As Long
        Dim I As Long
        Dim info As Variant
        Dim rep As Variant
        Dim ARQUIVO  As String
        Dim caminhop  As String
        Dim hora, skills, data As String
    

data = Sheets("INICIO").Range("B1")
periodo = Format(data, "dd/mm/yyyy")
CentreVu = "10.4.0.90"

caminho = Application.ActiveWorkbook.Path
ARQUIVO = caminho & "\cms.txt"
'ARQUIVO = Sheets("INICIO").Range("B9").Value
ORIGEM = ActiveWorkbook.Name

L = 13

Do While Sheets("INICIO").Cells(L, 2) <> ""

skills = Sheets("INICIO").Cells(L, 2).Value

        Set acsApp = CreateObject("ACSUP.cvsApplication")
        CentreVuOpen = acsApp.Servers.Count
                    
        If (CentreVuOpen > 0) Then
            For I = 1 To CentreVuOpen
                If (acsApp.Servers.Item(I).Name = CentreVu) Then
                    Set acsSrv = acsApp.Servers.Item(I)
                    flg = True
                    Exit For
                End If
            Next
        Else
            MsgBox "Antes de tudo conecte o CMS!!!", vbOKOnly, "Planejamento"
            Exit Sub
        End If
        If Not flg Then
            MsgBox ("O servidor não foi localizado")
        End If
        
    Set acsCatalog = acsSrv.Reports
        acsSrv.Reports.ACD = "1"
        Set info = acsSrv.Reports.Reports("Historical\Designer\Desempenho do Servico (INTERVALO) Speedy - Recep.")
If info Is Nothing Then
        MsgBox "O relatório Real-Time\Designer\Agentes não foi encontrado no DAC 1.", vbCritical Or vbOKOnly, "Avaya CMS Supervisor"
    Exit Sub
Else
                b = acsSrv.Reports.CreateReport(info, rep)
            If b Then
                rep.Window.Top = 0
                rep.Window.Left = 0
                rep.Window.Width = 0
                rep.Window.Height = 0
                
                
          rep.SetProperty "Grupos/Especialidades", skills
          rep.SetProperty "Data", data
          rep.SetProperty "Horários", "00:00-23:30"
          rep.SetProperty "DACs", "1"
                
                
              
                b = rep.Run
                b = rep.ExportData(ARQUIVO, 9, 0, True, True, True)
                If Not acsSrv.Interactive Then acsSrv.ActiveTasks.Remove rep.TaskID
                rep.Quit
                Set rep = Nothing
            End If
        End If
Set info = Nothing








PASTA = Sheets("INICIO").Cells(L, 1)

Sheets(PASTA).Visible = True

Workbooks.OpenText Filename:=ARQUIVO, Origin _
    :=xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
    xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
    Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
    Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
    Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15 _
    , 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array(29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array(42, 1), Array(43, 1), Array(44, 1), Array(45, 1), Array(46, 1), Array(47, 1), Array(48, 1), Array(49, 1)), TrailingMinusNumbers:=True
    
    
DESTINO = ActiveWorkbook.Name

Cells.Select
Cells.Copy

Windows(ORIGEM).Activate
Sheets(PASTA).Select
Range("A1").Select
Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        Selection.Replace What:=",000000000", Replacement:="0", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
                
                Windows(DESTINO).Activate
                Windows(DESTINO).Close
       

                For L1 = 9 To 70
                
                    If Cells(L1, 1).Value = Cells(6, 1).Value Then
                        Cells(L1, 1).Select
                        Range(Selection, Selection.End(xlDown)).Select
                        Range(Selection, Selection.End(xlToRight)).Select
                        Selection.Cut
                        Range("Z6").Select
                        ActiveSheet.Paste
                        Range("A1").Select
                        
                        
                    End If
            Next
            
            Sheets(PASTA).Visible = False
            Application.CutCopyMode = False
            L = L + 1
            
            Loop
   
             Sheets("INICIO").Select

Application.CutCopyMode = False

Application.StatusBar = "CMS ATUALIZADO"

End Sub

