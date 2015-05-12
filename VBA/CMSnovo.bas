Attribute VB_Name = "CMS"
Sub CMS_RUN()

    Application.Calculation = xlManual
    

    Dim linha As Integer
    Dim Arquivo As String

Application.ScreenUpdating = False

    linha = 7 'linha que tem o primeiro dado de CMS

    With Sheets("INICIO")
    
        Do Until .Cells(linha, 2) = vbNullString
        
            Arquivo = "\\brsjcsrv01\Operacoes\Speedy\Planejamento\_ComumSpeedy\Desempenho\TEMP_CMS\CMS_" & .Cells(linha, 2) & "_" & Format(.Cells(2, 3), "yyyymmdd") & ".xls"
            
            Call Extração_CMS(.Cells(linha, 3), .Cells(linha, 4), Arquivo, _
                .Cells(2, 3), Format(TimeValue("00:00:00"), "hh:mm") & "-" & Format(.Cells(4, 3), "hh:mm"))

            Call Edita_Extração(Arquivo, .Cells(linha, 2))

            linha = linha + 1
            
        Loop
        
    End With

Application.ScreenUpdating = True

Application.Calculation = xlAutomatic

Application.StatusBar = "CMS OK"

End Sub
Sub Extração_CMS(CentreVu As String, Skills As String, Arquivo As String, data As String, Hora As String)


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
    Select Case CentreVu
        Case Is = "10.189.0.92": Set info = acsSrv.Reports.Reports("Historical\Designer\Desempenho do Servico (INTERVALO) [para Grupos Diversos - MIS] PLANEJAMENTO")
        Case Is = "10.4.0.90": Set info = acsSrv.Reports.Reports("Historical\Designer\Desempenho do Servico (INTERVALO) [para Grupos Diversos - MIS]")
    End Select
    
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
            
            'rep.ReportView.Add "G3,1,3;-1,2,0", "Grid1"
            rep.SetProperty "Grupos/Especialidades", Skills
            rep.SetProperty "Data", data
            rep.SetProperty "Horários", Hora
            
            b = rep.Run
            b = rep.ExportData(Arquivo, 9, 0, True, True, True)
            
        If Not acsSrv.Interactive Then acsSrv.ActiveTasks.Remove rep.TaskID
            rep.Quit
            Set rep = Nothing
        End If
        
    End If
    Set info = Nothing


End Sub

Sub Edita_Extração(Arquivo As String, Site As String)

Dim L

'    Dim x As Integer
'    Dim linha As Integer
'    Dim hora_inicial As Date
    Dim file_main As Workbook
    Dim file_extracao As Workbook
'    Dim Cell As Object

    Set file_main = ThisWorkbook
    
    Sheets("CMS_" & Site).Select
    Cells.Select
    Selection.ClearContents
    
    'abre o arquivo de extração do CMS
    Workbooks.Open Arquivo
    Set file_extracao = ActiveWorkbook
        
    Cells.Select
    Cells.Replace What:=",000000000", Replacement:="0", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    
    Cells.Select
    Selection.Copy
    
    
    file_main.Activate
          
    Sheets("CMS_" & Site).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
L = 7

Do While Sheets("CMS_" & Site).Cells(L, 1).Value <> ""

    If Cells(L + 1, 1).Value = "Horário" Then
    
    Cells(L + 1, 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Cut
    Range("AW7").Select
    ActiveSheet.Paste
    
    End If
    
    L = L + 1

Loop


    file_extracao.Close SaveChanges:=False

End Sub
