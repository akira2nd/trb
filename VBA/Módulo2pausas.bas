Attribute VB_Name = "Módulo2"
' Estouros | e marca os estouros e i marca a quantidade de operadores
    Private e_Lanche, e_Desc, i_Lanche, i_Desc, T_lanche, t_Desc, f_Lanche, f_Desc As Double
    Private qunt_Lanche, qunt_Desc As Integer
Sub Executa_2()

    Dim app As Application
    Set app = Application

    app.ScreenUpdating = False
    
    qunt_Lanche = 0
    qunt_Desc = 0
    
    app.StatusBar = "Importando dados do CMS"
    CMS
    app.StatusBar = "Arrumando Horários"
    arruma_hora
    app.StatusBar = "Dimensionando Equipes"
    Nome_Super
    app.StatusBar = "Extraindo Login/Logout"
    CMS_LoginLogout
    app.StatusBar = "Exportando dados para o Relatório"
    coloca_dados
    
    Range("H8").Value = Format(f_Lanche, "hh:mm:ss")
    Range("I8").Value = Format(f_Desc, "hh:mm:ss")
    
    Range("H9").Value = qunt_Lanche
    Range("I9").Value = qunt_Desc
    
    Range("A1").Select
    
    app.StatusBar = ""
    app.ScreenUpdating = True
    
End Sub
Sub CMS_2()
    
    Dim periodo As String
    Dim flg As Boolean
    Dim b As Long
    Dim I As Long
    Dim info As Variant
    Dim rep As Variant
    Dim Caminho  As String
    Dim caminhop  As String
    Dim hora, skills, Data, L As String
    
    '------------------------------------------
    Sheets("Base_Temp").Select
    Cells.Select
    Selection.ClearContents
    '------------------------------------------
    'hora = Sheets("Capa").Cells(21, 3)
    'HORARIO = Format(hora, "hh:mm:ss")
    skills = Sheets("Capa").Cells(6, 2)
    '------------------------------------------
    PLANILHA = Application.ActiveWorkbook.Name
    '------------------------------------------
    Caminho = Application.ActiveWorkbook.Path
    '----------------------------------------
    CentreVu = "10.4.0.90"
    ARQUIVO = Caminho & "\" & "Base CMS.xls"
    '------------------------------------------
    Data = Sheets("Capa").Cells(5, 2)
    periodo = Format(Data, "dd/mm/yyyy")
    '------------------------------------------
    
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
    Set info = acsSrv.Reports.Reports("Historical\Designer\Pausas Grupo/Agentes (Diário)")
    
' verifica se foi possível setar o tipo de relatório a ser criado
    If info Is Nothing Then
    
        MsgBox "O relatório Real-Time\Designer\Agentes não foi encontrado no DAC 1.", vbCritical Or vbOKOnly, "Avaya CMS Supervisor"
        Exit Sub
        
    Else

' cria o relatório do CMS
        b = acsSrv.Reports.CreateReport(info, rep)
        If b Then
        ' Seta a caixa do relatório do CMS para "quase" não ser vista
            rep.Window.Top = 0
            rep.Window.Left = 0
            rep.Window.Width = 0
            rep.Window.Height = 0
        
        ' seta as propriedades do relatório
            rep.SetProperty "Grupos", skills
            rep.SetProperty "Datas", Data
            'rep.SetProperty "Horários", HORARIO
            rep.SetProperty "DACs", "1"
            
        ' roda e exporta o relatório
            b = rep.Run
            b = rep.ExportData(ARQUIVO, 9, 0, True, True, True)
        
            If Not acsSrv.Interactive Then acsSrv.ActiveTasks.Remove rep.TaskID
            rep.Quit
            Set rep = Nothing
        End If
    End If
    
' define como nulo a variável
    Set info = Nothing

' copia as informações das planilhas para importação
    PLANILHA = Application.ActiveWorkbook.Name
    Workbooks.OpenText Filename:=ARQUIVO
    fechar = Application.ActiveWorkbook.Name
    Cells.Select
    Selection.Copy
    
' cola no arquivo
    Windows(PLANILHA).Activate
    Sheets("Base_Temp").Visible = True
    Sheets("Base_Temp").Select
    Cells.Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("Base CMS.xls").Close
    Selection.Replace What:=",000000000", Replacement:="0", LookAt:=xlPart _
    , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Rows(6).Delete Shift:=xlShiftUp
    
    Range("N6").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[-13],5)*1"
    Selection.AutoFill Destination:=Range(Selection, ActiveCell.Cells(Range("B6").End(xlDown).Row - 5, 1))
    
    Range("N6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Range("A6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("N6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

End Sub
Sub arruma_hora()

    Dim cell As Object
    
    
    Sheets("Base_Temp").Select
    
    Range("N5").Value = "Operador"
    Range("O5").Value = "Supervisor"
    
    Range("AA6:AL6").Select
    For Each cell In Selection
        With cell
            '.FormulaR1C1 = "=TEXT(RC[-25],""hh:mm:ss"")"
            .FormulaR1C1 = "=TIME(,,RC[-25])"
            .NumberFormat = "hh:mm:ss"
        End With
    Next
    Selection.AutoFill Destination:=Range(Selection, ActiveCell.Cells(Range("B6").End(xlDown).Row - 6, 1))
    
    Range(Selection, Range(ActiveCell.Address).End(xlDown)).Copy
    Range("B6").PasteSpecial xlPasteValuesAndNumberFormats
    Range("AA6:AL6", Range("AA6").End(xlDown)).Clear
    

End Sub
Sub Nome_Super()

    Sheets("Base_Temp").Select
    
    Range("N6").FormulaR1C1 = "=VLOOKUP(RC[-13],BD_Ilhas!C[-13]:C[-11],2,0)"
    Range("O6").FormulaR1C1 = "=VLOOKUP(RC[-14],BD_Ilhas!C[-14]:C[-12],3,0)"
    
    Range("N6:O6").Select
    Selection.AutoFill Destination:=Range(Selection, ActiveCell.Cells(Range("B6").End(xlDown).Row - 6, 1))
    With Range("N:O")
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats
    End With
    
    Range("A6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Base_Temp").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Base_Temp").Sort.SortFields.Add Key:=Range( _
        "O6:O65000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Base_Temp").Sort.SortFields.Add Key:=Range( _
        "A6:A65000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Base_Temp").Sort
        .SetRange Range("A6:O65000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Range("N:O").Replace What:="#N/D", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    'ActiveSheet.Range("$A$5:$O$65000").AutoFilter Field:=15, Criteria1:="#N/D" _
    , Operator:=xlOr, Criteria2:="="
    'Range("A6:O65000").SpecialCells(xlCellTypeVisible).Clear
    ActiveSheet.AutoFilterMode = False
    
    Range("O65000").End(xlUp).Select
    'MsgBox ActiveCell.Cells(1, 15).Value
    Do
        If IsError(ActiveCell.Value) Then
            Range(ActiveCell.Address, ActiveCell.Cells(1, -13).Address).Clear
            ActiveCell.Offset(-1, 0).Select
        Else
            Exit Do
        End If
    Loop

End Sub
Sub coloca_dados()

    Dim super_atual As String
    Dim Linha As Long
    
    Linha = 4
    
    Sheets("CAPA").Select
    Range("A10:Z10", Range("A65000")).Clear
    Range("A8").Select
    
    On Error Resume Next
    
' -1 para ele colocar o total de pausas no último supervisor também
    Do Until Sheets("CMS").Cells(Linha - 1, 32).Value = vbNullString
        If Sheets("Base_temp").Cells(Linha, 32).Value <> super_atual Then
        
        If Linha > 4 Then
            With Sheets("BD")
                .Select
                .Range("L3:AC3").Copy
            End With
            Sheets("CAPA").Select
            With ActiveCell
                .PasteSpecial
                
                T_lanche = e_Lanche - (i_Lanche * Sheets("BD").Range("M11"))
                t_Desc = e_Desc - (i_Desc * Sheets("BD").Range("M12"))
                
                f_Lanche = f_Lanche + T_lanche
                f_Desc = f_Desc + t_Desc
                
                .Cells(1, 8).Value = Format(T_lanche, "hh:mm:ss")
                .Cells(1, 9).Value = Format(t_Desc, "hh:mm:ss")
                .Offset(1, 0).Select
            End With
        End If
    
    ' sai do sub se acabar os teles a preencher
        If Sheets("CMS").Cells(Linha, 32).Value = vbNullString Then Exit Do
        
        ' seleciona 3 linhas abaixo da atual
            ActiveCell.Offset(3, 0).Select
        ' cria um novo grupo de supervisão
            super_atual = Sheets("CMS").Cells(Linha, 32).Value
            Range(ActiveCell.Address, Range(ActiveCell.Cells(1, 18).Address)).Merge
            With ActiveCell
                .Value = "Supervisor(a) " & super_atual
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                With .Font
                    .FontStyle = "Negrito"
                    .Size = 14
                End With
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark2
                    .TintAndShade = -0.249946592608417
                    .PatternTintAndShade = 0
                End With
            End With
            ActiveCell.Offset(1, 0).Select
            With Sheets("BD")
                .Select
                .Range("L1:AC1").Copy
            End With
            Sheets("CAPA").Select
            ActiveCell.PasteSpecial
            ActiveCell.Offset(1, 0).Select
       
       Sheets("BD").Range("L2").Value = Sheets("CMS").Cells(Linha, 1).Value
       
        ' preenche o tele
           With ActiveCell
            ' Login
            .Range(Cells(1, 1), Cells(1, 18)).Value = Sheets("BD").Range("L2:AC2").Value
            
                '.Value = Sheets("BD").Range("L2").Value
            ' Nome
                '.Cells(1, 2).Value = Sheets("BD").Range("M2").Value
            ' Tempo em Serviço
                '.Cells(1, 3).Value = Sheets("BD").Range("N2").Value
            ' Tempo em Pausa
                '.Cells(1, 4).Value = Sheets("BD").Range("O2").Value
            ' Sistema
                '.Cells(1, 5).Value = Sheets("BD").Range("P2").Value
            ' Lanche
                '.Cells(1, 6).Value = Sheets("BD").Range("Q2").Value
            ' Descanso
                '.Cells(1, 7).Value = Sheets("BD").Range("R2").Value
            ' Ambulatório
                '.Cells(1, 8).Value = Sheets("BD").Range("S2").Value
            ' Defeito
                '.Cells(1, 9).Value = Sheets("BD").Range("T2").Value
            ' BackOffice
                '.Cells(1, 10).Value = Sheets("BD").Range("U2").Value
            ' Treinamento
                '.Cells(1, 11).Value = Sheets("BD").Range("V2").Value
            ' Feedback
                '.Cells(1, 12).Value = Sheets("BD").Range("W2").Value
            ' Particular
                '.Cells(1, 13).Value = Sheets("BD").Range("X2").Value
            ' Reunião
                '.Cells(1, 14).Value = Sheets("BD").Range("Y2").Value
            ' Qunt Deslogue
                '.Cells(1, 15).Value = Sheets("BD").Range("Z2").Value
            ' Tempo deslogue
                '.Cells(1, 16).Value = Sheets("BD").Range("AA2").Value
                
                '.Cells(1, 17).Value = Sheets("BD").Range("AB2").Value
                
                '.Cells(1, 18).Value = Sheets("BD").Range("AC2").Value
                
                
            
            End With
        ' formata os valores para horaa
            Range(ActiveCell.Cells(1, 3).Address, ActiveCell.Cells(1, 4).Address).NumberFormat = "hh:mm:ss"
            Range(ActiveCell.Cells(1, 7).Address, ActiveCell.Cells(1, 18).Address).NumberFormat = "hh:mm:ss"
        ' formata o valores geral
            Range(ActiveCell.Cells(1, 5).Address, ActiveCell.Cells(1, 6).Address).NumberFormat = "General"
        ' centraliza o login e as pausas
            With ActiveCell
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            With Range(ActiveCell.Cells(1, 2).Address, ActiveCell.Cells(1, 2).Address)
                .VerticalAlignment = xlCenter
            End With
            With Range(ActiveCell.Cells(1, 4).Address, ActiveCell.Cells(1, 18).Address)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        ' coloca a borda nas linhas
            With Range(ActiveCell.Address, ActiveCell.Cells(1, 18).Address).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = -0.149937437055574
                .Weight = xlThin
            End With

        ' chama sub de estouro das pausas
            e_Lanche = 0
            e_Desc = 0
            i_Lanche = 0
            i_Desc = 0
            
            If ActiveCell.Cells(1, 8).Value > Sheets("BD").Range("M11") Then
                e_Lanche = e_Lanche + ActiveCell.Cells(1, 8).Value
                i_Lanche = i_Lanche + 1
            End If
            If ActiveCell.Cells(1, 9).Value > Sheets("BD_Ilhas").Range("K12") Then
                e_Desc = e_Desc + ActiveCell.Cells(1, 9).Value
                i_Desc = i_Desc + 1
            End If
            Call Estouro_Pausa

        Else
        
        
        ' preenche o tele
            With ActiveCell
            ' Login
                .Value = Sheets("Base_temp").Cells(Linha, 1).Value
            ' Nome
                .Cells(1, 2).Value = Sheets("Base_temp").Cells(Linha, 14).Value
            ' Tempo em Serviço
                .Cells(1, 4).Value = Sheets("Base_temp").Cells(Linha, 2).Value
            ' Tempo em Pausa
                .Cells(1, 5).Value = Sheets("Base_temp").Cells(Linha, 3).Value
            ' Sistema
                .Cells(1, 7).Value = Sheets("Base_temp").Cells(Linha, 4).Value
            ' Lanche
                .Cells(1, 8).Value = Sheets("Base_temp").Cells(Linha, 5).Value
            ' Descanso
                .Cells(1, 9).Value = Sheets("Base_temp").Cells(Linha, 6).Value
            ' Ambulatório
                .Cells(1, 10).Value = Sheets("Base_temp").Cells(Linha, 7).Value
            ' Defeito
                .Cells(1, 11).Value = Sheets("Base_temp").Cells(Linha, 8).Value
            ' BackOffice
                .Cells(1, 12).Value = Sheets("Base_temp").Cells(Linha, 9).Value
            ' Treinamento
                .Cells(1, 13).Value = Sheets("Base_temp").Cells(Linha, 10).Value
            ' Feedback
                .Cells(1, 14).Value = Sheets("Base_temp").Cells(Linha, 11).Value
            ' Particular
                .Cells(1, 15).Value = Sheets("Base_temp").Cells(Linha, 12).Value
            ' Reunião
                .Cells(1, 16).Value = Sheets("Base_temp").Cells(Linha, 13).Value
            ' Qunt Deslogue
                .Cells(1, 17).Value = Sheets("Base_temp").Cells(Linha, 16).Value
            ' Tempo deslogue
                .Cells(1, 18).Value = Sheets("Base_temp").Cells(Linha, 17).Value
            End With
        ' formata os valores para horaa
            Range(ActiveCell.Cells(1, 4).Address, ActiveCell.Cells(1, 18).Address).NumberFormat = "hh:mm:ss"
        ' formata o valores geral
            ActiveCell.Cells(1, 17).NumberFormat = "General"
        ' centraliza o login e as pausas
            With ActiveCell
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            With Range(ActiveCell.Cells(1, 2).Address, ActiveCell.Cells(1, 2).Address)
                .VerticalAlignment = xlCenter
            End With
            With Range(ActiveCell.Cells(1, 4).Address, ActiveCell.Cells(1, 18).Address)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        ' coloca a borda nas linhas
            With Range(ActiveCell.Address, ActiveCell.Cells(1, 18).Address).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = -0.149937437055574
                .Weight = xlThin
            End With
            
        ' chama sub de estouro das pausas
            If ActiveCell.Cells(1, 8).Value > Sheets("BD_Ilhas").Range("K11") Then
                e_Lanche = e_Lanche + ActiveCell.Cells(1, 8).Value
                i_Lanche = i_Lanche + 1
            End If
            If ActiveCell.Cells(1, 9).Value > Sheets("BD_Ilhas").Range("K12") Then
                e_Desc = e_Desc + ActiveCell.Cells(1, 9).Value
                i_Desc = i_Desc + 1
            End If
            Call Estouro_Pausa
            
        End If
        
        ActiveCell.Offset(1, 0).Select
        Linha = Linha + 1
    Loop
    
    Columns("B:B").EntireColumn.AutoFit
    
End Sub
Sub Estouro_Pausa()

' pausas "AMARELAS"
    Dim y_servico, y_pausa, y_sist, y_lanche, y_desc, y_ambu, y_defeito, y_bkoffice, y_trn, y_feed, y_part, y_reun As Range
    
    Set y_servico = Sheets("BD_Ilhas").Range("K20")
    Set y_pausa = Sheets("BD_Ilhas").Range("K21")
    Set y_sist = Sheets("BD_Ilhas").Range("K10")
    Set y_lanche = Sheets("BD_Ilhas").Range("K11")
    Set y_desc = Sheets("BD_Ilhas").Range("K12")
    Set y_ambu = Sheets("BD_Ilhas").Range("K13")
    Set y_defeito = Sheets("BD_Ilhas").Range("K14")
    Set y_bkoffice = Sheets("BD_Ilhas").Range("K15")
    Set y_trn = Sheets("BD_Ilhas").Range("K16")
    Set y_feed = Sheets("BD_Ilhas").Range("K17")
    Set y_part = Sheets("BD_Ilhas").Range("K18")
    Set y_reun = Sheets("BD_Ilhas").Range("K19")
    
' pausas "VERMELHAS"
    Dim r_servico, r_pausa, r_sist, r_lanche, r_desc, r_ambu, r_defeito, r_bkoffice, r_trn, r_feed, r_part, r_reun As Range
    
    Set r_servico = Sheets("BD_Ilhas").Range("L20")
    Set r_pausa = Sheets("BD_Ilhas").Range("L21")
    Set r_sist = Sheets("BD_Ilhas").Range("L10")
    Set r_lanche = Sheets("BD_Ilhas").Range("L11")
    Set r_desc = Sheets("BD_Ilhas").Range("L12")
    Set r_ambu = Sheets("BD_Ilhas").Range("L13")
    Set r_defeito = Sheets("BD_Ilhas").Range("L14")
    Set r_bkoffice = Sheets("BD_Ilhas").Range("L15")
    Set r_trn = Sheets("BD_Ilhas").Range("L16")
    Set r_feed = Sheets("BD_Ilhas").Range("L17")
    Set r_part = Sheets("BD_Ilhas").Range("L18")
    Set r_reun = Sheets("BD_Ilhas").Range("L19")

    With ActiveCell
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Tempo em Serviço
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If .Cells(1, 4).Value < y_servico Then
            With .Cells(1, 4).Font
                .FontStyle = "Negrito"
            End With
            With .Cells(1, 4).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .Color = 6737151
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
        If .Cells(1, 4).Value < r_servico Then
            With .Cells(1, 4).Font
                .FontStyle = "Negrito"
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With .Cells(1, 4).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Tempo em Pausa
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If .Cells(1, 5).Value > y_pausa Then
            With .Cells(1, 5).Font
                .FontStyle = "Negrito"
            End With
            With .Cells(1, 5).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .Color = 6737151
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
        If .Cells(1, 5).Value > r_pausa Then
            With .Cells(1, 5).Font
                .FontStyle = "Negrito"
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With .Cells(1, 5).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Pausa Sistema
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If .Cells(1, 6).Value > y_sist Then
            With .Cells(1, 6).Font
                .FontStyle = "Negrito"
            End With
            With .Cells(1, 6).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .Color = 6737151
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
        If .Cells(1, 6).Value > r_sist Then
            With .Cells(1, 6).Font
                .FontStyle = "Negrito"
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With .Cells(1, 6).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Pausa Lanche
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If .Cells(1, 8).Value > y_lanche Then
            With .Cells(1, 8).Font
                .FontStyle = "Negrito"
            End With
            With .Cells(1, 8).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .Color = 6737151
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            qunt_Lanche = qunt_Lanche + 1
        End If
        If .Cells(1, 8).Value > r_lanche Then
            With .Cells(1, 8).Font
                .FontStyle = "Negrito"
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With .Cells(1, 8).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Pausa Descanso
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If .Cells(1, 9).Value > y_desc Then
            With .Cells(1, 9).Font
                .FontStyle = "Negrito"
            End With
            With .Cells(1, 9).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .Color = 6737151
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            
            qunt_Desc = qunt_Desc + 1
        End If
        If .Cells(1, 9).Value > r_desc Then
            With .Cells(1, 9).Font
                .FontStyle = "Negrito"
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With .Cells(1, 9).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Pausa Ambulatório
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If .Cells(1, 10).Value > y_ambu Then
            With .Cells(1, 10).Font
                .FontStyle = "Negrito"
            End With
            With .Cells(1, 10).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .Color = 6737151
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
        If .Cells(1, 10).Value > r_ambu Then
            With .Cells(1, 10).Font
                .FontStyle = "Negrito"
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With .Cells(1, 10).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Pausa Defeito
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If .Cells(1, 11).Value > y_defeito Then
            With .Cells(1, 11).Font
                .FontStyle = "Negrito"
            End With
            With .Cells(1, 11).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .Color = 6737151
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
        If .Cells(1, 11).Value > r_defeito Then
            With .Cells(1, 11).Font
                .FontStyle = "Negrito"
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With .Cells(1, 11).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Pausa BackOffice
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If .Cells(1, 12).Value > y_bkoffice Then
            With .Cells(1, 12).Font
                .FontStyle = "Negrito"
            End With
            With .Cells(1, 12).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .Color = 6737151
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
        If .Cells(1, 12).Value > r_bkoffice Then
            With .Cells(1, 12).Font
                .FontStyle = "Negrito"
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With .Cells(1, 12).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Pausa Treinamento
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If .Cells(1, 13).Value > y_trn Then
            With .Cells(1, 13).Font
                .FontStyle = "Negrito"
            End With
            With .Cells(1, 13).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .Color = 6737151
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
        If .Cells(1, 13).Value > r_trn Then
            With .Cells(1, 13).Font
                .FontStyle = "Negrito"
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With .Cells(1, 13).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Pausa Feedback
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If .Cells(1, 14).Value > y_feed Then
            With .Cells(1, 14).Font
                .FontStyle = "Negrito"
            End With
            With .Cells(1, 14).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .Color = 6737151
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
        If .Cells(1, 14).Value > r_feed Then
            With .Cells(1, 14).Font
                .FontStyle = "Negrito"
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With .Cells(1, 14).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Pausa Particular
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If .Cells(1, 15).Value > y_part Then
            With .Cells(1, 15).Font
                .FontStyle = "Negrito"
            End With
            With .Cells(1, 15).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .Color = 6737151
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
        If .Cells(1, 15).Value > r_part Then
            With .Cells(1, 15).Font
                .FontStyle = "Negrito"
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With .Cells(1, 15).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Pausa Reunião
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If .Cells(1, 16).Value > y_reun Then
            With .Cells(1, 16).Font
                .FontStyle = "Negrito"
            End With
            With .Cells(1, 16).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .Color = 6737151
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
        If .Cells(1, 16).Value > r_reun Then
            With .Cells(1, 16).Font
                .FontStyle = "Negrito"
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With .Cells(1, 16).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
        
    
'    ' Pausa Descanso
'        If .Cells(1, 9).Value > pausa Then
'            With .Cells(1, 9).Font
'                .FontStyle = "Negrito"
'                .ThemeColor = xlThemeColorDark1
'                .TintAndShade = 0
'                .ThemeFont = xlThemeFontMinor
'            End With
'            With .Cells(1, 9).Interior
'                .Pattern = xlSolid
'                .PatternColorIndex = xlAutomatic
'                .ThemeColor = xlThemeColorAccent2
'                .TintAndShade = 0
'                .PatternTintAndShade = 0
'            End With
'        End If
'    ' Pausa Particular
'        If .Cells(1, 15).Value > "0" And .Cells(1, 15).Value < particular Then
'            With .Cells(1, 15).Interior
'                .Pattern = xlSolid
'                .PatternColorIndex = xlAutomatic
'                .ThemeColor = xlThemeColorAccent2
'                .Color = 6737151
'                .TintAndShade = 0
'                .PatternTintAndShade = 0
'            End With
'            With .Cells(1, 15).Font
'                .FontStyle = "Negrito"
'                .PatternColorIndex = xlAutomatic
'                .TintAndShade = 0
'                .ThemeFont = xlThemeFontMinor
'            End With
'        End If
'        If .Cells(1, 15).Value >= particular Then
'            With .Cells(1, 15).Font
'                .FontStyle = "Negrito"
'                .ColorIndex = 0
'                .ThemeColor = xlThemeColorDark1
'                .TintAndShade = 0
'                .ThemeFont = xlThemeFontMinor
'            End With
'            With .Cells(1, 15).Interior
'                .Pattern = xlSolid
'                .PatternColorIndex = xlAutomatic
'                .ThemeColor = xlThemeColorAccent2
'                .TintAndShade = 0
'                .PatternTintAndShade = 0
'            End With
'        End If
    End With


    Set y_servico = Nothing
    Set y_pausa = Nothing
    Set y_sist = Nothing
    Set y_lanche = Nothing
    Set y_desc = Nothing
    Set y_ambu = Nothing
    Set y_defeito = Nothing
    Set y_bkoffice = Nothing
    Set y_trn = Nothing
    Set y_feed = Nothing
    Set y_part = Nothing
    Set y_reun = Nothing
    
    Set r_servico = Nothing
    Set r_pausa = Nothing
    Set r_sist = Nothing
    Set r_lanche = Nothing
    Set r_desc = Nothing
    Set r_ambu = Nothing
    Set r_defeito = Nothing
    Set r_bkoffice = Nothing
    Set r_trn = Nothing
    Set r_feed = Nothing
    Set r_part = Nothing
    Set r_reun = Nothing


End Sub
Sub Soma_Pausas()

    ActiveCell.Value = Sum()
End Sub
Sub CMS_LoginLogout()

    Dim flg As Boolean
    Dim b As Long
    Dim I As Long
    Dim info As Variant
    Dim rep As Variant
    Dim Caminho  As String
    Dim caminhop  As String
    Dim hora, skills, Data, L As String
    
    '------------------------------------------
    'hora = Sheets("Capa").Cells(21, 3)
    'HORARIO = Format(hora, "hh:mm:ss")
    skills = Sheets("Capa").Cells(6, 2)
    '------------------------------------------
    PLANILHA = Application.ActiveWorkbook.Name
    '------------------------------------------
    Caminho = Application.ActiveWorkbook.Path
    '----------------------------------------
    CentreVu = "10.4.0.90"
    ARQUIVO = Caminho & "\" & "Base CMS.xls"
    '------------------------------------------
    Data = Sheets("Capa").Cells(5, 2)
    periodo = Format(Data, "dd/mm/yyyy")
    '------------------------------------------
    
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
    Set info = acsSrv.Reports.Reports("Historical\Designer\Login/Logout (Especialidade) [Grupos Diversos]")
    
' verifica se foi possível setar o tipo de relatório a ser criado
    If info Is Nothing Then
    
        MsgBox "O relatório Historical\Designer\Login/Logout (Especialidade) [Grupos Diversos] não foi encontrado no DAC 1.", vbCritical Or vbOKOnly, "Avaya CMS Supervisor"
        Exit Sub
        
    Else

' cria o relatório do CMS
        b = acsSrv.Reports.CreateReport(info, rep)
        If b Then
        ' Seta a caixa do relatório do CMS para "quase" não ser vista
            rep.Window.Top = 0
            rep.Window.Left = 0
            rep.Window.Width = 0
            rep.Window.Height = 0
        
        ' seta as propriedades do relatório
            rep.SetProperty "Grupo", skills
            rep.SetProperty "Data", Data
            'rep.SetProperty "Horários", HORARIO
            rep.SetProperty "DACs", "1"
            
        ' roda e exporta o relatório
            b = rep.Run
            b = rep.ExportData(ARQUIVO, 9, 0, True, True, True)
        
            If Not acsSrv.Interactive Then acsSrv.ActiveTasks.Remove rep.TaskID
            rep.Quit
            Set rep = Nothing
        End If
    End If
    
' define como nulo a variável
    Set info = Nothing

' copia as informações das planilhas para importação
    PLANILHA = Application.ActiveWorkbook.Name
    Workbooks.OpenText Filename:=ARQUIVO
    fechar = Application.ActiveWorkbook.Name
    Range("C:C,E:E").Delete Shift:=xlShiftLeft
    Range("A:D").Copy
    
' cola no arquivo
    Windows(PLANILHA).Activate
    Sheets("Base_Temp").Visible = True
    Sheets("Base_Temp").Select
    Range("T1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("Base CMS.xls").Close SaveChanges:=False
    
    
    Range("AA4").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[-7],5)*1"
    Range("AA4").Select
    Selection.AutoFill Destination:=Range(Selection, ActiveCell.Cells(Range("T4").End(xlDown).Row - 3, 1))
        
    Range("AA4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    Range("T4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AA4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.ClearContents
       
    With Range("X4", Range("U4").End(xlDown).Offset(0, 3).Address)
        .FormulaR1C1 = "=IF(RC[-4]<>R[-1]C[-4],0,RC[-2]-R[-1]C[-1])"
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats
    End With
    
    Range("P6").FormulaR1C1 = "=IF((COUNTIF(C[4],RC[-15])-1)<0,0,COUNTIF(C[4],RC[-15])-1)"
    Range("Q6").FormulaR1C1 = "=SUMIF(C[3],RC[-16],C[7])/60"
    With Range("P6:Q6")
        .AutoFill Destination:=Range(Range("P6:Q6"), Range("O6").End(xlDown).Offset(0, 2).Address)
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats
    End With
    Range("Q6", Range("Q6").End(xlDown)).NumberFormat = "hh:mm:ss" '.NumberFormat = "[$-F400]h:mm:ss AM/PM"
    
    Range("R6").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-17],C[2]:C[4],3,0)"
    Range("S6").Select
    ActiveCell.FormulaR1C1 = "=INDEX(C[1]:C[4],MATCH(RC[-18],C[1],0)+RC[-3],4)"
    
    Range("R6:S6").Select
    Selection.AutoFill Destination:=Range(Selection, ActiveCell.Cells(Range("B6").End(xlDown).Row - 5, 1))
    
    Range("Q6").Select
    Selection.Copy
    Columns("R:S").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("R6:S6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
    
    

End Sub

