Attribute VB_Name = "RODAR"
Sub Executa()

    Dim L 'Col
    'Dim skill
    Dim app As Application
    Set app = Application

    app.ScreenUpdating = False
    
    'qunt_Lanche = 0
    'qunt_Desc = 0

L = 5
Do While Sheets("INICIO").Cells(L, 1).Value <> ""

ABA = Sheets("INICIO").Cells(L, 1).Value
    Sheets(ABA).Select
    Rows("11:4621").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp

L = L + 1

Loop



L = 5
'Col = 2


Do While Sheets("INICIO").Cells(L, 3).Value <> ""

If Sheets("INICIO").Cells(L, 5).Value = "S" Then
    
    app.StatusBar = "Importando dados do CMS"
    Call CMS(L)
    app.StatusBar = "Extraindo Login/Logout"
    Call CMS_LoginLogout(L)
    app.StatusBar = "Arrumando Horários"
    formulas
    
ABA = Sheets("INICIO").Cells(L, 1).Value
    app.StatusBar = "Dimensionando Equipes " & ABA
    Call ARRUMA_CAPA(ABA)
    

Range("A:A").Select
Selection.Clear

End If

L = L + 1
    
Loop
    'app.StatusBar = "Exportando dados para o Relatório"
    
    
    'Range("H8").Value = Format(f_Lanche, "hh:mm:ss")
    'Range("I8").Value = Format(f_Desc, "hh:mm:ss")
    
    'Range("H9").Value = qunt_Lanche
    'Range("I9").Value = qunt_Desc
    
    Range("A1").Select
    
    Sheets("INICIO").Select
MsgBox "Relatório Concluído", , "Planejamento"
    
    
    app.StatusBar = ""
    app.ScreenUpdating = True
    
End Sub

Sub CMS(L)
    
    Dim periodo As String
    Dim flg As Boolean
    Dim b As Long
    Dim I As Long
    Dim info As Variant
    Dim rep As Variant
    Dim Caminho  As String
    Dim caminhop  As String
    Dim hora, Data As String
    
    '------------------------------------------
    Sheets("CMS").Select
    Cells.Select
    Selection.ClearContents
    '------------------------------------------
    PLANILHA = Application.ActiveWorkbook.Name
    '------------------------------------------
    Caminho = Application.ActiveWorkbook.Path
    '----------------------------------------
    CentreVu = Sheets("INICIO").Cells(L, 2).Value
    ARQUIVO = Caminho & "\" & "Base CMS.xls"
    '------------------------------------------
    skills = Sheets("INICIO").Cells(L, 3)
    Data = Sheets("INICIO").Cells(2, 2)
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
    Set info = acsSrv.Reports.Reports("Historical\Designer\Desempenho da Equipe/Agente MIS(INTERVALO) - Planejamento")
    
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
            rep.SetProperty "Grupo/Especialidade", skills
            rep.SetProperty "Data", Data
            rep.SetProperty "DAC", "1"
            rep.SetProperty "Horário:", "00:00-23:30"
            
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
    
    Workbooks.OpenText Filename:=ARQUIVO, Origin _
    :=xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
    xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
    Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
    Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
    Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15 _
    , 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array(29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array(42, 1), Array(43, 1), Array(44, 1), Array(45, 1), Array(46, 1), Array(47, 1), Array(48, 1), Array(49, 1)), TrailingMinusNumbers:=True
    
    
    fechar = Application.ActiveWorkbook.Name
    Cells.Select
    Selection.Copy
    
' cola no arquivo
    Windows(PLANILHA).Activate
    Sheets("CMS").Visible = True
    Sheets("CMS").Select
    Cells.Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("Base CMS.xls").Close
    Selection.Replace What:=",000000000", Replacement:="0", LookAt:=xlPart _
    , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    
    If CentreVu <> "10.189.0.91" Then
    Rows(4).Delete Shift:=xlShiftUp
    
    Else
    Rows(3).Delete Shift:=xlShiftUp
    Rows(4).Delete Shift:=xlShiftUp
    
    End If
        '''''''''''''''''''''''''''''''''''''''''' A
    Range("AE4").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[-30],5)*1"
    Selection.AutoFill Destination:=Range(Selection, ActiveCell.Cells(Range("A4").End(xlDown).Row - 3, 1))
    
    Range("AE4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    Range("AE4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Range("AE3").Select
    ActiveCell.FormulaR1C1 = "operador"
    Range("AF3").Select
    ActiveCell.FormulaR1C1 = "supervisor"
    ''''''''''''''''''''''''''''''''''''''''''''
        
    If Sheets("INICIO").Cells(L, 4).Value = "" Then
    Range("AE4").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-30],BD!C[-28]:C[-27],2,0)"
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-30],BD!C[-28]:C[-27],2,0),""LOGIN NÃO CADASTRADO NO WFM"")"
    
    Range("AF4").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-31],BD!C[-29]:C[-26],4,0)"
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-31],BD!C[-29]:C[-26],4,0),""SEM SUPERVISOR NO WFM"")"
    Else
    
    If Sheets("INICIO").Cells(L, 4).Value = "SBA|" Then
    Range("AE4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(""SBA|""&RC[-30],BD!C[-28]:C[-27],2,0),""LOGIN NÃO CADASTRADO NO WFM"")"
        
    Range("AF4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(""SBA|""&RC[-31],BD!C[-29]:C[-26],4,0),""SEM SUPERVISOR NO WFM"")"
        
    Else
    If Sheets("INICIO").Cells(L, 4).Value = "SBC|" Then
    Range("AE4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(""SBC|""&RC[-30],BD!C[-28]:C[-27],2,0),""LOGIN NÃO CADASTRADO NO WFM"")"
        
    Range("AF4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(""SBC|""&RC[-31],BD!C[-29]:C[-26],4,0),""SEM SUPERVISOR NO WFM"")"
        
    End If
    End If
    End If
    
    
    Range("AE4:AF4").Select
    Selection.AutoFill Destination:=Range(Selection, ActiveCell.Cells(Range("A4").End(xlDown).Row - 3, 1))
    
    Range("AE4:AF6000").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Application.CutCopyMode = False
    
    Range(Selection, Selection.End(xlToLeft)).Select
    ActiveWorkbook.Worksheets("CMS").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("CMS").Sort.SortFields.Add Key:=Range("AF4:AF6000") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("CMS").Sort.SortFields.Add Key:=Range("AE4:AE6000") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("CMS").Sort
        .SetRange Range("A4:AF6000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
End Sub

Sub CMS_LoginLogout(L)

    Dim flg As Boolean
    Dim b As Long
    Dim I As Long
    Dim info As Variant
    Dim rep As Variant
    Dim Caminho  As String
    Dim caminhop  As String
    Dim hora, skills, Data As String
    
    '------------------------------------------
    Sheets("CMS").Select
    'Cells.Select
    'Selection.ClearContents
    '------------------------------------------
    skills = Sheets("INICIO").Cells(L, 3)
    '------------------------------------------
    PLANILHA = Application.ActiveWorkbook.Name
    '------------------------------------------
    Caminho = Application.ActiveWorkbook.Path
    '----------------------------------------
    CentreVu = Sheets("INICIO").Cells(L, 2).Value
    ARQUIVO = Caminho & "\" & "Base CMS.xls"
    '------------------------------------------
    Data = Sheets("INICIO").Cells(2, 2)
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
    
    If CentreVu = "10.4.0.90" Then
    Set info = acsSrv.Reports.Reports("Historical\Designer\Login/Logout (Especialidade) [GRUPO DIVERSOS]")
    
    Else
    
    Set info = acsSrv.Reports.Reports("Historical/Designer/Login/Logout (ESPECIALIDADE) [GRUPO DIVERSOS]")
    End If
    
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
    Workbooks.OpenText Filename:=ARQUIVO, Origin _
    :=xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
    xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
    Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
    Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
    Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15 _
    , 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array(29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array(42, 1), Array(43, 1), Array(44, 1), Array(45, 1), Array(46, 1), Array(47, 1), Array(48, 1), Array(49, 1)), TrailingMinusNumbers:=True
    

    fechar = Application.ActiveWorkbook.Name
    
    'If (CentreVu = "10.189.0.90" Or CentreVu = "10.2.0.90") Then
          
    Range("A:A,D:D,F:F").Delete Shift:=xlShiftLeft
    Range("A:E").Copy
        
    'Else
        
    'Range("C:C,E:E").Delete Shift:=xlShiftLeft
    'Range("A:E").Copy
    
    'End If
' cola no arquivo
    Windows(PLANILHA).Activate
    Sheets("CMS").Visible = True
    Sheets("CMS").Select
    Range("AR1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("Base CMS.xls").Close SaveChanges:=False
    
    
    Range("AQ4").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[1],5)*1"
    
    Range("AQ4").Select
    Selection.AutoFill Destination:=Range(Selection, ActiveCell.Cells(Range("AR4").End(xlDown).Row - 3, 1))
    
    Range("AQ4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("AR4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.NumberFormat = "General"
    
    Range("AQ4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
  
End Sub

Sub formulas()
Attribute formulas.VB_ProcData.VB_Invoke_Func = " \n14"
   
Dim excluir

'' utilizar para formula tempo

    Range("C3:AB3").Select
    Selection.Copy
    
    Range("BA3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("BA4").Select
    ActiveCell.FormulaR1C1 = "=RC[-50]/(24*60*60)"
    Selection.AutoFill Destination:=Range("BA4:BZ4"), Type:=xlFillDefault
    Range("BA4:BZ4").Select
    Selection.AutoFill Destination:=Range(Selection, ActiveCell.Cells(Range("A4").End(xlDown).Row - 3, 1))
    
    Range("BA3:BZ3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "[h]:mm:ss"

    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Selection.Cut
    Range("C3").Select
    ActiveSheet.Paste
    
    'formula tempo deslog
    Range("AW4").Select
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C[-5]<>RC[-5],0,RC[-3]-R[-1]C[-2])"
    Range("AW4").Select
    Selection.AutoFill Destination:=Range(Selection, ActiveCell.Cells(Range("AR4").End(xlDown).Row - 3, 1))
    
    Range("AW4").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
'arrumando formulas
    Range("AG3").Select
    ActiveCell.FormulaR1C1 = "qnt deslog"
    Range("AH3").Select
    ActiveCell.FormulaR1C1 = "temp deslog"
    Range("AI3").Select
    ActiveCell.FormulaR1C1 = "login"
    Range("AJ3").Select
    ActiveCell.FormulaR1C1 = "logout"
    
    Range("AG4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((COUNTIF(C[11],RC[-32])-1)<0,0,COUNTIF(C[11],RC[-32])-1)"
    
    Range("AH4").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(C[10],RC[-33],C[15])"
    
    Range("AI4").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-34],C[9]:C[11],3,0)"
    
    Range("AJ4").Select
    ActiveCell.FormulaR1C1 = "=INDEX(C[8]:C[11],MATCH(RC[-35],C[8],0)+RC[-3],4)"
    

    Range("AG4:AJ4").Select
    Selection.AutoFill Destination:=Range(Selection, ActiveCell.Cells(Range("A4").End(xlDown).Row - 3, 1))
    
    Range("AG4:AJ4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("AH:AJ").Select
    Selection.NumberFormat = "[h]:mm:ss"
    

excluir = 4

Do While Sheets("CMS").Cells(excluir, 1).Value <> ""

    If Sheets("CMS").Cells(excluir, 32).Value = "SEM SUPERVISOR NO WFM" Then
        
    Rows(excluir).Delete
    
    Else
    
    excluir = excluir + 1
    
    End If

Loop
    
    
    
End Sub

Sub ARRUMA_CAPA(ABA)

Dim LA, Lcms
Dim sup
Dim T_dsl, T_pausa, T_sist, T_lan, T_dec, T_ambu, T_def, T_bko, T_trei, T_feed, T_part, T_reun

Sheets(ABA).Select
Sheets(ABA).Range("11:65000").Select
'Selection.Clear
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete Shift:=xlUp

T_dsl = 0
T_pausa = 0
T_sist = 0
T_lan = 0
T_dec = 0
T_ambu = 0
T_def = 0
T_bko = 0
T_trei = 0
T_feed = 0
T_part = 0
T_reun = 0


LA = 11
Lcms = 4

        sup = Sheets("CMS").Cells(Lcms, 32).Value
        Sheets(ABA).Range(Cells(LA, 2), Cells(LA, 20)).Select
        Selection.Merge
        With ActiveCell
              .Value = "Supervisor(a) " & sup
              .HorizontalAlignment = xlCenter
              .VerticalAlignment = xlCenter
              With .Font
                    .FontStyle = "Negrito"
                    .Size = 14
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = -0.249977111117893
                .PatternTintAndShade = 0
            End With
            
            End With
            ActiveCell.Offset(1, 0).Select
            With Sheets("BD")
                .Select
                .Range("L1:AD1").Copy
            End With
            Sheets(ABA).Select
            ActiveCell.PasteSpecial

LA = LA + 2

Do While Sheets("CMS").Cells(Lcms, 32).Value <> vbNullString
    
Sheets("BD").Range("L2").Value = Sheets("CMS").Cells(Lcms, 1).Value
Calculate
If Sheets("BD").Range("N2").Value = "" Then

    Lcms = Lcms + 1

Else

     
    If sup <> Sheets("CMS").Cells(Lcms, 32).Value Then
        
        Sheets("BD").Select
        Range("L3:AD3").Copy
        
        Sheets(ABA).Select
        Cells(LA, 2).Select
        Cells(LA, 2).PasteSpecial
                
        Cells(LA, 7).Value = T_dsl
        Cells(LA, 9).Value = T_pausa
        Cells(LA, 10).Value = T_sist
        Cells(LA, 11).Value = T_lan
        Cells(LA, 12).Value = T_dec
        Cells(LA, 13).Value = T_ambu
        Cells(LA, 14).Value = T_def
        Cells(LA, 15).Value = T_bko
        Cells(LA, 16).Value = T_trei
        Cells(LA, 17).Value = T_feed
        Cells(LA, 18).Value = T_part
        Cells(LA, 19).Value = T_reun
            
With Range(Cells(LA, 7), Cells(LA, 20))
        .Select
        .NumberFormat = "[h]:mm:ss"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
End With
        
        With Selection.Font
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
        .Bold = True
        End With

        LA = LA + 2
        sup = Sheets("CMS").Cells(Lcms, 32).Value
        
        Sheets(ABA).Range(Cells(LA, 2), Cells(LA, 20)).Select
        Selection.Merge
        With ActiveCell
              .Value = "Supervisor(a) " & sup
              .HorizontalAlignment = xlCenter
              .VerticalAlignment = xlCenter
              With .Font
                    .FontStyle = "Negrito"
                    .Size = 14
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
            
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = -0.249977111117893
                .PatternTintAndShade = 0
            End With
                

            End With
            ActiveCell.Offset(1, 0).Select
            With Sheets("BD")
                .Select
                .Range("L1:AD1").Copy
            End With
            Sheets(ABA).Select
            ActiveCell.PasteSpecial
            
            LA = LA + 2
            
T_dsl = 0
T_pausa = 0
T_sist = 0
T_lan = 0
T_dec = 0
T_ambu = 0
T_def = 0
T_bko = 0
T_trei = 0
T_feed = 0
T_part = 0
T_reun = 0
            
            
End If


    Sheets(ABA).Range(Cells(LA, 2), Cells(LA, 20)).Value = Sheets("BD").Range("L2:AD2").Value
    
            ' formata os valores para horaa
            Range(Cells(LA, 4).Address, Cells(LA, 5).Address).NumberFormat = "hh:mm:ss"
            Range(Cells(LA, 7).Address, Cells(LA, 20).Address).NumberFormat = "hh:mm:ss"
        ' formata o valores geral
            Cells(LA, 6).NumberFormat = "General"
            Range(LA & ":" & LA).Select
            
            
        ' centraliza o login e as pausas
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
        ' coloca a borda nas linhas
            With Range(ActiveCell.Address, ActiveCell.Cells(1, 20).Address).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = -0.149937437055574
                .Weight = xlThin
            End With
            
Call Pausas(LA, ABA)
            
T_dsl = T_dsl + Cells(LA, 7).Value
T_pausa = T_pausa + Cells(LA, 9).Value
T_sist = T_sist + Cells(LA, 10).Value
T_lan = T_lan + Cells(LA, 11).Value
T_dec = T_dec + Cells(LA, 12).Value
T_ambu = T_ambu + Cells(LA, 13).Value
T_def = T_def + Cells(LA, 14).Value
T_bko = T_bko + Cells(LA, 15).Value
T_trei = T_trei + Cells(LA, 16).Value
T_feed = T_feed + Cells(LA, 17).Value
T_part = T_part + Cells(LA, 18).Value
T_reun = T_reun + Cells(LA, 19).Value
            
            
LA = LA + 1
Lcms = Lcms + 1

'Else


'Lcms = Lcms + 1

End If

Loop

        Sheets("BD").Select
        Range("L3:AD3").Copy
        
        Sheets(ABA).Select
        Cells(LA, 2).Select
        Cells(LA, 2).PasteSpecial
                
        Cells(LA, 7).Value = T_dsl
        Cells(LA, 9).Value = T_pausa
        Cells(LA, 10).Value = T_sist
        Cells(LA, 11).Value = T_lan
        Cells(LA, 12).Value = T_dec
        Cells(LA, 13).Value = T_ambu
        Cells(LA, 14).Value = T_def
        Cells(LA, 15).Value = T_bko
        Cells(LA, 16).Value = T_trei
        Cells(LA, 17).Value = T_feed
        Cells(LA, 18).Value = T_part
        Cells(LA, 19).Value = T_reun
        
        With Range(Cells(LA, 7), Cells(LA, 20))
        .Select
        .NumberFormat = "[h]:mm:ss"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
End With
        
        With Selection.Font
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
        .Bold = True
        End With


End Sub

Sub Pausas(LA, ABA)

If Sheets(ABA).Cells(LA, 11).Value >= Sheets("BD").Range("N11").Value Then
    
    Sheets("BD").Select
    Sheets("BD").Range("N9").Select
    Selection.Copy
    Sheets(ABA).Select
    Sheets(ABA).Cells(LA, 11).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
Else

If (Sheets(ABA).Cells(LA, 11).Value >= Sheets("BD").Range("M11").Value And Sheets(ABA).Cells(LA, 11).Value <= Sheets("BD").Range("N11").Value) Then
    
    Sheets("BD").Select
    Sheets("BD").Range("M9").Select
    Selection.Copy
    Sheets(ABA).Select
    Sheets(ABA).Cells(LA, 11).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    
    
End If
End If

If Sheets(ABA).Cells(LA, 12).Value >= Sheets("BD").Range("N12").Value Then
    
    Sheets("BD").Select
    Sheets("BD").Range("N9").Select
    Selection.Copy
    Sheets(ABA).Select
    Sheets(ABA).Cells(LA, 12).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    

Else

If (Sheets(ABA).Cells(LA, 12).Value >= Sheets("BD").Range("M12").Value And Sheets(ABA).Cells(LA, 12).Value <= Sheets("BD").Range("N12").Value) Then
    
    Sheets("BD").Select
    Sheets("BD").Range("M9").Select
    Selection.Copy
    Sheets(ABA).Select
    Sheets(ABA).Cells(LA, 12).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    

End If
End If

If Sheets(ABA).Cells(LA, 8).Value < Sheets("BD").Range("N20").Value Then
    
    Sheets("BD").Select
    Sheets("BD").Range("N9").Select
    Selection.Copy
    Sheets(ABA).Select
    Sheets(ABA).Cells(LA, 8).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False


Else

If (Sheets(ABA).Cells(LA, 8).Value > Sheets("BD").Range("M20").Value) Then
    
    Sheets("BD").Select
    Sheets("BD").Range("M9").Select
    Selection.Copy
    Sheets(ABA).Select
    Sheets(ABA).Cells(LA, 8).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    


End If
End If


End Sub



