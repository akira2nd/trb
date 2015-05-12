Attribute VB_Name = "Módulo1"
Sub duplicadas()
Attribute duplicadas.VB_ProcData.VB_Invoke_Func = " \n14"

Dim ARQUIVO


Application.ScreenUpdating = False
Application.DisplayAlerts = False

    Sheets("PORTAL").Select
    
If Sheets("PORTAL").Cells(2, 1).Value <> "" Then

    Columns("BN:BO").Select
    Selection.ClearContents
    
    Range("BN1").Value = "DUPLICADOS"
    Range("BO1").Value = "HORA"
    
    Range("BN2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-8],RC[-8])"
    Range("BO2").Select
    ActiveCell.FormulaR1C1 = "=TIMEVALUE(HOUR(RC[-35])&"":""&MINUTE(RC[-35]))"
    
    Range("BN2:BO2").Select
        Selection.AutoFill Destination:=Range(Selection, ActiveCell.Cells(Range("A2").End(xlDown).Row - 1, 1))
    
    Calculate
    Columns("BN:BO").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    

    
    Sheets("RESUMO").Select
    Calculate

ARQUIVO = Sheets("INICIO").Cells(10, 2).Value
    
    Sheets(Array("RESUMO", "PA")).Select
    Sheets(Array("RESUMO", "PA")).Copy Before:=Sheets(1)
    
    Sheets("RESUMO (2)").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Application.CutCopyMode = False
    
    Sheets("PA (2)").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Application.CutCopyMode = False
    
    Sheets(Array("RESUMO (2)", "PA (2)")).Select
    Sheets(Array("RESUMO (2)", "PA (2)")).Move
    
    Sheets("PA (2)").Select
    Sheets("PA (2)").Name = "PA"
    
    Sheets("RESUMO (2)").Select
    Sheets("RESUMO (2)").Name = "RESUMO"
    
    'Sheets(Array("RESUMO", "PA")).Select
    'Sheets(Array("RESUMO", "PA")).Copy
    'Cells.Select
    'Selection.Copy
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Application.CutCopyMode = False
    
    'Range("B1").Select

    ActiveWorkbook.SaveAs Filename:=ARQUIVO
    
    ActiveWindow.Close

Application.ScreenUpdating = True
Application.DisplayAlerts = True

Sheets("INICIO").Select

MsgBox "OK"

Else

MsgBox "sem conteúdo no portal"

End If



End Sub
Sub VDS()

Application.ScreenUpdating = False
Application.StatusBar = "ATUALIZANDO VENDAS"
Application.DisplayAlerts = False


Dim ORIGEM, DESTINO

Sheets("VANTIVE").Visible = True

ORIGEM = ActiveWorkbook.Name

Sheets("VANTIVE").Select

Cells.Select
Selection.ClearContents
Selection.ClearContents
Selection.ClearContents

    
Application.StatusBar = "----> Atualizando Vendas <---"
Workbooks.OpenText Filename:= _
        "\\brsjcsrv01\Operacoes\Speedy\Planejamento\_ComumSpeedy\Estudos Particular\Tempo Real\Teste\Vendas_Ativo.txt", Origin:=xlWindows, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=True, _
        Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), Array( _
        3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10 _
        , 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), _
        Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array( _
        23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array(29, 1), _
        Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), Array( _
        36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array(42, 1), _
        Array(43, 1), Array(44, 1), Array(45, 1), Array(46, 1), Array(47, 1), Array(48, 1), Array( _
        49, 1), Array(50, 1), Array(51, 1), Array(52, 1), Array(53, 1), Array(54, 1), Array(55, 1), _
        Array(56, 1), Array(57, 1), Array(58, 1), Array(59, 1), Array(60, 1), Array(61, 1), Array( _
        62, 1), Array(63, 1), Array(64, 1), Array(65, 1), Array(66, 1), Array(67, 1), Array(68, 1), _
        Array(69, 1), Array(70, 1), Array(71, 1), Array(72, 1), Array(73, 1), Array(74, 1), Array( _
        75, 1), Array(76, 1), Array(77, 1), Array(78, 1), Array(79, 1), Array(80, 1)), _
        TrailingMinusNumbers:=True
        
        DESTINO = ActiveWorkbook.Name
        
    Cells.Select
    
    Selection.Copy
    Windows(ORIGEM).Activate
    Sheets("VANTIVE").Select
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Windows(DESTINO).Activate
    
    
    'ActiveWorkbook.SaveAs Filename:=COPIA, WriteResPassword:="PLAN01", ReadOnlyRecommended:=False
        'ActiveWorkbook.SaveAs Filename:= _
        COPIA, _
        FileFormat:=xlExcel8, Password:="", WriteResPassword:="PLAN01", _
        ReadOnlyRecommended:=False, CreateBackup:=False

    ActiveWindow.Close
    
    
    
    Sheets("VANTIVE").Select
    
    Range("BZ1").Value = "HORA"
    
    Range("BZ2").Select
    ActiveCell.FormulaR1C1 = _
        "=TIMEVALUE(HOUR(RC[-36])&"":""&MINUTE(RC[-36])&"":""&SECOND(RC[-36]))"

    Selection.AutoFill Destination:=Range(Selection, ActiveCell.Cells(Range("A2").End(xlDown).Row - 1, 1))

    Calculate
    Columns("BZ:BZ").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
        
Sheets("VANTIVE").Visible = False
  
Application.StatusBar = "CMS ATUALIZADO"

VDS_portal

End Sub

Sub VDS_portal()

Application.ScreenUpdating = False
Application.StatusBar = "CMS VENDAS"
Application.DisplayAlerts = False


Dim ORIGEM, DESTINO

Sheets("PORTAL").Visible = True

ORIGEM = ActiveWorkbook.Name

Sheets("PORTAL").Select

Cells.Select
Selection.ClearContents
Selection.ClearContents
Selection.ClearContents

    
Application.StatusBar = "----> Atualizando Vendas <---"
Workbooks.OpenText Filename:= _
        "\\brsjcsrv01\Operacoes\Speedy\Planejamento\_ComumSpeedy\Estudos Particular\Tempo Real\Teste\CADASTRO_ATIVO.xls", Origin:=xlWindows, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=True, _
        Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), Array( _
        3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10 _
        , 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), _
        Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array( _
        23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array(29, 1), _
        Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), Array( _
        36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array(42, 1), _
        Array(43, 1), Array(44, 1), Array(45, 1), Array(46, 1), Array(47, 1), Array(48, 1), Array( _
        49, 1), Array(50, 1), Array(51, 1), Array(52, 1), Array(53, 1), Array(54, 1), Array(55, 1), _
        Array(56, 1), Array(57, 1), Array(58, 1), Array(59, 1), Array(60, 1), Array(61, 1), Array( _
        62, 1), Array(63, 1), Array(64, 1), Array(65, 1), Array(66, 1), Array(67, 1), Array(68, 1), _
        Array(69, 1), Array(70, 1), Array(71, 1), Array(72, 1), Array(73, 1), Array(74, 1), Array( _
        75, 1), Array(76, 1), Array(77, 1), Array(78, 1), Array(79, 1), Array(80, 1)), _
        TrailingMinusNumbers:=True
        
        DESTINO = ActiveWorkbook.Name
        
    
    Rows("5:5").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    Selection.Copy
    Windows(ORIGEM).Activate
    Sheets("PORTAL").Select
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Windows(DESTINO).Activate
    
    
    'ActiveWorkbook.SaveAs Filename:=COPIA, WriteResPassword:="PLAN01", ReadOnlyRecommended:=False
        'ActiveWorkbook.SaveAs Filename:= _
        COPIA, _
        FileFormat:=xlExcel8, Password:="", WriteResPassword:="PLAN01", _
        ReadOnlyRecommended:=False, CreateBackup:=False

    ActiveWindow.Close
    
    
    
    Sheets("PORTAL").Select
    
  
 'Application.StatusBar = "VENDAS ATUALIZADO"

duplicadas

End Sub

