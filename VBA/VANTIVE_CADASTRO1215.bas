Attribute VB_Name = "VANTIVE_CADASTRO"
Sub Atualizar_hora_CADASTRO()
Dim ARQUIVO, data, PID_VANTIVE As String
Dim I


    ARQUIVO = Sheets("PREMISSAS").Cells(20, 2).Value


            For I = 98 To 114
    
    
            If Cells(I, 1).Value = "" Or Cells(I, 1).Value <> 1 Then
    
    I = I - 1
            
            If Cells(I, 1).Value = 1 Then
                GoTo continuar
            End If
            End
            End If
            Next

continuar:

    data = Cells(I, 3).Value
    PID_VANTIVE = "Aplicativo de vendas - SPEEDY"
    Call CONS_CADASTRO(PID_VANTIVE, data, ARQUIVO)
    
End Sub



Sub CONS_CADASTRO(PID_VANTIVE, data, ARQUIVO)
    
    
    'ABRINDO RELATORIO ATENDIMENTO RECEPITIVO
    
    AppActivate (PID_VANTIVE)
    Application.Wait Now + TimeValue("00:00:05")
    Application.SendKeys ("^{F4}")
    Application.SendKeys ("%{F}")
    Application.SendKeys ("{O}")
    Application.SendKeys ("{A 2}")
    Application.SendKeys ("~")
    Application.Wait Now + TimeValue("00:00:03")
    
    'ATENDIDO EM:
    
    Application.SendKeys ("{TAB 15}")
    Application.SendKeys ("B")
    Application.SendKeys ("{TAB}")
    Application.SendKeys (data)
    Application.Wait Now + TimeValue("00:00:04")
    

    'INFO STATUS
    Application.SendKeys ("{TAB 6}")
    Application.SendKeys ("ATIVO NORMAL")
    Application.SendKeys ("~")
    Application.Wait Now + TimeValue("00:00:04")
    


    'SALVANDO
        
    Application.SendKeys ("%{F}")
    Application.SendKeys ("{UP 5}")
    Application.SendKeys ("~")
    Application.SendKeys ("{TAB 6}")
    Application.SendKeys ("~")
    Application.Wait Now + TimeValue("00:00:04")
    Application.SendKeys ("{BACKSPACE 255}")
    Application.Wait Now + TimeValue("00:00:04")
    Application.SendKeys (ARQUIVO)
    Application.SendKeys ("~")
    Application.Wait Now + TimeValue("00:00:04")
    Application.SendKeys ("{TAB 6}")
    Application.Wait Now + TimeValue("00:00:02")
    Application.SendKeys ("~")
    Application.SendKeys ("~")
    Application.Wait Now + TimeValue("00:00:01")
    
    
    
End Sub
Sub CADASTRO()

Application.ScreenUpdating = False
Application.StatusBar = "ATUALIZANDO CADASTRO"
Application.DisplayAlerts = False





Dim ORIGEM, DESTINO
Dim L, CRITERIO_1, CRITERIO_2, VARIAVEL

Sheets("CADASTRO").Visible = True

ORIGEM = ActiveWorkbook.Name


Sheets("CADASTRO").Select

Cells.Select
Selection.ClearContents
    
Application.StatusBar = "----> Atualizando Cadastro <---"
Workbooks.OpenText Filename:= _
        "\\brsjcsrv01\Operacoes\Speedy\Planejamento\_ComumSpeedy\Estudos Particular\Tempo Real\Teste\CADASTRO_121520.txt", Origin:=xlWindows, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=True, Comma:=True, _
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
    Sheets("CADASTRO").Select
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Windows(DESTINO).Activate
        

    ActiveWindow.Close

   Sheets("CADASTRO").Select
Sheets("CADASTRO").Visible = False
    Sheets("CAPA").Select

Application.CutCopyMode = False
Application.StatusBar = "ATUALIZADO"
Sheets("CAPA").Select


End Sub
