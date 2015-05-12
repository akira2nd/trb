Attribute VB_Name = "Módulo2"
Sub Pega_Vantive()

Application.ScreenUpdating = False

    Dim plan_atual As Workbook, plan_extracao As Workbook
    
    Set plan_atual = ThisWorkbook
    
'    Sheets("SISCAM").Activate
'    Range("K:V").Select
'    Selection.ClearContents
    
    Sheets("PORTAL").Activate
    Cells.Select
    Selection.ClearContents
    
    Sheets("TOTAL").Activate
    Cells.Select
    Selection.ClearContents
    
        Workbooks.Open Filename:= _
        Sheets("INICIO").Cells(11, 3).Value _
        , UpdateLinks:=False, ReadOnly:=True, IgnoreReadOnlyRecommended:=True
        
        
    Set plan_extracao = ActiveWorkbook
    
    Sheets("Total").Select
    Range("BF:BF,BI:BI").Select
    Selection.Copy
    
    plan_atual.Activate
    
    Sheets("TOTAL").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Application.CutCopyMode = False
    plan_extracao.Close SaveChanges:=False
    
    
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

    
    Set plan_extracao = ActiveWorkbook
    
    Rows("5:5").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    Selection.Copy
    plan_atual.Activate
    Sheets("PORTAL").Select
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'    plan_extracao.Activate
'    Range("5:5").AutoFilter Field:=1, Criteria1:="TVA SJC"
'    Range("Y:Y,AA:AA").Copy
'
'    plan_atual.Activate
'    Sheets("SISCAM").Activate
'    Range("K1").PasteSpecial xlPasteValuesAndNumberFormats
'    Application.CutCopyMode = False
'
'    plan_extracao.Activate
'    Range("5:5").AutoFilter Field:=1, Criteria1:="TVA STO"
'    Range("Y:Y,AA:AA").Copy
'
'    plan_atual.Activate
'    Sheets("SISCAM").Activate
'    Range("R1").PasteSpecial xlPasteValuesAndNumberFormats
'    Application.CutCopyMode = False
'
'    plan_extracao.Close SaveChanges:=False
'
'    With Range("M6", Range("L6").End(xlDown).Offset(0, 1).Address)
'        .FormulaR1C1 = "=HOUR(RC[-2])"
'        .Copy
'        .PasteSpecial xlPasteValuesAndNumberFormats
'    End With
'
'    With Range("T6", Range("S6").End(xlDown).Offset(0, 1).Address)
'        .FormulaR1C1 = "=HOUR(RC[-2])"
'        .Copy
'        .PasteSpecial xlPasteValuesAndNumberFormats
'    End With
'    Application.CutCopyMode = False
    
    Application.CutCopyMode = False
    plan_extracao.Close SaveChanges:=False
        
        Sheets("PORTAL").Select
    
        Range("AS1").Value = "HORA"
        Range("AS2").Select
        ActiveCell.FormulaR1C1 = "=TEXT(RC[-19],""HH"")*1"
        
        Range("AT1").Select
        ActiveCell.FormulaR1C1 = "OS"
        Range("AT2").Select
        ActiveCell.FormulaR1C1 = "=IFERROR(LEN(RC[-5]),"""")"
        
            Range("AU1").Value = "duplicados SPD"
            Range("AU2").Select
            ActiveCell.FormulaR1C1 = "=COUNTIF(C[-7],RC[-7])"
            
            Range("AV1").Value = "duplicados SIA"
            Range("AV2").Select
            ActiveCell.FormulaR1C1 = "=COUNTIF(C[-7],RC[-7])"
            
        Range("AW1").Value = "TOTAL SPD"
        Range("AW2").Select
        ActiveCell.FormulaR1C1 = _
        "=COUNTIF(TOTAL!C1,PORTAL!RC[-9])"
        
        Range("AX1").Value = "TOTAL SIA"
        Range("AX2").Select
        ActiveCell.FormulaR1C1 = _
        "=COUNTIF(TOTAL!C2,PORTAL!RC[-9])"
        
        Range("AY1").Value = "DUP_ID"
        Range("AY2").Select
        ActiveCell.FormulaR1C1 = _
        "=COUNTIFS(C6,RC6,C28,""PRÉ VENDA"",C36,""*SPEEDY*"")"
        
        Range("AZ1").Value = "AUDITADAS"
        Range("AZ2").Select
        ActiveCell.FormulaR1C1 = _
        "=COUNTIFS(C[-46],RC[-46],C[-30],""<>""&RC[-30])"
        
        Range("AS2:AZ2").Select
                
        Selection.AutoFill Destination:=Range(Selection, ActiveCell.Cells(Range("A2").End(xlDown).Row - 1, 1))
        Columns("AS:AZ").Select
        Selection.NumberFormat = "General"
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False

Application.ScreenUpdating = True
Application.StatusBar = "Portal OK"

End Sub

Sub GRAVAR()

Dim grava, HResult
Dim L, L1

With Application
.Application.DisplayAlerts = False
.ScreenUpdating = False
.EnableEvents = False
End With

    
grava = Sheets("INICIO").Cells(10, 3).Value

    Sheets(Array("RESUMO", "OCUPACAO")).Copy
    
    Sheets(2).Select
    With Cells
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats
    End With
    
    Sheets(1).Select
    With Cells
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats
    End With
   
    Application.DisplayAlerts = False
    
    With ActiveWorkbook
        .SaveAs grava
        .Close
    End With
    
    
HResult = Sheets("INICIO").Range("C4").Value

For L = 19 To 36
    
    Sheets("ARRUMAR").Select
    
    If Sheets("ARRUMAR").Cells(L, 1).Value = HResult Then
    
        Range(Cells(L, 2), Cells(L, 8)).Value = Range("B13:H13").Value
        
    End If
    
Next

For L1 = 39 To 56
    
    Sheets("ARRUMAR").Select
    
    If Sheets("ARRUMAR").Cells(L1, 1).Value = HResult Then
    
        Range(Cells(L1, 2), Cells(L1, 8)).Value = Range("B16:H16").Value
        
    End If
    
Next
    
    
    
With Application
.Application.DisplayAlerts = True
.ScreenUpdating = True
.EnableEvents = True
End With
    
    
Application.StatusBar = "GRAVADO"

End Sub

Sub envia_planilha()
   
Dim DE, para, CC, Cco, Titulo, introducao
Dim Rel

If MsgBox("Deseja enviar o email?", vbYesNo, "Planejamento") = vbYes Then

With Application
.Application.DisplayAlerts = False
.ScreenUpdating = False
.EnableEvents = False
End With


Rel = Sheets("ARRUMAR").Range("I8").Value
DE = Sheets("ARRUMAR").Range("I3").Text
para = Sheets("ARRUMAR").Range("I4").Text
CC = Sheets("ARRUMAR").Range("I5").Text
Cco = Sheets("ARRUMAR").Range("I6").Text
Titulo = Sheets("ARRUMAR").Range("I7").Text

ORIGEM = ActiveWorkbook.Name

        Workbooks.Open Filename:=Rel, _
        UpdateLinks:=False, ReadOnly:=True, IgnoreReadOnlyRecommended:=True
        
DESTINO = ActiveWorkbook.Name

Windows(DESTINO).Activate
   ' Seleciona a planilha
   Sheets("Resumo").Activate
   ActiveSheet.Cells.Select
   
   ' mostrar o envelope na planilha ativa
   ActiveWorkbook.EnvelopeVisible = True
   
   ' atribui valor aos campos
   ' to = para **cc = CC **BCC = Cco **subject = título
   ' Attachments = Anexo
   With Application.ActiveSheet.MailEnvelope
      .Introduction = ""
      .Item.SentOnBehalfOfName = DE
      .Item.To = para
      .Item.CC = CC
      .Item.BCC = Cco
      .Item.Subject = Titulo
      .Item.Attachments.Add (Rel)
      .Item.Send
   End With
   
ActiveWorkbook.EnvelopeVisible = False
   
Windows(DESTINO).Close
'Activate.Window.Close savechanges:=False
   
With Application
.Application.DisplayAlerts = True
.ScreenUpdating = True
.EnableEvents = True
End With
   
MsgBox "Desempenho Enviado!"

Application.StatusBar = "Email enviado!"

End If

End Sub

Sub envia_planilha_2()
   
Dim DE, para, CC, Cco, Titulo, introducao
Dim Rel

If MsgBox("Deseja Enviar o e-mail?", vbYesNo, "Planejamento") = vbYes Then

With Application
.Application.DisplayAlerts = False
.ScreenUpdating = False
.EnableEvents = False
End With

Rel = Sheets("ARRUMAR").Range("L8").Value
DE = Sheets("ARRUMAR").Range("L3").Text
para = Sheets("ARRUMAR").Range("L4").Text
CC = Sheets("ARRUMAR").Range("L5").Text
Cco = Sheets("ARRUMAR").Range("L6").Text
Titulo = Sheets("ARRUMAR").Range("L7").Text

'ORIGEM = ActiveWorkbook.Name
'
'        Workbooks.Open Filename:=Rel, _
'        UpdateLinks:=False, ReadOnly:=True, IgnoreReadOnlyRecommended:=True
'
'DESTINO = ActiveWorkbook.Name

'Windows(DESTINO).Activate
   '' Seleciona a planilha
   Sheets("Resultado").Activate
   ActiveSheet.Cells.Select
   
   '' mostrar o envelope na planilha ativa
   ActiveWorkbook.EnvelopeVisible = True
   
   ' atribui valor aos campos
   ' to = para **cc = CC **BCC = Cco **subject = título
   ' Attachments = Anexo
   With Application.ActiveSheet.MailEnvelope
      .Introduction = ""
      .Item.SentOnBehalfOfName = DE
      .Item.To = para
      .Item.CC = CC
      .Item.BCC = Cco
      .Item.Subject = Titulo
      '.Item.Attachments.Add (Rel)
      .Item.Send
   End With
   
ActiveWorkbook.EnvelopeVisible = False
   
'Windows(DESTINO).Close
''Activate.Window.Close savechanges:=False
   
With Application
.Application.DisplayAlerts = True
.ScreenUpdating = True
.EnableEvents = True
End With
   
MsgBox "Desempenho Enviado!"

End If
   
   
   
End Sub
