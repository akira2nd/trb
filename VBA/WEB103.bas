Attribute VB_Name = "WEB"
Declare Sub Sleep Lib "Kernel32" (ByVal dwmilliseconds As Long)

Sub um_botao()


If MsgBox("Deseja rodar WEB das" & " " & Format(Sheets("CAPA").Cells(4, 13).Value, "HH:MM"), vbYesNo, "Planejamento") = vbYes Then

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Calculate
        
    Sheets("CAPA").Select
               Application.CutCopyMode = False
    
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.StatusBar = "Importando WEB"
        
    Sheets("Piloto").Visible = True
    Sheets("Assertividade").Visible = True
    
    Dim IE As New InternetExplorer
    Dim WshShell As Object
    Dim dtInicio As Date
    Dim objIE As New DataObject
  
    
    Sheets("Piloto").Select
    Sheets("Piloto").Cells.Clear
    'Rows("3:65536").Select
    'Selection.Delete Shift:=xlUp
        
 
    
        
    IE.Visible = False
        
    
    
    DATAI = Sheets("CAPA").Range("B1").Value
    'DATAI = InputBox("DATA INICIAL PARA A EXTRAÇÃO - DD/MM/AAAA", "Planejamento", Format((Now), "DD") & "/" & Format((Now), "mm") & "/" & Format((Now), "yyyy"))
    DATAF = DATAI

    IE.navigate "http://10.2.1.130/rel_telefonica/principal.asp"

    If Sheets("CAPA").Range("Z1") = True Then
    
    End If

    While IE.Busy Or IE.readyState <> READYSTATE_COMPLETE
        Sleep (1000)
        DoEvents
    Wend
    Wusuario = Sheets("PREMISSAS").Range("B24")
    senha = Sheets("PREMISSAS").Range("B25")

    IE.document.forms(0).re.Value = Wusuario
    IE.document.forms(0).senha.Value = senha
    IE.document.forms(0).submit
    
    While IE.Busy Or IE.readyState <> READYSTATE_COMPLETE
        Sleep (1000)
        DoEvents
    Wend

    IE.document.forms(0).cmbGrupo.Value = "VL"
    IE.document.forms(0).cmbGrupo.onchange

    While IE.Busy Or IE.readyState <> READYSTATE_COMPLETE
        Sleep (1000)
        DoEvents
    Wend

    IE.document.forms(0).site.selectedIndex = 1
    IE.document.forms(0).submit

    While IE.Busy Or IE.readyState <> READYSTATE_COMPLETE
        Sleep (1000)
        DoEvents
    Wend

    IE.document.forms(0).tipo_relatorio.selectedIndex = 2
    IE.document.forms(0).tipo_relatorio.onchange
    IE.document.forms(0).strDiaIniADetalhe.Value = Day(DATAI)
    IE.document.forms(0).strMesIniADetalhe.Value = Month(DATAI)
    IE.document.forms(0).strAnoIniADetalhe.Value = Year(DATAI)
    IE.document.forms(0).strDiaFimADetalhe.Value = Day(DATAF)
    IE.document.forms(0).strMesFimADetalhe.Value = Month(DATAF)
    IE.document.forms(0).strAnoFimADetalhe.Value = Year(DATAF)
    
    

    'APERTAR ENVIAR
    Application.StatusBar = "Abrindo Relátorio"
    IE.document.forms(0).Action = "rel_vl_analitico_detalhe.asp"
    IE.document.forms(0).submit
    While IE.Busy Or IE.readyState <> READYSTATE_COMPLETE
    Sleep (1000)
    DoEvents
    Wend
    

    'SALVAR VALORES NA MEMORIA
    Application.StatusBar = "...Salvando dados na Memoria..."
    objIE.SetText (IE.document.Body.outerHTML)
    objIE.PutInClipboard
    
    
   Application.StatusBar = "...Copiando os dados..."
    
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "TEMP"
    Sheets("TEMP").Select
    Worksheets("TEMP").Cells(1, 1).PasteSpecial 'COLA DA WEB NO EXCEL
   
    'FINALIZA


    IE.navigate "http://10.2.1.130/rel_telefonica/logout2.asp"

    While IE.Busy Or IE.readyState <> READYSTATE_COMPLETE
        Sleep (1000)
        DoEvents
    Wend

    IE.Quit

    Set IE = Nothing

    Sheets("TEMP").Select
    Cells.Select
    Selection.Copy
    Sheets("Piloto").Select
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    Sheets("TEMP").Select
    ActiveWindow.SelectedSheets.Delete
       

    'If Sheets("CAPA").Range("Z1") = False Then h = MsgBox("EXTRAÇÃO COMPLETA...AGUARDE SÓ MAIS UM MOMENTO!OBRIGADO", vbInformation, "Planejamento")
   
   
    Application.StatusBar = "...Finalizando Copia dos dados..."
    Call limpa_data_2
    Sheets("CAPA").Select
    Application.StatusBar = False
      
    Sheets("CAPA").Select
               Application.CutCopyMode = False
    
    CMS_export
    Sheets("CAPA").Select
               Application.CutCopyMode = False
    
    EXPORTAR
    Sheets("CAPA").Select
               Application.CutCopyMode = False
    
    ranking_novo
    Sheets("CAPA").Select
               Application.CutCopyMode = False
    
    GRAVAR_planilha
    Sheets("CAPA").Select
               Application.CutCopyMode = False
               
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    
End If


End Sub

Sub VENDAS_3()
    

If MsgBox("Deseja rodar WEB das" & " " & Format(Sheets("CAPA").Cells(4, 13).Value, "HH:MM"), vbYesNo, "Planejamento") = vbYes Then

    Calculate
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.StatusBar = "Importando WEB"
        
    Sheets("Piloto").Visible = True
    Sheets("Assertividade").Visible = True
    
    Dim IE As New InternetExplorer
    Dim WshShell As Object
    Dim dtInicio As Date
    Dim objIE As New DataObject
  
    
    Sheets("Piloto").Select
    Sheets("Piloto").Cells.Clear
    'Rows("3:65536").Select
    'Selection.Delete Shift:=xlUp
        
 
    
        
    IE.Visible = False
        
    
    
    DATAI = Sheets("CAPA").Range("B1").Value
    'DATAI = InputBox("DATA INICIAL PARA A EXTRAÇÃO - DD/MM/AAAA", "Planejamento", Format((Now), "DD") & "/" & Format((Now), "mm") & "/" & Format((Now), "yyyy"))
    DATAF = DATAI

    IE.navigate "http://10.2.1.130/rel_telefonica/principal.asp"

    If Sheets("CAPA").Range("Z1") = True Then
    
    End If

    While IE.Busy Or IE.readyState <> READYSTATE_COMPLETE
        Sleep (1000)
        DoEvents
    Wend
    Wusuario = Sheets("PREMISSAS").Range("B24")
    senha = Sheets("PREMISSAS").Range("B25")

    IE.document.forms(0).re.Value = Wusuario
    IE.document.forms(0).senha.Value = senha
    IE.document.forms(0).submit
    
    While IE.Busy Or IE.readyState <> READYSTATE_COMPLETE
        Sleep (1000)
        DoEvents
    Wend

    IE.document.forms(0).cmbGrupo.Value = "VL"
    IE.document.forms(0).cmbGrupo.onchange

    While IE.Busy Or IE.readyState <> READYSTATE_COMPLETE
        Sleep (1000)
        DoEvents
    Wend

    IE.document.forms(0).site.selectedIndex = 1
    IE.document.forms(0).submit

    While IE.Busy Or IE.readyState <> READYSTATE_COMPLETE
        Sleep (1000)
        DoEvents
    Wend

    IE.document.forms(0).tipo_relatorio.selectedIndex = 2
    IE.document.forms(0).tipo_relatorio.onchange
    IE.document.forms(0).strDiaIniADetalhe.Value = Day(DATAI)
    IE.document.forms(0).strMesIniADetalhe.Value = Month(DATAI)
    IE.document.forms(0).strAnoIniADetalhe.Value = Year(DATAI)
    IE.document.forms(0).strDiaFimADetalhe.Value = Day(DATAF)
    IE.document.forms(0).strMesFimADetalhe.Value = Month(DATAF)
    IE.document.forms(0).strAnoFimADetalhe.Value = Year(DATAF)
    
    

    'APERTAR ENVIAR
    Application.StatusBar = "Abrindo Relátorio"
    IE.document.forms(0).Action = "rel_vl_analitico_detalhe.asp"
    IE.document.forms(0).submit
    While IE.Busy Or IE.readyState <> READYSTATE_COMPLETE
    Sleep (1000)
    DoEvents
    Wend
    

    'SALVAR VALORES NA MEMORIA
    Application.StatusBar = "...Salvando dados na Memoria..."
    objIE.SetText (IE.document.Body.outerHTML)
    objIE.PutInClipboard
    
    
   Application.StatusBar = "...Copiando os dados..."
    
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "TEMP"
    Sheets("TEMP").Select
    Worksheets("TEMP").Cells(1, 1).PasteSpecial 'COLA DA WEB NO EXCEL
   
    'FINALIZA


    IE.navigate "http://10.2.1.130/rel_telefonica/logout2.asp"

    While IE.Busy Or IE.readyState <> READYSTATE_COMPLETE
        Sleep (1000)
        DoEvents
    Wend

    IE.Quit

    Set IE = Nothing

    Sheets("TEMP").Select
    Cells.Select
    Selection.Copy
    Sheets("Piloto").Select
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    Sheets("TEMP").Select
    ActiveWindow.SelectedSheets.Delete
       

    'If Sheets("CAPA").Range("Z1") = False Then h = MsgBox("EXTRAÇÃO COMPLETA...AGUARDE SÓ MAIS UM MOMENTO!OBRIGADO", vbInformation, "Planejamento")
   
   
    Application.StatusBar = "...Finalizando Copia dos dados..."
    Call limpa_data_2
    Sheets("CAPA").Select
    Application.StatusBar = False
    

 
 End If
  
 
 
 MsgBox "Processo WEB Finalizado", , "Planejamento"
 
   
End Sub

Sub limpa_data_2()
    
    Dim hs_limite As Date
    Dim lin As Long
    
    
    Application.StatusBar = "...Arrumando data..."
    hs_limite = Worksheets("CAPA").Range("M4").Value
    
    Worksheets("Piloto").Select
    
    'Monta formula
    
    lin = 18
    Do
    lin = lin + 1
    Loop Until Len(Cells(lin + 1, 1).Value) = 0
    Range(Cells(18, 58), Cells(lin, 58)).FormulaR1C1 = "=HOUR(LEFT(RC[-52],LEN(RC[-52])-1))/24+MINUTE(LEFT(RC[-52],LEN(RC[-52])-1))/24/60"
    Range(Cells(18, 58), Cells(lin, 58)).NumberFormat = "[hh]:mm"
    
    'Cola valor
    
    Range(Cells(18, 58), Cells(lin, 58)).Copy
    Cells(18, 6).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    'Exclui ítens fora do horário
    
    
    lin = 18
    Do
        If Cells(lin, 6).Value >= hs_limite Then Range(Cells(lin, 6), Cells(40000, 6)).EntireRow.Delete
        lin = lin + 1
    Loop Until Len(Cells(lin, 6).Value) = 0
    
      
    'Columns("A:AK").Select
    'Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    'Calculate
    
    Rows("16:16").Select
    Selection.Delete Shift:=xlUp
    
    Application.StatusBar = "...Batendo Assertividade..."
    Call Batendo_asservidade
    

End Sub


Sub Batendo_asservidade()


Application.DisplayAlerts = False
Application.ScreenUpdating = False


Dim TOTAL

TOTAL = Sheets("PREMISSAS").Range("b26")

Sheets("ASSERTIVIDADE").Select
Sheets("ASSERTIVIDADE").Cells.Clear

Columns("IC:IV").Select
Selection.Delete Shift:=xlToLeft
Selection.Delete Shift:=xlToLeft
Sheets("PILOTO").Select

Range("A16").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("ASSERTIVIDADE").Select

Range("A1").Select
'ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Columns("C:C").Select
Application.CutCopyMode = False
Selection.Insert Shift:=xlToRight
Range("C1").Select
ActiveCell.FormulaR1C1 = "duplicados"
    
    
    Columns("A:AK").Select
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Calculate
    
        
Range("C2").Select
ActiveCell.FormulaR1C1 = "=COUNTIF(C[-1],RC[-1])"
Range("C2").Select
Selection.Copy
    
Range("C2:" & "C" & TOTAL).Select
Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False
        
        
Sheets("PILOTO").Visible = False
Sheets("ASSERTIVIDADE").Visible = False
Sheets("CAPA").Select

    Application.CutCopyMode = False

End Sub
