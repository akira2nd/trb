Attribute VB_Name = "Módulo3"
Sub Gravar()

    Dim Caminho As String

    Caminho = Sheets("INICIO").Cells(2, 8)

'    Sheets(Array("CAPA", "ITV", "LINHAS_SJC", "AUDITORIA_STO", "SPD_OUTBOUND", "AUDITORIA_ITV", _
'        "AUDITORIA_LINHAS", "AUDITORIA_ATV", "TVA_OUTBOUND", "BKO_SBA", "RET_SBA", "RET_SBC")).Copy

Sheets(Array("CAPA", "ITV", "LINHAS_SJC", "SPD_OUTBOUND", "AUDITORIA_ITV", _
        "AUD_LIN_SJC", "AUDITORIA_ATV", "TVA_OUTBOUND", "AUD_TVA")).Copy
        
    Application.DisplayAlerts = False
    
'    Sheets(Array("CAPA", "ITV", "LINHAS_SJC", "AUDITORIA_STO", "SPD_OUTBOUND", "AUDITORIA_ITV", _
'        "AUDITORIA_LINHAS", "AUDITORIA_ATV", "TVA_OUTBOUND", "BKO_SBA", "RET_SBA", "RET_SBC")).Select

Sheets(Array("CAPA", "ITV", "LINHAS_SJC", "SPD_OUTBOUND", "AUDITORIA_ITV", _
        "AUD_LIN_SJC", "AUDITORIA_ATV", "TVA_OUTBOUND", "AUD_TVA")).Select



    
    Rows("1000:1000").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    
    'Sheets("CAPA").Select
    Range("A1").Select
    
    
    ActiveWorkbook.SaveAs Filename:=Caminho, WriteResPassword:="PLAN01", ReadOnlyRecommended:=False
    
    
    
    'ActiveWorkbook.SaveAs Filename:= _
        Caminho & "\Consolidado de Pausas " & Format(Now(), "dd") & ".xls" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

    Application.DisplayAlerts = True

    MsgBox "Arquivo Salvo", vbInformation, "Thiago Vale Cap"

End Sub
