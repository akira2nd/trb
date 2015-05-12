Attribute VB_Name = "GRAVAR"
Sub GRAVAR_planilha()


 Application.DisplayAlerts = False
 Application.ScreenUpdating = False


Dim caminho As Variant


caminho = Sheets("PREMISSAS").Range("B19")


 
 
    'Sheets(Array("Resumo", "10315-Geral", "10315-SJC", "Resumo Marca", "10315-STO", "AUDITORIA", "Ranking|Supervisores")).Select
    'Sheets(Array("Resumo", "10315-Geral", "10315-SJC", "Resumo Marca", "10315-STO", "AUDITORIA", "Ranking|Supervisores")).Copy
    'Sheets(Array("Resumo", "10315-Geral", "AUDITORIA", "Resumo Marca", "Ranking|Supervisores")).Select
    'Sheets(Array("Resumo", "10315-Geral", "AUDITORIA", "Resumo Marca", "Ranking|Supervisores")).Copy
    
     Sheets(Array("Resumo", "10315-Geral", "AUDITORIA", "Ranking|Supervisores")).Select
    Sheets(Array("Resumo", "10315-Geral", "AUDITORIA", "Ranking|Supervisores")).Copy
    
    
    Sheets("Resumo").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           Application.CutCopyMode = False
   
    Sheets("10315-Geral").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           Application.CutCopyMode = False

'    Sheets("10315-SJC").Select
'    Cells.Select
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'           Application.CutCopyMode = False
'
'    Sheets("10315-STO").Select
'    Cells.Select
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'           Application.CutCopyMode = False

    
    Sheets("AUDITORIA").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           Application.CutCopyMode = False
           
    Sheets("Ranking|Supervisores").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           Application.CutCopyMode = False
           
     'Sheets("Resumo Marca").Select
    'Cells.Select
    'Selection.Copy
    'Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           Application.CutCopyMode = False
           
        Sheets("Resumo").Select
    
    ActiveWorkbook.SaveAs Filename:=caminho, WriteResPassword:="PLAN01", ReadOnlyRecommended:=True
   
    
    ActiveWindow.Close
    Application.StatusBar = "...Gravado..."
    
    Sheets("CAPA").Select
       Application.CutCopyMode = False
    
End Sub
