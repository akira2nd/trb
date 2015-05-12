Attribute VB_Name = "Banco"
Sub Banco()


Application.StatusBar = "ATUALIZANDO BANCO"
Application.ScreenUpdating = False
Application.DisplayAlerts = False



Dim CRITERIO1, CRITERIO2, CRITERIO3, CRITERIO4, CRITERIO5, CRITERIO6
Dim caminho, ORIGEM, DESTINO

caminho = Sheets("PREMISSAS").Cells(18, 2)
ORIGEM = ActiveWorkbook.Name

    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "BANCO"
    Sheets("BANCO").Select
'Sheets("BANCO").Visible = True
'Sheets("BANCO").Select
'Cells.ClearContents

Workbooks.Open Filename:=caminho

DESTINO = ActiveWorkbook.Name
    
    Sheets("bd_Speedy").Select
     
     Windows(DESTINO).Activate

   
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
     Windows(ORIGEM).Activate
    
    Sheets("BANCO").Select
   
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Sheets("CAPA").Select
        
                
Windows(DESTINO).Activate
Windows(DESTINO).Close

Application.StatusBar = "BANCO ATUALIZADO"

 
Call Filtro

Application.CutCopyMode = False
End Sub

Sub Filtro()


Application.ScreenUpdating = False
Application.DisplayAlerts = False
Sheets("BASE_RANKING").Select
Range("G26:J65530").Clear

Sheets("BANCO").Select
   
    
   'SUPERVISOR
    
Range("AR2:AR6000").Select
Range("AR2:AR6000").Copy

Sheets("BASE_RANKING").Select
Range("G26").Select

Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Sheets("BANCO").Select

'OPERADOR
Range("B2:B6000").Select
Range("B2:B6000").Copy

Sheets("BASE_RANKING").Select
Range("H26").Select

Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Sheets("BANCO").Select

'RE
Range("D2:D6000").Select
Range("D2:D6000").Copy

Sheets("BASE_RANKING").Select
Range("I26").Select

Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Sheets("BANCO").Select
    
'LOGIN
Range("M2:M6000").Select
Range("M2:M6000").Copy

Sheets("BASE_RANKING").Select
Range("J26").Select

Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False

    Sheets("BANCO").Select
    ActiveWindow.SelectedSheets.Delete

Sheets("CAPA").Select
Application.CutCopyMode = False

End Sub
