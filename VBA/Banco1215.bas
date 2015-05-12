Attribute VB_Name = "Banco"
Sub Banco()


Application.StatusBar = "ATUALIZANDO BANCO"
Application.ScreenUpdating = False
Application.DisplayAlerts = False



Dim CRITERIO1, CRITERIO2, CRITERIO3, CRITERIO4, CRITERIO5, CRITERIO6
Dim caminho, ORIGEM, DESTINO

caminho = Sheets("PREMISSAS").Cells(18, 2)
ORIGEM = ActiveWorkbook.Name

 Sheets("BANCO").Visible = True
Sheets("BANCO").Select
Cells.ClearContents

Workbooks.Open Filename:=caminho

DESTINO = ActiveWorkbook.Name
    
    Sheets("bd_Speedy").Select
    Rows("1:1").Select
    Selection.AutoFilter
    
    L = 16
    
    Windows(ORIGEM).Activate
    
    'Do While Sheets("PREMISSAS").Cells(L, 13) <> ""
    
    
    'CRITERIO1 = Sheets("PREMISSAS").Cells(L, 13)
    
    'L = L + 1
    
    'CRITERIO2 = Sheets("PREMISSAS").Cells(L, 13)
    
    'L = L + 1
    
    'CRITERIO3 = Sheets("PREMISSAS").Cells(L, 13)
    
    'L = L + 1
    
    'CRITERIO4 = Sheets("PREMISSAS").Cells(L, 13)
    
    'L = L + 1
    
    'CRITERIO5 = Sheets("PREMISSAS").Cells(L, 13)
     
     'L = L + 1
     
     'CRITERIO6 = Sheets("PREMISSAS").Cells(L, 13)
     
     Windows(DESTINO).Activate
    
    
  
    'ActiveSheet.Range("$A$1:$AX$2094").AutoFilter Field:=29, Criteria1:=Array(CRITERIO1, CRITERIO2, CRITERIO3, CRITERIO4, CRITERIO5, CRITERIO6), Operator:=xlFilterValues
   
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

L = L + 1


 'Loop
 
 Sheets("BANCO").Visible = False
 
Call Filtro

Application.CutCopyMode = False
End Sub
