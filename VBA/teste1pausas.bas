Attribute VB_Name = "teste1"
Sub Banco()


Application.StatusBar = "ATUALIZANDO BANCO"
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim Caminho, ORIGEM, DESTINO

Caminho = Sheets("BD").Cells(6, 13)
ORIGEM = ActiveWorkbook.Name

    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "BANCO"
    Sheets("BANCO").Select
'Sheets("BANCO").Visible = True
'Sheets("BANCO").Select
'Cells.ClearContents

Workbooks.Open Filename:=Caminho

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
        
        Sheets("INICIO").Select
        
                
Windows(DESTINO).Activate
Windows(DESTINO).Close

Application.StatusBar = "BANCO ATUALIZADO"

 
Call Filtro

Application.CutCopyMode = False

MsgBox "Banco Atualizado", , "Planejamento"
End Sub

Sub Filtro()

Dim S

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Sheets("BD").Select
Range("C11:F65000").Clear

Sheets("BANCO").Select
   
    
   'SUPERVISOR
    
Range("AR2:AR65000").Select
Range("AR2:AR65000").Copy

Sheets("BD").Select
Range("F11").Select

Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Sheets("BANCO").Select

'OPERADOR
Range("B2:B65000").Select
Range("B2:B65000").Copy

Sheets("BD").Select
Range("D11").Select

Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Sheets("BANCO").Select

'RE
Range("D2:D65000").Select
Range("D2:D65000").Copy

Sheets("BD").Select
Range("E11").Select

Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Sheets("BANCO").Select
    
'LOGIN
Range("M2:M65000").Select
Range("M2:M65000").Copy

Sheets("BD").Select
Range("C11").Select

Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False

    Sheets("BANCO").Select
    ActiveWindow.SelectedSheets.Delete

Application.CutCopyMode = False

Sheets("BD").Select

'arruma agentes sem hierarquia
S = 11
Do While Cells(S, 5).Value <> ""
    
    If Cells(S, 6).Value = "" Then
        
        Cells(S, 6).Value = "SEM SUPERVISOR NO WFM"
        
    End If

S = S + 1

Loop

Sheets("INICIO").Select

End Sub
