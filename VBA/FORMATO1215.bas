Attribute VB_Name = "FORMATO"
Sub FORMATAR()
Attribute FORMATAR.VB_ProcData.VB_Invoke_Func = " \n14"

Dim L
Dim ARQUIVO, FORMATO, DESTINO
Dim ABA_F, ABA_REL

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.CutCopyMode = False

If MsgBox("Deseja arrumar a formatação dessa coisa?", vbYesNo, "Akira Pergunta") = vbYes Then

L = 32

ARQUIVO = Sheets("PREMISSAS").Cells(30, 2).Value

DESTINO = ActiveWorkbook.Name

Workbooks.Open Filename:=ARQUIVO

FORMATO = ActiveWorkbook.Name

Windows(DESTINO).Activate

Do While Sheets("PREMISSAS").Cells(L, 1).Value <> ""

ABA_REL = Sheets("PREMISSAS").Cells(L, 1).Value
ABA_F = Sheets("PREMISSAS").Cells(L, 2).Value

Windows(FORMATO).Activate
  
    Sheets(ABA_F).Select
    Cells.Select
    Selection.Copy
    
    Windows(DESTINO).Activate
    Sheets(ABA_REL).Select
    
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
Sheets("CAPA").Select

L = L + 1

Loop

Windows(FORMATO).Activate
ActiveWindow.Close

MsgBox ("Arrumado...")

Else

MsgBox "Se não vai arrumar pare de clicar nesse botão Natividade!", vbExclamation, "Akira ò.ó"

End If

End Sub

