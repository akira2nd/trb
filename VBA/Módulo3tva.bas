Attribute VB_Name = "Módulo3"
Sub TESTE()
Attribute TESTE.VB_ProcData.VB_Invoke_Func = " \n14"
'
' TESTE Macro
'

'
    ActiveCell.FormulaR1C1 = "=COUNTIFS(C6,RC6,C28,""PRÉ VENDA"",C36,""SPEEDY"")"
    Range("AY3").Select
End Sub
