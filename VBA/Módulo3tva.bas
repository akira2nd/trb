Attribute VB_Name = "M�dulo3"
Sub TESTE()
Attribute TESTE.VB_ProcData.VB_Invoke_Func = " \n14"
'
' TESTE Macro
'

'
    ActiveCell.FormulaR1C1 = "=COUNTIFS(C6,RC6,C28,""PR� VENDA"",C36,""SPEEDY"")"
    Range("AY3").Select
End Sub
