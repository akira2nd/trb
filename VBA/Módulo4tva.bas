Attribute VB_Name = "Módulo4"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveCell.FormulaR1C1 = "=COUNTIFS(C[-46],RC[-46],C[-30],""<>""&RC[-30])"
    Range("AZ3").Select
End Sub
