Attribute VB_Name = "APAGA"
Sub LIMPAR()
Attribute LIMPAR.VB_ProcData.VB_Invoke_Func = " \n14"
   
If MsgBox("Deseja limpar todos os dados?", vbYesNo, "Planejamento") = vbYes Then
   
   Range("C23:J41,C46:J64,C69:J87,N23:Q41,N46:P64,N69:Q87,C91:D127").Select
    Selection.ClearContents
    Selection.ClearContents
    Selection.ClearContents
 
 End If
End Sub
