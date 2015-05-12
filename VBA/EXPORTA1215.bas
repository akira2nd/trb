Attribute VB_Name = "EXPORTA"
Sub EXPORTA_CAPA()


Dim L, hora, L2

hora = Sheets("CAPA").Cells(5, 13)

For L = 23 To 40

If hora = Sheets("CAPA").Cells(L, 2) Then


'ILHA CLIENT_SPD
Sheets("CAPA").Range(Cells(L, 3), Cells(L, 10)).Value = Sheets("CAPA").Range(Cells(5, 29), Cells(5, 37)).Value

'ILHA AQUI_SPD
Sheets("CAPA").Range(Cells(L, 11), Cells(L, 18)).Value = Sheets("CAPA").Range(Cells(6, 29), Cells(6, 37)).Value

'3g
Sheets("CAPA").Range(Cells(L, 19), Cells(L, 26)).Value = Sheets("CAPA").Range(Cells(7, 29), Cells(7, 37)).Value

'linha 8
Sheets("CAPA").Range(Cells(L, 27), Cells(L, 34)).Value = Sheets("CAPA").Range(Cells(8, 29), Cells(8, 37)).Value

'''linha 9
'''Sheets("CAPA").Range(Cells(L, 11), Cells(L, 18)).Value = Sheets("CAPA").Range(Cells(6, 29), Cells(6, 37)).Value

End If
Next

'MIX SPEEDY
For L2 = 44 To 61
    If hora = Sheets("CAPA").Cells(L2, 2) Then
        
        Sheets("CAPA").Cells(L2, 3) = Sheets("CAPA").Cells(5, 43)
        Sheets("CAPA").Cells(L2, 5) = Sheets("CAPA").Cells(6, 43)
        Sheets("CAPA").Cells(L2, 7) = Sheets("CAPA").Cells(7, 43)
        Sheets("CAPA").Cells(L2, 9) = Sheets("CAPA").Cells(8, 43)
        Sheets("CAPA").Cells(L2, 11) = Sheets("CAPA").Cells(9, 43)

    End If
Next

Application.CutCopyMode = False
End Sub

Sub Gravar()



   Application.DisplayAlerts = False
   Application.ScreenUpdating = False
   Application.StatusBar = "...Gravando..."


    Dim caminho As Variant

caminho = Sheets("PREMISSAS").Range("B19")

 
    'Sheets("AUDITORIA").Select
 
'    Sheets(Array("Recebidas Trans Movel", "Recebidas da Plataforma", "SPEEDY_UNIFICADO", "CLIENTE VIVO SPEEDY", "AQUISIÇÃO VIVO SPEEDY", _
'        "AQUISIÇÃO INTERNET MÓVEL", "INTERNET MÓVEL", "AUDITORIA", "Resumo", _
'        "Ranking|Supervisores", "URA_SKILL_22")).Select
'    Sheets(Array("Recebidas Trans Movel", "Recebidas da Plataforma", "SPEEDY_UNIFICADO", "CLIENTE VIVO SPEEDY", "AQUISIÇÃO VIVO SPEEDY", "AQUISIÇÃO INTERNET MÓVEL", "INTERNET MÓVEL", "AUDITORIA", "Resumo", "Ranking|Supervisores", "URA_SKILL_22")).Copy


        Sheets(Array("SPEEDY_UNIFICADO", "CLIENTE VIVO SPEEDY", "AQUISIÇÃO VIVO SPEEDY", _
        "AQUISIÇÃO INTERNET MÓVEL", "AUDITORIA", "Resumo", _
        "Ranking|Supervisores")).Select
        Sheets(Array("SPEEDY_UNIFICADO", "CLIENTE VIVO SPEEDY", "AQUISIÇÃO VIVO SPEEDY", _
        "AQUISIÇÃO INTERNET MÓVEL", "AUDITORIA", "Resumo", _
        "Ranking|Supervisores")).Copy

   
'    Sheets("Recebidas Trans Movel").Select
'    Cells.Select
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'
'    Sheets("Recebidas da Plataforma").Select
'    Cells.Select
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Sheets("SPEEDY_UNIFICADO").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Sheets("CLIENTE VIVO SPEEDY").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Sheets("AQUISIÇÃO VIVO SPEEDY").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    
   
    
    Sheets("AUDITORIA").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
   
   Sheets("Resumo").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    
   Sheets("Ranking|Supervisores").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
   

   
    Sheets("Resumo").Select
    
    ActiveWorkbook.SaveAs Filename:=caminho, WriteResPassword:="PLAN01", ReadOnlyRecommended:=False
    Application.DisplayAlerts = False
    
    ActiveWindow.Close
    Application.StatusBar = "...Gravado..."
    
  Application.CutCopyMode = False
   Application.ScreenUpdating = True
   
     Sheets("CAPA").Select
    
End Sub

