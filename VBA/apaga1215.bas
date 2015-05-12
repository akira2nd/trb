Attribute VB_Name = "apaga"
Sub Limpar()
Attribute Limpar.VB_ProcData.VB_Invoke_Func = " \n14"

If MsgBox("Deseja Limpar todos os dados?", vbYesNo, "Planejamento") = vbYes Then
Application.ScreenUpdating = False
    Sheets("CAPA").Select
    Range("C23:AH40,C44:C61,E44:E61,G44:G61,I44:I61,K44:K61").Select
    Selection.ClearContents
    Selection.ClearContents
    Selection.ClearContents

LA = 16
    Do While Sheets("PREMISSAS").Cells(LA, 10) <> ""
    ABA = Sheets("PREMISSAS").Cells(LA, 10)
    Sheets(ABA).Visible = True
   
    Sheets(ABA).Select
    Cells.Select
    Selection.ClearContents
    Selection.ClearContents
    Selection.ClearContents
    
    Sheets(ABA).Visible = False

    LA = LA + 1
    Loop
    
    Sheets("CAPA").Select
    Range("C23").Select
    
End If
Application.ScreenUpdating = True
End Sub

Sub arruma_super()

Dim supervisor, arruma


supervisor = 10
arruma = 5

Sheets("Ranking|Supervisores").Select
    
    Range("B10:E65100").Select
    Selection.ClearContents
    
    Rows("10:65100").Select
    Selection.Rows.Ungroup

    Selection.EntireRow.Hidden = False


Do While Sheets("ARRUMAR").Cells(arruma, 6).Value <> ""


    Sheets("Ranking|Supervisores").Select
   
    Sheets("Ranking|Supervisores").Cells(supervisor, 3).Value = Sheets("ARRUMAR").Cells(arruma, 6).Value
    Sheets("Ranking|Supervisores").Cells(supervisor, 2).Value = "x"
    
    Rows(supervisor + 1 & ":" & supervisor + 50).Select
    Selection.Rows.Group
    
    arruma = arruma + 1
    supervisor = supervisor + 51
    

    
    Loop

ActiveSheet.Outline.ShowLevels RowLevels:=1

End Sub

