Attribute VB_Name = "Módulo1"
Sub limpar()


If MsgBox("Deseja limpar os dados?", vbYesNo, "Planejamento") = vbYes Then

Application.ScreenUpdating = False

Sheets("ARRUMAR").Select
Range("B19:H38,B39:H56").Select
Selection.ClearContents

Sheets("INICIO").Select

MsgBox "Limpeza OK!"

End If

End Sub
