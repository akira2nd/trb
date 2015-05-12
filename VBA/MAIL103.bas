Attribute VB_Name = "MAIL"
Sub envia_planilha()
   
Dim DE, para, CC, Cco, Titulo, introducao
Dim Rel


If MsgBox("Deseja enviar E-mail?", vbYesNo, "Planejamento") = vbYes Then


With Application
.Application.DisplayAlerts = False
.ScreenUpdating = False
.EnableEvents = False
End With


Rel = Sheets("ARRUMAR").Range("I8").Value
DE = Sheets("ARRUMAR").Range("I3").Text
para = Sheets("ARRUMAR").Range("I4").Text
CC = Sheets("ARRUMAR").Range("I5").Text
Cco = Sheets("ARRUMAR").Range("I6").Text
Titulo = Sheets("ARRUMAR").Range("I7").Text

ORIGEM = ActiveWorkbook.Name

        Workbooks.Open Filename:=Rel, _
        UpdateLinks:=False, ReadOnly:=True, IgnoreReadOnlyRecommended:=True
        
DESTINO = ActiveWorkbook.Name

Windows(DESTINO).Activate
   ' Seleciona a planilha
   Sheets("Resumo").Activate
   ActiveSheet.Cells.Select
   
   ' mostrar o envelope na planilha ativa
   ActiveWorkbook.EnvelopeVisible = True
   
   ' atribui valor aos campos
   ' to = para **cc = CC **BCC = Cco **subject = título
   ' Attachments = Anexo
   With Application.ActiveSheet.MailEnvelope
      .Introduction = ""
      .Item.SentOnBehalfOfName = DE
      .Item.To = para
      .Item.CC = CC
      .Item.BCC = Cco
      .Item.Subject = Titulo
      .Item.Attachments.Add (Rel)
      .Item.Send
   End With
   
ActiveWorkbook.EnvelopeVisible = False
   
Windows(DESTINO).Close
'Activate.Window.Close savechanges:=False
   
With Application
.Application.DisplayAlerts = True
.ScreenUpdating = True
.EnableEvents = True
End With
   
MsgBox "Desempenho Enviado", , "Planejamento"

Application.StatusBar = "Email enviado!"

End If


End Sub

