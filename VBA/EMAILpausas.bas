Attribute VB_Name = "EMAIL"
Sub envia_planilha()
   
Dim DE, para, CC, Cco, Titulo, introducao
Dim Rel
Dim link

If MsgBox("Deseja Enviar o e-mail?", vbYesNo, "Planejamento") = vbYes Then


link = Sheets("Resumo").Range("B6").Value
Sheets("Resumo").Select
    Range("B6").Select
    Selection.Hyperlinks(1).Address = link


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

'ORIGEM = ActiveWorkbook.Name
'
'        Workbooks.Open Filename:=Rel, _
'        UpdateLinks:=False, ReadOnly:=True, IgnoreReadOnlyRecommended:=True
'
'DESTINO = ActiveWorkbook.Name

'Windows(DESTINO).Activate
   '' Seleciona a planilha
   Sheets("Resumo").Activate
   ActiveSheet.Cells.Select
   
   '' mostrar o envelope na planilha ativa
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
      '.Item.Attachments.Add (Rel)
      .Item.Send
   End With
   
ActiveWorkbook.EnvelopeVisible = False
   
'Windows(DESTINO).Close
''Activate.Window.Close savechanges:=False
   
With Application
.Application.DisplayAlerts = True
.ScreenUpdating = True
.EnableEvents = True
End With
   
MsgBox "Desempenho Enviado!"

End If
   
   
   
End Sub



Sub Envia_Email()

'Sheets("PREMISSAS").Visible = True

Dim ofolder  As Outlook.MAPIFolder
Dim oitem As Outlook.MailItem
Dim oOutlook As New Outlook.Application
Dim moOutlook As Outlook.Namespace
'Dim rng As Range
    
    
rng = Sheets("Resumo").Range("A:AA").Value 'SpecialCells(xlCellTypeVisible)
    
    
ActiveWorkbook.EnvelopeVisible = True
    
    If (Sheets("ARRUMAR").Range("I3").Text) <> "" Then
    Set moOutlook = oOutlook.GetNamespace("MAPI")
    Set ofolder = moOutlook.GetDefaultFolder(olFolderInbox)
    Set oitem = ofolder.Items.Add(olMailItem)
    With oitem
    .Recipients.Add (Sheets("ARRUMAR").Range("I3").Text)
    If IsNull(Sheets("ARRUMAR").Range("I3").Text) = False Then
          '.To =
          .CC = Sheets("ARRUMAR").Range("I5").Text
          .BCC = Sheets("ARRUMAR").Range("I6").Text
    End If
            .Subject = Sheets("ARRUMAR").Range("I7").Text
            .BodyFormat = olFormatHTML
            .Body = rng
            '.HTMLBody = rng
            '.HTMLBody = Sheets("Resumo").Range("A:AA")
            .Attachments.Add (Sheets("ARRUMAR").Range("I8").Value)
'            .Importance = olImportanceHigh
            .Send
    End With
    
    Set oitem = Nothing
    Set ofolder = Nothing
    Else:
      MsgBox "digite o distinatário do email"
    End If
    
    
    Sheets("CAPA").Activate
    
    
End Sub


