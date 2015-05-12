Attribute VB_Name = "CMS"
Sub CMS_export()

Application.DisplayAlerts = False
Application.ScreenUpdating = False


Dim caminho, L, PASTA, ARQUIVO, ORIGEM, DESTINO


ORIGEM = ActiveWorkbook.Name

Sheets("CMS_SJC").Visible = True
Sheets("CMS_SBC").Visible = True
Sheets("CMS_TRIAGEM").Visible = True
'Sheets("CMS_509SJC").Visible = True
Sheets("CMS_AUD_LIN").Visible = True
Sheets("CMS_LIN_TRANF").Visible = True
'Sheets("CMS_STO").Visible = True
'Sheets("CMS_942").Visible = True
'Sheets("CMS_899STO").Visible = True
'Sheets("CMS_509STO").Visible = True
'Sheets("CMS_AUD_STO").Visible = True

L = 16

Do While Sheets("PREMISSAS").Cells(L, 9) <> ""


caminho = Sheets("PREMISSAS").Cells(16, 2)
ARQUIVO = Sheets("PREMISSAS").Cells(L, 9)
PASTA = Sheets("PREMISSAS").Cells(L, 10)


Workbooks.OpenText Filename:=caminho & ARQUIVO, Origin:=xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1))

DESTINO = ActiveWorkbook.Name

Cells.Select
Cells.Copy


Windows(ORIGEM).Activate

Sheets(PASTA).Visible = True

Sheets(PASTA).Select
Range("A1").Select
Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                
                
                Windows(DESTINO).Activate
                Windows(DESTINO).Close
       


            If PASTA <> "VDN_TRANSFER" Then
                For L1 = 10 To 70
                
                    If Cells(L1, 1).Value = "Horário" Then
                        Cells(L1, 1).Select
                        Range(Selection, Selection.End(xlDown)).Select
                        Range(Selection, Selection.End(xlToRight)).Select
                        Selection.Cut
                        Range("Z6").Select
                        ActiveSheet.Paste
                        Range("A1").Select
                        
                        
                    End If
            Next
            End If
            
            Sheets(PASTA).Visible = False
            
            L = L + 1
            
            Loop
   
             Sheets("CAPA").Visible = True

    Application.CutCopyMode = False
    
Call CMS_RANKING
End Sub

Sub CMS_RANKING()
Application.ScreenUpdating = False

Dim skill
Dim Sk, ShCMS

Dim DIA As String
Dim col As Integer
    Dim ramal As String
    
        Dim periodo As String
        Dim flg As Boolean
        Dim b As Long
        Dim I As Long
        Dim info As Variant
        Dim rep As Variant
        Dim caminho  As String
        Dim caminhop  As String
        Dim hora, skills, data As String
        
        '------------------------------------------
  
        '------------------------------------------
        
Sk = 2

Do While Sheets("BASE_RANKING").Cells(Sk, 7).Value <> ""

ShCMS = Sheets("BASE_RANKING").Cells(Sk, 6).Value

        data = Sheets("CAPA").Range("B1")
        periodo = Format(data, "dd/mm/yyyy")
        hora = Sheets("CAPA").Cells(4, 13).Value - 0.01
        hora = Format(hora, "hh:30")

        skills = Sheets("BASE_RANKING").Cells(Sk, 7).Value
        '------------------------------------------
        PLANILHA = Application.ActiveWorkbook.Name
        '------------------------------------------
        caminho = Application.ActiveWorkbook.Path
        caminho = caminho & "\"
        '----------------------------------------

        
        CentreVu = Sheets("BASE_RANKING").Cells(Sk, 8).Value
        ARQUIVO = caminho & "\TEMP_CMS\" & "CMS_LIN.xls"
        
        '------------------------------------------
        


        Set acsApp = CreateObject("ACSUP.cvsApplication")
        CentreVuOpen = acsApp.Servers.Count
                    
        If (CentreVuOpen > 0) Then
            For I = 1 To CentreVuOpen
                If (acsApp.Servers.Item(I).Name = CentreVu) Then
                    Set acsSrv = acsApp.Servers.Item(I)
                    flg = True
                    Exit For
                End If
            Next
        Else
            MsgBox "Antes de tudo conecte o CMS!!!", vbOKOnly, "Planejamento"
            Exit Sub
        End If
        If Not flg Then
            MsgBox ("O servidor não foi localizado")
        End If
        
    Set acsCatalog = acsSrv.Reports
        acsSrv.Reports.ACD = "1"
        Set info = acsSrv.Reports.Reports("Historical\Designer\Desempenho da Equipe/Agente (INTERVALO) - Planejamento")
If info Is Nothing Then
        MsgBox "O relatório Real-Time\Designer\Agentes não foi encontrado no DAC 1.", vbCritical Or vbOKOnly, "Avaya CMS Supervisor"
    Exit Sub
Else
                b = acsSrv.Reports.CreateReport(info, rep)
            If b Then
                rep.Window.Top = 0
                rep.Window.Left = 0
                rep.Window.Width = 0
                rep.Window.Height = 0
                
                
                rep.ReportView.Add "G3,1,3;-1,2,0", "Grid1"
                rep.SetProperty "Grupo/Especialidade", skills
                rep.SetProperty "Data", data
                rep.SetProperty "Horários", "00:00-" & hora
                
                
              
                b = rep.Run
                b = rep.ExportData(ARQUIVO, 9, 0, True, True, True)
                If Not acsSrv.Interactive Then acsSrv.ActiveTasks.Remove rep.TaskID
                rep.Quit
                Set rep = Nothing
            End If
        End If
Set info = Nothing

    

       PLANILHA = Application.ActiveWorkbook.Name
        
        Workbooks.OpenText Filename:= _
        ARQUIVO
        
        fechar = Application.ActiveWorkbook.Name
        
        Cells.Select
        Selection.Copy
        
        
        Windows(PLANILHA).Activate
        
        
        Sheets(ShCMS).Visible = True
        Sheets(ShCMS).Select
        Cells.Select
                ActiveSheet.Paste
        Application.CutCopyMode = False
        
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[-2],5)"
    'Selection.AutoFill Destination:=Range("C1:C400"), Type:=xlFillDefault
    Selection.AutoFill Destination:=Range(Selection, ActiveCell.Cells(Range("A6").End(xlDown).Row, 1))
        
        
        
        
        Windows("Cms_LIN.xls").Close
        Selection.Replace What:=",000000000", Replacement:="0", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
 


       Sheets("CAPA").Select
       Sheets(ShCMS).Visible = False
       
Sk = Sk + 1
       
       Loop
       
       
Application.StatusBar = "Atualizado CMS"
Application.CutCopyMode = False
Application.ScreenUpdating = True
       
End Sub
