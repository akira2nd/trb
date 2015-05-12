Attribute VB_Name = "RANKING"
Sub Filtro()
Attribute Filtro.VB_ProcData.VB_Invoke_Func = " \n14"


Application.ScreenUpdating = False
Application.DisplayAlerts = False
Sheets("Ranking_Operador").Select
Range("c21:h6000, j21:j6000, l21:l6000, n21:n6000, p21:p6000").Clear

Sheets("BANCO").Visible = True
Sheets("BANCO").Select
   
    Range("J1").Select
    
    'Selection.AutoFilter
    'ActiveSheet.Range("$A$1:$AX$652").AutoFilter Field:=34, Criteria1:="ATIVO"
    
   'SUPERVISOR
    
Range("AR2:AR5000").Select
Range("AR2:AR5000").Copy

Sheets("Ranking_Operador").Select
Range("C21").Select

Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Sheets("BANCO").Select

'OPERADOR
Range("B2:B5000").Select
Range("B2:B5000").Copy

Sheets("Ranking_Operador").Select
Range("D21").Select

Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Sheets("BANCO").Select

'RE
Range("D2:D5000").Select
Range("D2:D5000").Copy

Sheets("Ranking_Operador").Select
Range("E21").Select

Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Sheets("BANCO").Select
    
'LOGIN

Range("M2:M5000").Select
Range("M2:M5000").Copy

Sheets("Ranking_Operador").Select
Range("F21").Select

Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False



Sheets("BANCO").Select
Selection.AutoFilter
Sheets("BANCO").Visible = False

Sheets("CAPA").Select
Application.CutCopyMode = False

End Sub

Sub CMS_RANKING()
Application.ScreenUpdating = False


Dim PLANILHA, CentreVu, ARQUIVO, acsApp, CentreVuOpen, acsSrv, acsCatalog, fechar

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
        
        
        data = Sheets("CAPA").Range("B1")
        periodo = data 'format(data, "dd/mm/yyyy")
        hora = "23:30" 'Sheets("CAPA").Cells(4, 13).Value - 0.01
        'hora = format(hora, "hh:30")

        skills = Sheets("PREMISSAS").Range("B11")
        '------------------------------------------
        PLANILHA = Application.ActiveWorkbook.Name
        '------------------------------------------
        caminho = Application.ActiveWorkbook.Path
        caminho = caminho & "\"
        '----------------------------------------

        
        CentreVu = "10.4.0.90"
        ARQUIVO = caminho & "\TEMP_CMS\" & "CMS_UNI.xls"
        
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
                rep.SetProperty "Horários", "06:00-" & hora
                
                
              
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
        
        Sheets("CMS_RANKING").Visible = True
        Sheets("CMS_RANKING").Select
        Cells.Select
                ActiveSheet.Paste
        Application.CutCopyMode = False
        
        
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[-2],5)"
    Selection.AutoFill Destination:=Range("C1:C400"), Type:=xlFillDefault
           
        
        
        Windows("Cms_uni.xls").Close
        Selection.Replace What:=",000000000", Replacement:="0", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
 


       Sheets("CAPA").Select
       Sheets("CMS_RANKING").Visible = False
       
       
       
   
       
   
       
       
       
       Application.StatusBar = "Atualizado CMS"
Application.CutCopyMode = False
Application.ScreenUpdating = True
       
End Sub


Sub BUSCA_VENDAS()



Application.ScreenUpdating = False

Sheets("Ranking_Operador").Select

Range("H21:H5000").Clear
Range("J21:J5000").Clear
Range("L21:L5000").Clear
Range("N21:N5000").Clear

Dim L, A
Dim VEN_SPD As Variant
Dim VEN_TTD As Variant
Dim VEN_PII As Variant
Dim VEN_3G As Variant
Dim VEN_SUP As Variant


L = 21
A = 2
VEN_SPD = 0
VEN_TTD = 0
VEN_PII = 0
VEN_3G = 0
VEN_SUP = 0

Do While Sheets("Ranking_Operador").Cells(L, 5) <> ""

 Do While Sheets("VANTIVE").Cells(A, 38) <> ""
 
  
  

RE_VANTIVE = Sheets("VANTIVE").Cells(A, 38)
RE_RANKING = Sheets("Ranking_Operador").Cells(L, 5)
  
  



   
   If RE_VANTIVE = RE_RANKING Then
  
  
   If Sheets("VANTIVE").Cells(A, 17) = "SPEEDY" Then
   
      VEN_SPD = VEN_SPD + 1
      
      Sheets("Ranking_Operador").Cells(L, 8) = VEN_SPD
      
      GoTo VAI:
      
      Else
      
   
   If Sheets("VANTIVE").Cells(A, 16) = "TTD" Then

      VEN_TTD = VEN_TTD + 1
      
      Sheets("Ranking_Operador").Cells(L, 10) = VEN_TTD
      
      GoTo VAI:
      
      Else
      
  
   If Sheets("VANTIVE").Cells(A, 16) = "FFE" Then

      VEN_PII = VEN_PII + 1
      
      Sheets("Ranking_Operador").Cells(L, 12) = VEN_PII
      
      GoTo VAI:
        
      Else
     
   If (Sheets("VANTIVE").Cells(A, 16) = "V1G") Or (Sheets("VANTIVE").Cells(A, 16) = "V2G") Then

      VEN_3G = VEN_3G + 1
      
      Sheets("Ranking_Operador").Cells(L, 14) = VEN_3G
      
      GoTo VAI:
    Else
    If Sheets("VANTIVE").Cells(A, 16) = "V250" Then

      VEN_3G = VEN_3G + 1
      
      Sheets("Ranking_Operador").Cells(L, 14) = VEN_3G
      
      GoTo VAI:
    Else
    If Sheets("VANTIVE").Cells(A, 16) = "V4G" Then

      VEN_3G = VEN_3G + 1
      
      Sheets("Ranking_Operador").Cells(L, 14) = VEN_3G
      
      GoTo VAI:
    Else
    If (Sheets("VANTIVE").Cells(A, 16) = "V8G") Or (Sheets("VANTIVE").Cells(A, 16) = "V150") Then

      VEN_3G = VEN_3G + 1
      
      Sheets("Ranking_Operador").Cells(L, 14) = VEN_3G
      
      GoTo VAI:
    Else
    
    If Sheets("VANTIVE").Cells(A, 16) = "SUP" Then

      VEN_SUP = VEN_SUP + 1
      
      Sheets("Ranking_Operador").Cells(L, 16) = VEN_SUP
      
      GoTo VAI:
    
    
   L = L + 1
    
   End If
   End If

   End If

   End If

   End If

   End If
    
   End If
   End If
   End If
VAI:
 
A = A + 1

   Loop


L = L + 1
A = 2
VEN_SPD = 0
VEN_TTD = 0
VEN_PII = 0
VEN_3G = 0
VEN_SUP = 0

Application.StatusBar = "CONTABILIZADOS" & " " & L & " " & "OPERADORES"


  Loop
Application.CutCopyMode = False

Call operadores
End Sub


Sub BUSCA_CHAMADAS()
      
      
      Application.ScreenUpdating = False
      
      
      
      Sheets("Ranking_Operador").Select

       Dim L, A
       Dim LOGIN_RANKING As Variant
       Dim LOGIN_CMS As Variant
       
       L = 21
       A = 6
       
       
       
       Do While Sheets("Ranking_Operador").Cells(L, 5) <> ""
       
       Do While Sheets("CMS_RANKING").Cells(A, 1) <> ""
       
       
       LOGIN_RANKING = Sheets("Ranking_Operador").Cells(L, 6)
       LOGIN_CMS = Sheets("CMS_RANKING").Cells(A, 1)
       
       
       
       
       If LOGIN_RANKING = LOGIN_CMS Then
       
       Sheets("Ranking_Operador").Cells(L, 7) = Sheets("CMS_RANKING").Cells(A, 2)
       
     L = L + 1
     A = 6
     
     Application.StatusBar = "CONTABILIZADOS" & " " & L & " " & "LOGINS"
     
     
     GoTo VAI:
     
     
       End If
       
     A = A + 1
       
       Loop
       
     L = L + 1
     A = 6
     
     Application.StatusBar = "CONTABILIZADOS" & " " & L & " " & "LOGINS"
     
     
VAI:
       
       Loop
 Application.CutCopyMode = False
End Sub

Sub operadores()


supervisor = 10
LRS = 11

Sheets("Ranking|Supervisores").Select

Do While Sheets("Ranking|Supervisores").Cells(supervisor, 3) <> ""
                
        Sheets("Ranking|Supervisores").Range(Cells(LRS, 3), Cells(LRS + 49, 9)).Select
        Selection.ClearContents
        Selection.ClearContents
        
        supervisor = supervisor + 51
        LRS = supervisor + 1
Loop

LRO = 21
supervisor = 10
LRS = 11

Do While Sheets("Ranking|Supervisores").Cells(supervisor, 3) <> ""
                
            If Sheets("Ranking_Operador").Cells(LRO, 3).Value = Sheets("Ranking|Supervisores").Cells(supervisor, 3).Value Then
            
                Sheets("Ranking|Supervisores").Cells(LRS, 3) = Sheets("Ranking_Operador").Cells(LRO, 4)
                Sheets("Ranking|Supervisores").Cells(LRS, 4) = Sheets("Ranking_Operador").Cells(LRO, 7)
                Sheets("Ranking|Supervisores").Cells(LRS, 5) = Sheets("Ranking_Operador").Cells(LRO, 8)
                Sheets("Ranking|Supervisores").Cells(LRS, 6) = Sheets("Ranking_Operador").Cells(LRO, 10)
                Sheets("Ranking|Supervisores").Cells(LRS, 7) = Sheets("Ranking_Operador").Cells(LRO, 12)
                Sheets("Ranking|Supervisores").Cells(LRS, 8) = Sheets("Ranking_Operador").Cells(LRO, 14)
                Sheets("Ranking|Supervisores").Cells(LRS, 9) = Sheets("Ranking_Operador").Cells(LRO, 16)
            
            'Sheets("Ranking_GERAL").Select
            'Sheets("Ranking_GERAL").Range("D" & LRO & ":M" & LRO).Select
            'Selection.Copy
            
            'Sheets("Ranking_Supervisores").Select
            'Sheets("Ranking_Supervisores").Cells(LRS, 3).Select
            'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
            Application.CutCopyMode = False
            
                LRS = LRS + 1
                LRO = LRO + 1
        
        Else
        
            LRO = LRO + 1
        End If
        
        If Sheets("Ranking_Operador").Cells(LRO, 5) = "" Then
        
            LRO = 21
            supervisor = supervisor + 51
            LRS = supervisor + 1
            
        End If
        
Application.StatusBar = "Supervisor" & " " & Sheets("Ranking|Supervisores").Cells(supervisor, 3).Value
        
        Loop


End Sub

Sub operador()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim supervisor
Dim LRS, LRO, LRC


supervisor = 10
LRS = 11
LRC = 16

Sheets("Ranking|Supervisores").Select

Do While Sheets("Ranking|Supervisores").Cells(supervisor, 3) <> ""
                
        Sheets("Ranking|Supervisores").Range(Cells(LRS, 3), Cells(LRS + 49, 16)).Select
        Selection.ClearContents
                
        supervisor = supervisor + 51
        LRS = supervisor + 1
Loop

LRO = 21
supervisor = 10
LRS = 11

Do While Sheets("Ranking|Supervisores").Cells(supervisor, 3) <> ""
                
            
                            
            If Sheets("Ranking_Operador").Cells(LRO, 3).Value = Sheets("Ranking|Supervisores").Cells(supervisor, 3).Value Then
                
                Sheets("Ranking_Operador").Cells(LRC, 5).Value = Sheets("Ranking_Operador").Cells(LRO, 5).Value
            
               
                Sheets("Ranking|Supervisores").Range(Cells(LRS, 3), Cells(LRS, 16)).Value = Sheets("Ranking_Operador").Range("D16:Q16").Value
        
                
            Application.CutCopyMode = False
            
                LRS = LRS + 1
                LRO = LRO + 1
        
        Else
        
            LRO = LRO + 1
        End If
        
        If Sheets("Ranking_Operador").Cells(LRO, 5) = "" Then
        
            LRO = 21
            supervisor = supervisor + 51
            LRS = supervisor + 1
            
        End If
        
Application.StatusBar = "Supervisor" & " " & Sheets("Ranking|Supervisores").Cells(supervisor, 3).Value
        
        Loop


End Sub

Sub ranking_novo() 'modificado dia 20/11/2011

Application.Calculation = xlAutomatic
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim LR, LS, LRO, LVD 'LR'linha Ranking e 'LS'linha Supervisor LRO LINHA RANKING OPERADOR LVD linha de vendas
Dim SomaVD 'coloka um numero para somase das chamadas e vendas

On Error GoTo cont

    Sheets("Ranking|Supervisores").Select
    Rows("10:10").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Selection.Rows.Ungroup
cont:
    Selection.EntireRow.Hidden = False
    Selection.Font.Bold = False
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

LR = 10
LS = 5
LVD = LR

Sheets("Ranking|Supervisores").Cells(LR, 3).Value = Sheets("ARRUMAR").Cells(LS, 6).Value
Sheets("Ranking|Supervisores").Cells(LR, 2).Value = "x"
    
    Range(Cells(LR, 3), Cells(LR, 16)).Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With

LR = 11
LRO = 21
SomaVD = 1

Do While Sheets("ARRUMAR").Cells(LS, 6).Value <> ""

    If Sheets("Ranking_Operador").Cells(LRO, 3).Value = Sheets("ARRUMAR").Cells(LS, 6).Value Then
        
        Sheets("Ranking_Operador").Cells(16, 5).Value = Sheets("Ranking_Operador").Cells(LRO, 5).Value

        Sheets("Ranking|Supervisores").Range(Cells(LR, 3), Cells(LR, 16)).Value = Sheets("Ranking_Operador").Range("D16:Q16").Value
        
        Sheets("Ranking|Supervisores").Cells(LR, 2).Value = SomaVD
                
            Application.CutCopyMode = False
                
        
                LR = LR + 1
                LRO = LRO + 1
        
        Else
        
            LRO = LRO + 1
        End If
        
        If Sheets("Ranking_Operador").Cells(LRO, 5) = "" Then
                         
    Cells(LVD, 6).Select
    ActiveCell.FormulaR1C1 = "=IF(SUMIF(C2,R[1]C2,C)=0,""-"",SUMIF(C2,R[1]C2,C))"
    Selection.AutoFill Destination:=Range(Cells(LVD, 6), Cells(LVD, 11)), Type:=xlFillDefault
    Cells(LVD, 12).Select
    ActiveCell.FormulaR1C1 = "=IFERROR(SUM(RC[-5]:RC[-3])/RC[-6],""-"")"
    Cells(LVD, 13).Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-6]/RC6,""-"")"
    Selection.AutoFill Destination:=Range(Cells(LVD, 13), Cells(LVD, 16)), Type:=xlFillDefault
                             
                             
            LRO = 21
            LS = LS + 1
            LR = LR + 3
            SomaVD = SomaVD + 1
            
    Rows(LVD + 1 & ":" & LR - 1).Select
    Selection.Rows.Group
        
            LVD = LR
            
            
Sheets("Ranking|Supervisores").Cells(LR, 3).Value = Sheets("ARRUMAR").Cells(LS, 6).Value
Sheets("Ranking|Supervisores").Cells(LR, 2).Value = "x"
    
    Range(Cells(LR, 3), Cells(LR, 16)).Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
             
LR = LR + 1
             
        End If
        
Application.StatusBar = "Supervisor" & " " & Sheets("Ranking|Supervisores").Cells(LVD, 3).Value
        
        Loop

    ActiveSheet.Outline.ShowLevels RowLevels:=1
    Range("A10").Select

End Sub
