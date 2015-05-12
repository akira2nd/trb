Attribute VB_Name = "Ranking"

Sub operadores()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim supervisor
Dim LRS, LBR, LRC


supervisor = 10
LRS = 11
LRC = 21

Sheets("Ranking|Supervisores").Select

Do While Sheets("Ranking|Supervisores").Cells(supervisor, 11) <> ""
                
        Sheets("Ranking|Supervisores").Range(Cells(LRS, 11), Cells(LRS + 49, 38)).Select
        Selection.ClearContents
                
        supervisor = supervisor + 51
        LRS = supervisor + 1
Loop

LBR = 26
supervisor = 10
LRS = 11

Do While Sheets("Ranking|Supervisores").Cells(supervisor, 11) <> ""
                
            
                            
            If Sheets("BASE_RANKING").Cells(LBR, 7).Value = Sheets("Ranking|Supervisores").Cells(supervisor, 11).Value Then
                
                Sheets("BASE_RANKING").Cells(LRC, 9).Value = Sheets("BASE_RANKING").Cells(LBR, 9).Value
            
               
                Sheets("Ranking|Supervisores").Range(Cells(LRS, 11), Cells(LRS, 38)).Value = Sheets("BASE_RANKING").Range("H21:AI21").Value
        
                
            Application.CutCopyMode = False
            
                LRS = LRS + 1
                LBR = LBR + 1
        
        Else
        
            LBR = LBR + 1
        End If
        
        If Sheets("BASE_RANKING").Cells(LBR, 9) = "" Then
        
            LBR = 26
            supervisor = supervisor + 51
            LRS = supervisor + 1
            
        End If
        
Application.StatusBar = "Supervisor" & " " & Sheets("Ranking|Supervisores").Cells(supervisor, 11).Value
        
        Loop


End Sub


Sub arruma_super()

Dim supervisor, arruma


supervisor = 10
arruma = 5

Sheets("Ranking|Supervisores").Select
    
    Range("K10:M65100").Select
    Selection.ClearContents
    
    Rows("10:65100").Select
    Selection.Rows.Ungroup

    Selection.EntireRow.Hidden = False


Do While Sheets("ARRUMAR").Cells(arruma, 6).Value <> ""


    Sheets("Ranking|Supervisores").Select
   
    Sheets("Ranking|Supervisores").Cells(supervisor, 11).Value = Sheets("ARRUMAR").Cells(arruma, 6).Value
    Sheets("Ranking|Supervisores").Cells(supervisor, 10).Value = "x"


    Rows(supervisor + 1 & ":" & supervisor + 50).Select
    Selection.Rows.Group
    
    arruma = arruma + 1
    supervisor = supervisor + 51
    

    
    Loop

ActiveSheet.Outline.ShowLevels RowLevels:=1

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

Sheets("Ranking|Supervisores").Cells(LR, 11).Value = Sheets("ARRUMAR").Cells(LS, 6).Value
Sheets("Ranking|Supervisores").Cells(LR, 10).Value = "x"
    
    Range(Cells(LR, 11), Cells(LR, 43)).Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With

LR = 11
LRO = 26
SomaVD = 1

''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sheets("BASE_RANKING").Select
    
    Columns("A:D").Select
    Selection.ClearContents

    Range("G25:J25").Select
    Selection.AutoFilter
    ActiveSheet.Range("$G$25:$J$8024").AutoFilter Field:=1, Criteria1:=Sheets("ARRUMAR").Cells(LS, 6).Value
    Range("G25").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy

    Range("A25").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("G25").Select
    Selection.AutoFilter

Sheets("Ranking|Supervisores").Select
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Do While Sheets("ARRUMAR").Cells(LS, 6).Value <> ""

'1 - 7
    If Sheets("BASE_RANKING").Cells(LRO, 1).Value = Sheets("ARRUMAR").Cells(LS, 6).Value Then
        '3 - 9
        Sheets("BASE_RANKING").Cells(21, 9).Value = Sheets("BASE_RANKING").Cells(LRO, 3).Value
        
        'Calculate
        Sheets("Ranking|Supervisores").Range(Cells(LR, 11), Cells(LR, 42)).Value = Sheets("BASE_RANKING").Range("H21:AM21").Value
        
        Sheets("Ranking|Supervisores").Cells(LR, 10).Value = SomaVD
                
            Application.CutCopyMode = False
                
        
                LR = LR + 1
                LRO = LRO + 1
        
        Else
        
            LRO = LRO + 1
        End If
        
        If Sheets("BASE_RANKING").Cells(LRO, 3) = "" Then

    Cells(LVD, 14).Select
    ActiveCell.FormulaR1C1 = "=IF(SUMIF(C10,R[1]C10,C)=0,""-"",SUMIF(C10,R[1]C10,C))"
    Selection.AutoFill Destination:=Range(Cells(LVD, 14), Cells(LVD, 38)), Type:=xlFillDefault
    'Cells(LVD, 12).Select
    'ActiveCell.FormulaR1C1 = "=IFERROR(SUM(RC[-5]:RC[-3])/RC[-6],""-"")"
    Cells(LVD, 39).Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-24]/RC14,""-"")"
    Cells(LVD, 40).Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-16]/RC14,""-"")"
    Selection.AutoFill Destination:=Range(Cells(LVD, 40), Cells(LVD, 42)), Type:=xlFillDefault
                             
                             
            LRO = 26
            LS = LS + 1
            LR = LR + 3
            SomaVD = SomaVD + 1
            
    Rows(LVD + 1 & ":" & LR - 1).Select
    Selection.Rows.Group
        
            LVD = LR
            
            
Sheets("Ranking|Supervisores").Cells(LR, 11).Value = Sheets("ARRUMAR").Cells(LS, 6).Value
Sheets("Ranking|Supervisores").Cells(LR, 10).Value = "x"
    
    Range(Cells(LR, 11), Cells(LR, 43)).Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
             
LR = LR + 1

''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sheets("BASE_RANKING").Select
    
    Columns("A:D").Select
    Selection.ClearContents

    Range("G25:J25").Select
    Selection.AutoFilter
    ActiveSheet.Range("$G$25:$J$8024").AutoFilter Field:=1, Criteria1:=Sheets("ARRUMAR").Cells(LS, 6).Value
    Range("G25").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy

    Range("A25").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("G25").Select
    Selection.AutoFilter


Sheets("Ranking|Supervisores").Select
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''


             
        End If
        
Application.StatusBar = "Supervisor" & " " & Sheets("Ranking|Supervisores").Cells(LVD, 11).Value
        
        Loop

    ActiveSheet.Outline.ShowLevels RowLevels:=1
    Range("A10").Select

End Sub


