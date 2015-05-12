Attribute VB_Name = "CMS"
Sub CMS_G()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim caminho, L, PASTA, ARQUIVO, ORIGEM, DESTINO
Dim L1


ORIGEM = ActiveWorkbook.Name

L = 16

Do While Sheets("PREMISSAS").Cells(L, 10) <> ""


caminho = Sheets("PREMISSAS").Cells(17, 2)
ARQUIVO = Sheets("PREMISSAS").Cells(L, 9)
PASTA = Sheets("PREMISSAS").Cells(L, 10)

Sheets(PASTA).Visible = True

Workbooks.OpenText Filename:=caminho & ARQUIVO, Origin _
    :=xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
    xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
    Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
    Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
    Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15 _
    , 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array(29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array(42, 1), Array(43, 1), Array(44, 1), Array(45, 1), Array(46, 1), Array(47, 1), Array(48, 1), Array(49, 1)), TrailingMinusNumbers:=True
    
    
DESTINO = ActiveWorkbook.Name

Cells.Select
Cells.Copy

Windows(ORIGEM).Activate
Sheets(PASTA).Select
Range("A1").Select
Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        Selection.Replace What:=",000000000", Replacement:="0", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
                
                Windows(DESTINO).Activate
                Windows(DESTINO).Close
       
       
        If PASTA <> "VDN_TRANSFER" Then
                    For L1 = 10 To 70
                    
                        If Cells(L1, 1).Value = Cells(7, 1).Value Then
                            Cells(L1 - 1, 1).Select
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
            Application.CutCopyMode = False
            L = L + 1
            
            Loop

   
             Sheets("CAPA").Select

Call CMS_RANKING
Application.CutCopyMode = False
End Sub
