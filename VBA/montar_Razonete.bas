Attribute VB_Name = "montar_Razonete"
Sub razonete()

Application.ScreenUpdating = False

Dim movimento As Long
Dim debCr As String, nContas As String
Dim lprinc As Integer, lsec As Integer, cR As Integer, lRD As Integer, lrC As Integer, lF As Integer
Dim parar As Integer

Dim lB As Integer

parar = 0
lprinc = 2
lsec = lprinc
cR = 3
lRD = 3
lrC = 3
lB = 10

Sheets("Balancete").Select
Rows("10:30").Select
Selection.ClearContents

Sheets("Lancamentos").Range("K:K").ClearContents
Sheets("Razonete").Select
Cells.Select
Selection.Clear

Do While parar <> 2
    If Sheets("Lancamentos").Cells(lprinc, 11).Value = "ok" Then
        lprinc = lprinc + 1
        parar = 0
    Else
    If (Sheets("Lancamentos").Cells(lprinc, 9).Value = "") Then
        parar = parar + 1
        lprinc = lprinc + 1
    Else
        lsec = lprinc
        parar = 0
        nContas = Right(Sheets("Lancamentos").Cells(lsec, 9).Value, Len(Sheets("Lancamentos").Cells(lsec, 9).Value) - 4)
        Do While parar <> 2
            If Sheets("Lancamentos").Cells(lsec, 9).Value = "" Then
                parar = parar + 1
                lsec = lsec + 1
            Else
                parar = 0
                If (Sheets("Lancamentos").Cells(lsec, 11).Value <> "ok") And (nContas = Right(Sheets("Lancamentos").Cells(lsec, 9).Value, Len(Sheets("Lancamentos").Cells(lsec, 9).Value) - 4)) Then
                    movimento = Sheets("Lancamentos").Cells(lsec, 10).Value
                    debCr = Left(Sheets("Lancamentos").Cells(lsec, 9).Value, 1)
                    If debCr = "D" Then
                        Sheets("Razonete").Cells(3, cR).Value = nContas
                        Range(Cells(3, cR), Cells(3, cR + 1)).Select
                        With Selection
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlBottom
                            .WrapText = False
                            .Orientation = 0
                            .AddIndent = False
                            .IndentLevel = 0
                            .ShrinkToFit = False
                            .ReadingOrder = xlContext
                            .MergeCells = True
                        End With
                        With Selection.Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        
                        Sheets("Razonete").Cells(lRD + 1, cR).Value = movimento
                        Cells(lRD + 1, cR).Select
                        With Selection.Borders(xlEdgeRight)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        
                        lRD = lRD + 1
                        Sheets("Lancamentos").Cells(lsec, 11).Value = "ok"
                    Else
                        Sheets("Razonete").Cells(3, cR).Value = nContas
                        Range(Cells(3, cR), Cells(3, cR + 1)).Select
                        With Selection
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlBottom
                            .WrapText = False
                            .Orientation = 0
                            .AddIndent = False
                            .IndentLevel = 0
                            .ShrinkToFit = False
                            .ReadingOrder = xlContext
                            .MergeCells = True
                        End With
                        With Selection.Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With

                        Sheets("Razonete").Cells(lrC + 1, cR + 1).Value = movimento
                        Cells(lrC + 1, cR + 1).Select
                        With Selection.Borders(xlEdgeLeft)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        
                        lrC = lrC + 1
                        Sheets("Lancamentos").Cells(lsec, 11).Value = "ok"
                    End If
                End If
                lsec = lsec + 1
            End If
        Loop
        
        lF = Application.Max(lrC, lRD)
        
''''''''''''''''''''''''''balancete nome''''''''''''''''''''''''''''''''
        Sheets("Balancete").Cells(lB, 4).Value = nContas
''''''''''''''''''''''''''balancete fim'''''''''''''''''''''''''''''''''''''

        Range(Cells(lF + 1, cR), Cells(lF + 1, cR + 1)).Select
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        Cells(lF + 1, cR).Value = Application.Sum(Range(Cells(4, cR), Cells(lF, cR)))
        Cells(lF + 1, cR + 1).Value = Application.Sum(Range(Cells(4, cR + 1), Cells(lF, cR + 1)))
''''''''''''''''''''''''''balancete movimento''''''''''''''''''''''''''''''''
        Sheets("Balancete").Cells(lB, 5).Value = Application.Sum(Range(Cells(4, cR), Cells(lF, cR)))
        Sheets("Balancete").Cells(lB, 6).Value = Application.Sum(Range(Cells(4, cR + 1), Cells(lF, cR + 1)))
''''''''''''''''''''''''''balancete fim'''''''''''''''''''''''''''''''''''''
        Cells(lF + 1, cR + 1).Select
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        
        If Cells(lF + 1, cR).Value > Cells(lF + 1, cR + 1).Value Then
            Cells(lF + 2, cR).Value = (Cells(lF + 1, cR).Value - Cells(lF + 1, cR + 1).Value)
''''''''''''''''''''''''''balancete positivo''''''''''''''''''''''''''''''''
            Sheets("Balancete").Cells(lB, 7).Value = (Cells(lF + 1, cR).Value - Cells(lF + 1, cR + 1).Value)
''''''''''''''''''''''''''balancete fim'''''''''''''''''''''''''''''''''''''
        Else
            Cells(lF + 2, cR + 1).Value = (Cells(lF + 1, cR + 1).Value - Cells(lF + 1, cR).Value)
''''''''''''''''''''''''''balancete negativo''''''''''''''''''''''''''''''''
            Sheets("Balancete").Cells(lB, 8).Value = (Cells(lF + 1, cR + 1).Value - Cells(lF + 1, cR).Value)
''''''''''''''''''''''''''balancete fim'''''''''''''''''''''''''''''''''''''
        End If

        Range(Cells(lF + 2, cR), Cells(lF + 2, cR + 1)).Select
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        lprinc = lprinc + 1
        lRD = 3
        lrC = 3
        cR = cR + 3
        parar = 0
        lB = lB + 1
    End If
    End If
Loop

Application.ScreenUpdating = True

End Sub
