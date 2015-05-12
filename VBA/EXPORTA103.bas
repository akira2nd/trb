Attribute VB_Name = "EXPORTA"
Sub EXPORTAR()


  Dim L, hora, L2, L3, L4

hora = Sheets("CAPA").Cells(5, 13)

  For L = 23 To 41

  If hora = Sheets("CAPA").Cells(L, 2) Then


'ILHA SJC
Sheets("CAPA").Cells(L, 3) = Sheets("CAPA").Cells(17, 22)
Sheets("CAPA").Cells(L, 4) = Sheets("CAPA").Cells(17, 23)
Sheets("CAPA").Cells(L, 5) = Sheets("CAPA").Cells(17, 24)
Sheets("CAPA").Cells(L, 6) = Sheets("CAPA").Cells(17, 25)
Sheets("CAPA").Cells(L, 7) = Sheets("CAPA").Cells(17, 26)
Sheets("CAPA").Cells(L, 8) = Sheets("CAPA").Cells(17, 27)
Sheets("CAPA").Cells(L, 9) = Sheets("CAPA").Cells(17, 28)
Sheets("CAPA").Cells(L, 10) = Sheets("CAPA").Cells(17, 29)
Sheets("CAPA").Cells(L, 10) = Sheets("CAPA").Cells(17, 30)

'Assertividade sem lirp/lirw/lire

    Sheets("CAPA").Cells(L, 14) = Sheets("CAPA").Cells(17, 34)
    Sheets("CAPA").Cells(L, 15) = Sheets("CAPA").Cells(17, 35)
    
'3G
    Sheets("CAPA").Cells(L, 16) = Sheets("CAPA").Cells(17, 37)
'FWT
    Sheets("CAPA").Cells(L, 17) = Sheets("CAPA").Cells(17, 38)
        
  End If


  Next


'********************************************************************

  For L2 = 46 To 64

  If hora = Sheets("CAPA").Cells(L2, 2) Then


'''''''''''''ILHA STO
Sheets("CAPA").Cells(L2, 3) = Sheets("CAPA").Cells(18, 22)
Sheets("CAPA").Cells(L2, 4) = Sheets("CAPA").Cells(18, 23)
Sheets("CAPA").Cells(L2, 5) = Sheets("CAPA").Cells(18, 24)
Sheets("CAPA").Cells(L2, 6) = Sheets("CAPA").Cells(18, 25)
Sheets("CAPA").Cells(L2, 7) = Sheets("CAPA").Cells(18, 26)
Sheets("CAPA").Cells(L2, 8) = Sheets("CAPA").Cells(18, 27)
Sheets("CAPA").Cells(L2, 9) = Sheets("CAPA").Cells(18, 28)
Sheets("CAPA").Cells(L2, 10) = Sheets("CAPA").Cells(18, 29)
Sheets("CAPA").Cells(L2, 10) = Sheets("CAPA").Cells(18, 30)

''''''''''''Assertividade sem lirp/lirw/lire

Sheets("CAPA").Cells(L2, 14) = Sheets("CAPA").Cells(18, 34)
Sheets("CAPA").Cells(L2, 15) = Sheets("CAPA").Cells(18, 35)


  End If


  Next


'***********************************************************************


  For L3 = 69 To 87

  If hora = Sheets("CAPA").Cells(L3, 2) Then


'ILHA GERAL

Sheets("CAPA").Cells(L3, 3) = Sheets("CAPA").Cells(19, 22)
Sheets("CAPA").Cells(L3, 4) = Sheets("CAPA").Cells(19, 23)
Sheets("CAPA").Cells(L3, 5) = Sheets("CAPA").Cells(19, 24)
Sheets("CAPA").Cells(L3, 6) = Sheets("CAPA").Cells(19, 25)
Sheets("CAPA").Cells(L3, 7) = Sheets("CAPA").Cells(19, 26)
Sheets("CAPA").Cells(L3, 8) = Sheets("CAPA").Cells(19, 27)
Sheets("CAPA").Cells(L3, 9) = Sheets("CAPA").Cells(19, 28)
Sheets("CAPA").Cells(L3, 10) = Sheets("CAPA").Cells(19, 29)
Sheets("CAPA").Cells(L3, 10) = Sheets("CAPA").Cells(19, 30)

'Assertividade sem lirp/lirw/lire

Sheets("CAPA").Cells(L3, 14) = Sheets("CAPA").Cells(19, 34)
Sheets("CAPA").Cells(L3, 15) = Sheets("CAPA").Cells(19, 35)

'3G
Sheets("CAPA").Cells(L3, 16) = Sheets("CAPA").Cells(19, 37)
'FWT
Sheets("CAPA").Cells(L3, 17) = Sheets("CAPA").Cells(19, 38)

  End If


  Next
  
'vendas separadas 30 minutos
For L4 = 91 To 127
  
  If Sheets("CAPA").Cells(L4, 1).Value = "J" Then
        
        Sheets("CAPA").Range(Cells(L4, 3), Cells(L4 + 1, 4)).Value = Sheets("CAPA").Range("AN17:AO18").Value
        
  End If
Next


    Application.CutCopyMode = False
End Sub

