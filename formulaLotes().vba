Sub FormulaLotes()
'
Dim i As Integer

Application.ScreenUpdating = False
Application.EnableEvents = False

For i = 21 To 27
    ' Fills cells from B to D in row i
    Range("B" & i & ":D" & i).FormulaR1C1 = "=IF(R[-18]C[22]="""","""",R[-18]C[22])"
    
    ' Fills cells from E to G in row i
    Range("E" & i & ":G" & i).FormulaR1C1 = "=IF(R[-11]C[19]="""","""",R[-11]C[19])"
    
    ' Fills cells from H to J in row i
    Range("H" & i & ":J" & i).FormulaR1C1 = "=IF(R[-4]C[16]="""","""",R[-4]C[16])"
    
    ' Fills cells from K to M in row i
    Range("K" & i & ":M" & i).FormulaR1C1 = "=IF(R[3]C[13]="""","""",R[3]C[13])"

Next i

Range("S22:V22").Select 'RESET GRAMMER FORMULA
ActiveCell.FormulaR1C1 = _
    "=IFERROR(IF(R[-18]C[-14]=""GRAMMER DO BRASIL LTDA"",IF(R[-15]C[-13]="""","""",R[24]C[-15]),""""),"""")"
Range("AA3").Select
ActiveCell.FormulaR1C1 = _
    "=IF(RC[-2]="""","""",IFERROR(XLOOKUP(RC[-2],Dados!C[-19],Dados!C[-18]),""""))"

'Materials of the lots
Range("AA4").Select
ActiveCell.FormulaR1C1 = _
    "=IF(RC[-2]="""","""",IFERROR(XLOOKUP(RC[-2],Dados!C[-19],Dados!C[-18]),""""))"
Range("AA5").Select
ActiveCell.FormulaR1C1 = _
    "=IF(RC[-2]="""","""",IFERROR(XLOOKUP(RC[-2],Dados!C[-19],Dados!C[-18]),""""))"
Range("AA6").Select
ActiveCell.FormulaR1C1 = _
    "=IF(RC[-2]="""","""",IFERROR(XLOOKUP(RC[-2],Dados!C[-19],Dados!C[-18]),""""))"

'Additional fields
Range("O22:R22").Select
ActiveCell.FormulaR1C1 = _
    "=IF(OR(R[-18]C[-10]=""WORK ELETRO SISTEMAS IND COM E"",R[-18]C[-10]=""METALURGICA FORMIGARI LTDA""),""Customer Code"",IF(OR(R[-18]C[-10]=""INDUSTRIA METALURGICA A PEDRO LTDA"",R[-18]C[-10]=""MGK SOLUCOES INDUSTRIAIS""),""Coating (g/m²)"",IF(R[-18]C[-10]=""GRAMMER DO BRASIL LTDA"",""Customer Code"",IF(R3C6=Dados!R6C2,""Rev Point (g/m²)"",""""))))"
Range("O23:R23").Select
ActiveCell.FormulaR1C1 = _
    "=IF(R[-19]C[-10]=""MGK SOLUCOES INDUSTRIAIS"",""Sup"",IF(R[-19]C[-10]=""WORK ELETRO SISTEMAS IND COM E"",""Purchase Order"",IF(R3C6=Dados!R6C2,""Rev Average (g/m²)"",IF(R[-19]C[-10]=""GRAMMER DO BRASIL LTDA"",""Hardness (HRB)"",""""))))"
Range("O24:R24").Select
ActiveCell.FormulaR1C1 = _
    "=IF(R[-20]C[-10]=""MGK SOLUCOES INDUSTRIAIS"",""Inf"",IF(R3C6=Dados!R6C2,""Total Resin (mg/m²)"",""""))"
Range("O25:R25").Select
ActiveCell.FormulaR1C1 = _
    "=IF(R[-21]C[-10]=""MGK SOLUCOES INDUSTRIAIS"",""Total"",IF(R3C6=Dados!R6C2,""Hardness (HRB)"",""""))"

'Invoice fields
Range("E4:V4").Select
ActiveCell.FormulaR1C1 = _
    "=IF(R[10]C[25]="""","""",IFERROR(VLOOKUP(R[10]C[25],Clientes!C[-3]:C[1],2,0),""""))"
Range("L5:O5").Select
ActiveCell.FormulaR1C1 = _
    "=IF(RC[-7]="""","""",IFERROR(VLOOKUP(RC[-7]&R[1]C[-2],'Banco de Dados'!C1:C15,5,0),""""))"
Range("U5:V5").Select
ActiveCell.FormulaR1C1 = _
    "=IF(RC[-16]="""","""",IFERROR(VLOOKUP(RC[-16]&R[1]C[-11],'Banco de Dados'!C1:C15,6,0),""""))"
Range("E6:G6").Select
ActiveCell.FormulaR1C1 = _
    "=IF(R[-1]C="""","""",IFERROR(VLOOKUP(R[-1]C&RC[5],'Banco de Dados'!C1:C15,7,0),""""))"
Range("F7:V7").Select
ActiveCell.FormulaR1C1 = _
    "=IF(R[-2]C[-1]="""","""",IFERROR(VLOOKUP(R[-2]C[-1]&R[-1]C[4],'Banco de Dados'!C1:C15,8,0),""""))"
Range("F8:G8").Select
ActiveCell.FormulaR1C1 = _
    "=IF(R[-3]C[-1]="""","""",IFERROR(VLOOKUP(R[-3]C[-1]&R[-2]C[4],'Banco de Dados'!C1:C15,9,0),""""))"
Range("AD3").Select
ActiveCell.FormulaR1C1 = _
    "=IF(R[2]C[-25]="""","""",IFERROR(VLOOKUP(R[2]C[-25]&R[3]C[-20],'Banco de Dados'!C1:C15,14,0),""""))"
Range("AD6").Select
ActiveCell.FormulaR1C1 = _
    "=IF(R[-1]C[-25]="""","""",IFERROR(VLOOKUP(R[-1]C[-25]&RC[-20],'Banco de Dados'!C1:C15,12,0),""""))"
Range("AD12").Select
ActiveCell.FormulaR1C1 = _
    "=IF(R[-3]C[2]=0,"""",IF(R[-3]C[2]=1,CONCAT(R[-3]C[2],"" "",""Item""),CONCAT(R[-3]C[2],"" "",""Items"")))"
Range("AD14").Select
ActiveCell.FormulaR1C1 = _
    "=IF(R[-9]C[-25]="""","""",IFERROR(VLOOKUP(R[-9]C[-25]&R[-8]C[-20],'Banco de Dados'!C1:C15,3,0),""""))"
Range("AA14").Select
ActiveCell.FormulaR1C1 = _
    "=IFERROR(XLOOKUP(R3C30,Clientes_SAP!C[-26],Clientes_SAP!C[-24]),"""")"
Range("Y14").Select
ActiveCell.FormulaR1C1 = _
    "=IF(R[-9]C[7]="""","""",IF(OR(R[-9]C[7]=""6101/AA"",R[-9]C[7]=""5101/AA""),""Direct Sale"",IF(OR(R[-9]C[7]=""6122/AA"",R[-9]C[7]=""5122/AA""),""Triangular Sale"",IF(OR(R[-9]C[7]=""5102/AA"",R[-9]C[7]=""6102/AA""),""Third Party Goods R13"",IF(OR(R[-9]C[7]=""6924/AA"",R[-9]C[7]=""5924/AA""),""Triangular Shipment"",""N/D"")))))"

'Issuance auxiliaries
Range("AF2:AI2").Select
ActiveCell.FormulaR1C1 = _
    "=IF(R14C30="""","""",XLOOKUP(R14C30,Clientes!C2,Clientes!C3))"
Range("AF5:AG5").Select
ActiveCell.FormulaR1C1 = _
    "=IF(RC[-27]="""","""",IFERROR(VLOOKUP(RC[-27]&R[1]C[-22],'Banco de Dados'!C1:C15,13,0),""""))"
Range("B9:V9").Select


'LINES ABOVE RESET ESSENTIAL FORMULAS FOR THE SHEET TO FUNCTION
Range("E5:G5").Select


Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub
