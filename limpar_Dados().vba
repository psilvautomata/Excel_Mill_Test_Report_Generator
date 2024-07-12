Sub limpar_Dados()
'
' Clear data
'

'
Application.ScreenUpdating = False
Application.EnableEvents = False

Range("E5").Select
Selection.ClearContents 'INVOICE FIELD
Range("T8:V8").ClearContents 'MATERIAL FIELD
Range("R6:V6").ClearContents  'PRODUCTION ORDER FIELD
Range("I11:V14").ClearContents 'CHEMICAL COMPOSITION FIELD
Range("C17:H20").ClearContents 'MECHANICAL PROPERTIES FIELD
Range("S23:V27").ClearContents 'ADDITIONAL INFORMATION FIELD
Range("Y3:Y6").ClearContents 'LOTES MP FIELD
Range("H8").Value = "KG" 'RETURN UNIT VALUE TO DEFAULT
Range("X3:X32").ClearContents 'LOTES SOUFER FIELD
Range("AA3:AA6").ClearContents 'LOT MATERIALS FIELD
Range("F3").Value = Sheets("Dados").Range("B1").Value 'DEFAULT MATERIAL VALUE FOR PIPE
Range("M8").Value = "NBR 6591" 'DEFAULT VALUE FOR STANDARD
Range("AX1").Value = Range("R3").Value 'SET PARAMETER AS CURRENT INVOICE NUMBER
Range("R3").Value = Range("AX1") + 1 'CURRENT INVOICE NUMBER + 1
Rows(6).RowHeight = 18
Columns("A:V").ColumnWidth = 6.57 'DEFAULT COLUMN WIDTH
Columns("W:W").ColumnWidth = 1 'DEFAULT COLUMN WIDTH
Range("E5:G5").Select 'CONVENIENTLY SELECT INVOICE CELL

Call FormulaLotes 'RESET LOT FORMULAS

Application.ScreenUpdating = True
Application.EnableEvents = True

    
End Sub
