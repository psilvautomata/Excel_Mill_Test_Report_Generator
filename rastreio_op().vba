Sub rastreio_op()

Application.ScreenUpdating = False
Application.EnableEvents = False

    Worksheets("Consulta").Activate
    Range("G2").Copy
    Worksheets("Soufer").Activate
    Range("R6").PasteSpecial
    Sheets("Consulta").Activate
    Range("H4:H1000").Copy
    Range("AE4").PasteSpecial
    Range("AE4:AE1000").RemoveDuplicates Columns:=1, Header:=xlNo
    Range("AE4:AE1000").Sort key1:=Range("AE4"), order1:=xlAscending
    Range("AE4:AE1000").Copy
    'ActiveSheet.PivotTables("Tabela dinâmica6").PivotSelect "OP[All]", xlLabelOnly, True
    'Selection.SpecialCells(xlCellTypeConstants).Select
    Sheets("Soufer").Activate
    Range("X3").Activate
    ActiveCell.FormulaR1C1 = "=UNIQUE('Consulta'!R[4]C[-16]:R[200]C[-16])"
    Range("X3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("E5").Select
    
Application.ScreenUpdating = True
Application.EnableEvents = True
    
End Sub
