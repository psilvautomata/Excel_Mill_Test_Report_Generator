Sub limpar_Parcial()
'
' Clear data and keep the same invoice
'

'
Dim max As Long
Dim sts As Long
Dim value As Integer
Dim i As Variant
Dim a As Variant
Dim resuCell As Variant

Application.ScreenUpdating = False
Application.EnableEvents = False

Sheets("Soufer").Activate

max = Range("AF9").Value

a = Range("J6").Value
    
If WorksheetFunction.CountIf(Sheets("Soufer").Range("AK9:AK38"), a) > 0 Then
    Set resuCell = Range("AK9:AK38").Find(What:=a, LookIn:=xlValues, LookAt:=xlWhole)
    resuCell.Select
    sts = ActiveCell.Offset(0, 1).Value
    
End If

If sts < max Then

    Range("AK" & sts + 9).Copy
    Range("J6").PasteSpecial
    
    Range("R6:V6").ClearContents 'PRODUCTION ORDER FIELD
    Range("I11:V14").ClearContents 'CHEMICAL COMPOSITION FIELD
    Range("X3:X32").ClearContents 'LOTES SOUFER FIELD
    Range("Y3:Y6").ClearContents 'LOTES MP FIELD
    Range("S23:V27").ClearContents 'ADDITIONAL INFORMATION FIELD
    Range("C17:H20").ClearContents 'MECHANICAL PROPERTIES FIELD
    Range("T8:V8").ClearContents 'MATERIAL FIELD
    Range("AA3:AA6").ClearContents 'LOT MATERIALS FIELD
    Range("AX1").Value = Range("R3").Value 'SET PARAMETER AS CURRENT INVOICE NUMBER
    Range("R3").Value = Range("AX1") + 1 'CURRENT INVOICE NUMBER + 1
    Columns("A:V").ColumnWidth = 6.57 'DEFAULT COLUMN WIDTH
    Columns("W:W").ColumnWidth = 1 'DEFAULT COLUMN WIDTH
    Rows(6).RowHeight = 18
    Range("E5:G5").Select 'CONVENIENTLY SELECT INVOICE CELL
    
    Call rastreio_op
    Call concat_sub
    Call rastreio_op
    
Else
    Call limpar_Dados
    
End If

Call FormulaLotes 'RESET BATCHES FORMULAS

Application.ScreenUpdating = True
Application.EnableEvents = True


End Sub
