Sub moverMenu()

    Dim ws As Worksheet
    Dim menuDropDown As Shape
    Dim ultimaLinha As Long
    
    For Each ws In ThisWorkbook.Sheets
    
    If ws.Name = "1103" Or ws.Name = "1109" Then

        ' Name of the dropdown menu (replace with the correct name)
        Set menuDropDown = ws.Shapes("Agrupar 1")
    
        ' Find the last filled row in column B
        ultimaLinha = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
        ' Move the dropdown menu to a position below the last filled row
        menuDropDown.Top = ws.Cells(ultimaLinha + 2, "B").Top

    End If
    
    Next

End Sub
