Dim ultimaAlteracaoJ6 As Boolean

Sub concat_sub()

Application.ScreenUpdating = False
Application.EnableEvents = False

Dim x As Variant

Sheets("Soufer").Activate
If Range("E5").value = "" Then
    MsgBox "ERRO - Nota Fiscal Vazia"
    Exit Sub ' Ends the macro if cell E5 is empty
End If

' Disables screen updating and events to improve performance
Application.ScreenUpdating = False
Application.EnableEvents = False

' Activates the "Soufer" sheet
Sheets("Soufer").Activate

' Concatenates the values of cells E5 and J6
x = Range("E5").value & Range("J6").value

' Activates the "Consulta" sheet
Sheets("Consulta").Activate

' Sets the value in cell B1 to the concatenated value
Range("B1").value = x

' Clears the clipboard
Application.CutCopyMode = False

Application.Wait Now + TimeSerial(0, 0, 0.5)

' Returns to the "Soufer" sheet
Sheets("Soufer").Activate
Range("J6").Select
    
' Re-enables screen updating and events
Application.ScreenUpdating = True
Application.EnableEvents = True

' Indicates that the last change was in J6
ultimaAlteracaoJ6 = True

Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub
