Sub enviado()

ActiveCell.Offset(0, 4).Select
Selection.value = "Enviado"
ActiveCell.Offset(1, -4).Activate 'Ctr+Q function to "Sent" status

End Sub
