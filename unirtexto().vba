Sub unirtexto()
'
' Merge Soufer Batches
'

'
Application.ScreenUpdating = False 'FREEZE SCREEN

Sheets("Dados").Activate 'ACTIVATE SHEET DADOS
Range("L1:L44").Select 'REFERENCE CELL FOR MERGED BATCHES
Selection.Copy 'COPY VALUES
    
Sheets("Soufer").Activate
Range("X3").Select 'PASTE LOCATION
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False 'PASTE THE VALUES OF MERGED BATCHES
Range("X32:X115").Select 'CLEAR VALUES BELOW THE LIMIT ALLOWED BY THE CERTIFICATE MODEL
Application.CutCopyMode = False
Selection.ClearContents
Columns("D:D").ColumnWidth = 8.14 'DEFAULT COLUMN WIDTH
Columns("G:G").ColumnWidth = 8.14 'DEFAULT COLUMN WIDTH
Columns("J:J").ColumnWidth = 8.14 'DEFAULT COLUMN WIDTH
Columns("W:W").ColumnWidth = 3.86 'DEFAULT COLUMN WIDTH
Range("B4").Select 'SELECT B4 TO RETURN
    
Application.ScreenUpdating = True 'UNFREEZE SCREEN

End Sub
