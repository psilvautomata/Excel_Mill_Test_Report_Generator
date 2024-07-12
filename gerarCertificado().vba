Sub gerarCertificado()

Application.ScreenUpdating = False
Application.EnableEvents = False

Dim GerarCertificadoAtiva As Boolean

GerarCertificadoAtiva = True

Workbooks.Open Filename:="\Your\Path" 'OPENS THE DATABASE FILE

Workbooks("Controle de Certificados.xlsm").Activate 'ACTIVATES THE SHEET IN THE BD_CERTIFICADOS FILE

max = Worksheets("Soufer").Range("AC3").value 'SETS MAX VALUE WHICH IS EQUAL TO THE NUMBER OF ENTRIES IN MP BATCH

Range("I11:V14").ClearContents 'CLEARS CHEMICAL COMPOSITION DATA
Range("C17:H20").ClearContents 'CLEARS MECHANICAL PROPERTIES DATA
'Range("AA3:AA6").ClearContents 'CLEARS BATCH MATERIALS DATA

For i = 11 To max
    
    Workbooks("Controle de Certificados.xlsm").Activate 'ACTIVATES CERTIFICATES CONTROL FILE
    varLote = Worksheets("Soufer").Range("C" & i).value 'GETS BATCH VALUES
    
    Workbooks("BD_Certificados.xlsm").Activate 'RETURNS TO DATABASE---------------------------------------------|
    Worksheets("Dados").Activate 'ACTIVATES DATA SHEET                                                          |
    Worksheets("Dados").Range("A2").value = varLote 'INSERTS THE COPIED VALUE IN THE PREVIOUS STEP IN CELL A2 TO GET COMPOSITION
    Worksheets("Dados").Range("B2:O2").Select 'SELECTS CHEMICAL COMPOSITION DATA FOR THE INSERTED BATCH---------|
    Selection.Copy '                                                                                            |
    '                                                                                                           |
    Mat = Worksheets("Dados").Range("S2").value 'MATERIAL LINKING FOR CELL S2                                   |
    LE = Worksheets("Dados").Range("Q2").value 'YIELD STRENGTH LINKING FOR CELL Q2                              |
    LR = Worksheets("Dados").Range("R2").value 'TENSILE STRENGTH LINKING FOR CELL R2                            |
    Along = Worksheets("Dados").Range("P2").value 'ELONGATION LINKING FOR CELL P2                               |
    '                                                                                                           |
    Workbooks("Controle de Certificados.xlsm").Activate 'ACTIVATES CERTIFICATES CONTROL FILE                    |
    Worksheets("Soufer").Range("I" & i).Select 'SELECTS THE INFORMATION DROP LOCATION                           |
    '                                                                                                           |
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False 'PASTES THE COPIED CHEMICAL COMPOSITION VALUES--------------------------------|
       
    Worksheets("Soufer").Range("C" & i + 6).value = Along 'DROP LOCATION FOR ELONGATION LINKING
    Worksheets("Soufer").Range("E" & i + 6).value = LE 'DROP LOCATION FOR YIELD STRENGTH LINKING
    Worksheets("Soufer").Range("G" & i + 6).value = LR 'DROP LOCATION FOR TENSILE STRENGTH LINKING
    Worksheets("Soufer").Range("T8").value = Mat 'DROP LOCATION FOR MATERIAL LINKING
    'Worksheets("Soufer").Range("AA" & i - 8).Value = Mat 'DROP LOCATION FOR MATERIAL FOR ALL BATCHES
    
Next

Workbooks("BD_Certificados.xlsm").Close SaveChanges:=False 'CLOSES BD_CERTIFICADOS SHEET WITHOUT SAVING ANYTHING

Application.ScreenUpdating = True
Application.EnableEvents = True

GerarCertificadoAtiva = False

MsgBox ("Data imported successfully!") 'MESSAGE

End Sub
