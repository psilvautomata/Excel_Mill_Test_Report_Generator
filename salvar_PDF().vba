Sub salvar_PDF()

'SAVE FILE AS PDF

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Sheets("Soufer").Activate 'ACTIVATE SHEET SOUFER

folder = "\Your\Path" 'PDF SAVE PATH
certificate_number = Sheets("Soufer").Range("R3").value 'CERTIFICATE NUMBER
client = Sheets("Soufer").Range("E4").value 'CLIENT NAME
invoice_number = Sheets("Soufer").Range("E5").value 'INVOICE NUMBER
proper_case_client_name = StrConv(client, vbProperCase) 'CONVERT STRING TO PROPER CASE
file_name = "Certificado_" & certificate_number & " - " & proper_case_client_name & " " & invoice_number & ".pdf" 'CERTIFICATE FILE NAME FORMAT

final_name = folder & file_name 'FINAL FILE PATH

'Application.Wait (Now + TimeValue("0:00:02")) 'WAIT FOR 2 SECONDS

Application.Dialogs(xlDialogPrinterSetup).Show
    
Dim response As VbMsgBoxResult

response = MsgBox("Deseja salvar o arquivo?", vbOKCancel + vbQuestion, "Salvar Arquivo") 'ASK USER TO SAVE THE FILE
    
If response = vbOK Then
        ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, PrToFileName:=final_name, IgnorePrintAreas:=False

        MsgBox "Salvo com sucesso!" + Chr(13) + Chr(13) & file_name 'SUCCESS MESSAGE
Else
        MsgBox "Cancelado.", vbInformation, "Operação Cancelada" 'CANCEL MESSAGE

End If
    
End Sub
