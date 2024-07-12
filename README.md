# $Excel$ $Mill$ $Test$ $Report$ $Generator$

[General_Mill_Report](https://drive.google.com/file/d/1DnFpvdknJpdV1xJ__K8aYQpKMeTcO-5)

[Complete_System_Emission](https://drive.google.com/file/d/1IRkTykcuH0UFWD889sFt-ifwqnv6iaa8)

<p align="Justify"> This workbook generates mill reports from mill's batches informations as chemichal compositions and mechanical properties.</p>

### *gerarCertificado()*

<p align="Justify">Automates the process of importing data from a database Excel file into a certificate control spreadsheet. It starts by disabling screen updating and events to optimize performance. It sets a control variable GerarCertificadoAtiva to True and opens the database file BD_Certificados.xlsm. The code then activates the control spreadsheet Controle de Certificados.xlsm and sets a maximum value based on the number of entries in the MP batch from cell AC3 in the Soufer sheet. It clears specific cells for chemical composition and mechanical properties data.</p>

<p align="Justify">For each entry from row 11 to the maximum value, it activates the control spreadsheet, retrieves the batch values from column C, and switches back to the database file. It activates the Dados sheet, inserts the batch value into cell A2 to fetch the chemical composition data, selects and copies the data, and captures material, yield strength, tensile strength, and elongation values from specific cells.</p>

<p align="Justify">The control spreadsheet is reactivated, the data is pasted into the appropriate cells, and the captured values are placed in their respective cells for elongation, yield strength, tensile strength, and material. After processing all entries, the database file is closed without saving, screen updating and events are re-enabled, and GerarCertificadoAtiva is set to False. Finally, a message box confirms the successful import of data.</p>

### *concat_sub()*

<p align="Justify">Concatenates the contents of cells E5 and J6 from the "Soufer" sheet. It first checks if cell E5 is empty; if it is, it displays an error message and stops execution. Next, it disables screen updating and events to enhance performance, concatenates the values from the specified cells, and places the result in cell B1 of the "Consulta" sheet. After clearing the clipboard and waiting for half a second, the code returns to the "Soufer" sheet and selects cell J6. Finally, it re-enables screen updating and events, indicating that the last change was made in cell J6.</p>

### *rastreio_op()*

<p align="Justify">The code activates the "Consulta" worksheet and copies the value from cell G2, which is subsequently pasted into cell R6 of the "Soufer" worksheet. Returning to the "Consulta" worksheet, the code copies the range H4 and pastes it into AE4. It removes any duplicates from AE4, sorts the data in ascending order, and then copies this range again. The script then activates the "Soufer" worksheet, places a preset formula in cell X3 to extract unique values from 'Consulta'!AE4, and pastes the values (not the formula) into cell X3. Finally, it selects cell E5 in the "Soufer".</p>

### *selecionar_Todos()* and *moverMenu()*

<p align="Justify">The macros filter data in sheets "1103" and "1109" based on various priority criteria. After applying the filters, they call the MoverMenu subroutine, which repositions a dropdown menu named "Agrupar 1" to just below the last filled row in column B.</p>

### *limpar_Dados*

<p align="Justify">Clears various fields in the "Soufer" worksheet, resetting fields such as material, production order, chemical composition, mechanical properties, additional information, and MP and Soufer batches. It also sets default values for unit, material, and standard, updates the invoice number, and resets the formulas for batches.</p>

### *limpar_Parcial*

<p align="Justify">Performs a partial reset, maintaining the same invoice number while clearing similar fields as Limpar_Dados. It checks if the current state is less than the maximum allowed; if so, it copies data from a specified range and clears various fields. If not, it calls Limpar_Dados to perform a full reset.</p>

### *salvar_PDF*

<p align="Justify">It saves the active sheet ("Soufer") as a PDF file. It sets the destination folder path and constructs the PDF file name using the certificate number, client name, and invoice number. It then prompts the user to confirm if they want to save the file. If confirmed, it saves the PDF file to the specified path and displays a success message. If the user cancels, it shows a cancellation message.</p>

### *unirtexto* 

<p align="Justify">It combines data from the "Dados" sheet into the "Soufer" sheet.The macro then activates the "Dados" sheet and copies the Soufer batches numbers from the range L1. Next, it activates the "Soufer" sheet and pastes these values starting at cell X3. After pasting, it clears the contents of the range X32 to ensure no old data remains below the newly pasted values. The macro then resets specific column widths to their default values and selects cell B4 to return the focus. Finally, it re-enables screen updating.</p>

