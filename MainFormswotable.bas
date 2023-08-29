Attribute VB_Name = "MainForms"
Sub MainRoutine()

    Dim msgResponse As VbMsgBoxResult

    ' Step 1: Ask the user if they have correctly highlighted the words
    msgResponse = MsgBox("Have you correctly highlighted the words to fill? Make sure to leave space before and after if needed for text comprehension.", vbYesNo + vbQuestion, "Confirmation")
    
    If msgResponse = vbNo Then
        MsgBox "Please highlight the words correctly and run the program again.", vbExclamation, "Operation Cancelled"
        Exit Sub
    End If

    ' Step 2: Run CreateFormFieldsForHighlightedGroups
    CreateFormFieldsForHighlightedGroups
    
    ' Step 4: ExportFieldsToExcel
    ExportFieldsToExcel
    
    ' Pause and wait for user to finish editing Excel file
    msgResponse = MsgBox("Please edit the Excel file and press 'Yes' once you are done.", vbYesNo + vbQuestion, "Edit Excel File")
    If msgResponse = vbNo Then
        MsgBox "Operation cancelled. Please run the program again when you are ready.", vbExclamation, "Operation Cancelled"
        Exit Sub
    End If

    ' Step 5: ImportFieldsFromExcel
    ImportFieldsFromExcel

    ' Final Message
    MsgBox "All operations completed successfully!", vbInformation, "Success"

End Sub




