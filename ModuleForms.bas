Attribute VB_Name = "ModuleForms"
Sub CreateFormFieldsForHighlightedGroups()

    Dim doc As Document
    Dim searchRange As Range, foundRange As Range
    Dim ff As FormField
    Dim defaultText As String
    
    Set doc = ActiveDocument
    Set searchRange = doc.Content

    With searchRange.Find
        .ClearFormatting
        .Highlight = True
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop

        While .Execute
            Set foundRange = searchRange.Duplicate
            defaultText = foundRange.Text
            Set ff = doc.FormFields.Add(Range:=foundRange, Type:=wdFieldFormTextInput)
            ff.Result = defaultText
            ff.Range.HighlightColorIndex = wdNoHighlight
            searchRange.Start = ff.Range.End + 1
        Wend
    End With

    ' Protect the document for form filling
    doc.Protect Type:=wdAllowOnlyFormFields, NoReset:=True

    MsgBox "Form created successfully!"

End Sub

Sub DuplicateOperationTable()

    Dim numOfOperations As Integer
    Dim doc As Document
    Dim tbl As Table
    Dim rng As Range
    Dim i As Integer

    ' Reference the active document
    Set doc = ActiveDocument

    ' Unprotect the document
    doc.Unprotect

    ' Check if the document has tables
    If doc.Tables.count = 0 Then
        MsgBox "No tables found in the document!"
        Exit Sub
    End If
    
    ' Reference the first table (assuming it's your template)
    Set tbl = doc.Tables(1)

    ' Ask user for the number of operations
    numOfOperations = InputBox("Enter the number of operations:", "Number of Operations")

    ' Loop to copy and paste the table
    For i = 1 To numOfOperations
        If i > 1 Then
            ' Duplicate the table and maintain form fields
            DuplicateTableWithFormFields tbl
            
            ' Set the tbl object to the last table for the next iteration
            Set tbl = doc.Tables(doc.Tables.count)
        End If
        
        ' Update the operation number of the table
        tbl.Cell(1, 1).Range.Text = "Opération n°" & i
    Next i

    ' Protect the document again for form filling
    doc.Protect Type:=wdAllowOnlyFormFields, NoReset:=True

    ' Inform the user
    MsgBox numOfOperations & " tables generated successfully!"
End Sub

Sub DuplicateTableWithFormFields(tbl As Table)
    Dim rngCopy As Range
    Dim rngPaste As Range
    Dim formFieldsPositions() As Long
    Dim i As Integer
    Dim count As Integer
    Dim newFieldRange As Range
    
    ' Capture positions of form fields
    count = 0
    For Each ff In tbl.Range.FormFields
        ReDim Preserve formFieldsPositions(count)
        formFieldsPositions(count) = ff.Range.Start - tbl.Range.Start
        count = count + 1
    Next ff

    ' Copy and paste the table
    Set rngCopy = tbl.Range
    rngCopy.Copy
    Set rngPaste = rngCopy.Duplicate
    rngPaste.Collapse wdCollapseEnd
    rngPaste.Paste
    
    ' Recreate the form fields in the new table
    For i = 0 To count - 1
        Set newFieldRange = rngPaste.Tables(1).Range.Characters(formFieldsPositions(i)).Duplicate
        ActiveDocument.FormFields.Add Range:=newFieldRange, Type:=wdFieldFormTextInput
    Next i
End Sub

Sub ExportFieldsToExcel()
    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object
    Dim doc As Document
    Dim ff As FormField
    Dim rngContext As Range
    Dim i As Integer
    Dim startContext As Long
    Dim endContext As Long
    Const ContextLength As Long = 40
    Dim msgResponse As VbMsgBoxResult

    ' Inform the user about Excel instance
    msgResponse = MsgBox("This operation will open a new Excel file and close the existing ones automatically to keep only one active. Make sure to close any other Excel files you have open. Continue?", vbYesNo + vbQuestion, "Confirmation")
    
    If msgResponse = vbNo Then
        MsgBox "Operation cancelled.", vbExclamation, "Operation Cancelled"
        Exit Sub
    End If

    ' Close existing Excel application if any
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If
    On Error GoTo 0

    ' Create a new Excel application
    Set xlApp = CreateObject("Excel.Application")
    Set xlWorkbook = xlApp.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets(1)

    ' Reference the active document
    Set doc = ActiveDocument

    ' Set column widths
    xlWorksheet.Columns("A:C").ColumnWidth = 30

    ' Add header to the Excel
    xlWorksheet.Cells(1, 1).Value = "Field Name"
    xlWorksheet.Cells(1, 2).Value = "Field Value"
    xlWorksheet.Cells(1, 3).Value = "Context"

    ' Loop through each form field in the Word document
    i = 2
    For Each ff In doc.FormFields
        ' Get the context around the form field
        startContext = ff.Range.Start - ContextLength
        If startContext < 0 Then startContext = 0

        endContext = ff.Range.End + ContextLength
        If endContext > doc.Content.End Then endContext = doc.Content.End

        Set rngContext = doc.Range(startContext, endContext)

        ' Write data to Excel
        xlWorksheet.Cells(i, 1).Value = ff.Name
        xlWorksheet.Cells(i, 2).Value = ff.Result
        xlWorksheet.Cells(i, 3).Value = rngContext.Text
        i = i + 1
    Next ff

    ' Make Excel visible to the user
    xlApp.Visible = True
End Sub


Sub ImportFieldsFromExcel()
    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object
    Dim doc As Document
    Dim ff As FormField
    Dim i As Integer
    Dim lastRow As Long

    Set doc = ActiveDocument

    ' Open an existing Excel application
    Set xlApp = GetObject(, "Excel.Application")
    Set xlWorkbook = xlApp.ActiveWorkbook
    Set xlWorksheet = xlWorkbook.ActiveSheet

    ' Determine the last row with data in the Excel worksheet
    lastRow = xlWorksheet.Cells(xlWorksheet.Rows.count, 1).End(-4162).Row ' -4162 is xlUp

    ' Loop through each row in the Excel worksheet
    For i = 1 To lastRow
        For Each ff In doc.FormFields
            If ff.Name = xlWorksheet.Cells(i, 1).Value Then
                ff.Result = xlWorksheet.Cells(i, 2).Value
                Exit For
            End If
        Next ff
    Next i

    ' Inform the user
    MsgBox "Data imported successfully!"
End Sub




