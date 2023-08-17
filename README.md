

## README for Word-to-Excel Automation

This guide will walk you through preparing a Word document for automation, converting highlighted sections into form fields, duplicating operations, and then exporting/importing these fields to/from Excel for quick data entry.

### Preparing Your Word Document
1. **Highlight the Relevant Sections**: Begin by highlighting the sections of your document in yellow that you wish to convert into form fields. Make sure to write within these highlighted sections, clearly indicating what each section corresponds to.

2. **Create Form Fields**:
    - Run the `Createforms` macro.
    - This will replace all the yellow-highlighted sections with form fields, making them ready for data entry.
    - (Note: This step may be done in advance if preferred.)

### Duplicating tables (adjust it if needed):
3. **Choose the Number of duplicates**:
    - Run the `Dupliquer operation` macro.
    - When prompted, enter the number of operations/duplicates you wish to create. This will duplicate the operation tables accordingly in your document.

At this point, you can choose to fill out the form fields manually or automate the process with Excel. it is currently configured for a specific case, and for the first table of the document. 

### Exporting & Importing Data with Excel:
4. **Export to Excel**:
    - Run the `Subexport` macro.
    - This will export the content of the form fields in your Word document into an Excel sheet.
    - You can now quickly enter or modify values in Excel.
    - Save the Excel file to any location but keep it open.
    - The version with context will assist you in understanding the content better. It will also expand the columns for a clearer view.

5. **Import Back to Word**:
    - Ensure the Excel file with the data is still open.
    - Run the `Subimport` macro.
    - This will fetch the data from the opened Excel sheet and populate the form fields in your Word document with the respective values.

---

Please note: Always remember to save your documents and sheets regularly to prevent any data loss.