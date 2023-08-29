## README for Word-to-Excel Automation
**installation**
1. Enable developer mode in the settings
2. go to VBA
3. right click import module
4. choose the bas file to import it
   
This guide will walk you through preparing a Word document for automation, converting highlighted sections into form fields, duplicating operations, and then exporting/importing these fields to/from Excel for quick data entry.

### Preparing Your Word Document
1. **Highlight the Relevant Sections**: Begin by highlighting the sections of your document in yellow that you wish to convert into form fields. Make sure to write within these highlighted sections, clearly indicating what each section corresponds to.

2. **Create Form Fields**:
    - Run the `Createforms` macro.
    - This will replace all the yellow-highlighted sections with form fields, making them ready for data entry.
    - (Note: This step may be done in advance if preferred.)


https://github.com/yoman38/Word-Forms/assets/124726056/0aed7b46-fcbb-4de4-b559-ce57bb9260e5


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


https://github.com/yoman38/Word-Forms/assets/124726056/6fa89a30-c365-403e-8712-035e275341ab


5. **Import Back to Word**:
    - Ensure the Excel file with the data is still open.
    - Run the `Subimport` macro.
    - This will fetch the data from the opened Excel sheet and populate the form fields in your Word document with the respective values.

https://github.com/yoman38/Word-Forms/assets/124726056/b854d873-766d-42c5-a290-5ea6f8568303


---

Please note: Always remember to save your documents and sheets regularly to prevent any data loss.
