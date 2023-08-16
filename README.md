# Bridge-Report
Takes PDF bridge reports and creates an excel spreadsheet with table data.

Connects to a SharePoint where the PDF's are uploaded into a folder.

Uses pdfPlumber library to extract all data from PDF.
I used this over Tabula or other libraries because of security on my work laptop which wouldn't allow me to add to PATH.
Cycles through each file and extracts table data which is stored to a local excel spreadsheet

Creates a masterlist with the ID and what PDF it references.
