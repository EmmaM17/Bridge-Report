# Bridge-Report
Takes PDF bridge reports and creates an excel spreadsheet with table data.
--------------------------------------------------------------------------
PDF_to_excelSP connects to a SharePoint where the PDF's are uploaded into a folder.

Uses pdfPlumber library to extract all data from PDF.
I used this over Tabula or other libraries because of security on my work laptop which wouldn't allow me to add to PATH.
Cycles through each file and extracts table data which is stored to a local excel spreadsheet

Creates a masterlist with the ID and the name of what PDF it references.

Creates an online version and uploads it back to SharePoint in an output folder

NLP_SP takes the Bridge Report excel file stores the tables in dataframes.
It also takes the training data and stores that to a dataframe.

Description and catergory data is simplified for the Support Vector Model to process.

Prediction data is created and is checked for it's accuracy. (Can output a confusion matrix)

The real descriptions from the bridge report data frame is used to predict it's categories.
These predictions are then added to the excel spreadsheet.

The data is reuploaded back to SharePoint for it to be used in a PowerBi report that visually represents findings/trends

