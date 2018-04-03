# ExcelToJsonConverter
Converts all data within a given worksheet into a json document. 

Workflow:
1. There are 3 parameters:
- SpreadsheetPath: The full path to the excel spreadsheet
- WorkSheetIndex: The index of the worksheet (Starting index is 1, not 0)
- OutputPath: The full path where to output the json document
2. When reading the spreadsheet (Using the SpreadsheetPath and WorkSheetIndex values), it will consider the first row as headers for each column. If it encounters a null it will stop collecting the headers and set the column index at that point (To ensure that all other rows match correctly to the header). The column data for each row after will be collected and matched to its header and stored into a single json object (i.e. JObject). 

Once all rows have been collected they will written to the json document using the OutputPath provided.

A sample spreadsheet and output file have been provided in './source/Samples'
