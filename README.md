Excel toolkit.
Table SHOULD be edited by advanced GUI applications, BUT converted to any other format. 

# Features
  - Convert Excel sheet to MarkDown Table
  - HyperLink cell in Excel sheet will be **retained** as `[text](url)` format 
  - CrossLine cell in Excel sheet will be **expanded** to multirow
  - Empty columns on the right side will be **trimed**, which is **detected** by the first 100 rows. 
  - [New]: Support parse sheet to markdown inner clipboard directly.

# Useage:
  - `exceltk.exe -t md -xls example.xls` 
  - `exceltk.exe -t md -xls example.xls -sheet sheetname`
  - `exceltk.exe -t md -xls example.xlsx` 
  - `exceltk.exe -t md -xls example.xlsx -sheet sheetname`
  - `exceltk.exe -t cm`, Now you can copy sheet from excel, then paster to any editor, which will be Markdown table.

# Download:
  - [exceltk0.0.5 debug version](https://github.com/fanfeilong/exceltk/blob/master/pub/exceltk.0.0.5.7z)

