Excel toolkit. [![Build Status](https://travis-ci.org/fanfeilong/exceltk.svg?branch=master)](https://travis-ci.org/fanfeilong/exceltk)
Table SHOULD be edited by advanced GUI applications, BUT converted to any other format. 

# Features
  - Convert Excel sheet to MarkDown Table
  - HyperLink cell in Excel sheet will be **retained** as `[text](url)` format 
  - CrossLine cell in Excel sheet will be **expanded** to multirow
  - Empty columns on the right side will be **trimed**, which is **detected** by the first 100 rows. 
  - Support parse sheet to markdown inner clipboard directly.
  - Support set the precision of decimal

# Useage:
  - `exceltk.exe -t md -xls example.xls` 
  - `exceltk.exe -t md -xls example.xls -sheet sheetname`
  - `exceltk.exe -t md -xls example.xlsx` 
  - `exceltk.exe -t md -xls example.xlsx -sheet sheetname`
  - `exceltk.exe -t md -p 2 -xls example.xls`, where `-p 2` setting the decimal precision to 2
  - `exceltk.exe -t md -bhead -xls example.xls`, which will use the first row to replace table header, and keep the head empty, so that 
  the table will auto response in small screen device, this is just a simply solution.
  - `exceltk.exe -t cm`, Now you can copy sheet from excel, then paster to any editor, which will be Markdown table.

# Download:
  - http://fanfeilong.github.io/exceltk0.0.7.7z
  - http://files.cnblogs.com/files/math/exceltk0.0.7.7z

# 3rd projects

ExcelTk integrated the following projects
- [Excel Data Reader](https://github.com/ExcelDataReader/ExcelDataReader)
- [SharpZip](https://github.com/icsharpcode/SharpZipLib)
