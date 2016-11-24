#!/bin/bash
dotnet restore
dotnet run -t md -xls test/test1.xlsx
dotnet run -t md -xls test/test2.xlsx
dotnet run -t md -xls test/test3.xlsx
dotnet run -t md -xls test/test4.xlsx
dotnet run -t md -xls test/test5.xlsx
dotnet run -t md -xls test/test6_crossline.xlsx
dotnet run -t md -xls test/test7_large_tail_empty_columns.xlsx
dotnet run -t md -xls test/test8.xls
dotnet run -t -t md -bhead -xls test\test10_bodyhead.xlsx 
