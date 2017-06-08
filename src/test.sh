#!/bin/bash
dotnet restore
dotnet run -p Exceltk/Exceltk.csproj -t md -xls test/test1.xlsx
dotnet run -p Exceltk/Exceltk.csproj -t md -xls test/test2.xlsx
dotnet run -p Exceltk/Exceltk.csproj -t md -xls test/test3.xlsx
dotnet run -p Exceltk/Exceltk.csproj -t md -xls test/test4.xlsx
dotnet run -p Exceltk/Exceltk.csproj -t md -xls test/test5.xlsx
dotnet run -p Exceltk/Exceltk.csproj -t md -xls test/test6_crossline.xlsx
dotnet run -p Exceltk/Exceltk.csproj -t md -xls test/test7_large_tail_empty_columns.xlsx
dotnet run -p Exceltk/Exceltk.csproj -t md -xls test/test8.xls
dotnet run -p Exceltk/Exceltk.csproj -t md -bhead -xls test/test10_bodyhead.xlsx 
dotnet run -p Exceltk/Exceltk.csproj -t md -bhead -xls test/test9_formula.xlsx 
dotnet run -p Exceltk/Exceltk.csproj -t md -bhead -xls test/test11_form.xls
