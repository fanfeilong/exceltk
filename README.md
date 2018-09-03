Excel toolkit. [![Build Status](https://travis-ci.org/fanfeilong/exceltk.svg?branch=master)](https://travis-ci.org/fanfeilong/exceltk)
Table SHOULD be edited by advanced GUI applications, BUT converted to any other format. 


# Convert Excel sheet to MarkDown Table
  - HyperLink cell in Excel sheet will be **retained** as `[text](url)` format 
  - CrossLine cell in Excel sheet will be **expanded** to multirow
  - Empty columns on the right side will be **trimed**, which is **detected** by the first 100 rows. 
  - Support set the precision of decimal
  - Support to set markdown table aligin
  - Convert newline in cell text into `<br/>`
  - Cross sheet Hyperlink formula support, link formula like `HYPERLINK(test_sheet!C9,...)` will be extract as `[text](url)` format automatic
  - Hyperlink formula support, link formula like `HYPERLINK(C9,...)` will be extract as `[text](url)` format automatic

### Usage:
  - `exceltk.exe -t md -xls example.xls` 
  - `exceltk.exe -t md -xls example.xls -sheet sheetname`
  - `exceltk.exe -t md -xls example.xlsx` 
  - `exceltk.exe -t md -xls example.xlsx -sheet sheetname`
  - `exceltk.exe -t md -p 2 -xls example.xls`, where `-p 2` setting the decimal precision to 2
  - `exceltk.exe -t md -bhead -xls example.xls`, which will use the first row to replace table header, and keep the head empty, so that 
  the table will auto response in small screen device, this is just a simply solution.
  - `exceltk.exe -t cm`, Now you can copy sheet from excel, then paster to any editor, which will be Markdown table.
  - `exceltk -t md -a r -xls example.xlsx`, where the `-a` option can be followd by a aligin character
    - `-a l`: aligin left
    - `-a r`: aligin right
    - `-a c`: aligin center

# Convert Excel to Json 
  chagne the `-t` option to `json`
  - `exceltk.exe -t json -xls example.xls `

# Convert Excel to TeX
  change the `-t` option to `tex`
  - `exceltk.exe -t tex -xls example.xls`
  - using `-st n` option to split table into multitable
  - using `-sn` option to adjust number, for example, `1234656` will be split into `1 2 3 4 5 6`, it the table width is too large, this is useful

# Download:

## 0.1.3
  - mac: https://github.com/fanfeilong/exceltk/blob/master/pub/exceltk.0.1.3.pkg
  - windows: http://files.cnblogs.com/files/math/exceltk.0.1.3.zip

## 0.0.9 for windows
  - http://fanfeilong.github.io/exceltk0.0.9.7z
  - http://files.cnblogs.com/files/math/exceltk0.0.9.7z


# 3rd projects

ExcelTk integrated the following projects
- [Excel Data Reader](https://github.com/ExcelDataReader/ExcelDataReader)
- [SharpZip](https://github.com/icsharpcode/SharpZipLib)

# How to build

## Build on MacOS
1. install .NET Core SDK 2.0.0-preview1-005977 
2. cd to the project dir
3. run the following script step by step.
  - `dotnet restore src/exceltk.sln`
  - `dotnet build src/exceltk.sln` 
  - `dotnet run --project src/Exceltk/Exceltk.csproj -t md -xls src/test/test1.xlsx`
4. the `dotnet restore`, this will take long time to install nupack files for publish target runtime, you can comment the following config in `src/Exceltk/Exceltk.csproj` to ignore it.
```
  <PropertyGroup>
    <RuntimeIdentifiers>win-x86;osx.10.10-x64</RuntimeIdentifiers>
  </PropertyGroup>
```
5. run the following script to publish 
```
dotnet publish -r osx.10.10-x64 src/exceltk.sln -c Release
```

## Build on Windows
1. you can also build for windows with the .NET Core SDK, and publish it
```
dotnet publish -r win-x86 src/exceltk.sln -c Release
```
2. you can also build with the visual studio by load the `src/exceltk_vs.sln`


## Build on Linux (example by ubuntu-x64)
1. install .NET Core SDK 
2. cd to the project dir
3. append the `ubuntu-x64` to following RuntimeIdentifiers in `src/Exceltk/Exceltk.csproj`. you can find other RuntimeIdentifiers at: https://docs.microsoft.com/en-us/dotnet/core/rid-catalog
```
  <PropertyGroup>
    <RuntimeIdentifiers>win-x86;osx.10.10-x64;ubuntu-x64</RuntimeIdentifiers>
  </PropertyGroup>
```
4. run the following script step by step.
  - `dotnet restore src/exceltk.sln`
  - `dotnet build src/exceltk.sln` 
  - `dotnet run --project src/Exceltk/Exceltk.csproj -t md -xls src/test/test1.xlsx`

5. run the following script to publish 
```
dotnet publish -r ubuntu-x64 src/exceltk.sln -c Release
```



