Excel toolkit. [![Build Status](https://travis-ci.org/fanfeilong/exceltk.svg?branch=master)](https://travis-ci.org/fanfeilong/exceltk)
Table SHOULD be edited by advanced GUI applications, BUT converted to any other format. 


# Convert Excel sheet to MarkDown Table
  - HyperLink cell in Excel sheet will be **retained** as `[text](url)` format 
  - CrossLine cell in Excel sheet will be **expanded** to multirow
  - Empty columns on the right side will be **trimed** (sampled from the first 100 rows for speed; columns that appear later are still kept when encountered during the full read). 
  - Support set the precision of decimal
  - Support to set markdown table aligin
  - Convert newline in cell text into `<br/>`
  - Cross sheet Hyperlink formula support, link formula like `HYPERLINK(test_sheet!C9,...)` will be extract as `[text](url)` format automatic
  - Hyperlink formula support, link formula like `HYPERLINK(C9,...)` will be extract as `[text](url)` format automatic
  - MultiMarkdown mode (`-mmd`): emit HTML tables with `rowspan`/`colspan` for Excel merged cells (pipe Markdown cannot express rowspan)

# Convert CSV to MarkDown Table
  - Same markdown options as Excel (`-p`, `-a`, `-bhead`, `-pretty`)
  - Quoted fields and embedded commas are supported (RFC 4180 style)
  - Input via `-xls file.csv` or `-csv file.csv`

# Pretty MarkDown tables
  - Use `-pretty` to pad cell text and separators so columns line up in the source
  - Works with `-a l|c|r` alignment

# MultiMarkdown merged cells
  - Use `-mmd` to emit an HTML `<table>` with `rowspan`/`colspan` from Excel merge regions
  - HTML tables are valid in MultiMarkdown documents; standard pipe Markdown cannot express rowspan
  - Currently supported for `.xlsx` worksheets that declare `mergeCells`

### Usage:
  - `exceltk.exe -t md -xls example.xls` 
  - `exceltk.exe -t md -xls example.xls -sheet sheetname`
  - `exceltk.exe -t md -xls example.xlsx` 
  - `exceltk.exe -t md -xls example.xlsx -sheet sheetname`
  - `exceltk.exe -t md -xls example.csv`
  - `exceltk.exe -t md -csv example.csv`
  - `exceltk.exe -t md -pretty -xls example.xlsx`, pad columns so the markdown source looks aligned
  - `exceltk.exe -t md -mmd -xls example.xlsx`, MultiMarkdown-friendly HTML table preserving merged cells
  - `exceltk.exe -t md -p 2 -xls example.xls`, where `-p 2` setting the decimal precision to 2
  - `exceltk.exe -t md -bhead -xls example.xls`, which will use the first row to replace table header, and keep the head empty, so that 
  the table will auto response in small screen device, this is just a simply solution.
  - `exceltk -t md -a r -xls example.xlsx`, where the `-a` option can be followd by a aligin character
    - `-a l`: aligin left
    - `-a r`: aligin right
    - `-a c`: aligin center

# Removed: clipboard monitor (`-t cm`)
  - `-t cm` (GUI clipboard watcher) existed only in **0.0.9** on Windows and was **removed** afterwards
  - Download 0.0.9 if you still need it: http://files.cnblogs.com/files/math/exceltk0.0.9.7z
  - Current versions convert files with `-t md|json|tex` instead

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

Requires [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0) (LTS). A `global.json` at the repo root pins SDK 10 with `rollForward: latestMajor`.

Quick check (any OS):

```bash
dotnet build src/Exceltk/Exceltk.csproj -c Release
dotnet run --project src/Exceltk/Exceltk.csproj -c Release -- -t md -xls src/test/test1.xlsx
```

> Tip: put `--` before app args when using `dotnet run`, so options like `-t` / `-a` are not eaten by the `dotnet` CLI.

Optional: if `dotnet restore` feels slow because of publish RIDs, you can temporarily remove or comment the `RuntimeIdentifiers` block in `src/Exceltk/Exceltk.csproj`.

CI: GitHub Actions (`.github/workflows/ci.yml`) builds on .NET 10 and runs `src/test.sh`.

## Build on MacOS
```bash
dotnet publish -r osx-x64 src/Exceltk/Exceltk.csproj -c Release
# Apple Silicon:
dotnet publish -r osx-arm64 src/Exceltk/Exceltk.csproj -c Release
```

## Build on Windows
```bash
dotnet publish -r win-x86 src/Exceltk/Exceltk.csproj -c Release
```
Open `src/exceltk.sln` (SDK-style `Exceltk.csproj`) in Visual Studio / VS Code. The legacy `exceltk_vs.sln` / `Exceltk_vs.csproj` (.NET Framework) is obsolete.

## Build on Linux (example: ubuntu / linux-x64)
1. Install the [.NET 10 SDK](https://learn.microsoft.com/dotnet/core/install/linux)
2. From the repo root:

```bash
dotnet restore src/Exceltk/Exceltk.csproj
dotnet build src/Exceltk/Exceltk.csproj -c Release
dotnet run --project src/Exceltk/Exceltk.csproj -c Release -- -t md -xls src/test/test1.xlsx
```

3. Publish a linux binary:

```bash
dotnet publish -r linux-x64 src/Exceltk/Exceltk.csproj -c Release
# output under: src/bin/net10.0/linux-x64/publish/exceltk
./src/bin/net10.0/linux-x64/publish/exceltk -t md -xls src/test/test1.xlsx
```

`linux-x64` is the portable RID for Ubuntu and most x64 Linux distros. Other RIDs: https://learn.microsoft.com/dotnet/core/rid-catalog



