Excel toolkit. [![CI](https://github.com/fanfeilong/exceltk/actions/workflows/ci.yml/badge.svg)](https://github.com/fanfeilong/exceltk/actions/workflows/ci.yml)
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
  - Current versions convert files with `-t md|json|tex` instead
  - If you still need 0.0.9, look for an archived asset on [Releases](https://github.com/fanfeilong/exceltk/releases) (do not rely on README version matrices)

# Convert Excel to Json 
  change the `-t` option to `json`
  - `exceltk.exe -t json -xls example.xls `
  - JSON is emitted as an ExcelTk `TableDocument` (`marker=EXCELTK1`) so it can be imported back

# Convert Excel to TeX
  change the `-t` option to `tex`
  - `exceltk.exe -t tex -xls example.xls`
  - using `-st n` option to split table into multitable
  - using `-sn` option to adjust number, for example, `1234656` will be split into `1 2 3 4 5 6`, it the table width is too large, this is useful

# Format plugins (export / import)
  Output formats are plugins (`md`, `json`, `tex`, `img`). Each plugin can **export** Excel/CSV and **import** back to `.xlsx`.

  - Export: `exceltk -t <plugin> -xls file.xlsx [-sheet name] [-out prefix]`
  - Import: `exceltk -t <plugin> -import file.ext [-out out.xlsx]`

  Precise payloads use marker `EXCELTK1` (Markdown HTML comment, TeX `%` comment, JSON field, or PNG private chunk `tkXl`).

# Marked image plugin (`-t img`)
  - Export renders a PNG preview and embeds the table payload in a private PNG chunk
  - Import only accepts ExcelTk-marked PNGs; unmarked images are rejected
  - `exceltk -t img -xls example.xlsx`
  - `exceltk -t img -import exampleSheet1.png -out restored.xlsx`

# Download

Prebuilt binaries are published through **GitHub Releases** (CI builds self-contained packages for each RID on version tags `v*`).

- **Latest release**: https://github.com/fanfeilong/exceltk/releases/latest
- **All releases**: https://github.com/fanfeilong/exceltk/releases

Pick the asset that matches your OS:

| Asset RID | Platform |
| --- | --- |
| `linux-x64` | Linux x64 (`.tar.gz`) |
| `osx-x64` | macOS Intel (`.tar.gz`) |
| `osx-arm64` | macOS Apple Silicon (`.tar.gz`) |
| `win-x86` | Windows (`.zip`) |

Example asset name: `exceltk-0.1.4-linux-x64.tar.gz`.

To cut a release: tag `vX.Y.Z` and push; `.github/workflows/release.yml` publishes the assets.


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

Architecture note: binary/OpenXML reading is organized as **streaming packages** (emit a complete local format unit as soon as it is recognized) so a renderer can progress while data is still arriving — see [`docs/streaming-packages.md`](docs/streaming-packages.md).

Output formats are **plugins** with export/import — see [`docs/format-plugins.md`](docs/format-plugins.md).

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



