# Streaming package / progressive render

ExcelTk's binary and OpenXML readers are built around **packages**: complete local
format units that a parser can emit as soon as enough input has arrived.

## Mental model

```
network / file Stream
        │
        ▼
  PackageParser   ── recognizes a finished local unit ──►  Package
        │                                                    │
        │ (continue)                                         ▼
        │                                              Renderer / DataSet builder
```

Examples of a “finished local unit”:

| Format | Package | Completeness boundary |
| --- | --- | --- |
| BIFF (.xls) | `XlsBiffRecord` (: `BinaryPackage`) | one record header (`type`+`size`) + body |
| OpenXML sheet | `XmlCellPackage` / `XmlRowPackage` | finished `</c>` / `</row>` |
| OpenXML | `XmlMergePackage` / `XmlHyperlinkPackage` | finished empty element |

A progressive renderer can subscribe to `IPackageParser<T>.Parse()` (or
`IPullPackageParser.TryRead`) and paint each package without waiting for EOF.

## What this codebase does today

- **Binary**: `BinaryPackageParser` slices **one BIFF record** into an owned byte
  buffer per emit. OLE sector assembly still materializes the workbook stream
  (compound-file seeks); record emission itself is one-package-at-a-time.
- **OpenXML**: `ZipWorker` keeps the ZIP open and **streams entry inflate**
  (`GetInputStream`) instead of extracting the whole archive to disk.
  `XmlPackageParser` emits row/cell/merge/hyperlink packages from `XmlReader`.
- **Batch path**: markdown/json/tex still build a `DataSet` by consuming the
  package stream (convenient, not required for progressive UI).

## Non-seekable network ZIP

Fully forward-only OPC (single `ZipInputStream` pass over a non-seekable network
body) is a follow-up: sheet parts are not contiguous with workbook.xml. Seekable
streams (files, buffered downloads) are supported via `ZipFile` entry streaming.

## TCP demo

Run a loopback server/client that streams sheet XML and BIFF packages:

```bash
dotnet run --project src/Exceltk/Exceltk.csproj -- -t tcpstream \
  -xlsx src/test/test1.xlsx -biff src/test/test8.xls
```

(or from `src/`: `dotnet run --project Exceltk/Exceltk.csproj -- -t tcpstream`)

The receiver prints packages as they complete (`[xml-stream] row#…`, `[biff-stream] #…`)
and exits with `TCP stream test OK` / `Done!` on success. Wired into `src/test.sh`.
