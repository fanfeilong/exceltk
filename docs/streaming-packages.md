# Package / PackageParser (XUdt-style)

ExcelTk frames Office-defined protocol units — not UI progressive-render DTOs.

## Roles

| Role | Responsibility |
| --- | --- |
| **Package** | One Office unit (BIFF record type / SpreadsheetML element). Owns `EncodeBody` / `DecodeBody` of Office fields. |
| **PackageParser** | `PushData` / `Pump` / `TryFrameOne` state machine: recognize complete unit → empty package by type → `DecodeBody` → queue; `HavePackage` / `PopPackage` / `HaveUnParsedData` / `Reset`. Does not invent field layout. |
| **Reader** | Workbook semantics (SST, styles, DataSet) consuming popped packages / cursor records. |

Wire layouts are dictated by Office Excel — do not change them for streaming convenience.

## Encode / Decode model

Same shape as XUdt:

| Layer | Binary (`BinaryPackage`) | XML (`XmlPackage`) |
| --- | --- | --- |
| Header framing | Base writes/reads 2-byte type + 2-byte body size | Element start tag / framer owns fragment bytes |
| Body | Virtual `EncodeBody` / `DecodeBody` on each concrete package | `EncodeBody(XmlWriter)` / `Decode(XmlReader)` (bytes route via `IPackage.DecodeBody`) |
| Required size | `GetRequiredEncodeBodyBufferLength()` | Encode-to-`MemoryStream` via `EncodeBody(byte[])` helper |

Base only frames the header (or element wrapper). **Every concrete Package implements body encode/decode.** `OnBytesReady` is gone — field parsing lives in `DecodeBody`.

## Binary (.xls)

```
OLE/XlsStream (sector assembly)     ← NOT a Package
        │ workbook byte[]
        ▼
BinaryPackageParser.PushData
        │  CreateEmpty(BIFFRECORDTYPE) → concrete XlsBiff* Package
        ▼
BinaryPackage  (base in Reader.Binary)
  ├── EncodePackage = header + EncodeBody
  ├── DecodePackage / IPackage.DecodeBody = header → DecodeBody(body)
  ├── XlsBiffBOF, XlsBiffRKCell, XlsBiffLabelSSTCell, …
  └── BinaryPackage itself (unknown types: raw body passthrough)
```

**Each `XlsBiff*` record class in `Reader/Binary/` is a Package** (one Office BIFF command type = one class).

| Is a Package | Is NOT a Package |
| --- | --- |
| `BinaryPackage` + every `XlsBiff*` record type | `XlsStream`, `XlsHeader`, `XlsFat`, `XlsRootDirectory` (OLE) |
| | `XlsWorksheet`, `XlsWorkbookGlobals`, `ExcelBinaryReader` (workbook semantics) |
| | `BiffWorkbookCursor` (Seek/ReadAt index helper) |

- BIFF header = 2-byte type + 2-byte body length; body from Office header.
- Factory: `BinaryPackage.CreateEmpty(type)` → typed subclass → `DecodePackage` / `AttachSharedBytes` → `DecodeBody` fills Office fields.
- Every concrete `XlsBiff*` owns **field-level** `DecodeBody` (buffer → members) and `EncodeBody` (members → buffer). No raw-body cheat on known types.
- `BinaryPackage` base `CaptureRawBody` remains only for **unknown** record types.
- Opaque known types (`CONTINUE`, `QUICKTIP`) store `byte[] m_payload` as the Office body member — still field-level.
- **SST:** members `m_count`, `m_uniqueCount`, `m_stringData` (bytes after the 8-byte header). CONTINUE-spanned strings stay Reader-side (`Append` + `ReadStrings`).
- **HyperLink:** range, GUID, flags, optional description/frame blocks, and URL are members; Encode rebuilds that structure.
- Formula string results may look ahead to the next STRING package on the shared buffer.
- `XlsBiffBlankCell` is a mid-base for shared row/col/xf; subclasses call `base.DecodeBody` then extend (except types that rewrite the full body).

## OpenXML (.xlsx)

```
worksheet XML stream (file / NetworkStream / PushData buffer)
        │
        ▼
XmlPackageParser (XmlReader tokenize)
        │  complete SpreadsheetML element
        ▼
XmlPackage.CreateEmpty(localName) → Decode(reader) → queue
        │
        EncodeBody(XmlWriter) writes the same Office element shape
```

Office element packages only:

| Element | Package |
| --- | --- |
| `dimension` | `XmlDimensionPackage` |
| `row` | `XmlRowPackage` (attrs + child `c` cell packages) |
| `c` | `XmlCellPackage` (attrs `r`/`t`/`s`; children `v`/`t`/`f`) |
| `mergeCell` | `XmlMergePackage` |
| `hyperlink` | `XmlHyperlinkPackage` |

`worksheet` / `sheetData` / `mergeCells` / `hyperlinks` are **parser state** — not packages.

`ExcelOpenXmlReader.ReadSheetGlobals` obtains `XmlDimensionPackage` via `XmlPackageParser` when present; sheets without `dimension` still fall back to counting row/cell packages for size.

## TCP demo

```bash
dotnet run --project src/Exceltk/Exceltk.csproj -- -t tcpstream \
  -xlsx src/test/test1.xlsx -biff src/test/test8.xls
```

- **BIFF**: server sends raw workbook stream chunks; client `BinaryPackageParser.PushData` pops concrete `XlsBiff*` packages.
- **XML**: server sends worksheet XML chunks; client `XmlReader` + `XmlPackageParser.TryFrameOne` / `PopPackage`.
