# Format plugins

ExcelTk conversion outputs are **format plugins**. Each plugin implements:

- **Export**: `DataSet` → one or more artifacts (text or binary)
- **Import**: artifact stream → `DataSet` → written as `.xlsx`

Built-in plugins:

| `-t` | Extension | Notes |
| --- | --- | --- |
| `md` | `.md` | Pipe / MultiMarkdown tables; embeds `<!-- EXCELTK1 … -->` for precise import |
| `json` | `.json` | `TableDocument` JSON with `marker=EXCELTK1` |
| `tex` | `.tex` | TeX `tabular`; embeds `% EXCELTK1 …` |
| `img` | `.png` | Preview image + private PNG chunk `tkXl`; **unmarked PNGs are rejected** |

## CLI

```bash
# export
exceltk -t md -xls book.xlsx
exceltk -t img -xls book.xlsx -sheet Sheet1

# import back to Excel
exceltk -t md -import bookSheet1.md -out restored.xlsx
exceltk -t img -import bookSheet1.png -out restored.xlsx
```

## Extending

1. Implement `Exceltk.Format.IFormatPlugin`
2. Register in `FormatRegistry.BuiltInPlugins()` (or call `FormatRegistry.Register` at startup)

Precise round-trip data uses `TableDocument` (`marker=EXCELTK1`). The image plugin stores `TableDocument.ToMarkedBytes()` inside the `tkXl` chunk so visual pixels are not the source of truth.
