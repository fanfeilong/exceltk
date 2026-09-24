#!/bin/bash
set -euo pipefail

PROJECT=Exceltk/Exceltk.csproj

dotnet restore "$PROJECT"

run() {
  # App args after -- so dotnet CLI does not consume -t / -a / etc.
  dotnet run --project "$PROJECT" -- "$@"
}

run -t md -xls test/test1.xlsx
run -t md -xls test/test2.xlsx
run -t md -xls test/test3.xlsx
run -t md -xls test/test4.xlsx
run -t md -xls test/test5.xlsx
run -t md -xls test/test6_crossline.xlsx
run -t md -xls test/test7_large_tail_empty_columns.xlsx
run -t md -xls test/test8.xls
run -t md -bhead -xls test/test10_bodyhead.xlsx
run -t md -bhead -xls test/test9_formula.xlsx
run -t md -bhead -xls test/test11_form.xls
run -t md -xls test/test13_csv.csv
run -t md -csv test/test13_csv.csv
run -t md -pretty -xls test/test13_csv.csv
run -t md -pretty -a c -xls test/test13_csv.csv
run -t md -pretty -a r -xls test/test13_csv.csv
run -t md -mmd -xls test/test6_crossline.xlsx
run -t md -xls test/test14_late_columns.xlsx
run -t md -xls test/test10_nostyles.xlsx
run -t md -xls test/test8_issue8_hyperlinks.xlsx
run -t tcpstream -xlsx test/test1.xlsx -biff test/test8.xls

# Format plugins: export/import round-trip (md/json/img) + reject unmarked PNG
PLUGIN_OUT=$(mktemp -d)
trap 'rm -rf "$PLUGIN_OUT"' EXIT
run -t json -xls test/test1.xlsx -sheet Sheet1 -out "$PLUGIN_OUT/sheet"
JSON_FILE=$(ls "$PLUGIN_OUT"/sheet*.json | head -1)
run -t json -import "$JSON_FILE" -out "$PLUGIN_OUT/from-json.xlsx"
run -t img -xls test/test1.xlsx -sheet Sheet1 -out "$PLUGIN_OUT/img"
PNG_FILE=$(ls "$PLUGIN_OUT"/img*.png | head -1)
run -t img -import "$PNG_FILE" -out "$PLUGIN_OUT/from-img.xlsx"
run -t md -xls test/test1.xlsx -sheet Sheet1 -out "$PLUGIN_OUT/md"
MD_FILE=$(ls "$PLUGIN_OUT"/md*.md | head -1)
run -t md -import "$MD_FILE" -out "$PLUGIN_OUT/from-md.xlsx"
# Unmarked PNG must fail import
python3 - "$PLUGIN_OUT/unmarked.png" <<'PY'
import struct, zlib, sys
sig = b"\x89PNG\r\n\x1a\n"
def chunk(t, d):
    return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 6, 0, 0, 0)
raw = b"\x00" + b"\xff\x00\x00\xff"
open(sys.argv[1], "wb").write(sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))
PY
if run -t img -import "$PLUGIN_OUT/unmarked.png" -out "$PLUGIN_OUT/should-fail.xlsx"; then
  echo "ERROR: unmarked PNG import should have failed" >&2
  exit 1
fi
echo "Format plugin round-trip OK"
