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
