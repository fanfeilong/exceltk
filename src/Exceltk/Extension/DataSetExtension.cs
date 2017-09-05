using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

using Exceltk.Reader;

namespace Exceltk{
    public static class DataSetExtension{
        public static DataTable Shrink(this DataTable dataTable) {
            int columnCount=dataTable.Columns.Count;
            int rowCount=dataTable.Rows.Count;
            for (int r=rowCount-1; r>=0; r--) {
                for (int c=columnCount-1; c>=0; c--) {
                    object cell=dataTable.Rows[r][c];
                    string value="";
                    if (cell!=null) {
                        var xlsCell=cell as XlsCell;
                        if (cell is XlsCell) {
                            value=xlsCell.Value.ToString();
                        } else {
                            value=cell.ToString();
                        }
                    }

                    if (!string.IsNullOrEmpty(value.Trim())) {
                        goto CUT_COLUMN;
                    }
                }

                dataTable.Rows.RemoveAt(r);
            }

            CUT_COLUMN:
            columnCount=dataTable.Columns.Count;
            rowCount=dataTable.Rows.Count;
            for (int c=columnCount-1; c>=0; c--) {
                for (int r=rowCount-1; r>=0; r--) {
                    object cell=dataTable.Rows[r][c];
                    string value="";
                    if (cell!=null) {
                        var xlsCell=cell as XlsCell;
                        if (cell is XlsCell) {
                            value=xlsCell.Value.ToString();
                        } else {
                            value=cell.ToString();
                        }
                    }
                    if (!string.IsNullOrEmpty(value.Trim())) {
                        goto QUIT;
                    }
                }

                dataTable.Columns.RemoveAt(c);
            }

            QUIT:

            return dataTable;
        }

        public static DataTable RemoveColumnsByRow(this DataTable dataTable, int rowIndex, Func<XlsCell, bool> filter) {
            if (rowIndex>=dataTable.Rows.Count) {
                throw new ArgumentOutOfRangeException(string.Format("row index overflow: {0}", dataTable.Rows.Count));
            }
            DataRow row=dataTable.Rows[rowIndex];
            int index=0;
            var removeIndexs=new List<int>();
            foreach (object cell in row.ItemArray) {
                XlsCell value=null;
                var xlsCell=cell as XlsCell;
                if (xlsCell!=null) {
                    value=xlsCell;
                } else {
                    value=new XlsCell(cell);
                }

                if (filter(value)) {
                    removeIndexs.Add(index);
                }

                index++;
            }

            for (int i=removeIndexs.Count-1; i>=0; i--) {
                dataTable.Columns.RemoveAt(removeIndexs[i]);
            }

            return dataTable;
        }

        public static bool IsEmpty(this DataRow row) {
            foreach (object cell in row.ItemArray) {
                string value="";
                var xlsCell=cell as XlsCell;
                if (xlsCell!=null) {
                    value=xlsCell.Value.ToString();
                } else {
                    value=cell.ToString();
                }

                if (!string.IsNullOrEmpty(value.Trim())) {
                    return false;
                }
            }
            return true;
        }

        public static bool IsSingleByteEncoding(this Encoding encoding) {
            return encoding.IsSingleByte;
        }

        public static double Int64BitsToDouble(this long value) {
            return BitConverter.ToDouble(BitConverter.GetBytes(value), 0);
        }

        public static string ConvertEscapeChars(this string input) {
            var re=new Regex("_x([0-9A-F]{4,4})_");
            return re.Replace(input, m => (((char)UInt32.Parse(m.Groups[1].Value, NumberStyles.HexNumber))).ToString());
        }

        public static object ConvertFromOATime(this double value) {
            if ((value>=0.0)&&(value<60.0)) {
                value++;
            }
            return value.FromOADate();
        }

        public static int[] ReferenceToColumnAndRow(this string reference) {
            //split the string into row and column parts
            var matchLettersNumbers=new Regex("([a-zA-Z]*)([0-9]*)");
            string column=matchLettersNumbers.Match(reference).Groups[1].Value.ToUpper();
            string rowString=matchLettersNumbers.Match(reference).Groups[2].Value;

            //.net 3.5 or 4.5 we could do this awesomeness
            //return reference.Aggregate(0, (s,c)=>{s*26+c-'A'+1});
            //but we are trying to retain 2.0 support so do it a longer way
            //this is basically base 26 arithmetic
            int columnValue=0;
            int pow=1;

            //reverse through the string
            for (int i=column.Length-1; i>=0; i--) {
                int pos=column[i]-'A'+1;
                columnValue+=pow*pos;
                pow*=26;
            }

            return new int[] { int.Parse(rowString), columnValue };
        }
    }
}