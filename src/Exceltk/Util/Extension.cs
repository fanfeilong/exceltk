using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelToolKit {
    public static class Extension {
        private static readonly Regex re=new Regex("_x([0-9A-F]{4,4})_");

        public static DataTable Shrink(this DataTable dataTable) {
            int columnCount=dataTable.Columns.Count;
            int rowCount=dataTable.Rows.Count;
            for (int r=rowCount-1; r>=0; r--) {
                for (int c=columnCount-1; c>=0; c--) {
                    object cell=dataTable.Rows[r][c];
                    string value="";
                    var xlsCell=cell as XlsCell;
                    if (cell is XlsCell) {
                        value=xlsCell.Value.ToString();
                    } else {
                        value=cell.ToString();
                    }
                    if (value!=null&&!string.IsNullOrEmpty(value.Trim())) {
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
                    var xlsCell=cell as XlsCell;
                    if (cell is XlsCell) {
                        value=xlsCell.Value.ToString();
                    } else {
                        value=cell.ToString();
                    }
                    if (value!=null&&!string.IsNullOrEmpty(value.Trim())) {
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
                throw new ArgumentOutOfRangeException(string.Format("行下标超出范围，最大行数为： {0}", dataTable.Rows.Count));
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

                if (value!=null&&!string.IsNullOrEmpty(value.Trim())) {
                    return false;
                }
            }
            return true;
        }


        public static MarkDownTable ToMd(this string xls, string sheet) {
            FileStream stream=File.Open(xls, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader=null;
            if (Path.GetExtension(xls)==".xls") {
                excelReader=ExcelReaderFactory.CreateBinaryReader(stream);
            } else if (Path.GetExtension(xls)==".xlsx") {
                excelReader=ExcelReaderFactory.CreateOpenXmlReader(stream);
            } else {
                throw new ArgumentException("Not Support Format: ");
            }
            DataSet dataSet=excelReader.AsDataSet();
            DataTable dataTable=dataSet.Tables[sheet];

            var table=new MarkDownTable {
                Name=dataTable.TableName,
                Value=dataTable.ToMd()
            };

            excelReader.Close();

            return table;
        }

        public static IEnumerable<MarkDownTable> ToMd(this string xls) {
            FileStream stream=File.Open(xls, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader=null;
            if (Path.GetExtension(xls)==".xls") {
                excelReader=ExcelReaderFactory.CreateBinaryReader(stream);
            } else if (Path.GetExtension(xls)==".xlsx") {
                excelReader=ExcelReaderFactory.CreateOpenXmlReader(stream);
            } else {
                throw new ArgumentException("Not Support Format: ");
            }
            DataSet dataSet=excelReader.AsDataSet();

            foreach (DataTable dataTable in dataSet.Tables) {
                var table=new MarkDownTable {
                    Name=dataTable.TableName,
                    Value=dataTable.ToMd()
                };

                yield return table;
            }

            excelReader.Close();
        }

        private static string ToMd(this DataTable table) {
            table.Shrink();
            //table.RemoveColumnsByRow(0, string.IsNullOrEmpty);
            var sb=new StringBuilder();

            int i=0;
            foreach (DataRow row in table.Rows) {
                //if (row.IsEmpty())
                //{
                //    continue;
                //}

                sb.Append("|");
                foreach (object cell in row.ItemArray) {
                    string value="";
                    var xlsCell=cell as XlsCell;
                    if (xlsCell!=null) {
                        value=xlsCell.MarkDownText;
                    } else {
                        value=cell.ToString();
                    }

                    sb.Append(value).Append("|");
                }
                sb.Append("\r\n");
                if (i==0) {
                    sb.Append("|");
                    foreach (DataColumn col in table.Columns) {
                        sb.Append(":--|");
                    }
                    sb.Append("\r\n");
                }
                i++;
            }
            return sb.ToString();
        }

        public static bool IsSingleByteEncoding(this Encoding encoding) {
            return encoding.IsSingleByte;
        }

        public static double Int64BitsToDouble(this long value) {
            return BitConverter.ToDouble(BitConverter.GetBytes(value), 0);
        }

        public static string ConvertEscapeChars(this string input) {
            return re.Replace(input, m => (((char)UInt32.Parse(m.Groups[1].Value, NumberStyles.HexNumber))).ToString());
        }

        public static object ConvertFromOATime(this double value) {
            if ((value>=0.0)&&(value<60.0)) {
                value++;
            }
            return DateTime.FromOADate(value);
        }

        public static void FixDataTypes(this DataSet dataset) {
            var tables=new List<DataTable>(dataset.Tables.Count);
            bool convert=false;
            foreach (DataTable table in dataset.Tables) {
                if (table.Rows.Count==0) {
                    tables.Add(table);
                    continue;
                }
                DataTable newTable=null;
                for (int i=0; i<table.Columns.Count; i++) {
                    Type type=null;
                    foreach (DataRow row in table.Rows) {
                        if (row.IsNull(i))
                            continue;
                        Type curType=row[i].GetType();
                        if (curType!=type) {
                            if (type==null)
                                type=curType;
                            else {
                                type=null;
                                break;
                            }
                        }
                    }
                    if (type!=null) {
                        convert=true;
                        if (newTable==null)
                            newTable=table.Clone();
                        newTable.Columns[i].DataType=type;
                    }
                }
                if (newTable!=null) {
                    newTable.BeginLoadData();
                    foreach (DataRow row in table.Rows) {
                        newTable.ImportRow(row);
                    }

                    newTable.EndLoadData();
                    tables.Add(newTable);
                } else
                    tables.Add(table);
            }
            if (convert) {
                dataset.Tables.Clear();
                dataset.Tables.AddRange(tables.ToArray());
            }
        }

        public static void AddColumnHandleDuplicate(this DataTable table, string columnName) {
            //if a colum  already exists with the name append _i to the duplicates
            string adjustedColumnName=columnName;
            DataColumn column=table.Columns[columnName];
            int i=1;
            while (column!=null) {
                adjustedColumnName=string.Format("{0}_{1}", columnName, i);
                column=table.Columns[adjustedColumnName];
                i++;
            }

            table.Columns.Add(adjustedColumnName, typeof(Object));
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

            return new int[2] { int.Parse(rowString), columnValue };
        }

        #region Nested type: MarkDownTable

        public class MarkDownTable {
            public string Name {
                get;
                set;
            }
            public string Value {
                get;
                set;
            }
        }

        #endregion
    }
}