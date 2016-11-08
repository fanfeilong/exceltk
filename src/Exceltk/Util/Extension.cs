using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace ExcelToolKit {
    public static class Extension {
        // Number of 100ns ticks per time unit
        private const long TicksPerMillisecond=10000;
        private const long TicksPerSecond=TicksPerMillisecond*1000;
        private const long TicksPerMinute=TicksPerSecond*60;
        private const long TicksPerHour=TicksPerMinute*60;
        private const long TicksPerDay=TicksPerHour*24;

        // Number of milliseconds per time unit
        private const int MillisPerSecond=1000;
        private const int MillisPerMinute=MillisPerSecond*60;
        private const int MillisPerHour=MillisPerMinute*60;
        private const int MillisPerDay=MillisPerHour*24;

        // Number of days in a non-leap year
        private const int DaysPerYear=365;
        // Number of days in 4 years
        private const int DaysPer4Years=DaysPerYear*4+1;       // 1461
        // Number of days in 100 years
        private const int DaysPer100Years=DaysPer4Years*25-1;  // 36524
        // Number of days in 400 years
        private const int DaysPer400Years=DaysPer100Years*4+1; // 146097

        // Number of days from 1/1/0001 to 12/31/1600
        private const int DaysTo1601=DaysPer400Years*4;          // 584388
        // Number of days from 1/1/0001 to 12/30/1899
        private const int DaysTo1899=DaysPer400Years*4+DaysPer100Years*3-367;
        // Number of days from 1/1/0001 to 12/31/1969
        internal const int DaysTo1970=DaysPer400Years*4+DaysPer100Years*3+DaysPer4Years*17+DaysPerYear; // 719,162
        // Number of days from 1/1/0001 to 12/31/9999
        private const int DaysTo10000=DaysPer400Years*25-366;  // 3652059

        internal const long MinTicks=0;
        internal const long MaxTicks=DaysTo10000*TicksPerDay-1;
        private const long MaxMillis=(long)DaysTo10000*MillisPerDay;

        private const long FileTimeOffset=DaysTo1601*TicksPerDay;
        private const long DoubleDateOffset=DaysTo1899*TicksPerDay;
        // The minimum OA date is 0100/01/01 (Note it's year 100).
        // The maximum OA date is 9999/12/31
        private const long OADateMinAsTicks=(DaysPer100Years-DaysPerYear)*TicksPerDay;
        // All OA dates must be greater than (not >=) OADateMinAsDouble
        private const double OADateMinAsDouble=-657435.0;
        // All OA dates must be less than (not <=) OADateMaxAsDouble
        private const double OADateMaxAsDouble=2958466.0;

        private const int DatePartYear=0;
        private const int DatePartDayOfYear=1;
        private const int DatePartMonth=2;
        private const int DatePartDay=3;

        private static readonly int[] DaysToMonth365= {
            0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365};
        private static readonly int[] DaysToMonth366= {
            0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366};

        public static readonly DateTime MinValue=new DateTime(MinTicks, DateTimeKind.Unspecified);
        public static readonly DateTime MaxValue=new DateTime(MaxTicks, DateTimeKind.Unspecified);

        private const UInt64 TicksMask=0x3FFFFFFFFFFFFFFF;
        private const UInt64 FlagsMask=0xC000000000000000;
        private const UInt64 LocalMask=0x8000000000000000;
        private const Int64 TicksCeiling=0x4000000000000000;
        private const UInt64 KindUnspecified=0x0000000000000000;
        private const UInt64 KindUtc=0x4000000000000000;
        private const UInt64 KindLocal=0x8000000000000000;
        private const UInt64 KindLocalAmbiguousDst=0xC000000000000000;
        private const Int32 KindShift=62;

        private const String TicksField="ticks";
        private const String DateDataField="dateData";

        static long DoubleDateToTicks(double value) {
            // The check done this way will take care of NaN
            if (!(value<OADateMaxAsDouble)||!(value>OADateMinAsDouble))
                throw new ArgumentException("Arg_OleAutDateInvalid");

            // Conversion to long will not cause an overflow here, as at this point the "value" is in between OADateMinAsDouble and OADateMaxAsDouble
            long millis=(long)(value*MillisPerDay+(value>=0?0.5:-0.5));
            // The interesting thing here is when you have a value like 12.5 it all positive 12 days and 12 hours from 01/01/1899
            // However if you a value of -12.25 it is minus 12 days but still positive 6 hours, almost as though you meant -11.75 all negative
            // This line below fixes up the millis in the negative case
            if (millis<0) {
                millis-=(millis%MillisPerDay)*2;
            }

            millis+=DoubleDateOffset/TicksPerMillisecond;

            if (millis<0||millis>=MaxMillis)
                throw new ArgumentException("Arg_OleAutDateScale");
            return millis*TicksPerMillisecond;
        }

#if !OS_WINDOWS
        public static void Close(this XmlReader xmlReader) {
            xmlReader.Dispose();
        }
        public static void Close(this Stream stream){
            stream.Dispose();
        }
        public static void Close(this MemoryStream stream){
            stream.Dispose();
        }
        public static void Close(this BinaryReader stream){
            stream.Dispose();
        }
        public static DateTime FromOADate(this double d) {
            return new DateTime(DoubleDateToTicks(d), DateTimeKind.Unspecified);
        }
#else
        public static DateTime FromOADate(this double d) {
            return new DateTime(DoubleDateToTicks(d), DateTimeKind.Unspecified);
        }
#endif
        public static string GetUserName(){
#if OS_WINDOWS
            return Environment.UserName;
#else
            return "";
#endif
        }
        public static Encoding DefaultEncoding(){
            #if OS_WINDOWS
                return Encoding.Default;
            #else
                return Encoding.UTF8;
            #endif
        }

        private static readonly Regex re=new Regex("_x([0-9A-F]{4,4})_");

        public static DataTable Shrink(this DataTable dataTable) {
            int columnCount=dataTable.Columns.Count;
            int rowCount=dataTable.Rows.Count;
            for (int r=rowCount-1; r>=0; r--) {
                for (int c=columnCount-1; c>=0; c--) {
                    object cell=dataTable.Rows[r][c];
                    string value="";
                    if(cell!=null){
                        var xlsCell=cell as XlsCell;
                        if (cell is XlsCell) {
                            value=xlsCell.Value.ToString();
                        } else {
                            value=cell.ToString();
                        }
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
                    if(cell!=null){
                        var xlsCell = cell as XlsCell;
                        if (cell is XlsCell){
                            value = xlsCell.Value.ToString();
                        } else{
                            value = cell.ToString();
                        }
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

        public static string ToMd(this DataTable table,bool insertHeader=true) {
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
                    string value = GetCellValue(cell);
                    sb.Append(value).Append("|");
                }

                sb.Append("\r\n");
                if (i==0 && insertHeader) {
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

        private static string GetCellValue(object cell){
            if(cell==null){
                return "";
            }
            string value;
            var xlsCell = cell as XlsCell;
            if (xlsCell != null){
                value = xlsCell.MarkDownText;
            } else{
                value = cell.ToString();
            }

            // Decimal precision
            if(Config.DecimalPrecision>0){
                if (Regex.IsMatch(value, @"^(-?[0-9]{1,}[.][0-9]*)$")) {
                    var old = value;
                    value=string.Format(Config.DecimalFormat, Double.Parse(value));
                    //Console.Write("{0}/{1} ",old,value);
                }    
            }

            return value;
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
            return value.FromOADate();
        }

        //public static void FixDataTypes(this DataSet dataset) {
        //    var tables=new List<DataTable>(dataset.Tables.Count);
        //    bool convert=false;
        //    foreach (DataTable table in dataset.Tables) {
        //        if (table.Rows.Count==0) {
        //            tables.Add(table);
        //            continue;
        //        }
        //        DataTable newTable=null;
        //        for (int i=0; i<table.Columns.Count; i++) {
        //            Type type=null;
        //            foreach (DataRow row in table.Rows) {
        //                if (row.IsNull(i))
        //                    continue;
        //                Type curType=row[i].GetType();
        //                if (curType!=type) {
        //                    if (type==null)
        //                        type=curType;
        //                    else {
        //                        type=null;
        //                        break;
        //                    }
        //                }
        //            }
        //            if (type!=null) {
        //                convert=true;
        //                if (newTable==null)
        //                    newTable=table.Clone();
        //                newTable.Columns[i].DataType=type;
        //            }
        //        }
        //        if (newTable!=null) {
        //            newTable.BeginLoadData();
        //            foreach (DataRow row in table.Rows) {
        //                newTable.Rows.Add(row);
        //            }

        //            newTable.EndLoadData();
        //            tables.Add(newTable);
        //        } else
        //            tables.Add(table);
        //    }
        //    if (convert) {
        //        dataset.Tables.Clear();
        //        dataset.Tables.AddRange(tables.ToArray());
        //    }
        //}

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

        public static bool IsHexDigit(this char digit) {
            return (('0'<=digit&&digit<='9')||
                    ('a'<=digit&&digit<='f')||
                    ('A'<=digit&&digit<='F'));
        }
        public static int FromHex(this char digit) {
            if ('0'<=digit&&digit<='9') {
                return (int)(digit-'0');
            }

            if ('a'<=digit&&digit<='f')
                return (int)(digit-'a'+10);

            if ('A'<=digit&&digit<='F')
                return (int)(digit-'A'+10);

            throw new ArgumentException("digit");
        }
    }
}