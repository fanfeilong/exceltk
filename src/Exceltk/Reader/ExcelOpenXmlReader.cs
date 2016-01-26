using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Xml;
using ExcelToolKit.OpenXmlFormat;

namespace ExcelToolKit {
    public class ExcelOpenXmlReader : IExcelDataReader {
        #region Members

        private const string COLUMN="Column";
        private readonly List<int> m_defaultDateTimeStyles;
        private bool disposed;
        private object[] m_cellsValues;
        private int m_depth;
        private int m_emptyRowCount;
        private string m_exceptionMessage;
        private string m_instanceId=Guid.NewGuid().ToString();
        private bool m_isClosed;
        private bool m_isFirstRead;
        private bool m_isFirstRowAsColumnNames;
        private bool m_isValid;

        private string m_namespaceUri;
        private int m_resultIndex;
        private object[] m_savedCellsValues;
        private Stream m_sheetStream;
        private XlsxWorkbook m_workbook;
        private XmlReader m_xmlReader;
        private ZipWorker m_zipWorker;

        #endregion

        internal ExcelOpenXmlReader() {
            m_isValid=true;
            m_isFirstRead=true;

            m_defaultDateTimeStyles=new List<int>(new[]{
                14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47
            });
        }

        #region IExcelDataReader Members

        public void Initialize(Stream fileStream) {
            m_zipWorker=new ZipWorker();
            m_zipWorker.Extract(fileStream);

            if (!m_zipWorker.IsValid) {
                m_isValid=false;
                m_exceptionMessage=m_zipWorker.ExceptionMessage;
                Close();
            } else {
                m_isValid=true;
                ReadGlobals();
            }
        }

        public DataSet AsDataSet() {
            return AsDataSet(true);
        }

        public DataSet AsDataSet(bool convertOADateTime) {
            if (!m_isValid) {
                return null;
            } else {
                return ReadDataSet();
            }
        }

        public bool NextResult() {
            if (m_resultIndex<(ResultsCount-1)) {
                m_resultIndex++;
                m_isFirstRead=true;
                m_savedCellsValues=null;
                return true;
            } else {
                return false;
            }
        }

        public bool Read() {
            if (!m_isValid) {
                return false;
            }
            
            if (m_isFirstRead&&!InitializeSheetRead()) {
                return false;
            }

            return ReadSheetRow(m_workbook.Sheets[m_resultIndex]);
        }

        public void Close() {
            
            if (IsClosed) {
                return;
            }
            m_isClosed=true;

            if (m_xmlReader!=null) {
                m_xmlReader.Close();
                m_xmlReader = null;
            }

            if (m_sheetStream!=null) {
                m_sheetStream.Close();
                m_sheetStream = null;
            }

            if (m_zipWorker!=null) {
                m_zipWorker.Dispose();
                m_zipWorker = null;
            }
        }

        public bool IsFirstRowAsColumnNames {
            get {
                return m_isFirstRowAsColumnNames;
            }
            set {
                m_isFirstRowAsColumnNames=value;
            }
        }

        public bool IsValid {
            get {
                return m_isValid;
            }
        }

        public string ExceptionMessage {
            get {
                return m_exceptionMessage;
            }
        }

        public string Name {
            get {
                return (m_resultIndex>=0&&m_resultIndex<ResultsCount)
                     ? m_workbook.Sheets[m_resultIndex].Name
                     : null;
            }
        }

        public int Depth {
            get {
                return m_depth;
            }
        }

        public int ResultsCount {
            get {
                return m_workbook==null?-1:m_workbook.Sheets.Count;
            }
        }

        public bool IsClosed {
            get {
                return m_isClosed;
            }
        }

        public int FieldCount {
            get {
                return (m_resultIndex>=0&&m_resultIndex<ResultsCount)
                           ?m_workbook.Sheets[m_resultIndex].ColumnsCount
                           :-1;
            }
        }

        public bool GetBoolean(int i) {
            if (IsDBNull(i)) {
                return false;
            } else {
                return Boolean.Parse(m_cellsValues[i].ToString());
            }
        }

        public DateTime GetDateTime(int i) {
            if (IsDBNull(i)) {
                return DateTime.MinValue;
            } else {
                try {
                    return (DateTime)m_cellsValues[i];
                } catch (InvalidCastException) {
                    return DateTime.MinValue;
                }
            }
        }

        public decimal GetDecimal(int i) {
            if (IsDBNull(i)) {
                return decimal.MinValue;
            } else {
                return decimal.Parse(m_cellsValues[i].ToString());
            }
        }

        public double GetDouble(int i) {
            if (IsDBNull(i)) {
                return double.MinValue;
            } else {
                return double.Parse(m_cellsValues[i].ToString());
            }
        }

        public float GetFloat(int i) {
            if (IsDBNull(i)) {
                return float.MinValue;
            } else {
                return float.Parse(m_cellsValues[i].ToString());
            }
        }

        public short GetInt16(int i) {
            if (IsDBNull(i)) {
                return short.MinValue;
            } else {
                return short.Parse(m_cellsValues[i].ToString());
            }
        }

        public int GetInt32(int i) {
            if (IsDBNull(i)) {
                return int.MinValue;
            } else {
                return int.Parse(m_cellsValues[i].ToString());
            }
        }

        public long GetInt64(int i) {
            if (IsDBNull(i)) {
                return long.MinValue;
            } else {
                return long.Parse(m_cellsValues[i].ToString());
            }
        }

        public string GetString(int i) {
            if (IsDBNull(i)) {
                return null;
            } else {
                return m_cellsValues[i].ToString();
            }
        }

        public object GetValue(int i) {
            return m_cellsValues[i];
        }

        public bool IsDBNull(int i) {
            return (null==m_cellsValues[i])
                || (DBNull.Value==m_cellsValues[i]);
        }

        public object this[int i] {
            get {
                return m_cellsValues[i];
            }
        }

        public DataTable GetSchemaTable() {
            throw new NotSupportedException();
        }

        public int RecordsAffected {
            get {
                throw new NotSupportedException();
            }
        }

        public byte GetByte(int i) {
            throw new NotSupportedException();
        }

        public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length) {
            throw new NotSupportedException();
        }

        public char GetChar(int i) {
            throw new NotSupportedException();
        }

        public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length) {
            throw new NotSupportedException();
        }

        public IDataReader GetData(int i) {
            throw new NotSupportedException();
        }

        public string GetDataTypeName(int i) {
            throw new NotSupportedException();
        }

        public Type GetFieldType(int i) {
            throw new NotSupportedException();
        }

        public Guid GetGuid(int i) {
            throw new NotSupportedException();
        }

        public string GetName(int i) {
            throw new NotSupportedException();
        }

        public int GetOrdinal(string name) {
            throw new NotSupportedException();
        }

        public int GetValues(object[] values) {
            throw new NotSupportedException();
        }

        public object this[string name] {
            get {
                throw new NotSupportedException();
            }
        }

        #endregion

        #region Implement
        
        private void ReadGlobals() {
            m_workbook=new XlsxWorkbook(
                m_zipWorker.GetWorkbookStream(),
                m_zipWorker.GetWorkbookRelsStream(),
                m_zipWorker.GetSharedStringsStream(),
                m_zipWorker.GetStylesStream());

            CheckDateTimeNumFmts(m_workbook.Styles.NumFmts);
        }

        private void CheckDateTimeNumFmts(List<XlsxNumFmt> list) {
            if (list.Count==0)
                return;

            foreach (XlsxNumFmt numFmt in list) {
                if (string.IsNullOrEmpty(numFmt.FormatCode))
                    continue;
                string fc=numFmt.FormatCode.ToLower();

                int pos;
                while ((pos=fc.IndexOf('"'))>0) {
                    int endPos=fc.IndexOf('"', pos+1);

                    if (endPos>0)
                        fc=fc.Remove(pos, endPos-pos+1);
                }

                //it should only detect it as a date if it contains
                //dd mm mmm yy yyyy
                //h hh ss
                //AM PM
                //and only if these appear as "words" so either contained in [ ]
                //or delimted in someway
                //updated to not detect as date if format contains a #
                var formatReader=new FormatReader {
                    FormatString=fc
                };
                if (formatReader.IsDateFormatString()) {
                    m_defaultDateTimeStyles.Add(numFmt.Id);
                }
            }
        }

        private void ReadSheetGlobals(XlsxWorksheet sheet) {
            if (!ResetSheetReader(sheet)) {
                return;
            }

            //count rows and cols in case there is no dimension elements
            m_namespaceUri=null;
            int rows=0;
            int cols=0;
            int biggestColumn=0; 

            while (m_xmlReader.Read()) {
                if (m_xmlReader.NodeType==XmlNodeType.Element&&m_xmlReader.LocalName==XlsxWorksheet.N_worksheet) {
                    //grab the namespaceuri from the worksheet element
                    m_namespaceUri=m_xmlReader.NamespaceURI;
                }

                if (m_xmlReader.NodeType==XmlNodeType.Element&&m_xmlReader.LocalName==XlsxWorksheet.N_dimension) {
                    string dimValue=m_xmlReader.GetAttribute(XlsxWorksheet.A_ref);
                    sheet.Dimension=new XlsxDimension(dimValue);
                    break;
                }

                if (m_xmlReader.NodeType==XmlNodeType.Element&&m_xmlReader.LocalName==XlsxWorksheet.N_row) {
                    rows++;
                }

                // check cells so we can find size of sheet if can't work it out from dimension or 
                // col elements (dimension should have been set before the cells if it was available)
                // ditto for cols
                if (sheet.Dimension==null&&cols==0&&m_xmlReader.NodeType==XmlNodeType.Element&&m_xmlReader.LocalName==XlsxWorksheet.N_c) {
                    string refAttribute=m_xmlReader.GetAttribute(XlsxWorksheet.A_r);

                    if (refAttribute!=null) {
                        int[] thisRef=refAttribute.ReferenceToColumnAndRow();
                        if (thisRef[1]>biggestColumn) {
                            biggestColumn=thisRef[1];
                        }
                    }
                }
            }

            // if we didn't get a dimension element then use the calculated rows/cols to create it
            if (sheet.Dimension==null) {
                if (cols==0) {
                    cols=biggestColumn;
                }

                if (rows==0||cols==0) {
                    sheet.IsEmpty=true;
                    return;
                }

                sheet.Dimension=new XlsxDimension(rows, cols);

                //we need to reset our position to sheet data
                if (!ResetSheetReader(sheet)) {
                    return;
                }
            }

            // read up to the sheetData element. if this element is empty then 
            // there aren't any rows and we need to null out dimension
            m_xmlReader.ReadToFollowing(XlsxWorksheet.N_sheetData, m_namespaceUri);
            if (m_xmlReader.IsEmptyElement) {
                sheet.IsEmpty=true;
            }
        }

        private bool ResetSheetReader(XlsxWorksheet sheet)
        {
            if (m_sheetStream!=null) {
                m_sheetStream.Close();
                m_sheetStream=null;
            }

            if (m_xmlReader!=null) {
                m_xmlReader.Close();
                m_xmlReader=null;
            }

            m_sheetStream=m_zipWorker.GetWorksheetStream(sheet.Path);
            if (null==m_sheetStream) {
                return false;
            }

            m_xmlReader=XmlReader.Create(m_sheetStream);
            if (null==m_xmlReader) {
                return false;
            }

            return true;
        }

        private bool ReadSheetRow(XlsxWorksheet sheet) {
            if (sheet.ColumnsCount<0) {
                //Console.WriteLine("Columons Count Can NOT BE Negative");
                return false;
            }
            if (null==m_xmlReader)
                return false;

            if (m_emptyRowCount!=0) {
                m_cellsValues=new object[sheet.ColumnsCount];
                m_emptyRowCount--;
                m_depth++;

                return true;
            }

            if (m_savedCellsValues!=null) {
                m_cellsValues=m_savedCellsValues;
                m_savedCellsValues=null;
                m_depth++;

                return true;
            }

            bool isRow=false;
            bool isSheetData=(m_xmlReader.NodeType==XmlNodeType.Element&&
                                m_xmlReader.LocalName==XlsxWorksheet.N_sheetData);
            if (isSheetData) {
                isRow=m_xmlReader.ReadToFollowing(XlsxWorksheet.N_row, m_namespaceUri);
            } else {
                if (m_xmlReader.LocalName==XlsxWorksheet.N_row&&m_xmlReader.NodeType==XmlNodeType.EndElement) {
                    m_xmlReader.Read();
                }
                isRow=(m_xmlReader.NodeType==XmlNodeType.Element&&m_xmlReader.LocalName==XlsxWorksheet.N_row);
                //Console.WriteLine("isRow:{0}/{1}/{2}", isRow,m_xmlReader.NodeType, m_xmlReader.LocalName);
            }

            if (isRow) {
                m_cellsValues=new object[sheet.ColumnsCount];
                if (sheet.ColumnsCount>13) {
                    int i=sheet.ColumnsCount;
                }

                int rowIndex=int.Parse(m_xmlReader.GetAttribute(XlsxWorksheet.A_r));
                if (rowIndex!=(m_depth+1)) {
                    m_emptyRowCount=rowIndex-m_depth-1;
                }

                bool hasValue=false;
                string a_s=String.Empty;
                string a_t=String.Empty;
                string a_r=String.Empty;
                int col=0;
                int row=0;

                while (m_xmlReader.Read()) {
                    if (m_xmlReader.Depth==2) {
                        break;
                    }

                    if (m_xmlReader.NodeType==XmlNodeType.Element) {
                        hasValue=false;

                        if (m_xmlReader.LocalName==XlsxWorksheet.N_c) {
                            a_s=m_xmlReader.GetAttribute(XlsxWorksheet.A_s);
                            a_t=m_xmlReader.GetAttribute(XlsxWorksheet.A_t);
                            a_r=m_xmlReader.GetAttribute(XlsxWorksheet.A_r);
                            XlsxDimension.XlsxDim(a_r, out col, out row);
                        } else if (m_xmlReader.LocalName==XlsxWorksheet.N_v||m_xmlReader.LocalName==XlsxWorksheet.N_t) {
                            hasValue=true;
                        } else {
                            //Console.WriteLine("Error");
                        }
                    }

                    if (m_xmlReader.NodeType==XmlNodeType.Text&&hasValue) {
                        double number;
                        object o=m_xmlReader.Value;

                        var style=NumberStyles.Any;
                        CultureInfo culture=CultureInfo.InvariantCulture;

                        if (double.TryParse(o.ToString(), style, culture, out number))
                            o=number;

                        #region Read Cell Value

                        if (null!=a_t&&a_t==XlsxWorksheet.A_s) //if string
                        {
                            o=m_workbook.SST[int.Parse(o.ToString())].ConvertEscapeChars();
                        } // Requested change 4: missing (it appears that if should be else if)
                        else if (null!=a_t&&a_t==XlsxWorksheet.N_inlineStr) //if string inline
                        {
                            o=o.ToString().ConvertEscapeChars();
                        } else if (a_t=="b") //boolean
                        {
                            o=m_xmlReader.Value=="1";
                        } else if (a_t=="str") {
                            o=m_xmlReader.Value;
                        } else if (null!=a_s) //if something else
                        {
                            XlsxXf xf=m_workbook.Styles.CellXfs[int.Parse(a_s)];
                            if (xf.ApplyNumberFormat&&o!=null&&o.ToString()!=string.Empty&&
                                IsDateTimeStyle(xf.NumFmtId)) {
                                o=number.ConvertFromOATime();
                            } else if (xf.NumFmtId==49) {
                                o=o.ToString();
                            }
                        }

                        #endregion

                        if (col-1<m_cellsValues.Length) {
                            //Console.WriteLine(o);
                            if (string.IsNullOrEmpty(o.ToString())) {
                                //Console.WriteLine("Error");
                            }
                            m_cellsValues[col-1]=o;
                        } else {
                            //Console.WriteLine("Error");
                        }
                    } else {
                        if (m_xmlReader.LocalName==XlsxWorksheet.N_v) {
                            //Console.WriteLine("No Value");
                        }
                    }
                }

                if (m_emptyRowCount>0) {
                    m_savedCellsValues=m_cellsValues;
                    return ReadSheetRow(sheet);
                }
                m_depth++;

                return true;
            } else {
                //Console.WriteLine(m_xmlReader.LocalName.ToString());
                return false;
            }
        }

        private bool ReadHyperLinks(XlsxWorksheet sheet, DataTable table) {
            // ReadTo HyperLinks Node
            if (m_xmlReader==null) {
                //Console.WriteLine("m_xmlReader is null");
                return false;
            }

            //Console.WriteLine(m_xmlReader.Depth.ToString());

            m_xmlReader.ReadToFollowing(XlsxWorksheet.N_hyperlinks);
            if (m_xmlReader.IsEmptyElement) {
                //Console.WriteLine("not find hyperlink");
                return false;
            }

            // Read Realtionship Table
            //Console.WriteLine("sheetrel:{0}", sheet.Path);
            Stream sheetRelStream=m_zipWorker.GetWorksheetRelsStream(sheet.Path);
            var hyperDict=new Dictionary<string, string>();
            if (sheetRelStream!=null) {
                using (XmlReader reader=XmlReader.Create(sheetRelStream)) {
                    while (reader.Read()) {
                        if (reader.NodeType==XmlNodeType.Element&&reader.LocalName==XlsxWorkbook.N_rel) {
                            string rid=reader.GetAttribute(XlsxWorkbook.A_id);
                            hyperDict[rid]=reader.GetAttribute(XlsxWorkbook.A_target);
                        }
                    }
                    sheetRelStream.Close();
                }
            }


            // Read All HyperLink Node
            while (m_xmlReader.Read()) {
                if (m_xmlReader.NodeType!=XmlNodeType.Element)
                    break;
                if (m_xmlReader.LocalName!=XlsxWorksheet.N_hyperlink)
                    break;
                string aref=m_xmlReader.GetAttribute(XlsxWorksheet.A_ref);
                string display=m_xmlReader.GetAttribute(XlsxWorksheet.A_display);
                string rid=m_xmlReader.GetAttribute(XlsxWorksheet.A_rid);
                ////Console.WriteLine("{0}:{1}", aref.Substring(1), display);

                //Console.WriteLine("hyperlink:{0}",hyperDict[rid]);
                string hyperlink=display;
                if (hyperDict.ContainsKey(rid)) {
                    hyperlink=hyperDict[rid];
                }

                int col=-1;
                int row=-1;
                XlsxDimension.XlsxDim(aref, out col, out row);
                //Console.WriteLine("{0}:[{1},{2}]",aref, row, col);
                if (col>=1&&row>=1) {
                    row=row-1;
                    col=col-1;
                    if (row==0&&m_isFirstRowAsColumnNames) {
                        // TODO(fanfeilong):
                        string value=table.Columns[col].ColumnName;
                        var cell=new XlsCell(value);
                        cell.SetHyperLink(hyperlink);
                        table.Columns[col].DefaultValue=cell;
                    } else {
                        object value=table.Rows[row][col];
                        var cell=new XlsCell(value);
                        cell.SetHyperLink(hyperlink);
                        //Console.WriteLine(cell.MarkDownText);
                        table.Rows[row][col]=cell;
                    }
                }
            }

            // Close
            m_xmlReader.Close();
            if (m_sheetStream!=null) {
                m_sheetStream.Close();
            }

            return true;
        }

        private bool InitializeSheetRead() {
            if (ResultsCount<=0)
                return false;

            ReadSheetGlobals(m_workbook.Sheets[m_resultIndex]);

            if (m_workbook.Sheets[m_resultIndex].Dimension==null)
                return false;

            m_isFirstRead=false;

            m_depth=0;
            m_emptyRowCount=0;

            return true;
        }

        private bool IsDateTimeStyle(int styleId) {
            return m_defaultDateTimeStyles.Contains(styleId);
        }

        private Dictionary<int, XlsxDimension> DetectDemension() {
            var dict=new Dictionary<int, XlsxDimension>();
            for (int sheetIndex=0; sheetIndex<m_workbook.Sheets.Count; sheetIndex++) {
                XlsxWorksheet sheet=m_workbook.Sheets[sheetIndex];

                ReadSheetGlobals(sheet);

                if (sheet.Dimension!=null) {
                    m_depth=0;
                    m_emptyRowCount=0;

                    // 检测100行
                    int detectRows=Math.Min(sheet.Dimension.LastRow, 100);
                    int maxColumnCount=0;
                    while (detectRows>0) {
                        ReadSheetRow(sheet);
                        maxColumnCount=Math.Max(LastIndexOfNonNull(m_cellsValues)+1, maxColumnCount);
                        detectRows--;
                    }

                    // 如果实际检测出来的列个数小于元数据里的列数，
                    if (maxColumnCount<sheet.Dimension.LastCol) {
                        dict[sheetIndex]=new XlsxDimension(sheet.Dimension.LastRow, maxColumnCount);
                    } else {
                        dict[sheetIndex]=sheet.Dimension;
                    }
                } else {
                    dict[sheetIndex]=sheet.Dimension;
                }
            }
            return dict;
        }

        private int LastIndexOfNonNull(object[] cellsValues) {
            for (int i=cellsValues.Length-1; i>=0; i--) {
                if (cellsValues[i]!=null) {
                    return i;
                }
            }
            return 0;
        }

        private DataSet ReadDataSet() {
            var dataset=new DataSet();

            Dictionary<int, XlsxDimension> demensionDict=DetectDemension();

            for (int sheetIndex=0; sheetIndex<m_workbook.Sheets.Count; sheetIndex++) {
                XlsxWorksheet sheet=m_workbook.Sheets[sheetIndex];
                var table=new DataTable(m_workbook.Sheets[sheetIndex].Name);

                ReadSheetGlobals(sheet);
                sheet.Dimension=demensionDict[sheetIndex];

                if (sheet.Dimension==null) {
                    continue;
                }

                m_depth=0;
                m_emptyRowCount=0;

                // Reada Columns
                //Console.WriteLine("Read Columns");
                if (!m_isFirstRowAsColumnNames) {
                    // No Sheet Columns
                    //Console.WriteLine("SheetName:{0}, ColumnCount:{1}", sheet.Name, sheet.ColumnsCount);
                    for (int i=0; i<sheet.ColumnsCount; i++) {
                        table.Columns.Add(null, typeof(Object));
                    }
                } else if (ReadSheetRow(sheet)) {
                    // Read Sheet Columns
                    //Console.WriteLine("Read Sheet Columns");
                    for (int index=0; index<m_cellsValues.Length; index++) {
                        if (m_cellsValues[index]!=null&&m_cellsValues[index].ToString().Length>0) {
                            table.AddColumnHandleDuplicate(m_cellsValues[index].ToString());
                        } else {
                            table.AddColumnHandleDuplicate(string.Concat(COLUMN, index));
                        }
                    }
                } else {
                    continue;
                }

                // Read Sheet Rows
                //Console.WriteLine("Read Sheet Rows");
                table.BeginLoadData();
                //Console.WriteLine("SheetIndex Is:{0},Name:{1}",sheetIndex,sheet.Name);
                while (ReadSheetRow(sheet)) {
                    table.Rows.Add(m_cellsValues);
                }
                if (table.Rows.Count>0) {
                    dataset.Tables.Add(table);
                }

                // Read HyperLinks
                //Console.WriteLine("Read Sheet HyperLinks:{0}",table.Rows.Count);
                ReadHyperLinks(sheet, table);

                table.EndLoadData();
            }
            dataset.AcceptChanges();
            dataset.FixDataTypes();
            return dataset;
        }
        
        #endregion

        #region IDispose

        public void Dispose() {
            Dispose(true);

            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing) {
            // Check to see if Dispose has already been called.

            if (!disposed) {
                if (disposing) {
                    if (m_xmlReader!=null)
                        ((IDisposable)m_xmlReader).Dispose();
                    if (m_sheetStream!=null)
                        m_sheetStream.Dispose();
                    if (m_zipWorker!=null)
                        m_zipWorker.Dispose();
                }

                m_zipWorker=null;
                m_xmlReader=null;
                m_sheetStream=null;

                m_workbook=null;
                m_cellsValues=null;
                m_savedCellsValues=null;

                disposed=true;
            }
        }

        ~ExcelOpenXmlReader() {
            Dispose(false);
        }

        #endregion
    }
}