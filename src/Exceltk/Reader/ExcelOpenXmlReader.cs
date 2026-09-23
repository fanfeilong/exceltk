using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Xml;
using System.Text;
using Exceltk.Reader.Parser;
using Exceltk.Reader.Xml;

namespace Exceltk.Reader {
    public class ExcelOpenXmlReader : IExcelDataReader {
        #region Members

        private const string COLUMN = "Column";
        private readonly List<int> m_defaultDateTimeStyles;
        private bool disposed;
        private object[] m_cellsValues;
        private int m_depth;
        private int m_emptyRowCount;
        private string m_exceptionMessage;
        private string m_instanceId = Guid.NewGuid().ToString();
        private bool m_isClosed;
        private bool m_isValid;
        private bool m_ownsZipWorker;

        private string m_namespaceUri;
        private object[] m_savedCellsValues;
        private Stream m_sheetStream;
        private XlsxWorkbook m_workbook;
        private XmlReader m_xmlReader;
        private ZipWorker m_zipWorker;
        private IEnumerator<XmlRowPackage> m_rowPackages;

        #endregion

        #region ctor
        internal ExcelOpenXmlReader() {
            m_isValid = true;
            //m_isFirstRead = true;

            m_defaultDateTimeStyles = new List<int>(new[]{
                14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47
            });
        }
        #endregion

        #region IExcelDataReader Members

        public void Open(Stream fileStream) {
            m_zipWorker = new ZipWorker();
            m_zipWorker.Extract(fileStream);
            m_ownsZipWorker = true;

            if (!m_zipWorker.IsValid) {
                m_isValid = false;
                m_exceptionMessage = m_zipWorker.ExceptionMessage;
                Dispose();
            } else {
                m_isValid = true;
                ReadGlobals();
            }
        }

        public DataSet AsDataSet() {
            if (!m_isValid) {
                return null;
            } else {
                return ReadDataSet();
            }
        }

        public void Close() {

            if (m_isClosed) {
                return;
            }
            m_isClosed = true;

            if (m_xmlReader != null) {
                m_xmlReader.Close();
                m_xmlReader = null;
            }

            if (m_sheetStream != null) {
                m_sheetStream.Close();
                m_sheetStream = null;
            }

            if (m_rowPackages != null) {
                m_rowPackages.Dispose();
                m_rowPackages = null;
            }

            if (m_zipWorker != null) {
                if (m_ownsZipWorker) {
                    m_zipWorker.Dispose();
                }
                m_zipWorker = null;
            }
        }

        #endregion

        #region Implement

        private void ReadGlobals() {
            m_workbook = new XlsxWorkbook(
                m_zipWorker.GetWorkbookStream(),
                m_zipWorker.GetWorkbookRelsStream(),
                m_zipWorker.GetSharedStringsStream(),
                m_zipWorker.GetStylesStream());

            // Some workbooks omit styles.xml; treat that as empty styles instead of NRE (#10).
            if (m_workbook.Styles == null) {
                CheckDateTimeNumFmts(new List<XlsxNumFmt>());
            } else {
                CheckDateTimeNumFmts(m_workbook.Styles.NumFmts);
            }
        }

        private void CheckDateTimeNumFmts(List<XlsxNumFmt> list) {
            if (list == null || list.Count == 0) {
                return;
            }

            foreach (XlsxNumFmt numFmt in list) {
                if (string.IsNullOrEmpty(numFmt.FormatCode)) {
                    continue;
                }
                string fc = numFmt.FormatCode.ToLower();

                int pos;
                while ((pos = fc.IndexOf('"')) > 0) {
                    int endPos = fc.IndexOf('"', pos + 1);

                    if (endPos > 0) {
                        fc = fc.Remove(pos, endPos - pos + 1);
                    }
                }

                //it should only detect it as a date if it contains
                //dd mm mmm yy yyyy
                //h hh ss
                //AM PM
                //and only if these appear as "words" so either contained in [ ]
                //or delimted in someway
                //updated to not detect as date if format contains a #
                var formatReader = new FormatReader {
                    FormatString = fc
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
            m_namespaceUri = null;
            int rows = 0;
            int cols = 0;
            int biggestColumn = 0;

            while (m_xmlReader.Read()) {
                if (m_xmlReader.NodeType == XmlNodeType.Element && m_xmlReader.LocalName == XlsxWorksheet.N_worksheet) {
                    //grab the namespaceuri from the worksheet element
                    m_namespaceUri = m_xmlReader.NamespaceURI;
                }

                if (m_xmlReader.NodeType == XmlNodeType.Element && m_xmlReader.LocalName == XlsxWorksheet.N_dimension) {
                    string dimValue = m_xmlReader.GetAttribute(XlsxWorksheet.A_ref);
                    sheet.Dimension = new XlsxDimension(dimValue);
                    break;
                }

                if (m_xmlReader.NodeType == XmlNodeType.Element && m_xmlReader.LocalName == XlsxWorksheet.N_row) {
                    rows++;
                }

                // check cells so we can find size of sheet if can't work it out from dimension or 
                // col elements (dimension should have been set before the cells if it was available)
                // ditto for cols
                if (sheet.Dimension == null && cols == 0 && m_xmlReader.NodeType == XmlNodeType.Element && m_xmlReader.LocalName == XlsxWorksheet.N_c) {
                    string refAttribute = m_xmlReader.GetAttribute(XlsxWorksheet.A_r);

                    if (refAttribute != null) {
                        int[] thisRef = refAttribute.ReferenceToColumnAndRow();
                        if (thisRef[1] > biggestColumn) {
                            biggestColumn = thisRef[1];
                        }
                    }
                }
            }

            // if we didn't get a dimension element then use the calculated rows/cols to create it
            if (sheet.Dimension == null) {
                if (cols == 0) {
                    cols = biggestColumn;
                }

                if (rows == 0 || cols == 0) {
                    sheet.IsEmpty = true;
                    return;
                }

                sheet.Dimension = new XlsxDimension(rows, cols);

                //we need to reset our position to sheet data
                if (!ResetSheetReader(sheet)) {
                    return;
                }
            }

            // read up to the sheetData element. if this element is empty then 
            // there aren't any rows and we need to null out dimension
            Debug.Assert(m_namespaceUri!=null);
            m_xmlReader.ReadToFollowing(XlsxWorksheet.N_sheetData, m_namespaceUri);
            if (m_xmlReader.IsEmptyElement) {
                sheet.IsEmpty=true;
                m_rowPackages = null;
            } else {
                // Stream rows as XmlRowPackage entities via XmlPackageParser.
                if (m_rowPackages != null) {
                    m_rowPackages.Dispose();
                }
                var parser = new XmlPackageParser(m_xmlReader, m_namespaceUri);
                m_rowPackages = parser.ParseRows().GetEnumerator();
            }                
        }

        private bool ResetSheetReader(XlsxWorksheet sheet) {
            if (m_sheetStream != null) {
                m_sheetStream.Close();
                m_sheetStream = null;
            }

            if (m_xmlReader != null) {
                m_xmlReader.Close();
                m_xmlReader = null;
            }

            m_sheetStream = m_zipWorker.GetWorksheetStream(sheet.Path);
            if (null == m_sheetStream) {
                return false;
            }

            m_xmlReader = XmlReader.Create(m_sheetStream);
            if (null == m_xmlReader) {
                return false;
            }

            return true;
        }

        private HyperLinkIndex ReadHyperLinkFormula(string thisSheetName, string formula){
            var sb = new StringBuilder();
            var f = formula.Substring(10);

            //HYPERLINK(#REF!,RIGHT(#REF!,3))
            //HYPERLINK(#REF!,SUBSTITUTE(#REF!,"https://coding.net/u/",""))
            //Console.WriteLine(formula);
            if(formula.StartsWith("HYPERLINK(#REF!")){
                //var begin = formula.IndexOf("\"");
                //var rest = formula.Substring(begin);
                //var end = rest.IndexOf("\"");
                //var h = rest.Substring(0,end);
                //Console.WriteLine(h);
                return null;
            }

            for(var i=0;i<f.Length;i++){
                var c = f[i];
                
                if(c==','){
                    var link = sb.ToString();
                    var pos = link.IndexOf("!");

                    var sheetName = "";
                    int col = 0;
                    int row = 0;
                    if(pos>=0){
                        Console.WriteLine(pos);
                        Console.WriteLine(link);
                        sheetName = link.Substring(0, pos);
                        var cellName = link.Substring(pos+1);
                        XlsxDimension.XlsxDim(cellName, out col, out row);

                    }else{
                        sheetName = thisSheetName;
                        var cellName = link.ToString();
                        //Console.WriteLine(cellName);
                        XlsxDimension.XlsxDim(cellName, out col, out row);
                    }

                    return new HyperLinkIndex(){
                        Sheet = sheetName,
                        Col = col,
                        Row = row
                    };
                }

                sb.Append(c);
            }
            return null;
        }

        private bool ReadSheetRow(XlsxWorksheet sheet) {
            if (sheet.ColumnsCount < 0) {
                return false;
            }

            if (m_emptyRowCount != 0) {
                m_cellsValues = new object[sheet.ColumnsCount];
                m_emptyRowCount--;
                m_depth++;

                return true;
            }

            if (m_savedCellsValues != null) {
                m_cellsValues = m_savedCellsValues;
                m_savedCellsValues = null;
                m_depth++;

                return true;
            }

            if (m_rowPackages == null || !m_rowPackages.MoveNext()) {
                return false;
            }

            XmlRowPackage rowPackage = m_rowPackages.Current;
            m_cellsValues = new object[sheet.ColumnsCount];

            Debug.Assert(rowPackage.RowIndexAttribute != null);
            int rowIndex = int.Parse(rowPackage.RowIndexAttribute);

            if (rowIndex != (m_depth + 1)) {
                m_emptyRowCount = rowIndex - m_depth - 1;
            }

            foreach (XmlCellPackage cellPackage in rowPackage.Cells) {
                ApplyCellPackage(sheet, cellPackage);
            }

            if (m_emptyRowCount > 0) {
                m_savedCellsValues = m_cellsValues;
                return ReadSheetRow(sheet);
            }
            m_depth++;

            return true;
        }

        private void ApplyCellPackage(XlsxWorksheet sheet, XmlCellPackage cellPackage) {
            string a_s = cellPackage.StyleId;
            string a_t = cellPackage.CellType;
            string a_r = cellPackage.Reference;
            int col;
            int row;
            if (string.IsNullOrEmpty(a_r)) {
                return;
            }
            XlsxDimension.XlsxDim(a_r, out col, out row);

            HyperLinkIndex hyperlinkIndex = null;
            if (!string.IsNullOrEmpty(cellPackage.Formula) && cellPackage.Formula.StartsWith("HYPERLINK(")) {
                hyperlinkIndex = this.ReadHyperLinkFormula(sheet.Name, cellPackage.Formula);
            }

            double number;
            object o = cellPackage.ValueText;

            if (double.TryParse(o.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out number)) {
                o = number;
            }

            if (null != a_t && a_t == XlsxWorksheet.A_s) {
                if (m_workbook.SST != null) {
                    var sstStr = m_workbook.SST[int.Parse(o.ToString())];
                    o = sstStr.ConvertEscapeChars();
                }
            } else if (null != a_t && a_t == XlsxWorksheet.N_inlineStr) {
                o = o.ToString().ConvertEscapeChars();
            } else if (a_t == "b") {
                o = cellPackage.ValueText == "1";
            } else if (a_t == "str") {
                o = cellPackage.ValueText;
            } else if (null != a_s && m_workbook.Styles != null && m_workbook.Styles.CellXfs != null) {
                int styleIndex;
                if (int.TryParse(a_s, out styleIndex)
                    && styleIndex >= 0
                    && styleIndex < m_workbook.Styles.CellXfs.Count) {
                    XlsxXf xf = m_workbook.Styles.CellXfs[styleIndex];
                    if (xf.ApplyNumberFormat && o != null && o.ToString() != string.Empty &&
                        IsDateTimeStyle(xf.NumFmtId)) {
                        o = number.ConvertFromOATime();
                    } else if (xf.NumFmtId == 49) {
                        o = o.ToString();
                    }
                }
            }

            if (col >= 1) {
                EnsureRowCapacity(sheet, col);
                if (hyperlinkIndex != null) {
                    var co = new XlsCell(o);
                    co.HyperLinkIndex = hyperlinkIndex;
                    m_cellsValues[col - 1] = co;
                } else {
                    m_cellsValues[col - 1] = o;
                }
            }
        }

        private bool ReadMergeCells(XlsxWorksheet sheet, DataTable table) {
            // Restart the sheet stream so we can locate mergeCells independently of row reading.
            if (!ResetSheetReader(sheet)) {
                return false;
            }

            var parser = new XmlPackageParser(m_xmlReader, m_namespaceUri);
            foreach (XmlMergePackage mergePackage in parser.ParseMerges()) {
                string aref = mergePackage.Ref;
                if (string.IsNullOrEmpty(aref)) {
                    continue;
                }

                string[] parts = aref.Split(':');
                int c1, r1, c2, r2;
                XlsxDimension.XlsxDim(parts[0], out c1, out r1);
                if (parts.Length > 1) {
                    XlsxDimension.XlsxDim(parts[1], out c2, out r2);
                } else {
                    c2 = c1;
                    r2 = r1;
                }

                c1--; r1--; c2--; r2--;
                if (c1 < 0 || r1 < 0) {
                    continue;
                }

                int rowSpan = r2 - r1 + 1;
                int colSpan = c2 - c1 + 1;
                if (rowSpan < 1 || colSpan < 1) {
                    continue;
                }

                table.Merges.Add(new CellMerge {
                    Row = r1,
                    Col = c1,
                    RowSpan = rowSpan,
                    ColSpan = colSpan
                });
            }

            return table.Merges.Count > 0;
        }

        private bool ReadHyperLinks(XlsxWorksheet sheet, DataTable table) {
            // Restart so hyperlink parsing does not depend on prior mergeCells scanning.
            if (!ResetSheetReader(sheet)) {
                return false;
            }

            if (m_xmlReader == null) {
                return false;
            }

            // Read Relationship Table
            Stream sheetRelStream = m_zipWorker.GetWorksheetRelsStream(sheet.Path);
            var hyperDict = new Dictionary<string, string>();
            if (sheetRelStream != null) {
                using (XmlReader reader = XmlReader.Create(sheetRelStream)) {
                    while (reader.Read()) {
                        if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XlsxWorkbook.N_rel) {
                            string rid = reader.GetAttribute(XlsxWorkbook.A_id);
                            Debug.Assert(rid != null);
                            hyperDict[rid] = reader.GetAttribute(XlsxWorkbook.A_target);
                        }
                    }
                    sheetRelStream.Close();
                }
            }

            var parser = new XmlPackageParser(m_xmlReader, m_namespaceUri);
            bool any = false;
            foreach (XmlHyperlinkPackage linkPackage in parser.ParseHyperlinks()) {
                any = true;
                string aref = linkPackage.Ref;
                string display = linkPackage.Display;
                string rid = linkPackage.RelationshipId;
                string location = linkPackage.Location;
                string hyperlink = display;

                if (!string.IsNullOrEmpty(rid) && hyperDict.ContainsKey(rid)) {
                    hyperlink = hyperDict[rid];
                }

                if (string.IsNullOrEmpty(aref)) {
                    continue;
                }

                var dim = new XlsxDimension(aref);
                int c1 = dim.FirstCol - 1;
                int r1 = dim.FirstRow - 1;
                int c2 = dim.LastCol - 1;
                int r2 = dim.LastRow - 1;
                if (c1 < 0 || r1 < 0) {
                    continue;
                }

                for (int row = r1; row <= r2; row++) {
                    if (row >= table.Rows.Count) {
                        break;
                    }
                    for (int col = c1; col <= c2; col++) {
                        if (col >= table.Rows[row].Count) {
                            break;
                        }
                        object value = table.Rows[row][col];
                        var cell = value as XlsCell;
                        if (cell == null) {
                            cell = new XlsCell(value);
                        }
                        cell.SetHyperLink(hyperlink, location);
                        table.Rows[row][col] = cell;
                    }
                }
            }

            m_xmlReader.Close();
            if (m_sheetStream != null) {
                m_sheetStream.Close();
            }

            return any;
        }

        private bool IsDateTimeStyle(int styleId) {
            return m_defaultDateTimeStyles.Contains(styleId);
        }

        /// <summary>
        /// Grow the current row buffer (and sheet column count) when a cell appears
        /// beyond the width guessed during DetectDemension.
        /// </summary>
        private void EnsureRowCapacity(XlsxWorksheet sheet, int col1Based) {
            int needed = col1Based;
            if (m_cellsValues != null && m_cellsValues.Length >= needed) {
                return;
            }
            int oldLen = m_cellsValues == null ? 0 : m_cellsValues.Length;
            var grown = new object[needed];
            if (m_cellsValues != null && oldLen > 0) {
                Array.Copy(m_cellsValues, grown, oldLen);
            }
            m_cellsValues = grown;
            if (sheet.Dimension != null && sheet.Dimension.LastCol < needed) {
                sheet.Dimension.LastCol = needed;
            }
        }

        private static void EnsureTableWidth(DataTable table, int width) {
            while (table.Columns.Count < width) {
                int i = table.Columns.Count;
                table.Columns.Add(i.ToString(CultureInfo.InvariantCulture), typeof(Object));
            }
            // Pad earlier rows so ItemArray length matches column count.
            for (int r = 0; r < table.Rows.Count; r++) {
                object[] old = table.Rows[r].ItemArray;
                if (old != null && old.Length >= width) {
                    continue;
                }
                var padded = new object[width];
                if (old != null && old.Length > 0) {
                    Array.Copy(old, padded, old.Length);
                }
                table.Rows[r].ItemArray = padded;
            }
        }

        private Dictionary<int, XlsxDimension> DetectDemension() {
            var dict = new Dictionary<int, XlsxDimension>();
            for (int sheetIndex = 0; sheetIndex < m_workbook.Sheets.Count; sheetIndex++) {
                XlsxWorksheet sheet = m_workbook.Sheets[sheetIndex];

                ReadSheetGlobals(sheet);

                if (sheet.Dimension != null) {
                    m_depth = 0;
                    m_emptyRowCount = 0;

                    // Sample the first 100 rows to trim huge empty trailing dimensions
                    // cheaply. Columns that only appear later are still kept via
                    // EnsureRowCapacity during the full read (#14 + perf).
                    int detectRows = Math.Min(sheet.Dimension.LastRow, 100);
                    int maxColumnCount = 0;
                    while (detectRows > 0) {
                        if (!ReadSheetRow(sheet)) {
                            break;
                        }
                        maxColumnCount = Math.Max(LastIndexOfNonNull(m_cellsValues) + 1, maxColumnCount);
                        detectRows--;
                    }

                    // 
                    if (maxColumnCount < sheet.Dimension.LastCol) {
                        dict[sheetIndex] = new XlsxDimension(sheet.Dimension.LastRow, maxColumnCount);
                    } else {
                        dict[sheetIndex] = sheet.Dimension;
                    }
                } else {
                    dict[sheetIndex] = sheet.Dimension;
                }
            }
            return dict;
        }

        private static int LastIndexOfNonNull(object[] cellsValues) {
            for (int i = cellsValues.Length - 1; i >= 0; i--) {
                if (cellsValues[i] != null) {
                    return i;
                }
            }
            return 0;
        }

        private DataSet ReadDataSet() {
            var dataset = new DataSet();

            Dictionary<int, XlsxDimension> demensionDict = DetectDemension();

            for (int sheetIndex = 0; sheetIndex < m_workbook.Sheets.Count; sheetIndex++) {
                XlsxWorksheet sheet = m_workbook.Sheets[sheetIndex];
                var table = new DataTable(m_workbook.Sheets[sheetIndex].Name);

                ReadSheetGlobals(sheet);
                sheet.Dimension = demensionDict[sheetIndex];

                if (sheet.Dimension == null) {
                    continue;
                }

                m_depth = 0;
                m_emptyRowCount = 0;

                // Reada Columns
                for (int i = 0; i < sheet.ColumnsCount; i++) {
                    table.Columns.Add(i.ToString(CultureInfo.InvariantCulture), typeof(Object));
                }

                // Read Sheet Rows
                table.BeginLoadData();
                while (ReadSheetRow(sheet)) {
                    int width = m_cellsValues == null ? 0 : m_cellsValues.Length;
                    EnsureTableWidth(table, width);
                    object[] rowValues = m_cellsValues;
                    if (rowValues != null && rowValues.Length < table.Columns.Count) {
                        var padded = new object[table.Columns.Count];
                        Array.Copy(rowValues, padded, rowValues.Length);
                        rowValues = padded;
                    }
                    table.Rows.Add(rowValues);
                }

                if (table.Rows.Count > 0) {
                    dataset.Tables.Add(table);
                }

                // Read merged cells (before hyperlinks — both appear after sheetData)
                ReadMergeCells(sheet, table);

                // Read HyperLinks
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
                    if (m_xmlReader != null)
                        ((IDisposable)m_xmlReader).Dispose();
                    if (m_sheetStream != null)
                        m_sheetStream.Dispose();
                    if (m_zipWorker != null && m_ownsZipWorker)
                        m_zipWorker.Dispose();
                }

                m_zipWorker = null;
                m_rowPackages = null;
                m_xmlReader = null;
                m_sheetStream = null;

                m_workbook = null;
                m_cellsValues = null;
                m_savedCellsValues = null;

                disposed = true;
            }
        }

        ~ExcelOpenXmlReader() {
            Dispose(false);
        }

        #endregion
    }
}