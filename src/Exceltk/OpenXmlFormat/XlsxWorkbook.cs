using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace ExcelToolKit.OpenXmlFormat {
    internal class XlsxWorkbook {
        public const string N_sheet="sheet";
        public const string N_t="t";
        public const string N_si="si";
        public const string N_cellXfs="cellXfs";
        public const string N_numFmts="numFmts";
        public const string A_sheetId="sheetId";
        public const string A_name="name";
        public const string A_rid="r:id";
        public const string N_rel="Relationship";
        public const string A_id="Id";
        public const string A_target="Target";
        private XlsxSST _SST;
        private XlsxStyles _Styles;
        private List<XlsxWorksheet> sheets;

        private XlsxWorkbook() {
        }

        public XlsxWorkbook(Stream workbookStream, Stream relsStream, Stream sharedStringsStream, Stream stylesStream) {
            if (null==workbookStream)
                throw new ArgumentNullException();

            ReadWorkbook(workbookStream);

            ReadWorkbookRels(relsStream);

            ReadSharedStrings(sharedStringsStream);

            ReadStyles(stylesStream);
        }

        public List<XlsxWorksheet> Sheets {
            get {
                return sheets;
            }
            set {
                sheets=value;
            }
        }

        public XlsxSST SST {
            get {
                return _SST;
            }
        }

        public XlsxStyles Styles {
            get {
                return _Styles;
            }
        }


        private void ReadStyles(Stream xmlFileStream) {
            if (null==xmlFileStream)
                return;

            _Styles=new XlsxStyles();

            bool rXlsxNumFmt=false;

            using (XmlReader reader=XmlReader.Create(xmlFileStream)) {
                while (reader.Read()) {
                    if (!rXlsxNumFmt&&reader.NodeType==XmlNodeType.Element&&reader.LocalName==N_numFmts) {
                        while (reader.Read()) {
                            if (reader.NodeType==XmlNodeType.Element&&reader.Depth==1)
                                break;

                            if (reader.NodeType==XmlNodeType.Element&&reader.LocalName==XlsxNumFmt.N_numFmt) {
                                _Styles.NumFmts.Add(
                                    new XlsxNumFmt(
                                        int.Parse(reader.GetAttribute(XlsxNumFmt.A_numFmtId)),
                                        reader.GetAttribute(XlsxNumFmt.A_formatCode)
                                        ));
                            }
                        }

                        rXlsxNumFmt=true;
                    }

                    if (reader.NodeType==XmlNodeType.Element&&reader.LocalName==N_cellXfs) {
                        while (reader.Read()) {
                            if (reader.NodeType==XmlNodeType.Element&&reader.Depth==1)
                                break;

                            if (reader.NodeType==XmlNodeType.Element&&reader.LocalName==XlsxXf.N_xf) {
                                string xfId=reader.GetAttribute(XlsxXf.A_xfId);
                                string numFmtId=reader.GetAttribute(XlsxXf.A_numFmtId);

                                _Styles.CellXfs.Add(
                                    new XlsxXf(
                                        xfId==null?-1:int.Parse(xfId),
                                        numFmtId==null?-1:int.Parse(numFmtId),
                                        reader.GetAttribute(XlsxXf.A_applyNumberFormat)
                                        ));
                            }
                        }

                        break;
                    }
                }

                xmlFileStream.Close();
            }
        }

        private void ReadSharedStrings(Stream xmlFileStream) {
            if (null==xmlFileStream)
                return;

            _SST=new XlsxSST();

            using (XmlReader reader=XmlReader.Create(xmlFileStream)) {
                // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                bool bAddStringItem=false;
                string sStringItem="";

                while (reader.Read()) {
                    // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                    if (reader.NodeType==XmlNodeType.Element&&reader.LocalName==N_si) {
                        // Do not add the string item until the next string item is read.
                        if (bAddStringItem) {
                            // Add the string item to XlsxSST.
                            _SST.Add(sStringItem);
                        } else {
                            // Add the string items from here on.
                            bAddStringItem=true;
                        }

                        // Reset the string item.
                        sStringItem="";
                    }

                    if (reader.NodeType==XmlNodeType.Element&&reader.LocalName==N_t) {
                        // Append to the string item.
                        sStringItem+=reader.ReadElementContentAsString();
                    }
                }
                // Do not add the last string item unless we have read previous string items.
                if (bAddStringItem) {
                    // Add the string item to XlsxSST.
                    _SST.Add(sStringItem);
                }

                xmlFileStream.Close();
            }
        }


        private void ReadWorkbook(Stream xmlFileStream) {
            sheets=new List<XlsxWorksheet>();

            using (XmlReader reader=XmlReader.Create(xmlFileStream)) {
                while (reader.Read()) {
                    if (reader.NodeType==XmlNodeType.Element&&reader.LocalName==N_sheet) {
                        sheets.Add(new XlsxWorksheet(
                                       reader.GetAttribute(A_name),
                                       int.Parse(reader.GetAttribute(A_sheetId)), reader.GetAttribute(A_rid)));
                    }
                }

                xmlFileStream.Close();
            }
        }

        private void ReadWorkbookRels(Stream xmlFileStream) {
            using (XmlReader reader=XmlReader.Create(xmlFileStream)) {
                while (reader.Read()) {
                    if (reader.NodeType==XmlNodeType.Element&&reader.LocalName==N_rel) {
                        string rid=reader.GetAttribute(A_id);

                        for (int i=0; i<sheets.Count; i++) {
                            XlsxWorksheet tempSheet=sheets[i];

                            if (tempSheet.RID==rid) {
                                tempSheet.Path=reader.GetAttribute(A_target);
                                sheets[i]=tempSheet;
                                break;
                            }
                        }
                    }
                }

                xmlFileStream.Close();
            }
        }
    }
}