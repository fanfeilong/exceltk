using System.Text.RegularExpressions;
using Exceltk.Util;

using System;
using System.IO;

namespace Exceltk.Reader {

    public class HyperLinkIndex{
        public string Sheet{get;set;}
        public int Col{get;set;}
        public int Row{get;set;}
    }

    /// <summary>
    /// Merged cell region in 0-based coordinates (top-left + spans).
    /// </summary>
    public class CellMerge {
        public int Row{get;set;}
        public int Col{get;set;}
        public int RowSpan{get;set;}
        public int ColSpan{get;set;}
    }

    public class XlsCell {
        private readonly object m_object;
        private string m_hyperLink;
        private bool m_isHyperLink;
        private string prepareMarkDown;

        public XlsCell(object obj) {
            m_object=obj;
            prepareMarkDown = null;
            RowSpan=1;
            ColSpan=1;
        }

        public object Value {
            get {
                return m_object;
            }
        }

        public int RowSpan {
            get;
            set;
        }

        public int ColSpan {
            get;
            set;
        }

        /// <summary>
        /// True when this cell is covered by another cell's merge region
        /// and should be omitted from MultiMarkdown/HTML output.
        /// </summary>
        public bool IsMergeCovered {
            get;
            set;
        }

        public bool IsHyperLink {
            get {
                return m_isHyperLink;
            }
        }

        public string HyperLink {
            get {
                return m_hyperLink;
            }
        }

        private HyperLinkIndex m_hyperLinkIndex = null;
        public HyperLinkIndex HyperLinkIndex{
            set{
                m_hyperLinkIndex = value;
            }
            get{
                return m_hyperLinkIndex;
            }
        }

        public string GetMarkDownText(DataSet dataSet){
            string text=m_object==null ? "" : m_object.ToString();
            if(m_isHyperLink){
                if(prepareMarkDown!=null){
                    return prepareMarkDown;
                }else{
                    return string.Format("[{0}]({1})", text, HttpUtility.UrlPathEncode(m_hyperLink));
                }
            }else{
                if(m_hyperLinkIndex==null){
                    return text;
                }

                var table = dataSet.Tables[m_hyperLinkIndex.Sheet];
                //Console.WriteLine("{0}: <{1},{2}>, <{3},{4}>",table.TableName, m_hyperLinkIndex.Row, m_hyperLinkIndex.Col, table.Rows.Count, table.Columns.Count);

                string url = null;
                try{
                    var ro = table.Rows[m_hyperLinkIndex.Row-1].ItemArray[m_hyperLinkIndex.Col-1];

                    var rc = ro as XlsCell;
                    if(rc==null){
                        url = ro==null ? null : ro.ToString();
                    }else{
                        url = rc.HyperLink;
                    }
                }catch{
                    //Console.WriteLine("ERROR:{0}: <{1},{2}>, <{3},{4}>",table.TableName, m_hyperLinkIndex.Row, m_hyperLinkIndex.Col, table.Rows.Count, table.Columns.Count);
                }
                
                if(url!=null){
                    this.SetHyperLink(url);
                }

                if(m_isHyperLink){
                    if(prepareMarkDown!=null){
                        return prepareMarkDown;
                    }else{
                        return string.Format("[{0}]({1})", text, HttpUtility.UrlPathEncode(m_hyperLink));
                    }
                }else{
                    return text;
                }
            }
        }

        public string FilteUrl(string url){
            const string regex = @"https?://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?";
            var m=Regex.Match(url, regex);
            if(m.Captures.Count>0){
                return m.Groups[0].Value;
            }else{
                return "_blank";
            }
        }

        public void SetHyperLink(string url, string location = null){
            url = FilteUrl(url);

            m_isHyperLink=true;
            if(url=="_blank"){
                var text=m_object==null ? "" : m_object.ToString();
                var regex = @"\[.*\]\((.*)\)";
                var m = Regex.Match(text,regex);
                if(m.Captures.Count>0){
                    m_hyperLink=m.Groups[1].Value;
                    prepareMarkDown=text;
                }else{
                    regex=@"https?://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?";
                    m = Regex.Match(text, regex);
                    if(m.Captures.Count>0){
                        m_hyperLink = m.Groups[0].Value;
                    }else{
                        m_isHyperLink = false;
                    }
                }
            }else{
                if (location != null) {
                    url += "#" + location;
                }
                var text=m_object==null ? "" : m_object.ToString();
                var regex = string.Format(@"\[.*\]\({0}\)",url);
                var m = Regex.Match(text,regex);
                if(m.Captures.Count>0){
                    m_hyperLink=url;
                    prepareMarkDown=text;
                }else{
                    m_hyperLink = url;
                }
            }
        }

        public override string ToString() {
            return m_object==null ? "" : m_object.ToString();
        }
    }
}