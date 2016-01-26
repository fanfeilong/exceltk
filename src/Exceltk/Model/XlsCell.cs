using System.Web;

namespace ExcelToolKit {
    public class XlsCell {
        private readonly object m_object;
        private string m_hyperLink;
        private bool m_isHyperLink;

        public XlsCell(object obj) {
            m_object=obj;
        }

        public object Value {
            get {
                return m_object;
            }
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

        public string MarkDownText {
            get {
                return IsHyperLink
                           ?string.Format("[{0}]({1})", m_object, HttpUtility.UrlPathEncode(m_hyperLink))
                           :m_object.ToString();
            }
        }

        public void SetHyperLink(string url) {
            m_isHyperLink=true;
            m_hyperLink=url;
        }

        public override string ToString() {
            return m_object.ToString();
        }
    }
}