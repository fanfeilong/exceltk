using System.Text.RegularExpressions;
using ExcelToolKit.Util;

namespace ExcelToolKit {
    public class XlsCell {
        private readonly object m_object;
        private string m_hyperLink;
        private bool m_isHyperLink;
        private string prepareMarkDown;

        public XlsCell(object obj) {
            m_object=obj;
            prepareMarkDown = null;
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
                if(m_isHyperLink){
                    if(prepareMarkDown!=null){
                        return prepareMarkDown;
                    }else{
                        return string.Format("[{0}]({1})", m_object, HttpUtility.UrlPathEncode(m_hyperLink));
                    }
                }else{
                    return m_object.ToString();
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

        public void SetHyperLink(string url){
            url = FilteUrl(url);

            m_isHyperLink=true;
            if(url=="_blank"){
                var text=m_object.ToString();
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
                var text=m_object.ToString();
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
            return m_object.ToString();
        }
    }
}