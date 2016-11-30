#if OS_WINDOWS
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Drawing;
using System.Reflection;
using Exceltk;
using Exceltk.Reader;

namespace Exceltk.Clipborad {
    public class ClipboradMonitor : System.Windows.Forms.Form {
        [DllImport("User32.dll")]
        protected static extern int SetClipboardViewer(int hWndNewViewer);

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern bool ChangeClipboardChain(IntPtr hWndRemove, IntPtr hWndNewNext);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SendMessage(IntPtr hwnd, int wMsg, IntPtr wParam, IntPtr lParam);

        IntPtr nextClipboardViewer;
        RichTextBox richTextBox1;
        RadioButton radioButtonMdHead;
        bool isChecked = false;
        bool display = false;

        /// <summary>
        /// Required designer variable.
        /// </summary>
        private readonly System.ComponentModel.Container components = null;

        public ClipboradMonitor() {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();
            nextClipboardViewer = (IntPtr)SetClipboardViewer((int)Handle);

            //
            // TODO: Add any constructor code after InitializeComponent call
            //
            
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing) {
            ChangeClipboardChain(Handle, nextClipboardViewer);
            if (disposing) {
                if (components != null) {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }


        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            richTextBox1 = new System.Windows.Forms.RichTextBox();
            radioButtonMdHead = new RadioButton();
            var tableLayout = new TableLayoutPanel();
            var panenl = new Panel();

            const int w = 292;
            const int h = 273;
            const int t = 20;

            SuspendLayout();

            tableLayout.Dock = DockStyle.Fill;
            tableLayout.RowCount = 2;
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayout.Location = new System.Drawing.Point(w, h);
            tableLayout.Name = "main";
            tableLayout.BackColor = System.Drawing.Color.White;
            tableLayout.CellBorderStyle = TableLayoutPanelCellBorderStyle.None;
            tableLayout.Margin = new Padding(0, 0, 0, 0);

            panenl.Dock = DockStyle.Fill;
            panenl.BorderStyle = BorderStyle.None;
            tableLayout.Controls.Add(panenl);

            // markdown options
            radioButtonMdHead.Location = new System.Drawing.Point(0, 0);
            radioButtonMdHead.Size = new System.Drawing.Size(30,t);
            radioButtonMdHead.Dock = DockStyle.None;
            radioButtonMdHead.Anchor = AnchorStyles.Left|AnchorStyles.Top;
            radioButtonMdHead.BackColor = System.Drawing.Color.LightYellow;
            radioButtonMdHead.Text = "th";
            radioButtonMdHead.CheckedChanged += (s, e) => {
                isChecked = radioButtonMdHead.Checked;
            };
            radioButtonMdHead.Click += (s, e) => {
                if (radioButtonMdHead.Checked && !isChecked) {
                    radioButtonMdHead.Checked = false;
                } else {
                    radioButtonMdHead.Checked = true;
                    isChecked = false;
                }
            };
            panenl.Controls.Add(radioButtonMdHead);

            // richTextBox1
            richTextBox1.BorderStyle = BorderStyle.None;
            richTextBox1.Dock = DockStyle.Fill;
            richTextBox1.Location = new System.Drawing.Point(0, t);
            richTextBox1.Name = "markdownviewer";
            richTextBox1.ReadOnly = true;
            richTextBox1.Size = new System.Drawing.Size(w, h-t);
            richTextBox1.TabIndex = 0;
            richTextBox1.Text = "";
            richTextBox1.WordWrap = false;
            richTextBox1.BackColor = System.Drawing.Color.LightYellow;
            tableLayout.Controls.Add(richTextBox1);

            // Form
            AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            ClientSize = new System.Drawing.Size(w, h);
            Controls.Add(tableLayout);
            Name = "Exceltk";
            Text = "Exceltk - github.com/fanfeilong";

            Icon = Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location);

            ResumeLayout(false);
            
        }

        protected override void WndProc(ref System.Windows.Forms.Message m) {
            // defined in winuser.h
            const int WM_DRAWCLIPBOARD = 0x308;
            const int WM_CHANGECBCHAIN = 0x030D;

            switch (m.Msg) {
                case WM_DRAWCLIPBOARD:
                    ParseClipborad();
                    SendMessage(nextClipboardViewer, m.Msg, m.WParam, m.LParam);
                    break;

                case WM_CHANGECBCHAIN:
                    if (m.WParam == nextClipboardViewer)
                        nextClipboardViewer = m.LParam;
                    else
                        SendMessage(nextClipboardViewer, m.Msg, m.WParam, m.LParam);
                    break;

                default:
                    base.WndProc(ref m);
                    break;
            }
        }

        protected override void OnShown(EventArgs e) {
            base.OnShown(e);
            display = true;
        }

        void ParseClipborad() {

            if (!display) {
                return;
            }

            try {
                IDataObject iData = Clipboard.GetDataObject();
                if (iData!=null) {

                    if (iData.GetDataPresent(DataFormats.Html)){
                        var src = iData.GetData("Html Format");
                        OnChangeClipboardHtml(src);
                    } 
                }
            } catch (Exception e) {
                MessageBox.Show(e.ToString());
            }
        }

        void OnChangeClipboardHtml(object src) {
            var dataTable = GetTableFromHtml(src);
            if (dataTable != null) {
                var md = dataTable.ToMd(radioButtonMdHead.Checked);
                richTextBox1.Text = md;
                SetClipboard(md);
            }
        }

        void OnChangeClipboardRtf(object src) {
            var dataTable = GetTableFromRtfString(src as string);
            var md = dataTable.ToMd();
            richTextBox1.Text = md;
        }

        void OnChangeClipboardText(string src) {
            richTextBox1.Text = src;
        }

        static DataTable GetTableFromHtml(object src) {
            var dataSet = src.ParseDataSet();
            if (dataSet == null) {
                return null;
            }

            if (dataSet.Tables.Count > 0) {
                return dataSet.Tables[0];
            } else {
                return null;
            }
        }

        DataTable GetTableFromRtfString(string src) {

            File.WriteAllText("d:\\1.txt", src);

            int rowEnd = 0;
            int rowStart = 0;

            var dataTable = new DataTable("");
            bool firstRow = true;

            do {
                rowEnd = src.IndexOf(@"\row", rowEnd, StringComparison.OrdinalIgnoreCase);
                if (rowEnd < 0)
                    break;
                else if (src[rowEnd - 1] == '\\') {
                    rowEnd++;
                    continue;
                }

                rowStart = src.LastIndexOf(@"\trowd", rowEnd, StringComparison.OrdinalIgnoreCase);
                if (rowStart < 0)
                    break;
                else if (src[rowStart - 1] == '\\') {
                    rowEnd++;
                    continue;
                }

                string row = src.Substring(rowStart, rowEnd - rowStart);
                rowEnd++;

                int cellEnd = 0;
                int cellStart = 0;
                var dataRow = new List<string>();
                do {
                    cellEnd = row.IndexOf(@"\cell ", cellEnd, StringComparison.OrdinalIgnoreCase);
                    if (cellEnd < 0)
                        break;
                    else if (row[cellEnd - 1] == '\\') {
                        cellEnd++;
                        continue;
                    }

                    var cell = row.Substring(cellStart, cellEnd - cellStart);
                    var cellValue = CellValue(cell);
                    dataRow.Add(cellValue);

                    cellStart = cellEnd;
                    cellEnd++;
                }
                while (cellStart > 0);

                if (firstRow) {
                    firstRow = false;
                    foreach (string ColName in dataRow)
                        dataTable.Columns.Add(ColName);
                } else
                    dataTable.Rows.Add(dataRow.ToArray());
            }
            while ((rowStart > 0) && (rowEnd > 0));

            return dataTable;
        }

        static string CellValue(string cell) {
            int start = 0;

            while (start < cell.Length) {
                start = cell.IndexOf('\\', start);
                if (start < 0)
                    break;
                if (cell[start + 1] == '\\') {
                    cell = cell.Remove(start, 1);   //1 offset to erase space
                    start++; //skip "\"
                } else {
                    var end = cell.IndexOf(' ', start);
                    if (end < 0)
                        if (cell.Length > 0)
                            end = cell.Length - 1;
                        else
                            break;
                    cell = cell.Remove(start, end - start + 1);   //1 offset to erase space
                }

            }

            //Erase spaces at the end of the cell info.
            while (cell.Length > 0 && cell[cell.Length - 1] == ' ') {
                cell = cell.Remove(cell.Length - 1);
            }

            //Erase spaces at the beginning of the cell info.
            while (cell.Length > 0&&cell[0] == ' ') {
                cell = cell.Substring(1, cell.Length - 1);
            }

            return cell;
        }

        private static void SetClipboard(string value) {
            var dataObject = new DataObject();
            var bytes = Encoding.Unicode.GetBytes(value);
            var stream = new MemoryStream(bytes);
            dataObject.SetData("UnicodeText", false, stream);
            Clipboard.SetDataObject(dataObject, true);
        }
    }
}
#endif