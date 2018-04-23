using System;
using System.Collections.Generic;
using System.Windows.Forms;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit;

namespace RichEditDOCVARIABLEBasics {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();

            List<DetailInfo> details = new List<DetailInfo>();

            details.Add(new DetailInfo(1, "Detail1"));
            details.Add(new DetailInfo(2, "Detail2"));

            richEditControl1.Options.MailMerge.DataSource = details;

            richEditControl1.Document.CalculateDocumentVariable += new DevExpress.XtraRichEdit.CalculateDocumentVariableEventHandler(Document_CalculateDocumentVariable);

            richEditControl1.LoadDocument("Template.rtf");
            ShowFieldCodes();
        }

        private void button1_Click(object sender, EventArgs e) {
            richEditControl1.LoadDocument("Template.rtf");
            ShowFieldCodes();
        }

        private void button2_Click(object sender, EventArgs e) {
            MailMergeOptions myMergeOptions = richEditControl1.Document.CreateMailMergeOptions();
            myMergeOptions.MergeMode = MergeMode.NewParagraph;
  
            RichEditDocumentServer server = new RichEditDocumentServer();
            
            server.CalculateDocumentVariable += new CalculateDocumentVariableEventHandler(Document_CalculateDocumentVariable);
            richEditControl1.Document.MailMerge(myMergeOptions, server.Document);

            richEditControl1.CreateNewDocument();
            richEditControl1.Document.AppendDocumentContent(server.Document.Range);
        }

        void Document_CalculateDocumentVariable(object sender, DevExpress.XtraRichEdit.CalculateDocumentVariableEventArgs e) {
            int detailId = -1;

            if (Int32.TryParse(e.Arguments[0].Value, out detailId)) {
                RichEditDocumentServer server = new RichEditDocumentServer();

                string path = string.Format("{0}\\Detail{1}.rtf", System.IO.Directory.GetCurrentDirectory(), detailId.ToString());

                server.LoadDocument(path);

                e.Value = server;
                e.Handled = true;
            }
        }

        private void ShowFieldCodes() {
            Document doc = richEditControl1.Document;
            doc.BeginUpdate();
            foreach (Field f in doc.Fields) f.ShowCodes = true;
            doc.EndUpdate();
        }

    }

    public class DetailInfo {
        private int id;

        public int DetailId {
            get { return id; }
            set { id = value; }
        }
        private string description;

        public string Description {
            get { return description; }
            set { description = value; }
        }

        public DetailInfo(int id, string description) {
            this.DetailId = id;
            this.Description = description;
        }
    }

}