using System.IO.Packaging;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Xml.Linq;
using System;
using Office_File_Explorer.Helpers;
using System.Xml.Schema;
using System.Text;
using System.Xml;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmOpenXmlPartViewer : Form
    {
        List<PackagePart> pParts = new List<PackagePart>();
        Package package;
        bool hasXmlError;

        public FrmOpenXmlPartViewer(string path)
        {
            InitializeComponent();

            package = Package.Open(path, FileMode.Open, FileAccess.ReadWrite);

            TreeNode tRoot = new TreeNode();
            tRoot.Text = path;

            foreach (PackagePart part in package.GetParts())
            {
                tRoot.Nodes.Add(part.Uri.ToString());
                pParts.Add(part);
            }

            treeView1.Nodes.Add(tRoot);
        }

        string FormatXml(string xml)
        {
            try
            {
                XDocument doc = XDocument.Parse(xml);
                return doc.ToString();
            }
            catch (Exception)
            {
                return xml;
            }
        }

        private void FrmOpenXmlPartViewer_FormClosing(object sender, FormClosingEventArgs e)
        {
            package.Close();
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            try
            {
                if (e.Node.Text.EndsWith(".xml") || e.Node.Text.EndsWith(".rels"))
                {
                    // for customui files, allow additional editing
                    if (e.Node.Text.EndsWith("customUI.xml") || e.Node.Text.EndsWith("customUI14.xml"))
                    {
                        EnableCustomUIIcons();
                    }
                    else
                    {
                        DisableCustomUIIcons();
                    }

                    foreach (PackagePart pp in pParts)
                    {
                        if (pp.Uri.ToString() == treeView1.SelectedNode.Text)
                        {
                            using (StreamReader sr = new StreamReader(pp.GetStream()))
                            {
                                string contents = sr.ReadToEnd();
                                rtbPartContents.Rtf = XmlColorizer.Colorize(contents);
                                rtbPartContents.Rtf = FormatXml(rtbPartContents.Rtf);
                            }
                        }
                    }
                }
                else
                {
                    rtbPartContents.Text = "No Viewer For File Type";
                }
            }
            catch (Exception ex)
            {
                rtbPartContents.Text = "Error: " + ex.Message;
            }
        }

        private void FrmOpenXmlPartViewer_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { Close(); }
        }

        private void toolStripButtonModifyXml_Click(object sender, EventArgs e)
        {
            EnableModifyUI();
        }

        private void toolStripButtonSave_Click(object sender, EventArgs e)
        {
            foreach (PackagePart pp in pParts)
            {
                if (pp.Uri.ToString() == treeView1.SelectedNode.Text)
                {
                    MemoryStream ms = new MemoryStream();
                    using (TextWriter tw = new StreamWriter(ms))
                    {
                        tw.Write(rtbPartContents.Text);
                        tw.Flush();

                        ms.Position = 0;
                        Stream partStream = pp.GetStream(FileMode.OpenOrCreate, FileAccess.Write);
                        partStream.SetLength(0);
                        ms.WriteTo(partStream);
                    }

                    break;
                }
            }

            package.Flush();

            // update ui
            DisableModifyUI();
        }

        public void EnableModifyUI()
        {
            rtbPartContents.ReadOnly = false;
            toolStripButtonSave.Enabled = true;
            toolStripButtonModifyXml.Enabled = false;
        }

        public void DisableModifyUI()
        {
            rtbPartContents.ReadOnly = true;
            toolStripButtonSave.Enabled = false;
            toolStripButtonModifyXml.Enabled = true;
        }

        public void EnableCustomUIIcons()
        {
            toolStripButtonGenerateCallbacks.Enabled = true;
            toolStripButtonValidateXml.Enabled = true;
            toolStripDropDownButton1.Enabled = true;
        }

        public void DisableCustomUIIcons()
        {
            toolStripDropDownButton1.Enabled = false;
            toolStripButtonSave.Enabled = false;
            toolStripButtonGenerateCallbacks.Enabled = false;
            toolStripButtonValidateXml.Enabled = false;
            toolStripButtonInsertIcon.Enabled = false;
        }

        private void ShowError(string errorText)
        {
            MessageBox.Show(this, errorText, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        static void ValidationCallback(object sender, ValidationEventArgs args)
        {
            if (args.Severity == XmlSeverityType.Warning)
            {
                Console.Write("WARNING: ");
            }
            else if (args.Severity == XmlSeverityType.Error)
            {
                Console.Write("ERROR: ");
            }
        }

        public bool ValidateXml(bool showValidMessage)
        {
            if (rtbPartContents.Text == null || rtbPartContents.Text.Length == 0)
            {
                return false;
            }

            rtbPartContents.SuspendLayout();

            try
            {
                XmlTextReader xtr = new XmlTextReader(@".\Schemas\customui14.xsd");
                XmlSchema schema = XmlSchema.Read(xtr, ValidationCallback);

                XmlDocument xmlDoc = new XmlDocument();

                if (schema == null)
                {
                    return false;
                }

                xmlDoc.Schemas.Add(schema);
                xmlDoc.LoadXml(rtbPartContents.Text);

                if (xmlDoc.DocumentElement.NamespaceURI.ToString() != schema.TargetNamespace)
                {
                    StringBuilder errorText = new StringBuilder();
                    errorText.Append("Unknown Namespace".Replace("|1", xmlDoc.DocumentElement.NamespaceURI.ToString()));
                    errorText.Append("\n" + "CustomUI Namespace".Replace("|1", schema.TargetNamespace));

                    ShowError(errorText.ToString());
                    return false;
                }

                hasXmlError = false;
                xmlDoc.Validate(XmlValidationEventHandler);
            }
            catch (XmlException ex)
            {
                ShowError("Invalid Xml" + "\n" + ex.Message);
                return false;
            }

            rtbPartContents.ResumeLayout();

            if (!hasXmlError)
            {
                if (showValidMessage)
                {
                    MessageBox.Show(this, "Valid Xml", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                return true;
            }
            return false;
        }

        private void XmlValidationEventHandler(object sender, ValidationEventArgs e)
        {
            lock (this)
            {
                hasXmlError = true;
            }
            MessageBox.Show(this, e.Message, e.Severity.ToString(), MessageBoxButtons.OK,
                (e.Severity == XmlSeverityType.Error ? MessageBoxIcon.Error : MessageBoxIcon.Warning));
        }

        private void toolStripButtonValidateXml_Click(object sender, EventArgs e)
        {
            ValidateXml(true);
        }

        private void toolStripButtonGenerateCallbacks_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuInsertO14CustomUI_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuInsertO12CustomUIPart_Click(object sender, EventArgs e)
        {

        }

        private void xmlToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rtbPartContents.Text = Strings.xmlCustomOutspace;
        }

        private void customTabToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rtbPartContents.Text = Strings.xmlCustomTab;
        }

        private void excelCustomTabToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rtbPartContents.Text = Strings.xmlExcelCustomTab;
        }

        private void repurposeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rtbPartContents.Text = Strings.xmlRepurpose;
        }

        private void wordGroupOnInsertTabToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rtbPartContents.Text = Strings.xmlWordGroupInsertTab;
        }

        private void toolStripButtonInsertIcon_Click(object sender, EventArgs e)
        {

        }
    }
}
