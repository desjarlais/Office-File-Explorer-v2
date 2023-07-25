using System.IO.Packaging;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System;
using Office_File_Explorer.Helpers;
using System.Xml.Schema;
using System.Text;
using System.Xml;
using System.Diagnostics;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmOpenXmlPartViewer : Form
    {
        List<PackagePart> pParts = new List<PackagePart>();
        Package package;
        bool hasXmlError;

        public enum OpenXmlInnerFileTypes
        {
            Word,
            Excel,
            PowerPoint,
            XML,
            Image,
            Binary,
            Video,
            Audio,
            Text,
            Other
        }

        public FrmOpenXmlPartViewer(string path)
        {
            InitializeComponent();

            package = Package.Open(path, FileMode.Open, FileAccess.ReadWrite);

            TreeNode tRoot = new TreeNode();
            tRoot.Text = path;

            // update app icon
            if (GetFileType(path) == OpenXmlInnerFileTypes.Word)
            {
                treeView1.SelectedImageIndex = 0;
            }
            else if (GetFileType(path) == OpenXmlInnerFileTypes.Excel)
            {
                treeView1.SelectedImageIndex = 2;
            }
            else if (GetFileType(path) == OpenXmlInnerFileTypes.PowerPoint)
            {
                treeView1.SelectedImageIndex = 1;
            }

            foreach (PackagePart part in package.GetParts())
            {
                tRoot.Nodes.Add(part.Uri.ToString());

                // update file icon, need to update both the selected and normal image index
                if (GetFileType(part.Uri.ToString()) == OpenXmlInnerFileTypes.XML)
                {
                    tRoot.Nodes[tRoot.Nodes.Count - 1].ImageIndex = 3;
                    tRoot.Nodes[tRoot.Nodes.Count - 1].SelectedImageIndex = 3;
                }
                else if (GetFileType(part.Uri.ToString()) == OpenXmlInnerFileTypes.Image)
                {
                    tRoot.Nodes[tRoot.Nodes.Count - 1].ImageIndex = 4;
                    tRoot.Nodes[tRoot.Nodes.Count - 1].SelectedImageIndex = 4;
                }
                else
                {
                    tRoot.Nodes[tRoot.Nodes.Count - 1].ImageIndex = 5;
                    tRoot.Nodes[tRoot.Nodes.Count - 1].SelectedImageIndex = 5;
                }

                pParts.Add(part);
            }

            treeView1.Nodes.Add(tRoot);
        }

        private void FrmOpenXmlPartViewer_FormClosing(object sender, FormClosingEventArgs e)
        {
            package.Close();
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            try
            {
                // currently only displaying xml and rel file types
                if (GetFileType(e.Node.Text) == OpenXmlInnerFileTypes.XML)
                {
                    // customui files have additional editing options
                    if (e.Node.Text.EndsWith("customUI.xml") || e.Node.Text.EndsWith("customUI14.xml"))
                    {
                        EnableCustomUIIcons();
                    }
                    else
                    {
                        DisableCustomUIIcons();
                    }

                    // load file contents and format the xml
                    foreach (PackagePart pp in pParts)
                    {
                        if (pp.Uri.ToString() == treeView1.SelectedNode.Text)
                        {
                            using (StreamReader sr = new StreamReader(pp.GetStream()))
                            {
                                string contents = sr.ReadToEnd();
                                rtbPartContents.Rtf = XmlColorizer.Colorize(contents);
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

        public OpenXmlInnerFileTypes GetFileType(string path)
        {
            switch (Path.GetExtension(path))
            {
                case ".docx":
                case ".dotx":
                case ".dotm":
                case ".docm":
                    return OpenXmlInnerFileTypes.Word;
                case ".xlsx":
                case ".xlsm":
                case ".xltm":
                case ".xltx":
                case ".xlsb":  
                    return OpenXmlInnerFileTypes.Excel;
                case ".pptx":
                case ".pptm":
                case ".ppsx":
                case ".ppsm":
                case ".potx":
                case ".potm":
                    return OpenXmlInnerFileTypes.PowerPoint;
                case ".jpeg":
                case ".jpg":
                case ".bmp":
                case ".png":
                case ".gif":
                case ".emf":
                case ".wmf":
                    return OpenXmlInnerFileTypes.Image;
                case ".xml":
                case ".rels":
                    return OpenXmlInnerFileTypes.XML;
                case ".bin":
                    return OpenXmlInnerFileTypes.Binary;
                case ".mp4":
                case ".avi":
                case ".wmv":
                case ".mov":
                    return OpenXmlInnerFileTypes.Video;
                case ".mp3":
                case ".wav":
                case ".wma":
                    return OpenXmlInnerFileTypes.Audio;
                case ".txt":
                    return OpenXmlInnerFileTypes.Text;
                default:
                    return OpenXmlInnerFileTypes.Binary;
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

        /// <summary>
        /// write the modified xml back to the package
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// use the schema to validate the xml
        /// </summary>
        /// <param name="showValidMessage"></param>
        /// <returns></returns>
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

        private void AddPart(XMLParts partType)
        {
            OfficePart newPart = CreateCustomUIPart(partType);
            TreeNode partNode = ConstructPartNode(newPart);
            TreeNode currentNode = treeView1.Nodes[0];
            if (currentNode == null) return;

            treeView1.SuspendLayout();
            currentNode.Nodes.Add(partNode);
            rtbPartContents.Text = string.Empty;
            treeView1.SelectedNode = partNode;
            treeView1.ResumeLayout();
        }

        private TreeNode ConstructPartNode(OfficePart part)
        {
            TreeNode node = new TreeNode(part.Name);
            node.Tag = part.PartType;
            node.ImageIndex = 3;
            node.SelectedImageIndex = 3;
            return node;
        }

        private OfficePart CreateCustomUIPart(XMLParts partType)
        {
            string relativePath;
            string relType;

            switch (partType)
            {
                case XMLParts.RibbonX12:
                    relativePath = "/customUI/customUI.xml";
                    relType = Strings.CustomUIPartRelType;
                    break;
                case XMLParts.RibbonX14:
                    relativePath = "/customUI/customUI14.xml";
                    relType = Strings.CustomUI14PartRelType;
                    break;
                case XMLParts.QAT12:
                    relativePath = "/customUI/qat.xml";
                    relType = Strings.QATPartRelType;
                    break;
                default:
                    return null;
            }

            Uri customUIUri = new Uri(relativePath, UriKind.Relative);
            PackageRelationship relationship = package.CreateRelationship(customUIUri, TargetMode.Internal, relType);

            OfficePart part = null;
            if (!package.PartExists(customUIUri))
            {
                part = new OfficePart(package.CreatePart(customUIUri, "application/xml"), partType, relationship.Id);
            }
            else
            {
                part = new OfficePart(package.GetPart(customUIUri), partType, relationship.Id);
            }

            return part;
        }

        private void toolStripButtonValidateXml_Click(object sender, EventArgs e)
        {
            ValidateXml(true);
        }

        private void toolStripButtonGenerateCallbacks_Click(object sender, EventArgs e)
        {
            // if there is no callback , then there is no point in generating the callback code
            if (rtbPartContents.Text == null || rtbPartContents.Text.Length == 0)
            {
                return;
            }

            // if the xml is not valid, then there is no point in generating the callback code
            if (!ValidateXml(false))
            {
                return;
            }

            // if we have valid xml, then generate the callback code
            try
            {
                XmlDocument customUI = new XmlDocument();
                customUI.LoadXml(rtbPartContents.Text);
                StringBuilder callbacks = CallbackBuilder.GenerateCallback(customUI);

                // show the callback code in a new window
                FrmCallbackViewer fCallbacks = new FrmCallbackViewer(callbacks)
                {
                    Owner = this
                };
                fCallbacks.ShowDialog();

                if (callbacks == null || callbacks.Length == 0)
                {
                    MessageBox.Show(this, "No callbacks found", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void toolStripMenuInsertO14CustomUI_Click(object sender, EventArgs e)
        {
            AddPart(XMLParts.RibbonX14);
        }

        private void toolStripMenuInsertO12CustomUIPart_Click(object sender, EventArgs e)
        {
            AddPart(XMLParts.RibbonX12);
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
