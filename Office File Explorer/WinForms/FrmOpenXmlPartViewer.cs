using System.IO.Packaging;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Xml.Linq;
using System;
using Office_File_Explorer.Helpers;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmOpenXmlPartViewer : Form
    {
        List<PackagePart> pParts = new List<PackagePart>();
        PackagePart currentPart;
        Package package;

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

        public void EnableIcons()
        {
            toolStripButtonGenerateCallbacks.Enabled = true;
            toolStripButtonValidateXml.Enabled = true;
            toolStripDropDownButton1.Enabled = true;
        }

        public void DisableIcons()
        {
            toolStripDropDownButton1.Enabled = false;
            toolStripButtonSave.Enabled = false;
            toolStripButtonGenerateCallbacks.Enabled = false;
            toolStripButtonValidateXml.Enabled = false;
            toolStripButtonInsertIcon.Enabled = false;
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
                        //EnableIcons();
                    }
                    else
                    {
                        DisableIcons();
                    }

                    foreach (PackagePart pp in pParts)
                    {
                        if (pp.Uri.ToString() == treeView1.SelectedNode.Text)
                        {
                            currentPart = pp;
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
    }
}
