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

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
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

        private void FrmOpenXmlPartViewer_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { Close(); }
        }
    }
}
