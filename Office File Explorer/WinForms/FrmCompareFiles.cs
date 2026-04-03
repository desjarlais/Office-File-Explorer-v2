using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmCompareFiles : Form
    {
        private static readonly string FileFilter =
            "Open XML Files | *.docx; *.dotx; *.docm; *.dotm; *.xlsx; *.xlsm; *.xlst; *.xltm; *.pptx; *.pptm; *.potx; *.potm|" +
            "Binary Office Documents | *.doc; *.dot; *.xls; *.xlt; *.ppt; *.pot|" +
            "All Files | *.*";

        public FrmCompareFiles()
        {
            InitializeComponent();
        }

        private Dictionary<string, string> _leftParts = [];
        private Dictionary<string, string> _rightParts = [];

        private void BtnFileLeft_Click(object sender, EventArgs e)
        {
            _leftParts = OpenFileAndPopulateTree(label1, tvLeft);
        }

        private void BtnFileRight_Click(object sender, EventArgs e)
        {
            _rightParts = OpenFileAndPopulateTree(label2, tvRight);
        }

        private Dictionary<string, string> OpenFileAndPopulateTree(Label pathLabel, TreeView treeView)
        {
            using OpenFileDialog fDialog = new OpenFileDialog
            {
                Title = "Select Office File",
                Filter = FileFilter,
                RestoreDirectory = true,
                InitialDirectory = @"%userprofile%"
            };

            if (fDialog.ShowDialog() != DialogResult.OK)
            {
                return [];
            }

            string filePath = fDialog.FileName;
            pathLabel.Text = filePath;

            treeView.Nodes.Clear();

            try
            {
                using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
                TreeNode root = new TreeNode(filePath);
                Dictionary<string, string> parts = [];

                foreach (PackagePart part in package.GetParts())
                {
                    string uri = part.Uri.ToString();
                    root.Nodes.Add(uri);
                    parts[uri] = ReadPartContent(part);
                }

                treeView.Nodes.Add(root);
                treeView.ExpandAll();

                return parts;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to open file: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return [];
            }
        }

        private static string ReadPartContent(PackagePart part)
        {
            try
            {
                using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
                if (part.ContentType.Contains("xml") || part.ContentType.Contains("rels"))
                {
                    using StreamReader sr = new StreamReader(stream);
                    string raw = sr.ReadToEnd();

                    try
                    {
                        XmlDocument doc = new XmlDocument();
                        doc.LoadXml(raw);
                        using MemoryStream ms = new MemoryStream();
                        XmlWriterSettings settings = new XmlWriterSettings
                        {
                            Encoding = new UTF8Encoding(false),
                            Indent = true,
                            IndentChars = "  ",
                            NewLineChars = "\r\n",
                            NewLineHandling = NewLineHandling.Replace
                        };
                        using (XmlWriter writer = XmlWriter.Create(ms, settings))
                        {
                            doc.Save(writer);
                        }
                        return Encoding.UTF8.GetString(ms.ToArray());
                    }
                    catch
                    {
                        return raw;
                    }
                }
                else
                {
                    using MemoryStream ms = new MemoryStream();
                    stream.CopyTo(ms);
                    return Convert.ToHexString(ms.ToArray());
                }
            }
            catch (Exception ex)
            {
                return "[Error reading part: " + ex.Message + "]";
            }
        }

        private void TvLeft_AfterSelect(object sender, TreeViewEventArgs e)
        {
            string uri = e.Node?.Text ?? string.Empty;
            _leftParts.TryGetValue(uri, out string leftText);
            scintillaDiffControl1.TextLeft = leftText ?? string.Empty;
        }

        private void TvRight_AfterSelect(object sender, TreeViewEventArgs e)
        {
            string uri = e.Node?.Text ?? string.Empty;
            _rightParts.TryGetValue(uri, out string rightText);
            scintillaDiffControl1.TextRight = rightText ?? string.Empty;
        }
    }
}
