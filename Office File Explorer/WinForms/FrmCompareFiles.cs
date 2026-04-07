using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmCompareFiles : Form
    {
        private static readonly string FileFilter =
            "Open XML Files | *.docx; *.dotx; *.docm; *.dotm; *.xlsx; *.xlsm; *.xlst; *.xltm; *.pptx; *.pptm; *.potx; *.potm";

        public FrmCompareFiles()
        {
            InitializeComponent();
        }

        private Dictionary<string, string> _leftParts = [];
        private Dictionary<string, string> _rightParts = [];
        private bool _syncing;
        private string _fileSummary = string.Empty;

        private void BtnFileLeft_Click(object sender, EventArgs e)
        {
            _leftParts = OpenFileAndPopulateTree(tvLeft);
            HighlightDifferences();
        }

        private void BtnFileRight_Click(object sender, EventArgs e)
        {
            _rightParts = OpenFileAndPopulateTree(tvRight);
            HighlightDifferences();
        }

        private Dictionary<string, string> OpenFileAndPopulateTree(TreeView treeView)
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
            treeView.Nodes.Clear();

            try
            {
                Cursor = Cursors.WaitCursor;
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
            finally
            {
                Cursor = Cursors.Default;
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
            if (_syncing) return;

            try
            {
                Cursor = Cursors.WaitCursor;
                string uri = e.Node?.Text ?? string.Empty;
                _leftParts.TryGetValue(uri, out string leftText);
                scintillaDiffControl1.TextLeft = leftText ?? string.Empty;

                _syncing = true;
                SelectMatchingNode(tvRight, uri, _rightParts);
                _syncing = false;

                UpdateDiffCount();
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void TvRight_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (_syncing) return;

            try
            {
                Cursor = Cursors.WaitCursor;
                string uri = e.Node?.Text ?? string.Empty;
                _rightParts.TryGetValue(uri, out string rightText);
                scintillaDiffControl1.TextRight = rightText ?? string.Empty;

                _syncing = true;
                SelectMatchingNode(tvLeft, uri, _leftParts);
                _syncing = false;

                UpdateDiffCount();
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void SelectMatchingNode(TreeView targetTree, string nodeText, Dictionary<string, string> parts)
        {
            TreeNode match = FindNodeByText(targetTree.Nodes, nodeText);
            if (match is not null)
            {
                targetTree.SelectedNode = match;
                parts.TryGetValue(nodeText, out string text);
                if (targetTree == tvRight)
                    scintillaDiffControl1.TextRight = text ?? string.Empty;
                else
                    scintillaDiffControl1.TextLeft = text ?? string.Empty;
            }
            else
            {
                targetTree.SelectedNode = null;
                if (targetTree == tvRight)
                    scintillaDiffControl1.TextRight = string.Empty;
                else
                    scintillaDiffControl1.TextLeft = string.Empty;
            }
        }

        private void UpdateDiffCount()
        {
            string leftText = scintillaDiffControl1.TextLeft ?? string.Empty;
            string rightText = scintillaDiffControl1.TextRight ?? string.Empty;

            if (string.IsNullOrEmpty(leftText) && string.IsNullOrEmpty(rightText))
            {
                lblDiffCount.Text = "";
                return;
            }

            var diffModel = new SideBySideDiffBuilder().BuildDiffModel(leftText, rightText);
            int count = diffModel.OldText.Lines.Count(l => l.Type != ChangeType.Unchanged);
            string partDiff = count == 0 ? "Selected File: No differences" : $"Selected File: {count} difference(s)";
            lblDiffCount.Text = string.IsNullOrEmpty(_fileSummary) ? partDiff : _fileSummary + "  |  " + partDiff;
        }

        private static TreeNode FindNodeByText(TreeNodeCollection nodes, string text)
        {
            foreach (TreeNode node in nodes)
            {
                if (node.Text == text)
                    return node;

                TreeNode child = FindNodeByText(node.Nodes, text);
                if (child is not null)
                    return child;
            }

            return null;
        }

        /// <summary>
        /// Compare parts from both files and color-code tree nodes.
        /// Red = content differs, Blue = part only exists on one side, Default = identical.
        /// </summary>
        private void HighlightDifferences()
        {
            if (_leftParts.Count == 0 || _rightParts.Count == 0)
            {
                ResetNodeColors(tvLeft);
                ResetNodeColors(tvRight);
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;
                HighlightTreeNodes(tvLeft, _leftParts, _rightParts);
                HighlightTreeNodes(tvRight, _rightParts, _leftParts);

                tvLeft.ExpandAll();
                tvRight.ExpandAll();

                // Compute file-level change summary
                var allKeys = new HashSet<string>(_leftParts.Keys);
                allKeys.UnionWith(_rightParts.Keys);
                int modified = 0, uniqueLeft = 0, uniqueRight = 0, identical = 0;
                foreach (string key in allKeys)
                {
                    bool inLeft = _leftParts.ContainsKey(key);
                    bool inRight = _rightParts.ContainsKey(key);
                    if (inLeft && inRight)
                    {
                        if (!string.Equals(_leftParts[key], _rightParts[key], StringComparison.Ordinal))
                            modified++;
                        else
                            identical++;
                    }
                    else if (inLeft)
                        uniqueLeft++;
                    else
                        uniqueRight++;
                }

                if (modified == 0 && uniqueLeft == 0 && uniqueRight == 0)
                {
                    _fileSummary = "Total: Files are identical (" + identical + " parts)";
                }
                else
                {
                    var parts = new List<string>();
                    if (modified > 0) parts.Add(modified + " modified");
                    if (uniqueLeft > 0) parts.Add(uniqueLeft + " only in left");
                    if (uniqueRight > 0) parts.Add(uniqueRight + " only in right");
                    if (identical > 0) parts.Add(identical + " identical");
                    _fileSummary = "Total: " + string.Join(", ", parts);
                }

                lblDiffCount.Text = _fileSummary;

                // Auto-select the first part node so the diff view is immediately populated
                if (tvLeft.Nodes.Count > 0 && tvLeft.Nodes[0].Nodes.Count > 0)
                {
                    tvLeft.SelectedNode = tvLeft.Nodes[0].Nodes[0];
                }
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private static void HighlightTreeNodes(
            TreeView tree,
            Dictionary<string, string> thisParts,
            Dictionary<string, string> otherParts)
        {
            if (tree.Nodes.Count == 0)
                return;

            TreeNode root = tree.Nodes[0];
            bool rootHasDiffs = false;

            foreach (TreeNode node in root.Nodes)
            {
                string uri = node.Text;

                if (!otherParts.ContainsKey(uri))
                {
                    // part only exists on this side
                    node.ForeColor = Color.Blue;
                    node.ToolTipText = "Only in this file";
                    rootHasDiffs = true;
                }
                else if (!string.Equals(thisParts[uri], otherParts[uri], StringComparison.Ordinal))
                {
                    // part exists on both sides but content differs
                    node.ForeColor = Color.OrangeRed;
                    node.ToolTipText = "Content differs";
                    rootHasDiffs = true;
                }
                else
                {
                    // identical
                    node.ForeColor = tree.ForeColor;
                    node.ToolTipText = "Identical";
                }
            }

            root.ForeColor = rootHasDiffs ? Color.OrangeRed : tree.ForeColor;
            tree.ShowNodeToolTips = true;
        }

        private static void ResetNodeColors(TreeView tree)
        {
            foreach (TreeNode root in tree.Nodes)
            {
                root.ForeColor = tree.ForeColor;
                root.ToolTipText = string.Empty;
                foreach (TreeNode node in root.Nodes)
                {
                    node.ForeColor = tree.ForeColor;
                    node.ToolTipText = string.Empty;
                }
            }
        }
    }
}
