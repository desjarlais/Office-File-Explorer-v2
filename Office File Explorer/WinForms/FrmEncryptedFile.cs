// the code for opening the structured compound file format was mostly taken from here https://github.com/ironfede/openmcdf

using Office_File_Explorer.Helpers;
using Office_File_Explorer.OpenMcdf;
using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Xml;
using System.Xml.Schema;
using System.Reflection;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmEncryptedFile : Form
    {
        private CompoundFile cf;
        private CFStream cfStream;
        static bool isValid;
        public static List<string> validationErrors = new List<string>();

        public FrmEncryptedFile(FileStream fs, bool enableCommit)
        {
            InitializeComponent();

            // Load images
            tvEncryptedContents.ImageList = new ImageList();
            tvEncryptedContents.ImageList.Images.Add(Properties.Resources.folder);
            tvEncryptedContents.ImageList.Images.Add(Properties.Resources.BinaryFile);

            try
            {
                //Load file
                if (enableCommit)
                {
                    cf = new CompoundFile(fs, CFSUpdateMode.Update, CFSConfiguration.SectorRecycle | CFSConfiguration.NoValidationException | CFSConfiguration.EraseFreeSectors);
                }
                else
                {
                    cf = new CompoundFile(fs);
                }

                // populate treeview
                tvEncryptedContents.Nodes.Clear();
                TreeNode root = null;
                root = tvEncryptedContents.Nodes.Add("Root Entry", "Root");
                root.Tag = cf.RootStorage;
                AddNodes(root, cf.RootStorage);
                tvEncryptedContents.ExpandAll();
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, ex.Message);
                MessageBox.Show(ex.Message, "File Load Fail", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
        }

        /// <summary>
        /// Recursive addition of tree nodes foreach child of current item in the storage
        /// </summary>
        /// <param name="node">Current TreeNode</param>
        /// <param name="cfs">Current storage associated with node</param>
        private void AddNodes(TreeNode node, CFStorage cfs)
        {
            Action<CFItem> va = delegate (CFItem target)
            {
                TreeNode temp = node.Nodes.Add( target.Name, target.Name + (target.IsStream ? " (" + target.Size + " bytes )" : string.Empty));
                temp.Tag = target;
                
                // set images for treeview
                if (target.IsStream)
                {
                    temp.ImageIndex = 1;
                    temp.SelectedImageIndex = 1;
                }
                else
                {
                    temp.ImageIndex = 0;
                    temp.SelectedImageIndex = 0;

                    // Recursion into the storage
                    AddNodes(temp, (CFStorage)target);
                }
            };

            //Visit NON-recursively (first level only)
            cfs.VisitEntries(va, false);
        }

        private void tvEncryptedContents_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { Close(); }
        }

        private void tvEncryptedContents_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                TreeNode n = tvEncryptedContents.GetNodeAt(e.X, e.Y);
                if (n != null)
                {
                    tvEncryptedContents.SelectedNode = n;

                    // The tag property contains the underlying CFItem.
                    // CFItem target = (CFItem)n.Tag;
                    cfStream = n.Tag as CFStream;
                    if (cfStream != null)
                    {
                        byte[] buffer = new byte[cfStream.Size];
                        cfStream.Read(buffer, 0, buffer.Length);

                        StringBuilder sb = new StringBuilder();
                        foreach (byte b in buffer)
                        {
                            if (b != 0)
                            {
                                sb.Append(AppUtilities.ConvertByteToText(b.ToString()));
                            }
                        }

                        TxbOutput.Text = sb.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "NodeMouseClick Fail", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// write out any changes made in the textbox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnSave_Click(object sender, EventArgs e)
        {
            // write the stream changes and save
            cfStream.Write(Encoding.Default.GetBytes(TxbOutput.Text), 0, 0, Encoding.Default.GetByteCount(TxbOutput.Text));
            cf.Commit();

            // let the user know it worked, then close the stream and form
            MessageBox.Show("Stream changes saved.", "File Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
            cf.Close();
            Close();
        }

        private void FrmEncryptedFile_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (cf != null)
            {
                cf.Close();
            }
        }

        /// <summary>
        /// validate the labelinfo stream
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnValidateXml_Click(object sender, EventArgs e)
        {
            ValidateXml(true);
        }

        /// <summary>
        /// populate the validation errors
        /// </summary>
        /// <param name="displayOnly">only display ui for validate xml button</param>
        public void ValidateXml(bool displayOnly)
        {
            isValid = true;
            validationErrors.Clear();

            // currently only validating labelinfo streams
            if (tvEncryptedContents.SelectedNode.Name != "LabelInfo")
            {
                return;
            }

            // start validating the xml
            try
            {
                ValidationEventHandler eventHandler = new ValidationEventHandler(ValidationEventHandler);
                var xsdPath = new Uri(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)).LocalPath + "\\Schemas\\LabelInfo.xsd";
                XmlSchemaSet schema = new XmlSchemaSet();
                schema.Add(string.Empty, xsdPath);

                var settings = new XmlReaderSettings();
                settings.Schemas.Add("http://schemas.microsoft.com/office/2020/mipLabelMetadata", xsdPath);
                settings.ValidationType = ValidationType.Schema;
                settings.ValidationFlags |= XmlSchemaValidationFlags.ProcessInlineSchema;
                settings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
                settings.ValidationEventHandler += new ValidationEventHandler(ValidationEventHandler);

                using (TextReader textReader = new StringReader(TxbOutput.Text))
                {
                    XmlReader rd = XmlReader.Create(textReader, settings);
                    XDocument doc = XDocument.Load(rd);
                    doc.Validate(schema, eventHandler);
                }
            }
            catch (Exception ex)
            {
                // if there were xml validation errors, display a message with those details
                FileUtilities.WriteToLog(Strings.fLogFilePath, ex.Message);
                
                if (displayOnly)
                {
                    MessageBox.Show(ex.Message, "Xml Validation Errors", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                
                isValid = false;
            }
            finally
            {
                if (isValid && displayOnly)
                {
                    MessageBox.Show("Xml Is Valid.", "Xml Validation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // display schema errors
                    if (validationErrors.Count > 0)
                    {
                        StringBuilder sb = new StringBuilder();
                        int errorCount = 0;
                        foreach (string s in validationErrors)
                        {
                            errorCount++;
                            sb.Append(errorCount + Strings.wPeriod + s + "\r\n");
                        }

                        if (displayOnly)
                        {
                            MessageBox.Show(sb.ToString(), "Schema Validation Errors", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// write out schema validation errors
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void ValidationEventHandler(object sender, ValidationEventArgs e)
        {
            isValid = false;
            switch (e.Severity)
            {
                case XmlSeverityType.Error:
                    FileUtilities.WriteToLog(Strings.fLogFilePath, e.Message);
                    validationErrors.Add("Error at Line #" + e.Exception.LineNumber + " Position #" + e.Exception.LinePosition + Strings.wColonBuffer + e.Message);
                    break;
                case XmlSeverityType.Warning:
                    FileUtilities.WriteToLog(Strings.fLogFilePath, e.Message);
                    validationErrors.Add("Error at Line #" + e.Exception.LineNumber + " Position #" + e.Exception.LinePosition + Strings.wColonBuffer + e.Message);
                    break;
            }
        }

        /// <summary>
        /// attempt known fixes for corrupt / invalid labelinfo xml
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnFixXml_Click(object sender, EventArgs e)
        {
            // populate validation errors
            ValidateXml(false);

            // check existing validation errors against known corrupt labelinfo xml scenarios
            if (validationErrors.Count > 0)
            {
                string valToReplace = string.Empty;

                foreach (string s in validationErrors)
                {
                    // known issue fix - missing method attribute
                    // sometimes the method is not written out use a simple find/replace to get it back in
                    if (s.Contains("The required attribute 'method' is missing"))
                    {
                        TxbOutput.Text = TxbOutput.Text.Replace("enabled=\"1\"", "enabled=\"1\" method=\"Standard\"");
                    }

                    // known issue fix - siteid missing brackets
                    // example siteId="11111111-1111-1111-1111-111111111111"
                    // should be siteId="{11111111-1111-1111-1111-111111111111}"
                    if (s.Contains("The 'siteId' attribute is invalid"))
                    {
                        // first we need to pull the full text of the siteId attribute
                        string[] split = Regex.Split(TxbOutput.Text, @" +");
                        foreach (string sp in split)
                        {
                            if (sp.StartsWith("siteId=") && sp.Contains('{') == false && sp.Contains('}') == false)
                            {
                                string[] replace = sp.Split('"');
                                valToReplace = replace[0] + "\"{" + replace[1] + "}\"";
                                TxbOutput.Text = TxbOutput.Text.Replace(sp, valToReplace);
                            }
                        }
                    }
                }
            }
        }
    }
}