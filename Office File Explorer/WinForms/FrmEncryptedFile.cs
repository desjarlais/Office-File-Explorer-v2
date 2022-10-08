// the code for opening the structured compound file format was mostly taken from here https://github.com/ironfede/openmcdf

using Office_File_Explorer.Helpers;
using Office_File_Explorer.OpenMcdf;
using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmEncryptedFile : Form
    {
        private CompoundFile cf;
        private CFStream cfStream;

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
            MessageBox.Show("Stream saved.", "File Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
    }
}
