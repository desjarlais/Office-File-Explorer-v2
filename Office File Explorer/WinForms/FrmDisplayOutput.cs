using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Office_File_Explorer.Helpers;
using System;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmDisplayOutput : Form
    {
        public FrmDisplayOutput(StringBuilder rtfContent, bool isCallbackCode)
        {
            InitializeComponent();
            splitContainer1.Panel2Collapsed = true;

            if (isCallbackCode)
            {
                rtbRTFContent.Rtf = rtfContent.ToString();
            }
            else
            {
                rtbRTFContent.Text = rtfContent.ToString();
            }

            // format the headings
            //for (int i = 0; i < rtbRTFContent.Lines.Count(); i++)
            //{
            //    if (rtbRTFContent.Lines[i].StartsWith("---") && rtbRTFContent.Lines[i].EndsWith("---"))
            //    {
            //        rtbRTFContent.Select(rtbRTFContent.GetFirstCharIndexFromLine(i), rtbRTFContent.Lines[i].Length);
            //        rtbRTFContent.SelectionFont = new Font("Arial", 12, FontStyle.Bold);
            //    }
            //}
        }

        public FrmDisplayOutput(Image img)
        {
            InitializeComponent();
            splitContainer1.Panel1Collapsed = true;
            pictureBox1.Image = img;

            if (img.Height > pictureBox1.Size.Height || img.Width > pictureBox1.Size.Width) 
            {
                pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            }
            else
            {
                pictureBox1.SizeMode = PictureBoxSizeMode.CenterImage;
            }

            this.Text = "Image";
        }

        private void copySelectedTextToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                if (rtbRTFContent.Text.Length == 0)
                {
                    Clipboard.SetText(rtbRTFContent.SelectedText);
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "Copy Error: " + ex.Message);
            }
        }

        private void copyAllTextToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                if (rtbRTFContent.Text.Length == 0) { return; }
                StringBuilder buffer = new StringBuilder();
                foreach (string s in rtbRTFContent.Lines)
                {
                    buffer.Append(s);
                    buffer.Append('\n');
                }

                Clipboard.SetText(buffer.ToString());
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "Copy Error: " + ex.Message);
            }
        }
    }
}
