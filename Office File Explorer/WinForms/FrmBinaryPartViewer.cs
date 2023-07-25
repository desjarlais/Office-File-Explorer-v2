using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmBinaryPartViewer : Form
    {
        public FrmBinaryPartViewer(StringBuilder rtfContent)
        {
            InitializeComponent();
            splitContainer1.Panel2Collapsed = true;
            rtbRTFContent.Rtf = rtfContent.ToString();
        }

        public FrmBinaryPartViewer(Image img)
        {
            InitializeComponent();
            splitContainer1.Panel1Collapsed = true;
            pictureBox1.Image = img;
        }
    }
}
