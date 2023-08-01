using System.Drawing;
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
        }

        public FrmDisplayOutput(Image img)
        {
            InitializeComponent();
            splitContainer1.Panel1Collapsed = true;
            pictureBox1.Image = img;
        }
    }
}
