using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmRichTextBox : Form
    {
        public FrmRichTextBox(string rtfContent)
        {
            InitializeComponent();
            rtbOutput.Rtf = rtfContent;
        }
    }
}
