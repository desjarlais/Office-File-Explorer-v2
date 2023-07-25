using System.Text;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmCallbackViewer : Form
    {
        public FrmCallbackViewer(StringBuilder rtf)
        {
            InitializeComponent();
            rtbCallbacks.Text = rtf.ToString();
        }
    }
}
