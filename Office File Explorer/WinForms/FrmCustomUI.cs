using Office_File_Explorer.Helpers;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmCustomUI : Form
    {
        private string[] rtfContents = new string[(int)XMLParts.LastEntry];
        private OfficeDocument package = null;
        private string fPath;

        public FrmCustomUI(string fileName)
        {
            InitializeComponent();
            fPath = fileName;
            package = new OfficeDocument(fPath);
            PackageLoaded();
        }

        private void PackageLoaded()
        {
            foreach (OfficePart part in package.Parts)
            {
                int contentIndex = (int)part.PartType;
                rtbCustomUI.Rtf = XmlColorizer.Colorize(part.ReadContent());
                rtbCustomUI.Tag = part.PartType;
                rtfContents[contentIndex] = rtbCustomUI.Rtf;
            }
        }

        private void FrmCustomUI_FormClosing(object sender, FormClosingEventArgs e)
        {
            package.ClosePackage();
        }
    }
}
