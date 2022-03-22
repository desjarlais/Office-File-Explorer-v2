using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using Office_File_Explorer.Helpers;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmDeleteCustomProps : Form
    {
        CustomFilePropertiesPart part;

        public bool PartModified { get; set; }

        public FrmDeleteCustomProps(CustomFilePropertiesPart cfp)
        {
            InitializeComponent();
            PartModified = false;
            part = cfp;
            UpdateList();
        }

        public void UpdateList()
        {
            if (part is null)
            {
                lbProps.Items.Add(Strings.wCustomDocProps);
                return;
            }

            int count = 0;

            foreach (var v in CfpList(part))
            {
                count++;
                lbProps.Items.Add(count + Strings.wPeriod + v);
            }

            lbProps.SelectedIndex = 0;
        }

        public List<string> CfpList(CustomFilePropertiesPart part)
        {
            List<string> val = new List<string>();
            foreach (CustomDocumentProperty cdp in part.RootElement)
            {
                val.Add(cdp.Name);
            }
            return val;
        }

        private void BtnOk_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void BtnDeleteProp_Click(object sender, System.EventArgs e)
        {
            string[] valToDelete = lbProps.SelectedItem.ToString().Split('.');
            string val = valToDelete[1].TrimStart();
            foreach (CustomDocumentProperty cdp in part.RootElement)
            {
                if (val == cdp.Name)
                {
                    cdp.Remove();
                    PartModified = true;
                    lbProps.Items.RemoveAt(lbProps.SelectedIndex);
                    lbProps.Items.Clear();
                    UpdateList();
                    lbProps.SelectedIndex = 0;
                }
            }
        }

        private void FrmDeleteCustomProps_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }
    }
}
