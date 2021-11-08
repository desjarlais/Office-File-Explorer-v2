using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmBatchDeleteCustomProps : Form
    {
        public string PropName { get; set; }

        public FrmBatchDeleteCustomProps()
        {
            InitializeComponent();
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            PropName = tbPropName.Text;
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            PropName = "Cancel";
            Close();
        }
    }
}
