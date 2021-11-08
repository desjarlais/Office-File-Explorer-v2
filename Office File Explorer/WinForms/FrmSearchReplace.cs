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
    public partial class FrmSearchReplace : Form
    {
        public FrmSearchReplace()
        {
            InitializeComponent();
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            if (Owner is FrmMain f)
            {
                f.FindTextProperty = tbFind.Text;
                f.ReplaceTextProperty = tbReplace.Text;
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
