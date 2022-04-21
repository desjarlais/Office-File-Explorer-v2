using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmBase64 : Form
    {
        public FrmBase64()
        {
            InitializeComponent();
        }

        private void TxbEncoded_TextChanged(object sender, EventArgs e)
        {
            if (IsBase64String(txbEncoded.Text))
            {
                txbResult.Text = Base64Decode(txbEncoded.Text);
            }
            else
            {
                txbResult.Text = "Base64 Encoded Value Is Invalid!";
            }
        }

        public static string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = Convert.FromBase64String(base64EncodedData);
            return Encoding.UTF8.GetString(base64EncodedBytes);
        }

        public static bool IsBase64String(string base64)
        {
            Span<byte> buffer = new Span<byte>(new byte[base64.Length]);
            return Convert.TryFromBase64String(base64, buffer, out int bytesParsed);
        }

        private void FrmBase64_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }

        private void BtnParseLicense_Click(object sender, EventArgs e)
        {
            OpenFileDialog fDialog = new OpenFileDialog
            {
                Title = "Select Office 365 License File.",
                Filter = "License Files | *.*",
                RestoreDirectory = true,
                InitialDirectory = @"%userprofile%"
            };

            if (fDialog.ShowDialog() == DialogResult.OK && Path.GetExtension(fDialog.FileName) == string.Empty)
            {
                string text = File.ReadAllText(fDialog.FileName);
                txbResult.Text = Base64Decode(text);
            }
            else
            {
                txbResult.Text = "Not A Valid License File!";
            }
        }
    }
}