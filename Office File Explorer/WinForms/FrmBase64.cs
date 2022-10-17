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
            if (e.KeyCode == Keys.Escape) { Close(); }
        }

        /// <summary>
        /// Office 365 uses a file with no extension in the user profile for license information
        /// parse the first "License" portion of the encoded value and decode it
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnParseLicense_Click(object sender, EventArgs e)
        {
            OpenFileDialog fDialog = new OpenFileDialog
            {
                Title = "Select License File.",
                Filter = "License Files | *.*",
                RestoreDirectory = true,
                InitialDirectory = @"%userprofile%"
            };

            if (fDialog.ShowDialog() == DialogResult.OK && Path.GetExtension(fDialog.FileName) == string.Empty)
            {
                // get the decoded value from the file
                byte[] bytes = File.ReadAllBytes(fDialog.FileName);

                // remove initial 12 chars, then the last " char before adding the first split into the encoded textbox
                // the TextChanged event should handle decoding the value
                string encodedText = Encoding.Unicode.GetString(bytes).Remove(0, 12);
                encodedText = encodedText.Replace('"', ' ');
                string[] result = encodedText.Split(new char[] { ',' });
                txbEncoded.Text = result[0];
            }
            else
            {
                txbResult.Text = "Not A Valid License File!";
            }
        }
    }
}