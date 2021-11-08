using System;
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
            //return Convert.TryFromBase64String(base64.PadRight(base64.Length / 4 * 4 + (base64.Length % 4 == 0 ? 0 : 4), '='), new Span<byte>(new byte[base64.Length]), out _);
            return Convert.TryFromBase64String(base64, buffer, out int bytesParsed);
        }
    }
}
