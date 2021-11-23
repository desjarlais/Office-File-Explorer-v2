using System;
using System.Threading;
using System.Windows.Forms;

namespace Office_File_Explorer
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // create a named mutex
            Mutex mutex = new Mutex(false, "brandesoft office file explorer v2", out bool noInstanceCurrently);

            // let the user know if we already exist.
            if (noInstanceCurrently == false)
            {
                MessageBox.Show("The application is already running.", "Application Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FrmMain());
        }
    }
}
