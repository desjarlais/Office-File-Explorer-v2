﻿using Office_File_Explorer.Helpers;
using System;
using System.Reflection;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmAbout : Form
    {
        public FrmAbout()
        {
            InitializeComponent();

            // display app info
            lblProduct.Text = AssemblyProduct;
            lblVersion.Text = AssemblyVersion; 
        }

        #region Assembly Attribute Accessors

        public static string AssemblyVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }

        public string AssemblyDescription
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0)
                {
                    return string.Empty;
                }
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        public string AssemblyProduct
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0)
                {
                    return string.Empty;
                }
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        public string AssemblyCopyright
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                {
                    return string.Empty;
                }
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        public string AssemblyCompany
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                {
                    return string.Empty;
                }
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
        #endregion

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            AppUtilities.PlatformSpecificProcessStart("https://github.com/desjarlais/Office-File-Explorer-v2");
        }

        private void FrmAbout_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { Close(); }
        }
    }
}
