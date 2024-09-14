using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace Keyman
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            this.MinimizeBox = false;
            this.MaximizeBox = false;
            this.TopMost = true;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.AcceptButton = this.okButton;
            this.CancelButton = this.cancelButton;
            this.Activated += new EventHandler(Form2_Activated);
        }

        private void Form2_Activated(object sender, EventArgs e)
        {
            this.textBox1.Focus();
        }

        public string InputValue
        {
            get { return textBox1.Text; }
            set { textBox1.Text = value; }
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}