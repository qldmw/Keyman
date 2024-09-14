namespace Keyman
{
    partial class Form2
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;


        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            // textBox1
            this.textBox1.Location = new System.Drawing.Point(15, 15);
            this.textBox1.MaxLength = 10;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(150, 20);
            // okButton
            this.okButton.Location = new System.Drawing.Point(15, 50);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(70, 23);
            this.okButton.Text = "OK";
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // cancelButton
            this.cancelButton.Location = new System.Drawing.Point(95, 50);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(70, 23);
            this.cancelButton.Text = "Cancel";
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // InputDialog
            this.ClientSize = new System.Drawing.Size(180, 85);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.cancelButton);
        }

        #endregion
    }
}