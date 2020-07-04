using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CaptionBuilder
{
    public partial class frmSave : Form
    {
        public frmSave()
        {
            InitializeComponent();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            lblSaving.Visible = true;
            progressBar1.Visible = true;
           
            int i;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 200;

            for (i = 0; i <= 200; i += 50)
            {
                progressBar1.Value = i;
                System.Threading.Thread.Sleep(500);
            }

            Form1.saveStatus = "Save";
           
            this.Close();
        }

        private void frmSave_FormClosing(object sender, FormClosingEventArgs e)
        { 

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void btnDontSave_Click(object sender, EventArgs e)
        {
            Form1.saveStatus = "DontSave";
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Form1.saveStatus = "Cancel";
            this.Close();

        }
    }
}
