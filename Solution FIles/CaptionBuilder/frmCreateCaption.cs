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
    public partial class frmCreateCaption : Form
    {
        public frmCreateCaption()
        {
            InitializeComponent();
           
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void cboCategory_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtCaptionText.Enabled = true;

        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            if (txtCaptionText.Text != "")
            {
                Form1.captionText = txtCaptionText.Text;
                //check if save button is clicked
                Form1.clickSave = true;
                
                this.Close();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            
            this.Close();
            
        }

        private void frmCreateCaption_Load(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void txtCaptionText_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                btnCreate.PerformClick();
            }
        }
    }
}
