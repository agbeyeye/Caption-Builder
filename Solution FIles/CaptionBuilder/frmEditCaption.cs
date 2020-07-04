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
    public partial class frmEditCaption : Form
    {
        private string captionText="";
        int index=0;
        public frmEditCaption(string captionText,int index)
        {
            InitializeComponent();
            this.index = index;
            if (index > 0)
            {
                this.Text = "Add Caption Extension";
                
            }
            this.captionText = captionText;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (txtCaption.Text != "")
            {
                Form1.EditedCaptionText = txtCaption.Text;
                Form1.clickSave = true;
                //check if save button is clicked for auto edit open
                Form1.endNextEdit = false;
                this.Close();
            }else if(txtCaption.Text=="" && index > 0)
            {
                Form1.EditedCaptionText = "";
                Form1.clickSave = true;
                //check if save button is clicked for auto edit open
                Form1.endNextEdit = false;
                this.Close();
            }
           
        }

        private void frmEditCaption_Load(object sender, EventArgs e)
        {
            Form1.endNextEdit = true;
            txtCaption.Text = captionText;
            txtCaption.Focus();
            //txtCaption.SelectAll();
        }

        private void txtCaption_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                btnSave.PerformClick();
            }
        }
    }
}
