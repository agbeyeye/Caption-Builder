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
    public partial class frmNewScope : Form
    {
        
       // Data collection=new Data();

        public frmNewScope()
        {
            InitializeComponent();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (txtScopeName.Text != "")
            {

                // collection.NewScope(txtScopeName.Text);
                Form1.newScopeName = txtScopeName.Text;
                Form1.clickSave = true;
                this.Close();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void txtScopeName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                btnSave.PerformClick();
            }
        }

        private void frmNewScope_Load(object sender, EventArgs e)
        {
            txtScopeName.Focus();
        }
    }
}
