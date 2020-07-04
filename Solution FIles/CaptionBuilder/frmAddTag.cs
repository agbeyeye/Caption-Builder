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
    public partial class frmAddTag : Form
    {
        public frmAddTag()
        {
            InitializeComponent();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (txtTagText.Text != "")
            {
                Form1.tagText = txtTagText.Text;
                Form1.clickSave = true;
                Close();
            }
        }

        private void txtTagText_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                btnSave.PerformClick();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmAddTag_Load(object sender, EventArgs e)
        {

        }
    }
}
