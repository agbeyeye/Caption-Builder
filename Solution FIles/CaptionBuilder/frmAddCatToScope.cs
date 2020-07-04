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
    public partial class frmAddCatToScope : Form
    {
        
        public frmAddCatToScope(ListBox.ObjectCollection lbx)
        {
            InitializeComponent();
            if (lbx.Count > 0)
            {
                foreach (var item in lbx)
                {
                    listBox1.Items.Add(item.ToString());
                }
            }
           
        }

        private void frmAddCatToScope_Load(object sender, EventArgs e)
        {
           
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex > -1)
            {
                btnOk.Enabled = true;
            }
            else
            {
                btnOk.Enabled = false;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex > -1)
            {
                Form1.category = listBox1.SelectedItem.ToString();
                Form1.clickSave = true;
                Close();
            }
        }
    }
}
