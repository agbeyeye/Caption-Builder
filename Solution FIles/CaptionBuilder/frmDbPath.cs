using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace CaptionBuilder
{
    public partial class frmDbPath : Form
    {
        public frmDbPath()
        {
            InitializeComponent();
        }
        string targetDatabaseLocation= "DatabaseLocation.txt";
       // string dblocation = "";
        private void frmDbPath_Load(object sender, EventArgs e)
        {
            //get location of targeted db
            using (var reader = new StreamReader(targetDatabaseLocation))
            {
                txtDbPath.Text = reader.ReadToEnd();
            }
            txtDbPath.SelectAll();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (txtDbPath.Text != "")
            {
                using (StreamWriter writer = new StreamWriter(targetDatabaseLocation))
                {
                    writer.Write(txtDbPath.Text);
                }
            }
            Close();
        }

        private void txtDbPath_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                btnSave.PerformClick();
            }
        }
    }
}
