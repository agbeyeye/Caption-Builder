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
    public partial class frmAddCaption : Form
    {
        private List<Scope> scopes;
        private List<Caption> captions;
        public frmAddCaption( List<Caption> createdCaptions)
        {
            InitializeComponent();
           
            this.captions = createdCaptions;
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnAdd.Enabled = true;
          
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        
        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                if (listView1.FocusedItem.Index >= 0)
                {
                    Form1.captionText = listView1.FocusedItem.Text;
                    Form1.category = listView1.FocusedItem.SubItems[0].Text;
                    Form1.clickSave = true;
                    this.Close();
                }
            }catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmAddCaption_Load(object sender, EventArgs e)
        {
            txtSearch.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtSearch.AutoCompleteSource = AutoCompleteSource.CustomSource;
            AutoCompleteStringCollection autoCollection= new AutoCompleteStringCollection();
            if (captions != null)
            {


                foreach (var item in captions)
                {
                    autoCollection.Add(item.name);
                    ListViewItem item1 = new ListViewItem(item.name);
                    item1.SubItems.Add(item.category.name);
                    listView1.Items.Add(item1);

                }
            }
            txtSearch.AutoCompleteCustomSource = autoCollection;
        }

        private void txtSearch_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                //listView1.Items.Clear();
                //ListViewItem item1 = new ListViewItem(txtSearch.Text);
                //item1.SubItems.Add(getCategory(txtSearch.Text));
                //listView1.Items.Add(item1);
                MessageBox.Show("hello");
            }
          
        }

        private void loadListview()
        {
            listView1.Items.Clear();
            foreach (var item in captions)
            {
               
                ListViewItem item1 = new ListViewItem(item.name);
                item1.SubItems.Add(item.category.name);
                listView1.Items.Add(item1);

            }
        }

        private string getCategory(string captionString)
        {
            string category = "";
            foreach(var item in captions)
            {
                if (item.name == captionString)
                {
                    category = item.category.name;
                    break;
                }
            }

            return category;
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            if (txtSearch.Text == String.Empty)
            {
                loadListview();
            }
        }

        private void txtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData==Keys.Enter)
            {
                try
                {
                    string category = getCategory(txtSearch.Text);
                    listView1.Items.Clear();
                    if (category != String.Empty)
                    {
                        ListViewItem item1 = new ListViewItem(txtSearch.Text);
                        item1.SubItems.Add(category);
                        listView1.Items.Add(item1);
                        
                    }


                }
                catch(Exception)
                {
                    return;
                }
            }
        }
    }
}
