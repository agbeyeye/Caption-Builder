using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace CaptionBuilder
{
    public partial class Form1 : Form
    {
        Controller collection;
        frmNewScope newScope;
        frmNewCategory newCategory;
        frmAddCatToScope catToScope;
        frmAddTag tagForm;
        frmAddCompleteItem completeItem;
        frmDbPath setDbPath;
        Model db;
        //keep trace of caption arrangement by user
        List<Caption> captionsArrangement;
        //keep trace of category arrangement
        List<Category> categoryArrangement;
        //name of new scope to be created
        public static string newScopeName { get; set; }
        //keep  track of whether save button has been clicked or not on other forms
        public static bool clickSave { get; set; }
        //name given to a newly added category
        public static string newCategoryName { get; set; }
        //stores the name of category to be added to scope from addCategoryToScope form
        public static string category { get; set; }
        //stores the caption text from edit caption text form
        public static string captionText { get; set; }
        //stores the tag box number of a newly created tag to be added to the db
        public static int tagNumer { get; set; }
        //stores the name of the new tag from add tag form
        public static string tagText { get; set; }
        //stores the name of a new complete item
        public static string completeItemText { get; set; }


        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);
        enum GetWindow_Cmd : uint
        {
            GW_HWNDFIRST = 0,
            GW_HWNDLAST = 1,
            GW_HWNDNEXT = 2,
            GW_HWNDPREV = 3,
            GW_OWNER = 4,
            GW_CHILD = 5,
            GW_ENABLEDPOPUP = 6
        }
        [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        public static extern IntPtr GetParent(IntPtr hWnd);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        public Form1()
        {

            InitializeComponent();
            //ckbNext.CheckState = CheckState.Checked;


        }

        private void Form1_Load(object sender, EventArgs e)
        {
           
            collection = new Controller();
            db = collection.GetData();
            try {
                if (db.Scopes != null)
                {
                    lbxScopes.Items.Clear();
                    foreach (var scope in db.Scopes.ToList())
                    {
                        lbxScopes.Items.Add(scope.name);
                    }
                }

                if (db.Categories != null)
                {
                    
                    lbxComplete.Items.Clear();
                    lbxCategories.Items.Clear();
                    lbxTag1.Items.Clear();
                    lbxTag2.Items.Clear();
                    lbxTag3.Items.Clear();
                    lbxTag4.Items.Clear();
                    //load categories into list box
                    loadCategoryListbox(db);
                    //load data added caption data into the listview
                    //load edit textbox for scope 
                    loadEditTextBox();
                    //load edit textbox for category
                    loadCategoryEditTextBox();
                    //load edit textbox for tag1
                    loadTag1EditTextBox();
                    //load edit text box for tag2
                    loadTag2EditTextBox();
                    //load edit textbox for tag3
                    loadTag3EditTextBox();
                    //load edit textbox for tag4
                    loadTag4EditTextBox();
                    //load edit textbox for complete listbox
                    loadCompleteEditTextBox();
                    //load tags from db into the tags listbox
                    loadTags();
                    //load complete item from db into complete listbox
                    loadCompleteItems();
                    //load captions from db into the grid
                    if (db.Captions != null)
                    { loadGrid(db, "All", String.Empty, String.Empty); }
                }
            }
            catch (Exception) {
                MessageBox.Show("Please make sure database location provided is correct", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                setDbPath = new frmDbPath();
                setDbPath.ShowDialog();

            }
            
         
            //set drag and drop effects on all listbox
            lbxCategories.AllowDrop = true;
            lbxScopes.AllowDrop = true;
            lbxComplete.AllowDrop = true;
            lbxTag1.AllowDrop = true;
            lbxTag2.AllowDrop = true;
            lbxTag3.AllowDrop = true;
            lbxTag4.AllowDrop = true;
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //de-select items in scope and tags listbox anytime category item is selected
            lbxScopes.SelectedIndex = -1;
            lbxTag1.SelectedIndex = -1;
            lbxTag2.SelectedIndex = -1;
            lbxTag3.SelectedIndex = -1;
            lbxTag4.SelectedIndex = -1;
            lbxComplete.SelectedIndex = -1;

            if (lbxCategories.SelectedIndex > -1) { btnCreateCaption.Enabled = true; } else { btnCreateCaption.Enabled = false; }
            if (lbxCategories.SelectedIndex != -1 && lbxScopes.SelectedIndex == -1)
            {
                loadGrid(db, "Category", lbxCategories.SelectedItem.ToString(), String.Empty);
            }
            else if (lbxCategories.SelectedIndex > -1 && lbxScopes.SelectedIndex > -1)
            {
                loadGrid(db, "Mix", lbxCategories.SelectedItem.ToString(), lbxScopes.SelectedItem.ToString());
            }
            else if (lbxCategories.SelectedIndex == -1 && lbxScopes.SelectedIndex == -1)
            {
                loadGrid(db, "All", String.Empty, String.Empty);
            }
            else if (lbxCategories.SelectedIndex == -1 && lbxScopes.SelectedIndex != -1)
            {
                loadGrid(db, "Scope", String.Empty, lbxScopes.SelectedItem.ToString());

            }

        }

        //Add new scope to  model
        private void btnAddScope_Click(object sender, EventArgs e)
        {
            newScope = new frmNewScope();
            newScope.ShowDialog();
            if (clickSave)
            {
                if (!lbxScopes.Items.Contains(newScopeName))
                {
                    db = collection.NewScope(newScopeName);
                    //loadScopes(db);
                    lbxScopes.Items.Add(newScopeName);
                }
                else
                {
                    MessageBox.Show("Scope is already existing");
                }

                clickSave = false;
            }


        }

        //Load updates of scopes from model
        private void loadScopes(Model obj)
        {


            lbxScopes.Items.Clear();
            foreach (var scope in obj.Scopes.ToList())
            {
                lbxScopes.Items.Add(scope.name);
            }
        }

        //Load updated categories from model
        private void loadCategoryListbox(Model obj)
        {
            lbxCategories.Items.Clear();
            foreach (var category in db.Categories.ToList())
            {
                lbxCategories.Items.Add(category.name);
            }
        }


        //Add new category to a scope
        private void btnAddCatgory_Click(object sender, EventArgs e)
        {

            newCategory = new frmNewCategory();
            newCategory.ShowDialog();
            if (clickSave)
            {
                if (!lbxCategories.Items.Contains(newCategoryName))
                {
                    db = collection.NewCategory(newCategoryName);
                    loadCategoryListbox(db);
                }
                else
                {
                    MessageBox.Show("Item is already present in the list");
                }

                clickSave = false;
            }
        }

        private void lbxScopes_SelectedIndexChanged(object sender, EventArgs e)
        {
            //de-select items in category and tags listbox anytime scope item is selected
            lbxCategories.SelectedIndex = -1;
            lbxTag1.SelectedIndex = -1;
            lbxTag2.SelectedIndex = -1;
            lbxTag3.SelectedIndex = -1;
            lbxTag4.SelectedIndex = -1;
            lbxComplete.SelectedIndex = -1;

            if (lbxScopes.SelectedIndex > -1) { btnAddCaption.Enabled = true; btnCatToScope.Enabled = true; } else { btnAddCaption.Enabled = false; btnCatToScope.Enabled = false; }
            if (lbxCategories.SelectedIndex == -1 && lbxScopes.SelectedIndex != -1)
            {
                loadGrid(db, "Scope", String.Empty, lbxScopes.SelectedItem.ToString());
            }
            else if (lbxCategories.SelectedIndex > -1 && lbxScopes.SelectedIndex > -1)
            {
                loadGrid(db, "Mix", lbxCategories.SelectedItem.ToString(), lbxScopes.SelectedItem.ToString());
            }
            else if (lbxCategories.SelectedIndex == -1 && lbxScopes.SelectedIndex == -1)
            {
                loadGrid(db, "All", String.Empty, String.Empty);
            }
            else if (lbxCategories.SelectedIndex != -1 && lbxScopes.SelectedIndex == -1)
            {
                loadGrid(db, "Category", lbxCategories.SelectedItem.ToString(), String.Empty);
            }



        }

        frmCreateCaption newCaption;
        frmAddCaption addCaption;
        private void btnCreateCaption_Click(object sender, EventArgs e)
        {
            newCaption = new frmCreateCaption();
            newCaption.ShowDialog();
            if (clickSave)
            {
                db = collection.CreateCaption(captionText, new Category { name = lbxCategories.SelectedItem.ToString() });

                if (lbxCategories.SelectedIndex == -1 && lbxScopes.SelectedIndex != -1)
                {
                    loadGrid(db, "Scope", String.Empty, lbxScopes.SelectedItem.ToString());
                }
                else if (lbxCategories.SelectedIndex > -1 && lbxScopes.SelectedIndex > -1)
                {
                    loadGrid(db, "Mix", lbxCategories.SelectedItem.ToString(), lbxScopes.SelectedItem.ToString());
                }
                else if (lbxCategories.SelectedIndex == -1 && lbxScopes.SelectedIndex == -1)
                {
                    loadGrid(db, "All", String.Empty, String.Empty);
                }
                else if (lbxCategories.SelectedIndex != -1 && lbxScopes.SelectedIndex == -1)
                {
                    loadGrid(db, "Category", lbxCategories.SelectedItem.ToString(), String.Empty);
                }
                clickSave = false;

                dataGridView1.Focus();
                dataGridView1.CurrentCell = dataGridView1[0, dataGridView1.RowCount - 1];

            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            
            saveScope();
            saveCategory();
            saveCaptionsAndCategoriesOrder();
            saveComplete();
            saveTags();
            collection.SaveData(db);
            MessageBox.Show("Saved successfully");
        }

        private void btnAddCaption_Click(object sender, EventArgs e)
        {

            addCaption = new frmAddCaption(db.Captions);
            addCaption.ShowDialog();
            if (clickSave)
            {
                db = collection.AddCaptionToScope(captionText, category, lbxScopes.SelectedItem.ToString());
                if (lbxCategories.SelectedIndex == -1 && lbxScopes.SelectedIndex != -1)
                {
                    loadGrid(db, "Scope", String.Empty, lbxScopes.SelectedItem.ToString());
                }
                else if (lbxCategories.SelectedIndex > -1 && lbxScopes.SelectedIndex > -1)
                {
                    loadGrid(db, "Mix", lbxCategories.SelectedItem.ToString(), lbxScopes.SelectedItem.ToString());
                }
                else if (lbxCategories.SelectedIndex == -1 && lbxScopes.SelectedIndex == -1)
                {
                    loadGrid(db, "All", String.Empty, String.Empty);
                }
                else if (lbxCategories.SelectedIndex != -1 && lbxScopes.SelectedIndex == -1)
                {
                    loadGrid(db, "Category", lbxCategories.SelectedItem.ToString(), String.Empty);
                }
                clickSave = false;
            }
        }

        frmSave saveform;
        public static string saveStatus = "";

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            saveform = new frmSave();
            saveform.ShowDialog();
            if (saveStatus == "Save")
            {
                saveScope();
                saveCategory();
                saveCaptionsAndCategoriesOrder();
                saveComplete();
                saveTags();
                collection.SaveData(db);

                //this.Close();
            }
            else if (saveStatus == "Cancel")
            {
                e.Cancel = true;
            }

        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lbxCategories.SelectedIndex = -1;

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lbxScopes.SelectedIndex = -1;

        }

        List<string> oldScopes;
        int scopeIndex;
        private void editToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (lbxScopes.SelectedIndex > -1)
                {
                    oldScopes = new List<string>();
                    scopeIndex = lbxScopes.SelectedIndex;
                    foreach (var item in lbxScopes.Items) { oldScopes.Add(item.ToString()); }
                    CreateEditBox1(lbxScopes.GetItemRectangle(lbxScopes.SelectedIndex), lbxScopes.SelectedItem.ToString(), lbxScopes);
                }
                else { MessageBox.Show("Make sure to select an item"); }
            }
            catch (Exception ex)
            { MessageBox.Show(ex.Message); }
        }

        //load tags
        private void loadTags()
        {
            if (db.Tags != null)
            {
                lbxTag1.Items.Clear();
                lbxTag2.Items.Clear();
                lbxTag3.Items.Clear();
                lbxTag4.Items.Clear();
                foreach (var tag in db.Tags.ToList())
                {
                    if (tag.tagNumber == 1)
                    {
                        lbxTag1.Items.Add(tag.description);
                    }
                    else if (tag.tagNumber == 2)
                    {
                        lbxTag2.Items.Add(tag.description);
                    }
                    else if (tag.tagNumber == 3)
                    {
                        lbxTag3.Items.Add(tag.description);
                    }
                    else if (tag.tagNumber == 4)
                    {
                        lbxTag4.Items.Add(tag.description);
                    }
                }
            }
        }

        //load datagrid
        List<int> gridTilesIndex = new List<int>();
        private void loadGrid(Model obj, string filter, string categoryFilter, string scopeFilter)
        {
            int categoryCaptionCount = 0;
            gridTilesIndex = new List<int>();
            int captionIndex = 0;
            dataGridView1.ForeColor = Color.Blue;
            dataGridView1.Rows.Clear();
            if (obj.Captions != null)
            {


                if (filter == "All")
                {
                    dataGridView1.Rows.Add(obj.Captions.Count + obj.Categories.Count);
                    //for each category in the db
                    for (int i = 0; i < obj.Categories.Count; i++)
                    {
                        //get number of captions under a category
                        for (int j = 0; j < obj.Captions.Count; j++)
                        {
                            if (obj.Categories[i].name == obj.Captions[j].category.name)
                            {
                                categoryCaptionCount++;
                                break;
                            }
                        }
                        //if captions exist under a category then display them in datagrid
                        if (categoryCaptionCount > 0)
                        {
                            categoryCaptionCount = 0;
                            //ccreate tile for category in the datagrid
                            mergeCell(0, 7, captionIndex, obj.Categories[i].name);
                            //save records of tiles indices
                            gridTilesIndex.Add(captionIndex);
                            //increase row index 
                            captionIndex += 1;
                            //create rows in the datagridview
                            for (int j = 0; j < obj.Captions.Count; j++)
                            {
                                //for each category
                                if (obj.Categories[i].name == obj.Captions[j].category.name)
                                {
                                    //add captions into rows of datagrid
                                    dataGridView1.Rows[captionIndex].Cells[0].Value = obj.Captions[j].name;
                                    //add extensions of captions
                                    if (db.Captions[j].extensions.Length > 0)
                                    {
                                        //foreach extension of a caption
                                        for (int w = 0; w < db.Captions[j].extensions.Length; w++)
                                        {
                                            //add extensions to cell 2,3,4, and 5
                                            dataGridView1.Rows[captionIndex].Cells[w + 1].Value = obj.Captions[j].extensions[w];
                                        }
                                    }




                                    //set toggle button for next photo
                                    //dataGridView1.Rows[captionIndex].Cells[5].Value = "Y";
                                    if (obj.Captions[j].goToNextImage)
                                    {
                                        dataGridView1.Rows[captionIndex].Cells[5].Value = "Y";
                                        dataGridView1.Rows[captionIndex].Cells[5].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                                        dataGridView1.Rows[captionIndex].Cells[5].Style.ForeColor = Color.Black;
                                        dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Style.BackColor = Color.White;
                                        dataGridView1.Rows[captionIndex].Cells[5].Style.SelectionBackColor = SystemColors.Window;
                                        dataGridView1.Rows[captionIndex].Cells[5].Style.SelectionForeColor = SystemColors.ControlText;

                                    }
                                    else
                                    {
                                        dataGridView1.Rows[captionIndex].Cells[5].Value = "";
                                        dataGridView1.Rows[captionIndex].Cells[5].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                                        dataGridView1.Rows[captionIndex].Cells[5].Style.ForeColor = Color.Red;
                                        dataGridView1.Rows[captionIndex].Cells[5].Style.BackColor = Color.Red;
                                        dataGridView1.Rows[captionIndex].Cells[5].Style.SelectionBackColor = Color.Red;
                                        dataGridView1.Rows[captionIndex].Cells[5].Style.SelectionForeColor = Color.Red;
                                    }

                                    //add category name to row
                                    dataGridView1.Rows[captionIndex].Cells[6].Value = obj.Captions[j].category.name;
                                    dataGridView1.Rows[captionIndex].Cells[7].Value = obj.Captions[j].scopes;
                                    //move to the next row

                                    captionIndex += 1;

                                }
                            }
                        }
                    }
                }
                else if (filter == "Category")
                {

                    //get number of captions under a category
                    for (int j = 0; j < obj.Captions.Count; j++)
                    {
                        if (obj.Captions[j].category.name == categoryFilter)
                        {
                            categoryCaptionCount++;

                        }
                    }

                    //if captions exist under a category then display them in datagrid
                    if (categoryCaptionCount > 0)
                    {
                        dataGridView1.Rows.Add(categoryCaptionCount + 1);
                        //ccreate tile for category in the datagrid
                        mergeCell(0, 6, captionIndex, categoryFilter);
                        //save records of tiles indices
                        gridTilesIndex.Add(captionIndex);
                        //increase row index 
                        captionIndex += 1;
                        //create rows in the datagridview
                        for (int j = 0; j < obj.Captions.Count; j++)
                        {
                            //for each category
                            if (obj.Captions[j].category.name == categoryFilter)
                            {
                                //add captions into rows of datagrid
                                dataGridView1.Rows[captionIndex].Cells[0].Value = obj.Captions[j].name;
                                //add extensions of captions
                                if (db.Captions[j].extensions.Length > 0)
                                {
                                    //foreach extension of a caption
                                    for (int w = 0; w < db.Captions[j].extensions.Length; w++)
                                    {
                                        //add extensions to cell 2,3,4, and 5
                                        dataGridView1.Rows[captionIndex].Cells[w + 1].Value = obj.Captions[j].extensions[w];
                                    }
                                }
                                //set toggle field
                                // dataGridView1.Rows[captionIndex].Cells[5].Value = "Y";
                                if (obj.Captions[j].goToNextImage)
                                {
                                    dataGridView1.Rows[captionIndex].Cells[5].Value = "Y";
                                    dataGridView1.Rows[captionIndex].Cells[5].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                                    dataGridView1.Rows[captionIndex].Cells[5].Style.ForeColor = Color.Black;
                                    dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Style.BackColor = Color.White;
                                    dataGridView1.Rows[captionIndex].Cells[5].Style.SelectionBackColor = SystemColors.Window;
                                    dataGridView1.Rows[captionIndex].Cells[5].Style.SelectionForeColor = SystemColors.ControlText;

                                }
                                else
                                {
                                    dataGridView1.Rows[captionIndex].Cells[5].Value = "";
                                    dataGridView1.Rows[captionIndex].Cells[5].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                                    dataGridView1.Rows[captionIndex].Cells[5].Style.ForeColor = Color.Red;
                                    dataGridView1.Rows[captionIndex].Cells[5].Style.BackColor = Color.Red;
                                    dataGridView1.Rows[captionIndex].Cells[5].Style.SelectionBackColor = Color.Red;
                                    dataGridView1.Rows[captionIndex].Cells[5].Style.SelectionForeColor = Color.Red;
                                }


                                //add category name to row
                                dataGridView1.Rows[captionIndex].Cells[6].Value = obj.Captions[j].category.name;
                                //move to the next row
                                captionIndex += 1;

                            }
                        }
                        categoryCaptionCount = 0;
                    }

                }
                else if (filter == "Scope")
                {
                    //for each category in the db
                    for (int i = 0; i < obj.Categories.Count; i++)
                    {
                        categoryCaptionCount = 0;
                        //get number of captions under a category
                        for (int j = 0; j < obj.Captions.Count; j++)
                        {
                            //if row is a category title
                            if (obj.Captions[j].scopes != null)
                            {


                                if (obj.Categories[i].name == obj.Captions[j].category.name && obj.Captions[j].scopes.Contains(scopeFilter))
                                {
                                    //increase counter of captions under category
                                    categoryCaptionCount++;
                                }
                            }
                        }
                        //if captions exist under a category then display them in datagrid
                        if (categoryCaptionCount > 0)
                        {
                            //set the height of grid
                            dataGridView1.Rows.Add(categoryCaptionCount + 1);
                            // categoryCaptionCount = 0;
                            //ccreate tile for category in the datagrid
                            mergeCell(0, 6, captionIndex, obj.Categories[i].name);
                            //save records of tiles indices
                            gridTilesIndex.Add(captionIndex);
                            //increase row index 
                            captionIndex += 1;
                            //create rows in the datagridview
                            for (int j = 0; j < obj.Captions.Count; j++)
                            {
                                //for each category
                                if (obj.Categories[i].name == obj.Captions[j].category.name && obj.Captions[j].scopes.Contains(scopeFilter))
                                {
                                    //add captions into rows of datagrid
                                    dataGridView1.Rows[captionIndex].Cells[0].Value = obj.Captions[j].name;
                                    //add extensions of captions
                                    if (db.Captions[j].extensions.Length > 0)
                                    {
                                        //foreach extension of a caption
                                        for (int w = 0; w < db.Captions[j].extensions.Length; w++)
                                        {
                                            //add extensions to cell 2,3,4, and 5
                                            dataGridView1.Rows[captionIndex].Cells[w + 1].Value = obj.Captions[j].extensions[w];
                                        }
                                    }
                                    //set toggle field
                                    //dataGridView1.Rows[captionIndex].Cells[5].Value = "Y";
                                    if (obj.Captions[j].goToNextImage)
                                    {
                                        dataGridView1.Rows[captionIndex].Cells[5].Value = "Y";
                                        dataGridView1.Rows[captionIndex].Cells[5].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                                        dataGridView1.Rows[captionIndex].Cells[5].Style.ForeColor = Color.Black;
                                        dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Style.BackColor = Color.White;
                                        dataGridView1.Rows[captionIndex].Cells[5].Style.SelectionBackColor = SystemColors.Window;
                                        dataGridView1.Rows[captionIndex].Cells[5].Style.SelectionForeColor = SystemColors.ControlText;

                                    }
                                    else
                                    {
                                        dataGridView1.Rows[captionIndex].Cells[5].Value = "";
                                        dataGridView1.Rows[captionIndex].Cells[5].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                                        dataGridView1.Rows[captionIndex].Cells[5].Style.ForeColor = Color.Red;
                                        dataGridView1.Rows[captionIndex].Cells[5].Style.BackColor = Color.Red;
                                        dataGridView1.Rows[captionIndex].Cells[5].Style.SelectionBackColor = Color.Red;
                                        dataGridView1.Rows[captionIndex].Cells[5].Style.SelectionForeColor = Color.Red;
                                    }

                                    //add category name to row
                                    dataGridView1.Rows[captionIndex].Cells[6].Value = obj.Captions[j].category.name;
                                    //increase caption index for next row
                                    captionIndex += 1;

                                }
                            }
                        }
                    }
                }
                else if (filter == "Mix")
                {
                    //get number of captions under a category
                    for (int j = 0; j < obj.Captions.Count; j++)
                    {
                        if (obj.Captions[j].category.name == categoryFilter && obj.Captions[j].scopes.Contains(scopeFilter))
                        {
                            categoryCaptionCount++;
                        }
                    }
                    //if captions exist under a category then display them in datagrid
                    if (categoryCaptionCount > 0)
                    {
                        dataGridView1.Rows.Add(categoryCaptionCount + 1);
                        //ccreate tile for category in the datagrid
                        mergeCell(0, 6, captionIndex, categoryFilter);
                        //save records of tiles indices
                        gridTilesIndex.Add(captionIndex);
                        //increase row index 
                        captionIndex += 1;
                        //create rows in the datagridview
                        for (int j = 0; j < obj.Captions.Count; j++)
                        {
                            //for each category
                            if (obj.Captions[j].category.name == categoryFilter && obj.Captions[j].scopes.Contains(scopeFilter))
                            {
                                //add captions into rows of datagrid
                                dataGridView1.Rows[captionIndex].Cells[0].Value = obj.Captions[j].name;
                                //add extensions of captions
                                if (db.Captions[j].extensions.Length > 0)
                                {
                                    //foreach extension of a caption
                                    for (int w = 0; w < db.Captions[j].extensions.Length; w++)
                                    {
                                        //add extensions to cell 2,3,4, and 5
                                        dataGridView1.Rows[captionIndex].Cells[w + 1].Value = obj.Captions[j].extensions[w];
                                    }
                                }
                                //set toggle field

                                if (obj.Captions[j].goToNextImage)
                                {
                                    dataGridView1.Rows[captionIndex].Cells[5].Value = "Y";
                                    dataGridView1.Rows[captionIndex].Cells[5].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                                    dataGridView1.Rows[captionIndex].Cells[5].Style.ForeColor = Color.Black;
                                    dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Style.BackColor = Color.White;
                                    dataGridView1.Rows[captionIndex].Cells[5].Style.SelectionBackColor = SystemColors.Window;
                                    dataGridView1.Rows[captionIndex].Cells[5].Style.SelectionForeColor = SystemColors.ControlText;

                                }
                                else
                                {
                                    dataGridView1.Rows[captionIndex].Cells[5].Value = "";
                                    dataGridView1.Rows[captionIndex].Cells[5].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                                    dataGridView1.Rows[captionIndex].Cells[5].Style.ForeColor = Color.Red;
                                    dataGridView1.Rows[captionIndex].Cells[5].Style.BackColor = Color.Red;
                                    dataGridView1.Rows[captionIndex].Cells[5].Style.SelectionBackColor = Color.Red;
                                    dataGridView1.Rows[captionIndex].Cells[5].Style.SelectionForeColor = Color.Red;
                                }


                                //add category name to row
                                dataGridView1.Rows[captionIndex].Cells[6].Value = obj.Captions[j].category.name;

                                //increase caption index for next row
                                captionIndex += 1;

                            }
                        }
                        categoryCaptionCount = 0;
                    }
                }
            }

            for(int i = dataGridView1.RowCount-1;i>=0; i--)
            {
                if (dataGridView1.Rows[i].Cells[0].Value == null)
                {
                    dataGridView1.Rows.RemoveAt(i);
                }
            }

        }

        private void mergeCell(int start, int length, int rowIndex, string categoryName)
        {
            HMergedCell pCell;
            int nOffset = start;
            try
            {
                for (int j = nOffset; j < nOffset + length; j++)
                {
                    dataGridView1.Rows[rowIndex].Cells[j] = new HMergedCell();
                    pCell = (HMergedCell)dataGridView1.Rows[rowIndex].Cells[j];
                    pCell.LeftColumn = nOffset;
                    pCell.RightColumn = nOffset + length - 1;
                }

                dataGridView1.Rows[rowIndex].Cells[0].Value = categoryName;
                dataGridView1.Rows[rowIndex].ReadOnly = true;
            }
            catch (Exception)
            {
                return;
            }
        }

        #region Edit macro
        TextBox editBox;
        TextBox lvEditBox;
        TextBox categoryEditBox;
        private void loadEditTextBoxForListView()
        {
            lvEditBox = new TextBox();
            lvEditBox.Location = new System.Drawing.Point(0, 0);
            lvEditBox.Size = new System.Drawing.Size(0, 0);
            lvEditBox.Hide();
            //listView1.Controls.AddRange(new System.Windows.Forms.Control[] { this.lvEditBox });
            lvEditBox.Text = "";
            lvEditBox.BackColor = Color.Beige;
            lvEditBox.Font = new Font("Varanda", 13, FontStyle.Regular, GraphicsUnit.Pixel);
            //editBox.ForeColor = Color.Blue;
            lvEditBox.BorderStyle = BorderStyle.FixedSingle;
        }
        private void loadEditTextBox()
        {
            editBox = new TextBox();
            editBox.Location = new System.Drawing.Point(0, 0);
            editBox.Size = new System.Drawing.Size(0, 0);
            editBox.Hide();
            lbxComplete.Controls.AddRange(new System.Windows.Forms.Control[] { this.editBox });
            editBox.Text = "";
            editBox.BackColor = Color.Beige;
            editBox.Font = new Font("Varanda", 10, FontStyle.Regular, GraphicsUnit.Pixel);
            //editBox.ForeColor = Color.Blue;
            editBox.BorderStyle = BorderStyle.FixedSingle;


        }
        private void loadCategoryEditTextBox()
        {
            categoryEditBox = new TextBox();
            categoryEditBox.Location = new System.Drawing.Point(0, 0);
            categoryEditBox.Size = new System.Drawing.Size(0, 0);
            categoryEditBox.Hide();
            lbxComplete.Controls.AddRange(new System.Windows.Forms.Control[] { this.categoryEditBox });
            categoryEditBox.Text = "";
            categoryEditBox.BackColor = Color.Beige;
            categoryEditBox.Font = new Font("Varanda", 10, FontStyle.Regular, GraphicsUnit.Pixel);
            //editBox.ForeColor = Color.Blue;
            categoryEditBox.BorderStyle = BorderStyle.FixedSingle;


        }
        private void FocusOver(object sender, System.EventArgs e)
        {
            try
            {
                if (editBox.Text != "")
                {
                    lbxScopes.Items[lbxScopes.SelectedIndex] = editBox.Text;
                }
                editBox.Hide();
                db = collection.EditScope(oldScopes[scopeIndex], lbxScopes.SelectedItem.ToString());
                if (lbxCategories.SelectedIndex == -1)
                {
                    loadGrid(db, "Scope", string.Empty, lbxScopes.SelectedItem.ToString());
                }
                else
                {
                    loadGrid(db, "Mix", lbxCategories.SelectedItem.ToString(), lbxScopes.SelectedItem.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void FocusOverCategory(object sender, System.EventArgs e)
        {
            if (categoryEditBox.Text != "")
            {
                lbxCategories.Items[lbxCategories.SelectedIndex] = categoryEditBox.Text;
            }
            categoryEditBox.Hide();
            //update listview after edit
            Model obj = collection.EditCategory(oldCats[selectedindex], lbxCategories.SelectedItem.ToString());
            if (lbxScopes.SelectedIndex == -1)
            {
                loadGrid(obj, "Category", lbxCategories.SelectedItem.ToString(), string.Empty);
            }
            else
            {

                loadGrid(obj, "Mix", lbxCategories.SelectedItem.ToString(), lbxScopes.SelectedItem.ToString());

            }
        }

        private void FocusOverAddedCaptionEditbox(object sender, System.EventArgs e)
        {

        }
        private void EditOver(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            try
            {
                if (e.KeyChar == 13)
                {
                    if (editBox.Text != "")
                    {
                        lbxScopes.Items[lbxScopes.SelectedIndex] = editBox.Text;
                    }

                    editBox.Hide();
                    db = collection.EditScope(oldScopes[scopeIndex], lbxScopes.SelectedItem.ToString());
                    if (lbxCategories.SelectedIndex == -1)
                    {
                        loadGrid(db, "Scope", string.Empty, lbxScopes.SelectedItem.ToString());
                    }
                    else
                    {
                        loadGrid(db, "Mix", lbxCategories.SelectedItem.ToString(), lbxScopes.SelectedItem.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void EditOverCategory(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {

            if (e.KeyChar == 13)
            {
                if (categoryEditBox.Text != "")
                {
                    lbxCategories.Items[lbxCategories.SelectedIndex] = categoryEditBox.Text;
                }

                categoryEditBox.Hide();
                //update listivew after edit
                Model obj = collection.EditCategory(oldCats[selectedindex], lbxCategories.SelectedItem.ToString());
                if (lbxScopes.SelectedIndex == -1)
                {
                    loadGrid(obj, "Category", lbxCategories.SelectedItem.ToString(), string.Empty);
                }
                else
                {

                    loadGrid(obj, "Mix", lbxCategories.SelectedItem.ToString(), lbxScopes.SelectedItem.ToString());

                }
            }
        }

        private void EditOverAddedCaption(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {


            if (e.KeyChar == 13)
            {
                // string lvOldName = listView1.FocusedItem.Text;
                if (lvEditBox.Text != "")
                {
                    // listView1.FocusedItem.Text = lvEditBox.Text;

                    // db = collection.EditAddedCaption(lvOldName, listView1.FocusedItem.Group.ToString(), listView1.FocusedItem.Text);
                }

                lvEditBox.Hide();


            }
        }

        private void lbxScopes_KeyPress(object sender, KeyPressEventArgs e)
        {
            //if (e.KeyChar == 13)
            //    CreateEditBox(sender);
        }

        private void lbxScopes_DoubleClick(object sender, EventArgs e)
        {
            // CreateEditBox(sender);
        }

        private void lbxScopes_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyData == Keys.F2)
            //    CreateEditBox(sender);
        }


        private void CreateEditBox1(Rectangle r, string itemText, ListBox lbx)
        {

            editBox.Location = new System.Drawing.Point(r.X, r.Y);
            editBox.Size = new System.Drawing.Size(r.Width, r.Height - 15);
            editBox.Show();
            lbx.Controls.AddRange(new System.Windows.Forms.Control[] { this.editBox });
            editBox.Text = itemText;
            editBox.Focus();
            editBox.SelectAll();
            editBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EditOver);
            editBox.LostFocus += new System.EventHandler(this.FocusOver);
        }

        private void CreateCategoryEditBox(Rectangle r, string itemText, ListBox lbx)
        {

            categoryEditBox.Location = new System.Drawing.Point(r.X, r.Y);
            categoryEditBox.Size = new System.Drawing.Size(r.Width, r.Height - 10);
            categoryEditBox.Show();
            lbx.Controls.AddRange(new System.Windows.Forms.Control[] { this.categoryEditBox });
            categoryEditBox.Text = itemText;
            categoryEditBox.Focus();
            categoryEditBox.SelectAll();
            categoryEditBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EditOverCategory);
            categoryEditBox.LostFocus += new System.EventHandler(this.FocusOverCategory);
        }

        private void CreateAddedCaptionsEditBox(Rectangle r, string itemText, DataGridView lbx)
        {

            lvEditBox.Location = new System.Drawing.Point(r.X, r.Y);
            lvEditBox.Size = new System.Drawing.Size(r.Width, r.Height);
            lvEditBox.Show();
            lbx.Controls.AddRange(new System.Windows.Forms.Control[] { this.lvEditBox });
            lvEditBox.Text = itemText;
            lvEditBox.Focus();
            lvEditBox.SelectAll();
            lvEditBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.EditOverAddedCaption);
            lvEditBox.LostFocus += new System.EventHandler(this.FocusOverAddedCaptionEditbox);
        }
        #endregion

        private void panel24_Paint(object sender, PaintEventArgs e)
        {

        }

        private void tableLayoutPanel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void listView1_ItemActivate(object sender, EventArgs e)
        {

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {


        }

        private void listView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {

        }

        private void listView1_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void lbxCategories_Click(object sender, EventArgs e)
        {



        }

        //edit category item
        List<string> oldCats;
        int selectedindex;
        private void Item1_Menu_Click(object sender, EventArgs e)
        {


        }

        //edit added captions in listView
        private void EditListItem_Click(object sender, EventArgs e)
        {


        }

        private void editToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                if (lbxCategories.SelectedIndex > -1)
                {


                    oldCats = new List<string>();
                    selectedindex = lbxCategories.SelectedIndex;
                    foreach (var item in lbxCategories.Items) { oldCats.Add(item.ToString()); }
                    CreateCategoryEditBox(lbxCategories.GetItemRectangle(lbxCategories.SelectedIndex), lbxCategories.SelectedItem.ToString(), lbxCategories);
                }
                else
                {
                    MessageBox.Show("Make sure to select an item");
                }
            }
            catch (Exception ex)
            { MessageBox.Show(ex.Message); }
        }
        //remove category from listbox
        private void remove_Menu_Click(object sender, EventArgs e)
        {

        }
        private void deleteToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (lbxCategories.SelectedIndex > -1)
            {
                DialogResult result = MessageBox.Show("All Captions in this category will be permanently deleted. Do want to continue?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                  //  try
                    {
                        string value = lbxCategories.SelectedItem.ToString();
                         Model obj = collection.RemoveCategory(value);
                        if (lbxScopes.SelectedIndex == -1)
                        {
                            loadGrid(db, "All", string.Empty, String.Empty);
                        }
                        else
                        {

                            //loadGrid(db, "Scope", String.Empty, lbxScopes.SelectedItem.ToString());
                        }
                        lbxCategories.Items.RemoveAt(lbxCategories.SelectedIndex);
                    }
                    //catch (Exception ex)
                    //{
                    //    MessageBox.Show(ex.Message);
                    //}
                }
            }
            else { MessageBox.Show("Please select an item"); }
        }

        //remove scope 
        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (lbxScopes.SelectedIndex > -1)
            {
                DialogResult result = MessageBox.Show("All Captions under this scope will be permanently deleted. Do want to continue?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        Model db = collection.RemoveScope(lbxScopes.SelectedItem.ToString());
                        if (lbxCategories.SelectedIndex == -1)
                        {
                            loadGrid(db, "All", string.Empty, String.Empty);
                        }
                        else
                        {
                            loadGrid(db, "Category", lbxCategories.SelectedItem.ToString(), String.Empty);
                        }
                        lbxScopes.Items.RemoveAt(lbxScopes.SelectedIndex);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            else
            {
                MessageBox.Show("Please select an item");
            }
        }

        //remove added caption 
        private void removeListItem_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            dataGridView1.CurrentCell.ReadOnly = false;

        }

        //edit caption from grid
        private void gridEditMenu_Click(object sender, EventArgs e)
        {
            editGrid();

        }
        private void editGrid()
        {
            string oldtext = "";
            //if row is caption row and selected cell is caption cell or sub caption cell
            if (dataGridView1.CurrentCell.ColumnIndex < 5 && dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].GetType().Name != "HMergeCell")
            {


                try
                {
                    frmEdit = new frmEditCaption(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[dataGridView1.CurrentCell.ColumnIndex].Value.ToString(), dataGridView1.CurrentCell.ColumnIndex);
                    oldtext = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[dataGridView1.CurrentCell.ColumnIndex].Value.ToString();
                }
                catch (Exception)
                {
                    frmEdit = new frmEditCaption("", dataGridView1.CurrentCell.ColumnIndex);
                    oldtext = "";
                }

                frmEdit.ShowDialog();
                if (clickSave)
                {
                    dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[dataGridView1.CurrentCell.ColumnIndex].Value = EditedCaptionText;
                    if (dataGridView1.CurrentCell.ColumnIndex == 0)
                    {
                        db = collection.EditCaption(oldtext, dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString(), EditedCaptionText);
                    }
                    else
                    {
                        db = collection.AddExtensionToCaption(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString(), dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString(), EditedCaptionText, dataGridView1.CurrentCell.ColumnIndex - 1);
                    }
                    clickSave = false;
                  

                }
            }
        }


        //remove a caption from the grid
        private void gridRemoveMenu_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("This caption will be permanently deleted. Do want to continue?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                try
                {
                    if (lbxScopes.SelectedIndex >= 0)
                    {

                        db = collection.RemoveCaptionFromScope(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[dataGridView1.CurrentCell.ColumnIndex].Value.ToString(), dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString(), lbxScopes.SelectedItem.ToString());
                    }
                    else if (lbxScopes.SelectedIndex == -1)
                    {
                        db = collection.RemoveCaption(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[dataGridView1.CurrentCell.ColumnIndex].Value.ToString(), dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString());
                    }
                    dataGridView1.Rows.RemoveAt(dataGridView1.CurrentCell.RowIndex);



                }
                catch (Exception ex)
                {
                    dataGridView1.Rows.RemoveAt(dataGridView1.CurrentCell.RowIndex);


                }
            }
        }

        private void dataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {

        }

        public static string EditedCaptionText;
        frmEditCaption frmEdit;
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                //if cell is not null
                if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[dataGridView1.CurrentCell.ColumnIndex].Value != null)
                {
                    //if row is caption text row and cell contains caption data
                    if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].GetType().Name!="HMergedCell" && dataGridView1.CurrentCell.ColumnIndex < 5)
                    {
                        string captionStr = "";
                        if (lbxTag1.SelectedIndex > -1) { captionStr = captionStr + lbxTag1.SelectedItem.ToString() + " "; }
                        //append selected item in tag 1 listbox to caption string
                        if (lbxTag2.SelectedIndex > -1) { captionStr = captionStr + lbxTag2.SelectedItem.ToString() + " "; }
                        //append selected item in tag 3 listbox to caption string
                        if (lbxTag3.SelectedIndex > -1) { captionStr = captionStr + lbxTag3.SelectedItem.ToString() + " "; }
                        //append selected item in tag 4 listbox to caption string
                        if (lbxTag4.SelectedIndex > -1) { captionStr = captionStr + lbxTag4.SelectedItem.ToString() + " "; }

                         captionStr = captionStr + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()+" ";

                        if (dataGridView1.CurrentCell.ColumnIndex > 0)
                        {
                            captionStr +=  dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[dataGridView1.CurrentCell.ColumnIndex].Value.ToString()+" ";
                        }

                        IntPtr lastWindowHandle = GetWindow(Process.GetCurrentProcess().MainWindowHandle, (uint)GetWindow_Cmd.GW_HWNDNEXT);
                        while (true)
                        {
                            IntPtr temp = GetParent(lastWindowHandle);
                            if (temp.Equals(IntPtr.Zero)) break;
                            lastWindowHandle = temp;
                        }
                        SetForegroundWindow(lastWindowHandle);
                        //System.Threading.Thread.Sleep(1000);
                        SendKeys.SendWait(captionStr);

                        //if next image is set to true then
                        if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value.ToString() == "Y")
                        {
                            //send next image command
                            SendKeys.SendWait("%{n}");
                            //reset current window to caption builder app
                            //SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);
                        }

                       

                    }

                    //setting toggle fields for next image
                    //if the selected cell is a Y cell
                    if (e.ColumnIndex == 5 && dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[dataGridView1.CurrentCell.ColumnIndex].Value.ToString() == "")
                    {
                        // if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[dataGridView1.CurrentCell.ColumnIndex].Style.BackColor==Color.Red)
                        //{
                        db = collection.EditCaptionNextImageStatus(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString(), dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString(), true);
                        dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value = "Y";
                        dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Style.ForeColor = Color.Black;
                        dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Style.BackColor = Color.White;
                        dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Style.SelectionBackColor = SystemColors.Window;
                        dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Style.SelectionForeColor = SystemColors.ControlText;
                    }
                    else if (dataGridView1.CurrentCell.ColumnIndex == 5 && dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[dataGridView1.CurrentCell.ColumnIndex].Value.ToString() == "Y")
                    {
                        db = collection.EditCaptionNextImageStatus(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString(), dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString(), false);
                        //dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Style.ForeColor = Color.Red;
                        dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value = "";
                        dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Style.BackColor = Color.Red;
                        dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Style.SelectionBackColor = Color.Red;
                        dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Style.SelectionForeColor = Color.Red;
                        //}
                    }
                }
            }
            catch (Exception)
            {
                return;
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void ckbNext_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lbxComplete.SelectedIndex = -1;
        }

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        #region drag code for list box
        private void lbx_MouseDown(object sender, MouseEventArgs e)
        {

            ListBox lst = sender as ListBox;

            // Only use the right mouse button.
            if (e.Button != MouseButtons.Left) return;

            // Find the item under the mouse.
            int index = lst.IndexFromPoint(e.Location);
            lst.SelectedIndex = index;
            if (index < 0) return;

            // Drag the item.
            DragItem drag_item = new DragItem(lst, index, lst.Items[index]);
            lst.DoDragDrop(drag_item, DragDropEffects.Move);
        }

        private void lbx_DragEnter(object sender, DragEventArgs e)
        {
            ListBox lst = sender as ListBox;

            // Allow a Move for DragItem objects that are
            // dragged to the control that started the drag.
            if (!e.Data.GetDataPresent(typeof(DragItem)))
            {
                // Not a DragItem. Don't allow it.
                e.Effect = DragDropEffects.None;
            }
            else if ((e.AllowedEffect & DragDropEffects.Move) == 0)
            {
                // Not a Move. Do not allow it.
                e.Effect = DragDropEffects.None;
            }
            else
            {
                // Get the DragItem.
                DragItem drag_item = (DragItem)e.Data.GetData(typeof(DragItem));

                // Verify that this is the control that started the drag.
                if (drag_item.Client != lst)
                {
                    // Not the congtrol that started the drag. Do not allow it.
                    e.Effect = DragDropEffects.None;
                }
                else
                {
                    // Allow it.
                    e.Effect = DragDropEffects.Move;
                }
            }

        }

        private void lbx_DragOver(object sender, DragEventArgs e)
        {
            // Do nothing if the drag is not allowed.
            if (e.Effect != DragDropEffects.Move) return;

            ListBox lst = sender as ListBox;

            // Select the item at this screen location.
            lst.SelectedIndex = MyExtension.IndexFromScreenPoint(lst, new Point(e.X, e.Y));
        }

        private void lbx_DragDrop(object sender, DragEventArgs e)
        {
            // Drop the item here.

            ListBox lst = sender as ListBox;

            // Get the ListBox item data.
            DragItem drag_item = (DragItem)e.Data.GetData(typeof(DragItem));

            // Get the index of the item at this position.
            int new_index = MyExtension.IndexFromScreenPoint(lst, new Point(e.X, e.Y));

            // If the item was dropped after all
            // of the items, move it to the end.
            if (new_index == -1) new_index = lst.Items.Count - 1;

            // Remove the item from its original position.
            lst.Items.RemoveAt(drag_item.Index);

            // Insert the item in its new position.
            lst.Items.Insert(new_index, drag_item.Item);

            // Select the item.
            lst.SelectedIndex = new_index;

        }
        #endregion
        #region drag code  for datagrid
        private Rectangle dragBoxFromMouseDown;
        private int rowIndexFromMouseDown;
        private int rowIndexOfItemUnderMouseToDrop;
        private void dataGridView1_MouseMove(object sender, MouseEventArgs e)
        {

            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                // If the mouse moves outside the rectangle, start the drag.
                if (dragBoxFromMouseDown != Rectangle.Empty &&
                !dragBoxFromMouseDown.Contains(e.X, e.Y))
                {
                    // Proceed with the drag and drop, passing in the list item.                    
                    DragDropEffects dropEffect = dataGridView1.DoDragDrop(
                          dataGridView1.Rows[rowIndexFromMouseDown],
                          DragDropEffects.Move);
                }

            }
        }

        private void dataGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            rowIndexFromMouseDown = dataGridView1.HitTest(e.X, e.Y).RowIndex;

            if (rowIndexFromMouseDown != -1)
            {
                // Remember the point where the mouse down occurred. 
                // The DragSize indicates the size that the mouse can move 
                // before a drag event should be started.                
                Size dragSize = SystemInformation.DragSize;

                // Create a rectangle using the DragSize, with the mouse position being
                // at the center of the rectangle.
                dragBoxFromMouseDown = new Rectangle(
                          new Point(
                            e.X - (dragSize.Width / 2),
                            e.Y - (dragSize.Height / 2)),
                      dragSize);
            }
            else
                // Reset the rectangle if the mouse is not over an item in the ListBox.
                dragBoxFromMouseDown = Rectangle.Empty;
        }


        private void dataGridView1_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;

        }

        private void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            // The mouse locations are relative to the screen, so they must be 
            // converted to client coordinates.
            Point clientPoint = dataGridView1.PointToClient(new Point(e.X, e.Y));

            // Get the row index of the item the mouse is below. 
            rowIndexOfItemUnderMouseToDrop = dataGridView1.HitTest(clientPoint.X, clientPoint.Y).RowIndex;

            // If the drag operation was a move then remove and insert the row.
            if (e.Effect == DragDropEffects.Move)
            {
                try
                {
                    //when dragging category tile in datagrid view
                    if (dataGridView1.Rows[rowIndexOfItemUnderMouseToDrop].Cells[0].GetType().Name.ToString() == "HMergedCell" && lbxCategories.Items.Contains(dataGridView1.Rows[rowIndexFromMouseDown].Cells[0].Value.ToString()))
                    {
                        DataGridViewRow rowToMove = e.Data.GetData(typeof(DataGridViewRow)) as DataGridViewRow;
                        dataGridView1.Rows.RemoveAt(rowIndexFromMouseDown);
                        int count = 0;

                        dataGridView1.Rows.Insert(rowIndexOfItemUnderMouseToDrop, rowToMove);
                        string cat = dataGridView1.Rows[rowIndexOfItemUnderMouseToDrop].Cells[0].Value.ToString();


                        //when dragging item from buttom to top of grid
                        if (rowIndexOfItemUnderMouseToDrop < rowIndexFromMouseDown)
                        {
                            int j = 1;
                            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                            {
                                //omit category tile rows
                                if (!lbxCategories.Items.Contains(dataGridView1.Rows[i].Cells[0].Value.ToString()))
                                {
                                    //if row is in moved category
                                    if (cat == dataGridView1.Rows[i].Cells[6].Value.ToString())
                                    {
                                        DataGridViewRow row = dataGridView1.Rows[i];
                                        //remove row
                                        dataGridView1.Rows.RemoveAt(i);
                                        //insert row at the dropped location
                                        dataGridView1.Rows.Insert(rowIndexOfItemUnderMouseToDrop + j, row);
                                        j++;
                                        // i = 0;
                                        ///count++;
                                    }
                                }
                            }
                        }
                        //when dragging item from top of grid to buttom
                        else if (rowIndexOfItemUnderMouseToDrop > rowIndexFromMouseDown)
                        {
                            string catAbove = dataGridView1.Rows[rowIndexOfItemUnderMouseToDrop - 1].Cells[0].Value.ToString();
                            int j = 0;
                            for (int i = dataGridView1.Rows.Count - 1; i >= 0; i--)
                            {
                                //omit category tiles 
                                if (dataGridView1.Rows[i].Cells[0].GetType().Name.ToString() != "HMergedCell")
                                {
                                    if (cat == dataGridView1.Rows[i].Cells[6].Value.ToString())
                                    {
                                        DataGridViewRow row = dataGridView1.Rows[i];
                                        dataGridView1.Rows.RemoveAt(i);


                                        dataGridView1.Rows.Insert(rowIndexOfItemUnderMouseToDrop - j, row);


                                        j++;
                                        // i = 0;
                                        count++;
                                    }

                                }

                            }
                            int w = 1;
                            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                            {
                                //omit category tiles 
                                if (dataGridView1.Rows[i].Cells[0].GetType().Name.ToString() != "HMergedCell")
                                {
                                    if (catAbove == dataGridView1.Rows[i].Cells[6].Value.ToString())
                                    {
                                        DataGridViewRow row = dataGridView1.Rows[i];
                                        dataGridView1.Rows.RemoveAt(i);
                                        dataGridView1.Rows.Insert(rowIndexFromMouseDown + w, row);
                                        w++;
                                    }

                                }
                            }
                        }
                    }
                    //when dragging caption row in datagridview, if drag occurs between captions then
                    // use this comparison when caption can only be moved between categories [dataGridView1.Rows[rowIndexOfItemUnderMouseToDrop].Cells[6].Value.ToString() ]
                    else if (dataGridView1.Rows[rowIndexOfItemUnderMouseToDrop].Cells[0].GetType().Name.ToString() != "HMergedCell" && !lbxCategories.Items.Contains(dataGridView1.Rows[rowIndexFromMouseDown].Cells[0].Value.ToString()))
                    {
                        DataGridViewRow rowToMove = e.Data.GetData(typeof(DataGridViewRow)) as DataGridViewRow;
                        dataGridView1.Rows.RemoveAt(rowIndexFromMouseDown);

                        //if dragging from top to buttom
                        if (rowIndexOfItemUnderMouseToDrop > rowIndexFromMouseDown)
                        {
                            //insert dragged row
                            dataGridView1.Rows.Insert(rowIndexOfItemUnderMouseToDrop, rowToMove);


                        }
                        //if dragging caption from buttom to top
                        else
                        {
                            //insert dragged row
                            dataGridView1.Rows.Insert(rowIndexOfItemUnderMouseToDrop, rowToMove);

                        }

                    }
                }
                catch (Exception) { return; }
            }

        }

        private void dataGridView1_MouseUp(object sender, MouseEventArgs e)
        {

        }

        private void dataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Right) == MouseButtons.Right)
            {
                if (dataGridView1.CurrentCell.ColumnIndex < 5 && !gridTilesIndex.Contains(dataGridView1.CurrentCell.RowIndex))
                {
                    ContextMenuStrip gridMenu = new ContextMenuStrip();

                    ToolStripMenuItem removeSubMenu = new ToolStripMenuItem();
                    ToolStripMenuItem editSubMenu = new ToolStripMenuItem();

                    removeSubMenu.Text = "Remove Caption";
                    editSubMenu.Text = "Edit Caption";
                    removeSubMenu.Click += new EventHandler(gridRemoveMenu_Click);
                    editSubMenu.Click += new EventHandler(gridEditMenu_Click);
                    gridMenu.Items.Add(editSubMenu);
                    if (dataGridView1.CurrentCell.ColumnIndex == 0) { gridMenu.Items.Add(removeSubMenu); }


                    dataGridView1.ContextMenuStrip = gridMenu;

                }
                else
                {
                    dataGridView1.ContextMenuStrip = null;
                }
            }
        }

        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {

        }
        #endregion

        #region Tags and complete items
        //load complete items
        private void loadCompleteItems()
        {
            if (db.CompleteItems != null)
            {
                if (db.CompleteItems.Count > 0)
                {
                    foreach (var item in db.CompleteItems.ToList())
                    {
                        lbxComplete.Items.Add(item.text);
                    }
                }
            }
        }
        //add tag
        private void btnAddTag1_Click(object sender, EventArgs e)
        {
            tagForm = new frmAddTag();
            tagForm.ShowDialog();
            if (clickSave)
            {
                // db = collection.AddTag(tagText, 1);
                lbxTag1.Items.Add(tagText);
                clickSave = false;
            }
        }

        private void btnAddTag2_Click(object sender, EventArgs e)
        {
            tagForm = new frmAddTag();
            tagForm.ShowDialog();
            if (clickSave)
            {
                // db = collection.AddTag(tagText, 2);
                lbxTag2.Items.Add(tagText);
                clickSave = false;
            }
        }

        private void btnAddTag3_Click(object sender, EventArgs e)
        {
            tagForm = new frmAddTag();
            tagForm.ShowDialog();
            if (clickSave)
            {
                ///db = collection.AddTag(tagText, 3);
                lbxTag3.Items.Add(tagText);
                clickSave = false;
            }
        }

        private void btnAddTag4_Click(object sender, EventArgs e)
        {
            tagForm = new frmAddTag();
            tagForm.ShowDialog();
            if (clickSave)
            {
                // db = collection.AddTag(tagText, 4);
                lbxTag4.Items.Add(tagText);
                clickSave = false;
            }
        }

        private void lbxTag1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                btnAddTag1.PerformClick();
            }
        }
        private void lbxTag2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                btnAddTag2.PerformClick();
            }
        }
        private void lbxTag3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                btnAddTag3.PerformClick();
            }
        }
        private void lbxTag4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                btnAddTag4.PerformClick();
            }
        }

        //edit tag
        TextBox tag1EditBox;
        TextBox tag2EditBox;
        TextBox tag3EditBox;
        TextBox tag4EditBox;
        TextBox completeEditBox;
        private void loadTag1EditTextBox()
        {
            tag1EditBox = new TextBox();
            tag1EditBox.Location = new System.Drawing.Point(0, 0);
            tag1EditBox.Size = new System.Drawing.Size(0, 0);
            tag1EditBox.Hide();
            lbxTag1.Controls.AddRange(new System.Windows.Forms.Control[] { this.tag1EditBox });
            tag1EditBox.Text = "";
            tag1EditBox.BackColor = Color.Beige;
            tag1EditBox.Font = new Font("Varanda", 13, FontStyle.Regular, GraphicsUnit.Pixel);
            //editBox.ForeColor = Color.Blue;
            tag1EditBox.BorderStyle = BorderStyle.FixedSingle;
        }
        private void loadTag2EditTextBox()
        {
            tag2EditBox = new TextBox();
            tag2EditBox.Location = new System.Drawing.Point(0, 0);
            tag2EditBox.Size = new System.Drawing.Size(0, 0);
            tag2EditBox.Hide();
            lbxTag2.Controls.AddRange(new System.Windows.Forms.Control[] { this.tag2EditBox });
            tag2EditBox.Text = "";
            tag2EditBox.BackColor = Color.Beige;
            tag2EditBox.Font = new Font("Varanda", 13, FontStyle.Regular, GraphicsUnit.Pixel);
            //editBox.ForeColor = Color.Blue;
            tag2EditBox.BorderStyle = BorderStyle.FixedSingle;
        }
        private void loadTag3EditTextBox()
        {
            tag3EditBox = new TextBox();
            tag3EditBox.Location = new System.Drawing.Point(0, 0);
            tag3EditBox.Size = new System.Drawing.Size(0, 0);
            tag3EditBox.Hide();
            lbxTag3.Controls.AddRange(new System.Windows.Forms.Control[] { this.tag3EditBox });
            tag3EditBox.Text = "";
            tag3EditBox.BackColor = Color.Beige;
            tag3EditBox.Font = new Font("Varanda", 13, FontStyle.Regular, GraphicsUnit.Pixel);
            //editBox.ForeColor = Color.Blue;
            tag3EditBox.BorderStyle = BorderStyle.FixedSingle;
        }
        private void loadTag4EditTextBox()
        {
            tag4EditBox = new TextBox();
            tag4EditBox.Location = new System.Drawing.Point(0, 0);
            tag4EditBox.Size = new System.Drawing.Size(0, 0);
            tag4EditBox.Hide();
            lbxTag4.Controls.AddRange(new System.Windows.Forms.Control[] { this.tag4EditBox });
            tag4EditBox.Text = "";
            tag4EditBox.BackColor = Color.Beige;
            tag4EditBox.Font = new Font("Varanda", 13, FontStyle.Regular, GraphicsUnit.Pixel);
            //editBox.ForeColor = Color.Blue;
            tag4EditBox.BorderStyle = BorderStyle.FixedSingle;
        }
        private void loadCompleteEditTextBox()
        {
            completeEditBox = new TextBox();
            completeEditBox.Location = new System.Drawing.Point(0, 0);
            completeEditBox.Size = new System.Drawing.Size(0, 0);
            completeEditBox.Hide();
            lbxComplete.Controls.AddRange(new System.Windows.Forms.Control[] { this.completeEditBox });
            completeEditBox.Text = "";
            completeEditBox.BackColor = Color.Beige;
            completeEditBox.Font = new Font("Varanda", 13, FontStyle.Regular, GraphicsUnit.Pixel);
            //editBox.ForeColor = Color.Blue;
            completeEditBox.BorderStyle = BorderStyle.FixedSingle;
        }
        private void tag1FocusOver(object sender, System.EventArgs e)
        {
            try
            { string selectedDescription = lbxTag1.Items[lbxTag1.SelectedIndex].ToString();
                if (tag1EditBox.Text != "")
                {
                    lbxTag1.Items[lbxTag1.SelectedIndex] = tag1EditBox.Text;
                }
                tag1EditBox.Hide();
                db = collection.EditTag(selectedDescription, lbxTag1.SelectedItem.ToString(), 1);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void tag1EditOver(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {

            if (e.KeyChar == 13)
            {
                string selectedDescription = lbxTag1.Items[lbxTag1.SelectedIndex].ToString();
                if (tag1EditBox.Text != "")
                {
                    lbxTag1.Items[lbxTag1.SelectedIndex] = tag1EditBox.Text;
                }
                tag1EditBox.Hide();
                db = collection.EditTag(selectedDescription, lbxTag1.SelectedItem.ToString(), 1);

            }
        }

        private void tag2FocusOver(object sender, System.EventArgs e)
        {
            try
            {
                string selectedDescription = lbxTag2.Items[lbxTag2.SelectedIndex].ToString();
                if (tag2EditBox.Text != "")
                {
                    lbxTag2.Items[lbxTag2.SelectedIndex] = tag2EditBox.Text;
                }
                tag2EditBox.Hide();
                db = collection.EditTag(selectedDescription, lbxTag2.SelectedItem.ToString(), 2);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void tag2EditOver(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {

            if (e.KeyChar == 13)
            {
                string selectedDescription = lbxTag2.Items[lbxTag2.SelectedIndex].ToString();
                if (tag2EditBox.Text != "")
                {
                    lbxTag2.Items[lbxTag2.SelectedIndex] = tag2EditBox.Text;
                }
                tag2EditBox.Hide();
                db = collection.EditTag(selectedDescription, lbxTag2.SelectedItem.ToString(), 2);

            }
        }

        private void tag3FocusOver(object sender, System.EventArgs e)
        {
            try
            {
                string selectedDescription = lbxTag3.Items[lbxTag3.SelectedIndex].ToString();
                if (tag3EditBox.Text != "")
                {
                    lbxTag3.Items[lbxTag3.SelectedIndex] = tag3EditBox.Text;
                }
                tag3EditBox.Hide();
                db = collection.EditTag(selectedDescription, lbxTag3.SelectedItem.ToString(), 3);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void tag3EditOver(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {

            if (e.KeyChar == 13)
            {
                string selectedDescription = lbxTag3.Items[lbxTag3.SelectedIndex].ToString();
                if (tag3EditBox.Text != "")
                {
                    lbxTag3.Items[lbxTag3.SelectedIndex] = tag3EditBox.Text;
                }
                tag3EditBox.Hide();
                db = collection.EditTag(selectedDescription, lbxTag3.SelectedItem.ToString(), 3);

            }
        }

        private void tag4FocusOver(object sender, System.EventArgs e)
        {
            try
            {
                string selectedDescription = lbxTag4.Items[lbxTag4.SelectedIndex].ToString();
                if (tag4EditBox.Text != "")
                {
                    lbxTag4.Items[lbxTag4.SelectedIndex] = tag4EditBox.Text;
                }
                tag4EditBox.Hide();
                db = collection.EditTag(selectedDescription, lbxTag4.SelectedItem.ToString(), 4);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void completeFocusOver(object sender, System.EventArgs e)
        {
            try
            {
                string selectedDescription = lbxComplete.Items[lbxComplete.SelectedIndex].ToString();
                if (completeEditBox.Text != "")
                {
                    lbxComplete.Items[lbxComplete.SelectedIndex] = completeEditBox.Text;
                }
                completeEditBox.Hide();
                //db = collection.EditTag(selectedDescription, lbxTag4.SelectedItem.ToString(), 4);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void tag4EditOver(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {

            if (e.KeyChar == 13)
            {
                string selectedDescription = lbxTag4.Items[lbxTag4.SelectedIndex].ToString();
                if (tag4EditBox.Text != "")
                {
                    lbxTag4.Items[lbxTag4.SelectedIndex] = tag4EditBox.Text;
                }
                tag4EditBox.Hide();
                db = collection.EditTag(selectedDescription, lbxTag4.SelectedItem.ToString(), 4);

            }
        }
        private void completeEditOver(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {

            if (e.KeyChar == 13)
            {
                string selectedDescription = lbxComplete.Items[lbxComplete.SelectedIndex].ToString();
                if (completeEditBox.Text != "")
                {
                    lbxComplete.Items[lbxComplete.SelectedIndex] = completeEditBox.Text;
                }
                completeEditBox.Hide();
                // db = collection.EditTag(selectedDescription, lbxTag4.SelectedItem.ToString(), 4);

            }
        }

        private void CreateTagEditBox1(Rectangle r, string itemText, ListBox lbx)
        {


            tag1EditBox.Location = new System.Drawing.Point(r.X, r.Y);
            tag1EditBox.Size = new System.Drawing.Size(r.Width, r.Height);
            tag1EditBox.Show();
            lbx.Controls.AddRange(new System.Windows.Forms.Control[] { this.tag1EditBox });
            tag1EditBox.Text = itemText;
            tag1EditBox.Focus();
            tag1EditBox.SelectAll();
            tag1EditBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.tag1EditOver);
            tag1EditBox.LostFocus += new System.EventHandler(this.tag1FocusOver);
        }
        private void CreateTagEditBox2(Rectangle r, string itemText, ListBox lbx)
        {
            tag2EditBox.Location = new System.Drawing.Point(r.X, r.Y);
            tag2EditBox.Size = new System.Drawing.Size(r.Width, r.Height);
            tag2EditBox.Show();
            lbx.Controls.AddRange(new System.Windows.Forms.Control[] { this.tag2EditBox });
            tag2EditBox.Text = itemText;
            tag2EditBox.Focus();
            tag2EditBox.SelectAll();
            tag2EditBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.tag2EditOver);
            tag2EditBox.LostFocus += new System.EventHandler(this.tag2FocusOver);
        }
        private void CreateTagEditBox3(Rectangle r, string itemText, ListBox lbx)
        {
            tag3EditBox.Location = new System.Drawing.Point(r.X, r.Y);
            tag3EditBox.Size = new System.Drawing.Size(r.Width, r.Height);
            tag3EditBox.Show();
            lbx.Controls.AddRange(new System.Windows.Forms.Control[] { this.tag3EditBox });
            tag3EditBox.Text = itemText;
            tag3EditBox.Focus();
            tag3EditBox.SelectAll();
            tag3EditBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.tag3EditOver);
            tag3EditBox.LostFocus += new System.EventHandler(this.tag3FocusOver);
        }
        private void CreateTagEditBox4(Rectangle r, string itemText, ListBox lbx)
        {
            tag4EditBox.Location = new System.Drawing.Point(r.X, r.Y);
            tag4EditBox.Size = new System.Drawing.Size(r.Width, r.Height);
            tag4EditBox.Show();
            lbx.Controls.AddRange(new System.Windows.Forms.Control[] { this.tag4EditBox });
            tag4EditBox.Text = itemText;
            tag4EditBox.Focus();
            tag4EditBox.SelectAll();
            tag4EditBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.tag4EditOver);
            tag4EditBox.LostFocus += new System.EventHandler(this.tag4FocusOver);


        }
        private void CreateCompleteEditBox4(Rectangle r, string itemText, ListBox lbx)
        {
            completeEditBox.Location = new System.Drawing.Point(r.X, r.Y);
            completeEditBox.Size = new System.Drawing.Size(r.Width, r.Height);
            completeEditBox.Show();
            lbx.Controls.AddRange(new System.Windows.Forms.Control[] { this.completeEditBox });
            completeEditBox.Text = itemText;
            completeEditBox.Focus();
            completeEditBox.SelectAll();
            completeEditBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.completeEditOver);
            completeEditBox.LostFocus += new System.EventHandler(this.completeFocusOver);


        }

        private void lbxTag1_MouseUp(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Right) == MouseButtons.Right)
            {
                getMenuStrip(lbxTag1);
            }
        }
        private void lbxTag2_MouseUp(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Right) == MouseButtons.Right)
            {
                getMenuStrip(lbxTag2);
            }
        }
        private void lbxTag3_MouseUp(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Right) == MouseButtons.Right)
            {
                getMenuStrip(lbxTag3);
            }
        }
        private void lbxTag4_MouseUp(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Right) == MouseButtons.Right)
            {
                getMenuStrip(lbxTag4);
            }
        }

        private void lbxComplete_MouseUp(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Right) == MouseButtons.Right)
            {
                getMenuStrip(lbxComplete);
            }
        }
        private void getMenuStrip(ListBox lbx)
        {
            if (lbx.SelectedIndex > -1)
            {
                ContextMenuStrip menu = new ContextMenuStrip();

                ToolStripMenuItem removeSubMenu = new ToolStripMenuItem();
                ToolStripMenuItem editSubMenu = new ToolStripMenuItem();

                removeSubMenu.Text = "Remove Tag";
                editSubMenu.Text = "Edit Tag";
                if (lbx.Name == lbxTag1.Name)
                {
                    removeSubMenu.Click += new EventHandler(tag1RemoveMenu_Click);
                    editSubMenu.Click += new EventHandler(tag1EditMenu_Click);
                } else if (lbx.Name == lbxTag2.Name)
                {
                    removeSubMenu.Click += new EventHandler(tag2RemoveMenu_Click);
                    editSubMenu.Click += new EventHandler(tag2EditMenu_Click);
                }
                else if (lbx.Name == lbxTag3.Name)
                {
                    removeSubMenu.Click += new EventHandler(tag3RemoveMenu_Click);
                    editSubMenu.Click += new EventHandler(tag3EditMenu_Click);
                }
                else if (lbx.Name == lbxTag4.Name)
                {
                    removeSubMenu.Click += new EventHandler(tag4RemoveMenu_Click);
                    editSubMenu.Click += new EventHandler(tag4EditMenu_Click);
                }
                else if (lbx.Name == lbxComplete.Name)
                {
                    removeSubMenu.Click += new EventHandler(completeRemoveMenu_Click);
                    editSubMenu.Click += new EventHandler(completeEditMenu_Click);
                }

                menu.Items.Add(editSubMenu);
                menu.Items.Add(removeSubMenu);


                lbx.ContextMenuStrip = menu;

            }
            else
            {
                lbx.ContextMenuStrip = null;
            }
        }

        //edit and remove for tag 1 listbox
        private void tag1RemoveMenu_Click(object sender, EventArgs e)
        {
            lbxTag1.Items.RemoveAt(lbxTag1.SelectedIndex);
        }

        private void tag1EditMenu_Click(object sender, EventArgs e)
        {
            CreateTagEditBox1(lbxTag1.GetItemRectangle(lbxTag1.SelectedIndex), lbxTag1.SelectedItem.ToString(), lbxTag1);
        }
        //edit and remove for tag 2 listbox
        private void tag2RemoveMenu_Click(object sender, EventArgs e)
        {
            lbxTag2.Items.RemoveAt(lbxTag2.SelectedIndex);
        }

        private void tag2EditMenu_Click(object sender, EventArgs e)
        {
            CreateTagEditBox2(lbxTag2.GetItemRectangle(lbxTag2.SelectedIndex), lbxTag2.SelectedItem.ToString(), lbxTag2);
        }

        //edit and remove for tag 3 listbox
        private void tag3RemoveMenu_Click(object sender, EventArgs e)
        {
            lbxTag3.Items.RemoveAt(lbxTag3.SelectedIndex);
        }

        private void tag3EditMenu_Click(object sender, EventArgs e)
        {
            CreateTagEditBox3(lbxTag3.GetItemRectangle(lbxTag3.SelectedIndex), lbxTag3.SelectedItem.ToString(), lbxTag3);
        }
        //edit and remove for tag 4 listbox
        private void tag4RemoveMenu_Click(object sender, EventArgs e)
        {
            lbxTag4.Items.RemoveAt(lbxTag4.SelectedIndex);
        }

        private void tag4EditMenu_Click(object sender, EventArgs e)
        {
            CreateTagEditBox4(lbxTag4.GetItemRectangle(lbxTag4.SelectedIndex), lbxTag4.SelectedItem.ToString(), lbxTag4);
        }
        //edit and remove for complete listbox
        private void completeRemoveMenu_Click(object sender, EventArgs e)
        {
            lbxComplete.Items.RemoveAt(lbxComplete.SelectedIndex);
        }

        private void completeEditMenu_Click(object sender, EventArgs e)
        {
            CreateCompleteEditBox4(lbxComplete.GetItemRectangle(lbxComplete.SelectedIndex), lbxComplete.SelectedItem.ToString(), lbxComplete);
        }

        //de-selecting tag listbox items
        private void linkLabel7_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lbxTag4.SelectedIndex = -1;
        }

        private void linkLabel6_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lbxTag3.SelectedIndex = -1;
        }



        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lbxTag1.SelectedIndex = -1;
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lbxTag2.SelectedIndex = -1;
        }

        //selecting an item in any of the tag list box automatically de-select items in scope and category list box
        private void lbxTag1_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbxCategories.SelectedIndex = -1;
            lbxScopes.SelectedIndex = -1;

        }
        private void lbxTag2_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbxCategories.SelectedIndex = -1;
            lbxScopes.SelectedIndex = -1;

        }
        private void lbxTag3_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbxCategories.SelectedIndex = -1;
            lbxScopes.SelectedIndex = -1;

        }

        private void lbxTag4_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbxCategories.SelectedIndex = -1;
            lbxScopes.SelectedIndex = -1;

        }


        private void btnComplete_Click(object sender, EventArgs e)
        {
            completeItem = new frmAddCompleteItem();
            completeItem.ShowDialog();
            if (clickSave)
            {
                clickSave = false;
                //db = collection.AddCompleteItem(completeItemText);
                lbxComplete.Items.Add(completeItemText);
            }
        }

        private void lbxComplete_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbxScopes.SelectedIndex = -1;
            lbxCategories.SelectedIndex = -1;
            //if item is selected in complete list box send caption string to photo app
            if (lbxComplete.SelectedIndex > -1)
            {
                string captionStr = "";
                //append selected item  in tag 1 listbox to caption string
                if (lbxTag1.SelectedIndex > -1) { captionStr = captionStr + lbxTag1.SelectedItem.ToString() + " "; }
                //append selected item in tag 1 listbox to caption string
                if (lbxTag2.SelectedIndex > -1) { captionStr = captionStr + lbxTag2.SelectedItem.ToString() + " "; }
                //append selected item in tag 3 listbox to caption string
                if (lbxTag3.SelectedIndex > -1) { captionStr = captionStr + lbxTag3.SelectedItem.ToString() + " "; }
                //append selected item in tag 4 listbox to caption string
                if (lbxTag4.SelectedIndex > -1) { captionStr = captionStr + lbxTag4.SelectedItem.ToString() + " "; }
                //append item in complete listbox to caption string
                captionStr = captionStr + lbxComplete.SelectedItem.ToString()+" ";

                //get previous active window
                IntPtr lastWindowHandle = GetWindow(Process.GetCurrentProcess().MainWindowHandle, (uint)GetWindow_Cmd.GW_HWNDNEXT);
                while (true)
                {
                    IntPtr temp = GetParent(lastWindowHandle);
                    if (temp.Equals(IntPtr.Zero)) break;
                    lastWindowHandle = temp;
                }
                SetForegroundWindow(lastWindowHandle);
                //send string to photo app
                SendKeys.SendWait(captionStr);
                //Next photo command
                SendKeys.SendWait("%{n}");
                //get caption builder window back
               // SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);



            }
        }

        private void lbxComplete_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                btnComplete.PerformClick();
            }
        }
        #endregion

        private void btnCatToScope_Click(object sender, EventArgs e)
        {
            catToScope = new frmAddCatToScope(lbxCategories.Items);
            catToScope.ShowDialog();
            if (clickSave)
            {
                clickSave = false;
                foreach (var caption in db.Captions.ToList())
                {
                    if (caption.category.name == category)
                    {
                        db = collection.AddCaptionToScope(caption.name, caption.category.name, lbxScopes.SelectedItem.ToString());
                    }
                }
                if (lbxCategories.SelectedIndex == -1 && lbxScopes.SelectedIndex != -1)
                {
                    loadGrid(db, "Scope", String.Empty, lbxScopes.SelectedItem.ToString());
                }
                else if (lbxCategories.SelectedIndex > -1 && lbxScopes.SelectedIndex > -1)
                {
                    loadGrid(db, "Mix", lbxCategories.SelectedItem.ToString(), lbxScopes.SelectedItem.ToString());
                }
                else if (lbxCategories.SelectedIndex == -1 && lbxScopes.SelectedIndex == -1)
                {
                    loadGrid(db, "All", String.Empty, String.Empty);
                }
                else if (lbxCategories.SelectedIndex != -1 && lbxScopes.SelectedIndex == -1)
                {
                    loadGrid(db, "Category", lbxCategories.SelectedItem.ToString(), String.Empty);
                }

            }
        }

        #region save from front end //This is to keep the user's arrangement of items on the UI
        private void SaveDataFromUI()
        {

        }

        private void saveScope()
        {
            try
            {


                List<Scope> scopes = new List<Scope>();
                foreach (var scope in lbxScopes.Items)
                {
                    scopes.Add(new Scope { name = scope.ToString() });
                }
                db.Scopes = scopes;
            }
            catch (Exception) { return; }
        }

        private void saveCategory()
        {
            try
            {
                categoriesInUserOrder = new List<CaptionBuilder.Category>();
                foreach (var cat in lbxCategories.Items)
                {
                    categoriesInUserOrder.Add(new Category { name = cat.ToString() });
                }
                //reset categories
                db.Categories = categoriesInUserOrder;
            }
            catch (Exception) { return; }
           
        }
        //save caption order according to user arrangement
        List<Category> categoriesInUserOrder;
        private void saveCaptionsAndCategoriesOrder()
        {
            
            captionsArrangement = new List<Caption>();
            string currentCategory = "";
            //when all caption string is being saved 
            if (dataGridView1.RowCount > 0)
            {

                if (lbxCategories.SelectedIndex == -1 && lbxScopes.SelectedIndex == -1)
                {
                    //for every row
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        if (dataGridView1.Rows[i].Cells[0].Value != null)
                        {


                            //if row is category title
                            if (dataGridView1.Rows[i].Cells[0].GetType().Name.ToString() == "HMergedCell")
                            {
                                currentCategory = dataGridView1.Rows[i].Cells[0].Value.ToString();

                            }
                            //if row is a caption string
                            else
                            {
                                //store extensions in temporary extenx variable
                                string[] extenx = new string[4] { null, null, null, null };
                                //stores the next photo toggle status of the caption 
                                bool nextPhoto = true;
                                //when toggle field is red
                                if (dataGridView1.Rows[i].Cells[5].Value == null) { nextPhoto = false; }
                                //set extensions according to data in columns 2,3,4,5
                                if (dataGridView1.Rows[i].Cells[1].Value != null) { extenx[0] = dataGridView1.Rows[i].Cells[1].Value.ToString(); }
                                if (dataGridView1.Rows[i].Cells[2].Value != null) { extenx[1] = dataGridView1.Rows[i].Cells[2].Value.ToString(); }
                                if (dataGridView1.Rows[i].Cells[3].Value != null) { extenx[2] = dataGridView1.Rows[i].Cells[3].Value.ToString(); }
                                if (dataGridView1.Rows[i].Cells[4].Value != null) { extenx[3] = dataGridView1.Rows[i].Cells[4].Value.ToString(); }
                                //create new caption
                                captionsArrangement.Add(new Caption
                                {

                                    name = dataGridView1.Rows[i].Cells[0].Value.ToString(),
                                    category = new Category { name = currentCategory },
                                    extensions = extenx,
                                    scopes = (List<string>)dataGridView1.Rows[i].Cells[7].Value,
                                    goToNextImage = nextPhoto
                                });

                            }

                        }
                    }
                    //reset captions in database
                    db.Captions = captionsArrangement;


                }
                //when scope data is rearranged and is being saved
                else if (lbxScopes.SelectedIndex > -1)
                {
                    //stores the index of the most current category title
                    int currentCategoryIndex = 0;
                    //for every row of caption under selected scope
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {

                        //if row is category title
                        if (dataGridView1.Rows[i].Cells[0].GetType().Name == "HMergedCell")
                        {
                            currentCategory = dataGridView1.Rows[i].Cells[0].Value.ToString();
                            //categoriesInUserOrder.Add(new Category { name = currentCategory });
                        }
                        //if the row is a caption row
                        else
                        {
                            //look for matching caption in db and rearrange
                            for (int j = 0; j < db.Captions.Count; j++)
                            {
                                //if caption is found in db
                                if (dataGridView1.Rows[i].Cells[0].Value.ToString() == db.Captions[j].name)
                                {
                                    //store current caption
                                    Caption currentCaption = db.Captions[j];
                                    currentCaption.category.name = currentCategory;
                                    //remove caption
                                    db.Captions.Remove(db.Captions[j]);
                                    //re-add caption to sort in datagrid order
                                    db.Captions.Add(currentCaption);
                                    break;
                                }
                            }
                        }
                    }
                }
                //when captions under category is being saved
                else if (lbxCategories.SelectedIndex > -1)
                {
                    //for every row of caption under selected scope
                    for (int i = 1; i < dataGridView1.RowCount; i++)
                    {
                        //look for matching caption in db and rearrange
                        for (int j = 0; j < db.Captions.Count; j++)
                        {
                            //when caption is found in db
                            if (dataGridView1.Rows[i].Cells[0].Value.ToString() == db.Captions[j].name && dataGridView1.Rows[0].Cells[0].Value.ToString() == db.Captions[j].category.name)
                            {
                                //store  caption
                                Caption currentCaption = db.Captions[j];

                                //remove caption
                                db.Captions.Remove(db.Captions[j]);
                                //re-add caption to sort in datagrid order
                                db.Captions.Add(currentCaption);
                                break;
                            }
                        }
                    }
                }
            }

            //collection.SaveData(db);
        }

        //save complete
        List<Complete> saveCompleteItemList;
        private void saveComplete()
        {
            try
            {
                saveCompleteItemList = new List<Complete>();
                foreach (var item in lbxComplete.Items)
                {
                    saveCompleteItemList.Add(new Complete() { text = item.ToString() });
                }
                db.CompleteItems = saveCompleteItemList;
            }
            catch (Exception) { return; }
        }

        //save tags
        List<Tag> saveTagList;
        private void saveTags()
        {
            try
            {
                saveTagList = new List<CaptionBuilder.Tag>();
                //save items in tag1 listbox
                //if listbox is not empty
                if (lbxTag1.Items.Count > 0)
                {
                    foreach (var item in lbxTag1.Items)
                    {
                        saveTagList.Add(new CaptionBuilder.Tag() { description = item.ToString(), tagNumber = 1 });
                    }
                }
                //save items in tag2 listbox
                //if listbox is not empty
                if (lbxTag2.Items.Count > 0)
                {
                    foreach (var item in lbxTag2.Items)
                    {
                        saveTagList.Add(new Tag() { description = item.ToString(), tagNumber = 2 });
                    }
                }
                //save items in tag3 listbox
                //if listbox is not empty
                if (lbxTag3.Items.Count > 0)
                {
                    foreach (var item in lbxTag3.Items)
                    {
                        saveTagList.Add(new CaptionBuilder.Tag() { description = item.ToString(), tagNumber = 3 });
                    }
                }
                //save items in tag1 listbox
                //if listbox is not empty
                if (lbxTag4.Items.Count > 0)
                {
                    foreach (var item in lbxTag4.Items)
                    {
                        saveTagList.Add(new CaptionBuilder.Tag() { description = item.ToString(), tagNumber = 4 });
                    }
                }
                db.Tags = saveTagList;
            }
            catch (Exception) { return; }
        }

        #endregion

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter && dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].GetType().Name!="HMergedCell")
            {
                editGrid();            

            }
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dataGridView1.Focus();
        }

        private void dataGridView1_TabIndexChanged(object sender, EventArgs e)
        {

        }
        public static bool endNextEdit{get;set;}

        //auto forward datagrid cell during caption text entering
        private void dataGridView1_KeyUp(object sender, KeyEventArgs e)
        {
            //if tab key is pressed
            if (e.KeyData == Keys.Tab)
            { int x = 0;
                
                while (x < 4)
                {
                    if (!endNextEdit)
                    {
                        //open edit box
                        editGrid();
                        if (dataGridView1.CurrentCell.ColumnIndex < 4)
                        {
                            dataGridView1.CurrentCell = dataGridView1[dataGridView1.CurrentCell.ColumnIndex + 1, dataGridView1.CurrentCell.RowIndex];
                        }
                        x += 1;
                        
                    }
                    else { break; }
                   
                }
                if (x==4 && btnCreateCaption.Enabled == true)
                {
                    btnCreateCaption.PerformClick();
                }
                           
            }
        }

        private void btnDbPath_Click(object sender, EventArgs e)
        {
            setDbPath = new frmDbPath();
            setDbPath.ShowDialog();

        }

        private void btnReload_Click(object sender, EventArgs e)
        {
            Form1_Load(sender,e);
        }
    }


    public class DragItem
    {
        public ListBox Client;
        public int Index;
        public object Item;

        public DragItem(ListBox client, int index, object item)
        {
            Client = client;
            Index = index;
            Item = item;
        }

    }

    static class MyExtension
    {
        // Return the index of the item that is
        // under the point in screen coordinates.
        public static int IndexFromScreenPoint(this ListBox lst, Point point)
        {
            // Convert the location to the ListBox's coordinates.
            point = lst.PointToClient(point);

            // Return the index of the item at that position.
            return lst.IndexFromPoint(point);
        }
    }

}

