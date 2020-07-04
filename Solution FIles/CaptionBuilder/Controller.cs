using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.IO;
using System.Windows.Forms;


namespace CaptionBuilder
{
     class Controller
    {
        Model  db;
        string jsonData;
        string path = "CaptionDatabase.json";
        string dbPath = "";
        readonly string dbName = "CaptionDatabase.json";
        string fullPath = "";
        string dblocation = "";




        public Controller()
        {
            try
            {
                //initiaze root path of database (get My documents path of host machine)
                string rootPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string targetDatabaseLocation = "DatabaseLocation.txt";
                //directory path for db
                string dbPath = $@"{rootPath}\Caption Builder\database";
                //if db folder does not exist, create 
                if (!Directory.Exists(dbPath))
                {                    
                    
                    Model emptyModel = new Model();
                    Directory.CreateDirectory(dbPath);
                    jsonData = JsonConvert.SerializeObject(emptyModel, Formatting.Indented);
                    using (StreamWriter writer = new StreamWriter(dbPath+"/"+dbName))
                    {
                        writer.Write(jsonData);
                    }
                   

                }
                if (!File.Exists(targetDatabaseLocation))
                {
                    using (StreamWriter writer = new StreamWriter(targetDatabaseLocation))
                    {
                        writer.Write(dbPath);
                    }
                  
                }
                //get location of targeted db
                using (var reader = new StreamReader(targetDatabaseLocation))
                {
                    dblocation = reader.ReadToEnd();
                }

                //read data from database
                using (var reader = new StreamReader($@"{dblocation}\{dbName}"))
                {
                    jsonData = reader.ReadToEnd();
                }

                db = JsonConvert.DeserializeObject<Model>(jsonData);


            }
            catch (Exception ex)
            {
                return;
            }
        }

        public Model GetData()
        {

            return db;
        }

        public void SaveData(Model model)
        {
            try
            {
                //save data to db
                jsonData = JsonConvert.SerializeObject(model,Formatting.Indented);
                using (StreamWriter writer = new StreamWriter($@"{dblocation}\{dbName}"))
                {
                    writer.Write(jsonData);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured during saving, try again");
            }
        }

        public Model NewScope(string name)
        {
            Scope scope = new Scope
            {
                name = name
                
            };
            if (db.Scopes == null) { db.Scopes = new List<Scope>(); }
            db.Scopes.Add(scope);

            return db;
        }

        public Model RemoveScope(string name)
        {
            if (db.Captions != null)
            {


                foreach (var addedCaption in db.Captions.ToList())
                {
                    if (addedCaption.scopes.Count == 1 && addedCaption.scopes.Contains(name))
                    {


                        db.Captions.Remove(addedCaption);
                    }
                    else if (addedCaption.scopes.Count > 1 && addedCaption.scopes.Contains(name))
                    {
                        addedCaption.scopes.Remove(name);
                    }
                }

            }

            foreach (var scope in db.Scopes.ToList())
            {
                if (scope.name == name)
                {
                    db.Scopes.Remove(scope);
                }
            }

            return db;
        }

        public Model EditScope(string oldName, string newName)
        {
            if (db.Captions != null)
            {


                foreach (var addedCaption in db.Captions.ToList())
                {
                    if (addedCaption.scopes.Contains(oldName))
                    {
                        addedCaption.scopes[addedCaption.scopes.IndexOf(oldName)] = newName;
                    }
                }
            }
            foreach (var scope in db.Scopes.ToList())
            {
                if (scope.name == oldName)
                {
                    scope.name = newName;
                }
            }


            return db;
        }

        public Model NewCategory(string categoryName)
        {
            Category category = new Category
            {
                name = categoryName,
                
            };
            if (db.Categories == null) { db.Categories = new List<Category>(); }
            db.Categories.Add(category);

            return db;

        }

        public Model RemoveCategory(string categoryName)
        {
            if (db.Captions != null)
            {


                foreach (var addedCaption in db.Captions.ToList())
                {
                    if (addedCaption.category.name == categoryName)
                    {
                        db.Captions.Remove(addedCaption);
                    }
                }
            }
            foreach (var category in db.Categories.ToList())
            {
                if (category.name == categoryName)
                {
                    db.Categories.Remove(category);
                }
            }
            return db;
        }

        public Model EditCategory(string oldName,string newName)
        {
            if (db.Captions != null)
            {


                foreach (var addedCaption in db.Captions.ToList())
                {
                    if (addedCaption.category.name == oldName)
                    {
                        addedCaption.category.name = newName;
                    }
                }
            }
            foreach (var category in db.Categories.ToList())
            {
                if (category.name == oldName)
                {
                    category.name=newName;
                }
            }
            return db;
        }

        public Model CreateCaption(string captionName, Category category)
        {
            Caption caption = new Caption
            {
                name = captionName,
                category = category,
                extensions = new string[4],
                scopes = new List<string>(),
                goToNextImage = true
                

            };
            if (db.Captions == null) { db.Captions = new List<Caption>(); }
            db.Captions.Add(caption);

            return db;
        }

        public Model AddCaptionToScope(string name,string category,string scope)
        {
            if (db.Captions.Count > 0)
            {
                foreach (var cap in db.Captions.ToList())
                {
                    if (cap.name == name && cap.scopes.Contains(scope)==false)
                    {
                        cap.scopes.Add(scope);
                        
                    }
                }

            }
            
            return db;
        }

        public Model RemoveCaptionFromScope(string caption,string category, string scope)
        {
            foreach (var cap in db.Captions.ToList())
            {
                if (cap.name == caption && cap.category.name ==category && cap.scopes.Contains(scope))
                {
                   
                        cap.scopes.Remove(scope);
                  
                }
            }

            return db;

        }

        public Model RemoveCaption(string caption,string category)
        {
            foreach (var cap in db.Captions.ToList())
            {
                if (cap.name == caption)
                {
                    db.Captions.Remove(cap);
                }
            }


            return db;

        }


        public Model EditCaption(string captionString, string category, string newCaptionString)
        {
            foreach (var cap in db.Captions.ToList())
            {
                if (cap.name == captionString && cap.category.name == category)
                {
                    cap.name = newCaptionString;
                }
            }


            return db;
        }

        public Model EditCaptionNextImageStatus(string captionString, string category, bool status)
        {
            foreach (var cap in db.Captions.ToList())
            {
                if (cap.name == captionString && cap.category.name == category)
                {
                    cap.goToNextImage = status;
                }
            }


            return db;
        }

        public Model AddExtensionToCaption(string caption,string category,string extensionStr,int index)
        {
            foreach(var cap in db.Captions.ToList())
            {
                if(cap.name==caption && cap.category.name==category)
                {
                    cap.extensions[index]=extensionStr;
                }
            }
            return db;
        }


        public Model RemoveExtensionToCaption(string caption, string category, string extensionStr,int index)
        {
            foreach (var cap in db.Captions.ToList())
            {
                if (cap.name == caption && cap.category.name == category && cap.extensions.Contains(extensionStr))
                {
                   // cap.extensions.Remove(extensionStr);
                }
            }
            return db;
        }


        //add a tag
        public Model AddTag(string description , int number)
        {
            Tag tag = new Tag
            {
                description = description,
                tagNumber = number
            };
            if (db.Tags == null) { db.Tags = new List<Tag>(); }
            db.Tags.Add(tag);

            return db;
        }

        public Model EditTag(string oldDesc, string newDesc, int tagNumber)
        {
            if (db.Tags != null)
            {


                foreach (var tag in db.Tags.ToList())
                {
                    if (tag.description == oldDesc && tag.tagNumber == tagNumber)
                    {
                        tag.description = newDesc;
                        break;
                    }
                }
            }
            return db;
        }

        //remove tag
        public Model RemoveTag(string description,int tagNumber)
        {
            if (db.Tags != null)
            {


                foreach (var tag in db.Tags.ToList())
                {
                    if (tag.description == description && tag.tagNumber == tagNumber)
                    {
                        db.Tags.Remove(tag);
                        break;
                    }
                }
            }
            return db;
        }

        //Add  complete item
        public Model AddCompleteItem(string text)
        {
            Complete completeItem = new Complete
            {
               text=text
            };
            if (db.CompleteItems == null) { db.CompleteItems = new List<Complete>(); }
            
            db.CompleteItems.Add(completeItem);

            return db;
        }

        //edit complete item
        public Model EditCompleteItem(string oldText,string newText)
        {
            if (db.CompleteItems != null)
            {


                foreach (var item in db.CompleteItems.ToList())
                {
                    if (item.text == oldText)
                    {
                        item.text = newText;
                    }
                }
            }
            return db;
        }

        //remove complete item
        public Model RemoveCompleteItem(string text)
        {
            if (db.CompleteItems != null)
            {
                foreach (var item in db.CompleteItems.ToList())
                {
                    if (item.text == text)
                    {
                        db.CompleteItems.Remove(item);
                    }
                }
            }

            return db;
        }
    }
}
