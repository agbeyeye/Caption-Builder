using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaptionBuilder
{
    public class Model
    {
        public List<Scope> Scopes { get; set; }
        public List<Category> Categories { get; set; }

        public List<Caption> Captions { get; set; }

        public List<Tag> Tags { get; set; }

        public List<Complete> CompleteItems { get; set; }

    }

    public class Scope
    {
        public string name { get; set; }
    }

    public class Category
    {
        public string name { get; set; }
    }

    public class Caption
    {
        public string name { get; set; }
        public Category category { get; set; }
        public string[] extensions { get; set; }
        public List<string> scopes { get; set; }
        public bool goToNextImage { get; set; }
    }

    public class Tag
    {
        public string description { get; set; }
        public int tagNumber { get; set; }
    }

    public class Complete
    {
        public string text { get; set; }
    }
}
