using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaseRunnerModel.Attributes
{
    [AttributeUsage(AttributeTargets.Class,AllowMultiple =false)]
    public class ScriptAttribute:Attribute
    {
        public ScriptAttribute(string Name) : this(Name, null, null) { }

        public ScriptAttribute(string Name, string Description) : this(Name, Description,null) { }

        public ScriptAttribute(string Name,string Description,string HelpLink)
        {
            this.Name = Name;
            this.Description = Description;
            this.HelpLink = HelpLink;
        }

        public string Name { get; set; }

        public string Description { get; set; }

        public string HelpLink { get; set; }
    }
}
