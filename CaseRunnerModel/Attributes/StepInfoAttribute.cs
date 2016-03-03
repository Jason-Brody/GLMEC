using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaseRunnerModel.Attributes
{
    [AttributeUsage(AttributeTargets.Method,AllowMultiple =false)]
    public class StepAttribute:Attribute
    {
        public int Order { get; set; }

        public string Name { get; set; }

        public string Description { get; set; }
    }
}
