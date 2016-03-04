using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Young.Data.Attributes;

namespace TestScript.Case6
{
    public class HRP1001DataModel
    {
        [ColMapping("ObjectID")]
        public string ObjectId { get; set; }

        [ColMapping("User name")]
        public string UserName { get; set; }
    }
}
