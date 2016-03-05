using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Young.Data.Attributes;

namespace TestScript.Case6
{
    public class EDIDCDataModel
    {
        [ColMapping("IDoc number")]
        public string IDocNumber { get; set; }

        [ColMapping("Status")]
        public string Status { get; set; }

        [ColMapping("Message Variant")]
        public string MessageVariant { get; set; }

        [ColMapping("Message function")]
        public string MessageFunction { get; set; }

        [ColMapping("Created on")]
        public string CreateDt { get; set; }

        [ColMapping("Message Type")]
        public string MessageType { get; set; }
    }
}
