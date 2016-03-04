using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Young.Data.Attributes;

namespace TestScript.Case6
{
    class EDP21DataModel
    {
        [ColMapping("Message code")]
        public string MessageCode { get; set; }

        public string Partner { get; set; }

        [ColMapping("Message function")]
        public string MessageFunction { get; set; }

        [ColMapping("Message Type")]
        public string MessageType { get; set; }

        public string Agent { get; set; }
    }
}
