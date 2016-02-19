using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestScript.Case1
{
    public class AccountModel
    {
        [Alias("G/L Acct")]
        public string Account { get; set; }

        public bool IsComplete { get; set; }

        public TimeSpan Period { get; set; }

        public int DataCount { get; set; }
    }
}
