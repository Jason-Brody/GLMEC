using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Young.Data.Attributes;

namespace TestScript.Case4
{
    class CSKSDataModel
    {
        [ColMapping("Valid From")]
        public string ValidFrom { get; set; }

        [ColMapping("to")]
        public string ValidTo { get; set; }

        [ColMapping("Cost Ctr")]
        public string CostCenter { get; set; }

        [ColMapping("Log.System")]
        public string LogicalSystem { get; set; }
    }
}
