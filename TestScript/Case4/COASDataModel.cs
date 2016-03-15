using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Young.Data.Attributes;

namespace TestScript.Case4
{
    class COASDataModel
    {
        [ColMapping("Order")]
        public string InternalOrder { get; set; }

        [ColMapping("Resp. CCtr")]
        public string CostCenter { get; set; }

        [ColMapping("Log.System")]
        public string LogicalSystem { get; set; }
    }
}
