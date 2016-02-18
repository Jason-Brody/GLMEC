using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestScript.Case1
{
    class Case1DataModel:LoginDataModel
    {
        public string GLAccountFrom { get; set; }

        public string GLAccountTo { get; set; }

        public string CompanyCode { get; set; }

        public string PostingDateFrom { get; set; }

        public string PostingDateTo { get; set; }

        public string Layout { get; set; }
    }
}
