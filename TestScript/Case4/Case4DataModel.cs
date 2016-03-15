using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestScript.Case4
{
    public class Case4DataModel:LoginDataModel
    {
        [Required]
        public string CompanyCode { get; set; }

        [Required]
        public string GLAccountFilePath { get; set; }

        public string PostingDateFrom { get; set; }

        public string PostingDateTo { get; set; }
    }
}
