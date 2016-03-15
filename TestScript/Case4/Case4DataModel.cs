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

        [Required]
        public string PostingDateFrom { get; set; }

        [Required]
        public string PostingDateTo { get; set; }
    }
}
