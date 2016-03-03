using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPShareDataLib;
using System.ComponentModel.DataAnnotations;

namespace GL_MEC_006
{
    public class DataModel:LoginDataModel
    {
        [Required]
        public string IDocNumber { get; set; }
    }
}
