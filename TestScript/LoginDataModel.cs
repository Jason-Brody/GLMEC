using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestScript
{
    public class LoginDataModel
    {
        [Required]
        public string UserName { get; set; }

        [Required]
        public string Password { get; set; }

        [Required]
        public string Address { get; set; }

        [Required]
        public string Language { get; set; }

        [Required]
        public string Client { get; set; }


        public string MaximumNo { get; set; } = "2147483646";
    }
}
