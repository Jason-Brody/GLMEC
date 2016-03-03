﻿using System;
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
        [Display(Name = "SAP User Name")]
        public string UserName { get; set; }

        [Required]
        [Display(Name = "SAP User Password")]
        public string Password { get; set; }

        [Required]
        [Display(Name = "SAP Server Address")]
        public string Address { get; set; }

        [Required]
        [Display(Name = "SAP Server Language")]
        public string Language { get; set; }

        [Required]
        [Display(Name = "SAP Server Client")]
        public string Client { get; set; }
    }
}
