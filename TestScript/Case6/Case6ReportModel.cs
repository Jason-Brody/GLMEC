using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Young.Excel.Interop.Attributes;

namespace TestScript.Case6
{
    class Case6ReportModel
    {
        [ExcelHeaderStyle(BackgroundColor = 14336204, IsFontBold = true)]
        [Display(Name ="Idoc No")]
        public string IDocNumber { get; set; }

        [ExcelHeaderStyle(BackgroundColor = 14336204, IsFontBold = true)]
        [Display(Name = "Creation date")]
        public string CreateDate { get; set; }

        [ExcelHeaderStyle(BackgroundColor = 14336204, IsFontBold = true)]
        [Display(Name = "Current Status")]
        public string CurrentStatus { get; set; }

        [ExcelHeaderStyle(BackgroundColor = 65535, IsFontBold = true)]
        [Display(Name = "Last Date of Status 69")]
        public string LastDateOfStatus69 { get; set; }

        [ExcelHeaderStyle(BackgroundColor = 65535, IsFontBold = true)]
        [Display(Name = "Last Date of Status 51")]
        public string LastDateOfStatus51 { get; set; }

        [ExcelHeaderStyle(BackgroundColor = 14336204, IsFontBold = true)]
        [Display(Name = "Message Type")]
        public string MessageType { get; set; }

        [ExcelHeaderStyle(BackgroundColor = 14336204,IsFontBold =true)]
        [Display(Name = "Message Variant")]
        public string MessageVariant { get; set; }

        [ExcelHeaderStyle(BackgroundColor = 14336204, IsFontBold = true)]
        [Display(Name = "Message function")]
        public string MessageFunction { get; set; }

        [ExcelHeaderStyle(BackgroundColor = 5296274, IsFontBold = true)]
        [Display(Name = "Partner number")]
        public string PartnerNumber { get; set; }

        [ExcelHeaderStyle(BackgroundColor = 5296274, IsFontBold = true)]
        [Display(Name = "Agent number")]
        public string AgentNumber { get; set; }

        [ExcelHeaderStyle(BackgroundColor = 14336204, IsFontBold = true)]
        [Display(Name = "User ID")]
        public string UserId { get; set; }

        [ExcelHeaderStyle(BackgroundColor = 9420794, IsFontBold = true)]
        [Display(Name = "Existing in HRP1001?")]
        public string IsExisted { get; set; }
    }
}
