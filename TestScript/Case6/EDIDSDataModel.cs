using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Young.Data.Attributes;
using Young.Excel.Interop.Attributes;

namespace TestScript.Case6
{
    class EDIDSDataModel
    {
        public EDIDSDataModel() { }

        public EDIDSDataModel(EDIDSDataModel data)
        {
            this.IDocNumber = data.IDocNumber;
            this.DataStatusError = data.DataStatusError;
            this.TimeStatusError = data.TimeStatusError;
            this.StatusCounter = data.StatusCounter;
            this.Status = data.Status;
            this.UserID = data.UserID;
            this.LastDateOfStatus51 = data.LastDateOfStatus51;
            this.LastDateOfStatus69 = data.LastDateOfStatus69;
        }

        [ExcelHeaderStyle(BackgroundColor = 14857357)]
        [ColMapping("IDoc number")]
        [Display(Name = "IDoc number")]
        public string IDocNumber { get; set; }

        [ExcelHeaderStyle(BackgroundColor = 14857357)]
        [ColMapping("Date status error")]
        [Display(Name = "Date status error")]
        public string DataStatusError { get; set; }

        [ExcelHeaderStyle(BackgroundColor = 14857357,NumberFormat = "h: mm:ss;@")]
        [ColMapping("Time status error")]
        [Display(Name = "Time status error")]
        public string TimeStatusError { get; set; }

        [ExcelHeaderStyle(BackgroundColor = 14857357)]
        [ColMapping("Status counter")]
        [Display(Name = "Status counter")]
        public string StatusCounter { get; set; }

        [ExcelHeaderStyle(BackgroundColor = 14857357)]
        [ColMapping("Status")]
        public string Status { get; set; }

        [ExcelHeaderStyle(BackgroundColor = 14857357)]
        [ColMapping("Name of person to be notified")]
        [Display(Name = "Name of person to be notified")]
        public string UserID { get; set; }
        
        [Ignore]
        public string LastDateOfStatus69 { get; set; }

        [Ignore]
        public string LastDateOfStatus51 { get; set; }
    }
}
