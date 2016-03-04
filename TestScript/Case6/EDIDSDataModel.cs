using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Young.Data.Attributes;

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

        [ColMapping("IDoc number")]
        public string IDocNumber { get; set; }

        [ColMapping("Date status error")]
        public string DataStatusError { get; set; }

        [ColMapping("Time status error")]
        public string TimeStatusError { get; set; }

        [ColMapping("Status counter")]
        public string StatusCounter { get; set; }

        [ColMapping("Status")]
        public string Status { get; set; }

        [ColMapping("Name of person to be notified")]
        public string UserID { get; set; }

        public string LastDateOfStatus69 { get; set; }

        public string LastDateOfStatus51 { get; set; }
    }
}
