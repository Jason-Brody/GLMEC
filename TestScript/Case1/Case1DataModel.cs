using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;

namespace TestScript.Case1
{
    public class Case1DataModel:LoginDataModel
    {
        public string GLAccountFrom { get; set; }

        public string GLAccountTo { get; set; }

        [Required]
        [Display(Name = "CompanyCode")]
        public string CompanyCode { get; set; }

        [Required]
        [Display(Name = "PostingDateFrom")]
        public string PostingDateFrom
        {
            get
            {
                return _postingStart.ToString("dd.MM.yyyy");
            }
            set
            {
                _postingStart = setDate(value);
            }
        }

        [Required]
        [Display(Name = "PostingDateTo")]
        public string PostingDateTo
        {
            get
            {
                return _postingEnd.ToString("dd.MM.yyyy");
            }
            set
            {
                _postingEnd = setDate(value);
            }
        }

        [Required]
        [Display(Name ="Layout")]
        public string Layout { get; set; }

        private DateTime _postingStart;
        private DateTime _postingEnd;

        public DateTime PostingStartDate { get { return _postingStart; } }

        public DateTime PostingEndDate { get { return _postingEnd; } }

        private DateTime setDate(string date)
        {
            var dataList = date.Split('.');
            var dd = int.Parse(dataList[0]);
            var MM = int.Parse(dataList[1]);
            var yyyy = int.Parse(dataList[2]);
            return new DateTime(yyyy, MM, dd);
        }

    }
}
