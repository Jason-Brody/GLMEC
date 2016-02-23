using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestScript.Case1
{
    public class Case1DataModel:LoginDataModel
    {
        public string GLAccountFrom { get; set; }

        public string GLAccountTo { get; set; }

        public string CompanyCode { get; set; }

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
