﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestScript.Case1
{
    public class Case1DocInfoModel
    {
        [Alias("Cocd")]
        public string CompanyCode { get; set; }

        [Alias("DocumentNo")]
        public string DocumentNumber { get; set; }

        public string DocType { get; set; }

        public string DocDate { get; set; }

        public string PostingDate { get; set; }

        public string EntryDate { get; set; }

        public string TransDate { get; set; }

        public string UserName { get; set; }

        public string DocCurrency { get; set; }

        public string LocalCurrency { get; set; }

        [Alias("LCur2")]
        public string GroupCurrency { get; set; }

        public string ExchangeRateType { get; set; }

        public string LogicalSystem { get; set; }

        public string RefKey1 { get; set; }
    }
}
