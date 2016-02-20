using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestScript.Case1
{
    public class Case1ReportDataModel
    {
        [Alias("LOGICAL SYSTEM")]
        public string LogicalSystem { get; set; }

        [Alias("COMPANY CODE")]
        public string CompanyCode { get; set; }

        [Alias("Account")]
        public string Account { get; set; }

        [Alias("DOCUMENT NUMBER")]
        public string DocumentNumber { get; set; }

        [Alias("DOC TYPE")]
        public string DocType { get; set; }

        [Alias("ENTRY DATE")]
        public string EntryDate { get; set; }

        [Alias("POSTING DATE")]
        public string PostingDate { get; set; }

        [Alias("DOCUMENT DATE")]
        public string DocumentDate { get; set; }

        [Alias("TRANSLATION DATE")]
        public string TranslationDate { get; set; }

        [Alias("USER NAME")]
        public string UserName { get; set; }

        [Alias("DOC CURRENCY")]
        public string DocCurrency { get; set; }

        [Alias("Amount in Document Currency")]
        public float AmtInDocCur { get; set; }

        [Alias("LOCAL CURRENCY")]
        public string LocalCur { get; set; }

        [Alias("Amount in Local Currency")]
        public float AmtInlocalCur { get; set; }

        [Alias("GROUP CURRENCY")]
        public string GroupCur { get; set; }

        [Alias("Amount in Group Currency")]
        public float AmtInGroupCur { get; set; }

        [Alias("TC/GC")]
        public float TC_GC
        {
            get
            {
                if (AmtInGroupCur != 0)
                    return AmtInDocCur / AmtInGroupCur;
                return 0;
            }
        }

        [Alias("LC/GC")]
        public float LC_GC
        {
            get
            {
                if (AmtInGroupCur != 0)
                    return AmtInlocalCur / AmtInGroupCur;
                return 0;
            }
        }

        [Alias("OB08 Ex Rate TC/GC")]
        public float OB08ExTC_GC { get; set; }

        [Alias("OB08 Ex Rate LC/GC")]
        public float OB08ExLC_GC { get; set; }

        [Alias("TC/GC delta")]
        public float TC_GC_Delta { get; set; }

        [Alias("LC/GC delta")]
        public float LC_GC_Delta { get; set; }

        [Alias("Delta LC/GC")]
        public float Delta_LC_GC
        {
            get
            {
                if (Rateof_LC_USD != 0)
                    return AmtInlocalCur / Rateof_LC_USD - AmtInGroupCur;
                return 0;
            }
        }

        [Alias("Delta GC/LC In Loc")]
        public float Delta_GC_LC
        {
            get
            {
                return AmtInGroupCur * Rateof_LC_USD - AmtInlocalCur;
            }
        }

        [Alias("Delta TC/LC/GC")]
        public string Delta_TC_LC_GC
        {
            get
            {
                if (DocCurrency.ToLower() == LocalCur.ToLower())
                    return (AmtInDocCur - AmtInlocalCur).ToString();
                if (DocCurrency.ToLower() == "usd")
                    return "Amt in Document Currency – Amt in USD";
                if (Rateof_DC_USD == 0)
                    return "0";
                return (AmtInDocCur / Rateof_DC_USD - AmtInGroupCur).ToString();
            }
        }

        public float Rateof_LC_USD { get; set; }

        public float Rateof_DC_USD { get; set; }
        ///Curr. Doc Cur
        ///LCurr Local Cur
        ///Lcur2 Group Cur
    }
}
