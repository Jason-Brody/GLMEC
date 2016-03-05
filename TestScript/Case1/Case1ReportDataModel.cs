using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Young.Data.Attributes;
using Young.Excel.Interop.Attributes;
using System.ComponentModel.DataAnnotations;

namespace TestScript.Case1
{
    public class Case1ReportDataModel
    {
        
        [ExcelHeaderStyle(65535, 10.85)]
        [Display(Name = "LOGICAL SYSTEM")]
        public string LogicalSystem { get; set; }

        [ExcelHeaderStyle(14857357, 11.70)]
        [Display(Name = "COMPANY CODE")]
        public string CompanyCode { get; set; }

        [ExcelHeaderStyle(14857357, 9)]
        [Display(Name = "Account")]
        public string Account { get; set; }

        [ExcelHeaderStyle(14857357,14)]
        [Display(Name ="DOCUMENT NUMBER")]
        public string DocumentNumber { get; set; }

        [ExcelHeaderStyle(14857357, 6.55)]
        [Display(Name = "DOC TYPE")]
        public string DocType { get; set; }

        [ExcelHeaderStyle(14857357, 5.5)]
        [Display(Name = "ENTRY DATE")]
        public string EntryDate { get; set; }

        [ExcelHeaderStyle(14857357, 12.00)]
        [Display(Name = "POSTING DATE")]
        public string PostingDate { get; set; }

        [ExcelHeaderStyle(14857357, 12.00)]
        [Display(Name = "DOCUMENT DATE")]
        public string DocumentDate { get; set; }

        [ExcelHeaderStyle(65535, 12.00)]
        [Display(Name = "TRANSLATION DATE")]
        public string TranslationDate { get; set; }

        [ExcelHeaderStyle(65535, 12.00)]
        [Display(Name = "USER NAME")]
        public string UserName { get; set; }

        [ExcelHeaderStyle(14857357, 7.0)]
        [Display(Name = "DOC CURRENCY")]
        public string DocCurrency { get; set; }

        [ExcelHeaderStyle("#,##0.00",14857357, 15.5)]
        [Display(Name = "Amount in Document Currency")]
        public float AmtInDocCur { get; set; }

        [ExcelHeaderStyle(14857357,8.43)]
        [Display(Name = "LOCAL CURRENCY")]
        public string LocalCur { get; set; }

        [ExcelHeaderStyle("#,##0.00", 14857357, 11.57)]
        [Display(Name = "Amount in Local Currency")]
        public float AmtInlocalCur { get; set; }

        [ExcelHeaderStyle(14857357, 8.71)]
        [Display(Name = "GROUP CURRENCY")]
        public string GroupCur { get; set; }

        [ExcelHeaderStyle("#,##0.00", 14857357, 11.43)]
        [Display(Name = "Amount in Group Currency")]
        public float AmtInGroupCur { get; set; }

        [ExcelFormula("=L2/P2")]
        [ExcelHeaderStyle("0.00000",14857357, 7.86)]
        [Display(Name = "TC/GC")]
        public float TC_GC
        {
            get
            {
                if (AmtInGroupCur != 0)
                    return AmtInDocCur / AmtInGroupCur;
                return 0;
            }
        }

        [ExcelFormula("=N2/P2")]
        [ExcelHeaderStyle("0.00000", 14857357, 10.57)]
        [Display(Name = "LC/GC")]
        public float LC_GC
        {
            get
            {
                if (AmtInGroupCur != 0)
                    return AmtInlocalCur / AmtInGroupCur;
                return 0;
            }
        }

        private float _ob08TC_GC;

        
        [ExcelHeaderStyle("0.00000", 5287936, 9.29)]
        [Display(Name = "OB08 Ex Rate TC/GC")]
        public float OB08ExTC_GC
        {
            get
            {
                if (DocCurrency == GroupCur)
                    return 1;
                return _ob08TC_GC;
            }
            set
            {
                _ob08TC_GC = value;
            }
        }

        private float _ob08LC_GC;

        [ExcelHeaderStyle("0.00000", 5287936, 9.14)]
        [Display(Name = "OB08 Ex Rate LC/GC")]
        public float OB08ExLC_GC
        {
            get
            {
                if (LocalCur == GroupCur)
                    return 1;
                return _ob08LC_GC;
            }
            set
            {
                _ob08LC_GC = value;
            }
        }

        [ExcelFormula("=Q2-S2")]
        [ExcelHeaderStyle("0.00000", 14857357, 8.71)]
        [Display(Name = "TC/GC")]
        public float TC_GC_Delta
        {
            get
            {
                return TC_GC - OB08ExTC_GC;
            }
        }

        [ExcelFormula("=R2-T2")]
        [ExcelHeaderStyle("0.00000", 14857357, 8.71)]
        [Display(Name = "LC/GC")]
        public float LC_GC_Delta {
            get
            {
                return LC_GC - OB08ExLC_GC;
            }
        }

        [ExcelFormula("=(N2/S2)-P2")]
        [ExcelHeaderStyle("0.00", 49407, 12.00)]
        [Display(Name = "Delta LC/GC")]
        public float Delta_LC_GC
        {
            get
            {
                if (OB08ExLC_GC != 0)
                    return AmtInlocalCur / OB08ExLC_GC - AmtInGroupCur;
                return 0;
            }
        }

        [ExcelFormula("=(P2*S2)-N2")]
        [ExcelHeaderStyle("0.00", 49407, 12.00)]
        [Display(Name = "Delta GC/LC In Loc")]
        public float Delta_GC_LC
        {
            get
            {
                return AmtInGroupCur * OB08ExTC_GC - AmtInlocalCur;
            }
        }

        [ExcelFormula("=(L2/S2)-P2")]
        [ExcelHeaderStyle("0.00", 49407, 13.14)]
        [Display(Name = "Delta TC/LC/GC")]
        public string Delta_TC_LC_GC
        {
            get
            {
                if (OB08ExTC_GC == 0)
                    return "ERROR";
                return (AmtInDocCur / OB08ExTC_GC - AmtInGroupCur).ToString();
            }
        }

        
        ///Curr. Doc Cur
        ///LCurr Local Cur
        ///Lcur2 Group Cur
    }
}
