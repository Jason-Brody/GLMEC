using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;
using TestScript.Case1;
using TestScript;
using System.Data;
using System.IO;
using Young.Data;
using SAPAutomation;
using SAPFEWSELib;
using Ex = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace CaseTrigger
{
    class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine(Utils.FillNumber("200000160"));
            var _dt = ExcelHelper.Current.Open("Case1_MTD_Analysis.xlsx").Read("Case1_MTD_Analysis");
            ExcelHelper.Current.Close();
            var data = _dt.Rows[0].ToEntity<Case1DataModel>();
            var datas = mergeData(@"E:\GitHub\GLMEC\CaseTrigger\bin\Debug\ReportData\DE50\Datas");
            datas.ExportToExcel("test.xlsx", "MTD Analysis");
            Console.WriteLine(DateTime.Now);
            Console.WindowHeight = 1;
            Console.WindowWidth = 1;
            Case1_MTD_Analysis case1 = new Case1_MTD_Analysis();
            case1.Run();


        }

      


        private static void changeLayout(Dictionary<string, int> columns)
        {
            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("btn[32]").Press();
            var displayedColumnsGrid = SAPTestHelper.Current.PopupWindow.FindById<GuiGridView>("usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell");
            if (displayedColumnsGrid.RowCount > 0)
            {
                displayedColumnsGrid.SelectAll();
                SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("APP_FL_SING").Press();
            }
            var columnSetGrid = SAPTestHelper.Current.PopupWindow.FindById<GuiGridView>("usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell");

            string selectedRow = "";

            for (int c = 0; c < columnSetGrid.RowCount; c++)
            {
                var col = columnSetGrid.GetCellValue(c, "SELTEXT");

                if (columns.ContainsKey(col))
                {
                    selectedRow += c.ToString() + ",";
                    columns[col] = c;
                }

            }

            columnSetGrid.SelectedRows = selectedRow;
            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("APP_WL_SING").Press();

            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[0]").Press();
        }

        static void OB08(DataTable dt)
        {
            
        }


        private static List<Case1ReportDataModel> mergeData(string dir)
        {
            List<Case1ReportDataModel> report = new List<Case1ReportDataModel>();

            //Dictionary<string, int> dic = new Dictionary<string, int>();
            //dic.Add("LOGICAL SYSTEM", 1);
            //dic.Add("COMPANY CODE", 2);
            //dic.Add("ACCOUNT", 3);
            //dic.Add("DOCUMENT NUMBER", 4);
            //dic.Add("DOC TYPE", 5);
            //dic.Add("ENTRY DATE", 6);
            //dic.Add("POSTING DATE", 7);
            //dic.Add("DOCUMENT DATE", 8);
            //dic.Add("TRANSLATION DATE", 9);
            //dic.Add("USER NAME", 10);
            //dic.Add("DOC CURRENCY", 11);
            //dic.Add("Amount in Document Currency", 12);
            //dic.Add("LOCAL CURRENCY", 13);
            //dic.Add("Amount in Local Currency", 14);
            //dic.Add("GROUP CURRENCY", 15);
            //dic.Add("Amount in Group Currency", 16);
            //dic.Add("TC/GC", 17);
            //dic.Add("LC/GC", 18);
            //dic.Add("OB08 Ex Rate TC/GC", 19);
            //dic.Add("OB08 Ex Rate LC/GC", 20);
            //dic.Add("TC/GC delta", 21);
            //dic.Add("LC/GC delta", 22);
            //dic.Add("Delta LC/GC", 23);
            //dic.Add("Delta GC/LC In Loc", 24);
            //dic.Add("Delta TC/LC/GC", 25);


            //DataTable mergeTable = new DataTable();
            //foreach (var item in dic)
            //{
            //    mergeTable.Columns.Add(item.Key);
            //}

            foreach (var f in Directory.GetFiles(dir))
            {
                DataTable dt = Utils.ReadStringToTable(f, (s, h) =>
                {
                    string splitChar = "|";
                    if (!s.Contains(splitChar) || s == h || s.Contains("*"))
                        return null;

                    var vals = s.Split(splitChar.ToCharArray().First());
                    var returnVals = new List<string>();
                    for (int i = 0; i < vals.Count(); i++)
                    {
                        returnVals.Add(vals[i].Trim());
                    }
                    return returnVals;

                });

                foreach (DataRow dr in dt.Rows)
                {
                    Case1ReportDataModel rp = new Case1ReportDataModel();
                    rp.DocumentNumber = dr[1].ToString();
                    rp.DocType = dr[2].ToString();
                    rp.DocumentDate = dr[3].ToString();
                    rp.AmtInlocalCur = Utils.GetAmount(dr[4].ToString());
                    rp.LocalCur = dr[5].ToString();
                    rp.PostingDate = dr[6].ToString();
                    rp.CompanyCode = dr[7].ToString();
                    rp.AmtInGroupCur = Utils.GetAmount(dr[8].ToString());
                    rp.GroupCur = dr[9].ToString();
                    rp.Account = dr[10].ToString();
                    rp.AmtInDocCur = Utils.GetAmount(dr[11].ToString());
                    rp.DocCurrency = dr[12].ToString();
                    rp.EntryDate = dr[13].ToString();

                    report.Add(rp);
                    //var amtInGroupCur = Utils.GetAmount(dr[8].ToString());
                    //var amtInDocCur = Utils.GetAmount(dr[11].ToString());
                    //var amdInLolCur = Utils.GetAmount(dr[4].ToString());

                    //DataRow newDr = mergeTable.NewRow();
                    //newDr["DOCUMENT NUMBER"] = dr[1];
                    //newDr["DOC TYPE"] = dr[2];
                    //newDr["DOCUMENT DATE"] = dr[3];
                    //newDr["Amount in Local Currency"] = amdInLolCur;
                    //newDr["LOCAL CURRENCY"] = dr[5];
                    //newDr["POSTING DATE"] = dr[6];
                    //newDr["COMPANY CODE"] = dr[7];
                    //newDr["Amount in Group Currency"] = amtInGroupCur;
                    //newDr["GROUP CURRENCY"] = dr[9];
                    //newDr["ACCOUNT"] = dr[10];
                    //newDr["Amount in Document Currency"] = amtInDocCur;
                    //newDr["DOC CURRENCY"] = dr[12];
                    //newDr["ENTRY DATE"] = dr[13];
                    //newDr["TC/GC"] = amtInDocCur / amtInGroupCur;
                    //newDr["LC/GC"] = amdInLolCur / amtInGroupCur;


                    //mergeTable.Rows.Add(newDr);
                }
            }

            report.Export("test.csv", "|");
            return report;

        }

        static float getAmount(string val)
        {
            val.Replace(",", "");
            var pos = val.IndexOf('-');
            if(val.Contains("-")&&val.IndexOf('-')==val.Length-1)
            {
                val = "-" + val.Substring(0, val.Length - 1);
            }
            return float.Parse(val);
        }




    }
}
