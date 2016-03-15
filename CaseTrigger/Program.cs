using SAPAutomation;
using SAPFEWSELib;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using TestScript;
using TestScript.Case1;
using TestScript.Case6;
using Young.Data;
using Young.Data.Extension;
using Ex = Microsoft.Office.Interop.Excel;
using ScriptRunner.Interface;

namespace CaseTrigger
{
    class Program
    {
        static void DoIt()
        {
            try
            {
                Console.WriteLine("inner try");
                int i = 0;
                Console.WriteLine(12 / i); // oops
            }
            catch (Exception e)
            {
                Console.WriteLine("inner catch");
                throw e; // or "throw", or "throw anything"
            }
            finally
            {
                Console.WriteLine("inner finally");
            }
        }


        const string KEY_64 = "VavicApp";//注意了，是8个字符，64位

        const string IV_64 = "VavicApp";

        public static string Encode(string data)
        {
            byte[] byKey = System.Text.ASCIIEncoding.ASCII.GetBytes(KEY_64);
            byte[] byIV = System.Text.ASCIIEncoding.ASCII.GetBytes(IV_64);

            DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
            int i = cryptoProvider.KeySize;
            MemoryStream ms = new MemoryStream();
            CryptoStream cst = new CryptoStream(ms, cryptoProvider.CreateEncryptor(byKey, byIV), CryptoStreamMode.Write);

            StreamWriter sw = new StreamWriter(cst);
            sw.Write(data);
            sw.Flush();
            cst.FlushFinalBlock();
            sw.Flush();
            return Convert.ToBase64String(ms.GetBuffer(), 0, (int)ms.Length);

        }

        public static string Decode(string data)
        {
            byte[] byKey = System.Text.ASCIIEncoding.ASCII.GetBytes(KEY_64);
            byte[] byIV = System.Text.ASCIIEncoding.ASCII.GetBytes(IV_64);

            byte[] byEnc;
            try
            {
                byEnc = Convert.FromBase64String(data);
            }
            catch
            {
                return null;
            }

            DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
            MemoryStream ms = new MemoryStream(byEnc);
            CryptoStream cst = new CryptoStream(ms, cryptoProvider.CreateDecryptor(byKey, byIV), CryptoStreamMode.Read);
            StreamReader sr = new StreamReader(cst);
            return sr.ReadToEnd();
        }

        class Person
        {
            public string FirstName { get; set; }
            public string LastName { get; set; }
        }

        class Pet
        {
            public string Name { get; set; }
            public Person Owner { get; set; }
        }

        public static void LeftOuterJoinExample()
        {
            Person magnus = new Person { FirstName = "Magnus", LastName = "Hedlund" };
            Person terry = new Person { FirstName = "Terry", LastName = "Adams" };
            Person charlotte = new Person { FirstName = "Charlotte", LastName = "Weiss" };
            Person arlene = new Person { FirstName = "Arlene", LastName = "Huff" };

            Pet barley = new Pet { Name = "Barley", Owner = terry };
            Pet boots = new Pet { Name = "Boots", Owner = terry };
            Pet whiskers = new Pet { Name = "Whiskers", Owner = charlotte };
            Pet bluemoon = new Pet { Name = "Blue Moon", Owner = terry };
            Pet daisy = new Pet { Name = "Daisy", Owner = magnus };

            // Create two lists.
            List<Person> people = new List<Person> { magnus, terry, charlotte, arlene };
            List<Pet> pets = new List<Pet> { barley, boots, whiskers, bluemoon, daisy };

            var query = from person in people
                        join pet in pets on person equals pet.Owner into gj
                        from subpet in gj.DefaultIfEmpty()
                        select new { person.FirstName, PetName = (subpet == null ? String.Empty : subpet.Name) };

            foreach (var v in query)
            {
                Console.WriteLine("{0,-15}{1}", v.FirstName + ":", v.PetName);
            }
        }


        
        public static string ChooseNode(GuiTree tree, string path)
        {
            var paths = path.Split(new string[] { "->" }, StringSplitOptions.None);
            var initialLevel = 0;
            var myKey = "";
            foreach (var key in tree.GetAllNodeKeys())
            {
                var level = tree.GetHierarchyLevel(key);
                if(level == initialLevel)
                {
                    var node = tree.GetNodeTextByKey(key);
                    if (node.ToLower().Trim() == paths[initialLevel].ToLower().Trim())
                    {
                        initialLevel++;
                        if (initialLevel == paths.Count())
                        {
                            myKey = key;
                            break;
                        }
                    }
                }
            }
            if (myKey != "")
            {
                List<string> keyList = new List<string>();
                var parentKey = tree.GetParent(myKey);
                while(parentKey.Trim()!="")
                {
                    keyList.Add(parentKey);
                    parentKey = tree.GetParent(parentKey);
                }
                var count = keyList.Count();
                if(count>0)
                {
                    for (int i = count - 1; i >= 0; i--)
                    {
                        tree.ExpandNode(keyList[i]);
                    }
                    
                }
                tree.SelectNode(myKey);
            }
               
            return myKey;
           
        }

        static void Main(string[] args)
        {



            //SAPTestHelper.Current.SetSession();
            //var tree = SAPTestHelper.Current.MainWindow.FindDescendantByProperty<GuiTree>();

            ////SAPTestHelper.Current.SAPGuiSession.FindById<GuiTree>("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell").ExpandNode("         53");
            ////SAPTestHelper.Current.SAPGuiSession.FindById<GuiTree>("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell").SelectNode("         66");

            ////var n = tree.SelectedNode;
            ////var color = tree.GetNodeTextColor("         62");
            ////tree.ChooseNode("General Ledger Line Items->Order");


            //var a1 = tree.SelectedItemNode();
            ////SAPTestHelper.Current.SAPGuiSession.FindById<GuiToolbarControl>("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell").PressButton("TAKE");


            //foreach (var item in tree.GetAllNodeKeys())
            //{
            //    var top = tree.Top;
            //    var header = tree.GetNodeTextByKey(item);
            //    var b = tree.GetHierarchyLevel(item);
            //    var a = tree.GetNodePathByKey(item);
            //    Console.WriteLine(header + "color:" + tree.GetNodeTextColor(item));
            //}
            DataTable dt = ExcelHelper.Current.Open("Case1_MTD_Analysis.xlsx").Read("Case6_WorkFlow");
            var myDatas = dt.ToList<Case6DataModel>();
            foreach (var d in myDatas)
            {
                Case6_Workflow script = new Case6_Workflow();
                var runner = new ScriptEngine<Case6DataModel>(script);

                runner.Run(d);
                // runner.Run(d);

            }
















        }

        public static void SetAccess(string windowName, CancellationToken token)
        {
            if (token.IsCancellationRequested)
                token.ThrowIfCancellationRequested();

            bool isPress = false;
            while (!isPress)
            {

                var e = TreeWalker.ControlViewWalker.GetFirstChild(AutomationElement.RootElement);

                while (e != null)
                {
                    if (e.Current.Name.ToLower().Contains(windowName.ToLower()))
                        break;
                    var tempE = TreeWalker.ControlViewWalker.GetNextSibling(e);
                    e = tempE;


                }
                if (e != null)
                {
                    var condition1 = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
                    var condition2 = new PropertyCondition(AutomationElement.AccessKeyProperty, "Alt+A");
                    var andCondition = new AndCondition(condition1, condition2);
                    var btnElement = e.FindFirst(TreeScope.Descendants, andCondition);
                    if (btnElement != null)
                    {
                        var btn = btnElement.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;

                        btn.Invoke();
                        isPress = true;
                    }
                }
            }

        }

        

        private static void format(object worksheet)
        {
            var sheet= worksheet as Ex.Worksheet;

            var range = sheet.Cells[2, 1] as Ex.Range;

            range.EntireRow.Insert();

            range = sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, 25]] as Ex.Range;
            range.Interior.Color = 16777215;

            range = sheet.Range[sheet.Cells[1, 21], sheet.Cells[1, 25]] as Ex.Range;
            range.Copy();

            range = sheet.Range[sheet.Cells[2, 21], sheet.Cells[2, 25]] as Ex.Range;
            range.PasteSpecial(Ex.XlPasteType.xlPasteAll);

            range = sheet.Range[sheet.Cells[1, 21], sheet.Cells[1, 22]] as Ex.Range;
            range.Value = "";
            range.Merge();

            range.Value = "Check between OB08 and TC";
            range.Interior.Color = 255;
            range.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Ex.XlVAlign.xlVAlignBottom;
            range.Font.Bold = true;
            range.Font.Size = 10;

            range = sheet.Range[sheet.Cells[1, 23], sheet.Cells[1, 25]] as Ex.Range;
            range.Value = "";
            range.Merge();

            range.Value = "Delta Check";
            range.Interior.Color = 49407;
            range.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Ex.XlVAlign.xlVAlignCenter;
            range.Font.Bold = true;
            range.Font.Size = 10;

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
                DataTable dt = Young.Data.Utils.ReadStringToTable(f, (s, h) =>
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
                    rp.AmtInlocalCur = SAPAutomation.Utils.GetAmount(dr[4].ToString());
                    rp.LocalCur = dr[5].ToString();
                    rp.PostingDate = dr[6].ToString();
                    rp.CompanyCode = dr[7].ToString();
                    rp.AmtInGroupCur = SAPAutomation.Utils.GetAmount(dr[8].ToString());
                    rp.GroupCur = dr[9].ToString();
                    rp.Account = dr[10].ToString();
                    rp.AmtInDocCur = SAPAutomation.Utils.GetAmount(dr[11].ToString());
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

            report.ExportToFile("test.csv", "|");
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
