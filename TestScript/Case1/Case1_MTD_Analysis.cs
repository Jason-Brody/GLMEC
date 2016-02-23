using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPAutomation;
using Young.Data;
using System.Data;
using SAPFEWSELib;
using System.IO;
using System.Threading;
using CaseRunnerModel;
using Ex = Microsoft.Office.Interop.Excel;
using Young.Data.Extension;
using Young.Excel.Interop.Extensions;

namespace TestScript.Case1
{
    public class Case1_MTD_Analysis : IRunner
    {
        private string _workDir;
        private string _tempDir;
        private string _glAccountFileName;
        private string _accountPath;
        private string _reportSourceDataDir;
        private Dictionary<string, int> _dataOutputLayout;
        private List<Case1ReportDataModel> _reportData;
        private List<Case1DocInfoModel> _docInfoes;
        private string _docInfoFileName = "SE16_BKPF.txt";
        private string _reportFileName = "report.csv";

        public Case1_MTD_Analysis()
        {
            _steps.Add(new StepInfo() { Id = 1, Name = "Read Config Data", IsProcessKnown = false });
            _steps.Add(new StepInfo() { Id = 2, Name = "Login to LH", IsProcessKnown = false });
            _steps.Add(new StepInfo() { Id = 3, Name = "Get GL Accounts", IsProcessKnown = true });
            _steps.Add(new StepInfo() { Id = 4, Name = "Get Report Data", IsProcessKnown = true });
            _steps.Add(new StepInfo() { Id = 5, Name = "Merge Report Data", IsProcessKnown = true });
            _steps.Add(new StepInfo() { Id = 6, Name = "Get Currency Rate", IsProcessKnown = true });
            _steps.Add(new StepInfo() { Id = 7, Name = "Get Detail info by doc number", IsProcessKnown = true });
            _steps.Add(new StepInfo() { Id = 8, Name = "Generate Report", IsProcessKnown = true });
            _steps.Add(new StepInfo() { Id = 9, Name = "Format Report", IsProcessKnown = true });
        }

        private void process(StepInfo step)
        {
            if (OnProcess != null)
                OnProcess(step);
        }
        

        private DataTable _dt;

        private List<StepInfo> _steps = new List<StepInfo>();

        public event ProcessHander OnProcess;

        public List<StepInfo> GetSteps
        {
            get
            {
                return _steps;
            }
        }

        public CaseInfo Info
        {
            get
            {
                return new CaseInfo() { Name = nameof(Case1_MTD_Analysis) };
            }
        }

        private void initialStep(StepInfo s)
        {
            process(s);
            s.IsComplete = false;
            s.CurrentProcess = 0;
            s.TotalProcess = 0;
        }

        private void setCondig(Case1DataModel data)
        {
            _steps.Where(s=>s.Id > 1).ToList().ForEach(s => { s.IsComplete = false; s.CurrentProcess = 0; s.TotalProcess = 0; });
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            _workDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ReportData", data.CompanyCode);
            _tempDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Temp");
            _reportSourceDataDir = Path.Combine(_workDir, "Datas");
            _glAccountFileName = $"{data.CompanyCode}.txt";
            _accountPath = @"Case1\Accountlist.txt";
            _dataOutputLayout = new Dictionary<string, int>() { };
            _dataOutputLayout.Add("Company Code",-1);
            _dataOutputLayout.Add("Account", -1);
            _dataOutputLayout.Add("Document Number", -1);
            _dataOutputLayout.Add("Document Type", -1);
            _dataOutputLayout.Add("Entry Date", -1);
            _dataOutputLayout.Add("Posting Date", -1);
            _dataOutputLayout.Add("Document Date", -1);
            _dataOutputLayout.Add("Document currency", -1);
            _dataOutputLayout.Add("Local Currency", -1);
            _dataOutputLayout.Add("Local currency 2", -1);
            _dataOutputLayout.Add("Amount in local currency", -1);
            _dataOutputLayout.Add("Amount in doc. curr.", -1);
            _dataOutputLayout.Add("Amount in loc.curr.2", -1);


        }

        public void Run()
        {

            ReadData();
            foreach (DataRow dr in _dt.Rows)
            {
                var data = dr.ToEntity<Case1DataModel>();
                setCondig(data);
                login(data);

                getGLAccount(data);

                getData(data, formatGlAccountData());

                _reportData = mergeData(_reportSourceDataDir);

                var curDic = getCurrency();

                getDocInfo(data);

                generateReport(curDic);

                formatReport();

                SAPTestHelper.Current.SAPGuiConnection.CloseConnection();
            }
        }

        private DataTable formatGlAccountData()
        {
            
            var accountDatas = Young.Data.Utils.ReadStringToTable(Path.Combine(_workDir,_glAccountFileName), (s,s1) =>
            {
                string splitchar = "|";
                if (!s.Contains(splitchar))
                    return null;
                var vals = s.Split(splitchar.ToCharArray().First());
                var returnVals = new List<string>();
                for (int i = 0; i < vals.Count(); i++)
                {
                    returnVals.Add(vals[i].Trim());
                }
                return returnVals;
            });
            return accountDatas;
        }

        private List<Case1ReportDataModel> mergeData(string dir)
        {

            var step = _steps.Where(s => s.Id == 5).First();
            initialStep(step);
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
            step.TotalProcess = Directory.GetFiles(dir).Count();
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
                    rp.DocumentNumber =SAPAutomation.Utils.FillNumber(dr[1].ToString());
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
                step.CurrentProcess++;
            }
            step.IsComplete = true;
            return report;

        }

        private Dictionary<string,float> getCurrency()
        {
            var step = _steps.Where(s => s.Id == 6).First();
            initialStep(step);

            var curList = new List<string>();
            curList.AddRange(_reportData.Where(c=>c.DocCurrency.ToLower()!="usd").GroupBy(g => g.DocCurrency).Select(s => s.Key));
            curList.AddRange(_reportData.Where(c => c.LocalCur.ToLower() != "usd").GroupBy(g => g.LocalCur).Select(s => s.Key));
            curList = curList.GroupBy(s => s).Select(s => s.Key).ToList();
            Dictionary<string, float> rateDic = new Dictionary<string, float>();
            curList.ForEach(s => rateDic.Add(s, 0));

            DateTime date = DateTime.Now;
            var firstDayOfMonth = new DateTime(date.Year, date.Month, 1);
            string d = firstDayOfMonth.ToString("dd.MM.yyyy");

            if(rateDic.Count>0)
            {
                SAPTestHelper.Current.SAPGuiSession.StartTransaction("OB08");
                SAPTestHelper.Current.MainWindow.SendKey(SAPKeys.Ctrl_F4);
            }

            step.TotalProcess = rateDic.Count;

            for (int i =0;i<rateDic.Count;i++)
            {
                var cur = rateDic.ElementAt(i).Key;
                SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("VIM_POSI_PUSH").Press();
                SAPTestHelper.Current.PopupWindow.FindById<GuiCTextField>("usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").Text = "M";
                SAPTestHelper.Current.PopupWindow.FindById<GuiCTextField>("usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[1,21]").Text = cur;
                SAPTestHelper.Current.PopupWindow.FindById<GuiCTextField>("usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[2,21]").Text = "USD";
                SAPTestHelper.Current.PopupWindow.FindById<GuiCTextField>("usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[3,21]").Text = d;
                SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[0]").Press();

                var table = SAPTestHelper.Current.MainWindow.FindByName<GuiTableControl>("SAPL0SAPTCTRL_V_TCURR");

                if (table.RowCount > 0)
                {
                    if (table.GetCell(0,5).Text.ToLower()== cur.ToLower())
                    {
                        rateDic[cur] = SAPAutomation.Utils.GetAmount(table.GetCell(0, 2).Text) / SAPAutomation.Utils.GetAmount(table.GetCell(0, 9).Text);
                    }
                }

                step.CurrentProcess++;

            }
            step.IsComplete = true;
            return rateDic;
           
        }

        //private void changeLayout(Dictionary<string,int> columns)
        //{
        //    SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("btn[32]").Press();
        //    var displayedColumnsGrid = SAPTestHelper.Current.PopupWindow.FindById<GuiGridView>("usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell");
        //    if (displayedColumnsGrid.RowCount > 0)
        //    {
        //        displayedColumnsGrid.SelectAll();
        //        SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("APP_FL_SING").Press();
        //    }
        //    var columnSetGrid = SAPTestHelper.Current.PopupWindow.FindById<GuiGridView>("usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell");

        //    string selectedRow = "";

        //    for (int c = 0; c < columnSetGrid.RowCount; c++)
        //    {
        //        var col = columnSetGrid.GetCellValue(c, "SELTEXT");

        //        if (columns.ContainsKey(col))
        //        {
        //            selectedRow += c.ToString() + ",";
        //            columns[col] = c;
        //        }

        //    }

        //    columnSetGrid.SelectedRows = selectedRow;
        //    SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("APP_WL_SING").Press();

        //    SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[0]").Press();
        //}

        private void getDocInfo(Case1DataModel data)
        {
            var step = _steps.Where(s => s.Id == 7).First();
            initialStep(step);

            if (_reportData.Count>0)
            {
                SAPTestHelper.Current.SAPGuiSession.StartTransaction("SE16");
                SAPTestHelper.Current.MainWindow.FindByName<GuiCTextField>("DATABROWSE-TABLENAME").Text = "BKPF";
                SAPTestHelper.Current.MainWindow.SendKey(SAPKeys.Enter);
                SAPTestHelper.Current.MainWindow.FindByName<GuiCTextField>("I1-LOW").Text = data.CompanyCode;
                SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("%_I2_%_APP_%-VALU_PUSH").Press();
                var table = SAPTestHelper.Current.PopupWindow.FindByName<GuiTableControl>("SAPLALDBSINGLE");
                var docList = _reportData.Select(r => new List<Tuple<int, string>>() { new Tuple<int, string>(1, r.DocumentNumber) }).ToList();
                step.TotalProcess = docList.Count;
                table.SetBatchValues(docList,i=> { step.CurrentProcess++; });
                SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[8]").Press();

                var datetimeArray = data.PostingDateFrom.Split('.');

                var startFiscalYear = data.PostingStartDate.AddMonths(2).Year;
                var endFiscalYear = data.PostingEndDate.AddMonths(2).Year;

                SAPTestHelper.Current.MainWindow.FindByName<GuiTextField>("I3-LOW").Text = startFiscalYear.ToString();

                if (startFiscalYear != endFiscalYear)
                    SAPTestHelper.Current.MainWindow.FindByName<GuiTextField>("I3-HIGH").Text = endFiscalYear.ToString();


                SAPTestHelper.Current.MainWindow.FindByName<GuiCTextField>("I6-LOW").Text = data.PostingDateFrom;
                SAPTestHelper.Current.MainWindow.FindByName<GuiCTextField>("I6-HIGH").Text = data.PostingDateTo;
                SAPTestHelper.Current.MainWindow.FindByName<GuiTextField>("MAX_SEL").Text = "10000";

                SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("btn[8]").Press();

                var grid = SAPTestHelper.Current.MainWindow.FindById<GuiGridView>("usr/cntlGRID1/shellcont/shell");

                if(grid.RowCount>0)
                {
                    ///change layout
                    Dictionary<string, int> columns = new Dictionary<string, int>();
                    columns.Add("Logical System", -1);
                    columns.Add("Company Code", -1);
                    columns.Add("Document Number", -1);
                    columns.Add("Document Date", -1);
                    columns.Add("Document Type", -1);
                    columns.Add("Entry Date", -1);
                    columns.Add("Posting Date", -1);
                    columns.Add("Translation Date", -1);
                    columns.Add("User name", -1);
                    columns.Add("Currency", -1);
                    columns.Add("Local Currency", -1);
                    columns.Add("Local currency 2", -1);
                    columns.Add("Exchange Rate Type", -1);
                    columns.Add("Ref.key (header) 1", -1);
                    UIHelper.ChangeLayout(columns);

                    UIHelper.ExportFile("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]", _workDir, _docInfoFileName);
                }
            }

            step.IsComplete = true;
        }

        private void generateReport(Dictionary<string,float> curDic)
        {
            var step = _steps.Where(s => s.Id == 8).First();
            initialStep(step);

            var dt = Young.Data.Utils.ReadStringToTable(Path.Combine(_workDir, _docInfoFileName), (s, h) =>
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

            _docInfoes = new List<Case1DocInfoModel>();

            step.CurrentProcess = 0;
            step.TotalProcess = dt.Rows.Count + _reportData.Count;

            foreach (DataRow dr in dt.Rows)
            {
                Case1DocInfoModel di = new Case1DocInfoModel();
                di.CompanyCode = dr[1].ToString();
                di.DocumentNumber = SAPAutomation.Utils.FillNumber(dr[2].ToString());
                di.DocType = dr[3].ToString();
                di.DocDate = dr[4].ToString();
                di.PostingDate = dr[5].ToString();
                di.EntryDate = dr[6].ToString();
                di.TransDate = dr[7].ToString();
                di.UserName = dr[8].ToString();
                di.DocCurrency = dr[9].ToString();
                di.LocalCurrency = dr[10].ToString();
                di.GroupCurrency = dr[11].ToString();
                di.ExchangeRateType = dr[12].ToString();
                di.LogicalSystem = dr[14].ToString();
                di.RefKey1 = dr[15].ToString();

                _docInfoes.Add(di);

                step.CurrentProcess++;
            }

            var docDic = _docInfoes.ToDictionary(d => d.DocumentNumber);

            foreach (var r in _reportData)
            {
                if (curDic.ContainsKey(r.DocCurrency))
                    r.OB08ExTC_GC = curDic[r.DocCurrency];

                if (curDic.ContainsKey(r.LocalCur))
                    r.OB08ExLC_GC = curDic[r.LocalCur];

                if(docDic.ContainsKey(r.DocumentNumber))
                {
                    r.LogicalSystem = docDic[r.DocumentNumber].LogicalSystem;
                    r.TranslationDate = docDic[r.DocumentNumber].TransDate;
                    r.UserName = docDic[r.DocumentNumber].UserName;
                }

                step.CurrentProcess++;
            }

            _reportData.ExportToFile(Path.Combine(_workDir, _reportFileName), "|");

            step.IsComplete = true;
        }

        private void formatReport()
        {
            var step = _steps.Where(s => s.Id == 9).First();
            initialStep(step);

            string sheetName = "MTD Analysis";

            _reportData.ExportToExcel(Path.Combine(_workDir,"result.xlsx"), sheetName,(s)=> {

                var sheet = s as Ex.Worksheet;

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
            });

            step.IsComplete = true;

        }

        private void login(LoginDataModel data)
        {
            var step = _steps.Where(s => s.Id == 2).First();
            initialStep(step);
            SAPLogon l = new SAPLogon();
            l.StartProcess();
            l.OpenConnection(data.Address);
            l.Login(data.UserName, data.Password, data.Client, data.Language);
            SAPTestHelper.Current.SetSession(l);
            step.IsComplete = true;
        }

        private void getData(Case1DataModel data, DataTable accountData)
        {
            var step = _steps.Where(s => s.Id == 4).First();
            initialStep(step);

            List<AccountModel> accounts = new List<AccountModel>();

            foreach (DataRow dr in accountData.Rows)
            {
                AccountModel acct = dr.ToEntity<AccountModel>();
                accounts.Add(acct);
            }

            step.TotalProcess = accounts.Count;
            for (int i = 0; i < accounts.Count; i++)
            {
                //List<List<Tuple<int, string>>> testAccounts = accounts.Skip(i*10).Take(10).Select(
                //    c => new List<Tuple<int, string>>() { new Tuple<int, string>(1, c.Account) }).ToList();

                SAPTestHelper h = SAPTestHelper.Current;

                h.SAPGuiSession.StartTransaction("FAGLL03");
                h.MainWindow.FindByName<GuiCTextField>("SD_SAKNR-LOW").Text = accounts[i].Account;
                h.MainWindow.FindByName<GuiCTextField>("SD_BUKRS-LOW").Text = data.CompanyCode;

                //Set Multiple Accounts
                //h.MainWindow.FindByName<GuiButton>("%_SD_SAKNR_%_APP_%-VALU_PUSH").Press();
                //var table = h.PopupWindow.FindByName<GuiTableControl>("SAPLALDBSINGLE");
                //table.SetBatchValues(testAccounts);
                //h.PopupWindow.FindByName<GuiButton>("btn[8]").Press();

                h.MainWindow.FindByName<GuiRadioButton>("X_AISEL").Select();
                h.MainWindow.FindByName<GuiCTextField>("SO_BUDAT-LOW").Text = data.PostingDateFrom;
                h.MainWindow.FindByName<GuiCTextField>("SO_BUDAT-HIGH").Text = data.PostingDateTo;
                h.MainWindow.FindByName<GuiCTextField>("PA_VARI").Text = data.Layout;

                DateTime start = DateTime.Now;
                //click execute button
                h.MainWindow.FindByName<GuiButton>("btn[8]").Press();

                accounts[i].DataCount = SAPTestHelper.Current.MainWindow.FindById<GuiGridView>("usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").RowCount;

                if (accounts[i].DataCount>0)
                {
                    //change layout
                    UIHelper.ChangeLayout(_dataOutputLayout);
                    

                    DateTime end = DateTime.Now;

                    accounts[i].Period = end.Subtract(start);

                    string file = $"{accounts[i].Account}.txt";

                    accounts[i].DataCount = SAPTestHelper.Current.MainWindow.FindById<GuiGridView>("usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").RowCount;

                    UIHelper.ExportFile("mbar/menu[0]/menu[3]/menu[2]", _reportSourceDataDir, file);
                }

                

                string processFile = $"{_workDir}\\result.txt";

                using (StreamWriter sw = new StreamWriter(processFile, true))
                {
                    string line = $"{accounts[i].Account},{accounts[i].DataCount},{accounts[i].Period.Seconds}";
                    sw.WriteLine(line);
                }

                step.CurrentProcess++;

            }

            step.IsComplete = true;

        }

        private void getGLAccount(Case1DataModel data)
        {           
            var step = _steps.Where(s => s.Id == 3).First();
            initialStep(step);

            List<List<Tuple<int, string>>> accounts = new List<List<Tuple<int, string>>>();
            using (StreamReader sr = new StreamReader(_accountPath))
            {
                while (!sr.EndOfStream)
                {
                    accounts.Add(new List<Tuple<int, string>>() { new Tuple<int, string>(1, sr.ReadLine()) });
                }

            }

            step.TotalProcess = accounts.Count;

            SAPTestHelper.Current.SAPGuiSession.StartTransaction("SE16");

            SAPTestHelper.Current.SAPGuiSession.FindById<GuiCTextField>("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "SKB1";
            SAPTestHelper.Current.SAPGuiSession.FindById<GuiMainWindow>("wnd[0]").SendVKey(0);
            SAPTestHelper.Current.SAPGuiSession.FindById<GuiCTextField>("wnd[0]/usr/ctxtI1-LOW").Text = data.CompanyCode;

            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("%_I2_%_APP_%-VALU_PUSH").Press();

            var table = SAPTestHelper.Current.PopupWindow.FindByName<GuiTableControl>("SAPLALDBSINGLE");
            table.SetBatchValues(accounts,(i)=> { step.CurrentProcess++; });

            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[8]").Press();
            SAPTestHelper.Current.MainWindow.FindByName<GuiTextField>("MAX_SEL").Text = "10000";
            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("btn[8]").Press();


            UIHelper.ExportFile("mbar/menu[0]/menu[10]/menu[3]/menu[2]", _workDir, _glAccountFileName);
           
            step.IsComplete = true;
        }

        //private void outputFile(string outputmenuId,string dir,string fileName)
        //{
        //    SAPTestHelper.Current.MainWindow.FindById<GuiMenu>(outputmenuId).Select();
        //    SAPTestHelper.Current.PopupWindow.FindByName<GuiRadioButton>("SPOPLI-SELFLAG").Select();
        //    SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[0]").Press();

        //    var filePath = Path.Combine(dir, fileName);

        //    if (!Directory.Exists(dir))
        //        Directory.CreateDirectory(dir);

        //    if (File.Exists(filePath))
        //        File.Delete(filePath);

        //    SAPTestHelper.Current.PopupWindow.FindByName<GuiCTextField>("DY_PATH").Text = dir;
        //    SAPTestHelper.Current.PopupWindow.FindByName<GuiCTextField>("DY_FILENAME").Text = fileName;

        //    var windowName = SAPTestHelper.Current.MainWindow.Text;

        //    var ts = new CancellationTokenSource();
        //    var ct = ts.Token;

        //    Task.Run(() => { UIHelper.SetAccess(windowName,ct); });
        //    SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[0]").Press();
        //    ts.Cancel();
        //}

        private void ReadData()
        {
            var step = _steps.First();
            initialStep(step);
            _dt = ExcelHelper.Current.Open("Case1_MTD_Analysis.xlsx").Read("Case1_MTD_Analysis");
            ExcelHelper.Current.Close();
            step.IsComplete = true;
        }
    }
}
