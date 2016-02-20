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

namespace TestScript.Case1
{
    public class Case1_MTD_Analysis : IRunner
    {
        private string _workDir;
        private string _tempDir;
        private string _glAccountFileName;
        private string _accountPath;
        private string _reportSourceDataDir;

        public Case1_MTD_Analysis()
        {
            _steps.Add(new StepInfo() { Id = 1, Name = "Read GL Account Data from txt file", IsProcessKnown = false });
            _steps.Add(new StepInfo() { Id = 2, Name = "Login to LH", IsProcessKnown = false });
            _steps.Add(new StepInfo() { Id = 3, Name = "Get GL Accounts", IsProcessKnown = true });
            _steps.Add(new StepInfo() { Id = 4, Name = "Get Report Data", IsProcessKnown = true });
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

        

        private void setCondig(Case1DataModel data)
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            _workDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ReportData", data.CompanyCode);
            _tempDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Temp");
            _reportSourceDataDir = Path.Combine(_workDir, "Datas");
            _glAccountFileName = $"{data.CompanyCode}.txt";
            _accountPath = @"Case1\Accountlist.txt";
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

                mergeData(_reportSourceDataDir);
            }
        }

        private DataTable formatGlAccountData()
        {
            var accountDatas = Utils.ReadStringToTable(Path.Combine(_workDir,_glAccountFileName), (s,s1) =>
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

        private DataTable mergeData(string dir)
        {
            DataTable mergeTable = null;

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

                if (mergeTable == null)
                    mergeTable = dt.Copy();
                else
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        DataRow newDr = mergeTable.NewRow();
                        for (int i = 0; i < mergeTable.Columns.Count; i++)
                        {
                            newDr[i] = dr[i];
                        }
                        mergeTable.Rows.Add(newDr);
                    }

                }


            }

            return mergeTable;

        }

        private void login(LoginDataModel data)
        {
            var step = _steps.Where(s => s.Id == 2).First();
            //OnProcess(step);
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
            //OnProcess(step);
            List<AccountModel> accounts = new List<AccountModel>();

            foreach (DataRow dr in accountData.Rows)
            {
                AccountModel acct = dr.ToEntity<AccountModel>();
                accounts.Add(acct);
            }

            step.TotalProcess = accounts.Count;
            step.CurrentProcess = 1;
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


                //var dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,$"ReportData\\{data.CompanyCode}");
                if (accounts[i].DataCount>0)
                {
                    //click filter button
                    h.MainWindow.FindByName<GuiButton>("btn[38]").Press();



                    var grid = h.PopupWindow.FindById<GuiGridView>("usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER2_FILT/shellcont/shell");
                    if (grid.RowCount > 0)
                    {
                        grid.SelectAll();
                       h.PopupWindow.FindByName<GuiButton>("APP_FL_SING").Press();
                    }

                    grid = h.PopupWindow.FindById<GuiGridView>("usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell");
                    for (int j = 0; j < grid.RowCount; j++)
                    {
                        if (grid.GetCellValueByDisplayColumn(j, "Column Name") == "Document Date")
                        {
                            grid.SelectedRows = j.ToString();
                            h.PopupWindow.FindByName<GuiButton>("APP_WL_SING").Press();
                            h.PopupWindow.FindByName<GuiButton>("600_BUTTON").Press();
                            h.PopupWindow.FindByName<GuiCTextField>("%%DYN001-LOW").Text = data.DocDateFrom;
                            h.PopupWindow.FindByName<GuiCTextField>("%%DYN001-HIGH").Text = data.DocDateTo;
                            h.PopupWindow.FindByName<GuiButton>("btn[0]").Press();
                            break;
                        }
                    }

                    DateTime end = DateTime.Now;

                    accounts[i].Period = end.Subtract(start);

                   

                    string file = $"{accounts[i].Account}.txt";

                    accounts[i].DataCount = SAPTestHelper.Current.MainWindow.FindById<GuiGridView>("usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").RowCount;

                    outputFile("mbar/menu[0]/menu[3]/menu[2]", _reportSourceDataDir, file);
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
            //OnProcess(step);

            List<List<Tuple<int, string>>> accounts = new List<List<Tuple<int, string>>>();
            using (StreamReader sr = new StreamReader(_accountPath))
            {
                while (!sr.EndOfStream)
                {
                    accounts.Add(new List<Tuple<int, string>>() { new Tuple<int, string>(1, sr.ReadLine()) });
                }

            }

            SAPTestHelper.Current.SAPGuiSession.StartTransaction("SE16");

            SAPTestHelper.Current.SAPGuiSession.FindById<GuiCTextField>("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "SKB1";
            SAPTestHelper.Current.SAPGuiSession.FindById<GuiMainWindow>("wnd[0]").SendVKey(0);
            SAPTestHelper.Current.SAPGuiSession.FindById<GuiCTextField>("wnd[0]/usr/ctxtI1-LOW").Text = data.CompanyCode;

            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("%_I2_%_APP_%-VALU_PUSH").Press();

            var table = SAPTestHelper.Current.PopupWindow.FindByName<GuiTableControl>("SAPLALDBSINGLE");
            table.SetBatchValues(accounts);

            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[8]").Press();
            SAPTestHelper.Current.MainWindow.FindByName<GuiTextField>("MAX_SEL").Text = "10000";
            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("btn[8]").Press();


            outputFile("mbar/menu[0]/menu[10]/menu[3]/menu[2]", _workDir, _glAccountFileName);
           
            step.IsComplete = true;
        }

        private void outputFile(string outputmenuId,string dir,string fileName)
        {
            SAPTestHelper.Current.MainWindow.FindById<GuiMenu>(outputmenuId).Select();
            SAPTestHelper.Current.PopupWindow.FindByName<GuiRadioButton>("SPOPLI-SELFLAG").Select();
            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[0]").Press();

            var filePath = Path.Combine(dir, fileName);

            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            if (File.Exists(filePath))
                File.Delete(filePath);

            SAPTestHelper.Current.PopupWindow.FindByName<GuiCTextField>("DY_PATH").Text = dir;
            SAPTestHelper.Current.PopupWindow.FindByName<GuiCTextField>("DY_FILENAME").Text = fileName;

            var windowName = SAPTestHelper.Current.MainWindow.Text;

            var ts = new CancellationTokenSource();
            var ct = ts.Token;

            Task.Run(() => { UIHelper.SetAccess(windowName,ct); });
            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[0]").Press();
            ts.Cancel();
        }

        private void ReadData()
        {
            var step = _steps.First();
            //OnProcess(step);
            _dt = ExcelHelper.Current.Open("Case1_MTD_Analysis.xlsx").Read("Case1_MTD_Analysis");
            ExcelHelper.Current.Close();
            step.IsComplete = true;
        }
    }
}
