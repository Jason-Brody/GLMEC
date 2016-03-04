using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CaseRunnerModel;
using CaseRunnerModel.Attributes;
using SAPAutomation;
using SAPFEWSELib;
using System.IO;
using Young.Data;
using Young.Data.Extension;

namespace TestScript.Case6
{
    public class Case6_Workflow : IScriptInitial
    {
        private Case6DataModel _data;
        private List<string> _idocList;
        private string _workDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Case6");
        private string _edp21Folder;
        private int _maximum = 100000;



        public void InitialData(object data)
        {
            _data = data as Case6DataModel;
            Tools.MasterDataVerification(_data);
            _idocList = Tools.GetDatas(Path.Combine(_workDir, "IDocList.txt"));


            _workDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Case6");
            _edp21Folder = Path.Combine(_workDir, "EDP21");

        }



        [Step(Order = 1, Name = "Login to SAP")]
        public void Login()
        {
            UIHelper.Login(_data);
        }

        [Step(Order = 2, Name = "TCode Verification")]
        public void TCodeVerifition()
        {
            UIHelper.SE16TableAccessVerification("EDIDC");
            UIHelper.SE16TableAccessVerification("EDIDS");
            UIHelper.SE16TableAccessVerification("EDP21");
            UIHelper.SE16TableAccessVerification("HRP1001");
        }

        private List<EDIDCDataModel> _EDIDCDataList = null;

        private string _EDIDCFileName = "EDIDC.txt";

        [Step(Order = 3, Name = "Get IDoc Information")]
        public void GetIDocInformation()
        {
            UIHelper.GoToSE16Table("EDIDC");
            getIDocInfo();
            var grid = SAPTestHelper.Current.MainWindow.FindDescendantByProperty<GuiGridView>();
            if (grid.RowCount > 0)
            {
                Dictionary<string, int> columns = new Dictionary<string, int>();
                columns.Add("IDoc number", -1);
                columns.Add("IDoc Status", -1);
                columns.Add("Message Type", -1);
                columns.Add("Message Variant", -1);
                columns.Add("Message function", -1);
                columns.Add("Created at", -1);
                UIHelper.ChangeLayout(columns);
                UIHelper.ExportFile("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]", _workDir, _EDIDCFileName);

                var dt = Tools.ReadToTable(Path.Combine(_workDir, _EDIDCFileName));
                _EDIDCDataList = dt.ToList<EDIDCDataModel>();
                _EDIDCDataList.ExportToFile(Path.Combine(_workDir, "EDIDC-Backup.txt"), "|");
            }
        }

        private void getIDocInfo()
        {
            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("%_I1_%_APP_%-VALU_PUSH").Press();
            var table = SAPTestHelper.Current.PopupWindow.FindByName<GuiTableControl>("SAPLALDBSINGLE");
            table.SetBatchValues(_idocList);
            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[8]").Press();
            SAPTestHelper.Current.MainWindow.FindByName<GuiTextField>("MAX_SEL").Text = _maximum.ToString();
            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("btn[8]").Press();
        }

        private List<EDIDSDataModel> _EDIDSDataList = null;
        private string _EDIDSFileName = "EDIDS.txt";

        [Step(Order = 4, Name = "Last date for status 69/51")]
        public void GetIDocInfomation2()
        {
            UIHelper.GoToSE16Table("EDIDS");
            getIDocInfo();
            var grid = SAPTestHelper.Current.MainWindow.FindDescendantByProperty<GuiGridView>();
            if (grid.RowCount > 0)
            {
                Dictionary<string, int> columns = new Dictionary<string, int>();
                columns.Add("IDoc number", -1);
                columns.Add("Date status error", -1);
                columns.Add("Time status error", -1);
                columns.Add("Status counter", -1);
                columns.Add("IDoc Status", -1);
                columns.Add("Name of person to be notified", -1);
                UIHelper.ChangeLayout(columns);
                UIHelper.ExportFile("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]", _workDir, _EDIDSFileName);

                var dt = Tools.ReadToTable(Path.Combine(_workDir, _EDIDSFileName));
                _EDIDSDataList = dt.ToList<EDIDSDataModel>();
                _EDIDSDataList.ExportToFile(Path.Combine(_workDir, "EDIDS-Backup.txt"), "|");
            }
        }



        [Step(Order = 5, Name = "Pick up agent and partner number")]
        public void PickupAgent()
        {
            var items = _EDIDCDataList.GroupBy(g => new { g.MessageFunction, g.MessageType, g.MessageVariant }).ToList();
            int fileIndex = 1;

            foreach (var item in items)
            {
                UIHelper.GoToSE16Table("EDP21");
                SAPTestHelper.Current.MainWindow.FindByName<GuiCTextField>("I4-LOW").Text = item.Key.MessageType;
                SAPTestHelper.Current.MainWindow.FindByName<GuiTextField>("I5-LOW").Text = item.Key.MessageVariant;
                SAPTestHelper.Current.MainWindow.FindByName<GuiTextField>("I6-LOW").Text = item.Key.MessageFunction;
                SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("btn[8]").Press();

                var grid = SAPTestHelper.Current.MainWindow.FindDescendantByProperty<GuiGridView>();
                if (grid.RowCount > 0)
                {
                    Dictionary<string, int> columns = new Dictionary<string, int>();
                    columns.Add("Partner No.", -1);
                    columns.Add("Message Type", -1);
                    columns.Add("Message code", -1);
                    columns.Add("Message function", -1);
                    columns.Add("Permitted users", -1);
                    UIHelper.ChangeLayout(columns);
                    UIHelper.ExportFile("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]", _edp21Folder, fileIndex + ".txt");
                }
                fileIndex++;
            }
        }

        private List<EDP21DataModel> _edp21DataList = null;

        [Step(Order = 6, Name = "Merge data from step 5")]
        public void MergeData()
        {
            _edp21DataList = new List<EDP21DataModel>();
            foreach (var f in Directory.GetFiles(_edp21Folder))
            {
                var dt = Tools.ReadToTable(f);
                var list = dt.ToList<EDP21DataModel>();
                _edp21DataList.AddRange(list);
            }
            _edp21DataList.ExportToFile(Path.Combine(_workDir, "EDP21-Backup.txt"), "|");
        }

        private List<HRP1001DataModel> _hRP1001DataList;
        private string _HRP1001FileName = "HRP1001.txt";
        [Step(Order = 7, Name = "Check User in HRP1001")]
        public void CheckUser()
        {
            var objIdList = _edp21DataList.GroupBy(g => g.Agent).Select(g => g.Key).ToList();
            UIHelper.GoToSE16Table("HRP1001");
            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("%_I2_%_APP_%-VALU_PUSH").Press();
            var table = SAPTestHelper.Current.PopupWindow.FindByName<GuiTableControl>("SAPLALDBSINGLE");
            table.SetBatchValues(objIdList);
            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[8]").Press();
            SAPTestHelper.Current.MainWindow.FindByName<GuiTextField>("MAX_SEL").Text = _maximum.ToString();
            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("btn[8]").Press();
            var grid = SAPTestHelper.Current.MainWindow.FindDescendantByProperty<GuiGridView>();
            if (grid.RowCount > 0)
            {
                Dictionary<string, int> columns = new Dictionary<string, int>();
                columns.Add("User name", -1);
                columns.Add("Object ID", -1);
                UIHelper.ChangeLayout(columns);
                UIHelper.ExportFile("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]", _workDir, "HRP1001.txt");

                var dt = Tools.ReadToTable(Path.Combine(_workDir, _HRP1001FileName));

                _hRP1001DataList = dt.ToList<HRP1001DataModel>();
                _hRP1001DataList = _hRP1001DataList.GroupBy(g => new { g.ObjectId, g.UserName }).Select(g => new HRP1001DataModel() { ObjectId = g.Key.ObjectId, UserName = g.Key.UserName }).ToList();
                _hRP1001DataList.ExportToFile(Path.Combine(_workDir, "HRP1001-Backup.txt"), "|");

            }


        }


        [Step(Order = 8, Name = "Generate Report")]
        public void GenerateReport()
        {
            var idocList = _EDIDSDataList.GroupBy(g => g.IDocNumber).Select(g => g.Key).ToList();
            List<EDIDSDataModel> edidsDataList = new List<EDIDSDataModel>();
            foreach(var item in idocList)
            {
                var idocInfo51 = _EDIDSDataList.Where(i => i.IDocNumber == item && i.Status == "51").OrderByDescending(o => o.StatusCounter).FirstOrDefault();
                var idocInfo69 = _EDIDSDataList.Where(i => i.IDocNumber == item && i.Status == "69").OrderByDescending(o => o.StatusCounter).FirstOrDefault();

                if(idocInfo51 != null)
                {
                    idocInfo51.LastDateOfStatus51 = idocInfo51.DataStatusError;
                }



                EDIDSDataModel dataItem = new EDIDSDataModel()
            }


            var items = from edidc in _EDIDCDataList
                        join edids in _EDIDSDataList on edidc.IDocNumber equals edids.IDocNumber
                        join edp21 in _edp21DataList
                        on new { a = edidc.MessageFunction, b = edidc.MessageType, c = edidc.MessageVariant }
                        equals
                        new { a = edp21.MessageFunction, b = edp21.MessageType, c = edp21.MessageCode }
                        select new Case6ReportModel
                        {
                            AgentNumber = edp21.Agent,
                            CreateDate = edidc.CreateDt,
                            CurrentStatus = edidc.Status,
                            //DataStatusError = edids.DataStatusError,
                            IDocNumber = edidc.IDocNumber,
                            MessageFunction = edidc.MessageFunction,
                            MessageType = edidc.MessageType,
                            MessageVariant = edidc.MessageVariant,
                            PartnerNumber = edp21.Partner,
                            //TimeStatusError = edids.TimeStatusError,
                            UserId = edids.UserID
                        };
            items.ToList().ExportToFile(Path.Combine(_workDir, "Report.csv"), ",");
        }

    }
}
