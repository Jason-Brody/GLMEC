using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ScriptRunner.Interface.Attributes;
using SAPAutomation;
using SAPFEWSELib;
using System.IO;
using Young.Data;
using Young.Data.Extension;
using Young.Excel.Interop.Extensions;
using ScriptRunner.Interface;

namespace TestScript.Case6
{
    [Script("GL_MEC_Case006_WorkFlow")]
    public class Case6_Workflow : IScriptRunner<Case6DataModel>
    {
        private Case6DataModel _data;
        private List<string> _idocList;
        private int _maximum = 100000;
        private Case6FileConfig _fileConfig = null;
        private IProgress<ProgressInfo> _progress = null;

        

        public void SetInputData(Case6DataModel data, IProgress<ProgressInfo> MyProgress)
        {
            _data = data;
            _progress = MyProgress;
            Tools.MasterDataVerification(_data);
            _fileConfig = new Case6FileConfig(Path.Combine(Environment.CurrentDirectory, "Case6"));
            _idocList = Tools.GetDatas(Path.Combine(_fileConfig.WorkFolder, "IDocList.txt"));
        }

       



        [Step(Id = 1, Name = "Login to SAP")]
        public void Login()
        {
            UIHelper.Login(_data);
        }

        [Step(Id = 2, Name = "TCode Verification")]
        public void TCodeVerifition()
        {
            UIHelper.SE16TableAccessVerification("EDIDC");
            UIHelper.SE16TableAccessVerification("EDIDS");
            UIHelper.SE16TableAccessVerification("EDP21");
            UIHelper.SE16TableAccessVerification("HRP1001");
        }

        private List<EDIDCDataModel> _EDIDCDataList = null;

        [Step(Id = 3, Name = "Get IDoc Information")]
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
                columns.Add("Created on", -1);
                UIHelper.ChangeLayout(columns);
                UIHelper.ExportFile("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]", _fileConfig.SourceOutputFolder, _fileConfig.EDIDCFileName);
                _EDIDCDataList = Tools.GetDataEntites<EDIDCDataModel>(_fileConfig.GetSourcePath(_fileConfig.EDIDCFileName));
                
            }
        }

        private void getIDocInfo()
        {
            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("%_I1_%_APP_%-VALU_PUSH").Press();
            var table = SAPTestHelper.Current.PopupWindow.FindByName<GuiTableControl>("SAPLALDBSINGLE");
            table.SetBatchValues(_idocList,i=> {
                _progress.Report(new ProgressInfo() { IsProgressKnow= true, Current = i++, Total = _idocList.Count });
            });
            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[8]").Press();
            SAPTestHelper.Current.MainWindow.FindByName<GuiTextField>("MAX_SEL").Text = _maximum.ToString();
            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("btn[8]").Press();
        }

        private List<EDIDSDataModel> _EDIDSDataList = null;

        [Step(Id = 4, Name = "Last date for status 69/51")]
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
                UIHelper.ExportFile("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]", _fileConfig.SourceOutputFolder, _fileConfig.EDIDSFileName);

                _EDIDSDataList = Tools.GetDataEntites<EDIDSDataModel>(_fileConfig.GetSourcePath(_fileConfig.EDIDSFileName));
                //_EDIDSDataList.ExportToFile(Path.Combine(_workDir, "EDIDS-Backup.txt"), "|");
            }
        }

        [Step(Id = 5, Name = "Pick up agent and partner number")]
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
                    UIHelper.ExportFile("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]", _fileConfig.GetSourcePath(_fileConfig.EDP21Folder), fileIndex + ".txt");
                }
                fileIndex++;
            }
        }

        private List<EDP21DataModel> _edp21DataList = null;

        [Step(Id = 6, Name = "Merge data from step 5")]
        public void MergeData()
        {
            _edp21DataList = new List<EDP21DataModel>();
            foreach (var f in Directory.GetFiles(_fileConfig.GetSourcePath(_fileConfig.EDP21Folder)))
            {
                var dt = Tools.ReadToTable(f);
                var list = dt.ToList<EDP21DataModel>();
                _edp21DataList.AddRange(list);
            }
            _edp21DataList.ExportToFile(_fileConfig.GetSourcePath(_fileConfig.EDP21FileName), "|");
        }

        private List<HRP1001DataModel> _hRP1001DataList;
        [Step(Id = 7, Name = "Check User in HRP1001")]
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
                UIHelper.ExportFile("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]", _fileConfig.SourceOutputFolder, _fileConfig.HRP1001FileName);

                _hRP1001DataList = Tools.GetDataEntites<HRP1001DataModel>(_fileConfig.GetSourcePath(_fileConfig.HRP1001FileName));
               

            }
        }

        private void readData<T>(ref List<T> obj, string fileName) where T :class,new()
        {
            if(obj == null || obj.Count == 0)
            {
                obj = Tools.GetDataEntites<T>(_fileConfig.GetSourcePath(fileName));
            }
            obj.ExportToFile(_fileConfig.GetFullPath(fileName), "|");
        }

        [Step(Id = 8, Name = "Generate Report")]
        public void GenerateReport()
        {
            readData(ref _EDIDCDataList, _fileConfig.EDIDCFileName);
            readData(ref _EDIDSDataList, _fileConfig.EDIDSFileName);
            readData(ref _hRP1001DataList, _fileConfig.HRP1001FileName);
            MergeData();
            //更新EDIDSList, 找到51和69的数据
            var idocList = _EDIDSDataList.GroupBy(g => g.IDocNumber).Select(g => g.Key).ToList();
            List<EDIDSDataModel> edidsDataList = new List<EDIDSDataModel>();
            foreach(var item in idocList)
            {
                var idocInfo51 = _EDIDSDataList.Where(i => i.IDocNumber == item && i.Status == "51").OrderByDescending(o => o.StatusCounter).FirstOrDefault();
                var idocInfo69 = _EDIDSDataList.Where(i => i.IDocNumber == item && i.Status == "69").OrderByDescending(o => o.StatusCounter).FirstOrDefault();
                EDIDSDataModel dataItem = new EDIDSDataModel();
                dataItem.IDocNumber = item;
                dataItem.LastDateOfStatus51 = idocInfo51?.DataStatusError;
                dataItem.LastDateOfStatus69 = idocInfo69?.DataStatusError;

                var user1 = idocInfo51?.UserID;
                var user2 = idocInfo69?.UserID;

                if(user1 == user2)
                {
                    dataItem.StatusCounter = idocInfo51?.StatusCounter +":"+idocInfo69?.StatusCounter;
                    edidsDataList.Add(dataItem);
                }
                else
                {
                    dataItem.UserID = user1;
                    dataItem.StatusCounter = idocInfo51?.StatusCounter;
                    dataItem.Status = idocInfo51?.Status;
                    edidsDataList.Add(dataItem);
                    var dataItem2 = new EDIDSDataModel(dataItem);
                    dataItem2.UserID = user2;
                    dataItem2.StatusCounter = idocInfo69?.StatusCounter;
                    dataItem2.Status = idocInfo69?.Status;
                    edidsDataList.Add(dataItem2);
                }
            }

            edidsDataList.ExportToFile(Path.Combine(_fileConfig.WorkFolder, "EDIDS-Update.txt"), "|");
            var hRP100DataList = _hRP1001DataList.GroupBy(g => new { g.ObjectId,g.UserName}).Select(g=>new HRP1001DataModel() { ObjectId = g.Key.ObjectId,UserName=g.Key.UserName}).ToList();

            var items = from edidc in _EDIDCDataList
                        join edids in edidsDataList on edidc.IDocNumber equals edids.IDocNumber
                        join edp21 in _edp21DataList
                        on new { a = edidc.MessageFunction, b = edidc.MessageType, c = edidc.MessageVariant }
                        equals
                        new { a = edp21.MessageFunction, b = edp21.MessageType, c = edp21.MessageCode }
                        select new Case6ReportModel
                        {
                            AgentNumber = edp21.Agent,
                            CreateDate = edidc.CreateDt,
                            CurrentStatus = edidc.Status,
                            LastDateOfStatus51 = edids.LastDateOfStatus51,
                            LastDateOfStatus69 = edids.LastDateOfStatus69,
                            IDocNumber = edidc.IDocNumber,
                            MessageFunction = edidc.MessageFunction,
                            MessageType = edidc.MessageType,
                            MessageVariant = edidc.MessageVariant,
                            PartnerNumber = edp21.Partner,
                            UserId = edids.UserID,
                        };

            var reportDatas = items.ToList();

            foreach(var data in reportDatas)
            {
                if (hRP100DataList.Where(c => c.ObjectId == data.AgentNumber && c.UserName == data.UserId).FirstOrDefault() != null)
                    data.IsExisted = "Y";
                else
                    data.IsExisted = "N";
            }

            reportDatas.ExportToFile( _fileConfig.GetFullPath(_fileConfig.ReportCSVFile), ",");


            _EDIDSDataList.ExportToExcel("IDOC_STATUS_CHANGE_HISTORY", null, "");
            //for(int i = idocList.Count-1;i>=0;i--)
            //{
            //    var historyDatas = _EDIDSDataList.Where(s => s.IDocNumber == idocList[i]).ToList();
            //    historyDatas.ExportToExcel($"IDOC_STATUS_CHANGE_HISTORY_{i}", null, "");
                   
            //}

            reportDatas.ExportToExcel("Report", null, _fileConfig.GetFullPath(_fileConfig.ReportExcelFile));


        }

        [Step(Id = 9,Name ="Exit SAP")]
        public void CloseSAP()
        {
            SAPTestHelper.Current.SAPGuiConnection.CloseConnection();
        }
    }
}
