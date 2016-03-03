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

namespace TestScript.Case6
{
    public class Case6_Workflow : IScriptInitial
    {
        private Case6DataModel _data;
        private List<string> _idocList;
        private string _workDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Case6");
        private string _idocOutputFile = "IDocInfo.txt";
        private string _idocOutputFile2 = "IDocInfo2.txt";

        public void InitialData(object data)
        {
            _data = data as Case6DataModel;
            Tools.MasterDataVerification(_data);
            _idocList = Tools.GetDatas("test.txt");
        }

        [Step(Order = 1,Name ="Login to SAP")]
        public void Login()
        {
            UIHelper.Login(_data);
        }

        [Step(Order =2,Name ="TCode Verification")]
        public void TCodeVerifition()
        {
            UIHelper.SE16TableAccessVerification("EDIDC");
            UIHelper.SE16TableAccessVerification("EDIDS");
            UIHelper.SE16TableAccessVerification("EDP21");
            UIHelper.SE16TableAccessVerification("HRP1001");
        }

        [Step(Order =3,Name ="Get IDoc Information")]
        public void GetIDocInformation()
        {
            UIHelper.GoToSE16Table("EDIDC");
            getIDocInfo();
            var grid = SAPTestHelper.Current.MainWindow.FindDescendantByProperty<GuiGridView>();
            if(grid.RowCount>0)
            {
                Dictionary<string, int> columns = new Dictionary<string, int>();
                columns.Add("IDoc number", -1);
                columns.Add("IDoc Status", -1);
                columns.Add("Message Type", -1);
                columns.Add("Message Variant", -1);
                columns.Add("Message function", -1);
                columns.Add("Created at", -1);
                UIHelper.ChangeLayout(columns);
                UIHelper.ExportFile("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]", _workDir, _idocOutputFile);
            }
        }

        private void getIDocInfo()
        {
            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("%_I1_%_APP_%-VALU_PUSH").Press();
            var table = SAPTestHelper.Current.PopupWindow.FindByName<GuiTableControl>("SAPLALDBSINGLE");
            table.SetBatchValues(_idocList);
            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[8]").Press();
            SAPTestHelper.Current.MainWindow.FindByName<GuiTextField>("MAX_SEL").Text = "10000";
            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("btn[8]").Press();
        }

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
                UIHelper.ExportFile("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]", _workDir, _idocOutputFile2);
            }
        }

        [Step(Order =5,Name ="Pick up agent and partner number")]
        public void PickupAgent()
        {
            UIHelper.GoToSE16Table("EDP21");
        }

    }
}
