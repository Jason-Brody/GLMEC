using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ScriptRunner.Interface;
using ScriptRunner.Interface.Attributes;
using SAPAutomation;
using SAPFEWSELib;
using System.IO;

namespace TestScript.Case4
{
    [Script("GL MEC Case004 Parallel Ledger Reconcilication")]
    public class Case4_Parallel_Ledger_Reconcilication : IScriptRunner<Case4DataModel>
    {
        private Case4DataModel _data;
        private IProgress<ProgressInfo> _process;
        private Case4OutputModel _outputModel;

        public Case4_Parallel_Ledger_Reconcilication()
        {
            _outputModel = new Case4OutputModel();
        }

        public void SetInputData(Case4DataModel data, IProgress<ProgressInfo> MyProgress)
        {
            _data = data;
            _process = MyProgress;
            Tools.MasterDataVerification(data);
            var glAccountFile = Path.Combine(_outputModel.WorkDir, _data.GLAccountFilePath);
            _outputModel.GLAccounts = Tools.GetDatas(glAccountFile);
        }

        [Step(Id = 1, Name = "Login to SAP")]
        public void Login()
        {
            UIHelper.Login(_data);
        }

        [Step(Id =2,Name ="Check User Configuration in TCode SU3")]
        public void CheckUserConfig()
        {
            UIHelper.CheckUserConfig();
        }

        [Step(Id =3,Name ="TCode Verification")]
        public void TCodeVerification()
        {
            UIHelper.SE16TableAccessVerification("CSKS");
            UIHelper.SE16TableAccessVerification("COAS");
            UIHelper.SAPAccessVerification("FAGLL03");
        }

        [Step(Id=4,Name = "Cost Centers(CC) and Internal Order(IO) Validation")]
        public void step3()
        {
            UIHelper.GoToSE16Table("CSKS");
            SAPTestHelper.Current.MainWindow.FindByName<GuiCTextField>("I7-LOW").Text = _data.CompanyCode;
            SAPTestHelper.Current.MainWindow.FindByName<GuiTextField>("MAX_SEL").Text = _data.MaximumNo;
            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("btn[8]").Press();
            string columns = "Cost Center,Valid From,Valid To,Logical System";
            UIHelper.Export("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]", columns,_outputModel.CSKSFile);
        }

        [Step(Id =5,Name = "Extract Internal Orders")]
        public void Step4()
        {
            UIHelper.GoToSE16Table("COAS");
            SAPTestHelper.Current.MainWindow.FindByName<GuiCTextField>("I5-LOW").Text = _data.CompanyCode;
            SAPTestHelper.Current.MainWindow.FindByName<GuiTextField>("MAX_SEL").Text = _data.MaximumNo;
            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("btn[8]").Press();
            string columns = "Order,Logical System,Responsible CCtr";
            UIHelper.Export("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]", columns, _outputModel.COASFile);
        }

        [Step(Id =6,Name ="Extract of line item for the above cost centers")]
        public void Step5()
        {
            string systemName = SAPTestHelper.Current.SAPGuiSession.Info.SystemName;
            switch (systemName)
            {
                case "LH7":
                    _outputModel.CSKS = Tools.GetDataEntites<CSKSDataModel>(_outputModel.CSKSFile).Where(c => c.LogicalSystem != "" && c.LogicalSystem != "FINTLH7100").ToList();
                    break;
                case "LH4":
                    _outputModel.CSKS = Tools.GetDataEntites<CSKSDataModel>(_outputModel.CSKSFile).Where(c => c.LogicalSystem != "" && c.LogicalSystem != "FINTLH4100").ToList();
                    break;
                default:
                    _outputModel.CSKS = Tools.GetDataEntites<CSKSDataModel>(_outputModel.CSKSFile).Where(c => c.LogicalSystem != "").ToList();
                    break;
            }
            
            if(_outputModel.CSKS.Count < 1)
            {
                throw new Exception("No Cost Centers Found");
            }
            extratFAGLL03File(_outputModel.CSKS.GroupBy(g => g.CostCenter).Select(s => s.Key).ToList(), "%_%%DYN001_%_APP_%-VALU_PUSH", _outputModel.CostCenterFile);
        }

        [Step(Id = 7, Name = "Extract of line item for the above internal order")]
        public void Step6()
        {
            string systemName = SAPTestHelper.Current.SAPGuiSession.Info.SystemName;
            switch(systemName)
            {
                case "LH7":
                    _outputModel.COAS = Tools.GetDataEntites<COASDataModel>(_outputModel.COASFile).Where(c => c.LogicalSystem != "" && c.LogicalSystem != "FINTLH7100").ToList();
                    break;
                case "LH4":
                    _outputModel.COAS = Tools.GetDataEntites<COASDataModel>(_outputModel.COASFile).Where(c => c.LogicalSystem != "" && c.LogicalSystem != "FINTLH4100").ToList();
                    break;
                default:
                    _outputModel.COAS = Tools.GetDataEntites<COASDataModel>(_outputModel.COASFile).Where(c => c.LogicalSystem != "").ToList();
                    break;
            }
            if(_outputModel.COAS.Count < 1)
            {
                throw new Exception("No Internal Orders Found");
            }
            extratFAGLL03File(_outputModel.COAS.GroupBy(g => g.InternalOrder).Select(s => s.Key).ToList(), "%_%%DYN002_%_APP_%-VALU_PUSH", _outputModel.InternalOrderFile);
        }

        private void extratFAGLL03File(List<string> filterDatas,string filterButtonName,string file)
        {
            SAPTestHelper.Current.SAPGuiSession.StartTransaction("FAGLL03");
            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("%_SD_SAKNR_%_APP_%-VALU_PUSH").Press();

            SAPTestHelper.Current.PopupWindow.FindDescendantByProperty<GuiTableControl>().SetBatchValues(_outputModel.GLAccounts);
            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[8]").Press();

            SAPTestHelper.Current.MainWindow.FindByName<GuiCTextField>("SD_BUKRS-LOW").Text = _data.CompanyCode;
            SAPTestHelper.Current.MainWindow.FindByName<GuiRadioButton>("X_AISEL").Select();
            SAPTestHelper.Current.MainWindow.FindByName<GuiCTextField>("SO_BUDAT-LOW").Text = _data.PostingDateFrom;
            SAPTestHelper.Current.MainWindow.FindByName<GuiCTextField>("SO_BUDAT-HIGH").Text = _data.PostingDateTo;

            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("btn[25]").Press();
            var btn = SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("btn[5]");
            btn?.Press();

            var tree = SAPTestHelper.Current.MainWindow.FindDescendantByProperty<GuiTree>();
            tree.ChooseNode("General Ledger Line Items->Cost Center");
            tree.ChooseNode("General Ledger Line Items->Order");
            SAPTestHelper.Current.SAPGuiSession.FindById<GuiToolbarControl>("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell").PressButton("TAKE");
            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>(filterButtonName).Press();
            SAPTestHelper.Current.PopupWindow.FindDescendantByProperty<GuiTableControl>().SetBatchValues(filterDatas);
            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[8]").Press();
            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("btn[11]").Press();


            SAPTestHelper.Current.MainWindow.FindByName<GuiButton>("btn[8]").Press();

            string columns = "Company Code,G/L Account,Document Number,Business Area,Profit Center,Cost Center,Order,Document currency,Amount in doc. curr.,Local Currency,Amount in local currency,Local currency 2,Amount in loc.curr.2";
            UIHelper.Export("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]", columns, file);
        }
    }
}
