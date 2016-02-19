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

namespace TestScript.Case1
{
    public class Case1_MTD_Analysis
    {
        private DataTable _dt;



        public void Run()
        {


           

            

            ReadData();
            foreach (DataRow dr in _dt.Rows)
            {
                var data = dr.ToEntity<Case1DataModel>();
                login(data);
                getGLAccount(data);
                //getData(data);
            }
        }

        private void login(LoginDataModel data)
        {
            SAPLogon l = new SAPLogon();
            l.StartProcess();
            l.OpenConnection(data.Address);
            l.Login(data.UserName, data.Password, data.Client, data.Language);
            SAPTestHelper.Current.SetSession(l);
        }
        private void getData(Case1DataModel data)
        {
            SAPTestHelper h = SAPTestHelper.Current;
            h.SAPGuiSession.StartTransaction("FAGLL03");
            h.MainWindow.FindByName<GuiCTextField>("SD_SAKNR-LOW").Text = data.GLAccountFrom;
            h.MainWindow.FindByName<GuiCTextField>("SD_SAKNR-HIGH").Text = data.GLAccountTo;
            h.MainWindow.FindByName<GuiCTextField>("SD_BUKRS-LOW").Text = data.CompanyCode;
            h.MainWindow.FindByName<GuiRadioButton>("X_AISEL").Select();
            h.MainWindow.FindByName<GuiCTextField>("SO_BUDAT-LOW").Text = data.PostingDateFrom;
            h.MainWindow.FindByName<GuiCTextField>("SO_BUDAT-HIGH").Text = data.PostingDateTo;
            h.MainWindow.FindByName<GuiCTextField>("PA_VARI").Text = data.Layout;
        }

        private void getGLAccount(Case1DataModel data)
        {
            List<List<Tuple<int, string>>> accounts = new List<List<Tuple<int, string>>>();
            using (StreamReader sr = new StreamReader(@"Case1\Accountlist.txt"))
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

            SAPTestHelper.Current.MainWindow.FindById<GuiMenu>("mbar/menu[0]/menu[10]/menu[3]/menu[2]").Select();
            SAPTestHelper.Current.PopupWindow.FindByName<GuiRadioButton>("SPOPLI-SELFLAG").Select();
            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[0]").Press();

            var dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"temp");
            var fileName = data.CompanyCode + ".txt";

            var filePath = Path.Combine(dir, fileName);

            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            if (File.Exists(filePath))
                File.Delete(filePath);



            SAPTestHelper.Current.PopupWindow.FindByName<GuiCTextField>("DY_PATH").Text =dir;
            SAPTestHelper.Current.PopupWindow.FindByName<GuiCTextField>("DY_FILENAME").Text = fileName;

           

            var windowName = SAPTestHelper.Current.MainWindow.Text;

            
            Task.Run(() => { UIHelper.SetAccess(windowName); });
            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[0]").Press();
           

        }

        public void ReadData()
        {
            _dt = ExcelHelper.Current.Open("Case1_MTD_Analysis.xlsx").Read("Case1_MTD_Analysis");
            ExcelHelper.Current.Close();
        }
    }
}
