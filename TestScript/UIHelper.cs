using SAPAutomation;
using SAPFEWSELib;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;

namespace TestScript
{
    class UIHelper
    {
        public static void ChangeLayout(HashSet<string> columns)
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

                if (columns.Contains(col))
                {
                    selectedRow += c.ToString() + ",";
                }

            }

            columnSetGrid.SelectedRows = selectedRow;
            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("APP_WL_SING").Press();

            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[0]").Press();
        }
        public static void ChangeLayout(Dictionary<string, int> columns)
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

        public static void ExportFile(string outputmenuId, string dir, string fileName)
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
            var args = "\"" + windowName + "\"";

            var ts = new CancellationTokenSource();
            var ct = ts.Token;

            Task.Run(() => { Utils.SetAccess(windowName, ct); });
            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[0]").Press();
            ts.Cancel();
            
        }

        public static void ExportFile(string outputmenuId,  string fileName)
        {
            SAPTestHelper.Current.MainWindow.FindById<GuiMenu>(outputmenuId).Select();
            SAPTestHelper.Current.PopupWindow.FindByName<GuiRadioButton>("SPOPLI-SELFLAG").Select();
            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[0]").Press();

            FileInfo f = new FileInfo(fileName);
            if (!f.Directory.Exists)
                f.Directory.Create();
            if (f.Exists)
                f.Delete();

            

            SAPTestHelper.Current.PopupWindow.FindByName<GuiCTextField>("DY_PATH").Text = f.DirectoryName;
            SAPTestHelper.Current.PopupWindow.FindByName<GuiCTextField>("DY_FILENAME").Text = f.Name;

            var windowName = SAPTestHelper.Current.MainWindow.Text;
            var args = "\"" + windowName + "\"";

            var ts = new CancellationTokenSource();
            var ct = ts.Token;

            Task.Run(() => { Utils.SetAccess(windowName, ct); });
            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[0]").Press();
            ts.Cancel();

        }

        public static bool Export(string outputMenuId,string columnsDivideByComma, string fileName)
        {
            var grid = SAPTestHelper.Current.MainWindow.FindDescendantByProperty<GuiGridView>();
            if (grid.RowCount > 0)
            {
                Dictionary<string, int> columns = new Dictionary<string, int>();
                columnsDivideByComma.Split(',').ToList().ForEach(s => columns.Add(s, -1));
                UIHelper.ChangeLayout(columns);
                UIHelper.ExportFile("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]", fileName);
                return true;
            }
            return false;
        }

        public static void SAPAccessVerification(string TCode, Func<Tuple<bool, string>> otherVerification = null)
        {
            SAPTestHelper.Current.SAPGuiSession.StartTransaction(TCode);
            if (SAPTestHelper.Current.SAPGuiSession.Info.Transaction != TCode)
                throw new Exception($"Can't access to TCode:{TCode}");

            if (otherVerification != null)
            {
                var result = otherVerification();
                if (!result.Item1)
                    throw new Exception(result.Item2);
            }
        }

        public static void SE16TableAccessVerification(string tableName)
        {
            SAPAccessVerification("SE16", () =>
            {
                var page = SAPTestHelper.Current.SAPGuiSession.Info.ScreenNumber;
                SAPTestHelper.Current.MainWindow.FindByName<GuiCTextField>("DATABROWSE-TABLENAME").Text = tableName;
                SAPTestHelper.Current.MainWindow.SendKey(SAPKeys.Enter);
                var page1 = SAPTestHelper.Current.SAPGuiSession.Info.ScreenNumber;
                if (page == page1)
                {
                    return new Tuple<bool, string>(false, $"Don't have access to TCODE:SE16 ,table:{tableName}");
                }
                else
                {
                    return new Tuple<bool, string>(true, "");
                }
            });
        }

        public static void Login(LoginDataModel data)
        {
            SAPLogon l = new SAPLogon();
            l.StartProcess();
            l.OpenConnection(data.Address);
            l.Login(data.UserName, data.Password, data.Client, data.Language);
            SAPTestHelper.Current.SetSession(l);

            SAPTestHelper.Current.OnRequestError += (s, e) => { throw new Exception(e.Message); };
        }

        public static void GoToSE16Table(string tableName)
        {
            SAPTestHelper.Current.SAPGuiSession.StartTransaction("SE16");
            SAPTestHelper.Current.MainWindow.FindByName<GuiCTextField>("DATABROWSE-TABLENAME").Text = tableName;
            SAPTestHelper.Current.MainWindow.SendKey(SAPKeys.Enter);
        }

        public static void CheckUserConfig()
        {
            SAPTestHelper.Current.SAPGuiSession.StartTransaction("su3");
            SAPTestHelper.Current.MainWindow.FindByName<GuiTab>("DEFA").Select();

            var decimalNotation = SAPTestHelper.Current.MainWindow.FindByName<GuiComboBox>("SUID_ST_NODE_DEFAULTS-DCPFM");

            bool isChange = false;

            if (decimalNotation.Value != "1,234,567.89")
            {
                decimalNotation.Value = "1,234,567.89";
                isChange = true;
            }

            var dateFormat = SAPTestHelper.Current.MainWindow.FindByName<GuiComboBox>("SUID_ST_NODE_DEFAULTS-DATFM");
            if (dateFormat.Value != "DD.MM.YYYY")
            {
                dateFormat.Value = "DD.MM.YYYY";
                isChange = true;
            }

            if (isChange)
                SAPTestHelper.Current.MainWindow.SendKey(SAPKeys.Ctrl_S);
            else
                SAPTestHelper.Current.MainWindow.SendKey(SAPKeys.F3);
        }
    }
}
