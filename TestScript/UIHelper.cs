using SAPAutomation;
using SAPFEWSELib;
using System;
using System.Collections.Generic;
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
        public static void SetAccess(string windowName,CancellationToken token)
        {
            

            bool isPress = false;
            while(!isPress)
            {
                if (token.IsCancellationRequested)
                    break;

                var e = TreeWalker.ControlViewWalker.GetFirstChild(AutomationElement.RootElement);

                while (e != null)
                {
                    if (e.Current.Name == windowName)
                        break;
                    var tempE = TreeWalker.ControlViewWalker.GetNextSibling(e);
                    e = tempE;
                }

                if (e != null)
                {
                    var condition1 = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.CheckBox);
                    var checkboxElement = e.FindFirst(TreeScope.Descendants, condition1);
                    if (checkboxElement != null)
                    {
                        var checkbox = checkboxElement.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                        if (checkbox.Current.ToggleState == ToggleState.Off)
                        {
                            checkbox.Toggle();
                        }
                    }

                    condition1 = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
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

            var ts = new CancellationTokenSource();
            var ct = ts.Token;

            Task.Run(() => { UIHelper.SetAccess(windowName, ct); });
            SAPTestHelper.Current.PopupWindow.FindByName<GuiButton>("btn[0]").Press();
            ts.Cancel();
        }
    }
}
