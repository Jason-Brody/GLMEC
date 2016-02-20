using System;
using System.Collections.Generic;
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
            if (token.IsCancellationRequested)
                token.ThrowIfCancellationRequested();

            bool isPress = false;
            while(!isPress)
            {
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
    }
}
