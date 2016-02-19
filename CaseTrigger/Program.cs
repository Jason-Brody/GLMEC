using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;
using TestScript.Case1;

namespace CaseTrigger
{
    class Program
    {
        static void Main(string[] args)
        {

            //var e = TreeWalker.ControlViewWalker.GetFirstChild(AutomationElement.RootElement);

            //while (e != null)
            //{
            //    Console.WriteLine(e.Current.Name);
            //    var tempE = TreeWalker.ControlViewWalker.GetNextSibling(e);
            //    e = tempE;


            //}

            Console.WindowHeight = 1;
            Console.WindowWidth = 1;
            Case1_MTD_Analysis case1 = new Case1_MTD_Analysis();
            case1.Run();


        }





    }
}
