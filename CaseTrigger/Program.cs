using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestScript.Case1;

namespace CaseTrigger
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.SetWindowSize(20, 20);
            Case1_MTD_Analysis case1 = new Case1_MTD_Analysis();
            case1.Run();
        }
    }
}
