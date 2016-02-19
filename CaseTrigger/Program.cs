using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;
using TestScript.Case1;
using TestScript;
using System.Data;

namespace CaseTrigger
{
    class Program
    {
        static void Main(string[] args)
        {

            //string file = @"temp\DE50.txt";
            //var table = Utils.ReadStringToTable(file, (s) =>
            //{
            //    string splitchar = "|";
            //    if (!s.Contains(splitchar))
            //        return null;
            //    var vals = s.Split(splitchar.ToCharArray().First()).ToList();
            //    var returnVals = new List<string>();
            //    vals.ForEach(temps => returnVals.Add(temps.Trim()));
            //    return returnVals;
            //});

            //List<AccountModel> accounts = new List<AccountModel>();
            
            //foreach (DataRow dr in table.Rows)
            //{
            //    AccountModel acct = dr.ToEntity<AccountModel>();
            //    accounts.Add(acct);

            //}

            //int b = accounts.Count / 100;

            //for(int i =0;i< b;i++)
            //{
            //    List<List<Tuple<int, string>>> testAccounts = accounts.Skip(i * 100).Take(100).Select(
            //       c => new List<Tuple<int, string>>() { new Tuple<int, string>(1, c.Account) }).ToList();
            //}
           

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
