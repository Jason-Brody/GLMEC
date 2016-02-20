using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;
using TestScript.Case1;
using TestScript;
using System.Data;
using System.IO;

namespace CaseTrigger
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(DateTime.Now);
            mergeData(@"C:\Work\GLMEC\CaseTrigger\bin\Debug\ReportData\DE50\Datas");
            Console.WriteLine(DateTime.Now);
            Console.WindowHeight = 1;
            Console.WindowWidth = 1;
            Case1_MTD_Analysis case1 = new Case1_MTD_Analysis();
            case1.Run();


        }


        static void mergeData(string dir)
        {
            DataTable mergeTable = null;
            
            foreach (var f in Directory.GetFiles(dir))
            {
                DataTable dt = Utils.ReadStringToTable(f, (s, h) =>
                {
                    string splitChar = "|";
                    if (!s.Contains(splitChar) || s == h || s.Contains("*"))
                        return null;

                    var vals = s.Split(splitChar.ToCharArray().First());
                    var returnVals = new List<string>();
                    for (int i = 0; i < vals.Count(); i++)
                    {
                        returnVals.Add(vals[i].Trim());
                    }
                    return returnVals;

                });

                if (mergeTable == null)
                    mergeTable = dt.Copy();
                else
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        DataRow newDr = mergeTable.NewRow();
                        for (int i = 0; i < mergeTable.Columns.Count; i++)
                        {
                            newDr[i] = dr[i];
                        }
                        mergeTable.Rows.Add(newDr);
                    }

                }

                
            }

            Console.WriteLine(DateTime.Now);

        }




    }
}
