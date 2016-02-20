using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestScript
{
    public class Utils
    {
        

        public static DataTable ReadStringToTable(string filePath, Func<string,string,List<string>> LineFunc)
        {
            string tempString = "";
            DataTable table = null;
            string headerRow = "";
            using (StreamReader sr = new StreamReader(filePath))
            {
                while (!sr.EndOfStream)
                {
                    tempString = sr.ReadLine();
                    var vals = LineFunc(tempString,headerRow);
                    if (vals != null && vals.Count() > 0)
                    {
                        if (table == null)
                        {
                            table = new DataTable();
                            headerRow = tempString;
                            for (int i = 0; i < vals.Count(); i++)
                            {
                                if (vals[i] == "")
                                {
                                    vals[i] = "Header_Temp_" + i.ToString();
                                }
                                table.Columns.Add(vals[i]);
                            }
                        }
                        else
                        {
                            DataRow dr = table.NewRow();
                            for (int i = 0; i < vals.Count(); i++)
                            {
                                dr[i] = vals[i];
                            }
                            table.Rows.Add(dr);
                        }
                    }
                }
            }
            return table;
        }
    }
}
