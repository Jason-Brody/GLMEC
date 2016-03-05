using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Data;
using Young.Data;

namespace TestScript
{
    public class Tools
    {
        public static void MasterDataVerification<T>(T data) where T : class
        {
            if (data == null)
                throw new ArgumentNullException(typeof(T).Name);

            var props = typeof(T).GetProperties();

            foreach (var prop in props)
            {
                var require = prop.GetCustomAttribute<RequiredAttribute>();
                if (require != null)
                {
                    object value = prop.GetValue(data);
                    var str = value?.ToString();
                    if (string.IsNullOrEmpty(str))
                    {
                        var alias = prop.GetCustomAttribute<DisplayAttribute>();
                        throw new ArgumentNullException(alias == null ? prop.Name : alias.Name);
                    }
                }
            }

        }

        public static List<string> GetDatas(string filePath)
        {
            List<string> datas = new List<string>();
            using (StreamReader sr = new StreamReader(filePath))
            {
                while (!sr.EndOfStream)
                {
                    datas.Add( sr.ReadLine());
                }

            }
            return datas;
        }

        public static DataTable ReadToTable(string path)
        {
            DataTable dt = Young.Data.Utils.ReadStringToTable(path, (s, h) =>
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

            return dt;
        }

        public static List<T> GetDataEntites<T>(string filePath) where T:class,new()
        {
            var dt = Tools.ReadToTable(filePath);
            List<T> datas = dt.ToList<T>();
            return datas;
        }
    }
}
